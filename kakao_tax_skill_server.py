"""
카카오톡 채널(오픈빌더) 세무·회계 Q&A 챗봇 - 스킬 서버 스켈레톤

동작 방식:
1. 사용자가 카카오톡 채널에 질문을 보냄
2. 오픈빌더가 이 서버의 /skill 엔드포인트로 POST 요청 (webhook)
3. 이 서버가 Claude API를 호출해 답변 생성
4. 오픈빌더가 요구하는 JSON 포맷으로 답변을 반환 -> 사용자 채팅창에 출력

실행 전 준비:
  pip install fastapi uvicorn httpx
  export ANTHROPIC_API_KEY="sk-ant-..."
  export LAWGOKR_OC="본인의-law.go.kr-인증키"   # 없으면 법령 조회 없이 Claude 기본 지식으로만 답변
  uvicorn kakao_tax_skill_server:app --host 0.0.0.0 --port 8000

배포 시:
  - 이 서버가 인터넷에서 접근 가능한 HTTPS 주소여야 오픈빌더 웹훅 등록이 가능합니다.
    (예: Render, Railway, Fly.io, AWS/GCP 등에 배포)
  - 오픈빌더 > 스킬 관리 > 스킬 서버 URL에 "https://내도메인/skill" 등록
  - ANTHROPIC_API_KEY, LAWGOKR_OC 모두 절대 코드에 하드코딩하지 말고 서버 환경변수로 관리하세요.

주의:
  - law.go.kr API 응답의 실제 XML 태그명(법령일련번호, 조문내용 등)은 국가법령정보 공동활용
    공식 가이드(open.law.go.kr)와 실제 호출 결과를 보고 한 번 검증/조정이 필요합니다.
    이 코드는 뼈대이며, 태그명이 다르면 해당 부분만 맞춰 수정하면 됩니다.
"""

import os
import re
import logging
import xml.etree.ElementTree as ET

import httpx
from fastapi import FastAPI, Request

app = FastAPI()
logger = logging.getLogger("kakao_tax_skill")

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_MODEL = "claude-sonnet-4-6"

# 국가법령정보센터(open.law.go.kr) 오픈API 인증키 - 반드시 환경변수로만 관리
LAWGOKR_OC = os.environ.get("LAWGOKR_OC", "")
LAWGOKR_SEARCH_URL = "https://www.law.go.kr/DRF/lawSearch.do"
LAWGOKR_SERVICE_URL = "https://www.law.go.kr/DRF/lawService.do"

# 질문에서 관련 세법을 유추하기 위한 키워드 매핑 (필요에 따라 계속 추가)
TAX_LAW_KEYWORDS = {
    "소득세": "소득세법",
    "법인세": "법인세법",
    "부가가치세": "부가가치세법",
    "부가세": "부가가치세법",
    "상속세": "상속세및증여세법",
    "증여세": "상속세및증여세법",
    "국세기본": "국세기본법",
    "취득세": "지방세법",
    "재산세": "지방세법",
    "지방세": "지방세법",
}

DISCLAIMER = (
    "\n\n※ 본 답변은 AI가 생성한 참고용 정보이며, "
    "구체적인 판단·처리는 담당 세무사·회계사와 상담하시기 바랍니다."
)

SYSTEM_PROMPT = """당신은 한국의 세무·회계 실무를 안내하는 상담 챗봇입니다.
- 국세, 지방세, K-IFRS, K-GAAP, 감사기준서, 관련 법령(공인회계사법, 외감법 등)에 대한
  질문에 정확하고 실무적으로 답합니다.
- 확실하지 않은 내용, 최근 법 개정 여부가 불분명한 내용은 추측하지 말고
  "최신 법령을 확인해야 한다"고 명시하세요.
- 답변은 카카오톡 채팅창에서 읽기 좋도록 3~6문장 내외로 간결하게 작성하세요.
- 지나치게 긴 목록이나 표는 피하고, 핵심만 전달하세요.
- 회계기준(K-IFRS, K-GAAP) 관련 질문에는 기준서 원문을 그대로 인용하지 말고,
  핵심 내용을 요약해서 답하세요. 원문 조문 확인이 필요하다는 점을 답변에 자연스럽게 언급하세요.
"""

# 회계기준 질문에는 원문 확인 링크를 답변 끝에 덧붙임
KASB_NOTICE = (
    "\n\n📌 회계기준 원문은 한국회계기준원 회계기준열람서비스"
    "(https://db.kasb.or.kr/standard/)에서 확인하실 수 있습니다."
)

ACCOUNTING_KEYWORDS = ["IFRS", "GAAP", "회계기준", "기준서", "일반기업회계기준", "감가상각", "손상차손"]


def needs_kasb_notice(question: str, answer: str) -> bool:
    combined = f"{question} {answer}"
    return any(keyword.lower() in combined.lower() for keyword in ACCOUNTING_KEYWORDS)


def detect_law_name(question: str) -> str | None:
    """질문 문장에서 관련 세법명을 키워드 매칭으로 유추"""
    for keyword, law_name in TAX_LAW_KEYWORDS.items():
        if keyword in question:
            return law_name
    return None


async def search_law_mst(law_name: str) -> str | None:
    """법령명으로 검색해 법령일련번호(MST)를 반환. 실패 시 None."""
    if not LAWGOKR_OC:
        return None

    params = {
        "OC": LAWGOKR_OC,
        "target": "law",
        "type": "XML",
        "query": law_name,
        "display": 1,
    }
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            resp = await client.get(LAWGOKR_SEARCH_URL, params=params)
            resp.raise_for_status()
            root = ET.fromstring(resp.text)
    except Exception:
        logger.exception("law.go.kr 법령 검색 실패: %s", law_name)
        return None

    # 응답 스키마는 API 가이드 기준이며, 실제 태그명은 테스트 후 조정이 필요할 수 있음
    law_elem = root.find(".//law")
    if law_elem is None:
        return None
    mst_elem = law_elem.find("법령일련번호")
    return mst_elem.text if mst_elem is not None else None


async def fetch_law_summary(mst: str, max_chars: int = 1500) -> str | None:
    """법령일련번호로 본문을 조회해 앞부분 일부만 요약용으로 반환"""
    if not LAWGOKR_OC:
        return None

    params = {
        "OC": LAWGOKR_OC,
        "target": "law",
        "type": "XML",
        "MST": mst,
    }
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            resp = await client.get(LAWGOKR_SERVICE_URL, params=params)
            resp.raise_for_status()
            root = ET.fromstring(resp.text)
    except Exception:
        logger.exception("law.go.kr 법령 본문 조회 실패: MST=%s", mst)
        return None

    # 짧은 조문은 <조문내용>에 전체가 들어있지만, 항/호/목으로 나뉜 조문은
    # <조문내용>에 제목만 있고 실제 내용은 <항내용>/<호내용>/<목내용>에 들어있음.
    # 문서 순서대로(iter()는 문서 순서를 보장) 관련 태그를 모두 모아 이어붙임.
    CONTENT_TAGS = {"조문내용", "항내용", "호내용", "목내용"}
    texts = [
        elem.text.strip()
        for elem in root.iter()
        if elem.tag in CONTENT_TAGS and elem.text and elem.text.strip()
    ]
    combined = "\n".join(texts)
    if not combined:
        return None
    return combined[:max_chars]


async def get_law_context(question: str) -> str | None:
    """질문에서 관련 법령을 찾아 최신 조문 일부를 컨텍스트로 반환. 실패해도 None만 반환하고 진행."""
    law_name = detect_law_name(question)
    if not law_name:
        return None

    mst = await search_law_mst(law_name)
    if not mst:
        return None

    summary = await fetch_law_summary(mst)
    if not summary:
        return None

    return f"[참고: {law_name} 관련 조문 발췌]\n{summary}"


def build_kakao_response(text: str) -> dict:
    """오픈빌더 스킬 응답 규격(simpleText)에 맞춘 JSON 생성"""
    return {
        "version": "2.0",
        "template": {
            "outputs": [
                {"simpleText": {"text": text}}
            ]
        }
    }


async def ask_claude(question: str, law_context: str | None = None) -> str:
    if not ANTHROPIC_API_KEY:
        return "서버에 ANTHROPIC_API_KEY가 설정되지 않았습니다. 관리자에게 문의하세요."

    headers = {
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }

    if law_context:
        user_content = (
            f"{question}\n\n"
            f"(다음은 국가법령정보센터에서 조회한 관련 법령 조문 발췌입니다. "
            f"답변의 근거로 참고하되, 조문을 그대로 옮기지 말고 요약해서 답하세요.)\n"
            f"{law_context}"
        )
    else:
        user_content = question

    payload = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": 800,
        "system": SYSTEM_PROMPT,
        "messages": [{"role": "user", "content": user_content}],
    }

    try:
        async with httpx.AsyncClient(timeout=25) as client:
            resp = await client.post(ANTHROPIC_URL, headers=headers, json=payload)
            resp.raise_for_status()
            data = resp.json()
    except Exception:
        logger.exception("Claude API 호출 실패")
        return "일시적인 오류로 답변을 생성하지 못했습니다. 잠시 후 다시 시도해 주세요."

    text_blocks = [
        block.get("text", "")
        for block in data.get("content", [])
        if block.get("type") == "text"
    ]
    answer = "\n".join(text_blocks).strip()

    if not answer:
        return "답변을 생성하지 못했습니다. 질문을 조금 더 구체적으로 입력해 주세요."

    if needs_kasb_notice(question, answer):
        answer += KASB_NOTICE

    return answer + DISCLAIMER


@app.post("/skill")
async def kakao_skill(request: Request):
    body = await request.json()

    # 오픈빌더가 보내는 사용자 발화(질문) 위치
    try:
        user_utterance = body["userRequest"]["utterance"]
    except (KeyError, TypeError):
        return build_kakao_response("질문을 인식하지 못했습니다. 다시 입력해 주세요.")

    if not user_utterance or not user_utterance.strip():
        return build_kakao_response("질문 내용이 비어 있습니다.")

    question = user_utterance.strip()
    law_context = await get_law_context(question)
    answer = await ask_claude(question, law_context)
    return build_kakao_response(answer)


@app.get("/health")
async def health():
    return {"status": "ok"}
