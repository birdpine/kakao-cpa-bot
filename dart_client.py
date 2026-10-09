"""
OpenDART API 클라이언트

사용하는 API (https://opendart.fss.or.kr 개발가이드 참조):
  - corpCode.xml        : 공시대상회사 고유번호 전체 목록(ZIP). 비상장 외감법인 포함
  - list.json           : 공시검색 (감사보고서 / 사업보고서 목록)
  - document.xml        : 공시서류 원본파일(ZIP, 내부는 DART XML)
  - fnlttSinglAcntAll.json : 단일회사 전체 재무제표(XBRL, 사업보고서 제출회사만)

인증키는 환경변수 DART_API_KEY 로만 관리합니다.
"""

import io
import os
import re
import time
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from pathlib import Path

import httpx

DART_API_KEY = os.environ.get("DART_API_KEY", "")
DART_BASE_URL = "https://opendart.fss.or.kr/api"
CACHE_DIR = Path(os.environ.get("DART_CACHE_DIR", ".cache"))
CORP_CODE_TTL_SECONDS = 24 * 60 * 60

# list.json 상태코드: 000 정상, 013 조회된 데이터 없음
STATUS_OK = "000"
STATUS_NO_DATA = "013"

# 사업보고서 재무제표 보고서코드
REPRT_CODE_ANNUAL = "11011"


class DartError(Exception):
    """OpenDART가 오류 상태코드를 반환했거나 응답을 해석할 수 없는 경우"""


def _require_key() -> str:
    if not DART_API_KEY:
        raise DartError("DART_API_KEY 환경변수가 설정되어 있지 않습니다.")
    return DART_API_KEY


def normalize_name(name: str) -> str:
    """회사명 비교용 정규화: 공백·법인격 표기 제거, 영문 소문자화"""
    name = re.sub(r"\(주\)|㈜|주식회사|\(유\)|유한회사", "", name)
    return re.sub(r"\s+", "", name).lower()


# ---------------------------------------------------------------------------
# 회사 고유번호 목록 (corpCode.xml)
# ---------------------------------------------------------------------------

_corp_list: list[dict] | None = None
_corp_list_loaded_at: float = 0.0


def parse_corp_code_xml(xml_bytes: bytes) -> list[dict]:
    root = ET.fromstring(xml_bytes)
    corps = []
    for item in root.iter("list"):
        corp = {
            "corp_code": (item.findtext("corp_code") or "").strip(),
            "corp_name": (item.findtext("corp_name") or "").strip(),
            "stock_code": (item.findtext("stock_code") or "").strip(),
            "modify_date": (item.findtext("modify_date") or "").strip(),
        }
        corp["norm"] = normalize_name(corp["corp_name"])
        corps.append(corp)
    return corps


async def load_corp_list(client: httpx.AsyncClient) -> list[dict]:
    """corpCode.xml을 내려받아 메모리와 디스크에 하루 동안 캐시"""
    global _corp_list, _corp_list_loaded_at

    now = time.time()
    if _corp_list is not None and now - _corp_list_loaded_at < CORP_CODE_TTL_SECONDS:
        return _corp_list

    cache_file = CACHE_DIR / "CORPCODE.xml"
    if cache_file.exists() and now - cache_file.stat().st_mtime < CORP_CODE_TTL_SECONDS:
        xml_bytes = cache_file.read_bytes()
    else:
        resp = await client.get(f"{DART_BASE_URL}/corpCode.xml", params={"crtfc_key": _require_key()})
        resp.raise_for_status()
        xml_bytes = _first_xml_in_zip(resp.content)
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache_file.write_bytes(xml_bytes)

    _corp_list = parse_corp_code_xml(xml_bytes)
    _corp_list_loaded_at = now
    return _corp_list


def search_corps(corps: list[dict], query: str, limit: int = 50) -> list[dict]:
    """회사명 부분일치 검색. 정확히 일치 > 앞부분 일치 > 상장사 > 이름 짧은 순"""
    q = normalize_name(query)
    if not q:
        return []
    hits = [c for c in corps if q in c["norm"]]
    hits.sort(key=lambda c: (c["norm"] != q, not c["norm"].startswith(q), not c["stock_code"], len(c["norm"])))
    return [{k: v for k, v in c.items() if k != "norm"} for c in hits[:limit]]


def resolve_corp(corps: list[dict], name: str) -> tuple[dict | None, list[dict]]:
    """
    회사명 하나로 회사 확정. (확정된 회사, 후보 목록)
    정확히 일치하는 회사가 1개이거나, 부분일치 결과가 1개뿐이면 확정. 그 외에는 후보만 반환
    """
    hits = search_corps(corps, name)
    q = normalize_name(name)
    exact = [h for h in hits if normalize_name(h["corp_name"]) == q]
    if len(exact) == 1:
        return exact[0], exact
    if not exact and len(hits) == 1:
        return hits[0], hits
    return None, exact or hits


def find_corp(corps: list[dict], corp_code: str) -> dict | None:
    return next((c for c in corps if c["corp_code"] == corp_code), None)


# ---------------------------------------------------------------------------
# 공시 목록 (list.json)
# ---------------------------------------------------------------------------

REPORT_NAME_PATTERN = re.compile(r"감사보고서|사업보고서")


async def list_reports(client: httpx.AsyncClient, corp_code: str, years: int = 10,
                       begin: date | None = None) -> list[dict]:
    """감사보고서(외부감사관련, F)와 사업보고서(정기공시, A) 목록. 정정 시 최종본만"""
    end = date.today()
    begin = begin or end - timedelta(days=365 * years)
    reports: list[dict] = []

    for pblntf_ty in ("A", "F"):
        page_no = 1
        while True:
            data = await _get_json(client, "list.json", {
                "corp_code": corp_code,
                "bgn_de": begin.strftime("%Y%m%d"),
                "end_de": end.strftime("%Y%m%d"),
                "pblntf_ty": pblntf_ty,
                "last_reprt_at": "Y",
                "page_no": page_no,
                "page_count": 100,
            })
            if data is None:
                break
            for item in data.get("list", []):
                report_nm = item.get("report_nm", "")
                if not REPORT_NAME_PATTERN.search(report_nm):
                    continue
                reports.append({
                    "rcept_no": item.get("rcept_no", ""),
                    "report_nm": report_nm.strip(),
                    "fiscal_year": year_from_report_name(report_nm),
                    "rcept_dt": item.get("rcept_dt", ""),
                    "flr_nm": item.get("flr_nm", ""),
                    "is_annual": "사업보고서" in report_nm,
                    "is_consolidated": "연결" in report_nm,
                })
            if page_no >= int(data.get("total_page", 1) or 1):
                break
            page_no += 1

    reports.sort(key=lambda r: r["rcept_dt"], reverse=True)
    return reports


# ---------------------------------------------------------------------------
# 공시 원문 (document.xml)
# ---------------------------------------------------------------------------

async def fetch_document_files(client: httpx.AsyncClient, rcept_no: str) -> list[tuple[str, str]]:
    """공시서류 원본 ZIP을 받아 (파일명, 텍스트) 목록으로 반환. 본문 파일이 먼저 오도록 정렬"""
    resp = await client.get(f"{DART_BASE_URL}/document.xml",
                            params={"crtfc_key": _require_key(), "rcept_no": rcept_no})
    resp.raise_for_status()
    content = resp.content
    if not content.startswith(b"PK"):
        raise DartError(_error_message_from_body(content))

    files = []
    with zipfile.ZipFile(io.BytesIO(content)) as zf:
        for name in sorted(zf.namelist(), key=lambda n: (len(n), n)):
            if not name.lower().endswith((".xml", ".html", ".htm")):
                continue
            files.append((name, decode_document(zf.read(name))))
    if not files:
        raise DartError("공시 원문 ZIP 안에 문서 파일이 없습니다.")
    return files


def decode_document(raw: bytes) -> str:
    """DART 원문은 대부분 UTF-8이나, 과거 공시는 EUC-KR(CP949)인 경우가 있음"""
    head = raw[:200].decode("ascii", errors="ignore").lower()
    if "euc-kr" in head or "ks_c_5601" in head:
        return raw.decode("cp949", errors="replace")
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp949", errors="replace")


# ---------------------------------------------------------------------------
# XBRL 재무제표 (fnlttSinglAcntAll.json) - 사업보고서 제출회사
# ---------------------------------------------------------------------------

async def fetch_xbrl_statements(client: httpx.AsyncClient, corp_code: str, bsns_year: int,
                                fs_div: str, try_previous: bool = True) -> tuple[int, list[dict]]:
    """
    fs_div: CFS(연결) / OFS(별도)
    결산월이 12월이 아닌 회사는 보고서명의 연도와 bsns_year가 다를 수 있어
    조회 결과가 없으면 전년도로 한 번 더 조회합니다. (사업연도 기준 연도 정의 확인 필요)
    반환값: (실제 조회된 사업연도, 계정 목록)
    """
    for year in ((bsns_year, bsns_year - 1) if try_previous else (bsns_year,)):
        data = await _get_json(client, "fnlttSinglAcntAll.json", {
            "corp_code": corp_code,
            "bsns_year": str(year),
            "reprt_code": REPRT_CODE_ANNUAL,
            "fs_div": fs_div,
        })
        if data is not None and data.get("list"):
            return year, data["list"]
    raise DartError("XBRL 재무제표 데이터가 없습니다. '원문 파싱'으로 시도해 보세요.")


def year_from_report_name(report_nm: str) -> int | None:
    """'사업보고서 (2024.12)' -> 2024"""
    m = re.search(r"\((\d{4})\.\d{2}\)", report_nm)
    return int(m.group(1)) if m else None


# ---------------------------------------------------------------------------
# 공통
# ---------------------------------------------------------------------------

async def _get_json(client: httpx.AsyncClient, endpoint: str, params: dict) -> dict | None:
    """정상이면 dict, 데이터 없음(013)이면 None, 그 외 상태코드는 DartError"""
    resp = await client.get(f"{DART_BASE_URL}/{endpoint}", params={"crtfc_key": _require_key(), **params})
    resp.raise_for_status()
    data = resp.json()
    status = data.get("status")
    if status == STATUS_OK:
        return data
    if status == STATUS_NO_DATA:
        return None
    raise DartError(f"OpenDART 오류 [{status}] {data.get('message', '')}")


def _first_xml_in_zip(content: bytes) -> bytes:
    if not content.startswith(b"PK"):
        raise DartError(_error_message_from_body(content))
    with zipfile.ZipFile(io.BytesIO(content)) as zf:
        for name in zf.namelist():
            if name.lower().endswith(".xml"):
                return zf.read(name)
    raise DartError("ZIP 안에 XML 파일이 없습니다.")


def _error_message_from_body(content: bytes) -> str:
    """ZIP 대신 오류 XML/JSON이 온 경우 메시지 추출"""
    text = content[:2000].decode("utf-8", errors="replace")
    status = re.search(r"<status>(.*?)</status>|\"status\"\s*:\s*\"(\d+)\"", text)
    message = re.search(r"<message>(.*?)</message>|\"message\"\s*:\s*\"(.*?)\"", text)
    s = next((g for g in status.groups() if g), "") if status else ""
    m = next((g for g in message.groups() if g), "") if message else text[:200]
    return f"OpenDART 오류 [{s}] {m}".strip()
