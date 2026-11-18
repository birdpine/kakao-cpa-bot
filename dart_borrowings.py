"""
재무상태표에서 차입금 집계

포함 계정 (계정명 기준, 앞의 번호·공백·주석표시 제거 후 비교):
  단기차입금, 유동성장기부채(유동성장기차입금), 유동성사채, 사채, 장기차입금, 전환사채

조정 계정:
  - 차감: 현재가치할인차금, 사채할인발행차금, 전환권조정, 신주인수권조정
  - 가산: 사채할증발행차금, (사채)상환할증금
  → 바로 위의 차입금 계정(그 사이에 다른 조정 계정만 있는 경우 포함)에 붙여 순액을 만든다.
    원문에 괄호 없이 양수로 적혀 있어도 차감 계정은 음수로 처리한다.
  → 바로 위가 차입금 계정이 아니면(예: 임대보증금 현재가치할인차금) 집계에서 제외하고 명세에 표시한다.

참고로 명세에만 표시하고 집계에서 제외: 신주인수권부사채, 교환사채, 리스부채
(포함 범위가 바뀌면 CATEGORY_RULES만 고치면 된다)

금액은 재무상태표의 당기말 열에서 가져온다.
"""

import re
from dataclasses import dataclass

CATEGORIES = ["단기차입금", "유동성장기부채", "유동성사채", "사채", "장기차입금", "전환사채"]

DEDUCTION = "차감"
ADDITION = "가산"
EXCLUDED = "미포함"

# (분류, 패턴) - 위에서부터 먼저 일치하는 것 적용
CATEGORY_RULES = [
    (DEDUCTION, re.compile(r"^(현재가치할인차금|사채할인발행차금|전환권조정|신주인수권조정)")),
    (ADDITION, re.compile(r"^(사채할증발행차금|사채상환할증금|상환할증금)")),
    ("유동성사채", re.compile(r"^유동성(전환)?사채")),
    ("유동성장기부채", re.compile(r"^유동성장기(부채|차입금)")),
    ("단기차입금", re.compile(r"^단기차입금")),
    ("장기차입금", re.compile(r"^장기차입금")),
    ("전환사채", re.compile(r"^전환사채")),
    ("사채", re.compile(r"^(일반|공모|사모|무보증|보증)?사채$|^회사채$")),
    (EXCLUDED, re.compile(r"^(신주인수권부사채|교환사채|유동성리스부채|리스부채)")),
]

# 계정명 앞 번호: Ⅰ. / 1. / (1) / 1) / 가. / ① 등
_PREFIX = re.compile(r"^([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+|\d+|[가-하]|[①-⑳])[.)]\s*|^\(\s*(\d+|[가-하])\s*\)\s*")
# 계정명 뒤 주석 표시: (주석5), (주5,12), (주 5) 등
_NOTE_SUFFIX = re.compile(r"\(\s*주(석)?\s*[\d,\s]+\)$")


@dataclass
class BorrowingLine:
    category: str      # CATEGORIES 중 하나, 또는 EXCLUDED
    account: str       # 원문 계정명
    kind: str          # 차입금 / 차감 / 가산 / 미포함 사유
    amount: float      # 원 단위, 집계 부호 반영 (차감은 음수)


def normalize_account(name: str) -> str:
    s = re.sub(r"\s+", "", str(name).replace("\xa0", ""))
    s = _PREFIX.sub("", s)
    s = _NOTE_SUFFIX.sub("", s)
    return s


def classify(name: str) -> str | None:
    n = normalize_account(name)
    for category, pattern in CATEGORY_RULES:
        if pattern.search(n):
            return category
    return None


def current_period_cols(rows: list[list], header_rows: int) -> list[int]:
    """당기말 금액이 들어 있는 열 번호들 (상세 열·합계 열로 나뉜 양식이면 여러 개)"""
    if rows and [str(c) for c in rows[0][:2]] == ["계정명", "계정ID"]:
        return [3] if rows[0][2] == "자본구성요소" else [2]          # XBRL 표

    width = max((len(r) for r in rows), default=0)
    header = rows[:header_rows]
    current = _header_cols(header, r"\(\s*당\s*\)|당\s*기")
    prior = _header_cols(header, r"\(\s*전\s*\)|전\s*기")
    if not current:
        # '제 10 기 / 제 9 기'처럼 당·전 표시가 없으면 첫 번째 기수가 당기
        periods = _header_cols(header, r"제\s*\d+\s*기")
        current, prior = periods[:1], periods[1:2]
    if current:
        start = current[0]
        later = [c for c in prior if c > start]
        end = later[0] if later else width
        return list(range(start, end))

    # 머리글로 못 찾으면 숫자 열을 반으로 나눠 앞쪽을 당기로
    numeric = sorted({c for r in rows[header_rows:] for c, v in enumerate(r) if _is_number(v)})
    return numeric[: max(1, len(numeric) // 2)]


def _header_cols(header: list[list], pattern: str) -> list[int]:
    cols = set()
    for row in header:
        for c, v in enumerate(row):
            if c > 0 and isinstance(v, str) and re.search(pattern, v):
                cols.add(c)
    return sorted(cols)


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _current_amount(row: list, cols: list[int]):
    for c in cols:
        if c < len(row) and _is_number(row[c]):
            return row[c]
    return None


def extract_borrowings(rows: list[list], header_rows: int) -> list[BorrowingLine]:
    """재무상태표 표(원 단위)에서 차입금 관련 행 추출"""
    cols = current_period_cols(rows, header_rows)
    lines: list[BorrowingLine] = []
    attach_to: str | None = None    # 조정 계정을 붙일 직전 차입금 분류

    for row in rows[header_rows:]:
        if not row:
            continue
        account = str(row[0]).strip()
        amount = _current_amount(row, cols)
        category = classify(account)

        if category in CATEGORIES:
            attach_to = category
            if amount is not None:
                lines.append(BorrowingLine(category, account, "차입금", amount))
        elif category in (DEDUCTION, ADDITION):
            if amount is None:
                continue
            signed = -abs(amount) if category == DEDUCTION else abs(amount)
            if attach_to:
                lines.append(BorrowingLine(attach_to, account, category, signed))
            else:
                lines.append(BorrowingLine(EXCLUDED, account, "미포함(차입금 조정계정 아님)", signed))
        elif category == EXCLUDED:
            attach_to = None
            if amount is not None:
                lines.append(BorrowingLine(EXCLUDED, account, "미포함(포함 범위 외)", amount))
        elif amount is not None:
            # 금액이 있는 다른 계정이 나오면 조정 계정 연결을 끊음 (금액 없는 소제목 행은 무시)
            attach_to = None
    return lines
