"""
DART 공시 원문(XML/HTML)에서 재무제표 표를 찾아 구조화

원문 구조(감사보고서 기준, 대체로 아래 순서):
  [제목 표/문단] "재 무 상 태 표", "제 10(당) 기 2024년 12월 31일 현재", "(단위 : 원)"
  [본표]        과목 | 주석 | 제10(당)기 | 제9(전)기 ...

파싱 원칙:
  - 짧은 문구에서 재무제표명을 찾으면 그 다음의 '큰 표'(데이터 표)를 해당 재무제표로 본다.
    (감사의견 문단처럼 긴 문장 안의 '재무상태표'는 무시)
  - 재무제표별로 최초 1회만 채택한다. (사업보고서는 본문과 첨부 감사보고서에 중복 수록됨)
  - 금액은 표 위 '(단위 : 천원)' 등을 읽어 원 단위로 환산한다. 주당이익 행은 환산하지 않는다.
  - 주석번호 열, 머리글 행, 과목 열은 숫자로 바꾸지 않는다.
"""

import re
import warnings
from dataclasses import dataclass, field

from bs4 import BeautifulSoup, Tag, XMLParsedAsHTMLWarning


# DART XML 셀 태그: TD(일반), TH(머리글), TE(숫자), TU(머리글/구분)
CELL_TAGS = ("td", "th", "te", "tu")
HEADER_CELL_TAGS = ("th", "tu")

# (키, 정규화된 제목 패턴, 시트명). 위에서부터 검사 - 포괄손익계산서를 손익계산서보다 먼저
STATEMENT_PATTERNS = [
    ("BS", re.compile(r"재무상태표|대차대조표"), "재무상태표"),
    ("CIS", re.compile(r"포괄손익계산서"), "포괄손익계산서"),
    ("IS", re.compile(r"손익계산서"), "손익계산서"),
    ("SCE", re.compile(r"자본변동표"), "자본변동표"),
    ("CF", re.compile(r"현금흐름표"), "현금흐름표"),
    ("RE", re.compile(r"이익잉여금처분계산서|결손금처리계산서|이익잉여금처분\(안\)"), "이익잉여금처분계산서"),
]
STATEMENT_ORDER = [key for key, _, _ in STATEMENT_PATTERNS]

# 제목으로 인정할 문구의 최대 길이(공백 제거 후). 감사의견 문단 등 긴 문장 배제
TITLE_MAX_LEN = 30

# 주석 안의 표 제목 배제: 요약재무정보, 종속·관계기업 재무정보, 주석식 번호 '(2)', '1)', '가.'
# 사업보고서 본문의 '2-1. 연결 재무상태표' 같은 번호는 허용
NOTE_TITLE_PATTERN = re.compile(r"요약|기업|^\(?\d+\)|^\([가-힣a-z]\)|^[가-힣]\.")

UNIT_PATTERN = re.compile(r"단위\s*[:：]?\s*(백만원|천원|원)")
UNIT_MULTIPLIER = {"원": 1, "천원": 1_000, "백만원": 1_000_000}

NUMBER_PATTERN = re.compile(r"^(\()?\s*([-△▲(]?)\s*(\d{1,3}(?:,\d{3})+|\d+)(\.\d+)?\s*(\))?$")
DASH_VALUES = {"-", "－", "—", "–", "―", "‐"}

# 데이터 표로 볼 최소 조건
MIN_DATA_ROWS = 4
MIN_NUMERIC_CELLS = 3


@dataclass
class Statement:
    key: str
    title: str                 # 원문 제목 (예: 연결재무상태표)
    unit: str                  # 원문 단위
    unit_found: bool           # 단위 문구를 실제로 찾았는지 (못 찾으면 '원'으로 가정)
    period_texts: list[str]    # 제목 주변의 기간 문구 (제 10(당) 기 ...)
    rows: list[list]           # 셀 값: str 또는 int/float (원 단위 환산 완료)
    header_row_count: int
    numeric_cols: list[int] = field(default_factory=list)

    @property
    def consolidated(self) -> bool:
        return self.title.startswith("연결")


def statement_id(key: str, consolidated: bool) -> str:
    """ParseResult.statements의 키. 사업보고서는 연결·별도 재무제표가 함께 있으므로 구분"""
    return f"C_{key}" if consolidated else key


@dataclass
class ParseResult:
    statements: dict[str, Statement]    # statement_id -> Statement
    raw_tables: list[list[list[str]]]   # 재무제표를 하나도 못 찾았을 때 엑셀에 그대로 싣는 원문 표

    def get(self, key: str, consolidated: bool) -> Statement | None:
        return self.statements.get(statement_id(key, consolidated))

    def income_statement(self, consolidated: bool) -> Statement | None:
        """손익계산서가 없으면(단일 포괄손익계산서 양식) 포괄손익계산서"""
        return self.get("IS", consolidated) or self.get("CIS", consolidated)


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def compact(text: str) -> str:
    return re.sub(r"\s+", "", text.replace("\xa0", ""))


def detect_statement(text: str) -> tuple[str, str] | None:
    """짧은 제목 문구면 (키, 원문 제목) 반환"""
    c = compact(text)
    if not c or len(c) > TITLE_MAX_LEN:
        return None
    # 목차, 주석 참조 등 배제
    if "주석" in c or "목차" in c or c.endswith("참조") or NOTE_TITLE_PATTERN.search(c):
        return None
    for key, pattern, _ in STATEMENT_PATTERNS:
        m = pattern.search(c)
        if m:
            prefix = "연결" if "연결" in c[: m.start() + 1] else ""
            return key, prefix + m.group(0)
    return None


def detect_unit(text: str) -> str | None:
    m = UNIT_PATTERN.search(compact(text))
    return m.group(1) if m else None


def parse_number(text: str) -> float | int | None:
    """'1,234' / '(1,234)' / '△1,234' / '-' -> 숫자. 숫자가 아니면 None"""
    t = compact(text)
    if not t:
        return None
    if t in DASH_VALUES:
        return 0
    m = NUMBER_PATTERN.match(t)
    if not m:
        return None
    open_paren, sign, int_part, frac, close_paren = m.groups()
    if (open_paren or sign == "(") and not close_paren:
        return None
    if close_paren and not (open_paren or sign == "("):
        return None
    digits = int_part.replace(",", "")
    value: float | int = float(digits + frac) if frac else int(digits)
    negative = bool(open_paren) or sign in ("-", "△", "▲", "(")
    return -value if negative else value


def _own_rows(table: Tag) -> list[Tag]:
    """중첩 표의 행은 제외하고 이 표에 속한 행만"""
    return [tr for tr in table.find_all("tr") if tr.find_parent("table") is table]


def _own_cells(tr: Tag) -> list[Tag]:
    return [c for c in tr.find_all(CELL_TAGS) if c.find_parent("tr") is tr]


def table_to_grid(table: Tag) -> tuple[list[list[str]], list[bool]]:
    """COLSPAN/ROWSPAN을 펼친 텍스트 격자와, 행별 머리글 태그(TH/TU) 포함 여부"""
    grid: list[list[str]] = []
    header_flags: list[bool] = []
    pending: dict[tuple[int, int], str] = {}   # rowspan으로 아래 행에 채울 칸

    for r, tr in enumerate(_own_rows(table)):
        row: list[str] = []
        col = 0
        is_header = False

        def fill_pending():
            nonlocal col
            while (r, col) in pending:
                row.append(pending.pop((r, col)))
                col += 1

        for cell in _own_cells(tr):
            fill_pending()
            text = clean_text(cell.get_text(" "))
            if cell.name in HEADER_CELL_TAGS:
                is_header = True
            colspan = _span(cell, "colspan")
            rowspan = _span(cell, "rowspan")
            for i in range(colspan):
                # 병합 셀은 첫 칸에만 텍스트를 두고 나머지는 빈칸
                row.append(text if i == 0 else "")
                for k in range(1, rowspan):
                    pending[(r + k, col)] = ""
                col += 1
        fill_pending()
        grid.append(row)
        header_flags.append(is_header)

    width = max((len(row) for row in grid), default=0)
    for row in grid:
        row.extend([""] * (width - len(row)))
    return grid, header_flags


def _span(cell: Tag, attr: str) -> int:
    try:
        return max(1, min(int(cell.get(attr, 1)), 100))
    except (TypeError, ValueError):
        return 1


def is_data_table(grid: list[list[str]]) -> bool:
    if len(grid) < MIN_DATA_ROWS:
        return False
    numeric = sum(1 for row in grid for cell in row[1:] if cell and parse_number(cell) is not None
                  and compact(cell) not in DASH_VALUES)
    return numeric >= MIN_NUMERIC_CELLS


def count_header_rows(grid: list[list[str]], header_flags: list[bool]) -> int:
    """위에서부터 이어지는 머리글 행 수: TH/TU 행이거나, 과목/구분/기간 문구 행"""
    count = 0
    for row, flagged in zip(grid, header_flags):
        if not (flagged or _looks_like_header(row)):
            break
        count += 1
    return count


def _looks_like_header(row: list[str]) -> bool:
    first = row[0] if row else ""
    if detect_statement(first) or detect_unit(first):
        return True
    if compact(first) not in ("과목", "구분", "계정과목", ""):
        return False
    if any(_is_amount(c) for c in row[1:]):
        return False
    return re.search(r"과목|구분|제\d+|당기|전기|기말|기초|\d{4}", compact("".join(row))) is not None


def _is_amount(text: str) -> bool:
    """머리글의 연도(2024)와 구분되는 금액 셀: 천단위 구분기호나 괄호가 있는 숫자"""
    return parse_number(text) is not None and ("," in text or "(" in text)


def find_note_cols(grid: list[list[str]], header_rows: int) -> set[int]:
    cols = set()
    for row in grid[: max(header_rows, 1)]:
        for i, cell in enumerate(row):
            if compact(cell) in ("주석", "주석번호", "주"):
                cols.add(i)
    return cols


def convert_rows(grid: list[list[str]], header_rows: int, multiplier: int) -> tuple[list[list], list[int]]:
    note_cols = find_note_cols(grid, header_rows)
    numeric_cols: set[int] = set()
    rows: list[list] = []
    for r, row in enumerate(grid):
        if r < header_rows:
            rows.append(list(row))
            continue
        # 주당이익은 표 단위와 무관하게 원 단위로 표시됨
        row_multiplier = 1 if "주당" in compact(row[0] if row else "") else multiplier
        out: list = []
        for c, cell in enumerate(row):
            value = None if c == 0 or c in note_cols else parse_number(cell)
            if value is None:
                out.append(cell)
            else:
                numeric_cols.add(c)
                out.append(value * row_multiplier)
        rows.append(out)
    return rows, sorted(numeric_cols)


def _block_texts(el: Tag) -> list[str]:
    """제목 후보로 쓸 짧은 문구들"""
    if el.name == "table":
        grid, _ = table_to_grid(el)
        return [cell for row in grid for cell in row if cell]
    return [clean_text(el.get_text(" "))]


def parse_document(html_texts: list[str]) -> ParseResult:
    statements: dict[str, Statement] = {}
    raw_tables: list[list[list[str]]] = []

    current: tuple[str, str] | None = None
    unit: str | None = None
    periods: list[str] = []

    for html in html_texts:
        # DART 원문은 XML이지만 태그가 HTML 표와 같고 형식 오류가 잦아 관대한 HTML 파서로 읽음
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", XMLParsedAsHTMLWarning)
            soup = BeautifulSoup(html, "lxml")
        for el in soup.find_all(["table", "p", "title", "span"]):
            # 표 안의 문단은 표를 처리할 때 함께 읽음
            if el.name != "table" and el.find_parent("table") is not None:
                continue
            if el.name == "span" and el.find_parent("p") is not None:
                continue
            if el.name == "table" and el.find_parent("table") is not None:
                continue

            if el.name == "table":
                grid, header_flags = table_to_grid(el)
                if is_data_table(grid):
                    raw_tables.append(grid)
                    # 본표 첫 행에 제목/단위가 들어 있는 양식도 있음
                    for cell in (c for row in grid[:3] for c in row if c):
                        found = detect_statement(cell)
                        if found and current is None:
                            current, unit, periods = found, None, []
                        unit = unit or detect_unit(cell)
                    if current:
                        sid = statement_id(current[0], current[1].startswith("연결"))
                        if sid not in statements:
                            statements[sid] = _build_statement(current, unit, periods, grid, header_flags)
                    current, unit, periods = None, None, []
                    continue

            texts = _block_texts(el)
            for text in texts:
                found = detect_statement(text)
                if found:
                    current, unit, periods = found, None, []
            for text in texts:
                unit = detect_unit(text) or unit
                if current and re.search(r"제\s*\d+\s*(\(\s*[당전]\s*\))?\s*기", text) and len(text) < 80:
                    periods.append(text)

    return ParseResult(statements=statements, raw_tables=raw_tables)


def _build_statement(current: tuple[str, str], unit: str | None, periods: list[str],
                     grid: list[list[str]], header_flags: list[bool]) -> Statement:
    key, title = current
    header_rows = count_header_rows(grid, header_flags)
    rows, numeric_cols = convert_rows(grid, header_rows, UNIT_MULTIPLIER[unit or "원"])
    return Statement(
        key=key,
        title=title,
        unit=unit or "원",
        unit_found=unit is not None,
        period_texts=periods,
        rows=rows,
        header_row_count=header_rows,
        numeric_cols=numeric_cols,
    )


def sheet_name_for(key: str, title: str) -> str:
    base = next(name for k, _, name in STATEMENT_PATTERNS if k == key)
    return ("연결" + base) if title.startswith("연결") else base
