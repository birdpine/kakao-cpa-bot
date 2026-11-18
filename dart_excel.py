"""
재무제표 엑셀(xlsx) 생성

  - 시트 순서: 정보 → 재무상태표 → (포괄)손익계산서 → 자본변동표 → 현금흐름표 → 이익잉여금처분계산서
  - 금액 단위는 천원(기본) 또는 원. 천원은 원 금액 ÷ 1,000을 셀에 그대로 두고 표시만 정수로 하므로
    직접 합계를 내도 단수차이가 없음. 주당이익 행은 항상 원 단위
  - 천 단위 구분기호, 음수는 괄호, 0은 '-'
  - 머리글 행은 가운데 정렬·굵게
"""

import io
from dataclasses import dataclass, field
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from dart_borrowings import CATEGORIES, EXCLUDED, BorrowingLine, extract_borrowings
from dart_parser import STATEMENT_ORDER, ParseResult, Statement, sheet_name_for

AMOUNT_FORMAT = '#,##0;(#,##0);"-"'
# 출력 금액 단위 -> 원 금액을 나눌 값
UNIT_DIVISORS = {"원": 1, "천원": 1_000}
DEFAULT_UNIT = "천원"
DECIMAL_FORMAT = '#,##0.##;(#,##0.##);"-"'
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
HEADER_FONT = Font(bold=True)
TITLE_FONT = Font(bold=True, size=14)
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

XBRL_SHEETS = [
    ("BS", "재무상태표"),
    ("IS", "손익계산서"),
    ("CIS", "포괄손익계산서"),
    ("SCE", "자본변동표"),
    ("CF", "현금흐름표"),
]


def _info_sheet(wb: Workbook, meta: dict, notes: list[str]) -> None:
    ws = wb.active
    ws.title = "정보"
    ws["A1"] = "DART 재무제표 추출"
    ws["A1"].font = TITLE_FONT
    rows = [
        ("회사명", meta.get("corp_name", "")),
        ("고유번호", meta.get("corp_code", "")),
        ("종목코드", meta.get("stock_code", "") or "비상장"),
        ("보고서명", meta.get("report_nm", "")),
        ("접수번호", meta.get("rcept_no", "")),
        ("접수일자", meta.get("rcept_dt", "")),
        ("추출방식", meta.get("method", "")),
        ("DART 원문", f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={meta.get('rcept_no', '')}"),
        ("생성일시", datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]
    for i, (k, v) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=k).font = HEADER_FONT
        cell = ws.cell(row=i, column=2, value=v)
        if k == "DART 원문":
            cell.hyperlink = v
            cell.style = "Hyperlink"

    r = len(rows) + 4
    ws.cell(row=r, column=1, value="확인사항").font = HEADER_FONT
    for note in notes:
        r += 1
        ws.cell(row=r, column=1, value="•")
        ws.cell(row=r, column=2, value=note)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 90


def _write_grid(ws: Worksheet, rows: list[list], header_rows: int, start_row: int, divisor: int = 1) -> None:
    """divisor: 원 금액을 나눌 값 (천원 출력이면 1,000). 주당이익 행은 나누지 않음"""
    width = max((len(r) for r in rows), default=0)
    for r_idx, row in enumerate(rows):
        excel_row = start_row + r_idx
        scale = r_idx >= header_rows and divisor != 1 and not _is_per_share_row(row)
        for c_idx in range(width):
            value = row[c_idx] if c_idx < len(row) else ""
            if scale and _is_number(value):
                value = _divide(value, divisor)
                fmt = AMOUNT_FORMAT
            elif _is_number(value):
                fmt = DECIMAL_FORMAT if isinstance(value, float) else AMOUNT_FORMAT
            cell = ws.cell(row=excel_row, column=c_idx + 1, value=value if value != "" else None)
            cell.border = BORDER
            if r_idx < header_rows:
                cell.font = HEADER_FONT
                cell.fill = HEADER_FILL
                cell.alignment = CENTER
            elif _is_number(value):
                cell.number_format = fmt

    ws.column_dimensions["A"].width = 42
    for c in range(2, width + 1):
        ws.column_dimensions[get_column_letter(c)].width = 20
    ws.freeze_panes = ws.cell(row=start_row + header_rows, column=2)


def _is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _is_per_share_row(row: list) -> bool:
    """주당이익(손실) 행: 재무제표 단위와 무관하게 원 단위로 표시됨"""
    return bool(row) and "주당" in str(row[0]).replace(" ", "")


def _divide(value, divisor: int):
    q = value / divisor
    return int(q) if q == int(q) else q


def unit_note(source_unit: str, unit: str) -> str:
    note = f"(단위: {unit})"
    if source_unit != unit:
        note += f"  ※ 원문 단위 {source_unit} → {unit} 환산"
    return note


def unit_notice(unit: str) -> str:
    if unit == "원":
        return "금액은 원 단위입니다."
    return (f"금액은 {unit} 단위입니다. 셀 값은 원 금액 ÷ {UNIT_DIVISORS[unit]:,}로 소수까지 보존하고 "
            "표시만 정수로 반올림하므로, 직접 합계를 내도 단수차이가 생기지 않습니다. 주당이익 행은 원 단위입니다.")


def build_workbook_from_document(meta: dict, result: ParseResult, unit: str = DEFAULT_UNIT) -> bytes:
    wb = Workbook()
    notes = [
        unit_notice(unit),
        "원문 표를 자동으로 읽은 결과이므로, 합계 검증 및 DART 원문과의 대사를 권장합니다.",
    ]

    found = [st for cons in (True, False) for k in STATEMENT_ORDER if (st := result.get(k, cons))]
    for st in found:
        _write_table(wb.create_sheet(sheet_name_for(st.key, st.title)), table_from_statement(st), unit)
        if not st.unit_found:
            notes.append(f"{st.title}: 원문에서 단위 표시를 찾지 못해 '원'으로 가정했습니다. 원문 확인 필요.")

    for cons in sorted({st.consolidated for st in found}, reverse=True):
        missing = [name for k, name in (("BS", "재무상태표"), ("SCE", "자본변동표"), ("CF", "현금흐름표"))
                   if not result.get(k, cons)]
        if not result.income_statement(cons):
            missing.insert(1, "손익계산서")
        if missing:
            prefix = "연결" if cons else ""
            notes.append("찾지 못한 재무제표: " + ", ".join(prefix + m for m in missing))

    if not found:
        notes.append("재무제표 표를 식별하지 못해 원문의 숫자 표를 '원문표' 시트에 원문 단위 그대로 실었습니다.")
        for i, grid in enumerate(result.raw_tables[:30], start=1):
            ws = wb.create_sheet(f"원문표{i}")
            _write_grid(ws, grid, 0, start_row=1)

    _info_sheet(wb, {**meta, "method": "공시 원문(감사보고서) 표 파싱"}, notes)
    return _to_bytes(wb)


@dataclass
class SheetTable:
    """시트 하나에 쓸 표 (원문 파싱·XBRL 공통)"""
    title: str
    rows: list[list]
    header_rows: int
    source_unit: str = "원"           # 원문(공시) 단위. rows의 금액은 이미 원 단위로 환산되어 있음
    period_texts: list[str] = field(default_factory=list)


def table_from_statement(st: Statement) -> SheetTable:
    return SheetTable(st.title, st.rows, st.header_row_count, st.unit, st.period_texts[:3])


def table_from_xbrl(items: list[dict], sj_div: str, fs_div: str) -> SheetTable | None:
    part = sorted((it for it in items if it.get("sj_div") == sj_div), key=lambda it: int(it.get("ord") or 0))
    if not part:
        return None
    first = part[0]
    headers = ["계정명", "계정ID"]
    if sj_div == "SCE":
        headers.append("자본구성요소")
    period_names = [first.get("thstrm_nm") or "당기", first.get("frmtrm_nm") or "전기"]
    if first.get("bfefrmtrm_nm"):
        period_names.append(first["bfefrmtrm_nm"])
    headers += period_names

    rows: list[list] = [headers]
    for it in part:
        row = [it.get("account_nm", ""), it.get("account_id", "")]
        if sj_div == "SCE":
            row.append(it.get("account_detail", ""))
        amounts = [it.get("thstrm_amount"), it.get("frmtrm_amount"), it.get("bfefrmtrm_amount")]
        row += [_xbrl_amount(a) for a in amounts[: len(period_names)]]
        rows.append(row)
    name = dict(XBRL_SHEETS)[sj_div]
    return SheetTable(("연결" if fs_div == "CFS" else "") + name, rows, 1)


def xbrl_income_statement(items: list[dict], fs_div: str) -> SheetTable | None:
    """손익계산서가 없으면(단일 포괄손익계산서 양식) 포괄손익계산서"""
    return table_from_xbrl(items, "IS", fs_div) or table_from_xbrl(items, "CIS", fs_div)


def _write_table(ws: Worksheet, table: SheetTable, unit: str) -> None:
    ws["A1"] = table.title
    ws["A1"].font = TITLE_FONT
    line = 2
    for text in table.period_texts:
        ws.cell(row=line, column=1, value=text)
        line += 1
    ws.cell(row=line, column=1, value=unit_note(table.source_unit, unit))
    _write_grid(ws, table.rows, table.header_rows, start_row=line + 2, divisor=UNIT_DIVISORS[unit])
    if table.rows and table.rows[0][:2] == ["계정명", "계정ID"]:
        ws.column_dimensions["B"].width = 30


def build_workbook_from_xbrl(meta: dict, items: list[dict], bsns_year: int, fs_div: str,
                            unit: str = DEFAULT_UNIT) -> bytes:
    """fnlttSinglAcntAll 응답을 재무제표별 시트로"""
    wb = Workbook()
    fs_label = "연결" if fs_div == "CFS" else "별도"
    notes = [f"OpenDART XBRL 재무제표({fs_label}, 사업연도 {bsns_year}) 기준입니다.", unit_notice(unit)]
    for sj_div, name in XBRL_SHEETS:
        table = table_from_xbrl(items, sj_div, fs_div)
        if table:
            _write_table(wb.create_sheet(f"{fs_label}{name}"), table, unit)
    _info_sheet(wb, {**meta, "method": f"OpenDART XBRL 재무제표 ({fs_label})"}, notes)
    return _to_bytes(wb)


@dataclass
class YearResult:
    year: int
    report_nm: str = ""
    rcept_no: str = ""
    method: str = ""
    bs: SheetTable | None = None
    pl: SheetTable | None = None
    notes: list[str] = field(default_factory=list)


def build_annual_workbook(meta: dict, results: list[YearResult], fs_div: str, unit: str = DEFAULT_UNIT) -> bytes:
    """연도별 BS(2025), PL(2025) 시트. 첫 시트는 연도별 출처 요약"""
    wb = Workbook()
    ws = wb.active
    ws.title = "정보"
    fs_label = "연결" if fs_div == "CFS" else "별도"
    ws["A1"] = f"{meta.get('corp_name', '')} 연도별 재무제표 ({fs_label} 우선)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = (f"고유번호 {meta.get('corp_code', '')} · 종목코드 {meta.get('stock_code') or '비상장'} · "
                f"생성일시 {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    ws["A3"] = unit_notice(unit) + " 원문 파싱 결과는 합계 검증 및 DART 원문과의 대사를 권장합니다."

    headers = ["연도", "보고서", "접수번호", "추출방식", "시트", "확인사항"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=c, value=h)
        cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, HEADER_FILL, CENTER, BORDER

    for i, res in enumerate(sorted(results, key=lambda r: r.year, reverse=True), start=6):
        sheets = []
        for prefix, table in (("BS", res.bs), ("PL", res.pl)):
            if table:
                name = f"{prefix}({res.year})"
                _write_table(wb.create_sheet(name), table, unit)
                sheets.append(name)
        values = [res.year, res.report_nm or "해당 연도 보고서 없음", res.rcept_no, res.method,
                  ", ".join(sheets), " / ".join(res.notes)]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        if res.rcept_no:
            link = ws.cell(row=i, column=3)
            link.hyperlink = f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={res.rcept_no}"
            link.style = "Hyperlink"
            link.border = BORDER

    for col, width in zip("ABCDEF", (8, 30, 18, 26, 22, 60)):
        ws.column_dimensions[col].width = width

    _write_borrowings_sheet(wb, results, unit)
    return _to_bytes(wb)


BORROWINGS_SHEET = "차입금"


def _write_borrowings_sheet(wb: Workbook, results: list[YearResult], unit: str) -> None:
    """
    연도별 차입금 집계(위)와 원문 계정 명세(아래).
    집계 셀은 명세를 SUMIFS로 합산하는 수식이라 엑셀에서 바로 검증 가능
    """
    years = [r.year for r in sorted(results, key=lambda r: r.year, reverse=True) if r.bs]
    detail: list[tuple[int, BorrowingLine]] = []
    for res in sorted(results, key=lambda r: r.year, reverse=True):
        if res.bs:
            detail += [(res.year, line) for line in extract_borrowings(res.bs.rows, res.bs.header_rows)]

    ws = wb.create_sheet(BORROWINGS_SHEET, index=1)
    ws["A1"] = "차입금 집계 (재무상태표 당기말 기준)"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"(단위: {unit})"
    ws["A3"] = ("포함: " + ", ".join(CATEGORIES) + ". 현재가치할인차금·사채할인발행차금·전환권조정은 바로 위 "
                "차입금 계정에서 차감, 사채상환할증금·할증발행차금은 가산. 각 금액은 아래 명세의 SUMIFS")
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=max(4, len(years) + 1))
    ws.row_dimensions[3].height = 45

    divisor = UNIT_DIVISORS[unit]
    summary_header_row = 5
    first_cat_row = summary_header_row + 1
    total_row = first_cat_row + len(CATEGORIES)
    detail_header_row = total_row + 3
    first_detail_row = detail_header_row + 1
    last_detail_row = max(first_detail_row, first_detail_row + len(detail) - 1)

    def rng(col: str) -> str:
        return f"${col}${first_detail_row}:${col}${last_detail_row}"

    _header_cells(ws, summary_header_row, ["구분"] + years)
    for i, category in enumerate(CATEGORIES):
        r = first_cat_row + i
        ws.cell(row=r, column=1, value=category).border = BORDER
        for j, year in enumerate(years, start=2):
            cell = ws.cell(row=r, column=j,
                           value=f'=SUMIFS({rng("E")},{rng("A")},{year},{rng("B")},"{category}")')
            cell.number_format, cell.border = AMOUNT_FORMAT, BORDER

    ws.cell(row=total_row, column=1, value="차입금 합계").font = HEADER_FONT
    ws.cell(row=total_row, column=1).border = BORDER
    for j in range(2, len(years) + 2):
        col = get_column_letter(j)
        cell = ws.cell(row=total_row, column=j, value=f"=SUM({col}{first_cat_row}:{col}{total_row - 1})")
        cell.number_format, cell.border, cell.font = AMOUNT_FORMAT, BORDER, HEADER_FONT

    ws.cell(row=detail_header_row - 1, column=1, value="명세 (원문 재무상태표 계정)").font = HEADER_FONT
    _header_cells(ws, detail_header_row, ["연도", "분류", "원문 계정명", "처리", f"금액({unit})"])
    for i, (year, line) in enumerate(detail):
        r = first_detail_row + i
        values = [year, line.category, line.account, line.kind, _divide(line.amount, divisor)]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
        ws.cell(row=r, column=5).number_format = AMOUNT_FORMAT
        if line.category == EXCLUDED:
            for c in range(1, 6):
                ws.cell(row=r, column=c).font = Font(color="888888")
    if not detail:
        ws.cell(row=first_detail_row, column=1, value="재무상태표에서 차입금 계정을 찾지 못했습니다.")

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 26
    ws.column_dimensions["E"].width = 18
    for j in range(2, len(years) + 2):
        letter = get_column_letter(j)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, 18)


def _header_cells(ws: Worksheet, row: int, values: list) -> None:
    for c, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, HEADER_FILL, CENTER, BORDER


def _xbrl_amount(value: str | None):
    if value is None:
        return ""
    v = str(value).replace(",", "").strip()
    if v in ("", "-"):
        return ""
    try:
        return int(v)
    except ValueError:
        try:
            return float(v)
        except ValueError:
            return value


def _to_bytes(wb: Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
