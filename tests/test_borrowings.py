import io
import re

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook

import dart_client
from dart_borrowings import classify, current_period_cols, extract_borrowings
from dart_parser import parse_document
from kakao_tax_skill_server import app
from tests.test_dart import SAMPLE, _patch_dart, _report

BORROWING_BS = """
<P>재 무 상 태 표</P><P>(단위 : 천원)</P>
<TABLE BORDER="1"><TBODY>
<TR><TU>과목</TU><TU>주석</TU><TU COLSPAN="2">제 10(당) 기</TU><TU COLSPAN="2">제 9(전) 기</TU></TR>
<TR><TD>Ⅰ. 유동부채</TD><TD></TD><TE></TE><TE>2,000</TE><TE></TE><TE>1,500</TE></TR>
<TR><TD>  단기차입금</TD><TD>8</TD><TE>1,000</TE><TE></TE><TE>900</TE><TE></TE></TR>
<TR><TD>  유동성장기차입금</TD><TD>8</TD><TE>300</TE><TE></TE><TE>200</TE><TE></TE></TR>
<TR><TD>  유동성사채</TD><TD>9</TD><TE>500</TE><TE></TE><TE>-</TE><TE></TE></TR>
<TR><TD>  사채할인발행차금</TD><TD></TD><TE>(5)</TE><TE></TE><TE>-</TE><TE></TE></TR>
<TR><TD>  임대보증금</TD><TD></TD><TE>400</TE><TE></TE><TE>400</TE><TE></TE></TR>
<TR><TD>  현재가치할인차금</TD><TD></TD><TE>(20)</TE><TE></TE><TE>(25)</TE><TE></TE></TR>
<TR><TD>Ⅱ. 비유동부채</TD><TD></TD><TE></TE><TE>3,000</TE><TE></TE><TE>3,100</TE></TR>
<TR><TD>  1. 사채 (주석9)</TD><TD></TD><TE>1,000</TE><TE></TE><TE>1,500</TE><TE></TE></TR>
<TR><TD>  사채할인발행차금</TD><TD></TD><TE>(10)</TE><TE></TE><TE>(15)</TE><TE></TE></TR>
<TR><TD>  장기차입금</TD><TD>8</TD><TE>2,000</TE><TE></TE><TE>1,800</TE><TE></TE></TR>
<TR><TD>  현재가치할인차금</TD><TD></TD><TE>50</TE><TE>1,950</TE><TE>60</TE><TE>1,740</TE></TR>
<TR><TD>  전환사채</TD><TD>10</TD><TE>700</TE><TE></TE><TE>-</TE><TE></TE></TR>
<TR><TD>  전환권조정</TD><TD></TD><TE>(30)</TE><TE></TE><TE>-</TE><TE></TE></TR>
<TR><TD>  사채상환할증금</TD><TD></TD><TE>15</TE><TE></TE><TE>-</TE><TE></TE></TR>
<TR><TD>  신주인수권부사채</TD><TD></TD><TE>100</TE><TE></TE><TE>-</TE><TE></TE></TR>
<TR><TD>  리스부채</TD><TD></TD><TE>80</TE><TE></TE><TE>90</TE><TE></TE></TR>
</TBODY></TABLE>
"""


@pytest.mark.parametrize("name, expected", [
    ("단기차입금", "단기차입금"),
    ("  1. 단기차입금 (주석5,12)", "단기차입금"),
    ("유동성장기부채", "유동성장기부채"),
    ("유동성 장기차입금", "유동성장기부채"),
    ("유동성사채", "유동성사채"),
    ("유동성전환사채", "유동성사채"),
    ("사채", "사채"),
    ("(2) 사채", "사채"),
    ("장기차입금", "장기차입금"),
    ("전환사채", "전환사채"),
    ("신주인수권부사채", "신주인수권부사채"),
    ("유동성신주인수권부사채", "유동성사채"),
    ("사채할인발행차금", "조정계정"),
    ("현재가치할인차금", "조정계정"),
    ("전환권조정", "조정계정"),
    ("사채상환할증금", "조정계정"),
    ("교환사채", "미포함"),
    ("장기미지급금", None),
    ("단기대여금", None),
])
def test_classify_borrowing_accounts(name, expected):
    assert classify(name) == expected


def test_extract_borrowings_face_value():
    """조정계정(할인차금·전환권조정·상환할증금)은 반영하지 않고 차입금 액면만 집계"""
    st = parse_document([BORROWING_BS]).get("BS", False)
    assert current_period_cols(st.rows, st.header_row_count) == [2, 3]
    by_cat: dict[str, list] = {}
    for line in extract_borrowings(st.rows, st.header_row_count):
        by_cat.setdefault(line.category, []).append((line.account, line.kind, line.amount))

    assert by_cat["단기차입금"] == [("단기차입금", "차입금", 1_000_000)]
    assert by_cat["유동성장기부채"] == [("유동성장기차입금", "차입금", 300_000)]
    assert by_cat["유동성사채"] == [("유동성사채", "차입금", 500_000)]
    assert by_cat["사채"] == [("1. 사채 (주석9)", "차입금", 1_000_000)]
    assert by_cat["장기차입금"] == [("장기차입금", "차입금", 2_000_000)]
    assert by_cat["전환사채"] == [("전환사채", "차입금", 700_000)]
    assert by_cat["신주인수권부사채"] == [("신주인수권부사채", "차입금", 100_000)]
    # 조정계정과 범위 외 계정은 명세에 '미포함'으로만 표시
    excluded = [(account, kind) for account, kind, _ in by_cat["미포함"]]
    adjustment = "미포함(조정계정, 액면가액 기준)"
    assert excluded == [("사채할인발행차금", adjustment), ("현재가치할인차금", adjustment),
                        ("사채할인발행차금", adjustment), ("현재가치할인차금", adjustment),
                        ("전환권조정", adjustment), ("사채상환할증금", adjustment),
                        ("리스부채", "미포함(포함 범위 외)")]


def test_current_period_cols_variants():
    xbrl = [["계정명", "계정ID", "제 56 기", "제 55 기"], ["단기차입금", "x", 1, 2]]
    assert current_period_cols(xbrl, 1) == [2]
    no_mark = [["과목", "제 10 기", "제 9 기"], ["단기차입금", 1, 2]]
    assert current_period_cols(no_mark, 1) == [1]
    no_header = [["단기차입금", 1, 2]]
    assert current_period_cols(no_header, 0) == [1]


def _sumifs(ws, formula: str) -> float:
    """테스트용 SUMIFS 계산: =SUMIFS($E$a:$E$b,$A$a:$A$b,연도,$B$a:$B$b,"분류")"""
    m = re.match(r'=SUMIFS\(\$E\$(\d+):\$E\$(\d+),\$A\$\d+:\$A\$\d+,(\d+),\$B\$\d+:\$B\$\d+,"(.+)"\)', formula)
    first, last, year, category = int(m[1]), int(m[2]), int(m[3]), m[4]
    return sum(ws.cell(row=r, column=5).value or 0 for r in range(first, last + 1)
               if ws.cell(row=r, column=1).value == year and ws.cell(row=r, column=2).value == category)


def test_borrowings_sheet(monkeypatch):
    reports = [_report("20250320000002", "감사보고서 (2024.12)"), _report("20240320000001", "감사보고서 (2023.12)")]
    _patch_dart(monkeypatch, reports, {"20250320000002": [BORROWING_BS], "20240320000001": [SAMPLE]})
    resp = TestClient(app).get("/dart/api/annual", params={"corp_code": "00123456", "years": "2023,2024"})
    ws = load_workbook(io.BytesIO(resp.content))["차입금"]

    assert [c.value for c in ws[5]][:3] == ["구분", 2024, 2023]
    rows = {ws.cell(row=r, column=1).value: r for r in range(6, 14)}
    # 원문 천원 단위 그대로, 액면가액 기준 (현재가치할인차금 등 미반영)
    expected_2024 = {"단기차입금": 1_000, "유동성장기부채": 300, "유동성사채": 500, "사채": 1_000,
                     "장기차입금": 2_000, "전환사채": 700, "신주인수권부사채": 100}
    for category, amount in expected_2024.items():
        assert _sumifs(ws, ws.cell(row=rows[category], column=2).value) == amount
        assert _sumifs(ws, ws.cell(row=rows[category], column=3).value) == 0      # 2023: 차입금 계정 없음
    assert ws.cell(row=rows["차입금 합계"], column=2).value == "=SUM(B6:B12)"
    assert sum(expected_2024.values()) == 5_600

    # 명세: 미포함 행도 표시하되 SUMIFS 분류가 달라 합계에서 빠짐
    detail = [[c.value for c in r] for r in ws.iter_rows(min_row=16) if r[0].value]
    assert ["2024", "미포함", "리스부채"] == [str(detail[-1][0]), detail[-1][1], detail[-1][2]]
