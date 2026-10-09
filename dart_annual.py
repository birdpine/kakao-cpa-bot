"""
회사명 + 조회연도 -> 연도별 재무상태표(BS)·손익계산서(PL)

연도별 보고서 선택 규칙 (연도 = 보고서명 '(2025.12)'의 사업연도 종료연도):
  1. 사업보고서가 있으면(상장사 등) OpenDART XBRL 재무제표를 우선 사용
     - 비12월 결산법인은 XBRL 사업연도 표기가 달라 조회되지 않을 수 있음 → 원문 파싱으로 대체
  2. XBRL이 없거나 실패하면 사업보고서 원문, 그 다음 감사보고서 원문을 파싱
  3. 감사보고서는 선택한 구분(별도/연결)과 일치하는 보고서(감사보고서/연결감사보고서)를 먼저 사용
  4. 선택한 구분의 재무제표가 없으면 다른 구분으로 대체하고 확인사항에 표시
"""

import httpx

import dart_client
from dart_excel import SheetTable, YearResult, table_from_statement, xbrl_income_statement, table_from_xbrl
from dart_parser import parse_document


def pick_reports(reports: list[dict], year: int, consolidated: bool) -> list[dict]:
    """해당 연도 보고서를 시도 순서대로: 사업보고서 → 구분 일치 감사보고서 → 나머지 감사보고서"""
    cands = [r for r in reports if r.get("fiscal_year") == year]
    annual = [r for r in cands if r["is_annual"]]
    audit = sorted((r for r in cands if not r["is_annual"]), key=lambda r: r["is_consolidated"] != consolidated)
    return annual + audit


async def collect_year(client: httpx.AsyncClient, corp_code: str, year: int, fs_div: str,
                       reports: list[dict]) -> YearResult:
    consolidated = fs_div == "CFS"
    label = "연결" if consolidated else "별도"
    res = YearResult(year=year)
    candidates = pick_reports(reports, year, consolidated)
    if not candidates:
        res.notes.append("DART에 해당 사업연도 감사보고서·사업보고서가 없습니다.")
        return res

    for report in candidates:
        if report["is_annual"]:
            try:
                _, items = await dart_client.fetch_xbrl_statements(client, corp_code, year, fs_div, try_previous=False)
                bs, pl = table_from_xbrl(items, "BS", fs_div), xbrl_income_statement(items, fs_div)
                if bs or pl:
                    return _done(res, report, f"XBRL ({label})", bs, pl)
            except dart_client.DartError:
                res.notes.append("XBRL 미제공 → 원문 파싱")

        try:
            files = await dart_client.fetch_document_files(client, report["rcept_no"])
        except dart_client.DartError as exc:
            res.notes.append(f"{report['report_nm']} 원문 조회 실패: {exc}")
            continue
        parsed = parse_document([text for _, text in files])
        for cons in (consolidated, not consolidated):
            bs, pl = parsed.get("BS", cons), parsed.income_statement(cons)
            if bs or pl:
                if cons != consolidated:
                    res.notes.append(f"{label} 재무제표가 없어 {'연결' if cons else '별도'} 재무제표를 실었습니다.")
                bs_t = table_from_statement(bs) if bs else None
                pl_t = table_from_statement(pl) if pl else None
                for st in (bs, pl):
                    if st and not st.unit_found:
                        res.notes.append(f"{st.title} 단위 표시를 찾지 못해 '원'으로 가정 (원문 확인 필요)")
                return _done(res, report, "원문 파싱", bs_t, pl_t)
        res.notes.append(f"{report['report_nm']}에서 재무상태표·손익계산서를 찾지 못했습니다.")

    return res


def _done(res: YearResult, report: dict, method: str, bs: SheetTable | None, pl: SheetTable | None) -> YearResult:
    res.report_nm, res.rcept_no, res.method, res.bs, res.pl = report["report_nm"], report["rcept_no"], method, bs, pl
    if not bs:
        res.notes.append("재무상태표를 찾지 못했습니다.")
    if not pl:
        res.notes.append("손익계산서를 찾지 못했습니다.")
    return res
