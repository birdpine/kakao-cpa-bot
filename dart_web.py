"""
DART 재무제표 조회·엑셀 다운로드 웹페이지 (FastAPI 라우터)

  GET /dart                     조회 화면
  GET /dart/api/search?q=       회사 검색 (상장·비상장)
  GET /dart/api/reports?corp_code=   감사보고서·사업보고서 목록
  GET /dart/api/annual?corp_code=&years=2023,2024,2025&fs_div=OFS
                                연도별 BS(2025)·PL(2025) 시트 엑셀 (주 기능)
  GET /dart/api/excel?...       보고서 1건의 전체 재무제표 엑셀
      mode=doc   : 공시 원문 표 파싱 (비상장 감사보고서 포함 모든 보고서)
      mode=xbrl  : OpenDART XBRL 재무제표 (사업보고서만, fs_div=CFS|OFS)
"""

import logging
import re
from datetime import date
from urllib.parse import quote

import httpx
from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import HTMLResponse, Response

import dart_client
from dart_annual import collect_year
from dart_excel import build_annual_workbook, build_workbook_from_document, build_workbook_from_xbrl
from dart_parser import parse_document

router = APIRouter(prefix="/dart")
logger = logging.getLogger("dart_web")

XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
HTTP_TIMEOUT = httpx.Timeout(60.0, connect=10.0)


def _client() -> httpx.AsyncClient:
    return httpx.AsyncClient(timeout=HTTP_TIMEOUT)


def _dart_error(exc: Exception) -> HTTPException:
    if isinstance(exc, dart_client.DartError):
        return HTTPException(status_code=502, detail=str(exc))
    logger.exception("DART 호출 실패")
    return HTTPException(status_code=502, detail="DART 서버 호출에 실패했습니다. 잠시 후 다시 시도해 주세요.")


@router.get("", response_class=HTMLResponse)
async def page():
    return INDEX_HTML


@router.get("/api/search")
async def search(q: str = Query(..., min_length=1, max_length=50)):
    try:
        async with _client() as client:
            corps = await dart_client.load_corp_list(client)
    except Exception as exc:
        raise _dart_error(exc)
    return {"results": dart_client.search_corps(corps, q)}


@router.get("/api/reports")
async def reports(corp_code: str = Query(..., pattern=r"^\d{8}$")):
    try:
        async with _client() as client:
            items = await dart_client.list_reports(client, corp_code)
    except Exception as exc:
        raise _dart_error(exc)
    return {"reports": items}


@router.get("/api/annual")
async def annual(
    corp_code: str = Query(..., pattern=r"^\d{8}$"),
    years: str = Query(..., pattern=r"^\d{4}(,\d{4}){0,14}$"),
    corp_name: str = Query("", max_length=100),
    stock_code: str = Query("", max_length=10),
    fs_div: str = Query("OFS", pattern=r"^(CFS|OFS)$"),
):
    year_list = sorted({int(y) for y in years.split(",")}, reverse=True)
    if not all(1999 <= y <= date.today().year for y in year_list):
        raise HTTPException(status_code=400, detail="조회연도를 확인해 주세요.")
    try:
        async with _client() as client:
            reports = await dart_client.list_reports(client, corp_code, begin=date(min(year_list), 1, 1))
            results = [await collect_year(client, corp_code, y, fs_div, reports) for y in year_list]
    except Exception as exc:
        raise _dart_error(exc)

    meta = {"corp_code": corp_code, "corp_name": corp_name, "stock_code": stock_code}
    content = build_annual_workbook(meta, results, fs_div)
    span = f"{year_list[-1]}-{year_list[0]}" if len(year_list) > 1 else str(year_list[0])
    return _xlsx_response(content, f"{corp_name or corp_code}_재무제표_{span}.xlsx")


@router.get("/api/excel")
async def excel(
    rcept_no: str = Query(..., pattern=r"^\d{14}$"),
    corp_code: str = Query(..., pattern=r"^\d{8}$"),
    corp_name: str = Query("", max_length=100),
    stock_code: str = Query("", max_length=10),
    report_nm: str = Query("", max_length=200),
    rcept_dt: str = Query("", max_length=8),
    mode: str = Query("doc", pattern=r"^(doc|xbrl)$"),
    fs_div: str = Query("CFS", pattern=r"^(CFS|OFS)$"),
):
    meta = {"rcept_no": rcept_no, "corp_code": corp_code, "corp_name": corp_name,
            "stock_code": stock_code, "report_nm": report_nm, "rcept_dt": rcept_dt}
    try:
        async with _client() as client:
            if mode == "xbrl":
                year = dart_client.year_from_report_name(report_nm)
                if year is None:
                    raise HTTPException(status_code=400, detail="사업보고서명에서 사업연도를 찾지 못했습니다.")
                bsns_year, items = await dart_client.fetch_xbrl_statements(client, corp_code, year, fs_div)
                content = build_workbook_from_xbrl(meta, items, bsns_year, fs_div)
                suffix = "XBRL_연결" if fs_div == "CFS" else "XBRL_별도"
            else:
                files = await dart_client.fetch_document_files(client, rcept_no)
                result = parse_document([text for _, text in files])
                content = build_workbook_from_document(meta, result)
                suffix = "감사보고서" if "감사보고서" in report_nm else "원문"
    except HTTPException:
        raise
    except Exception as exc:
        raise _dart_error(exc)

    return _xlsx_response(content, f"{corp_name or corp_code}_{rcept_dt or rcept_no}_{suffix}.xlsx")


def _xlsx_response(content: bytes, filename: str) -> Response:
    filename = _safe_filename(filename)
    return Response(
        content=content,
        media_type=XLSX_MEDIA_TYPE,
        headers={"Content-Disposition": f"attachment; filename=\"dart.xlsx\"; filename*=UTF-8''{quote(filename)}"},
    )


def _safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\r\n]+', "_", name)


INDEX_HTML = """<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DART 재무제표 조회</title>
<style>
  :root { --bg:#f6f7f9; --card:#fff; --text:#1d2330; --muted:#667085; --line:#e4e7ec;
          --accent:#1f5fbf; --accent-soft:#e8f0fb; --warn:#b54708; }
  @media (prefers-color-scheme: dark) {
    :root { --bg:#14171c; --card:#1d2128; --text:#e6e9ef; --muted:#98a2b3; --line:#2e343e;
            --accent:#6ea2f0; --accent-soft:#22324a; --warn:#f2a65a; }
  }
  * { box-sizing: border-box; }
  body { margin:0; background:var(--bg); color:var(--text);
         font-family: -apple-system, "Apple SD Gothic Neo", "Malgun Gothic", "Noto Sans KR", sans-serif; }
  main { max-width: 960px; margin: 0 auto; padding: 24px 16px 64px; }
  h1 { font-size: 22px; margin: 0 0 4px; }
  .sub { color: var(--muted); font-size: 14px; margin-bottom: 20px; }
  form { display:flex; gap:8px; }
  input[type=text] { flex:1; min-width:0; padding:10px 12px; font-size:16px; border:1px solid var(--line);
                     border-radius:8px; background:var(--card); color:var(--text); }
  button { padding:10px 16px; font-size:14px; border-radius:8px; border:1px solid var(--accent);
           background:var(--accent); color:#fff; cursor:pointer; white-space:nowrap; }
  button.ghost { background:var(--card); color:var(--accent); }
  button:disabled { opacity:.5; cursor:wait; }
  section { background:var(--card); border:1px solid var(--line); border-radius:10px; margin-top:16px; overflow:hidden; }
  section h2 { font-size:15px; margin:0; padding:12px 16px; border-bottom:1px solid var(--line); }
  table { width:100%; border-collapse:collapse; font-size:14px; }
  th { text-align:center; color:var(--muted); font-weight:600; background:var(--accent-soft); }
  th, td { padding:9px 12px; border-bottom:1px solid var(--line); }
  tr.pick { cursor:pointer; }
  tr.pick:hover td { background:var(--accent-soft); }
  .badge { display:inline-block; font-size:12px; padding:1px 8px; border-radius:10px; border:1px solid var(--line); color:var(--muted); }
  .badge.listed { color:var(--accent); border-color:var(--accent); }
  .actions { display:flex; gap:6px; flex-wrap:wrap; justify-content:flex-end; }
  .actions button { padding:6px 10px; font-size:13px; }
  .msg { padding:14px 16px; color:var(--muted); font-size:14px; }
  .msg.err { color:var(--warn); }
  a { color:var(--accent); }
  .panel { padding:14px 16px; border-bottom:1px solid var(--line); }
  .panel .label { font-size:13px; color:var(--muted); margin:0 0 6px; }
  .years { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:12px; }
  .years label, .fs label { display:inline-flex; align-items:center; gap:4px; font-size:14px;
                            border:1px solid var(--line); border-radius:6px; padding:5px 9px; cursor:pointer; }
  .fs { display:flex; gap:6px; margin-bottom:14px; flex-wrap:wrap; }
  details summary { padding:12px 16px; cursor:pointer; font-size:14px; color:var(--muted); }
  @media (max-width: 640px) {
    .hide-sm { display:none; }
    td.act { display:block; border-bottom:none; padding-top:0; }
    .actions { justify-content:flex-start; }
  }
</style>
</head>
<body>
<main>
  <h1>DART 재무제표 조회</h1>
  <div class="sub">회사명을 검색하고 조회연도를 고르면 연도별 재무상태표(BS)·손익계산서(PL)를 엑셀 한 파일로 내려받습니다. 상장·비상장 모두 지원.</div>
  <form id="f">
    <input id="q" type="text" placeholder="회사명 (예: 삼성전자, ○○건설)" autocomplete="off" required>
    <button id="go">검색</button>
  </form>
  <section id="corpBox" hidden><h2>회사 선택</h2><div id="corps"></div></section>
  <section id="repBox" hidden>
    <h2 id="repTitle">보고서</h2>
    <div id="annual" class="panel"></div>
    <details id="perReport"><summary>보고서별 전체 재무제표 다운로드 (자본변동표·현금흐름표 포함)</summary><div id="reports"></div></details>
  </section>
</main>
<script>
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const errText = (b, status) => typeof b.detail === "string" ? b.detail : `요청을 처리하지 못했습니다 (${status})`;
const fmtDate = (d) => d && d.length === 8 ? `${d.slice(0,4)}-${d.slice(4,6)}-${d.slice(6)}` : esc(d);
let corp = null;

async function getJSON(url) {
  const r = await fetch(url);
  const body = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(errText(body, r.status));
  return body;
}

$("f").addEventListener("submit", async (e) => {
  e.preventDefault();
  const q = $("q").value.trim();
  if (!q) return;
  $("go").disabled = true;
  $("repBox").hidden = true;
  $("corpBox").hidden = false;
  $("corps").innerHTML = '<div class="msg">검색 중… (처음 검색 시 회사목록을 내려받느라 시간이 걸립니다)</div>';
  try {
    const { results } = await getJSON(`/dart/api/search?q=${encodeURIComponent(q)}`);
    if (!results.length) { $("corps").innerHTML = '<div class="msg">검색 결과가 없습니다.</div>'; return; }
    $("corps").innerHTML = `<table><thead><tr><th>회사명</th><th>구분</th><th class="hide-sm">고유번호</th></tr></thead><tbody>${
      results.map((c, i) => `<tr class="pick" data-i="${i}"><td>${esc(c.corp_name)}</td>
        <td style="text-align:center">${c.stock_code ? `<span class="badge listed">상장 ${esc(c.stock_code)}</span>` : '<span class="badge">비상장</span>'}</td>
        <td class="hide-sm" style="text-align:center">${esc(c.corp_code)}</td></tr>`).join("")}</tbody></table>`;
    $("corps").querySelectorAll("tr.pick").forEach(tr => tr.addEventListener("click", () => loadReports(results[tr.dataset.i])));
    if (results.length === 1) loadReports(results[0]);
  } catch (err) {
    $("corps").innerHTML = `<div class="msg err">${esc(err.message)}</div>`;
  } finally { $("go").disabled = false; }
});

async function loadReports(c) {
  corp = c;
  $("repBox").hidden = false;
  $("repTitle").textContent = `${c.corp_name} (${c.stock_code ? "상장 " + c.stock_code : "비상장"})`;
  $("annual").innerHTML = '<div class="msg" style="padding:0">공시 목록 조회 중…</div>';
  $("reports").innerHTML = "";
  $("repBox").scrollIntoView({ behavior: "smooth", block: "start" });
  try {
    const { reports } = await getJSON(`/dart/api/reports?corp_code=${c.corp_code}`);
    if (!reports.length) { $("annual").innerHTML = '<div class="msg err" style="padding:0">최근 10년 내 감사보고서·사업보고서가 없습니다.</div>'; return; }
    renderAnnual(reports);
    renderReports(reports);
  } catch (err) {
    $("annual").innerHTML = `<div class="msg err" style="padding:0">${esc(err.message)}</div>`;
  }
}

function renderAnnual(reports) {
  const years = [...new Set(reports.map(r => r.fiscal_year).filter(Boolean))].sort((a, b) => b - a);
  const hasConsolidated = reports.some(r => r.is_annual || r.is_consolidated);
  $("annual").innerHTML = `
    <p class="label">조회연도 (사업연도 종료연도)</p>
    <div class="years">${years.map((y, i) =>
      `<label><input type="checkbox" value="${y}" ${i < 3 ? "checked" : ""}>${y}</label>`).join("")}</div>
    <p class="label">재무제표 구분</p>
    <div class="fs">
      <label><input type="radio" name="fs" value="OFS" checked>별도</label>
      <label><input type="radio" name="fs" value="CFS" ${hasConsolidated ? "" : "disabled"}>연결${hasConsolidated ? "" : " (해당 없음)"}</label>
    </div>
    <button id="annualBtn">연도별 BS·PL 엑셀 다운로드</button>`;
  $("annualBtn").addEventListener("click", async () => {
    const ys = [...$("annual").querySelectorAll(".years input:checked")].map(i => i.value);
    if (!ys.length) { alert("조회연도를 하나 이상 선택해 주세요."); return; }
    const fs = $("annual").querySelector('input[name="fs"]:checked').value;
    const p = new URLSearchParams({ corp_code: corp.corp_code, corp_name: corp.corp_name,
      stock_code: corp.stock_code || "", years: ys.join(","), fs_div: fs });
    await fetchFile($("annualBtn"), `/dart/api/annual?${p}`, `연도별 재무제표 생성 중… (${ys.length}개 연도)`);
  });
}

function renderReports(reports) {
  $("reports").innerHTML = `<table><thead><tr><th>접수일</th><th>보고서명</th><th>엑셀 다운로드</th></tr></thead><tbody>${
    reports.map((r, i) => `<tr><td style="text-align:center;white-space:nowrap">${fmtDate(r.rcept_dt)}</td>
      <td><a href="https://dart.fss.or.kr/dsaf001/main.do?rcpNo=${esc(r.rcept_no)}" target="_blank" rel="noopener">${esc(r.report_nm)}</a></td>
      <td class="act"><div class="actions">
        ${r.is_annual ? `<button data-i="${i}" data-mode="xbrl" data-fs="CFS">XBRL 연결</button>
                         <button data-i="${i}" data-mode="xbrl" data-fs="OFS">XBRL 별도</button>` : ""}
        <button class="${r.is_annual ? "ghost" : ""}" data-i="${i}" data-mode="doc">원문 파싱</button>
      </div></td></tr>`).join("")}</tbody></table>`;
  $("reports").querySelectorAll("button").forEach(b => b.addEventListener("click", () => download(b, reports[b.dataset.i])));
}

async function download(btn, r) {
  const p = new URLSearchParams({ rcept_no: r.rcept_no, corp_code: corp.corp_code, corp_name: corp.corp_name,
    stock_code: corp.stock_code || "", report_nm: r.report_nm, rcept_dt: r.rcept_dt, mode: btn.dataset.mode });
  if (btn.dataset.fs) p.set("fs_div", btn.dataset.fs);
  await fetchFile(btn, `/dart/api/excel?${p}`, "생성 중…");
}

async function fetchFile(btn, url, busyLabel) {
  const label = btn.textContent;
  btn.disabled = true; btn.textContent = busyLabel;
  try {
    const res = await fetch(url);
    if (!res.ok) { const b = await res.json().catch(() => ({})); throw new Error(errText(b, res.status)); }
    const blob = await res.blob();
    const cd = res.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename\\*=UTF-8''([^;]+)/);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = m ? decodeURIComponent(m[1]) : "dart.xlsx";
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  } catch (err) {
    alert(err.message);
  } finally { btn.disabled = false; btn.textContent = label; }
}
</script>
</body>
</html>
"""
