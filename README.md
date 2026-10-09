# kakao-cpa-bot

## 1. 카카오톡 세무·회계 Q&A 스킬 서버
`kakao_tax_skill_server.py` 상단 주석 참조.

## 2. DART 재무제표 조회 페이지 (`/dart`)

회사명을 검색하고 조회연도를 고르면 연도별 **재무상태표·손익계산서**를 엑셀 한 파일로 내려받습니다.
시트 구성: `정보`(연도별 출처·확인사항) → `BS(2025)` → `PL(2025)` → `BS(2024)` → `PL(2024)` …

### 실행
```bash
pip install -r requirements.txt
export DART_API_KEY="OpenDART 인증키"     # https://opendart.fss.or.kr 에서 발급
uvicorn kakao_tax_skill_server:app --host 0.0.0.0 --port 8000
# 브라우저에서 http://localhost:8000/dart
```

### 연도별 데이터 출처
| 회사 구분 | 1순위 | 대체 |
|---|---|---|
| 사업보고서 제출회사(상장사 등) | OpenDART XBRL 재무제표 | 사업보고서 원문 표 파싱 |
| 비상장 외감법인 | 감사보고서 원문 표 파싱 | 별도↔연결 대체 시 `정보` 시트에 표시 |

- 연도 = 보고서명 `(2025.12)`의 사업연도 종료연도
- 원문 파싱 금액은 `(단위 : 천원)` 등을 읽어 **원 단위로 환산** (주당이익 행은 환산 제외), 주석번호 열은 문자로 유지
- 원문 파싱 결과는 자동 인식이므로 합계 검증·원문 대사를 권장

### 엑셀 시트에서 바로 조회 (VBA)
`excel/DartFS.bas` — 엑셀 시트에 회사명·연도를 입력하고 버튼을 누르면 `BS(2025)`, `PL(2025)` … 시트가 생성됩니다.

1. 새 통합문서 → **Excel 매크로 사용 통합문서(.xlsm)** 로 저장
2. `Alt+F11` → 파일 → 파일 가져오기 → `excel/DartFS.bas` 선택
   (`DartFS.bas`는 한글이 깨지지 않도록 CP949로 저장된 가져오기용, `DartFS_utf8.bas`는 열람용으로 내용은 같음)
3. `Alt+F8` → `SetupDartSheet` 실행 (최초 1회) → `조회` 시트와 버튼 생성
4. `조회` 시트에 회사명·시작연도·종료연도·구분(별도/연결)·서버주소 입력 → **[재무제표 가져오기]**

- 동명 회사가 여러 개면 후보 목록(고유번호)이 표시됨 → 고유번호를 B7에 입력 후 다시 실행
- 다시 조회하면 기존 `BS(*)`, `PL(*)`, `조회정보` 시트는 지우고 새로 만듦
- 서버(`uvicorn ...`)가 실행 중이어야 함: 내 PC에서 실행하면 `http://localhost:8000`, 배포했으면 그 주소
- Windows용 Excel 2013 이상 (Mac용 Excel은 WinHTTP가 없어 지원 안 함)

### 외부 배포 시 접근 제한
환경변수 `DART_WEB_TOKEN`을 설정하면 `/dart/api/*` 호출에 토큰이 필요합니다
(엑셀 `조회` 시트의 접근토큰 칸, 웹페이지는 `/dart?token=...`). 설정하지 않으면 누구나 내 DART 인증키 호출한도를 쓸 수 있습니다.

### API (엑셀 VBA·Power Query에서도 호출 가능)
| 경로 | 설명 |
|---|---|
| `GET /dart/api/search?q=회사명` | 회사 검색 (corp_code 반환) |
| `GET /dart/api/reports?corp_code=` | 감사보고서·사업보고서 목록 |
| `GET /dart/api/annual?corp_code=&years=2023,2024,2025&fs_div=OFS` | 연도별 BS·PL 엑셀 (`OFS` 별도 / `CFS` 연결). `corp_code` 대신 `corp_name`도 가능(동명 회사면 409 + 후보 목록) |
| `GET /dart/api/excel?rcept_no=&corp_code=&mode=doc` | 보고서 1건 전체 재무제표 엑셀 |

### 테스트
```bash
pip install pytest && python -m pytest -q
```
