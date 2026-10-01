# 아키텍처

## 구성 요소와 연결

- **통합문서 `KIS_PM_Dashboard.xlsm`** — 사용자가 매일 여는 유일한 결과물. 보이는 시트 15개(대시보드·시장·업종·대회종목·종목분석·뉴스·이벤트·포트폴리오·리스크·성과·시세판·매매일지·매매분석·종목DB·설정·가이드)와 숨김 시트 5개로 이뤄진다.
  - `_data`: 기존 [모두 새로 고침] 쿼리의 원천 표(지수·업종·수급·해외·순위·보유·가격 이력(보유·관심) 등)
  - `_sys`: 접근토큰 캐시 `tblToken` 하나만
  - `_calc`: 버튼 전용 계산·캐시 표 9개(분류·세션 달력·가격 지표·종목 수급·분기 실적·목표주가·KIS 추정·추정 스냅샷·신용/공매도/대차)와 [모두 새로 고침] 표인 증시 자금(`_data`에 두면 기존 계산열과 충돌해서 여기 둠)을 2행에 45열 간격으로
  - `_store`: 가격 이력 저장소 `tblPxStore`(약 13만 행이라 따로)
  - `_seed`: 빌더가 만드는 정적 표 — 실행 제어 `tblRunCtl`, 대회 명단 `tblContest`, 기본 테마표 `tblThemeBase`, 이관용 `…Seed` 표
- **Power Query 쿼리 77개** — KIS REST(실전 도메인, HTTPS GET)를 직접 부르고, KIS 종목·업종 마스터 zip(인증 없음)과 로컬 파일 두 개(`~/KIS/config/kis_devlp.yaml`, VALUESearch 내보내기 xlsx)를 읽는다. 통합문서 안의 표는 `Excel.CurrentWorkbook()`으로 읽는다(설정·실행 제어·자기 표·다른 계산 표).
- **VBA 모듈 `mod_refresh`** — 페이지의 도형 버튼 4개가 매크로를 부른다. 매크로는 `tblRunCtl`에 실행 모드를 적고 정해진 표들의 QueryTable을 순서대로 동기 새로 고침한 뒤 단계별로 판정해 이름 정의 칸(최근 조회 시각·상태)에 기록한다. 데이터를 직접 받지는 않는다.
- **파이썬 빌더** (`build_dashboard.py` + `migrate.py` + `xl_helpers.py` + `xlsx_tables.py` + `pages/*.py`) — 별도 숨김 Excel 인스턴스(COM DispatchEx)로 통합문서를 매번 새로 만든다. 입력: 쿼리 원본 `powerquery/*.pq`(통합문서 쿼리의 원본), 정적 데이터 `data/*.csv`, 페이지 레이아웃 모듈, VBA 텍스트 `vba/*.bas`, 기존 통합문서에서 파일 파싱으로 읽은 이관 데이터.
- **개발·검증 도구** (`tools/`) — 쿼리 단위 시험용 임시 통합문서(하네스), 조회 전용 KIS 파이썬 클라이언트, 명단 생성, 검증 스크립트. 결과물 통합문서에는 들어가지 않는다.

## 쿼리 계층 (위가 아래를 쓴다)

| 계층 | 쿼리 | 비고 |
|---|---|---|
| 공통 함수 | fnSettings·fnCfg·fnToken·fnKisGet·fnKisGetPage·fnFallback·fnNum·fnYmd 등 | KIS 호출은 반드시 fnKisGet(연속 조회는 fnKisGetPage) |
| 분류·소속 | T_Class(VALUESearch 파일) → fnClassify → fnPageRows | 대회편입·NICS·테마·섹터와 대상 종목 목록의 단일 출처 |
| 원천 저장소(버튼 전용) | T_Sessions, T_PxStore(자기참조), T_FlowU, T_Fin(자기참조·주간), T_Target, T_Est, T_EstSnap(자기참조·영구), T_CSL(주간), T_Events, T_SectorKRX, T_A_* 13개 | 자기 실행 모드에서만 KIS 호출 |
| 계산(버튼 전용, KIS 호출 없음) | T_PxMetrics → T_Company → T_ThemeAgg | 원천을 '적재된 표'로 읽음 |
| 모두 새로 고침 | 기존 15개(T_Token·T_Universe·T_Quote·T_Rank·T_NAV 등) + T_News·T_MktFunds | T_Universe·T_Rank·T_Quote는 fnClassify로 분류 반영 |

계산 쿼리는 원천 쿼리를 직접 참조하지 않고 통합문서에 적재된 표를 읽는다. 그래서 버튼의 새로 고침 순서가 곧 데이터 흐름 순서이고, 순서를 바꾸면 한 박자 늦은 값이 계산된다.

## 대표 흐름 — [시세] 버튼

1. `RefreshQuick` → `tblRunCtl`에 `mode=quick`, `started=지금`을 바로 쓴다(준비가 없는 버튼이라 기다리지 않음, 쓰기에 실패하면 실행 중단).
2. `T_Token` 새로 고침 — 남은 유효시간 3시간 이상이면 캐시 재사용, 미만이면 1회 발급. 판정: 만료까지 5분 이상 남은 토큰이 있어야 정상.
3. `T_Sessions` — KOSPI(0001) 일봉 약 300세션을 세션 달력으로, 날짜가 바뀐 첫 실행이면 국내휴장일조회로 오늘 개장 여부를 기록.
4. `T_PxStore` — `fnPageRows()` 종목을 멀티 시세 API(30종목/회)로 받아 저장소의 **마지막 세션 봉**을 추가·갱신하고 자기 표에 이어 쓴다.
5. `T_PxMetrics` → `T_Company`(페이지 행에 분류·지표·수급·실적·목표가·추정·신용·이벤트 결합, 103열) → `T_ThemeAgg`(대회 종목만 테마별 집계).
6. 단계별 판정(정상 / 일부 오류 n건 / 실패) → `시세_최근조회`·`시세_상태` 기록 → `mode=build` 복원 → 요약 창(보이는 Excel이고 `quiet`가 Y가 아닐 때만).

[전체]는 같은 틀에 분류·종목 수급·목표주가·KIS 추정·스냅샷·이벤트·분기 실적·신용/공매도/대차·KRX 업종과 이력 CSV 저장이 더해지고(Excel 세션 첫 실행이면 먼저 자기 표를 build 모드로 준비 새로 고침한 뒤 '15초 + 준비 시간'과 다음 초까지 기다려 활성화), [업종]은 토큰 → KRX 업종 → 테마 집계, [조회]는 토큰 → `tblA_*` 13개다.

## 빌드·이관 흐름

1. 이관 원본 선택(`--migrate-from` → 출력 위치의 `.xlsm` → 같은 이름 `.xlsx`) → `backup/`에 시각 붙여 복사.
2. `migrate.py`가 원본 파일을 **zip/XML로 직접 읽어** 이관 계획을 만든다(입력표 행, 설정 키별 값, 저장소 행, 토큰, 알 수 없는 시트). 원본의 토큰이 210분 미만이면 Excel을 띄우기 전에 멈춘다.
3. AccessVBOM을 켠 상태에서 새 Excel 인스턴스를 띄워 시트·입력표·정적 표·시드 → 쿼리 77개 추가 → 표 적재(페이지 표는 첫 새로 고침 전에 덮어쓰기 방식으로) → 토큰을 `tblToken`에 심음 → `mode=build`로 전체 새로 고침(버튼 쿼리는 KIS를 부르지 않고 시드·빈 표를 돌려줌, 분류 `T_Class`만 로컬 파일을 읽음) → 기존 시트·새 페이지 서식 → VBA 텍스트 삽입 → `.xlsm` 저장 → Excel 종료 → AccessVBOM 복원.
4. 행 수 대조를 로그(`backup/<이름>_<시각>_build.log`)에 남기고, 원본이 `.xlsx`였으면 백업 폴더로 옮긴다.

## 경계

- 통합문서로 들어가는 것: 접근토큰(`_sys!tblToken`), 설정 파일 경로(`cfg_path`), 받은 데이터. 들어가지 않는 것: 앱키·시크릿·계좌번호·HTS ID(쿼리가 새로 고칠 때마다 yaml에서 읽음).
- KIS 쪽으로 나가는 것: 시세·정보 조회 GET뿐. 주문·정정·취소·계좌 조회 경로는 쿼리·도구 어디에도 없다(개발용 클라이언트는 `/trading/` 경로를 거부).
- 버튼 전용 쿼리는 [모두 새로 고침]·파일 열기·표 우클릭 새로 고침에서 KIS를 부르지 않는다. 파일을 열 때 자동으로 도는 것은 `T_Token`뿐(30분마다도 확인).

## 외부 의존

- KIS Open API 실전 도메인(초당 약 20건 제한, 토큰 1분 1회·발급마다 카카오톡 알림)
- KIS 종목·업종 마스터 파일(`kospi_code.mst.zip`·`kosdaq_code.mst.zip`·`idxcode.mst.zip`)
- VALUESearch 내보내기 `수집기업_valuesearch.xlsx`(사용자가 갱신, 기본 위치 저장소 루트, [설정] `vs_path`)
- `~/KIS/config/kis_devlp.yaml`(앱키·시크릿·도메인)
- Microsoft 365 Excel 한국어판(동적 배열·LAMBDA), Windows, 빌드용 Python 3 + pywin32·pyyaml
