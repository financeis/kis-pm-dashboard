# powerquery — Power Query(M) 쿼리 원본

## 범위
- 통합문서에 들어가는 쿼리 77개의 원본이다. 빌더가 이 폴더의 `.pq`를 전부 쿼리로 넣고(파일 이름 = 쿼리 이름), 적재 목록에 있는 것만 표로 올린다. 통합문서 안의 쿼리를 직접 고치면 다음 빌드에서 사라지므로 수정은 여기서만 한다.
- 이 폴더가 하지 않는 것: 화면 서식·차트·이름 정의(페이지 모듈과 빌더), 버튼 순서와 판정(VBA), 이관(빌더). 표의 적재 위치도 여기서 정하지 않는다.

## 종류
- 공통 함수 `fn*`: `fnKisGet`(KIS GET + 공통 헤더, `EGW00201`만 최대 5회 백오프, 네트워크 1회 재시도, 설정·토큰 문제는 즉시 오류), `fnKisGetPage`(`tr_cont` 헤더 연속 조회), `fnToken`(캐시 토큰 읽기만), `fnCfg`(yaml), `fnSettings`(키/값), `fnFallback`(직전 표 유지), `fnNum`·`fnYmd`(변환), `fnClassify`·`fnPageRows`(분류·대상 종목의 단일 출처), `fnMultiQuote`(30종목/회), `fnPxRunCtl`(모드·시험 시계·`started`), `fnPxCalendar`, `fnFwdEps`, `fnTargetOpinions`·`fnTargetConsensus`, 종목분석 보조 `fnA*`.
- `T_Token`: 유일한 토큰 발급자(자기 표 `tblToken`을 다시 읽어 남은 3시간 이상이면 재사용).
- [모두 새로 고침] 쿼리: 기존 `T_Universe`·`T_Quote`·`T_Rank`·`T_Flow`·`T_Global`·`T_IndexNow`·`T_IndexHist`·`T_Sector`·`T_PriceHist`·`T_Holdings`·`T_Positions`·`T_Trades`·`T_NAV`·`T_Risk`와 `T_News`·`T_MktFunds`.
- 버튼 전용 쿼리: `T_Class`, `T_Sessions`, `T_PxStore`, `T_PxMetrics`, `T_FlowU`, `T_Fin`, `T_Target`, `T_Est`, `T_EstSnap`, `T_CSL`, `T_Events`, `T_SectorKRX`, `T_Company`, `T_ThemeAgg`, `T_A_*` 13개.

## 불변식 (어기면 화면·버튼이 깨짐)
- KIS 호출은 `fnKisGet`/`fnKisGetPage`로만. `T_Token` 외의 쿼리는 토큰을 발급하지 않고 `T_Token`을 참조하지 않는다(시험 하네스는 `T_Token`을 참조하는 쿼리를 거부한다).
- 버튼 전용 쿼리는 `fnPxRunCtl()`(또는 같은 규칙)의 모드가 자기 활성 모드일 때만 KIS를 부른다 — `T_Sessions`·`T_PxStore`: quick·full / `T_Class`·`T_FlowU`·`T_Fin`·`T_Target`·`T_Est`·`T_EstSnap`·`T_CSL`·`T_Events`: full / `T_SectorKRX`: sector·full / `T_A_*`: analysis. 그 밖에서는 자기 표 → `<표>Seed` → 열을 다 갖춘 빈 표. `T_Class`는 KIS를 부르지 않고 VALUESearch 파일을 모든 모드에서 읽는다. `T_PxMetrics`·`T_Company`·`T_ThemeAgg`는 KIS 없이 모든 모드에서 계산한다.
- KIS를 부르는 계산은 함수 단계(`Fetch = () as table => …` / `Survey`)에 두고 `Main = if <활성> and not <같은 실행에서 이미 조회> then Fetch() else <자기 표>`. '이미 조회' = `started`가 있고 자기 표 최신 `조회시각` ≥ `started`(종목분석은 종목코드도 같아야 함).
- 마지막 단계 `Tried = try Table.Buffer(Main)`, 실패 시 `Table.TransformColumnTypes(fnFallback("tbl…", 열, 사유), 형식)`. 종목별 `try` 안에서 값을 강제 평가하고, 실패 종목만 `오류: <사유>`(직전 값·조회시각 유지), 조회한 종목이 모두 실패하면 오류로 올린다.
- 모든 적재 표에 `상태`·`조회시각`. 상태 값은 `OK`/`데이터 없음`/`추정 없음`/`오류: …`/`이전 데이터(…)`만. 열 이름·순서는 소비자(페이지·VBA·다른 쿼리)와의 약속이라 바꾸면 같이 바꾼다.
- 대상 종목은 `fnPageRows()`, 분류는 `fnClassify`(결과는 유효 코드당 1행 — 종목코드로 다시 조인)로만. 종목코드는 텍스트, 금액 억원(KIS 백만원 ÷ 100, 원 ÷ 1e8), 비율 소수, 데이터 없음 null.
- 자기참조 저장소(`T_PxStore`·`T_Fin`·`T_EstSnap`·`T_CSL`·`T_Target`·`T_Est`)는 자기 표가 비었을 때만 시드를 읽는다. `T_EstSnap`은 행을 지우지 않는다. `T_PxStore`는 (종목코드, 일자) 유일·달력 밖 날짜 없음(full이 지움).

## 구현 패턴
- 모드·시계·`started`는 `fnPxRunCtl()`로 읽는다(`now_override`는 `T_Sessions`·`T_PxStore`·`T_Fin`·`T_Company`가 '오늘'로 씀). 설정은 `fnSettings()`를 `try`로 읽고 없거나 잘못된 값은 기본값(weekly_days 7, force_weekly N, target_window_months 6, profile_days 60, event_days_ahead 60, event_days_back 7, fin_lag_q 45, fin_lag_y 90).
- 주간 규칙(`T_Fin`·`T_CSL`, 목표가·추정의 전체 재조사): full이면서 (`force_weekly = Y` 또는 새 종목 또는 직전 오류 또는 오늘 − 조회시각 날짜 ≥ `weekly_days`)인 종목만 조회, 나머지는 행 유지.
- 페이지가 0이 될 수 있는 표는 `Table.FromRecords(…, 열목록, MissingField.UseNull)`로 열을 보장한다. 큰 표 결합은 `Table.Buffer`·조인으로 하고 행마다 레코드 목록을 만드는 패턴은 피한다(대회종목 표가 3.5~6초 → 2초).
- 날짜 인자: 종목 투자자 일별은 오늘을 15:40 전 거부, 투자의견은 100행/회 날짜 분할, 예탁원은 `CTS` + `tr_cont=N`. 헤더 주석에 필드 → 열 대응과 단위 근거를 적는다.

## 시험
- `tools/pq_harness.py`로 실제 API 실행: 활성 모드(작은 명단 ≤30종목, `started` 설정), build 모드(요청 0 — 끊어진 `cfg_path`나 요청 계수 모의 서버로 확인), 시드만 있는 경우, 대체 경로(설정 경로를 틀리게 → `이전 데이터…` + 날짜 형식 유지).
- 원천 응답으로 파이썬 독립 재계산(`tools/kis_dev.py` 조회 전용 클라이언트). 경계: 영문 포함 코드, 상장 직후 종목, 4월 결산, 은행(추정·매출 없음), 커버하지 않는 종목, 정지 종목.
- 장 시작 전·휴장일은 `now_override`로 모사하고 끝나면 비운다.
