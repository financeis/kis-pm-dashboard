# tools — 개발·검증 도구

## 범위
- `kis_dev.py`: 개발용 조회 전용 KIS 클라이언트와 토큰 원천 파싱(`token_source`, `read_token`, `remaining_minutes`, `require_token(min_minutes=210)`, `KisClient.get(path, tr_id, params, tr_cont="")`). 단독 실행 시 원천 경로·만료·남은 분·상태만 출력(`--min N` 미달이면 종료 코드 2).
- `pq_harness.py`: 전체 빌드 없이 쿼리를 숨김 Excel에서 실제로 돌리는 시험 도구(파이썬 API `Harness`와 CLI). 정적 입력표·시드·실행 모드·이름 정의·셀 수정, 동기 새로 고침, 결과 요약·CSV, 저장·PNG 렌더, 페이지 시험용 대역 빌더.
- `make_contest_universe.py`(대회 명단 생성), `theme_helper.py`(`extract` 검토 목록 / `check` 테마표 검증), `verify_*.py`(통합 검증: `verify_universe` 명단 경계 재판정, `verify_buttons` 버튼 실행·시간, `verify_recalc` 독립 재계산, `verify_screens` 기준점·렌더, `verify_migration` 이관 왕복, `verify_security` 비밀값 검사, `verify_regress` NAV·성과 회귀, `verify_common` 공용).
- 하지 않는 것: 결과물 통합문서의 기능(빌더·쿼리·VBA), 토큰 발급.

## 불변식
- 토큰을 발급하지 않는다. `/uapi/` 밖 경로(토큰 발급 등)와 `/trading/` 경로(주문·계좌)는 `KisClient`가 거부한다. 기본 초당 8건 제한·`EGW00201` 재시도.
- 토큰 원천(인자 > 환경변수 `KIS_TOKEN_SOURCE` > 이 폴더 위의 `KIS_PM_Dashboard.xlsm` > `.xlsx`; `pq_harness`와 `verify_common`은 그다음 git 공통 디렉터리 기준 주 저장소의 통합문서까지 찾음 — 별도 작업 폴더에서 돌릴 때)은 Excel로 열지 않고 `xlsx_tables`로 파싱한다. 남은 시간 210분 미만이면 Excel을 띄우거나 호출하기 전에 멈춘다(종료 코드 2).
- 토큰·앱키·시크릿·계좌번호·HTS ID를 출력·JSON·CSV·로그에 쓰지 않는다. 하네스는 `tblToken` 읽기·덤프·렌더를 거부하고, 다른 출력에 토큰 문자열이 섞이면 `<토큰 가림>`으로 바꾼다.
- 하네스는 시험 통합문서에 `T_Token`을 넣지 않고(요청·참조 시 종료 코드 3, Excel 기동 전), 토큰은 정적 `_sys!tblToken`으로 복사한다. 결과 파일은 저장소 밖에만 쓴다(저장소 경로 거부). Excel은 DispatchEx, 자기 PID만 종료하며 파이썬이 죽어도 작업 개체로 자기 Excel만 정리된다.
- 하네스는 빌더에서 `settings_rows`·`SAMPLE_WATCH`·`sample_trades`·`MACRO_ROWS`·`HOLIDAYS`·`TRADE_HEADERS`·`LOADS`·`Builder.nav_links`·`HEADER_RIGHT`·`HEADER_STATUS`를 가져다 쓴다 — 빌더에서 이름이 바뀌면 여기도 바꾼다.
- 검증 스크립트는 실제 통합문서를 버튼 실행 때만 Excel로 열고(한 번에 하나, 열기·실행 직전마다 원천 남은 시간 재확인, 실행 뒤 저장), 실패 경로·이관 왕복·시계 모사는 스크래치 복사본에서만 한다. 화면 캡처는 자기 Excel 창만 찍는다.

## 구현 패턴
- 하네스 새로 고침: 백그라운드 새로 고침 → `QueryTable.Refreshing` 0.1초 폴링(시간 초과 시 취소) → `AfterRefresh` 이벤트로 성공·실패 판정. 폴링 중 자기 Excel 창에만 `WM_TIMER`를 보내 완료 감지를 앞당긴다. 실패가 60초 안이면 이벤트를 끄고 동기로 한 번 다시 돌려 M 오류 메시지를 얻는다(그때 API 호출이 반복될 수 있음). `sync=True`면 처음부터 동기.
- 하네스 기본 정적 표: `tblSettings`(빌더 기본값), `tblWatch`·`tblTrades`(sample=True면 빌더 샘플), `tblMacro`, `tblHolidays`, `tblRunCtl`(mode=build), `tblContest`·`tblThemeBase`·`tblOverride`(빈 표). CSV는 이름이 `코드`로 끝나거나 0으로 시작하는 숫자가 있는 열을 텍스트로 둔다.
- 기본 표 이름: 빌더 `LOADS`에 있으면 그 이름, 아니면 `T_X` → `tblX`, `T_A_X` → `tblA_X`. 의존 쿼리는 M 본문의 식별자로 자동 추가.
- `make_contest_universe.py`: 마스터(종류 ST·FS·DR, 우선주·SPAC 제외)로 후보 → 마스터 시총이 하한의 절반 이상이면 일봉(FHKST03010100, `J`, 원주가)으로 판정, 상장주식수는 9/30 기준(이후 추가상장분 차감, 판단 못 하는 이벤트면 종료 코드 4), 출력 파일이 있으면 `--force` 없이 덮지 않음(종료 코드 5).

## 시험
- 하네스 자체: `T_IndexNow` 4행·상태 OK, 원천 파일 수정시각·크기 불변, 시험 통합문서 쿼리 목록에 `T_Token` 없음, `--min-minutes`를 크게 주면 Excel 없이 종료 코드 2, 정적 표·시드·이름·셀 수정·저장·렌더 동작.
- `kis_dev.py`: 주문 경로 거부, 출력에 토큰 문자열 없음.
- 검증 스크립트: 실제 통합문서 대상 결과(JSON)와 종료 코드로 판단하고, 원천 응답 재조회는 조회 전용·소량으로.
