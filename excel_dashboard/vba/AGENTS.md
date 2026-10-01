# vba — 버튼 매크로

## 범위
- `mod_refresh.bas` 하나: 공개 매크로 `RefreshQuick`([시세]), `RefreshFull`([전체]), `RefreshSector`([업종]), `RunAnalysis`([조회])와 자동화용 함수 `JudgeTable(표이름)`, `SaveEstSnapCsv()`, `RefreshedTables()`. 빌더가 이 텍스트를 모듈 `mod_refresh`에 그대로 넣는다(`Attribute` 줄 제거, 자동 삽입 줄 삭제).
- 하지 않는 것: 데이터 조회·계산(쿼리), 버튼 도형·이름 정의 생성(페이지 모듈), 레지스트리·다른 파일 변경(`history\est_snap.csv` 쓰기만 예외).

## 실행 순서 (단계 이름 = 쿼리)
- [시세] T_Token → T_Sessions → T_PxStore → T_PxMetrics → T_Company → T_ThemeAgg
- [전체] T_Token → T_Class → T_Sessions → T_PxStore → T_PxMetrics → T_FlowU → T_Target → T_Est → T_EstSnap → T_Events → T_Fin → T_CSL → T_Company → T_ThemeAgg → T_SectorKRX → 이력 CSV
- [업종] T_Token → T_SectorKRX → T_ThemeAgg
- [조회] T_Token → 이름이 `tblA_`로 시작하는 표 전부(`…Seed` 제외)
- 쿼리 `T_X`는 표 `tblX`로 찾는다(모든 시트에서). 표가 없으면 그 단계는 실패(사유 기록), 실행은 계속.

## 불변식
- 시작: [전체]는 이번 Excel 세션에서 아직 새로 고치지 않은 자기 표를 `mode=build`로 먼저 새로 고친다(모듈 수준 기록, `RefreshedTables()`로 확인). 준비가 실제로 표를 새로 고쳤으면 '15초 + 준비에 걸린 시간'(최대 120초) 동안 `DoEvents`로 기다린다(Excel이 준비한 쿼리들을 백그라운드에서 하나씩 다시 평가하는데 VBA에서는 끝을 알 수 없어서 — 고정 15초는 중복 호출이 남았음). 그다음 다음 초가 될 때까지 기다려 `mode`와 `started`를 쓰고, 둘 중 하나라도 쓰기에 실패하면 실행을 중단한다(낡은 `started`는 쿼리의 재조회를 조용히 막음). [시세]·[업종]·[조회]는 사전 준비를 하지 않는다.
- 새로 고침은 표마다 `BackgroundQuery = False` + `Refresh False`로 동기 실행하고 원래 `BackgroundQuery` 값을 되돌린다. 이미 백그라운드 새로 고침 중이면 최대 120초 기다린 뒤 그 단계를 포기한다. `Application.CalculateUntilAsyncQueriesDone`은 쓰지 않는다.
- 단계 판정: 새로 고침 오류 또는 `상태`가 `이전 데이터`로 시작하는 행 → 실패, `오류:`로 시작하는 행 → 일부 오류(종목코드 열이 있으면 서로 다른 종목 수, 없으면 행 수), 그 밖(0행 포함) → 정상. 토큰 단계는 `tblToken`의 `env=prod` 행에 만료까지 5분 이상 남은 토큰이 있어야 정상(상태 열 이름은 `status`). `상태` 열이 없는 표는 실패.
- 상태 문구: 실패 단계가 있으면 `실패: <단계들> ` & `ChrW(&H2014)` & ` 이전 데이터 표시 중`, 아니면 `일부 오류 n건`(합계), 아니면 `정상`. 끝난 시각은 `…_최근조회` 이름 칸, 문구는 `…_상태` 이름 칸(통합문서 범위 우선, 시트 범위도 찾음).
- [전체]가 끝나면 설정 `force_weekly`를 `N`으로 되돌리고(키가 없으면 추가하지 않고 요약에 경고, Esc 취소 시 그대로), `tblEstSnap`을 통합문서 폴더 `history\est_snap.csv`로 저장한다(UTF-8 BOM·CRLF, 날짜 ISO, 소수점 `.`, 따옴표 처리). 스냅샷 표가 비었는데 CSV가 있으면 덮지 않고 그 단계를 실패로 판정, 파일이 잠겼으면 실패하고 파일은 그대로. 스냅샷 표 행 수가 기존 CSV의 데이터 행(따옴표 안 줄바꿈은 세지 않음)보다 적거나 기존 CSV를 읽을 수 없으면 기존 파일을 두고 `history\est_snap_yyyymmdd_hhnnss.csv`로 따로 저장한 뒤 그 단계를 '일부 오류 1건'(`오류: …`)으로 판정한다.
- 정리 단계는 오류가 나도 `mode=build`, `Application.StatusBar` 복원, 화면 갱신·이벤트·커서 복원을 한다. 요약 창은 Excel이 보이고 `quiet`가 Y가 아닐 때만, 요약 문구는 항상 `tblRunCtl`의 `last_summary`에.
- 소스 문자는 cp949로 표현 가능해야 한다(한글·ASCII·`·` 가능). 그 밖의 문자는 `ChrW`로 만든다 — 빌더 삽입기가 거부한다. `Option Explicit`, 모든 변수 선언.

## 시험
- 정적 검사: 블록 짝(Sub/Function/If/With/For/Select/Do), `Option Explicit` 한 번, 공개 매크로 4개, 표·이름 문자열이 계약 목록에 있는지, cp949 왕복, 금지 메서드 없음.
- 통합문서 시험(스크래치 `.xlsm`, AccessVBOM은 잠금 파일로 상호 배제하며 Excel 시작 전에 켜고 종료 후 복원): 정적 대역 표(OK / 한 종목 여러 행 `오류:` / `이전 데이터…` / 상태 열 없음 / 표 없음)로 `JudgeTable` 결과, 가짜 토큰(만료 1시간·2분 뒤)으로 토큰 판정, 전체 매크로 실행 뒤 이름 칸·`mode=build`·`force_weekly=N`·CSV 내용과 BOM.
- 실제 동작: 실제 통합문서 복사본에서 COM으로 매크로 호출(숨김 Excel이면 요약 창 없음), 설정 경로를 틀리게 해 실패 문구와 직전 데이터 유지, `RefreshedTables()`로 사전 준비 추적. `Application.StatusBar`는 초기화 뒤 "FALSE" 문자열로 읽힐 수 있다.
