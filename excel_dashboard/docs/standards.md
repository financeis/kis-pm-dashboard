# 규칙

## 쿼리 (powerquery/*.pq)

- KIS 호출은 모두 `fnKisGet`(연속 조회는 `fnKisGetPage`)을 거친다. 직접 `Web.Contents`로 KIS를 부르면 위반이다(토큰·재시도·오류 처리 누락).
- 토큰을 발급하는 쿼리는 `T_Token` 하나뿐이다. 다른 쿼리가 `oauth2/tokenP`를 부르거나 `T_Token`을 참조하면 위반이다.
- 버튼 전용 쿼리는 자기 실행 모드(`tblRunCtl`의 `mode`; 빈칸·알 수 없는 값은 `build`)에서만 KIS를 부르고, 그 밖에서는 자기 표 → `<표이름>Seed` → 열을 모두 갖춘 빈 표 순서로 돌려준다. 예외는 `T_Class`뿐이다(로컬 VALUESearch 파일만 읽으므로 build 모드에서도 읽음).
- KIS를 부르는 계산은 **함수 단계**(`Fetch = () as table => …`)에 두고 활성 모드일 때만 호출한다. 값 단계에 두면 Excel의 재평가로 build 모드에서도 호출된다.
- 버튼 전용 쿼리는 같은 버튼 실행 안에서 다시 조회하지 않는다: `tblRunCtl`의 `started`가 있고 자기 표의 가장 늦은 `조회시각` ≥ `started`면 자기 표를 돌려준다(`started`가 비면 확인하지 않음).
- 마지막 단계는 `try Table.Buffer(Main)`이고, 실패하면 `Table.TransformColumnTypes(fnFallback("tbl…", 열목록, 사유), 형식)`으로 직전 표를 형식까지 다시 입혀 돌려준다. 형식을 다시 입히지 않으면 날짜가 일련번호로 보인다.
- 종목별 `try`는 그 안에서 값을 강제 평가한다(`List.Buffer`·`Table.Buffer`는 항목 오류를 올리지 않고 빈 행으로 적재함). 일부 종목 실패는 그 종목 행만 `오류: <사유>`, 조회한 종목이 모두 실패하면 오류로 올려 대체 경로로 보낸다.
- Power Query가 채우는 표는 모두 `상태`·`조회시각` 열을 둔다(토큰 표만 `status`·`checked`). 성공한 실행은 모든 행을 `OK`(또는 정상 값 `데이터 없음`·`추정 없음`)로 쓴다. 상태 값은 `OK` / `오류: <사유>` / `이전 데이터(갱신 실패: <사유>)` / `데이터 없음` / `추정 없음`만 쓴다.
- 대상 종목은 `fnPageRows()`로만 얻고, 대회편입·NICS·테마·섹터는 `fnClassify`로만 붙인다. 이 판정을 다른 쿼리에서 다시 구현하면 위반이다(화면마다 판정이 달라짐).
- 종목코드는 6자리 **텍스트**다(선행 0, 영문 포함 코드 `0126Z0` 등). `Number.From`으로 바꾸면 위반이다. 금액은 억원, 비율은 소수(표시는 %), 데이터 없음은 null(0 금지).
- 쿼리 이름 `T_X` → 표 `tblX`, `T_A_X` → `tblA_X`, 보조 함수는 `fn` 접두사. 기존 표의 열 이름은 바꾸지 않는다(수식·VBA·다른 쿼리가 이름으로 참조).

## 통합문서 계약 (빌더·페이지·VBA 공통)

- 표 이름, 이름 정의(`분석코드`, `시세_/전체_/업종_/분석_최근조회`와 `…_상태`, `대회코드목록`), 매크로 이름(`RefreshQuick`·`RefreshFull`·`RefreshSector`·`RunAnalysis`), `tblRunCtl` 키(`mode`·`started`·`now_override`·`quiet`·`last_summary`)는 고정이다. 바꾸려면 쿼리·VBA·페이지·빌더를 함께 바꾼다.
- 페이지 표의 적재 위치는 각 페이지 모듈의 `LOADS`가 정하고 빌더는 그대로 쓴다. 페이지 표는 적재 직후·첫 새로 고침 전에 `QueryTable.RefreshStyle = 0`(덮어쓰기)으로 바꾼다(`prepare_tables`). 이후 되돌리지 않는다.
- 페이지 표와 `_calc`·`_store` 표에는 빌더가 임시 행 추가·`ListRows.Add`/`Delete`를 하지 않는다(쌓인 표는 Excel이 거부하고, 행 삭제는 아래 표를 끌어올린다). 페이지 모듈 안에서는 아래에 다른 표가 없는 뉴스·이벤트 페이지만 빈 표 서식용 임시 행 1개를 넣었다 지운다.
- 버튼 전용 쿼리는 연결 속성 '모두 새로 고침 시 이 연결 새로 고침'을 끄고 파일 열 때 새로 고침도 끈다. 파일 열 때 도는 쿼리는 `T_Token`뿐이다.
- `tools/pq_harness.py`가 빌더에서 가져다 쓰는 이름(`settings_rows`, `SAMPLE_WATCH`, `sample_trades`, `MACRO_ROWS`, `HOLIDAYS`, `TRADE_HEADERS`, `LOADS`, `Builder.nav_links`, `Builder.HEADER_RIGHT`, `Builder.HEADER_STATUS`)은 이름과 형태를 유지한다.

## VBA (vba/*.bas)

- 소스는 UTF-8 텍스트로 관리하고 cp949로 표현 가능한 문자만 쓴다(한글·`·`·`▲▼` 가능). `—`·`✓` 같은 문자는 `ChrW`로 만든다. 빌더 삽입기는 cp949로 못 쓰는 문자가 있으면 줄 번호와 함께 거부한다.
- `Application.CalculateUntilAsyncQueriesDone` 금지(통합문서 표를 읽는 쿼리와 교착).
- `mode`·`started` 쓰기가 실패하면 실행을 중단한다(낡은 `started`가 조회를 조용히 막음). 정리 단계는 오류가 나도 `mode=build`·상태 표시줄·화면 갱신을 되돌린다. 한 단계 실패로 전체를 멈추지 않는다.
- 요약 창은 Excel이 보이고 `quiet`가 Y가 아닐 때만 띄우고, 요약 문구는 항상 `last_summary`에 쓴다(자동화가 창에 막히지 않게).

## 빌더 (build_dashboard.py, migrate.py, xl_helpers.py)

- Excel은 항상 별도 인스턴스(`xl_helpers.new_excel` = DispatchEx)로 띄우고 자기 PID만 종료한다. `Dispatch`·`GetObject`로 사용자가 연 Excel에 붙으면 위반이다.
- AccessVBOM은 `access_vbom()` 블록 안에서만 켜고, 그 블록이 Excel 시작부터 종료까지를 감싼다(Excel은 이 값을 인스턴스 시작 때 한 번 읽음).
- 기존 통합문서는 `migrate.py`(zip/XML 파싱)로만 읽는다. Excel로 열면 위반이다(토큰 발급 위험).
- 덮어쓰거나 옮기기 전에 반드시 `backup/`에 복사한다. 이관에 실패하면 새 통합문서로 바꾸지 않고 멈춘다(`--no-migrate`일 때만 새로).
- 토큰은 원천에서 파싱해 새 통합문서 `tblToken`에 메모리로만 심는다. 숨은 옵션 `--seed-token`(토큰 JSON 파일)은 남겨 두되 개발·운영에 쓰지 않는다.

## 한국어 Excel

- 숫자 서식은 현지화 코드로 넣는다(`[색10]`·`[색11]`·`G/표준` — `set_nf`·`NF`). 영문 코드(`[Red]`, `General`)는 한국어판에서 깨진다.
- 조건부 서식은 위치 인수로 추가하고, 추가 전에 대상의 첫 셀을 선택한다(수식 안 상대 참조 기준).
- 날짜는 `Value2`에 일련번호로 넣는다. LET 변수명·표 이름·이름 정의가 셀 주소(`D0`, `A1` 등)와 겹치지 않게 한다.

## 데이터 파일 (data/)

- `contest_universe_20260930.csv`·`themes_base.csv`·`nics_codes.csv`는 UTF-8 BOM, 종목코드 텍스트. 테마표의 종목코드 집합·순서는 명단과 같아야 하고 `tools/theme_helper.py check`가 통과해야 한다. `nics_codes.csv`와 `fnNicsCodes` 쿼리 안의 표는 같아야 한다.

## 의존성과 git

- 빌드는 Python 3 + `pywin32`·`pyyaml`(표준 라이브러리 외), 도구는 `requests`까지. 새 패키지를 들이면 README 준비 항목에 적는다. 파일 파싱은 표준 라이브러리(`xlsx_tables.py`)로 한다.
- `.xlsm`·`.xlsx`·`backup/`·`history/`·사용자 제공 자료(VALUESearch xlsx, Q.Pack PDF)와 `kis_devlp.yaml`(저장소 루트·`~/KIS/config`)은 커밋하지 않고 수정하지 않는다. 파일은 경로를 지정해 스테이징한다.
- 코드 주석·문서는 한국어. 파이썬 이름은 모듈·함수·변수 snake_case, 클래스 PascalCase, 상수 UPPER_SNAKE_CASE, 널리 알려진 것(URL·ID 등) 외의 축약어 금지. 공개 함수에는 목적·인자·반환·예외·예시를 담은 docstring.

## 합치기 전 검증

- `python -m compileall -q excel_dashboard` 종료 코드 0.
- 바꾼 쿼리는 하네스로 실제 실행(활성 모드·build 모드 둘 다, build에서 KIS 요청 0)과 파이썬 독립 재계산으로 확인.
- 빌더·VBA·페이지를 바꿨으면 **복사본**으로 이관 빌드 → 해당 `tools/verify_*.py` 실행(버튼·재계산·화면·이관·보안·회귀). 실제 통합문서 재빌드는 그다음.
- 빌드 뒤 AccessVBOM이 원래 상태인지, 통합문서에 비밀값이 없는지(`verify_security.py`) 확인한다.
