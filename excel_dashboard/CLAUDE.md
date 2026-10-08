# KIS PM 일일 대시보드 (excel_dashboard)

한국투자증권(KIS) Open API 위에 만든 **개인용 주식 운용 대시보드**다. Excel(Microsoft 365 한국어판) + Power Query가 KIS REST를 직접 조회하고, 파이썬 빌더가 Excel COM으로 `KIS_PM_Dashboard.xlsm`(VBA 버튼 포함)을 생성·재생성한다. 사용자는 1명 — KOSPI·KOSDAQ 개별종목 2개월 모의투자대회(수익률·샤프 평가) 참가자로, 대회 종목 527개와 보유·관심 수십 종목을 매일 본다. 로컬 Windows PC 한 대, 조회 전용, 네트워크 서비스·다중 사용자 없음.
이 저장소는 KIS 공식 예제 저장소에서 출발했지만 2026-10-08에 대시보드와 무관한 공식 예제(`examples_llm/`, `stocks_info/` 등)를 지워 대시보드 전용이 됐다. API 사양(URL·tr_id·파라미터)은 `kis-code-assistant` MCP나 KIS 공식 예제 저장소(`upstream`, github.com/koreainvestment/open-trading-api)에서 확인한다(로컬: `git show upstream/main:examples_llm/...`). 저장소는 **공개** GitHub `financeis/kis-pm-dashboard`(원격 `origin`, 2026-10-09 공개 전환)로 올리고, KIS 공식 저장소(`upstream`)에는 올리지 않는다.

## 프로젝트 구조

```
excel_dashboard/
├── CLAUDE.md                     → 이 안내(Claude Code용)
├── AGENTS.md                     → 같은 내용(Codex용)
├── docs/
│   ├── architecture.md           → 구성 요소·쿼리 계층·버튼/빌드 흐름·경계·외부 의존
│   ├── business-rules.md         → 대회 종목·분류·세션·지표·컨센서스·Fwd PER·색 기준·판정·데이터 보존 규칙
│   ├── security.md               → 토큰 발급·비밀값·권한·개발 중 토큰 규칙·AccessVBOM·기록
│   ├── standards.md              → 쿼리·통합문서 계약·VBA·빌더·한국어 Excel·git·합치기 전 검증 규칙
│   ├── engineering-notes.md      → Excel·Power Query 함정, KIS API 실측 동작, 반복 작업 체크리스트
│   ├── operations.md             → 준비·빌드·처음 한 번·매일 운영·설정 키·데이터 재생성·검증 명령
│   ├── contracts.md              → 버튼·이름 정의·실행 제어 표·입력표·출력표·파일 형식·빌드 인자
│   └── tracking/
│       ├── status.md             → 완료(검증 근거)·남은 일·막힌 것
│       ├── decisions/
│       │   ├── index.md          → 결정 목록
│       │   └── 0001~0009-*.md    → 개별 결정(배경·결정·대안·결과)
│       └── findings.md           → 미해결 문제
├── powerquery/
│   └── AGENTS.md                 → 쿼리 77개: 종류·모드·가드·대체 경로 불변식·시험
├── pages/
│   └── AGENTS.md                 → 새 페이지 4개: 모듈 계약·표 예약·색 기준점·열 묶음·차트
├── vba/
│   └── AGENTS.md                 → 버튼 매크로: 실행 순서·사전 준비·판정·CSV·cp949
├── tools/
│   └── AGENTS.md                 → 시험 하네스·조회 전용 클라이언트·명단 생성·검증 스크립트
└── data/
    └── AGENTS.md                 → 대회 명단·기본 테마표·NICS 코드표 형식과 갱신 규칙
```

빌더 본체(`build_dashboard.py`, `migrate.py`, `xl_helpers.py`, `xlsx_tables.py`)는 이 폴더 바로 아래에 있다.

## 반드시 지킬 것

1. **조회 전용** — KIS 주문·정정·취소·계좌 API를 어디서도 부르지 않는다.
2. **토큰은 `T_Token` 쿼리만 발급한다** — 발급마다 사용자에게 카카오톡 알림, 1분 1회 제한. 개발 중에는 토큰 원천 통합문서를 Excel로 열지 않고 파일을 파싱해 읽으며, 남은 시간 210분 미만이면 호출 없이 멈춘다. 앱키·시크릿·토큰·계좌번호·HTS ID는 어떤 출력·로그·파일·커밋에도 남기지 않는다.
3. **사용자 데이터 보존** — 통합문서는 백업 없이 덮거나 옮기지 않고, 이관에 실패하면 새로 만들지 않고 멈춘다. 실패 경로·이관·시계 모사 시험은 복사본으로만.
4. **AccessVBOM은 빌드 동안만** 켜고 성공·실패와 관계없이 원래 상태(지금은 '값 없음')로 되돌린다.
5. **Excel 자동화 안전** — 별도 인스턴스(DispatchEx)로만, 자기 PID만 종료, `Application.CalculateUntilAsyncQueriesDone` 금지, 사용자가 연 Excel 창·파일은 건드리지 않는다.

## 작업 전에 읽을 것

- 항상: `docs/standards.md`, `docs/engineering-notes.md`, 고칠 폴더의 `AGENTS.md`.
- 쿼리를 고치기 전: `powerquery/AGENTS.md`의 모드·가드·대체 경로 불변식과 engineering-notes의 '새로 고침 직후 다시 도는 쿼리'·'오류가 빈 행이 되는 버퍼'. 버튼 쿼리를 추가하면 engineering-notes의 체크리스트대로 빌더·VBA 등록까지.
- 버튼 순서·판정·VBA를 고치기 전: `vba/AGENTS.md`, `docs/contracts.md`의 버튼 표와 `tblRunCtl` 키(바꾸면 쿼리·페이지·빌더를 함께).
- 페이지 서식·차트를 고치기 전: `pages/AGENTS.md`의 표 행 추가 금지·RefreshStyle·기준점 가드·열 묶음 규칙.
- 빌드·이관을 고치거나 실제 통합문서를 다시 만들기 전: `docs/operations.md`의 빌드 절차와 종료 코드, `docs/engineering-notes.md`의 '빌더 동작'(새로 고침 순서·`_calc` 간격·시드·임시 행), `docs/security.md`의 개발 중 토큰 규칙. 먼저 `python excel_dashboard/tools/kis_dev.py`로 원천 토큰 남은 분을 확인하고 복사본으로 빌드해 본다.
- 대회 종목·분류·지표·색 기준을 건드리기 전: `docs/business-rules.md`(수치·규칙은 사용자가 정한 그대로 — 대회 기준은 '이상'이다).

## 문제가 생기면

- 즉시 사용자에게 알릴 것: 토큰이 발급됐거나(알림톡) 비밀값이 출력·파일·커밋에 노출됨 / 사용자 통합문서·백업 손상·분실 또는 매매일지·설정·수정표·스냅샷 이관 누락 / AccessVBOM이 원래 상태로 돌아오지 않음 / 주문·계좌 API 호출 흔적 / 대회 종목 판정이 규칙(9/30 시총 1,000억·5일 평균 거래대금 25억 이상 보통주)과 다름.
- 그 밖의 문제(데이터 품질, 속도, 화면 깨짐, API 동작 변화 등): 고칠 수 있으면 고치고 알게 된 사실을 `docs/engineering-notes.md`에, 지금 못 고치면 이유와 함께 `docs/tracking/findings.md`에 기록한다.
