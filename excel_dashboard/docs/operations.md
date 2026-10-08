# 운영

## 준비 (한 번)

1. Windows 11 + Microsoft 365 Excel 한국어판(동적 배열·LAMBDA 지원 버전).
2. `~/KIS/config/kis_devlp.yaml`에 실전 앱키 `my_app`, 시크릿 `my_sec`, 도메인 `prod`(실전 REST 주소). 모의투자 키는 없어도 된다(쓰지 않음).
3. Python 3(저장소에서는 `C:\Python313\python.exe`)에 `pywin32`·`pyyaml`(빌드), `requests`(도구). `pip install pywin32 pyyaml requests`.
4. VALUESearch에서 내보낸 `수집기업_valuesearch.xlsx`(시트 `Sheet2`, 열 `종목코드`·`691300.NICS 산업분류`·`691440.KSIC-세분류(11차)`·`691230.주요상품`·`691280.소속기업집단`)를 저장소 루트에 둔다. 다른 곳에 두면 빌드 뒤 [설정] `vs_path`를 그 절대 경로로 바꾼다.

## 빌드

```powershell
python excel_dashboard/build_dashboard.py                        # 같은 폴더 .xlsm(없으면 .xlsx)에서 이관해 다시 만듦
python excel_dashboard/build_dashboard.py --migrate-from <경로>   # 이관 원본 지정
python excel_dashboard/build_dashboard.py --out <경로.xlsm>      # 출력 위치(백업·로그는 그 폴더의 backup\)
python excel_dashboard/build_dashboard.py --no-migrate [--empty] # 이관 없이 새로(샘플 / 빈 매매일지)
python excel_dashboard/build_dashboard.py --cfg <yaml> --visible # 설정 파일 경로 바꾸기 / Excel 보이게
```

- 약 7.5~9분(쿼리 77개 추가에만 1.5~2분). 순서: 원본 선택 → `backup/<이름>_<시각>.<확장자>` 복사 → 파일 파싱 이관 계획 → 토큰 남은 시간 확인 → AccessVBOM 켜고 Excel 시작 → 생성·build 모드 새로 고침 → VBA 삽입 → `.xlsm` 저장 → Excel 종료·AccessVBOM 복원 → 행 수 대조 로그 → `.xlsx` 원본이었으면 백업 폴더로 이동.
- 종료 코드: 0 성공 / 1 빌드 실패(원본·기존 파일 그대로, 임시 폴더 정리) / 2 원천 토큰 남은 시간 210분 미만(Excel 미기동) / 3 이관·설정 문제(원본 없음·손상, Excel 미기동).
- 로그: `backup/<이름>_<시각>_build.log`. 확인할 줄: 표별 `원본 → 새` 행 수, `tblToken 이관(남은 N분) → 상태 '재사용'`, `AccessVBOM 원래 상태로 복원 확인: …`, 경고.
- 빌드하는 동안 그 통합문서를 Excel로 열어 두지 않는다(저장·이동 실패). 사용자가 열어 둔 다른 Excel 창은 건드리지 않는다.
- 종료 코드 2가 나면: 통합문서를 열어 [모두 새로 고침](남은 시간이 3시간 미만이면 `T_Token`이 1회 발급 — 카카오톡 알림) → 저장 → 다시 빌드.
- 백업 되돌리기: Excel을 닫고 `backup/`의 원하는 파일을 `KIS_PM_Dashboard.xlsm`(또는 `.xlsx`)로 복사. 백업본도 토큰을 담고 있다.

## 처음 한 번 (빌드 직후)

1. `KIS_PM_Dashboard.xlsm` 열기 → [콘텐츠 사용](매크로·데이터 연결) → '웹 콘텐츠 액세스' 창이 뜨면 익명 → 연결.
2. [설정]에서 대회 기간·초기자금·수수료율·거래세율·벤치마크 확인.
3. [모두 새로 고침](약 1분 15초) → [대회종목] [전체](처음 약 15분, KIS 약 5,900회). [전체]는 15:40 이후에 누르면 오늘 수급까지 들어간다.

## 매일 운영

| 시간 | 할 일 | 소요 |
|---|---|---|
| 장 전 | [모두 새로 고침] → [시세](마지막 세션 봉만 갱신) | 약 1분 15초 + 30~45초 |
| 장중 | 필요할 때 [모두 새로 고침](시세판·순위·추정가집계) + [시세], [업종], [조회] | [시세] 30~37초, [업종] 약 20초, [조회] 약 1분 |
| 매매 직후 | [매매일지] 한 줄 → [모두 새로 고침] | 약 1분 15초 |
| 15:40 이후 | [전체] | 평일 4~5분(통합문서를 연 뒤 첫 실행은 준비·대기로 약 1.5~3분 더), 주간 항목이 도는 날은 10분 이상 |

- 장 시작 직후 KIS에 오늘 KOSPI 일봉이 없으면 [시세] 상태가 `일부 오류 1건`(`오류: 오늘 세션 미확인`)이다 — 잠시 뒤 다시.
- [전체]가 끝나면 `history/est_snap.csv`(UTF-8 BOM)가 갱신된다. 스냅샷 표가 비면 기존 CSV를 덮지 않고 그 단계를 실패로 판정하고, 스냅샷 행이 기존 CSV보다 적으면 기존 파일을 두고 `history/est_snap_<시각>.csv`를 따로 만든다(일부 오류 1건) — 그런 파일이 생기면 스냅샷 표가 줄어든 이유(재생성·이관 문제)를 확인한다.

## 설정 키 ([설정] `tblSettings`)

| 키 | 의미 | 유효 범위·기본 |
|---|---|---|
| `cfg_path` | kis_devlp.yaml 경로 | 기본 `~/KIS/config/kis_devlp.yaml`의 절대 경로 |
| `vs_path` | VALUESearch 내보내기 파일 | 빌드 시점 저장소 루트의 `수집기업_valuesearch.xlsx` 절대 경로(앞뒤 따옴표 허용) |
| `sector_basis` | 기존 화면 '섹터'의 기준 | 대분류 / 업종 / 세부(기본) / 대테마 — 그 밖의 값은 세부 |
| `weekly_days` | 주간 항목 주기(달력일) | 7, 음수·문자면 7 |
| `force_weekly` | 다음 [전체]에서 주간 항목 강제 | N(기본) / Y — [전체] 뒤 N으로 돌아감 |
| `target_window_months` | 목표가 컨센서스 창 | 6(1~24) |
| `profile_days` | 종목분석 N일 매물대 세션 수 | 60(5~500) |
| `event_days_ahead` / `event_days_back` | 이벤트 기간 | 60 / 7(0~366) |
| `fin_lag_q` / `fin_lag_y` | 분기·결산 실적 공시 지연(일) | 45 / 90 — 바꾸면 그 종목을 다음에 받을 때 반영 |
| `stock_flow` | 시세판 종목별 수급 조회 | Y / N(호출 한도가 빠듯할 때) |

## 다른 PC에서 쓰기

- 통합문서는 git에 없으므로(토큰·매매기록) 직접 옮긴다: `KIS_PM_Dashboard.xlsm` + `history\`를 한 폴더에, `수집기업_valuesearch.xlsx`를 그 폴더나 바로 위 폴더에(저장소에 들어 있으므로 clone한 저장소의 `excel_dashboard\`에 통합문서를 두면 그대로 찾음), `kis_devlp.yaml`을 그 PC의 `%USERPROFILE%\KIS\config\`에(USB 등 — 메일·클라우드 금지).
- 메일·다운로드로 받은 파일은 [속성] → [차단 해제] 체크(인터넷 표시가 붙은 `.xlsm`은 Office가 매크로를 막음).
- 처음 열 때 [콘텐츠 사용]을 누르면 매크로(`Auto_Open` → `FixLocalPaths`)가 [설정] `cfg_path`·`vs_path`를 그 PC 위치로 바꾼다 — 지금 경로에 파일이 없고 표준 위치에 있을 때만. 모든 버튼도 시작할 때 같은 확인을 하고, 바꾸면 요약에 `[주의] 이 PC에 맞게 설정 경로를 바꿈: …`.
- 두 PC에서 번갈아 쓰면 통합문서 하나만 원본으로(동시에 열지 않기). 사본을 따로 쓰면 매매일지·가격 이력이 갈린다. 동기화 폴더는 Google Drive·Dropbox처럼 일반 폴더 경로로 보이는 것을 쓴다 — OneDrive에서 연 통합문서는 `ThisWorkbook.Path`가 인터넷 주소로 나올 수 있어 이력 CSV 저장(`history\`)과 경로 맞추기가 실패할 수 있다(이 PC에서 시험하지 않은 알려진 Excel 동작).
- 그 PC에서 다시 만들려면: `git clone https://github.com/financeis/kis-pm-dashboard.git` → `pip install pywin32 pyyaml requests` → `kis_devlp.yaml`을 `%USERPROFILE%\KIS\config\`에, 통합문서(+ `history\`)를 `excel_dashboard\`에 → `python excel_dashboard/build_dashboard.py`(통합문서 없이 새로 만들려면 `--no-migrate --empty` — 매매일지·설정이 비고 첫 [전체]가 약 15분). VALUESearch 파일은 저장소에 들어 있다. 이관 원본의 `cfg_path`·`vs_path`가 그 PC에 없으면 빌더가 기본 위치(`%USERPROFILE%\KIS\config\kis_devlp.yaml`, 저장소 루트)로 바꾸고 로그에 남긴다.
- git 원격: `origin` = 공개 `financeis/kis-pm-dashboard`(여기로만 푸시), `upstream` = KIS 공식 저장소(받기만).

## 데이터 재생성

- 대회 명단: `python excel_dashboard/tools/make_contest_universe.py --base-date 20260930 --sessions 20260922,20260923,20260928,20260929,20260930 --mcap-min-eok 1000 --turnover-min-eok 25 --out excel_dashboard/data/contest_universe_20260930.csv [--audit <파일>] [--force]` — 약 1,900회 호출·5~6분. 기존 파일은 `--force` 없이 덮지 않는다(종료 코드 5). 분할·합병 등 판단할 수 없는 이벤트가 있으면 종료 코드 4로 멈춘다. 결과를 바꾸면 빌드를 다시 해야 통합문서의 `tblContest`에 반영된다.
- 기본 테마표: `python excel_dashboard/tools/theme_helper.py extract --out-dir <저장소 밖 폴더>`(검토용 목록: 명단 + VALUESearch + KIS 테마 마스터) → `data/themes_base.csv` 편집 → `python excel_dashboard/tools/theme_helper.py check`(종료 코드 0이어야 함) → 빌드. 사용자 수정은 기본표가 아니라 [설정] 수정표에 한다.
- VALUESearch 파일 갱신: 같은 열 이름으로 내보내 `vs_path` 위치에 덮어쓰기 → [전체](분류 갱신) → [모두 새로 고침](종목DB 반영).

## 검증 명령

```powershell
python -m compileall -q excel_dashboard
python excel_dashboard/tools/kis_dev.py                       # 토큰 원천·만료·남은 분·상태(토큰 값 없음)
python excel_dashboard/tools/pq_harness.py --queries T_IndexNow --load T_IndexNow=tblIndexNow --show 5
python excel_dashboard/tools/verify_security.py --workbook excel_dashboard/KIS_PM_Dashboard.xlsm --extra excel_dashboard/history/est_snap.csv
python excel_dashboard/tools/verify_universe.py --token-source excel_dashboard/KIS_PM_Dashboard.xlsm
python excel_dashboard/tools/verify_buttons.py --workbook <복사본.xlsm> --token-source excel_dashboard/KIS_PM_Dashboard.xlsm ...
python excel_dashboard/tools/verify_recalc.py  --workbook <저장한.xlsm> --token-source <원천.xlsm>
python excel_dashboard/tools/verify_regress.py --workbook <저장한.xlsm> --token-source <원천.xlsm>
python excel_dashboard/tools/verify_screens.py --workbook <저장한.xlsm>
python excel_dashboard/tools/verify_migration.py --source excel_dashboard/KIS_PM_Dashboard.xlsm --work <스크래치 폴더> --token-source excel_dashboard/KIS_PM_Dashboard.xlsm
```

- 토큰이 필요한 도구는 원천의 남은 시간이 210분 미만이면 종료 코드 2로 멈춘다.
- 실패 경로·이관 왕복·장 시작 전 모사는 실제 통합문서가 아니라 복사본으로 한다. 결과 파일은 저장소 밖(스크래치 폴더)에 둔다.
