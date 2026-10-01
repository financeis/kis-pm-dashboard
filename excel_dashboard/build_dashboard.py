"""KIS PM 일일 대시보드(Excel + Power Query + VBA 버튼) 생성 스크립트 — 결과물 excel_dashboard/KIS_PM_Dashboard.xlsm.

    python excel_dashboard/build_dashboard.py            # 기존 통합문서에서 이관해 다시 만듦(없으면 샘플 포함 새로 만듦)
    python excel_dashboard/build_dashboard.py --empty    # (처음 만들 때) 매매일지·관심종목을 비운 상태로 생성
    python excel_dashboard/build_dashboard.py --cfg D:/my/kis_devlp.yaml   # 설정 파일 경로 지정(이관한 설정값보다 우선)
    python excel_dashboard/build_dashboard.py --migrate-from D:/old/KIS_PM_Dashboard.xlsx   # 이관 원본 지정
    python excel_dashboard/build_dashboard.py --no-migrate                # 이관 없이 새로 만듦(기존 파일은 백업)

- 이관: 원본 = --migrate-from > 출력 위치의 .xlsm > 같은 위치의 .xlsx. 원본은 Excel로 열지 않고 파일을 직접 읽는다
  (migrate.py — 열면 '파일 열 때 새로 고침'으로 토큰이 새로 발급될 수 있음). 빌드 전에 원본을 backup\\<이름>_YYYYMMDD_HHMMSS로
  복사하고, 매매일지·관심종목·설정(키별)·해외지표·휴장일·수정표·스냅샷·가격 저장소·세션 달력·주간 캐시·토큰 캐시를 옮긴다.
  이관할 입력이 있으면 샘플을 넣지 않는다. 원본이 .xlsx였으면 성공 뒤 백업 폴더로 옮긴다. 원본을 못 읽으면 멈춘다.
- 토큰: 원본 토큰의 남은 시간이 210분 미만이면 Excel을 띄우지 않고 멈춘다("토큰 갱신 필요"). 이관한 토큰을 새 통합문서의
  tblToken에 먼저 넣어 T_Token이 재사용하게 한다(새 발급·알림톡 없음). 토큰 값은 어디에도 출력하지 않는다.
- 빌드는 같은 폴더의 .build_tmp\\에서 하고, 성공하면 출력 위치로 옮긴다(실패하면 기존 통합문서는 그대로).
- 버튼 전용 쿼리는 빌드 때 build 모드(KIS 호출 없음)로 한 번만 새로 고쳐 열 구조를 만들고, '모두 새로 고침'에서 뺀다.
- VBA(vba/*.bas, UTF-8)는 코드 텍스트로 넣는다. 그동안만 레지스트리 AccessVBOM을 1로 켜고 성공·실패와 관계없이 원래대로 되돌린다.
- Power Query(M) 원본은 excel_dashboard/powerquery/*.pq 에 있으며, 이 스크립트가 통합문서에 그대로 넣습니다.
- 앱키/시크릿은 통합문서에 저장하지 않고 ~/KIS/config/kis_devlp.yaml 을 Power Query가 직접 읽습니다.
- Windows + Microsoft 365 Excel 필요 (동적 배열 함수: LET, FILTER, SORTBY, TAKE, VSTACK, HSTACK, XLOOKUP).
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import glob
import os
import re
import shutil
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xl_helpers import (  # noqa: E402
    C, NF, FONT, XL_CENTER, XL_LEFT, XL_RIGHT, XL_TOP, XL_VCENTER, XL_SHEET_HIDDEN, XL_LINE, XL_AREA,
    XL_COLUMN_CLUSTERED, XL_BAR_CLUSTERED, XL_LEGEND_TOP, XL_LEGEND_BOTTOM, XL_SECONDARY, XL_EDGE_BOTTOM,
    XL_MEDIUM, XL_SRC_RANGE, XL_YES, _vba_text, access_vbom, add_calc_column, add_names, add_query, add_vba_module, border,
    bottom_line, card, cf_databar, cf_expr, cf_heat, chart, ensure_table_style, excel_pid, excel_serial, fill, freeze,
    hyperlink, load_query, new_excel, put, quit_excel, refresh, rgb, save_xlsm, section, series, set_col_format,
    set_nf, set_palette, sheet_setup, style, style_axes, title_bar, validation_list, write_table,
)
import migrate  # noqa: E402
from pages import page_analysis, page_company, page_news_events, page_sector  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
PQ_DIR = os.path.join(HERE, "powerquery")
DEFAULT_OUT = os.path.join(HERE, "KIS_PM_Dashboard.xlsm")
DEFAULT_CFG = os.path.join(os.path.expanduser("~"), "KIS", "config", "kis_devlp.yaml")
VBA_FILES = [os.path.join(HERE, "vba", "mod_refresh.bas")]
CONTEST_CSV = os.path.join(HERE, "data", "contest_universe_20260930.csv")
THEMES_CSV = os.path.join(HERE, "data", "themes_base.csv")
BUILD_TMP_DIR = ".build_tmp"     # 출력 폴더 아래 작업 폴더: 같은 파일 이름으로 만들어(차트의 이름 참조 유지) 성공하면 옮김
TOKEN_MIN_MINUTES = 210          # 원천 토큰 남은 시간 기준(3시간 30분) — T_Token 재발급 기준(3시간)보다 길게 잡아 빌드 결과물도 발급 없이 재사용(docs/security.md의 '개발 중 토큰 규칙')


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# 페이지 모듈 (pages/*.py) — 시트 순서대로. 각 모듈: SHEET, LOADS = [(쿼리, 표, 머리글 셀)], build(builder), prepare_tables(builder)
PAGE_MODULES = [page_sector, page_company, page_analysis, page_news_events]

# 시트 순서: 보이는 시트 15개(새 페이지 4개는 시장 다음) + 숨김 시트.
#   _data 기존 PQ 원본 · _sys 토큰 캐시 · _calc 버튼 전용 계산·캐시 표(+ 증시 자금) · _store 가격 이력 저장소(약 13만 행이라 따로) ·
#   _seed 빌더가 만드는 정적 표(대회 명단·기본 테마표·실행 제어·이관 Seed)
SHEETS = ["대시보드", "시장", page_sector.SHEET, page_company.SHEET, page_analysis.SHEET, page_news_events.SHEET,
          "포트폴리오", "리스크", "성과", "시세판", "매매일지", "매매분석", "종목DB", "설정", "가이드",
          "_data", "_sys", "_calc", "_store", "_seed"]
HIDDEN_SHEETS = [s for s in SHEETS if s.startswith("_")]
NAV_ITEMS = [s for s in SHEETS if not s.startswith("_")]     # 모든 시트 상단 메뉴 링크(새 시트 포함)

# 쿼리 → (시트, 위치, 표이름). 기존 15개 표 — 빈 표 임시 행(ensure_rows)·계산열·서식은 이 표들에만 쓴다
LOADS = [
    ("T_Token", "_sys", "B2", "tblToken"),
    ("T_Universe", "종목DB", "B21", "tblUniverse"),
    ("T_IndexNow", "_data", "B2", "tblIndexNow"),
    ("T_IndexHist", "_data", "Y2", "tblIndexHist"),
    ("T_Sector", "_data", "AQ2", "tblSector"),
    ("T_Flow", "_data", "BC2", "tblFlow"),
    ("T_Global", "_data", "CJ2", "tblGlobal"),
    ("T_Rank", "_data", "BT2", "tblRank"),
    ("T_Trades", "매매분석", "B12", "tblTradeLog"),
    ("T_Holdings", "포트폴리오", "B12", "tblHoldings"),
    ("T_Positions", "_data", "DP2", "tblPositions"),
    ("T_Quote", "시세판", "B26", "tblQuote"),
    ("T_PriceHist", "_data", "CU2", "tblPriceHist"),
    ("T_NAV", "성과", "B58", "tblNAV"),
    ("T_Risk", "리스크", "B12", "tblRisk"),
]

# 버튼 전용 계산·캐시 표: _calc 시트 2행에 나란히(표마다 45열 간격 — 가장 넓은 tblPxMetrics가 38열, 겹치면 새로 고침 오류).
# 가격 저장소(약 13만 행)는 _store 시트에 따로. 이 표들과 페이지 표에는 임시 행·행 추가/삭제를 하지 않는다(표 행 삭제가
# 같은 열의 셀을 끌어올리고, 쌓인 페이지 표에서는 Excel이 행 추가를 거부함 — 페이지 작업 실측).
CALC_SPACING = 45
# 증시 자금(새 일반 쿼리 — [시장]이 구조적 참조로 읽으므로 시트는 상관없음)도 _calc 끝에 둔다. _data의 기존 표 오른쪽(EL2)에
# 두면 빌드가 tblPositions(_data 맨 오른쪽 DP2)에 계산열을 덧붙일 때 Excel이 그 행들의 오른쪽 칸을 밀어야 해서 '표에 있는 셀이
# 이동될 수 있기 때문에 이 작업은 수행되지 않습니다'로 실패한다(2026-10-01 빌드 실측). _data 안에는 8열이 들어갈 빈틈도 없다.
_CALC_TABLES = [("T_Class", "tblClass"), ("T_Sessions", "tblSessions"), ("T_PxMetrics", "tblPxMetrics"), ("T_FlowU", "tblFlowU"),
                ("T_Fin", "tblFin"), ("T_Target", "tblTarget"), ("T_Est", "tblEst"), ("T_EstSnap", "tblEstSnap"),
                ("T_CSL", "tblCSL"), ("T_MktFunds", "tblMktFunds")]
HIDDEN_LOADS = ([(q, "_calc", f"{col_letter(2 + k * CALC_SPACING)}2", t) for k, (q, t) in enumerate(_CALC_TABLES)]
                + [("T_PxStore", "_store", "B2", "tblPxStore")])
EXTRA_LOADS = HIDDEN_LOADS

# 버튼 전용 쿼리: '모두 새로 고침'에서 제외(연결 속성 RefreshWithRefreshAll = False), 파일 열 때 새로 고침 없음.
# 빌드·수동 새로 고침에서는 tblRunCtl mode = build라 KIS를 부르지 않는다. 일반 쿼리 = 기존 15개 + T_News + T_MktFunds
BUTTON_QUERIES = frozenset(["T_Class", "T_Sessions", "T_PxStore", "T_PxMetrics", "T_FlowU", "T_Fin", "T_Target", "T_Est",
                            "T_EstSnap", "T_CSL", "T_Events", "T_SectorKRX", "T_Company", "T_ThemeAgg"]
                           + [q for q, _, _ in page_analysis.LOADS])

# 빌드 때 새로 고침 순서(표 이름). T_Token(이관 토큰 재사용) → T_Class(분류, 로컬 파일만) → T_Universe(분류로 대회편입·섹터) →
# 버튼 전용 저장소(build 모드: 자기 표 → Seed → 빈 표, KIS 호출 없음 — 이관 이력이 바로 보임) → 기존 일반 쿼리(T_Quote가
# tblCSL·tblEvents를 읽으므로 저장소 다음) → 새 일반 쿼리 → 계산 쿼리(T_PxMetrics → T_Company → T_ThemeAgg) → 종목분석 표.
# 같은 새로 고침 안에서 다른 쿼리가 적재한 표는 직전 결과로 읽히므로 읽히는 표를 먼저 새로 고친다.
BUILD_REFRESH_ORDER = (["tblToken", "tblClass", "tblUniverse",
                        "tblSessions", "tblPxStore", "tblFlowU", "tblFin", "tblTarget", "tblEst", "tblEstSnap", "tblCSL",
                        "tblEvents", "tblSectorKRX"]
                       + [t for _, _, _, t in LOADS if t not in ("tblToken", "tblUniverse")]
                       + ["tblMktFunds", "tblNews", "tblPxMetrics", "tblCompany", "tblThemeAgg"]
                       + [t for _, t, _ in page_analysis.LOADS])

# 실행 제어 표 tblRunCtl (VBA mod_refresh가 읽고 씀). 값 열은 일반 서식(텍스트 서식이면 started가 글자로 저장됨), mode 말고는 빈칸
RUNCTL_ROWS = [("mode", "build"), ("started", None), ("now_override", None), ("quiet", None), ("last_summary", None)]
# 빌드 결과 계약 점검(docs/contracts.md의 이름 정의·버튼) — 없으면 빌드 실패(저장된 결과물을 출력 위치로 옮기지 않음)
REQUIRED_NAMES = ["시세_최근조회", "시세_상태", "전체_최근조회", "전체_상태", "업종_최근조회", "업종_상태",
                  "분석_최근조회", "분석_상태", "분석코드", "대회코드목록"]
REQUIRED_BUTTONS = [(page_company.SHEET, "btn_RefreshQuick", "RefreshQuick"), (page_company.SHEET, "btn_RefreshFull", "RefreshFull"),
                    (page_sector.SHEET, "btn_RefreshSector", "RefreshSector"), (page_analysis.SHEET, "btn_RunAnalysis", "RunAnalysis")]

# 이관 토큰 시드: T_Token을 잠시 이 식(임시 정적 표 tblTokenSeed를 읽음 — 토큰 값이 M 식에 들어가지 않게)으로 바꿔 한 번
# 새로 고친 뒤 원래 식으로 되돌린다. 그다음 T_Token은 자기 표(tblToken)의 유효 토큰을 재사용한다(남은 시간 3시간 이상).
TOKEN_SEED_TABLE = "tblTokenSeed"
TOKEN_SEED_FORMULA = (
    'let\n'
    f'    Src = Excel.CurrentWorkbook(){{[Name = "{TOKEN_SEED_TABLE}"]}}[Content],\n'
    '    Sel = Table.SelectColumns(Src, {"env", "token", "expires", "issued", "key_sig", "status", "checked"}, MissingField.UseNull),\n'
    '    Typed = Table.TransformColumnTypes(Sel, {{"env", type text}, {"token", type text}, {"expires", type datetime},\n'
    '        {"issued", type datetime}, {"key_sig", type text}, {"status", type text}, {"checked", type datetime}})\n'
    'in\n'
    '    Typed')

QUERY_DESC = {
    "T_Token": "KIS 접근토큰 캐시(자기참조). 남은 유효시간 3시간 미만일 때만 재발급",
    "T_Universe": "KOSPI·KOSDAQ 종목 마스터(인증 불필요)",
    "T_IndexNow": "국내 지수 현재가·시장폭", "T_IndexHist": "국내 지수 일별", "T_Sector": "업종 등락",
    "T_Flow": "시장별 투자자 순매수", "T_Global": "해외지수·환율·금리", "T_Rank": "순위(거래대금·등락률·수급)",
    "T_Trades": "매매 원장", "T_Holdings": "보유종목", "T_Positions": "종목별 손익(청산 포함)",
    "T_Quote": "보유·관심 시세판(추정가집계·신용/공매도·다음 이벤트 포함)", "T_PriceHist": "일봉+기술지표", "T_NAV": "일별 순자산·벤치마크",
    "T_Risk": "사전 위험(변동성·베타·위험기여·VaR)",
    # 새 일반 쿼리 ([모두 새로 고침])
    "T_News": "보유·관심 종목 뉴스·공시 제목 최근 100건", "T_MktFunds": "증시 자금(고객예탁금·신용융자·미수금·펀드·MMF) 일별",
    # 버튼 전용 쿼리 (모두 새로 고침 제외, 빌드·수동 새로 고침은 호출 없는 build 모드)
    "T_Class": "NICS 업종 분류(VALUESearch 파일, KIS 호출 없음) — [전체]·빌드에서 갱신",
    "T_Sessions": "세션 달력(KOSPI 일봉 날짜)·오늘 개장 여부 — [시세]·[전체]",
    "T_PxStore": "가격 이력 저장소(일봉, 변경분만 이어 붙임) — [시세]·[전체]",
    "T_PxMetrics": "기간 등락·최근 20일·60일 평균·52주 고가 대비(저장소로 계산) — [시세]·[전체]",
    "T_FlowU": "종목별 외국인·기관 순매수 1D·1W·1M — [전체]",
    "T_Fin": "분기 실적(분기 단독값)·EPS·BPS·ROE, 주 1회 — [전체]",
    "T_Target": "목표주가 컨센서스(증권사별 6개월) — [전체]",
    "T_Est": "KIS 추정 EPS(FY1·FY2) — [전체]",
    "T_EstSnap": "Fwd EPS·Fwd PER 스냅샷(영구 이력) — [전체]",
    "T_CSL": "신용잔고율·공매도 비중·대차 변화, 주 1회 — [전체]",
    "T_Events": "기업 이벤트 캘린더(예탁원 일정) — [전체]",
    "T_SectorKRX": "KRX 업종지수 기간 등락·업종별 투자자 순매수 — [업종]·[전체]",
    "T_Company": "대회종목 페이지 표 조립(KIS 호출 없음) — [시세]·[전체]",
    "T_ThemeAgg": "테마 집계(대테마·세부테마, KIS 호출 없음) — [시세]·[전체]·[업종]",
    "T_A_Info": "종목분석: 종목 정보·분류·현재가 — [조회]", "T_A_Investor": "종목분석: 투자자 일별 120세션 — [조회]",
    "T_A_TradeSize": "종목분석: 체결금액별 매매비중(당일) — [조회]", "T_A_PbarToday": "종목분석: 당일 매물대 — [조회]",
    "T_A_Profile": "종목분석: 최근 N세션 매물대 — [조회]", "T_A_Estimate": "종목분석: 외인·기관 추정가집계(당일) — [조회]",
    "T_A_CSL": "종목분석: 신용·공매도·대차 일별 — [조회]", "T_A_Targets": "종목분석: 증권사별 목표주가 — [조회]",
    "T_A_TargetTrend": "종목분석: 월말 목표주가 컨센서스 추이 — [조회]", "T_A_KisEst": "종목분석: KIS 추정실적 — [조회]",
    "T_A_Fin": "종목분석: 최근 8분기 실적 — [조회]", "T_A_News": "종목분석: 뉴스·공시 40건 — [조회]",
    "T_A_Events": "종목분석: 이 종목의 이벤트 — [조회]",
    # 보조 함수 (연결만)
    "fnClassify": "분류·대회편입 판정(명단 ± 수정표, NICS, 테마, 섹터 기준)", "fnPageRows": "대회종목 페이지 행(대회 종목 ∪ 보유·관심)",
    "fnPxRunCtl": "실행 모드·시험용 시계(tblRunCtl)", "fnARun": "종목분석 쿼리 공통 실행기(모드·코드 검사·직전 표 유지)",
}

# ---------------------------------------------------------------------------------------------------------------
# 입력표 기본값
# ---------------------------------------------------------------------------------------------------------------
VS_FILE_NAME = "수집기업_valuesearch.xlsx"   # NICS 분류 원천(VALUESearch 내보내기). 저장소 루트에 두며 git에는 넣지 않음


def main_repo_root() -> str:
    """이 체크아웃이 속한 **주 저장소**의 루트 폴더 절대 경로.

    git 워크트리에서 실행해도 사용자 파일(VALUESearch 파일·실제 통합문서)이 있는 주 저장소를 가리키도록
    `git rev-parse --git-common-dir`(주 저장소의 .git 폴더)의 상위 폴더를 쓴다. git이 없거나 실패하면 이 체크아웃의
    루트(excel_dashboard의 상위 폴더)로 대신한다.

    Returns:
        str: 예) C:\\Users\\me\\open-trading-api
    """
    root = os.path.dirname(HERE)
    try:
        import subprocess
        r = subprocess.run(["git", "-C", HERE, "rev-parse", "--git-common-dir"], capture_output=True, text=True,
                           encoding="utf-8", timeout=10)
        common = (r.stdout or "").strip()
        if r.returncode == 0 and common and not common.startswith('"'):
            common = os.path.normpath(common if os.path.isabs(common) else os.path.join(HERE, common))
            if os.path.basename(common).lower() == ".git":
                root = os.path.dirname(common)
    except Exception:  # noqa: BLE001 — git이 없어도 빌드는 계속(체크아웃 루트로 대신)
        pass
    return root


def default_vs_path() -> str:
    """설정 vs_path 기본값: 주 저장소 루트(main_repo_root — git 워크트리에서 빌드해도 주 저장소)의
    수집기업_valuesearch.xlsx 절대 경로. 파일이 있는지는 확인하지 않는다(없으면 T_Class가 상태에 사유를 남기고 직전 분류 유지).

    Returns:
        str: 예) C:\\Users\\me\\open-trading-api\\수집기업_valuesearch.xlsx
    """
    return os.path.join(main_repo_root(), VS_FILE_NAME)


def settings_rows(cfg_path: str, sample: bool):
    start = dt.date(2026, 9, 1) if sample else dt.date.today()
    end = dt.date(2026, 10, 30) if sample else dt.date.today() + dt.timedelta(days=61)
    # 나중에 더한 키는 맨 뒤에 붙인다 — build_inputs의 검증 목록이 행 위치(5·18·19행)로 걸려 있음
    return [
        ("cfg_path", "KIS 설정파일 경로", cfg_path, "앱키·시크릿을 읽을 kis_devlp.yaml 위치 (키는 통합문서에 저장하지 않음)"),
        ("start_date", "대회 시작일", start, "★ 실제 대회 시작일로 변경 (샘플: 2026-09-01)" if sample else "★ 대회 시작일"),
        ("end_date", "대회 종료일", end, "D-day·남은 거래일 계산"),
        ("init_capital", "초기 자금(원)", 100_000_000, "대회에서 받은 운용자금"),
        ("benchmark", "벤치마크", "KOSPI", "KOSPI / KOSDAQ / 혼합"),
        ("bm_kospi_weight", "혼합BM KOSPI 비중", 0.8, "벤치마크=혼합일 때 KOSPI 비중 (나머지 KOSDAQ, 일별 리밸런싱)"),
        ("rf_rate", "무위험수익률(연)", None, "비워두면 CD91일물 금리를 자동 사용 (샤프·소르티노·알파)"),
        ("fee_rate", "매매 수수료율", 0.00015, "매매일지 수수료 칸이 비었을 때 적용 (대회 규정 확인)"),
        ("tax_rate", "매도 거래세율", 0.002, "증권거래세+농특세, 매도 시만 (대회 규정 확인)"),
        ("max_weight", "종목당 최대 비중", 0.2, "초과 시 알림"),
        ("stop_loss", "기본 손절 기준", -0.08, "손절가 미입력 종목에 적용 (평균단가 대비)"),
        ("take_profit", "기본 목표 수익", 0.2, "목표가 미입력 종목에 적용 (평균단가 대비)"),
        ("move_alert", "급등락 알림 기준", 0.05, "일간 등락률 절댓값"),
        ("vol_alert", "거래량 급증 배수", 2, "전일 기준 20일 평균 거래량 대비"),
        ("target_vol", "목표 변동성(연)", 0.3, "사전 변동성이 넘으면 알림"),
        ("hist_days", "가격 이력 일수", 100, "시세판 기술지표용 (클수록 API 호출 증가)"),
        ("risk_days", "위험 계산 기간(일)", 60, "변동성·베타·VaR 계산 구간"),
        ("rank_market", "순위 시장", "전체", "전체 / 코스피 / 코스닥"),
        ("stock_flow", "종목별 수급 조회", "Y", "Y: 종목당 1회 추가 호출(외국인·기관 5/20일 순매수)"),
        ("vs_path", "VALUESearch 파일 경로", default_vs_path(),
         "NICS 업종 분류 원천(수집기업_valuesearch.xlsx, 시트 Sheet2). 기본값 = 빌드할 때 저장소 루트의 이 파일"),
        ("sector_basis", "섹터 기준", "세부",
         "대분류 / 업종 / 세부 / 대테마 — 포트폴리오·리스크·시세판·순위의 '섹터' (값이 없으면 세부→업종→대분류→KRX)"),
        ("weekly_days", "주간 항목 주기(일)", 7, "분기 실적·신용/공매도/대차·목표주가·추정 전체 조사를 [전체]에서 다시 받는 간격"),
        ("force_weekly", "주간 항목 강제 갱신", "N", "Y면 다음 [전체]에서 주간 항목을 바로 다시 받고 N으로 되돌림"),
        ("target_window_months", "목표주가 기간(개월)", 6, "증권사별 최근 N개월 안의 마지막 목표가로 컨센서스 계산"),
        ("profile_days", "매물대 기간(세션)", 60, "종목분석 N일 매물대 계산 구간"),
        ("event_days_ahead", "이벤트 기간(앞, 일)", 60, "오늘부터 며칠 뒤까지의 기업 이벤트를 받을지"),
        ("event_days_back", "이벤트 기간(뒤, 일)", 7, "며칠 전까지의 지난 이벤트도 보여 줄지"),
        ("fin_lag_q", "분기 실적 공시 지연(일)", 45, "1~3분기 실적을 분기말 + N일부터 '알려진 값'으로 봄(후행 PER·PBR 변화)"),
        ("fin_lag_y", "연간 실적 공시 지연(일)", 90, "4분기(연간) 실적을 결산일 + N일부터 '알려진 값'으로 봄"),
    ]


SAMPLE_WATCH = [
    ("000270", "자동차", None, None, "[샘플] 관심가·목표가를 입력하면 도달 시 알림"),
    ("068270", "바이오", None, None, ""),
    ("207940", "바이오", None, None, ""),
    ("034020", "원전·전력", None, None, ""),
    ("042700", "반도체장비", None, None, ""),
    ("373220", "2차전지", None, None, ""),
    ("003230", "음식료", None, None, ""),
    ("005380", "자동차", None, None, ""),
    ("277810", "로봇", None, None, ""),
    ("009540", "조선", None, None, ""),
]

MACRO_ROWS = [
    ("미국", "S&P500", "N", "SPX", 1), ("미국", "NASDAQ", "N", "COMP", 2), ("미국", "다우", "N", ".DJI", 3),
    ("미국", "필라델피아반도체", "N", "SOX", 4), ("미국", "VIX", "N", "VIX", 5),
    ("아시아", "니케이225", "N", "JP#NI225", 6), ("아시아", "항셍", "N", "HK#HS", 7), ("아시아", "상해종합", "N", "SHANG", 8),
    ("환율", "원/달러", "X", "FX@KRW", 9),
    ("금리", "미국10년", "I", "Y0202", 10), ("금리", "국고3년", "I", "Y0101", 11), ("금리", "국고10년", "I", "Y0106", 12),
    ("금리", "CD91일", "I", "Y0112", 13), ("금리", "회사채AA-", "I", "Y0102", 14),
]

HOLIDAYS = [
    (dt.date(2026, 10, 5), "개천절 대체공휴일"), (dt.date(2026, 10, 9), "한글날"), (dt.date(2026, 12, 25), "성탄절"),
    (dt.date(2026, 12, 31), "연말 휴장"), (dt.date(2027, 1, 1), "신정"), (dt.date(2027, 2, 8), "설날 연휴"),
    (dt.date(2027, 2, 9), "설날 대체공휴일"), (dt.date(2027, 3, 1), "삼일절"), (dt.date(2027, 5, 5), "어린이날"),
    (dt.date(2027, 5, 13), "부처님오신날"), (dt.date(2027, 8, 16), "광복절 대체공휴일"),
    (dt.date(2027, 9, 14), "추석 연휴"), (dt.date(2027, 9, 15), "추석"), (dt.date(2027, 9, 16), "추석 연휴"),
    (dt.date(2027, 10, 4), "개천절 대체공휴일"), (dt.date(2027, 10, 11), "한글날 대체공휴일"),
    (dt.date(2027, 12, 27), "성탄절 대체공휴일"), (dt.date(2027, 12, 31), "연말 휴장"),
]

TRADE_HEADERS = ["일자", "종목코드", "종목명", "구분", "수량", "단가", "금액", "수수료", "세금", "전략", "매매근거",
                 "목표가", "손절가", "메모", "확인"]

F_TRADE_CODE = 'IF(LEN([@종목코드])<6,RIGHT("000000"&[@종목코드],6),UPPER([@종목코드]))'
F_TRADE_NAME = ('=IF([@종목코드]="","",IFERROR(XLOOKUP(' + F_TRADE_CODE +
                ',tblUniverse[종목코드],tblUniverse[종목명]),"⚠ 코드 확인"))')
F_TRADE_AMT = '=IF(AND(ISNUMBER([@수량]),ISNUMBER([@단가])),[@수량]*[@단가],"")'
F_WATCH_NAME = ('=IF([@종목코드]="","",IFERROR(XLOOKUP(' + F_TRADE_CODE +
                ',tblUniverse[종목코드],tblUniverse[종목명]),"⚠ 코드 확인"))')
# 대회편입 = 대회 종목 여부(명단 ± 수정표, T_Universe가 fnClassify로 채움) → 종목DB에 있지만 대회 종목이 아니면 경고
F_TRADE_CHK = ('=IF([@종목코드]="","",LET(e,XLOOKUP(' + F_TRADE_CODE + ',tblUniverse[종목코드],tblUniverse[대회편입],"X"),'
               'IF(e="X","⚠ 종목DB에 없는 코드",IF(e="N","⚠ 대회 종목 아님",IF(OR([@구분]="매수",[@구분]="매도"),"✓","⚠ 구분 확인")))))')

# [설정] 시트 배치 — 기본 설정 표(B6, 29행 → B6:E35) 아래로 기존 안내 문구를 옮기고, ⑤ 수정표는 ④ 휴장일 오른쪽의
# 빈 열 묶음(W:AD)에 둔다. 다른 입력표(B:E·G:L·N:R·T:U) 아래에 두면 그 표의 행 삭제·삽입이 "표의 셀이 이동될 수 있어
# 수행되지 않습니다"로 막힌다(Excel은 아래쪽 표 일부만 미는 이동을 거부 — 하네스에서 관심종목 행 삭제로 확인).
SETTINGS_NOTE_ROW = 37            # 기존 B28·B29 안내 문구가 옮겨 오는 첫 행(설정표와 겹치지 않게), 그 다음 행은 ⑤로 가는 링크
OVERRIDE_TITLE_CELL = "W5"        # ⑤ 수정표 제목(①~④ 제목과 같은 5행), W6·W7은 사용법 안내
OVERRIDE_WARN_CELL = "AB5"        # 무시된 수정표 행 경고 칸(F_OVERRIDE_WARN) — 수정표 바로 위 제목 줄 오른쪽
OVERRIDE_ANCHOR = "W8"            # tblOverride 머리글 왼쪽 위 칸(build_inputs가 빈 표로 만듦, 사용자 행은 이관 단계에서 옮김)
OVERRIDE_HEADERS = ["종목코드", "대회편입", "대테마", "세부테마", "NICS 대분류", "NICS 업종", "NICS 세부", "메모"]
OVERRIDE_WIDTHS = [10, 9, 12, 14, 12, 14, 16, 24]   # W:AD 열 너비 (V열은 3폭 여백)
# 수정표에서 무시되는 행 수 경고: fnClassify와 같은 정리·정규화(U+00A0·U+3000 → 공백, 코드 0~31 문자 제거, 앞뒤 공백 제거,
# 대문자, 숫자만 6자리 미만이면 앞 0 채움) 뒤
#  ① 종목코드가 6자리 [0-9A-Z]가 아님(빈 코드 포함) ② 대회편입이 공란·추가·제외가 아님 ③ 종목DB(tblUniverse)에 없는 코드
#  중 하나인 행을 센다(fnClassify가 무시하는 행과 같다). ③은 종목DB에 종목코드가 하나라도 있을 때만 본다 — fnClassify도 종목DB가
#  비었으면 종목DB 확인을 하지 않는다. 메모만 있는 행·완전히 빈 행은 세지 않는다. 없으면 빈칸.
#  U+00A0은 UNICHAR(160)으로 찾는다 — 한국어 Excel의 CHAR(160)은 코드 페이지 949 기준이라 일반 공백(UNICODE 32)을 돌려줘
#  U+00A0을 바꾸지 못한다(2026-10-02 하네스 실측).
F_OVERRIDE_WARN = (
    '=LET(x_cln,LAMBDA(v_x,IFERROR(TRIM(CLEAN(SUBSTITUTE(SUBSTITUTE(v_x&"",UNICHAR(160)," "),UNICHAR(12288)," "))),"")),'
    'x_has,LAMBDA(s_x,k_x,IF(LEN(s_x)=0,FALSE,AND(ISNUMBER(FIND(MID(s_x,SEQUENCE(LEN(s_x)),1),k_x))))),'
    'x_u,SUM(--(LEN(tblUniverse[종목코드]&"")>0)),'
    'x_bad,MAP(tblOverride[종목코드],tblOverride[대회편입],tblOverride[대테마],tblOverride[세부테마],'
    'tblOverride[NICS 대분류],tblOverride[NICS 업종],tblOverride[NICS 세부],'
    'LAMBDA(a_1,a_2,a_3,a_4,a_5,a_6,a_7,LET(s_0,UPPER(x_cln(a_1)),'
    's_1,IF(AND(LEN(s_0)<6,x_has(s_0,"0123456789")),RIGHT("000000"&s_0,6),s_0),'
    'm_0,x_cln(a_2),n_e,LEN(s_0&m_0&x_cln(a_3)&x_cln(a_4)&x_cln(a_5)&x_cln(a_6)&x_cln(a_7))>0,'
    'v_c,AND(LEN(s_1)=6,x_has(s_1,"0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ")),'
    'v_m,OR(m_0="",m_0="추가",m_0="제외"),'
    'i_u,IF(v_c,OR(x_u=0,ISNUMBER(XMATCH(s_1,tblUniverse[종목코드]))),FALSE),'
    'AND(n_e,NOT(AND(v_c,v_m,i_u)))))),'
    'x_n,SUM(--x_bad),IF(x_n=0,"","⚠ 수정표 "&x_n&"행 무시됨(종목코드·값 확인)"))'
)

# 순위 블록([시장]·[대시보드]) — tblRank[대회] = ★/공란. 빈 칸이 0으로 보이지 않게 ★만 골라 씀
RANK_STAR = 'IF(tblRank[대회]="★","★","")'
NONCONTEST_GREY = "#9AA2AD"       # 대회 밖 종목 글꼴(보유·관심 파란 굵은 글씨 규칙이 우선)
MKT_RANK_TOP = 60                 # [시장] 순위 6블록 시작 행 (2열 × 3행, 블록당 12열)
MKT_FUNDS_TOP = 116               # [시장] 증시 자금 동향 구역 시작 행


def sample_trades():
    d = dt.date
    rows = [
        (d(2026, 9, 1), "005930", "매수", 60, 261000, "코어", "메모리 업황·실적 모멘텀 (예시)", None, None),
        (d(2026, 9, 1), "000660", "매수", 8, 1693000, "코어", "HBM 수요 (예시)", None, None),
        (d(2026, 9, 1), "035420", "매수", 50, 215000, "스윙", "플랫폼 저평가 반등 (예시)", None, 198000),
        (d(2026, 9, 1), "196170", "매수", 30, 301000, "스윙", "기술이전 이벤트 (예시)", 360000, 265000),
        (d(2026, 9, 2), "329180", "매수", 20, 422000, "코어", "수주 사이클 (예시)", None, None),
        (d(2026, 9, 2), "058470", "매수", 100, 64400, "스윙", "소켓 수요 회복 (예시)", 76000, 59000),
        (d(2026, 9, 8), "105560", "매수", 60, 173700, "방어", "배당·저베타로 변동성 완화 (예시)", None, None),
        (d(2026, 9, 10), "035420", "매도", 25, 208000, "스윙", "추세 이탈, 절반 축소 (예시)", None, None),
        (d(2026, 9, 15), "247540", "매수", 50, 107000, "스윙", "낙폭과대 분할매수 (예시)", None, None),
        (d(2026, 9, 17), "196170", "매도", 30, 254000, "스윙", "손절 규칙 실행 (예시)", None, None),
        (d(2026, 9, 22), "012450", "매수", 6, 1057000, "코어", "방산 수출 (예시)", None, None),
        (d(2026, 9, 22), "058470", "매도", 50, 75000, "스윙", "목표 근접, 절반 차익실현 (예시)", None, None),
        (d(2026, 9, 29), "005930", "매도", 20, 272500, "코어", "비중 조절 (예시)", None, None),
    ]
    out = []
    for (day, code, side, q, px, strat, why, tgt, stop) in rows:
        out.append([day, code, None, side, q, px, None, None, None, strat, why, tgt, stop, "[샘플] 삭제 후 사용", None])
    return out


# ---------------------------------------------------------------------------------------------------------------
# 이름 정의(설정값 등)
# ---------------------------------------------------------------------------------------------------------------
def setting_ref(key: str) -> str:
    return f'=INDEX(tblSettings[값],MATCH("{key}",tblSettings[키],0))'


NAMES = {
    "시작일": setting_ref("start_date"), "종료일": setting_ref("end_date"), "초기자금": setting_ref("init_capital"),
    "벤치마크": setting_ref("benchmark"), "무위험입력": setting_ref("rf_rate"), "수수료율": setting_ref("fee_rate"),
    "거래세율": setting_ref("tax_rate"), "최대비중": setting_ref("max_weight"), "손절기준": setting_ref("stop_loss"),
    "목표수익": setting_ref("take_profit"), "급등락기준": setting_ref("move_alert"), "거래량배수": setting_ref("vol_alert"),
    "목표변동성": setting_ref("target_vol"),
    "금리_CD91": '=IFERROR(XLOOKUP(1,(tblGlobal[이름]="CD91일")*(tblGlobal[최신여부]="Y"),tblGlobal[종가])/100,0.03)',
    "무위험수익률": "=IF(ISNUMBER(무위험입력),무위험입력,금리_CD91)",
    "현금잔고": "=초기자금+SUM(tblTradeLog[현금흐름])",
    "수익률배열": '=FILTER(tblNAV[일간수익률],tblNAV[구분]="거래일")',
    "BM배열": '=FILTER(tblNAV[BM일간],tblNAV[구분]="거래일")',
    "tblQuote_종목명": "=tblQuote[종목명]",
    "차트_섹터명": "=OFFSET(리스크!$T$13,0,0,MAX(1,COUNTA(리스크!$T$13:$T$30)),1)",
    "차트_섹터비중": "=OFFSET(리스크!$U$13,0,0,MAX(1,COUNTA(리스크!$T$13:$T$30)),1)",
    "차트_위험이름": "=OFFSET(리스크!$AH$13,0,0,MAX(1,COUNTA(리스크!$AH$13:$AH$60)),1)",
    "차트_위험비중": "=OFFSET(리스크!$AI$13,0,0,MAX(1,COUNTA(리스크!$AH$13:$AH$60)),1)",
    "차트_위험기여": "=OFFSET(리스크!$AJ$13,0,0,MAX(1,COUNTA(리스크!$AH$13:$AH$60)),1)",
    "차트_기여종목": "=OFFSET(성과!$B$39,0,0,MAX(1,COUNTA(성과!$B$39:$B$54)),1)",
    "차트_기여금액": "=OFFSET(성과!$C$39,0,0,MAX(1,COUNTA(성과!$B$39:$B$54)),1)",
}
# 계산열(tblHoldings[평가금액])을 참조하므로 계산열 추가 후 정의
NAMES_POST = {
    "주식평가액": "=SUM(tblHoldings[평가금액])",
    "순자산": "=현금잔고+주식평가액",
}


# 성과 지표 수식 (LET 기반, 오류 시 "-")
def perf(expr: str) -> str:
    return f'=IFERROR(LET(r,수익률배열,b,BM배열,rf,무위험수익률/252,n,COUNT(r),{expr}),"-")'


F_SHARPE = perf("IF(n<3,\"-\",(AVERAGE(r)-rf)/STDEV.S(r)*SQRT(252))")
F_SHARPE0 = perf("IF(n<3,\"-\",AVERAGE(r)/STDEV.S(r)*SQRT(252))")
F_VOL = perf("IF(n<3,\"-\",STDEV.S(r)*SQRT(252))")
F_SORTINO = perf("IF(n<3,\"-\",LET(d,IF(r<rf,r-rf,0),dd,SQRT(SUMSQ(d)/n),IF(dd=0,\"-\",(AVERAGE(r)-rf)/dd*SQRT(252))))")
F_BETA = perf("IF(n<3,\"-\",SLOPE(r,b))")
F_ALPHA = perf("IF(n<3,\"-\",(AVERAGE(r)-rf-SLOPE(r,b)*(AVERAGE(b)-rf))*252)")
F_CORR = perf("IF(n<3,\"-\",CORREL(r,b))")
F_TE = perf("IF(n<3,\"-\",STDEV.S(r-b)*SQRT(252))")
F_IR = perf("IF(n<3,\"-\",AVERAGE(r-b)*252/(STDEV.S(r-b)*SQRT(252)))")
F_HIT = perf("IF(n<1,\"-\",SUM(--(r>0))/n)")
F_HITBM = perf("IF(n<1,\"-\",SUM(--(r>b))/n)")
F_BEST = perf("IF(n<1,\"-\",MAX(r))")
F_WORST = perf("IF(n<1,\"-\",MIN(r))")
F_CAGR = perf("IF(n<1,\"-\",(1+TAKE(tblNAV[누적수익률],-1))^(252/n)-1)")
F_MDD = '=IFERROR(MIN(tblNAV[낙폭]),"-")'
F_CALMAR = '=IFERROR(LET(c,' + F_CAGR[1:] + ',m,MIN(tblNAV[낙폭]),IF(OR(m=0,c="-"),"-",c/ABS(m))),"-")'
F_CUM = '=IFERROR(TAKE(tblNAV[누적수익률],-1),"-")'
F_BMCUM = '=IFERROR(TAKE(tblNAV[BM누적],-1),"-")'
F_EXCESS = '=IFERROR(TAKE(tblNAV[누적수익률],-1)-TAKE(tblNAV[BM누적],-1),"-")'
F_NAVLAST = '=IFERROR(TAKE(tblNAV[순자산],-1),"-")'
F_DAYS = '=IFERROR(TAKE(tblNAV[경과일],-1),0)'
F_REMAIN = '=IF(TODAY()>종료일,0,NETWORKDAYS(TODAY()+1,종료일,tblHolidays[휴장일]))'
F_TURNOVER = '=IFERROR(SUM(tblNAV[매매금액])/AVERAGE(tblNAV[순자산]),"-")'
F_AVGEXPO = '=IFERROR(AVERAGE(FILTER(tblNAV[주식비중],tblNAV[구분]="거래일")),"-")'
F_CURDD = '=IFERROR(TAKE(tblNAV[낙폭],-1),"-")'


# ---------------------------------------------------------------------------------------------------------------
# 정적 표 쓰기 (대회 명단·기본 테마표·실행 제어·이관 Seed·이관 입력값)
# ---------------------------------------------------------------------------------------------------------------
# Excel은 COM으로 넣은 글자도 사람이 친 것처럼 해석한다("3/4" → 날짜, "-10%" → 숫자, "=…" → 수식). 이관한 글자 값이 바뀌지
# 않게, 텍스트 서식(@)이 아닌 열에 그렇게 읽힐 글자를 넣을 때는 앞에 작은따옴표를 붙인다(셀 값은 원래 글자 그대로 저장됨).
_RE_LOOKS_NUMERIC = re.compile(r"^\s*[+-]?(\d[\d,]*\.?\d*|\.\d+)([eE][+-]?\d+)?\s*%?\s*$")
_RE_LOOKS_DATE = re.compile(r"^\s*\d{1,4}\s*[-/.]\s*\d{1,2}(\s*[-/.]\s*\d{1,4})?\s*$|^\s*\d{1,2}:\d{2}(:\d{2})?\s*([AaPp][Mm])?\s*$")
_RE_LOOKS_ERROR = re.compile(r"^\s*#(N/A|NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|SPILL!|CALC!)\s*$", re.I)
INT32 = 2 ** 31


def needs_quote(s: str) -> bool:
    """이 글자를 일반 서식 칸에 넣으면 Excel이 숫자·날짜·수식·논리값·오류로 바꿔 읽는가."""
    t = s.strip()
    if not t:
        return False
    if s[:1] in "=+-@":
        return True
    if t.upper() in ("TRUE", "FALSE"):
        return True
    return bool(_RE_LOOKS_NUMERIC.match(t) or _RE_LOOKS_DATE.match(t) or _RE_LOOKS_ERROR.match(t))


def protect_text(rows: list, columns: list, text_cols=()) -> list:
    """이관 행의 글자 값 보호: 텍스트 서식이 아닌 열에서 Excel이 다르게 읽을 글자는 작은따옴표를 앞에 붙인다(write_table용)."""
    text = set(text_cols)
    out = []
    for r in rows:
        out.append([("'" + v) if (isinstance(v, str) and columns[j] not in text and needs_quote(v)) else v
                    for j, v in enumerate(r)])
    return out


def _date_nf(v) -> str:
    return NF["dt"] if isinstance(v, dt.datetime) and (v.hour or v.minute or v.second) else NF["date"]


def format_dates(lo, rows: list, columns: list, skip=()) -> None:
    """날짜 열이 아닌 열에 날짜 값을 넣은 칸에 날짜 서식을 준다(write_table은 일련번호만 써서 숫자로 보임 — 이관한 값이
    원본에서 날짜로 보이던 그대로 보이게). skip = 이미 날짜 서식을 거는 열."""
    body = lo.DataBodyRange
    if body is None:
        return
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            if isinstance(v, (dt.date, dt.datetime)) and j < len(columns) and columns[j] not in skip:
                set_nf(body.Cells(i + 1, j + 1), _date_nf(v))


def column_kinds(columns: list, rows: list, forced: dict | None = None) -> dict:
    """열 형식: text(@ 서식) / date / datetime / general(숫자·빈칸) / mixed(글자와 숫자가 섞임 — 글자는 따옴표 보호).
    종목코드·상태처럼 글자인 열은 forced로 text를 줄 수 있다."""
    forced = forced or {}
    kinds = {}
    for j, c in enumerate(columns):
        if c in forced:
            kinds[c] = forced[c]
            continue
        vals = [r[j] for r in rows if j < len(r) and r[j] is not None and not (isinstance(r[j], str) and r[j] == "")]
        if not vals:
            kinds[c] = "general"
        elif all(isinstance(v, str) for v in vals):
            kinds[c] = "text"
        elif all(isinstance(v, (dt.date, dt.datetime)) for v in vals):
            has_time = any(isinstance(v, dt.datetime) and (v.hour or v.minute or v.second) for v in vals)
            kinds[c] = "datetime" if has_time else "date"
        elif all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals):
            kinds[c] = "general"
        else:
            kinds[c] = "mixed"
    return kinds


def _value2(v, kind: str):
    """Python 값 → Range.Value2에 넣을 값 (날짜는 일련번호 — pywin32의 시간대 변환을 피함)."""
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (dt.datetime, dt.date)):
        return excel_serial(v)
    if isinstance(v, int):
        return float(v) if abs(v) >= INT32 else v      # 32비트 밖 정수는 실수로(배열 마샬링에서 VT_I8 거부 방지)
    if isinstance(v, str):
        if kind == "text":
            return v
        return ("'" + v) if needs_quote(v) else v
    return v


def write_static_table(ws, top: int, left: int, name: str, columns: list, rows: list, kinds: dict | None = None,
                       style_name: str = "TableStyleLight1"):
    """정적 표를 덩어리 쓰기(Value2 배열)로 만든다 — 13만 행 Seed도 셀 하나씩 쓰지 않는다. 빈 표는 빈 행 1개.

    Args:
        kinds: 열 형식(column_kinds 결과). 없으면 값으로 판단. text 열은 @ 서식(앞 0·영문 코드 유지).
    Returns:
        ListObject
    """
    kinds = kinds or column_kinds(columns, rows)
    ncol = len(columns)
    nrow = max(1, len(rows))
    cells = ws.Cells
    hdr = ws.Range(cells(top, left), cells(top, left + ncol - 1))
    hdr.NumberFormat = "@"
    hdr.Value2 = (tuple(str(c) for c in columns),)
    for j, c in enumerate(columns):
        col = ws.Range(cells(top + 1, left + j), cells(top + nrow, left + j))
        if kinds.get(c) == "text":
            col.NumberFormat = "@"
        else:
            set_nf(col, "General")
    grid = [tuple(_value2(r[j] if j < len(r) else None, kinds.get(c, "general")) for j, c in enumerate(columns)) for r in rows]
    chunk = max(1, 50000 // ncol)
    for s in range(0, len(grid), chunk):
        part = grid[s:s + chunk]
        ws.Range(cells(top + 1 + s, left), cells(top + s + len(part), left + ncol - 1)).Value2 = tuple(part)
    for j, c in enumerate(columns):
        if kinds.get(c) in ("date", "datetime"):
            set_nf(ws.Range(cells(top + 1, left + j), cells(top + nrow, left + j)), NF["date"] if kinds[c] == "date" else NF["dt"])
        elif kinds.get(c) == "mixed":                    # 글자·숫자와 섞인 날짜 칸은 칸마다 날짜 서식
            for i, r in enumerate(rows):
                v = r[j] if j < len(r) else None
                if isinstance(v, (dt.date, dt.datetime)):
                    set_nf(cells(top + 1 + i, left + j), _date_nf(v))
    lo = ws.ListObjects.Add(XL_SRC_RANGE, ws.Range(cells(top, left), cells(top + nrow, left + ncol - 1)), None, XL_YES)
    lo.Name = name
    lo.TableStyle = style_name
    return lo


def read_csv_table(path: str, text_cols=(), number_cols=()) -> tuple[list, list]:
    """버전 관리 데이터 CSV(UTF-8, BOM 허용) → (열, 행). text_cols는 글자 그대로(종목코드 앞 0·영문 유지), number_cols는 숫자로."""
    with open(path, encoding="utf-8-sig", newline="") as fh:
        raw = [r for r in csv.reader(fh) if any(x.strip() for x in r)]
    cols = [c.strip() for c in raw[0]]
    rows = []
    for r in raw[1:]:
        r = (r + [""] * len(cols))[:len(cols)]
        out = []
        for c, v in zip(cols, r):
            s = v.strip()
            if s == "":
                out.append(None)
            elif c in number_cols:
                f = float(s.replace(",", ""))
                out.append(int(f) if f.is_integer() and "." not in s else f)
            else:
                out.append(s)
        rows.append(out)
    return cols, rows


def query_add_order(texts: dict) -> list:
    """쿼리를 넣을 순서: 보조 함수(fn*)는 서로 참조하는 순서대로(참조되는 쪽 먼저), 그다음 적재 쿼리(T_*)는 이름순.
    참조는 주석을 뺀 M 코드에 다른 쿼리 이름이 낱말로 나오는지로 본다(순환이 있으면 남은 것은 이름순)."""
    def code(t: str) -> str:
        t = re.sub(r"/\*.*?\*/", " ", t, flags=re.S)
        return re.sub(r"//[^\n]*", " ", t)
    fns = sorted(n for n in texts if not n.startswith("T_"))
    deps = {n: {m for m in fns if m != n and re.search(r"(?<![\w#])" + re.escape(m) + r"(?!\w)", code(texts[n]))} for n in fns}
    order, done = [], set()
    while len(order) < len(fns):
        ready = [n for n in fns if n not in done and deps[n] <= done]
        if not ready:                                  # 순환(있으면) — 남은 것은 이름순
            ready = [n for n in fns if n not in done]
        for n in ready:
            order.append(n)
            done.add(n)
    return order + sorted(n for n in texts if n.startswith("T_"))


def key_sig_from_cfg(cfg_path: str) -> str | None:
    """T_Token.pq의 KeySig와 같은 앱키 체크섬(앱키 자체는 어디에도 쓰지 않음). 설정을 못 읽으면 None."""
    try:
        import yaml
        with open(cfg_path, encoding="utf-8") as fh:
            cfg = yaml.safe_load(fh) or {}
        app = str(cfg.get("my_app") or "").strip()
        return str(sum(ord(ch) for ch in app) * 7 + len(app))
    except Exception:  # noqa: BLE001 — 체크섬 비교는 경고용(실패해도 이관 토큰의 체크섬을 그대로 씀)
        return None


class Builder:
    """통합문서 하나를 만든다. 이관 계획(plan)·이관 토큰(token)은 main이 파일 파싱으로 읽어 넘긴다(Excel로 원본을 열지 않음).

    Args:
        out_path: 최종 출력(.xlsm). 메시지용 — Excel은 work_path에 저장한다.
        work_path: 빌드 중 저장 경로(출력 폴더의 .build_tmp\\ 아래 같은 파일 이름). None이면 out_path.
        plan: migrate.MigrationPlan 또는 None(이관 없음 — 샘플/빈 입력표).
        token: 이관할 migrate.TokenRecord(값은 메모리에서만 씀) 또는 None.
        est_snap: 이관 없이 만들 때 이력 CSV에서 복원한 tblEstSnap(migrate.TableData) 또는 None.
        vbom_lock: AccessVBOM 잠금 파일(개발 중 병렬 작업 상호 배제). 평소 None.
        log: 로그 줄을 모을 목록(main과 공유).
    """

    def __init__(self, out_path: str, cfg_path: str, sample: bool, visible: bool, seed_token: str | None = None, *,
                 work_path: str | None = None, plan=None, token=None, est_snap=None, vbom_lock: str | None = None,
                 log: list | None = None, cfg_explicit: bool = False):
        self.final_out = os.path.abspath(out_path)
        self.out = os.path.abspath(work_path or out_path)
        self.cfg_path = cfg_path
        self.cfg_explicit = cfg_explicit
        self.sample = sample
        self.visible = visible
        self.seed_token = seed_token
        self.plan = plan
        self.token = token
        self.est_snap = est_snap
        self.vbom_lock = vbom_lock
        self.xl = None
        self.pid = None
        self.temp_rows = []   # 빈 PQ 표에 임시로 넣은 행 (계산열·서식 정의용, 저장 직전 삭제)
        self.wb = None
        self.ws = {}
        self.lo = {}
        self.query_of = {}    # 표 이름 → 쿼리 이름 (적재한 표 전부)
        self.seeds = {}       # Seed 표 이름 → (자기 표 이름, 넣은 행 수)
        self.settings_info = {}
        self.vbom_message = ""
        self.log = log if log is not None else []

    # -------------------------------------------------------------------------------------------------------
    def say(self, msg):
        print(msg, flush=True)
        self.log.append(msg)

    def run(self, resume: str | None = None, checkpoint: str | None = None):
        """Excel 시작~종료 전체를 access_vbom으로 감싼다(Excel은 AccessVBOM을 인스턴스 시작 때 읽음 — 실측, docs/engineering-notes.md의 'AccessVBOM 읽는 시점').
        레지스트리는 성공·실패와 관계없이 원래 상태로 되돌리고 그 결과를 로그에 남긴다."""
        t0 = time.time()
        with access_vbom(log=self.say, lock_path=self.vbom_lock, owner="build_dashboard") as vbom:
            self.xl = new_excel(self.visible)
            self.pid = excel_pid(self.xl)
            try:
                if resume:
                    self.open_checkpoint(resume)
                else:
                    self.create_workbook()
                    self.build_inputs()
                    self.build_static_tables()
                    self.add_queries()
                    self.load_all()
                    self.seed_tokens()
                    self.refresh_all()      # 표의 열 구조는 첫 새로 고침 후에 생기므로 수식보다 먼저 실행
                    self.settle_seeds()
                    self.exclude_button_queries()
                    if checkpoint:
                        self.wb.SaveCopyAs(os.path.abspath(checkpoint))
                        self.say(f"체크포인트 저장: {checkpoint}")
                self.ensure_rows()
                self.set_input_formulas()
                self.add_names()
                self.add_calc_columns()
                self.add_names_post()
                self.format_data_tables()
                self.build_dashboard()
                self.build_market()
                self.build_portfolio()
                self.build_risk()
                self.build_performance()
                self.build_quote_board()
                self.build_trade_sheets()
                self.build_universe_sheet()
                self.build_settings_sheet()
                self.build_guide()
                self.build_pages()
                self.add_vba()
                self.check_contract()
                self.report_counts()
                self.finish()
            finally:
                try:
                    if self.wb is not None:
                        self.wb.Close(False)
                except Exception:
                    pass
                self.wb, self.ws, self.lo = None, {}, {}
                quit_excel(self.xl, self.pid)
                self.xl = None
        self.vbom_message = vbom.message
        self.say(f"Excel 작업 시간 {time.time() - t0:.0f}s")

    # -------------------------------------------------------------------------------------------------------
    def create_workbook(self):
        os.makedirs(os.path.dirname(self.out), exist_ok=True)
        if os.path.exists(self.out):
            if os.path.normcase(self.out) == os.path.normcase(self.final_out):
                raise RuntimeError(f"작업 경로가 출력 통합문서와 같습니다(지우지 않음): {self.out}")
            os.remove(self.out)     # 직전에 실패한 빌드가 남긴 작업 파일(빌더 소유 — 사용자 통합문서가 아님)
        wb = self.xl.Workbooks.Add()
        self.wb = wb
        try:
            wb.EnableAutoRecover = False   # 빌드 중 자동 복구 사본(토큰 표 포함)이 남지 않게 — 저장 직전에 다시 켬
        except Exception:  # noqa: BLE001
            pass
        # 팔레트: Color10=상승(빨강), Color11=하락(파랑)  → 숫자서식 [Color10]/[Color11]
        set_palette(wb, 10, C["up"])
        set_palette(wb, 11, C["down"])
        wb.Queries.FastCombine = True  # 개인정보 수준 무시 (Formula.Firewall 방지)
        self.table_style = "PM Navy" if ensure_table_style(wb) is not None else "TableStyleLight9"
        st = wb.Styles("Normal")
        st.Font.Name = FONT
        st.Font.Size = 10
        while wb.Worksheets.Count < len(SHEETS):
            wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        for i, name in enumerate(SHEETS):
            wb.Worksheets(i + 1).Name = name
            self.ws[name] = wb.Worksheets(i + 1)
        # 먼저 최종 파일 이름(.xlsm)으로 저장해 두어야 차트의 이름 참조가 최종 파일명을 가리킴(작업 폴더만 다름)
        save_xlsm(wb, self.out)
        self.say(f"통합문서 생성: {self.final_out} (작업 위치 {self.out})")

    def open_checkpoint(self, path: str):
        """개발용: 새로 고침까지 끝난 체크포인트 파일을 열어 이후 단계만 다시 실행(이관·Seed 단계는 건너뜀)."""
        os.makedirs(os.path.dirname(self.out), exist_ok=True)
        shutil.copyfile(path, self.out)
        self.wb = self.xl.Workbooks.Open(self.out)
        self.table_style = "PM Navy" if ensure_table_style(self.wb) is not None else "TableStyleLight9"
        for ws in self.wb.Worksheets:
            self.ws[ws.Name] = ws
            for lo in ws.ListObjects:
                self.lo[lo.Name] = lo
                # 쿼리 표의 SourceType: 만들 때 0(xlSrcExternal), 저장 뒤 다시 열면 3(xlSrcQuery)
                m = re.search(r"\[(T_[^\]]+)\]", str(lo.QueryTable.CommandText)) if lo.SourceType in (0, 3) else None
                if m:
                    self.query_of[lo.Name] = m.group(1)
        self.say(f"체크포인트 열기: {path} (표 {len(self.lo)}개)")

    # -------------------------------------------------------------------------------------------------------
    def merged_settings(self) -> list[tuple]:
        """설정 표 행 (키, 항목, 값, 설명). 이관 원본이 있으면 기존 키는 사용자 값을 유지하고 새 키는 기본값을 더한다(docs/business-rules.md의 '사용자 데이터 보존').
        순서는 빌더 순서(숫자 서식·검증 목록이 키로 찾음), 원본에만 있는 키는 뒤에 그대로 붙인다(사용자 데이터 보존).
        cfg_path는 --cfg를 직접 준 경우에만 그 값으로 바꾼다."""
        base = settings_rows(self.cfg_path, self.sample)
        base_keys = [k for k, *_ in base]
        if self.plan is None:
            self.settings_info = {"source": 0, "kept": [], "added": base_keys, "extra": []}
            return base
        src = {k: (label, v, desc) for k, label, v, desc in self.plan.settings}
        out, kept, added = [], [], []
        for k, label, v, desc in base:
            if k in src:
                val = src[k][1]
                if k == "cfg_path" and self.cfg_explicit and val != self.cfg_path:
                    self.say(f"설정 cfg_path: --cfg로 준 경로를 씀(이관 값 대신): {self.cfg_path}")
                    val = self.cfg_path
                out.append((k, label, val, desc))
                kept.append(k)
            else:
                out.append((k, label, v, desc))
                added.append(k)
        extra = [(k, label, v, desc) for k, label, v, desc in self.plan.settings if k not in base_keys]
        out.extend(extra)
        self.settings_info = {"source": self.plan.settings_count, "kept": kept, "added": added, "extra": [k for k, *_ in extra]}
        return out

    def _migrated(self, name: str):
        """이관 계획의 입력표 행(원본에 그 표가 있었으면 0행이어도 목록), 원본에 없었으면 None(기본값 사용)."""
        if self.plan is None or name not in self.plan.tables:
            return None
        return self.plan.tables[name].rows

    def build_inputs(self):
        ws = self.ws["설정"]
        sheet_setup(ws, bg=False, zoom=90, tab="#6B7785",
                    widths={"A": 2, "B": 15, "C": 20, "D": 44, "E": 58, "F": 3, "G": 10, "H": 16, "I": 12, "J": 10,
                            "K": 10, "L": 44, "M": 3, "N": 8, "O": 16, "P": 9, "Q": 11, "R": 7, "S": 3, "T": 12, "U": 20})
        srows = self.merged_settings()
        lo = write_table(ws, 6, 2, ["키", "항목", "값", "설명"], protect_text(srows, ["키", "항목", "값", "설명"]), "tblSettings",
                         style_name=self.table_style)
        self.lo["tblSettings"] = lo
        format_dates(lo, srows, ["키", "항목", "값", "설명"])
        for i, (k, _, v, _) in enumerate(srows):
            c = lo.DataBodyRange.Cells(i + 1, 3)
            if k in ("start_date", "end_date"):
                set_nf(c, NF["date"])
            elif k in ("init_capital",):
                set_nf(c, NF["krw"])
            elif k in ("bm_kospi_weight", "rf_rate", "max_weight", "stop_loss", "take_profit", "move_alert", "target_vol"):
                set_nf(c, "0.0%")
            elif k in ("fee_rate", "tax_rate"):
                set_nf(c, "0.000%")
            elif k == "vol_alert":
                set_nf(c, NF["x"])
            c.HorizontalAlignment = XL_LEFT
        style(lo.ListColumns("키").DataBodyRange, color=C["muted"], size=8)
        fill(lo.ListColumns("값").DataBodyRange, "input")
        style(lo.ListColumns("설명").DataBodyRange, color=C["muted"], size=9)
        keys = [k for k, *_ in srows]
        for key, choices in (("benchmark", "KOSPI,KOSDAQ,혼합"), ("rank_market", "전체,코스피,코스닥"), ("stock_flow", "Y,N")):
            if key in keys:
                validation_list(lo.DataBodyRange.Cells(keys.index(key) + 1, 3), choices)

        mig = self._migrated("tblWatch")
        if mig is not None:
            wrows = protect_text([[c, None, g, a, t, p] for (c, g, a, t, p) in mig],
                                 ["종목코드", "종목명", "그룹", "관심가", "목표가", "투자포인트"], ("종목코드",))
        else:
            wrows = [[c, None, g, a, t, p] for (c, g, a, t, p) in (SAMPLE_WATCH if self.sample else [])]
        lo = write_table(ws, 6, 7, ["종목코드", "종목명", "그룹", "관심가", "목표가", "투자포인트"], wrows, "tblWatch",
                         text_cols=("종목코드",), style_name=self.table_style)
        format_dates(lo, wrows, ["종목코드", "종목명", "그룹", "관심가", "목표가", "투자포인트"])
        for c in ("관심가", "목표가"):
            set_nf(lo.ListColumns(c).DataBodyRange, NF["krw"])
        fill(lo.ListColumns("종목코드").DataBodyRange, "input")
        self.lo["tblWatch"] = lo

        mig = self._migrated("tblMacro")
        mrows = protect_text(mig, ["구분", "이름", "시장코드", "심볼", "순서"]) if mig is not None else MACRO_ROWS
        lo = write_table(ws, 6, 14, ["구분", "이름", "시장코드", "심볼", "순서"], mrows, "tblMacro",
                         style_name=self.table_style)
        format_dates(lo, mrows, ["구분", "이름", "시장코드", "심볼", "순서"])
        self.lo["tblMacro"] = lo
        mig = self._migrated("tblHolidays")
        hrows = protect_text(mig, ["휴장일", "설명"]) if mig is not None else HOLIDAYS
        lo = write_table(ws, 6, 20, ["휴장일", "설명"], hrows, "tblHolidays", date_cols=("휴장일",),
                         style_name=self.table_style)
        format_dates(lo, hrows, ["휴장일", "설명"], skip=("휴장일",))
        self.lo["tblHolidays"] = lo

        # ⑤ 수정표 tblOverride: ④ 오른쪽의 빈 열 묶음 W:AD(종목코드는 텍스트). 이관 원본의 행을 그대로 옮기고,
        # 없으면 빈 표(머리글 + 빈 행 1개). 제목·사용법·경고 칸은 build_settings_sheet
        anchor = ws.Range(OVERRIDE_ANCHOR)
        ws.Columns(anchor.Column - 1).ColumnWidth = 3
        for j, w in enumerate(OVERRIDE_WIDTHS):
            ws.Columns(anchor.Column + j).ColumnWidth = w
        mig = self._migrated("tblOverride")
        orows = protect_text(mig, OVERRIDE_HEADERS, ("종목코드",)) if mig else []
        lo = write_table(ws, anchor.Row, anchor.Column, OVERRIDE_HEADERS, orows, "tblOverride", text_cols=("종목코드",),
                         style_name=self.table_style)
        format_dates(lo, orows, OVERRIDE_HEADERS)
        fill(lo.DataBodyRange, "input")
        validation_list(lo.ListColumns("대회편입").DataBodyRange, "추가,제외")
        self.lo["tblOverride"] = lo

        # 매매일지 — 이관 원본이 있으면 입력 열만 옮긴다(종목명·금액·확인은 set_input_formulas가 수식으로 다시 넣음)
        ws = self.ws["매매일지"]
        sheet_setup(ws, bg=False, zoom=90, tab="#E8890C",
                    widths={"A": 2, "B": 11, "C": 9, "D": 16, "E": 6, "F": 8, "G": 11, "H": 13, "I": 9, "J": 9,
                            "K": 8, "L": 36, "M": 11, "N": 11, "O": 18, "P": 22})
        mig = self._migrated("tblTrades")
        if mig is not None:
            src_cols = migrate.INPUT_TABLES["tblTrades"]
            trades = [[r[src_cols.index(h)] if h in src_cols else None for h in TRADE_HEADERS] for r in mig]
            trades = protect_text(trades, TRADE_HEADERS, ("종목코드",))
        else:
            trades = sample_trades() if self.sample else []
        lo = write_table(ws, 11, 2, TRADE_HEADERS, trades, "tblTrades", text_cols=("종목코드",), date_cols=("일자",),
                         style_name=self.table_style)
        format_dates(lo, trades, TRADE_HEADERS, skip=("일자",))
        for c in ("수량",):
            set_nf(lo.ListColumns(c).DataBodyRange, NF["num0"])
        for c in ("단가", "금액", "수수료", "세금", "목표가", "손절가"):
            set_nf(lo.ListColumns(c).DataBodyRange, NF["krw"])
        validation_list(lo.ListColumns("구분").DataBodyRange, "매수,매도")
        validation_list(lo.ListColumns("전략").DataBodyRange, "코어,스윙,모멘텀,가치,이벤트,방어,역추세", show_error=False)
        for c in ("종목명", "금액", "확인"):
            style(lo.ListColumns(c).DataBodyRange, color=C["muted"])
            lo.ListColumns(c).Range.Cells(1, 1).Interior.Color = rgb("#8A94A3")
        self.lo["tblTrades"] = lo
        self.say("입력표 생성 완료 (설정·관심종목·해외지표·휴장일·수정표·매매일지)")

    # -------------------------------------------------------------------------------------------------------
    def build_static_tables(self):
        """숨김 시트 _seed의 정적 표(Power Query가 아니라 빌더가 쓰는 표 — 상태·조회시각 열 없음).

        tblRunCtl(실행 제어 키/값, mode = build) · tblContest(대회 명단, data CSV) · tblThemeBase(기본 테마표, data CSV) ·
        Seed 표(이관한 저장소 이력 — 자기참조 쿼리가 자기 표가 비었을 때 읽음). 표마다 열 1칸 띄워 2행에 나란히 둔다.
        Seed: tblSessionsSeed·tblPxStoreSeed·tblEstSnapSeed·tblCSLSeed·tblFinSeed는 늘(원본에 없으면 빈 표),
        tblFlowUSeed·tblTargetSeed·tblEstSeed·tblEventsSeed·tblSectorKRXSeed는 원본에 행이 있을 때만 만든다.
        """
        ws = self.ws["_seed"]
        specs = [("tblRunCtl", ["키", "값"], [list(r) for r in RUNCTL_ROWS], {"키": "text", "값": "general"})]
        cols, rows = read_csv_table(CONTEST_CSV, number_cols=("시가총액_0930_억", "평균거래대금_5일_억", "거래일수", "상장주식수"))
        specs.append(("tblContest", cols, rows, column_kinds(cols, rows, {"종목코드": "text", "종목명": "text", "시장": "text",
                                                                          "비고": "text"})))
        cols, rows = read_csv_table(THEMES_CSV)
        specs.append(("tblThemeBase", cols, rows, {c: "text" for c in cols}))
        for name in migrate.STORE_TABLES:
            t = self.plan.tables.get(name) if self.plan is not None else None
            if t is None and name == "tblEstSnap" and self.est_snap is not None:
                t = self.est_snap               # 이관 없이 만들 때 이력 CSV에서 복원한 스냅샷
            if name in migrate.OPTIONAL_SEEDS and (t is None or not t.rows):
                continue
            columns = list(t.columns) if t is not None else list(migrate.STORE_COLUMNS[name])
            data = list(t.rows) if t is not None else []
            forced = {c: "text" for c in ("종목코드", "상태") if c in columns}
            specs.append((name + "Seed", columns, data, column_kinds(columns, data, forced)))
            self.seeds[name + "Seed"] = (name, len(data), t.origin if t is not None else "")
        col = 2
        for name, columns, data, kinds in specs:
            self.lo[name] = write_static_table(ws, 2, col, name, columns, data, kinds)
            col += len(columns) + 1
        seeded = [f"{n}={c}행" for n, (_, c, _) in self.seeds.items() if c]
        self.say(f"정적 표 생성: tblRunCtl(mode=build) · tblContest {len(specs[1][2])}행 · tblThemeBase {len(specs[2][2])}행 · "
                 f"Seed {len(self.seeds)}개" + (f" (이관: {', '.join(seeded)})" if seeded else " (모두 빈 표)"))

    def add_queries(self):
        """powerquery/*.pq를 통합문서 쿼리로 넣는다. 보조 함수(fn*)를 의존 순서대로 먼저, 적재 쿼리(T_*)를 나중에 넣는다
        (아직 없는 함수를 참조하는 쿼리를 먼저 넣으면 함수가 들어올 때마다 참조 쿼리를 다시 검사해 느려짐)."""
        t0 = time.time()
        texts = {}
        for f in glob.glob(os.path.join(PQ_DIR, "*.pq")):
            with open(f, encoding="utf-8") as fh:
                texts[os.path.splitext(os.path.basename(f))[0]] = fh.read()
        for name in query_add_order(texts):
            add_query(self.wb, name, texts[name], QUERY_DESC.get(name, ""))
        self.say(f"Power Query {len(texts)}개 추가 ({time.time() - t0:.0f}s)")

    def load_all(self):
        """쿼리를 표로 적재: 기존 15개 + 새 일반·숨김 표(EXTRA_LOADS) + 페이지 모듈 표(각 모듈의 시트·셀).
        페이지 표는 첫 새로 고침 전에 새로 고침 방식을 덮어쓰기(RefreshStyle 0)로 바꾼다 — 기본값(셀 삽입·삭제)에서는
        0행 새로 고침이 셀을 지워 같은 열 아래의 표를 끌어올린다(페이지 작업 실측). 이 값은 이후 되돌리지 않는다."""
        for q, sheet, cell, tname in LOADS + EXTRA_LOADS:
            st = "TableStyleLight1" if sheet.startswith("_") else self.table_style
            self.lo[tname] = load_query(self.ws[sheet], q, cell, tname, background=True, style_name=st)
            self.query_of[tname] = q
        for mod in PAGE_MODULES:
            for q, tname, cell in mod.LOADS:
                self.lo[tname] = load_query(self.ws[mod.SHEET], q, cell, tname, background=True, style_name=self.table_style)
                self.query_of[tname] = q
        for mod in PAGE_MODULES:
            if hasattr(mod, "prepare_tables"):
                mod.prepare_tables(self)
        for _, tname, _ in page_news_events.LOADS:       # 이 모듈은 prepare_tables가 없어 같은 설정을 여기서
            self.lo[tname].QueryTable.RefreshStyle = 0    # xlOverwriteCells
        self.say(f"쿼리 → 표 로드 설정 완료 (표 {len(self.query_of)}개: 기존 {len(LOADS)} · 새 숨김/일반 {len(EXTRA_LOADS)} · "
                 f"페이지 {sum(len(m.LOADS) for m in PAGE_MODULES)})")

    def seed_tokens(self):
        """이관 토큰을 새 통합문서의 tblToken에 넣는다(T_Token 첫 새로 고침 전 — 그래야 재사용, 새 발급·알림톡 없음).
        T_Token 식을 잠시 임시 정적 표를 읽는 식으로 바꿔 한 번 새로 고친 뒤 원래 식으로 되돌리고 임시 표를 지운다.
        토큰 값은 메모리 → 임시 표 칸 → tblToken 칸으로만 가고, 로그에는 남은 분만 쓴다. 넣지 못하면 빌드를 멈춘다."""
        if self.seed_token:
            self.seed_token_cache()                     # 개발용 옛 경로(JSON 파일) — 쓰지 않는 것을 권장(정식 경로는 원천 통합문서를 파싱해 메모리로만 옮기기, docs/security.md)
        elif self.token is None:
            self.say("⚠ 이관할 토큰 캐시 없음 — 첫 새로 고침에서 T_Token이 새 토큰을 발급합니다(KIS 알림톡 1건)")
        else:
            tok = self.token
            cfg_sig = key_sig_from_cfg(self.cfg_path)
            sig = tok.key_sig or cfg_sig
            if tok.key_sig and cfg_sig and tok.key_sig != cfg_sig:
                self.say("⚠ 이관 토큰의 앱키 체크섬이 현재 kis_devlp.yaml과 다릅니다 — T_Token이 새 토큰을 발급할 수 있습니다")
            now = dt.datetime.now().replace(microsecond=0)
            row = ["prod", tok.token, tok.expires, tok.issued, sig, "이관(원본 파일 파싱)", now]
            kinds = {"env": "text", "token": "text", "expires": "datetime", "issued": "datetime", "key_sig": "text",
                     "status": "text", "checked": "datetime"}
            lo_seed = write_static_table(self.ws["_sys"], 2, 12, TOKEN_SEED_TABLE, migrate.TOKEN_COLUMNS, [row], kinds)
            q = self.wb.Queries("T_Token")
            original = q.Formula
            try:
                q.Formula = TOKEN_SEED_FORMULA
                ok, msg, n = refresh(self.lo["tblToken"], "T_Token(이관 토큰 넣기)")
            finally:
                q.Formula = original
                lo_seed.Delete()                         # 표와 칸 내용을 함께 지움
            if not ok or n != 1:
                raise RuntimeError(f"토큰 캐시 이관 실패({msg}) — T_Token이 새 토큰을 발급하지 않도록 빌드를 멈춥니다")
            self.say(f"  ✓ 토큰 캐시 이관: 남은 {tok.minutes_left()}분 (토큰 값은 표시하지 않음)")
        # 토큰 쿼리: 파일 열 때 + 30분마다 자동 확인(유효하면 재사용, 알림톡 없음) — 시드 뒤에 켠다
        qt = self.lo["tblToken"].QueryTable
        qt.RefreshOnFileOpen = True
        qt.RefreshPeriod = 30

    def ensure_rows(self):
        """0행인 PQ 표는 계산열 수식·숫자서식을 넣을 본문이 없으므로 임시 행을 추가 (finish에서 삭제).
        표가 비었다가 다시 채워져도 계산열 수식·서식은 표 정의에 남아 유지됩니다."""
        for q, sheet, cell, tname in LOADS:
            lo = self.lo[tname]
            if lo.ListRows.Count == 0:
                lo.ListRows.Add()
                self.temp_rows.append(tname)
        if self.temp_rows:
            self.say("빈 표 임시 행: " + ", ".join(self.temp_rows))

    def set_input_formulas(self):
        """입력표의 계산열 (PQ 표 tblUniverse가 만들어진 뒤에만 입력 가능)."""
        for t in ("tblTrades", "tblWatch"):
            rng = self.lo[t].ListColumns("종목코드").DataBodyRange
            for i in range(1, rng.Rows.Count + 1):
                try:
                    rng.Cells(i, 1).Errors(3).Ignore = True   # xlNumberAsText 표시 숨김
                except Exception:
                    pass
        self.lo["tblWatch"].ListColumns("종목명").DataBodyRange.Formula2 = F_WATCH_NAME
        lo = self.lo["tblTrades"]
        lo.ListColumns("종목명").DataBodyRange.Formula2 = F_TRADE_NAME
        lo.ListColumns("금액").DataBodyRange.Formula2 = F_TRADE_AMT
        lo.ListColumns("확인").DataBodyRange.Formula2 = F_TRADE_CHK

    def seed_token_cache(self):
        """개발용: 이미 발급받은 유효 토큰을 캐시에 넣어 빌드 중 재발급(알림톡)을 피함. 토큰 값은 출력하지 않음."""
        import json
        import yaml
        t = json.load(open(self.seed_token, encoding="utf-8"))
        cfg = yaml.safe_load(open(self.cfg_path, encoding="utf-8"))
        app = str(cfg["my_app"]).strip()
        sig = str(sum(ord(ch) for ch in app) * 7 + len(app))   # T_Token.pq 의 KeySig 와 동일한 식
        exp = dt.datetime.strptime(t["exp"], "%Y-%m-%d %H:%M:%S")
        now = dt.datetime.now()
        m = lambda d: f"#datetime({d.year},{d.month},{d.day},{d.hour},{d.minute},{d.second})"  # noqa: E731
        static = ('#table(type table [env = text, token = text, expires = datetime, issued = datetime, key_sig = text, '
                  'status = text, checked = datetime], {{"prod", "' + t["tok"] + '", ' + m(exp) + ', ' + m(now) + ', "' + sig +
                  '", "재사용", ' + m(now) + '}})')
        q = self.wb.Queries("T_Token")
        original = q.Formula
        q.Formula = static
        ok, msg, _ = refresh(self.lo["tblToken"], "T_Token(seed)")
        q.Formula = original
        self.say(("  ✓ " if ok else "  ✗ ") + "토큰 캐시 시드" + ("" if ok else f" 실패: {msg}"))

    def _key_value(self, table: str, key: str):
        lo = self.lo[table]
        body = lo.DataBodyRange
        if body is None:
            return None
        kc, vc = lo.ListColumns("키").Index, lo.ListColumns("값").Index
        for r in range(1, body.Rows.Count + 1):
            k = body.Cells(r, kc).Value
            if k is not None and str(k).strip() == key:
                return body.Cells(r, vc).Value
        return None

    def _token_status(self) -> str:
        lo = self.lo["tblToken"]
        try:
            v = lo.ListColumns("status").DataBodyRange.Cells(1, 1).Value
            return "" if v is None else str(v)
        except Exception:  # noqa: BLE001 — 상태 확인은 로그용
            return "(확인 못 함)"

    def refresh_all(self):
        """빌드 중 한 번 새로 고침(BUILD_REFRESH_ORDER). 일반 쿼리는 KIS를 조회하고, 버튼 전용 쿼리는 tblRunCtl mode = build라
        호출 없이 자기 표 → Seed → 빈 표를 돌려줘 열 구조가 생긴다(이관 이력은 이때 바로 표에 보임)."""
        mode = self._key_value("tblRunCtl", "mode")
        if str(mode or "").strip().lower() != "build":
            raise RuntimeError(f"tblRunCtl mode가 build가 아닙니다({mode!r}) — 버튼 전용 쿼리가 KIS를 부르지 않게 멈춥니다")
        self.say("데이터 새로 고침 — 일반 쿼리는 KIS 조회, 버튼 전용 쿼리는 build 모드(KIS 호출 없음)…")
        t0 = time.time()
        missing = [t for t in BUILD_REFRESH_ORDER if t not in self.query_of]
        if missing:
            self.say("⚠ 새로 고침 순서에 있는데 적재되지 않은 표: " + ", ".join(missing))
        order = [t for t in BUILD_REFRESH_ORDER if t in self.query_of] + [t for t in self.query_of if t not in BUILD_REFRESH_ORDER]
        failed = []
        for tname in order:
            q = self.query_of[tname]
            ok, msg, n = refresh(self.lo[tname], q + (" [build]" if q in BUTTON_QUERIES else ""))
            self.say(("  ✓ " if ok else "  ✗ ") + msg)
            if not ok:
                failed.append((q, tname))
            if tname == "tblToken":
                st = self._token_status()
                self.say(f"    토큰 상태: {st}")
                if self.token is not None and st.startswith("신규"):
                    self.say("⚠ 이관 토큰을 재사용하지 못하고 새로 발급됨(알림톡) — 원본 토큰·앱키 확인")
        for attempt in range(2):          # 일시 오류(호출 제한 등) 대비 재시도
            if not failed:
                break
            time.sleep(15)
            retry, failed = failed, []
            for q, tname in retry:
                ok, msg, n = refresh(self.lo[tname], q)
                self.say(("  ✓ (재시도) " if ok else "  ✗ (재시도) ") + msg)
                if not ok:
                    failed.append((q, tname))
        if failed:
            raise RuntimeError("새로 고침 실패: " + ", ".join(q for q, _ in failed) +
                               " — kis_devlp.yaml(my_app/my_sec/prod)과 네트워크를 확인하세요.")
        self.say(f"새로 고침 완료 ({time.time() - t0:.0f}s)")

    def _empty_static(self, lo):
        """정적 표를 머리글 + 빈 행 1개로 줄인다(아래 칸은 지움). 같은 행의 다른 표에는 손대지 않는다."""
        ws = lo.Range.Worksheet
        top, left = lo.HeaderRowRange.Row, lo.HeaderRowRange.Column
        ncol, n = lo.ListColumns.Count, lo.ListRows.Count
        if n > 1:
            lo.Resize(ws.Range(ws.Cells(top, left), ws.Cells(top + 1, left + ncol - 1)))
            ws.Range(ws.Cells(top + 2, left), ws.Cells(top + n, left + ncol - 1)).Clear()
        if lo.DataBodyRange is not None:
            lo.DataBodyRange.ClearContents()

    def settle_seeds(self):
        """Seed로 넣은 이관 이력이 build 모드 새로 고침으로 자기 표에 옮겨졌으면 Seed를 비운다(같은 이력을 두 벌 들고 있지 않게 —
        가격 저장소는 한 벌에 약 9MB). 자기 표가 비었거나 오류 1행뿐이면 Seed를 그대로 두고 경고한다(다음 새로 고침에서 다시 읽음).
        이관한 행은 백업 폴더의 원본에도 남아 있다."""
        for seed, (own, n_seed, _origin) in self.seeds.items():
            if not n_seed:
                continue
            lo_own = self.lo.get(own)
            n_own = 0 if lo_own is None or lo_own.DataBodyRange is None else int(lo_own.ListRows.Count)
            only_error = False
            if n_own == 1:
                try:
                    st = lo_own.ListColumns("상태").DataBodyRange.Cells(1, 1).Value
                    only_error = isinstance(st, str) and st.startswith(("오류", "이전"))
                except Exception:  # noqa: BLE001
                    pass
            if n_own > 0 and not only_error:
                self._empty_static(self.lo[seed])
                self.say(f"  Seed 정리: {seed} {n_seed}행 → {own} {n_own}행으로 옮겨져 Seed를 비움")
            else:
                self.say(f"⚠ {own}이(가) 비어 있어 {seed}({n_seed}행)를 그대로 둡니다 — 버튼 실행 때 다시 읽습니다")

    def exclude_button_queries(self):
        """버튼 전용 쿼리를 '모두 새로 고침'에서 뺀다(연결 속성 '모두 새로 고침 시 이 연결 새로 고침' 해제)·파일 열 때
        새로 고침 없음. 일반 쿼리(기존 + T_News + T_MktFunds)는 포함 그대로. T_Token의 파일 열 때·30분 확인은 유지."""
        n_ex, n_in = 0, 0
        for tname, q in self.query_of.items():
            qt = self.lo[tname].QueryTable
            conn = qt.WorkbookConnection
            if q in BUTTON_QUERIES:
                conn.RefreshWithRefreshAll = False
                qt.RefreshOnFileOpen = False
                try:
                    conn.OLEDBConnection.RefreshOnFileOpen = False
                except Exception:  # noqa: BLE001 — QueryTable 쪽 설정과 같은 값(연결 종류에 따라 없을 수 있음)
                    pass
                n_ex += 1
            else:
                conn.RefreshWithRefreshAll = True
                if q != "T_Token":
                    qt.RefreshOnFileOpen = False
                n_in += 1
        self.say(f"모두 새로 고침: 일반 쿼리 {n_in}개 포함 · 버튼 전용 {n_ex}개 제외 (파일 열 때 새로 고침은 T_Token만)")

    # -------------------------------------------------------------------------------------------------------
    def add_calc_columns(self):
        q = "tblQuote"
        lk = lambda col: f'IFERROR(XLOOKUP([@종목코드],tblQuote[종목코드],tblQuote[{col}]),"")'  # noqa: E731
        ph = lambda k, col: f'IFERROR(XLOOKUP([@종목코드]&"|{k}",tblPriceHist[키],tblPriceHist[{col}]),"")'  # noqa: E731
        rk = lambda col: f'IFERROR(XLOOKUP([@종목코드],tblRisk[종목코드],tblRisk[{col}]),"")'  # noqa: E731

        # --- 보유종목 (평균단가 뒤 / 목표가 뒤에 끼워 넣어 핵심 열이 앞쪽에 오도록)
        lo = self.lo["tblHoldings"]
        hsig = ('=TEXTJOIN(" · ",TRUE,IF([@현재가]<=[@손절가],"손절이탈",IF([@현재가]<=[@손절가]*1.03,"손절근접","")),'
                'IF([@현재가]>=[@목표가],"목표도달",""),IF([@비중]>최대비중,"비중초과",""),'
                'IF(ABS([@등락률])>=급등락기준,IF([@등락률]>0,"급등","급락"),""),'
                'IF(N([@RSI14])>=70,"과매수",IF(AND(ISNUMBER([@RSI14]),N([@RSI14])<=30),"과매도","")),'
                'IF(AND(N([@외인5일억])>0,N([@기관5일억])>0),"외인·기관 순매수",""),[@유의])')
        cols = [
            ("현재가", '=IFERROR(XLOOKUP([@종목코드],tblQuote[종목코드],tblQuote[현재가]),[@평균단가])', NF["krw"], "평균단가"),
            ("등락률", '=IFERROR(XLOOKUP([@종목코드],tblQuote[종목코드],tblQuote[등락률]),0)', NF["pct_arrow"], "현재가"),
            ("평가금액", "=[@보유수량]*[@현재가]", NF["krw"], "등락률"),
            ("비중", "=IFERROR([@평가금액]/순자산,0)", NF["pct1"], "평가금액"),
            ("평가손익", "=[@평가금액]-[@매입금액]", NF["krw_pl"], "비중"),
            ("수익률", "=IFERROR([@평가손익]/[@매입금액],0)", NF["pct_pl"], "평가손익"),
            ("일간손익", '=[@보유수량]*IFERROR(XLOOKUP([@종목코드],tblQuote[종목코드],tblQuote[전일대비]),0)', NF["krw_pl"], "수익률"),
            ("기여도", "=IFERROR([@일간손익]/(순자산-SUM(tblHoldings[일간손익])),0)", NF["pct_pl"], "일간손익"),
            ("손절여유", '=IFERROR([@현재가]/[@손절가]-1,"")', NF["pct1"], "목표가"),
            ("목표여유", '=IFERROR([@목표가]/[@현재가]-1,"")', NF["pct1"], "손절여유"),
            ("신호", hsig, None, "목표여유"),
            ("RSI14", "=" + ph(0, "RSI14"), NF["num1"], "신호"),
            ("이격도20", "=" + ph(0, "이격도20"), NF["pct_pl"], "RSI14"),
            ("변동성", "=" + rk("변동성"), NF["pct1"], "이격도20"),
            ("베타", "=" + rk("베타"), NF["ratio2"], "변동성"),
            ("위험기여", "=" + rk("위험기여비중"), NF["pct1"], "베타"),
            ("고52주대비", "=" + lk("고52주대비"), NF["pct1"], "위험기여"),
            ("외인5일억", "=" + lk("외국인5일억"), NF["eok_pl"], "고52주대비"),
            ("기관5일억", "=" + lk("기관5일억"), NF["eok_pl"], "외인5일억"),
            ("유의", "=" + lk("유의"), None, "기관5일억"),
        ]
        # 신호 열이 뒤쪽 열(RSI14 등)을 참조하므로 열을 모두 만든 뒤 수식을 입력
        for name, f, nf, after in cols:
            add_calc_column(lo, name, '=""', None, after=after)
        for name, f, nf, after in cols:
            rng = lo.ListColumns(name).DataBodyRange
            rng.Formula2 = f
            if nf:
                set_nf(rng, nf)

        # --- 손익 귀속
        lo = self.lo["tblPositions"]
        for name, f, nf in [
            ("현재가", '=IF([@상태]="보유",IFERROR(XLOOKUP([@종목코드],tblQuote[종목코드],tblQuote[현재가]),[@평균단가]),"")', NF["krw"]),
            ("평가금액", '=IF([@상태]="보유",[@보유수량]*[@현재가],0)', NF["krw"]),
            ("평가손익", '=IF([@상태]="보유",[@평가금액]-[@매입금액],0)', NF["krw_pl"]),
            ("총손익", "=N([@실현손익])+[@평가손익]", NF["krw_pl"]),
            ("기여도", "=IFERROR([@총손익]/초기자금,0)", NF["pct_pl"]),
        ]:
            add_calc_column(lo, name, f, nf)

        # --- 시세판 (등락률 뒤: 신호·추세·RSI…, 기관20일억 뒤: 관심가·목표가…)
        lo = self.lo[q]
        qsig = ('=TEXTJOIN(" · ",TRUE,IF(AND(ISNUMBER([@고52주대비]),N([@고52주대비])>=-0.03),"52주고가근접",""),'
                'IF(AND(ISNUMBER([@거래량비율]),N([@거래량비율])>=거래량배수),"거래량급증",""),[@교차],'
                'IF(N([@RSI14])>=70,"과매수",IF(AND(ISNUMBER([@RSI14]),N([@RSI14])<=30),"과매도","")),'
                'IF(AND(ISNUMBER([@관심가]),[@현재가]<=N([@관심가])),"관심가도달",""),'
                'IF(AND(ISNUMBER([@목표가]),[@현재가]>=N([@목표가])),"목표가도달",""),'
                'IF(AND(N([@외국인5일억])>0,N([@기관5일억])>0),"외인·기관 순매수",""),'
                'IF(ABS(N([@등락률]))>=급등락기준,IF([@등락률]>0,"급등","급락"),""),[@유의])')
        cross = ('=LET(k,[@종목코드],x5_0,XLOOKUP(k&"|0",tblPriceHist[키],tblPriceHist[MA5],""),'
                 'x20_0,XLOOKUP(k&"|0",tblPriceHist[키],tblPriceHist[MA20],""),x5_1,XLOOKUP(k&"|1",tblPriceHist[키],tblPriceHist[MA5],""),'
                 'x20_1,XLOOKUP(k&"|1",tblPriceHist[키],tblPriceHist[MA20],""),IF(COUNT(x5_0,x20_0,x5_1,x20_1)<4,"",'
                 'IF(AND(x5_0>x20_0,x5_1<=x20_1),"골든크로스",IF(AND(x5_0<x20_0,x5_1>=x20_1),"데드크로스",""))))')
        qcols = [
            ("신호", qsig, None, "등락률"),
            ("추세", '=IF(OR([@MA20]="",[@MA60]=""),"",IF(AND([@현재가]>[@MA20],[@MA20]>[@MA60]),"상승",'
                    'IF(AND([@현재가]<[@MA20],[@MA20]<[@MA60]),"하락","혼조")))', None, "신호"),
            ("RSI14", "=" + ph(0, "RSI14"), NF["num1"], "추세"),
            ("거래량비율", '=IFERROR([@거래량]/XLOOKUP([@종목코드]&"|1",tblPriceHist[키],tblPriceHist[거래량20평균]),"")', NF["x"], "RSI14"),
            ("수익률20일", '=IFERROR([@현재가]/XLOOKUP([@종목코드]&"|20",tblPriceHist[키],tblPriceHist[종가])-1,"")', NF["pct_pl"], "거래량비율"),
            ("수익률60일", '=IFERROR([@현재가]/XLOOKUP([@종목코드]&"|60",tblPriceHist[키],tblPriceHist[종가])-1,"")', NF["pct_pl"], "수익률20일"),
            ("관심가", '=LET(v,XLOOKUP([@종목코드],tblWatch[종목코드],tblWatch[관심가],""),IF(AND(ISNUMBER(v),v>0),v,""))', NF["krw"], "기관20일억"),
            ("목표가", '=LET(v,XLOOKUP([@종목코드],tblWatch[종목코드],tblWatch[목표가],""),IF(AND(ISNUMBER(v),v>0),v,""))', NF["krw"], "관심가"),
            ("투자포인트", '=LET(v,XLOOKUP([@종목코드],tblWatch[종목코드],tblWatch[투자포인트],""),IF(v=0,"",v))', None, "목표가"),
            ("MA20", "=" + ph(0, "MA20"), NF["krw"], "투자포인트"),
            ("MA60", "=" + ph(0, "MA60"), NF["krw"], "MA20"),
            ("변동성20", "=" + ph(0, "변동성20"), NF["pct1"], "MA60"),
            ("교차", cross, None, "변동성20"),
        ]
        for name, f, nf, after in qcols:
            add_calc_column(lo, name, '=""', None, after=after)
        for name, f, nf, after in qcols:
            rng = lo.ListColumns(name).DataBodyRange
            rng.Formula2 = f
            if nf:
                set_nf(rng, nf)
        self.say("계산열 추가 완료 (보유종목·손익귀속·시세판)")

    def add_names(self):
        add_names(self.wb, NAMES)

    def add_names_post(self):
        add_names(self.wb, NAMES_POST)

    # -------------------------------------------------------------------------------------------------------
    def format_data_tables(self):
        """PQ 표 숫자서식 (새로 고침 후에도 유지: PreserveFormatting)."""
        fmt = {
            "tblIndexNow": {"현재": NF["num2"], "전일대비": NF["num2"], "등락률": NF["pct_arrow"], "거래대금억": NF["num0"],
                            "연중고가일": NF["date"], "연중저가일": NF["date"], "조회시각": NF["dt"]},
            "tblIndexHist": {"일자": NF["date"], "등락률": NF["pct_pl"], "대회누적": NF["pct_pl"]},
            "tblSector": {"등락률": NF["pct_arrow"], "거래비중": NF["pct1"], "조회시각": NF["dt"]},
            "tblFlow": {"일자": NF["date"]},
            "tblRank": {"등락률": NF["pct_arrow"], "조회시각": NF["dt"]},
            "tblGlobal": {"일자": NF["date"], "등락률": NF["pct_pl"]},
            "tblPriceHist": {"일자": NF["date"], "등락률": NF["pct_pl"]},
            "tblPositions": {"최초매수일": NF["date"], "최근매매일": NF["date"], "실현손익": NF["krw_pl"], "평균단가": NF["krw"],
                             "매입금액": NF["krw"], "매수금액": NF["krw"], "매도금액": NF["krw"]},
            "tblToken": {"expires": NF["dt"], "issued": NF["dt"], "checked": NF["dt"]},
        }
        for t, m in fmt.items():
            lo = self.lo[t]
            for c, nf in m.items():
                try:
                    set_col_format(lo, c, nf=nf)
                except Exception:
                    pass
        # 숨김 시트 (B1 = 안내 한 줄, 표는 2행부터)
        notes = {
            "_data": ("※ Power Query 원본 데이터 (수정하지 마세요)", C["muted"]),
            "_sys": ("※ KIS 접근토큰 캐시 — 통합문서를 외부에 공유할 때는 이 표의 내용을 지우세요(24시간 유효)", C["up"]),
            "_calc": ("※ 버튼 전용 데이터(분류·세션 달력·가격 지표·수급·실적·목표주가·추정·스냅샷·신용/공매도 — [시세]·[전체]가 "
                      "갱신)와 증시 자금(tblMktFunds — [모두 새로 고침]). 수정하지 마세요", C["muted"]),
            "_store": ("※ 가격 이력 저장소(일봉) — [시세]·[전체]가 이어 씁니다. 수정하지 마세요", C["muted"]),
            "_seed": ("※ 빌더가 만든 정적 표(실행 제어·대회 명단·기본 테마표·이관 초기값 Seed) — 수정하지 마세요", C["muted"]),
        }
        for s in HIDDEN_SHEETS:
            ws = self.ws[s]
            ws.Cells.Font.Name = FONT
            ws.Cells.Font.Size = 9
            text, color = notes.get(s, ("※ 숨김 데이터 (수정하지 마세요)", C["muted"]))
            put(ws, "B1", text, size=9, color=color)
        self.ws["_sys"].Range("C:C").ColumnWidth = 12

    # -------------------------------------------------------------------------------------------------------
    def nav_links(self, ws, row: int, start_col: int = 2):
        """상단 메뉴 줄: 보이는 시트 15개(SHEETS 순서 — 새 페이지 업종·대회종목·종목분석·뉴스·이벤트 포함)를 2열 간격으로.
        링크가 늘어 기존 메뉴 띠보다 길어지면 마지막 링크 뒤 칸까지 띠 색을 칠한다(밝은 글자가 흰 바탕에 묻히지 않게).
        대회종목 페이지는 이 링크를 읽어 자기 열 폭에 맞게 다시 놓는다(page_company._nav_row)."""
        c = start_col
        for sheet in NAV_ITEMS:
            cell = ws.Cells(row, c).Address
            hyperlink(ws, cell, f"'{sheet}'!A1", "› " + sheet)
            c += 2
        fill(ws.Range(ws.Cells(row, start_col), ws.Cells(row, c - 1)), "navy2")

    HEADER_RIGHT = ('="기준 "&IFERROR(TEXT(MAX(tblIndexNow[조회시각]),"yyyy-mm-dd hh:mm"),"-")&"  ·  토큰 "&'
                    'IFERROR(XLOOKUP("prod",tblToken[env],tblToken[status]),"-")')
    HEADER_STATUS = ('=LET(e,COUNTIF(tblQuote[상태],"오류*")+COUNTIF(tblIndexNow[상태],"오류*")+COUNTIF(tblGlobal[상태],"오류*")'
                     '+COUNTIF(tblRank[종목명],"오류*")+COUNTIF(tblPriceHist[상태],"오류*"),'
                     's,COUNTIF(tblQuote[상태],"이전*")+COUNTIF(tblIndexNow[상태],"이전*")+COUNTIF(tblGlobal[상태],"이전*")'
                     '+COUNTIF(tblPriceHist[상태],"이전*"),'
                     'tk,IFERROR(XLOOKUP("prod",tblToken[env],tblToken[token]),""),'
                     'IF(tk="","▲ 접근토큰 없음 — [모두 새로 고침] 한 번 더",'
                     'IF(s>0,"▲ 갱신 실패 — 이전 데이터 표시 중 (다시 새로 고침)",'
                     'IF(e=0,"● 데이터 정상","▲ 조회 오류 "&e&"건 — 다시 새로 고침"))))')

    # -------------------------------------------------------------------------------------------------------
    def build_dashboard(self):
        ws = self.ws["대시보드"]
        sheet_setup(ws, zoom=85, tab="navy", default_width=9.3, widths={"A": 1.5, "Z": 1.5})
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "PM DAILY DASHBOARD", "KOSPI · KOSDAQ 개별종목 모의투자 | 한국투자증권 Open API × Excel Power Query",
                  "B2:Y3", right_formula=self.HEADER_RIGHT, right_cell="Y2", right2_formula=self.HEADER_STATUS,
                  right2_cell="Y3")
        fill(ws.Range("B4:Y4"), "navy2")
        self.nav_links(ws, 4)
        ws.Rows(4).RowHeight = 16
        ws.Rows(5).RowHeight = 8

        # KPI 카드 8개 (행 6~9)
        spans = ["B6:D9", "E6:G9", "H6:J9", "K6:M9", "N6:P9", "Q6:S9", "T6:V9", "W6:Y9"]
        for r in (7, 8):
            ws.Rows(r).RowHeight = 17
        card(ws, spans[0], "총자산 (순자산)", "=순자산",
             '="초기 대비 "&TEXT(순자산-초기자금,"+#,##0;-#,##0")&"원"', nf=NF["krw"])
        card(ws, spans[1], "일간 손익", '=IFERROR(TAKE(tblNAV[일간손익],-1),0)',
             '=IFERROR(TEXT(TAKE(tblNAV[일간수익률],-1),"+0.00%;-0.00%")&"  ("&TEXT(TAKE(tblNAV[일자],-1),"mm/dd")&" 기준)","-")',
             nf=NF["krw_pl"])
        card(ws, spans[2], "누적 수익률", F_CUM,
             '=IFERROR("BM "&TEXT(TAKE(tblNAV[BM누적],-1),"+0.00%;-0.00%")&" · 초과 "&TEXT(TAKE(tblNAV[누적초과],-1),"+0.00%p;-0.00%p"),"-")',
             nf=NF["pct_pl"])
        card(ws, spans[3], "샤프지수 (연율화)", F_SHARPE,
             '="rf "&TEXT(무위험수익률,"0.00%")&" · rf=0: "&IFERROR(TEXT(' + F_SHARPE0[1:] + ',"0.00"),"-")', nf=NF["ratio2"])
        card(ws, spans[4], "변동성 (연율화)", F_VOL,
             '="목표 "&TEXT(목표변동성,"0%")&" · 사전 "&IFERROR(TEXT(XLOOKUP("PORT",tblRisk[종목코드],tblRisk[변동성]),"0.0%"),"-")',
             nf=NF["pct1"])
        card(ws, spans[5], "최대낙폭 (MDD)", F_MDD, '="현재 낙폭 "&IFERROR(TEXT(TAKE(tblNAV[낙폭],-1),"0.00%"),"-")',
             nf=NF["pct"], value_color=C["down"])
        card(ws, spans[6], "현금 비중", "=IFERROR(현금잔고/순자산,0)",
             '="보유 "&COUNTIF(tblHoldings[보유수량],">0")&"종목 · 현금 "&TEXT(현금잔고/10000,"#,##0")&"만원"', nf=NF["pct1"])
        card(ws, spans[7], "대회 진행", '="D+"&' + F_DAYS[1:],
             '="남은 거래일 "&' + F_REMAIN[1:] + '&"일 · ~"&TEXT(종료일,"mm/dd")', nf=None)

        ws.Rows(10).RowHeight = 8
        # 시장 스트립 (행 11~13): 12개 × 2열
        items = [
            ("KOSPI", 'XLOOKUP("KOSPI",tblIndexNow[지수],tblIndexNow[현재])', 'XLOOKUP("KOSPI",tblIndexNow[지수],tblIndexNow[등락률])', NF["num2"], NF["pct_arrow"]),
            ("KOSDAQ", 'XLOOKUP("KOSDAQ",tblIndexNow[지수],tblIndexNow[현재])', 'XLOOKUP("KOSDAQ",tblIndexNow[지수],tblIndexNow[등락률])', NF["num2"], NF["pct_arrow"]),
            ("KOSPI200", 'XLOOKUP("KOSPI200",tblIndexNow[지수],tblIndexNow[현재])', 'XLOOKUP("KOSPI200",tblIndexNow[지수],tblIndexNow[등락률])', NF["num2"], NF["pct_arrow"]),
            ("VKOSPI", 'XLOOKUP("VKOSPI",tblIndexNow[지수],tblIndexNow[현재])', 'XLOOKUP("VKOSPI",tblIndexNow[지수],tblIndexNow[등락률])', NF["num2"], NF["pct_arrow"]),
            ("외국인(코스피)", 'LET(d,MAXIFS(tblFlow[일자],tblFlow[시장],"KOSPI"),SUMIFS(tblFlow[외국인],tblFlow[시장],"KOSPI",tblFlow[일자],d))',
             'SUM(TAKE(FILTER(tblFlow[외국인],tblFlow[시장]="KOSPI"),5))', NF["eok_pl"], '"5일 "+#,##0"억";"5일 "-#,##0"억"'),
            ("기관(코스피)", 'LET(d,MAXIFS(tblFlow[일자],tblFlow[시장],"KOSPI"),SUMIFS(tblFlow[기관계],tblFlow[시장],"KOSPI",tblFlow[일자],d))',
             'SUM(TAKE(FILTER(tblFlow[기관계],tblFlow[시장]="KOSPI"),5))', NF["eok_pl"], '"5일 "+#,##0"억";"5일 "-#,##0"억"'),
        ]
        for nm in ("원/달러", "S&P500", "NASDAQ", "필라델피아반도체", "미국10년", "국고3년"):
            v = f'XLOOKUP(1,(tblGlobal[이름]="{nm}")*(tblGlobal[최신여부]="Y"),tblGlobal[종가])'
            if nm in ("미국10년", "국고3년"):
                chg = f'XLOOKUP(1,(tblGlobal[이름]="{nm}")*(tblGlobal[최신여부]="Y"),tblGlobal[전일대비])*100'
                items.append((nm, v, chg, '0.000"%"', NF["bp_pl"]))
            else:
                chg = f'XLOOKUP(1,(tblGlobal[이름]="{nm}")*(tblGlobal[최신여부]="Y"),tblGlobal[등락률])'
                items.append(("SOX" if nm == "필라델피아반도체" else nm, v, chg, NF["num2"], NF["pct_arrow"]))
        for i, (label, fv, fc, nfv, nfc) in enumerate(items):
            c0 = 2 + i * 2
            rng = ws.Range(ws.Cells(11, c0), ws.Cells(13, c0 + 1))
            fill(rng, "card")
            border(rng, "line")
            put(ws, ws.Cells(11, c0).Address, label, size=8, color=C["muted"], indent=1)
            put(ws, ws.Cells(12, c0).Address, formula=f"=IFERROR({fv},\"-\")", nf=nfv, bold=True, size=11, indent=1)
            put(ws, ws.Cells(13, c0).Address, formula=f"=IFERROR({fc},\"\")", nf=nfc, size=8, indent=1)
        ws.Rows(14).RowHeight = 8

        # 좌: 누적수익률 차트 / 우: 보유종목 요약
        section(ws, "B15:M15", "누적 수익률 vs 벤치마크")
        section(ws, "N15:Y15", "보유 종목 (비중순)", "상세는 [포트폴리오]", "Y15")
        hdrs = [("N", "종목"), ("P", "비중"), ("Q", "일간"), ("R", "수익률"), ("S", "평가손익"), ("U", "손절여유"), ("V", "신호")]
        for col, h in hdrs:
            put(ws, f"{col}16", h, bold=True, size=9, color=C["muted"])
        bottom_line(ws.Range("N16:Y16"))
        key = "tblHoldings[비중]"
        spill = {
            "N": "tblHoldings[종목명]", "P": "tblHoldings[비중]", "Q": "tblHoldings[등락률]", "R": "tblHoldings[수익률]",
            "S": "tblHoldings[평가손익]", "U": "tblHoldings[손절여유]", "V": "tblHoldings[신호]",
        }
        nfs = {"P": NF["pct1"], "Q": NF["pct_arrow"], "R": NF["pct_pl"], "S": NF["krw_pl"], "U": NF["pct1"], "V": None, "N": None}
        for col, ref in spill.items():
            put(ws, f"{col}17", formula=f'=IFERROR(TAKE(SORTBY({ref},{key},-1),14),"")', nf=nfs[col], size=9)
            set_nf(ws.Range(f"{col}17:{col}30"), nfs[col] or "General")
            ws.Range(f"{col}17:{col}30").Font.Size = 9
        style(ws.Range("V17:V30"), color=C["warn"], size=8)
        fill(ws.Range("N16:Y30"), "card")
        cf_expr(ws.Range("N17:Y30"), '=ISNUMBER(SEARCH("손절",$V17))', fill_color=C["up_l"])
        cf_expr(ws.Range("N17:Y30"), '=ISNUMBER(SEARCH("목표",$V17))', fill_color=C["good_l"])

        nav = self.lo["tblNAV"]
        x = nav.ListColumns("일자").DataBodyRange
        shp, ch = chart(ws, "B16:M30", XL_LINE)
        series(ch, "포트폴리오", x, nav.ListColumns("누적수익률").DataBodyRange, color=C["accent"], weight=2.5)
        series(ch, "벤치마크", x, nav.ListColumns("BM누적").DataBodyRange, color="#8A94A3", weight=1.75)
        series(ch, "KOSDAQ", x, nav.ListColumns("KOSDAQ누적").DataBodyRange, color="#E8890C", weight=1.0)
        style_axes(ch, y_nf="0.0%", x_nf="mm/dd", legend=XL_LEGEND_TOP)

        ws.Rows(31).RowHeight = 8
        # 알림 / 섹터 / 낙폭
        section(ws, "B32:J32", "오늘의 알림")
        put(ws, "B33", formula=self.alerts_formula(), size=9)
        ws.Range("B33:J47").Font.Size = 9
        fill(ws.Range("B33:J47"), "card")
        cf_expr(ws.Range("B33:J47"), '=LEFT($B33,1)="▼"', font=C["up"], bold=True)
        cf_expr(ws.Range("B33:J47"), '=LEFT($B33,1)="▲"', font=C["good"], bold=True)
        cf_expr(ws.Range("B33:J47"), '=LEFT($B33,1)="◆"', font=C["warn"])

        section(ws, "K32:Q32", "섹터 비중")
        shp, ch = chart(ws, "K33:Q47", XL_BAR_CLUSTERED)
        s = series(ch, "비중", f"='{self.wb.Name}'!차트_섹터명", f"='{self.wb.Name}'!차트_섹터비중", fill_color=C["accent"])
        style_axes(ch, y_nf="0%", category=False, legend=None)
        ch.Axes(1).ReversePlotOrder = True
        s.HasDataLabels = True
        set_nf(s.DataLabels(), "0%")
        s.DataLabels().Font.Size = 8

        section(ws, "R32:Y32", "낙폭 (Drawdown)")
        shp, ch = chart(ws, "R33:Y47", XL_AREA)
        series(ch, "낙폭", x, nav.ListColumns("낙폭").DataBodyRange, fill_color="#5B8FD9", transparency=0.35)
        style_axes(ch, y_nf="0.0%", x_nf="mm/dd", legend=None)

        ws.Rows(48).RowHeight = 8
        # 시장 주도주 3종 — 시장 전체 개별종목. 종목명 옆 칸 ★ = 대회 종목, 대회 밖은 회색 글꼴,
        # 보유·관심 파란 굵은 글씨가 회색보다 우선(규칙 우선순위 맨 앞). 종목명 칸 값은 그대로(보유·관심 강조가 이름으로 찾음)
        blocks = [("B", "거래대금 상위", "거래대금상위", "tblRank[거래대금억]", NF["eok"], "거래대금"),
                  ("J", "외국인 순매수 상위", "외국인순매수", "tblRank[지표]", NF["eok_pl"], "순매수"),
                  ("R", "상승률 상위 (ETF·SPAC 제외)", "상승률상위", "tblRank[거래대금억]", NF["eok"], "거래대금")]
        for col, title, lst, extra, nf_extra, extra_h in blocks:
            c0 = ws.Range(f"{col}49").Column
            section(ws, ws.Range(ws.Cells(49, c0), ws.Cells(49, c0 + 7)).Address, title)
            for off, head in ((0, "#"), (1, "종목"), (3, "대회"), (4, "현재가"), (5, "등락률"), (6, extra_h)):
                put(ws, ws.Cells(50, c0 + off).Address, head, bold=True, size=8, color=C["muted"],
                    h=XL_CENTER if off == 3 else None)
            bottom_line(ws.Range(ws.Cells(50, c0), ws.Cells(50, c0 + 7)))
            cond = f'tblRank[목록]="{lst}"'
            refs = [(0, "tblRank[순위]", "0"), (1, "tblRank[종목명]", None), (3, RANK_STAR, None),
                    (4, "tblRank[현재가]", NF["krw"]), (5, "tblRank[등락률]", NF["pct_arrow"]), (6, extra, nf_extra)]
            for off, ref, nf in refs:
                cell = ws.Cells(51, c0 + off)
                put(ws, cell.Address, formula=f'=IFERROR(TAKE(FILTER({ref},{cond}),10),"")', size=9)
                set_nf(ws.Range(ws.Cells(51, c0 + off), ws.Cells(60, c0 + off)), nf or "General")
                ws.Range(ws.Cells(51, c0 + off), ws.Cells(60, c0 + off)).Font.Size = 9
            fill(ws.Range(ws.Cells(50, c0), ws.Cells(60, c0 + 7)), "card")
            style(ws.Range(ws.Cells(51, c0), ws.Cells(60, c0)), color=C["muted"], h=XL_CENTER)
            style(ws.Range(ws.Cells(51, c0 + 3), ws.Cells(60, c0 + 3)), color=C["gold"], h=XL_CENTER)
            # 보유·관심 종목 강조(파란 굵은 글씨) → 대회 밖 회색(#·종목·대회 칸) → 파랑을 맨 앞 우선순위로
            first = ws.Cells(51, c0 + 1).Address.replace("$", "")
            blue = cf_expr(ws.Range(ws.Cells(51, c0 + 1), ws.Cells(60, c0 + 3)),
                           f'=COUNTIF(tblQuote_종목명,{first})>0', font=C["accent"], bold=True)
            cf_expr(ws.Range(ws.Cells(51, c0), ws.Cells(60, c0 + 3)),
                    f'=AND(${col_letter(c0 + 1)}51<>"",${col_letter(c0 + 3)}51<>"★")', font=NONCONTEST_GREY)
            blue.SetFirstPriority()
        put(ws, "B62", "※ 파란 굵은 글씨 = 보유·관심 종목 | ★ = 대회 종목 · 회색 = 대회 종목 아님 | 상승=빨강·하락=파랑 (국내 관례) | "
                       "데이터: 한국투자증권 Open API", size=8, color=C["muted"])
        freeze(ws, 4)
        self.say("대시보드 시트 완료")

    def alerts_formula(self) -> str:
        return (
            '=LET('
            'hn,tblHoldings[종목명],hp,tblHoldings[현재가],hs,tblHoldings[손절가],ht,tblHoldings[목표가],hw,tblHoldings[비중],'
            'hc,tblHoldings[등락률],hf,tblHoldings[유의],'
            'qn,tblQuote[종목명],qt,tblQuote[구분],qp,tblQuote[현재가],qb,tblQuote[관심가],qs,tblQuote[신호],'
            'al_1,FILTER("▼ 손절선 이탈 · "&hn&"  "&TEXT(hp,"#,##0")&" ≤ "&TEXT(hs,"#,##0"),hp<=hs,""),'
            'al_2,FILTER("◆ 손절선 3% 이내 · "&hn&"  "&TEXT(hp/hs-1,"0.0%"),(hp>hs)*(hp<=hs*1.03),""),'
            'al_3,FILTER("▲ 목표가 도달 · "&hn&"  "&TEXT(hp,"#,##0")&" ≥ "&TEXT(ht,"#,##0"),hp>=ht,""),'
            'al_4,FILTER("◆ 비중 한도 초과 · "&hn&"  "&TEXT(hw,"0.0%")&" > "&TEXT(최대비중,"0%"),hw>최대비중,""),'
            'al_5,FILTER("◆ 급변동 · "&hn&"  "&TEXT(hc,"+0.0%;-0.0%"),IFERROR(ABS(hc),0)>=급등락기준,""),'
            'al_6,FILTER("◆ 유의 종목 · "&hn&" ("&hf&")",hf<>"",""),'
            'al_7,FILTER("▲ 관심가 도달 · "&qn&"  "&TEXT(qp,"#,##0"),(qt="관심")*ISNUMBER(qb)*(qp<=IFERROR(qb+0,0)),""),'
            'al_8,FILTER("● 관심종목 신호 · "&qn&" — "&qs,(qt="관심")*(qs<>""),""),'
            'pv,IFERROR(XLOOKUP("PORT",tblRisk[종목코드],tblRisk[변동성]),0),dd,IFERROR(TAKE(tblNAV[낙폭],-1),0),'
            'cs,IFERROR(현금잔고/순자산,0),'
            'pf_1,IF(pv>목표변동성,"◆ 사전 변동성 "&TEXT(pv,"0%")&" > 목표 "&TEXT(목표변동성,"0%")&" — 비중·종목 수 점검",""),'
            'pf_2,IF(dd<-0.1,"▼ 고점 대비 낙폭 "&TEXT(dd,"0.0%")&" — 리스크 축소 검토",""),'
            'pf_3,IF(cs<0.03,"◆ 현금 비중 3% 미만 — 추가 매수 여력 부족",""),'
            'all,VSTACK(al_1,al_2,al_3,al_4,al_5,al_6,pf_1,pf_2,pf_3,al_7,al_8),'
            'ok,FILTER(all,all<>""),'
            'IFERROR(TAKE(ok,15),"● 특이사항 없음"))'
        )

    # -------------------------------------------------------------------------------------------------------
    def build_market(self):
        ws = self.ws["시장"]
        sheet_setup(ws, zoom=85, tab="accent", default_width=9.3, widths={"A": 1.5, "Z": 1.5})
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "MARKET MONITOR", "지수 · 시장폭 · 투자자 수급 · 업종 · 해외/환율/금리 · 주도주", "B2:Y3",
                  right_formula=self.HEADER_RIGHT, right_cell="Y2", right2_formula=self.HEADER_STATUS, right2_cell="Y3")
        fill(ws.Range("B4:Y4"), "navy2")
        self.nav_links(ws, 4)
        ws.Rows(5).RowHeight = 8

        # 지수 카드 4개 (각 6열) 행 6~10
        idx = ["KOSPI", "KOSDAQ", "KOSPI200", "VKOSPI"]
        for i, nm in enumerate(idx):
            c0 = 2 + i * 6
            rng = ws.Range(ws.Cells(6, c0), ws.Cells(10, c0 + 5))
            fill(rng, "card")
            border(rng, "line")
            lk = lambda col: f'XLOOKUP("{nm}",tblIndexNow[지수],tblIndexNow[{col}])'  # noqa: E731
            put(ws, ws.Cells(6, c0).Address, nm, bold=True, size=10, color=C["navy"], indent=1)
            put(ws, ws.Cells(7, c0).Address, formula=f'=IFERROR({lk("현재")},"-")', nf=NF["num2"], bold=True, size=16, indent=1)
            put(ws, ws.Cells(7, c0 + 3).Address, formula=f'=IFERROR({lk("등락률")},"")', nf=NF["pct_arrow"], bold=True, size=11)
            put(ws, ws.Cells(8, c0).Address, formula=f'=IFERROR("전일대비 "&TEXT({lk("전일대비")},"+#,##0.00;-#,##0.00")&"  ·  거래대금 "&TEXT({lk("거래대금억")}/10000,"#,##0.0")&"조","")',
                size=8, color=C["muted"], indent=1)
            if nm != "VKOSPI":
                put(ws, ws.Cells(9, c0).Address,
                    formula=f'=IFERROR("상승 "&{lk("상승")}&" (상한 "&{lk("상한")}&")  ·  하락 "&{lk("하락")}&" (하한 "&{lk("하한")}&")  ·  보합 "&{lk("보합")},"")',
                    size=8, color=C["muted"], indent=1)
            put(ws, ws.Cells(10, c0).Address,
                formula=f'=IFERROR("연중 고 "&TEXT({lk("연중고가")},"#,##0.00")&" ("&TEXT({lk("연중고가일")},"mm/dd")&")  저 "&TEXT({lk("연중저가")},"#,##0.00"),"")',
                size=8, color=C["muted"], indent=1)
            ws.Rows(7).RowHeight = 24
        ws.Rows(11).RowHeight = 8

        # 지수 차트용 도우미 (최근 60영업일 누적 등락률) — 숨김 열 AB:AE
        put(ws, "AB5", "차트 도우미(수정 금지)", size=8, color=C["muted"])
        put(ws, "AB6", formula='=TAKE(FILTER(tblIndexHist[일자],tblIndexHist[지수]="KOSPI"),-60)', nf=NF["mmdd"])
        for col, nm in (("AC", "KOSPI"), ("AD", "KOSDAQ"), ("AE", "KOSPI200")):
            put(ws, f"{col}6", formula=f'=IFERROR(LET(v,XLOOKUP("{nm}|"&TEXT(AB6#,"yyyy-mm-dd"),tblIndexHist[키],tblIndexHist[종가]),v/INDEX(v,1)-1),NA())')
        set_nf(ws.Range("AB6:AB65"), NF["mmdd"])
        ws.Columns("AB:AE").Hidden = True
        section(ws, "B12:M12", "KOSPI · KOSDAQ · KOSPI200 — 최근 60영업일 누적 등락률")
        shp, ch = chart(ws, "B13:M28", XL_LINE)
        series(ch, "KOSPI", ws.Range("AB6:AB65"), ws.Range("AC6:AC65"), color=C["up"], weight=2)
        series(ch, "KOSDAQ", ws.Range("AB6:AB65"), ws.Range("AD6:AD65"), color=C["down"], weight=2)
        series(ch, "KOSPI200", ws.Range("AB6:AB65"), ws.Range("AE6:AE65"), color="#8A94A3", weight=1.25)
        style_axes(ch, y_nf="0%", x_nf="mm/dd", legend=XL_LEGEND_TOP)

        # 투자자 수급 표
        section(ws, "N12:Y12", "투자자별 순매수 (억원)", "당일 / 5일 / 20일 누적", "Y12")
        hdr = ["시장", "기간", "외국인", "기관계", "개인", "연기금", "금융투자", "투신"]
        cols = ["N", "O", "P", "R", "T", "V", "W", "X"]
        for c, h in zip(cols, hdr):
            put(ws, f"{c}13", h, bold=True, size=8, color=C["muted"])
        bottom_line(ws.Range("N13:Y13"))
        r = 14
        for mk in ("KOSPI", "KOSDAQ"):
            for label, n in (("당일", 1), ("5일", 5), ("20일", 20)):
                put(ws, f"N{r}", mk if n == 1 else "", bold=True, size=9)
                put(ws, f"O{r}", label, size=9, color=C["muted"])
                for c, fld in zip(cols[2:], ["외국인", "기관계", "개인", "연기금", "금융투자", "투신"]):
                    put(ws, f"{c}{r}", formula=f'=IFERROR(SUM(TAKE(FILTER(tblFlow[{fld}],tblFlow[시장]="{mk}"),{n})),"")',
                        nf=NF["krw_pl"], size=9)
                r += 1
        fill(ws.Range("N13:Y19"), "card")
        # 수급 차트 (KOSPI 20일)
        put(ws, "AG5", "수급 도우미", size=8, color=C["muted"])
        put(ws, "AG6", formula='=SORT(TAKE(FILTER(tblFlow[일자],tblFlow[시장]="KOSPI"),20))', nf=NF["mmdd"])
        for c, fld in (("AH", "외국인"), ("AI", "기관계"), ("AJ", "개인")):
            put(ws, f"{c}6", formula=f'=IFERROR(XLOOKUP("KOSPI|"&TEXT(AG6#,"yyyy-mm-dd"),tblFlow[키],tblFlow[{fld}]),0)')
        set_nf(ws.Range("AG6:AG25"), NF["mmdd"])
        ws.Columns("AG:AJ").Hidden = True
        shp, ch = chart(ws, "N20:Y28", XL_COLUMN_CLUSTERED, "KOSPI 투자자 순매수 (최근 20일, 억원)")
        series(ch, "외국인", ws.Range("AG6:AG25"), ws.Range("AH6:AH25"), fill_color=C["up"])
        series(ch, "기관계", ws.Range("AG6:AG25"), ws.Range("AI6:AI25"), fill_color=C["down"])
        series(ch, "개인", ws.Range("AG6:AG25"), ws.Range("AJ6:AJ25"), fill_color="#B0B8C4")
        style_axes(ch, y_nf="#,##0", x_nf="mm/dd", legend=XL_LEGEND_TOP)

        ws.Rows(29).RowHeight = 8
        # 업종 히트맵
        for i, mk in enumerate(("KOSPI", "KOSDAQ")):
            c0 = 2 + i * 6
            section(ws, ws.Range(ws.Cells(30, c0), ws.Cells(30, c0 + 4)).Address, f"{mk} 업종 등락")
            for off, h in ((0, "업종"), (2, "지수"), (3, "등락률"), (4, "거래비중")):
                put(ws, ws.Cells(31, c0 + off).Address, h, bold=True, size=8, color=C["muted"])
            bottom_line(ws.Range(ws.Cells(31, c0), ws.Cells(31, c0 + 4)))
            cond = f'(tblSector[시장]="{mk}")*(tblSector[구분]="업종")'
            key = f'FILTER(tblSector[등락률],{cond})'
            for off, fld, nf in ((0, "업종명", None), (2, "지수", NF["num2"]), (3, "등락률", NF["pct_arrow"]), (4, "거래비중", NF["pct1"])):
                put(ws, ws.Cells(32, c0 + off).Address, formula=f'=IFERROR(SORTBY(FILTER(tblSector[{fld}],{cond}),{key},-1),"")', size=9)
                set_nf(ws.Range(ws.Cells(32, c0 + off), ws.Cells(58, c0 + off)), nf or "General")
                ws.Range(ws.Cells(32, c0 + off), ws.Cells(58, c0 + off)).Font.Size = 9
            fill(ws.Range(ws.Cells(31, c0), ws.Cells(58, c0 + 4)), "card")
            cf_heat(ws.Range(ws.Cells(32, c0 + 3), ws.Cells(58, c0 + 3)))
            cf_databar(ws.Range(ws.Cells(32, c0 + 4), ws.Cells(58, c0 + 4)), color="#9DB7E0")
        # 해외·환율·금리
        section(ws, "N30:Y30", "해외 지수 · 환율 · 금리", "최근 종가 기준", "Y30")
        for c, h in (("N", "구분"), ("O", "지표"), ("Q", "현재"), ("S", "전일대비"), ("U", "등락률"), ("W", "기준일")):
            put(ws, f"{c}31", h, bold=True, size=8, color=C["muted"])
        bottom_line(ws.Range("N31:Y31"))
        g = 'tblGlobal[최신여부]="Y"'
        for c, fld, nf in (("N", "구분", None), ("O", "이름", None), ("Q", "종가", NF["num2"]), ("S", "전일대비", "+#,##0.00;-#,##0.00"),
                           ("U", "등락률", NF["pct_arrow"]), ("W", "일자", NF["mmdd"])):
            put(ws, f"{c}32", formula=f'=IFERROR(SORTBY(FILTER(tblGlobal[{fld}],{g}),FILTER(tblGlobal[순서],{g}),1),"")', size=9)
            set_nf(ws.Range(f"{c}32:{c}47"), nf or "General")
            ws.Range(f"{c}32:{c}47").Font.Size = 9
        fill(ws.Range("N31:Y47"), "card")
        cf_expr(ws.Range("N32:Y47"), '=$N32="금리"', font=C["navy2"])
        put(ws, "N49", "※ 금리 행의 등락률은 금리 자체의 변화율, 전일대비는 %p", size=8, color=C["muted"])

        ws.Rows(59).RowHeight = 8
        # 주도주 순위 6블록 (2열 × 3행, 블록당 12열) — 시장 전체 개별종목(ETF·ETN·SPAC 제외).
        # 열: # · 종목(이름 칸 + 넘침 칸) · 대회(★) · 섹터(4칸 폭, 설정 섹터 기준 이름이 길어도 보이게) · 현재가 · 등락률 · 지표 · 여백.
        # 3블록 × 8열로는 ★ 칸을 넣으면 종목명이나 섹터가 한 칸으로 줄어 잘리므로 2열 배치로 바꿨다.
        blocks = [
            ("거래대금상위", "거래대금 상위", "tblRank[거래대금억]", NF["eok"], "거래대금"),
            ("거래량급증", "거래량 급증", "tblRank[지표]", '0%', "증가율"),
            ("상승률상위", "상승률 상위", "tblRank[거래대금억]", NF["eok"], "거래대금"),
            ("하락률상위", "하락률 상위", "tblRank[거래대금억]", NF["eok"], "거래대금"),
            ("외국인순매수", "외국인 순매수 (가집계)", "tblRank[지표]", NF["eok_pl"], "순매수"),
            ("기관순매수", "기관 순매수 (가집계)", "tblRank[지표]", NF["eok_pl"], "순매수"),
        ]
        n_rank = 15
        for bi, (lst, title, extra, nf_extra, extra_h) in enumerate(blocks):
            row0 = MKT_RANK_TOP + (bi // 2) * (n_rank + 3)
            c0 = 2 + (bi % 2) * 12
            last = row0 + 1 + n_rank
            section(ws, ws.Range(ws.Cells(row0, c0), ws.Cells(row0, c0 + 10)).Address, title)
            cols = ((0, "#", "tblRank[순위]", "0"), (1, "종목", "tblRank[종목명]", None), (3, "대회", RANK_STAR, None),
                    (4, "섹터", "tblRank[섹터]", None), (8, "현재가", "tblRank[현재가]", NF["krw"]),
                    (9, "등락률", "tblRank[등락률]", NF["pct_arrow"]), (10, extra_h, extra, nf_extra))
            for off, head, ref, nf in cols:
                put(ws, ws.Cells(row0 + 1, c0 + off).Address, head, bold=True, size=8, color=C["muted"],
                    h=XL_CENTER if off == 3 else None)
                put(ws, ws.Cells(row0 + 2, c0 + off).Address, formula=f'=IFERROR(TAKE(FILTER({ref},tblRank[목록]="{lst}"),{n_rank}),"")',
                    size=9)
                rr = ws.Range(ws.Cells(row0 + 2, c0 + off), ws.Cells(last, c0 + off))
                set_nf(rr, nf or "General")
                rr.Font.Size = 9
            bottom_line(ws.Range(ws.Cells(row0 + 1, c0), ws.Cells(row0 + 1, c0 + 10)))
            fill(ws.Range(ws.Cells(row0 + 1, c0), ws.Cells(last, c0 + 10)), "card")
            style(ws.Range(ws.Cells(row0 + 2, c0), ws.Cells(last, c0)), color=C["muted"], h=XL_CENTER)
            style(ws.Range(ws.Cells(row0 + 2, c0 + 3), ws.Cells(last, c0 + 3)), color=C["gold"], h=XL_CENTER)
            style(ws.Range(ws.Cells(row0 + 2, c0 + 4), ws.Cells(last, c0 + 4)), color=C["muted"], size=8)
            # 보유·관심 파란 굵은 글씨(종목명) → 대회 밖 회색(#·종목·대회·섹터 칸) → 파랑을 맨 앞 우선순위로
            r2 = row0 + 2
            first = ws.Cells(r2, c0 + 1).Address.replace("$", "")
            blue = cf_expr(ws.Range(ws.Cells(r2, c0 + 1), ws.Cells(last, c0 + 2)),
                           f'=COUNTIF(tblQuote_종목명,{first})>0', font=C["accent"], bold=True)
            cf_expr(ws.Range(ws.Cells(r2, c0), ws.Cells(last, c0 + 4)),
                    f'=AND(${col_letter(c0 + 1)}{r2}<>"",${col_letter(c0 + 3)}{r2}<>"★")', font=NONCONTEST_GREY)
            blue.SetFirstPriority()
        note_row = MKT_RANK_TOP + 3 * (n_rank + 3)
        put(ws, f"B{note_row}", "※ 시장 전체 개별종목(ETF·ETN·SPAC 제외) | ★ = 대회 종목 · 회색 = 대회 종목 아님 | 파란 굵은 글씨 = 보유·관심 종목 | "
                                "외국인/기관 순매수는 장중 가집계(09:30·11:20·13:20·14:30 입력) 기준", size=8, color=C["muted"])

        # 증시 자금 동향 — tblMktFunds: 일자 오름차순 약 100영업일, 금액 억원(빈칸 가능·0 없음), 공표 1~2일 늦음.
        # 항목마다 비어 있지 않은 값만으로 최신값과 1·5·20개 전 값의 차(1D·1W·1M)를 구함. 잔고는 조원, 증감은 억원.
        top = MKT_FUNDS_TOP
        ws.Rows(top - 1).RowHeight = 8
        section(ws, f"B{top}:Y{top}", "증시 자금 동향 (금융투자협회 집계)")
        put(ws, f"Y{top}", formula='="기준일 "&IFERROR(TEXT(MAX(tblMktFunds[일자]),"yyyy-mm-dd"),"-")&"  ·  공표가 1~2일 늦음"',
            size=8, color=C["muted"], h=XL_RIGHT)
        hdr = top + 1
        for c, head in (("B", "항목"), ("D", "잔고(조원)"), ("E", "1D(억)"), ("F", "1W(억)"), ("G", "1M(억)"), ("H", "1M 증감률"),
                        ("I", "기준일")):
            put(ws, f"{c}{hdr}", head, bold=True, size=8, color=C["muted"], h=XL_LEFT if c == "B" else XL_RIGHT)
        bottom_line(ws.Range(f"B{hdr}:L{hdr}"))
        funds = [("고객예탁금", "고객예탁금억"), ("신용융자 잔고", "신용융자억"), ("미수금", "미수금억"),
                 ("주식형 펀드(평가액)", "주식형억"), ("MMF", "MMF억")]
        for i, (label, col) in enumerate(funds):
            r = hdr + 1 + i
            v = f'FILTER(tblMktFunds[{col}],ISNUMBER(tblMktFunds[{col}]))'
            diff = lambda k, v=v: f'=IFERROR(LET(v_s,{v},n_s,ROWS(v_s),IF(n_s>{k},INDEX(v_s,n_s)-INDEX(v_s,n_s-{k}),"")),"")'  # noqa: E731
            put(ws, f"B{r}", label, bold=True, size=9)
            put(ws, f"D{r}", formula=f'=IFERROR(LET(v_s,{v},INDEX(v_s,ROWS(v_s))/10000),"-")', nf="#,##0.0", size=9)
            put(ws, f"E{r}", formula=diff(1), nf=NF["krw_pl"], size=9)
            put(ws, f"F{r}", formula=diff(5), nf=NF["krw_pl"], size=9)
            put(ws, f"G{r}", formula=diff(20), nf=NF["krw_pl"], size=9)
            put(ws, f"H{r}", formula=f'=IFERROR(LET(v_s,{v},n_s,ROWS(v_s),IF(n_s>20,INDEX(v_s,n_s)/INDEX(v_s,n_s-20)-1,"")),"")',
                nf=NF["pct_pl"], size=9)
            put(ws, f"I{r}", formula=f'=IFERROR(LET(d_s,FILTER(tblMktFunds[일자],ISNUMBER(tblMktFunds[{col}])),INDEX(d_s,ROWS(d_s))),"")',
                nf=NF["mmdd"], size=9, h=XL_RIGHT)
            bottom_line(ws.Range(f"B{r}:L{r}"), "line2")
        n0 = hdr + len(funds) + 2
        for j, t in enumerate(("※ 1D·1W·1M = 1·5·20영업일 전 공표값과의 차이. 공표가 1~2일 늦어 기준일이 오늘보다 앞섭니다.",
                               "※ 주식형 펀드 잔고는 평가액이라 주가 등락에 따라 움직입니다(자금 유출입과 다름).",
                               "※ MMF는 분기 말에 줄었다가 돌아오는 계절성이 있습니다. [모두 새로 고침]에서 갱신(KIS 호출 1회).")):
            put(ws, f"B{n0 + j}", t, size=8, color=C["muted"])
        fill(ws.Range(f"B{hdr}:L{top + 15}"), "card")
        # 60영업일 추이 차트 도우미 (숨김 AL:AN). 두 잔고의 수준이 약 3배 차이(고객예탁금 ≈100조, 신용융자 ≈30조)라
        # 0부터 시작하는 이중 축에서는 두 선이 납작해지므로, 창의 첫날 대비 변화율로 한 축에 그린다(지수 60일 누적 등락률
        # 차트와 같은 방식, 잔고 수준은 왼쪽 표). 빈 값은 NA()로 끊어 그림
        put(ws, "AL5", "자금 도우미", size=8, color=C["muted"])
        put(ws, "AL6", formula='=IFERROR(TAKE(tblMktFunds[일자],-60),NA())', nf=NF["mmdd"])
        for hc_col, fld in (("AM", "고객예탁금억"), ("AN", "신용융자억")):
            put(ws, f"{hc_col}6", formula=f'=IFERROR(LET(v_m,TAKE(tblMktFunds[{fld}],-60),b_m,INDEX(FILTER(v_m,ISNUMBER(v_m)),1),'
                                          f'IF(ISNUMBER(v_m),v_m/b_m-1,NA())),NA())')
        set_nf(ws.Range("AL6:AL65"), NF["mmdd"])
        ws.Columns("AL:AN").Hidden = True
        shp, ch = chart(ws, f"N{hdr}:Y{top + 15}", XL_LINE, "고객예탁금 · 신용융자 잔고 추이 (최근 60영업일, 첫날 대비 %)")
        series(ch, "고객예탁금", ws.Range("AL6:AL65"), ws.Range("AM6:AM65"), color=C["accent"], weight=2)
        series(ch, "신용융자 잔고", ws.Range("AL6:AL65"), ws.Range("AN6:AN65"), color=C["warn"], weight=1.75)
        style_axes(ch, y_nf="0%", x_nf="mm/dd", legend=XL_LEGEND_TOP)
        try:
            ch.Axes(1).TickLabelSpacing = 5     # 날짜 60개 → 5영업일 간격 눈금 이름(겹침 방지)
        except Exception:  # noqa: BLE001 — 축 서식 실패는 차트 자체에 영향 없음
            pass
        freeze(ws, 4)
        self.say("시장 시트 완료")

    # -------------------------------------------------------------------------------------------------------
    def build_portfolio(self):
        ws = self.ws["포트폴리오"]
        sheet_setup(ws, bg=False, zoom=85, tab="good", widths={"A": 1.5})
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "PORTFOLIO", "보유 종목 · 평가손익 · 손절/목표 · 기술지표 · 수급 · 위험기여", "B2:AR3",
                  right_formula=self.HEADER_RIGHT, right_cell="Q2", right2_formula=self.HEADER_STATUS, right2_cell="Q3")
        fill(ws.Range("B4:AR4"), "navy2")
        self.nav_links(ws, 4)
        lo = self.lo["tblHoldings"]
        widths = {"종목코드": 8, "종목명": 15, "시장": 8, "섹터": 12, "보유수량": 8, "평균단가": 11, "매입금액": 13, "실현손익": 12,
                  "진입일": 10, "보유일수": 7, "손절가": 10, "목표가": 10, "전략": 7, "매매근거": 22, "확인": 8, "현재가": 11,
                  "등락률": 9, "평가금액": 13, "비중": 7, "평가손익": 12, "수익률": 8, "일간손익": 11, "기여도": 8, "손절여유": 8,
                  "목표여유": 8, "RSI14": 7, "이격도20": 8, "변동성": 7, "베타": 6, "위험기여": 8, "고52주대비": 9, "외인5일억": 10,
                  "기관5일억": 10, "유의": 10, "신호": 30}
        for c, w in widths.items():
            set_col_format(lo, c, width=w)
        for c, nf in {"보유수량": NF["num0"], "평균단가": NF["krw"], "매입금액": NF["krw"], "실현손익": NF["krw_pl"],
                      "진입일": NF["date"], "보유일수": "0", "손절가": NF["krw"], "목표가": NF["krw"]}.items():
            set_col_format(lo, c, nf=nf)
        set_col_format(lo, "종목명", bold=True)
        # 카드 (행 6~9)
        spans = ["B6:D9", "E6:G9", "H6:J9", "K6:M9", "N6:P9", "Q6:S9", "T6:V9", "W6:Y9"]
        card(ws, spans[0], "순자산", "=순자산", '="초기 "&TEXT(초기자금,"#,##0")', nf=NF["krw"])
        card(ws, spans[1], "주식 평가액", "=주식평가액", '="주식비중 "&TEXT(IFERROR(주식평가액/순자산,0),"0.0%")', nf=NF["krw"])
        card(ws, spans[2], "현금", "=현금잔고", '="현금비중 "&TEXT(IFERROR(현금잔고/순자산,0),"0.0%")', nf=NF["krw"])
        card(ws, spans[3], "평가손익 (보유)", "=SUM(tblHoldings[평가손익])",
             '=IFERROR(TEXT(SUM(tblHoldings[평가손익])/SUM(tblHoldings[매입금액]),"+0.00%;-0.00%")&" (매입가 대비)","-")', nf=NF["krw_pl"])
        card(ws, spans[4], "실현손익 (누적)", "=SUM(tblPositions[실현손익])", '="청산 포함 전체 종목"', nf=NF["krw_pl"])
        card(ws, spans[5], "총손익", "=SUM(tblPositions[총손익])", '=TEXT(IFERROR(SUM(tblPositions[총손익])/초기자금,0),"+0.00%;-0.00%")&" (초기자금 대비)"',
             nf=NF["krw_pl"])
        card(ws, spans[6], "오늘 손익 (보유분)", "=SUM(tblHoldings[일간손익])", '="전일종가 대비 · 보유수량 기준"', nf=NF["krw_pl"])
        card(ws, spans[7], "포트 베타 · 종목수", '=IFERROR(XLOOKUP("PORT",tblRisk[종목코드],tblRisk[베타]),"-")',
             '=COUNTIF(tblHoldings[보유수량],">0")&"종목 · 최대비중 "&TEXT(IFERROR(MAX(tblHoldings[비중]),0),"0.0%")', nf=NF["ratio2"])
        section(ws, "B11:Y11", "보유 종목", "손절가/목표가는 매매일지 입력값(없으면 설정의 기본 %) · 신호는 자동 계산", "Y11")
        # 조건부 서식
        body = lo.DataBodyRange
        if body is not None:
            first_row = body.Row
            sig_col = lo.ListColumns("신호").Range.Column
            sig = f"${col_letter(sig_col)}{first_row}"
            rng = ws.Range(ws.Cells(first_row, lo.Range.Column), ws.Cells(first_row + 199, lo.Range.Column + lo.ListColumns.Count - 1))
            cf_expr(rng, f'=ISNUMBER(SEARCH("손절이탈",{sig}))', fill_color=C["up_l"])
            cf_expr(rng, f'=ISNUMBER(SEARCH("목표도달",{sig}))', fill_color=C["good_l"])
            w_col = lo.ListColumns("비중").Range.Column
            cf_databar(ws.Range(ws.Cells(first_row, w_col), ws.Cells(first_row + 199, w_col)))
            style(lo.ListColumns("신호").DataBodyRange, color=C["warn"], size=9)
        freeze(ws, 12, 3)
        self.say("포트폴리오 시트 완료")

    # -------------------------------------------------------------------------------------------------------
    def build_risk(self):
        ws = self.ws["리스크"]
        sheet_setup(ws, bg=False, zoom=85, tab="good", default_width=9.3, widths={"A": 1.5})
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "RISK", "현재 비중 기준 사전 위험 — 변동성 · 베타 · 위험기여 · VaR · 섹터 집중도", "B2:AF3",
                  right_formula=self.HEADER_RIGHT, right_cell="AF2", right2_formula=self.HEADER_STATUS, right2_cell="AF3")
        fill(ws.Range("B4:AF4"), "navy2")
        self.nav_links(ws, 4)
        port = lambda col: f'IFERROR(XLOOKUP("PORT",tblRisk[종목코드],tblRisk[{col}]),"-")'  # noqa: E731
        spans = ["B6:D9", "E6:G9", "H6:J9", "K6:M9", "N6:P9", "Q6:S9", "T6:V9", "W6:Y9"]
        card(ws, spans[0], "사전 변동성 (연)", "=" + port("변동성"), '="목표 "&TEXT(목표변동성,"0%")&" · 시장 "&TEXT(' + port("시장변동성") + ',"0%")', nf=NF["pct1"])
        card(ws, spans[1], "포트 베타 (KOSPI)", "=" + port("베타"), '="1.0 = 시장과 동일한 민감도"', nf=NF["ratio2"])
        card(ws, spans[2], "VaR 95% (1일, 모수)", "=" + port("VaR95"), '=IFERROR(TEXT(' + port("VaR95") + '/순자산,"0.00%")&" of NAV","-")', nf=NF["krw"])
        card(ws, spans[3], "VaR 95% (1일, 과거)", "=" + port("과거VaR95"), '="최근 "&' + port("관측일수") + '&"영업일 가상수익률"', nf=NF["krw"])
        card(ws, spans[4], "최대 위험기여 종목", '=IFERROR(INDEX(SORTBY(FILTER(tblRisk[이름],tblRisk[구분]="종목"),FILTER(tblRisk[위험기여비중],tblRisk[구분]="종목"),-1),1),"-")',
             '="위험의 "&IFERROR(TEXT(MAX(FILTER(tblRisk[위험기여비중],tblRisk[구분]="종목")),"0%"),"-")&" 차지"', value_size=13)
        card(ws, spans[5], "상위 3종목 비중", '=IFERROR(SUM(TAKE(SORT(FILTER(tblRisk[비중],tblRisk[구분]="종목"),,-1),3)),0)',
             '="집중도 (HHI "&TEXT(IFERROR(SUMSQ(FILTER(tblRisk[비중],tblRisk[구분]="종목")),0),"0.000")&")"', nf=NF["pct1"])
        card(ws, spans[6], "최대 섹터 비중", '=IFERROR(MAX(DROP(INDEX(리스크!T13#,0,2),-1)),0)', '=IFERROR("섹터: "&INDEX(리스크!T13#,1,1),"-")', nf=NF["pct1"])
        card(ws, spans[7], "분산 효과", '=IFERROR(1-' + port("변동성") + '/SUMPRODUCT(FILTER(tblRisk[비중],tblRisk[구분]="종목"),FILTER(tblRisk[변동성],tblRisk[구분]="종목")),"-")',
             '="1 - 포트변동성/가중평균변동성"', nf=NF["pct1"])
        section(ws, "B11:R11", "종목별 위험 기여 (최근 N영업일 일간수익률, 현재 비중 적용)")
        lo = self.lo["tblRisk"]
        for c, (w, nf) in {"구분": (10, None), "종목코드": (10, None), "이름": (17, None), "섹터": (13, None), "비중": (9, NF["pct1"]),
                           "평가금액": (13, NF["krw"]), "변동성": (9, NF["pct1"]), "베타": (8, NF["ratio2"]),
                           "상관계수": (10, NF["ratio2"]), "한계위험": (10, NF["pct1"]), "위험기여": (10, NF["pct1"]),
                           "위험기여비중": (13, NF["pct1"]), "수익률N": (10, NF["pct_pl"]), "관측일수": (10, "0"),
                           "VaR95": (12, NF["krw"]), "과거VaR95": (13, NF["krw"]), "시장변동성": (12, NF["pct1"])}.items():
            set_col_format(lo, c, nf=nf, width=w)
        body = lo.DataBodyRange
        if body is not None:
            r0 = body.Row
            c_share = lo.ListColumns("위험기여비중").Range.Column
            c_w = lo.ListColumns("비중").Range.Column
            cf_databar(ws.Range(ws.Cells(r0, c_share), ws.Cells(r0 + 60, c_share)), color=C["up"])
            cf_databar(ws.Range(ws.Cells(r0, c_w), ws.Cells(r0 + 60, c_w)))
            c_type = lo.ListColumns("구분").Range.Column
            rng = ws.Range(ws.Cells(r0, lo.Range.Column), ws.Cells(r0 + 60, lo.Range.Column + lo.ListColumns.Count - 1))
            cf_expr(rng, f'=${col_letter(c_type)}{r0}="포트폴리오"', fill_color=C["accent_l"], bold=True)
        put(ws, "B40", "읽는 법: 위험기여비중 합계 = 100%. 비중보다 위험기여비중이 크게 높은 종목이 샤프지수를 깎는 주범입니다. "
                       "변동성을 낮추려면 위험기여 상위 종목 비중을 줄이거나 상관이 낮은 종목으로 분산하세요.", size=9, color=C["muted"])

        # 섹터 비중 (동적 배열) — T열 12행부터
        section(ws, "T11:AF11", "섹터 비중 (현금 포함)")
        put(ws, "T12", "섹터", bold=True, size=8, color=C["muted"])
        put(ws, "U12", "비중", bold=True, size=8, color=C["muted"])
        put(ws, "V12", "종목수", bold=True, size=8, color=C["muted"])
        put(ws, "T13", formula='=IFERROR(LET(s,UNIQUE(FILTER(tblHoldings[섹터],tblHoldings[보유수량]>0)),w,SUMIFS(tblHoldings[비중],tblHoldings[섹터],s),'
                               'n,COUNTIFS(tblHoldings[섹터],s,tblHoldings[보유수량],">0"),'
                               'VSTACK(SORTBY(HSTACK(s,w,n),w,-1),HSTACK("현금",현금잔고/순자산,""))),"")', size=9)
        set_nf(ws.Range("U13:U30"), NF["pct1"])
        fill(ws.Range("T12:V30"), "card")
        cf_databar(ws.Range("U13:U30"))
        ws.Columns("T").ColumnWidth = 14
        # 비중 vs 위험기여 차트 도우미 (숨김 AH:AJ)
        put(ws, "AH12", "도우미", size=8, color=C["muted"])
        put(ws, "AH13", formula='=IFERROR(FILTER(tblRisk[이름],tblRisk[구분]="종목"),"")')
        put(ws, "AI13", formula='=IFERROR(FILTER(tblRisk[비중],tblRisk[구분]="종목"),"")')
        put(ws, "AJ13", formula='=IFERROR(FILTER(tblRisk[위험기여비중],tblRisk[구분]="종목"),"")')
        ws.Columns("AH:AJ").Hidden = True
        shp, ch = chart(ws, "X12:AF24", XL_BAR_CLUSTERED, "비중 vs 위험기여")
        series(ch, "비중", f"='{self.wb.Name}'!차트_위험이름", f"='{self.wb.Name}'!차트_위험비중", fill_color="#9DB7E0")
        series(ch, "위험기여", f"='{self.wb.Name}'!차트_위험이름", f"='{self.wb.Name}'!차트_위험기여", fill_color=C["up"])
        style_axes(ch, y_nf="0%", category=False, legend=XL_LEGEND_TOP)
        ch.Axes(1).ReversePlotOrder = True
        shp, ch = chart(ws, "X25:AF38", XL_BAR_CLUSTERED, "섹터 비중")
        s = series(ch, "비중", f"='{self.wb.Name}'!차트_섹터명", f"='{self.wb.Name}'!차트_섹터비중", fill_color=C["accent"])
        style_axes(ch, y_nf="0%", category=False, legend=None)
        ch.Axes(1).ReversePlotOrder = True
        freeze(ws, 4)
        self.say("리스크 시트 완료")

    # -------------------------------------------------------------------------------------------------------
    def build_performance(self):
        ws = self.ws["성과"]
        sheet_setup(ws, bg=False, zoom=85, tab="good", default_width=9.3, widths={"A": 1.5, "B": 20, "C": 17, "D": 34})
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "PERFORMANCE", "대회 평가지표(수익률·샤프) 중심 성과 분석 — 일별 NAV 재구성 기반", "B2:Y3",
                  right_formula=self.HEADER_RIGHT, right_cell="Y2", right2_formula=self.HEADER_STATUS, right2_cell="Y3")
        fill(ws.Range("B4:Y4"), "navy2")
        self.nav_links(ws, 4)
        section(ws, "B6:D6", "성과 지표", "기간: 대회 시작일 ~ 기준일", "D6")
        metrics = [
            ("기준일", '=IFERROR(TAKE(tblNAV[일자],-1),"-")', NF["date"], "마지막 거래일(장중이면 현재가 반영)"),
            ("경과 거래일", F_DAYS, "0\"일\"", "남은 거래일: 설정의 종료일·휴장일 기준"),
            ("남은 거래일", F_REMAIN, "0\"일\"", ""),
            ("순자산 (NAV)", F_NAVLAST, NF["krw"], "초기자금 + 누적 현금흐름 + 보유평가"),
            ("누적 수익률", F_CUM, NF["pct_pl"], "★ 대회 평가지표"),
            ('="벤치마크 ("&벤치마크&")"', F_BMCUM, NF["pct_pl"], "설정에서 KOSPI/KOSDAQ/혼합 선택"),
            ("초과 수익률", F_EXCESS, NF["pct_pl"], "포트 누적 - 벤치마크 누적"),
            ("KOSPI / KOSDAQ", '=IFERROR(TEXT(TAKE(tblNAV[KOSPI누적],-1),"+0.00%;-0.00%")&" / "&TEXT(TAKE(tblNAV[KOSDAQ누적],-1),"+0.00%;-0.00%"),"-")', None, "같은 기간 지수 수익률"),
            ("연율화 수익률", F_CAGR, NF["pct_pl"], "(1+누적)^(252/거래일)-1 — 단기는 과장됨"),
            ("연율화 변동성", F_VOL, NF["pct1"], "일간수익률 표준편차×√252"),
            ("샤프지수 (rf 적용)", F_SHARPE, NF["ratio2"], '★ (평균일간수익률-rf/252)/표준편차×√252'),
            ("샤프지수 (rf=0)", F_SHARPE0, NF["ratio2"], "대회가 rf를 쓰지 않는 경우"),
            ("소르티노", F_SORTINO, NF["ratio2"], "하방 변동성만 사용"),
            ("최대낙폭 (MDD)", F_MDD, NF["pct"], "고점 대비 최대 하락"),
            ("현재 낙폭", F_CURDD, NF["pct"], ""),
            ("칼마 비율", F_CALMAR, NF["ratio2"], "연율화 수익률 / |MDD|"),
            ("베타 (vs BM)", F_BETA, NF["ratio2"], "일간수익률 회귀 기울기"),
            ("알파 (연, 젠센)", F_ALPHA, NF["pct_pl"], "(Rp-rf) - β(Rb-rf), 연율화"),
            ("상관계수 (vs BM)", F_CORR, NF["ratio2"], ""),
            ("추적오차 (연)", F_TE, NF["pct1"], "초과수익률 표준편차×√252"),
            ("정보비율", F_IR, NF["ratio2"], "연 초과수익 / 추적오차"),
            ("일간 승률", F_HIT, NF["pct1"], "수익률 > 0 인 날 비율"),
            ("BM 대비 승률", F_HITBM, NF["pct1"], "벤치마크를 이긴 날 비율"),
            ("최고 / 최저 일간", '=IFERROR(TEXT(' + F_BEST[1:] + ',"+0.00%;-0.00%")&" / "&TEXT(' + F_WORST[1:] + ',"+0.00%;-0.00%"),"-")', None, ""),
            ("평균 주식비중", F_AVGEXPO, NF["pct1"], ""),
            ("회전율", F_TURNOVER, "0.00\"회\"", "누적 매매금액 / 평균 NAV"),
            ("무위험수익률(연)", "=무위험수익률", NF["pct"], "설정값 없으면 CD91일물"),
        ]
        fill(ws.Range("B7:D%d" % (6 + len(metrics))), "card")
        r = 7
        for label, f, nf, note in metrics:
            if label.startswith("="):
                put(ws, f"B{r}", formula=label, size=9, color=C["text"], indent=1)
            else:
                put(ws, f"B{r}", label, size=9, color=C["text"], indent=1)
            put(ws, f"C{r}", formula=f, nf=nf, bold=True, size=10, h=XL_RIGHT)
            put(ws, f"D{r}", note, size=8, color=C["muted"], indent=1)
            bottom_line(ws.Range(f"B{r}:D{r}"), "line2")
            if note.startswith("★"):
                fill(ws.Range(f"B{r}:D{r}"), "accent_l")
            r += 1

        nav = self.lo["tblNAV"]
        x = nav.ListColumns("일자").DataBodyRange
        section(ws, "F6:Y6", "누적 수익률 · 벤치마크 · 초과수익")
        shp, ch = chart(ws, "F7:Y21", XL_LINE)
        series(ch, "포트폴리오", x, nav.ListColumns("누적수익률").DataBodyRange, color=C["accent"], weight=2.5)
        series(ch, "벤치마크", x, nav.ListColumns("BM누적").DataBodyRange, color="#8A94A3", weight=1.75)
        series(ch, "KOSPI", x, nav.ListColumns("KOSPI누적").DataBodyRange, color=C["up"], weight=1.0)
        series(ch, "KOSDAQ", x, nav.ListColumns("KOSDAQ누적").DataBodyRange, color="#E8890C", weight=1.0)
        s = series(ch, "누적초과(막대)", x, nav.ListColumns("누적초과").DataBodyRange, kind=XL_COLUMN_CLUSTERED, fill_color="#C9D8F2")
        style_axes(ch, y_nf="0.0%", x_nf="mm/dd", legend=XL_LEGEND_TOP)
        section(ws, "F22:O22", "일간 수익률")
        shp, ch = chart(ws, "F23:O35", XL_COLUMN_CLUSTERED)
        s = series(ch, "일간수익률", x, nav.ListColumns("일간수익률").DataBodyRange, fill_color=C["accent"])
        try:
            s.InvertIfNegative = True
            s.InvertColor = rgb(C["down"])
            s.Format.Fill.ForeColor.RGB = rgb(C["up"])
        except Exception:
            pass
        style_axes(ch, y_nf="0.0%", x_nf="mm/dd", legend=None)
        section(ws, "P22:Y22", "낙폭 · 20일 변동성")
        shp, ch = chart(ws, "P23:Y35", XL_AREA)
        series(ch, "낙폭", x, nav.ListColumns("낙폭").DataBodyRange, fill_color="#5B8FD9", transparency=0.35)
        series(ch, "20일 변동성(우)", x, nav.ListColumns("변동성20").DataBodyRange, kind=XL_LINE, color=C["warn"], weight=1.75,
               axis=XL_SECONDARY)
        style_axes(ch, y_nf="0.0%", x_nf="mm/dd", y2_nf="0%", legend=XL_LEGEND_TOP)

        # 손익 귀속 (종목별)
        section(ws, "B37:D37", "종목별 손익 기여 (실현+평가)", "초기자금 대비", "D37")
        for c, h in (("B", "종목"), ("C", "총손익"), ("D", "기여도 · 상태")):
            put(ws, f"{c}38", h, bold=True, size=8, color=C["muted"])
        bottom_line(ws.Range("B38:D38"))
        put(ws, "B39", formula='=IFERROR(TAKE(SORTBY(tblPositions[종목명],tblPositions[총손익],-1),16),"")', size=9)
        put(ws, "C39", formula='=IFERROR(TAKE(SORTBY(tblPositions[총손익],tblPositions[총손익],-1),16),"")', size=9)
        put(ws, "D39", formula='=IFERROR(TAKE(SORTBY(TEXT(tblPositions[기여도],"+0.00%;-0.00%")&"  "&tblPositions[상태],tblPositions[총손익],-1),16),"")', size=9)
        set_nf(ws.Range("C39:C54"), NF["krw_pl"])
        fill(ws.Range("B38:D54"), "card")
        section(ws, "F37:Y37", "손익 기여 차트")
        shp, ch = chart(ws, "F38:Y54", XL_BAR_CLUSTERED)
        s = series(ch, "총손익", f"='{self.wb.Name}'!차트_기여종목", f"='{self.wb.Name}'!차트_기여금액", fill_color=C["up"])
        try:
            s.InvertIfNegative = True
            s.InvertColor = rgb(C["down"])
        except Exception:
            pass
        style_axes(ch, y_nf="#,##0", category=False, legend=None)
        ch.Axes(1).ReversePlotOrder = True
        s.HasDataLabels = True
        set_nf(s.DataLabels(), "#,##0")
        s.DataLabels().Font.Size = 8

        section(ws, "B56:Y56", "일별 순자산(NAV) 원장", "T_NAV 쿼리: 매매일지 + 수정주가 종가로 매일 재구성", "Y56")
        lo = nav
        for c, (w, nf) in {"구분": (6, None), "경과일": (6, "0"), "일자": (10, NF["date"]), "순자산": (13, NF["krw"]),
                           "현금": (12, NF["krw"]), "주식평가": (12, NF["krw"]), "주식비중": (8, NF["pct1"]),
                           "일간손익": (11, NF["krw_pl"]), "일간수익률": (9, NF["pct_pl"]), "누적수익률": (9, NF["pct_pl"]),
                           "KOSPI": (9, NF["num2"]), "KOSPI일간": (9, NF["pct_pl"]), "KOSPI누적": (9, NF["pct_pl"]),
                           "KOSDAQ": (9, NF["num2"]), "KOSDAQ일간": (9, NF["pct_pl"]), "KOSDAQ누적": (9, NF["pct_pl"]),
                           "BM일간": (8, NF["pct_pl"]), "BM누적": (8, NF["pct_pl"]), "초과일간": (8, NF["pct_pl"]),
                           "누적초과": (8, NF["pct_pl"]), "고점": (12, NF["krw"]), "낙폭": (8, NF["pct"]),
                           "변동성20": (8, NF["pct1"]), "누적샤프": (8, NF["ratio2"]), "매매금액": (12, NF["krw"]),
                           "보유종목수": (7, "0")}.items():
            set_col_format(lo, c, nf=nf)
        freeze(ws, 4)
        self.say("성과 시트 완료")

    # -------------------------------------------------------------------------------------------------------
    def build_quote_board(self):
        ws = self.ws["시세판"]
        sheet_setup(ws, bg=False, zoom=85, tab="accent", widths={"A": 1.5})
        ws.Rows(1).RowHeight = 6
        lo = self.lo["tblQuote"]
        # 표 오른쪽 끝(새 열로 넓어짐) — 제목 막대는 표 끝 + 2열까지(최소 기존 BD), 차트 도우미는 그 오른쪽
        tbl_last = lo.Range.Column + lo.ListColumns.Count - 1
        bar_end = col_letter(max(tbl_last + 2, 56))
        title_bar(ws, "QUOTE BOARD", "보유·관심 종목 시세 · 밸류에이션 · 52주 · 수급(추정가집계 포함) · 신용/공매도 · 다음 이벤트 · 기술적 신호 | "
                  "관심종목은 [설정] 시트에서 추가", f"B2:{bar_end}3", right_formula=self.HEADER_RIGHT, right_cell="R2",
                  right2_formula=self.HEADER_STATUS, right2_cell="R3")
        fill(ws.Range(f"B4:{bar_end}4"), "navy2")
        self.nav_links(ws, 4)
        widths = {"구분": 8, "종목코드": 9, "종목명": 15, "시장": 9, "섹터": 16, "대테마": 12, "현재가": 11, "전일대비": 9,
                  "등락률": 9, "시가": 10, "고가": 10, "저가": 10, "전일종가": 10, "거래량": 12, "거래대금억": 10,
                  "전일대비거래량": 9, "시가총액억": 12, "PER": 7, "PBR": 6, "EPS": 9, "BPS": 9, "외국인소진율": 8,
                  "고52주": 10, "고52주일": 10, "고52주대비": 9, "저52주": 10, "저52주일": 10, "저52주대비": 9,
                  "외국인당일주": 11, "프로그램당일주": 11, "외국인5일억": 10, "기관5일억": 10, "개인5일억": 10,
                  "외국인20일억": 10, "기관20일억": 10, "상한가": 10, "하한가": 10, "유의": 10, "DB유의": 10,
                  "상태": 6, "조회시각": 15, "거래량비율": 8, "MA20": 10, "MA60": 10, "추세": 6, "RSI14": 7,
                  "수익률20일": 9, "수익률60일": 9, "변동성20": 8, "교차": 10, "관심가": 10, "목표가": 10,
                  "투자포인트": 24, "신호": 34, "다음이벤트": 18, "신용잔고율": 8, "공매도비중5일": 9, "추정시점": 7,
                  "추정외국인주": 11, "추정기관주": 11, "추정합산억": 10}
        for c, w in widths.items():
            set_col_format(lo, c, width=w)
        nfs = {"현재가": NF["krw"], "전일대비": "+#,##0;-#,##0;0", "등락률": NF["pct_arrow"], "시가": NF["krw"], "고가": NF["krw"],
               "저가": NF["krw"], "전일종가": NF["krw"], "거래량": NF["num0"], "거래대금억": NF["num0"], "전일대비거래량": NF["pct1"],
               "시가총액억": NF["num0"], "PER": NF["num1"], "PBR": NF["ratio2"], "EPS": NF["num0"], "BPS": NF["num0"],
               "외국인소진율": NF["pct1"], "고52주": NF["krw"], "고52주일": NF["date"], "고52주대비": NF["pct1"],
               "저52주": NF["krw"], "저52주일": NF["date"], "저52주대비": NF["pct1"], "외국인당일주": "+#,##0;-#,##0;0",
               "프로그램당일주": "+#,##0;-#,##0;0", "외국인5일억": NF["eok_pl"], "기관5일억": NF["eok_pl"], "개인5일억": NF["eok_pl"],
               "외국인20일억": NF["eok_pl"], "기관20일억": NF["eok_pl"], "상한가": NF["krw"], "하한가": NF["krw"], "조회시각": NF["dt"],
               "신용잔고율": NF["pct"], "공매도비중5일": NF["pct1"], "추정외국인주": "+#,##0;-#,##0;0", "추정기관주": "+#,##0;-#,##0;0",
               "추정합산억": '[Color10]+#,##0.0"억";[Color11]-#,##0.0"억";0.0"억"'}
        for c, nf in nfs.items():
            set_col_format(lo, c, nf=nf)
        set_col_format(lo, "종목명", bold=True)
        set_col_format(lo, "추정시점", h=XL_CENTER)
        style(lo.ListColumns("신호").DataBodyRange, color=C["warn"], size=9)
        style(lo.ListColumns("다음이벤트").DataBodyRange, color=C["navy"], size=9)
        body = lo.DataBodyRange
        if body is not None:
            r0 = body.Row
            c0 = lo.Range.Column
            last = c0 + lo.ListColumns.Count - 1
            gcol = col_letter(lo.ListColumns("구분").Range.Column)
            cf_expr(ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + 150, last)), f'=LEFT(${gcol}{r0},2)="보유"', fill_color="#F3F8FF")
            tcol = col_letter(lo.ListColumns("추세").Range.Column)
            tr = ws.Range(f"{tcol}{r0}:{tcol}{r0 + 150}")
            cf_expr(tr, f'={tcol}{r0}="상승"', font=C["up"], bold=True)
            cf_expr(tr, f'={tcol}{r0}="하락"', font=C["down"], bold=True)
            rcol = col_letter(lo.ListColumns("RSI14").Range.Column)
            rr = ws.Range(f"{rcol}{r0}:{rcol}{r0 + 150}")
            cf_expr(rr, f'=AND(ISNUMBER({rcol}{r0}),{rcol}{r0}>=70)', font=C["up"], bold=True)
            cf_expr(rr, f'=AND(ISNUMBER({rcol}{r0}),{rcol}{r0}<=30)', font=C["down"], bold=True)

        # 종목 차트 섹션 (행 6~23)
        section(ws, "B6:F6", "종목 차트")
        put(ws, "B7", "종목코드 ▶", bold=True, size=10, color=C["navy"])
        sel = ws.Range("C7")
        set_nf(sel, "@")
        first_code = "005930"
        sel.Value = first_code
        style(sel, bold=True, size=12, fill=C["input"], h=XL_CENTER)
        border(sel, "warn")
        validation_list(sel, "=INDIRECT(\"tblQuote[종목코드]\")", show_error=False)
        put(ws, "D7", formula='=IFERROR(XLOOKUP(C7,tblQuote[종목코드],tblQuote[종목명]),"시세판에 없는 코드")', bold=True, size=12)
        info = [
            ("현재가", 'XLOOKUP($C$7,tblQuote[종목코드],tblQuote[현재가])', NF["krw"]),
            ("등락률", 'XLOOKUP($C$7,tblQuote[종목코드],tblQuote[등락률])', NF["pct_arrow"]),
            ("시가총액(억)", 'XLOOKUP($C$7,tblQuote[종목코드],tblQuote[시가총액억])', NF["num0"]),
            ("PER / PBR", 'TEXT(XLOOKUP($C$7,tblQuote[종목코드],tblQuote[PER]),"0.0")&" / "&TEXT(XLOOKUP($C$7,tblQuote[종목코드],tblQuote[PBR]),"0.00")', None),
            ("52주 고가 대비", 'XLOOKUP($C$7,tblQuote[종목코드],tblQuote[고52주대비])', NF["pct1"]),
            ("20 / 60일 수익률", 'TEXT(XLOOKUP($C$7,tblQuote[종목코드],tblQuote[수익률20일]),"+0.0%;-0.0%")&" / "&TEXT(XLOOKUP($C$7,tblQuote[종목코드],tblQuote[수익률60일]),"+0.0%;-0.0%")', None),
            ("추세 · RSI14", 'XLOOKUP($C$7,tblQuote[종목코드],tblQuote[추세])&" · "&TEXT(XLOOKUP($C$7,tblQuote[종목코드],tblQuote[RSI14]),"0")', None),
            ("20일 변동성", 'XLOOKUP($C$7,tblQuote[종목코드],tblQuote[변동성20])', NF["pct1"]),
            ("외국인 / 기관 5일", 'TEXT(XLOOKUP($C$7,tblQuote[종목코드],tblQuote[외국인5일억]),"+#,##0;-#,##0")&" / "&TEXT(XLOOKUP($C$7,tblQuote[종목코드],tblQuote[기관5일억]),"+#,##0;-#,##0")&" 억"', None),
            ("외국인 소진율", 'XLOOKUP($C$7,tblQuote[종목코드],tblQuote[외국인소진율])', NF["pct1"]),
            ("섹터 · 시장", 'XLOOKUP($C$7,tblQuote[종목코드],tblQuote[섹터])&" · "&XLOOKUP($C$7,tblQuote[종목코드],tblQuote[시장])', None),
            ("신호", 'XLOOKUP($C$7,tblQuote[종목코드],tblQuote[신호])', None),
        ]
        r = 9
        for label, f, nf in info:
            put(ws, f"B{r}", label, size=9, color=C["muted"], indent=1)
            put(ws, f"D{r}", formula=f'=IFERROR({f},"-")', nf=nf, bold=True, size=10, h=XL_LEFT)
            bottom_line(ws.Range(f"B{r}:F{r}"), "line2")
            r += 1
        fill(ws.Range("B7:F21"), "card")
        style(ws.Range("D20"), color=C["warn"], size=9)
        # 차트 도우미 (숨김 열) — 표 오른쪽 끝에서 6열 뒤(새 열로 표가 넓어져도 겹치거나 숨겨지지 않게; 기존 표 폭이면 BH)
        base = max(60, tbl_last + 6)
        hc = [col_letter(base + i) for i in range(5)]
        put(ws, f"{hc[0]}6", "차트 도우미", size=8, color=C["muted"])
        put(ws, f"{hc[0]}7", formula='=IFERROR(FILTER(tblPriceHist[일자],tblPriceHist[종목코드]=$C$7),NA())', nf=NF["mmdd"])
        put(ws, f"{hc[1]}7", formula='=IFERROR(FILTER(tblPriceHist[종가],tblPriceHist[종목코드]=$C$7),NA())')
        put(ws, f"{hc[2]}7", formula='=IFERROR(FILTER(IF(ISNUMBER(tblPriceHist[MA20]),tblPriceHist[MA20],NA()),tblPriceHist[종목코드]=$C$7),NA())')
        put(ws, f"{hc[3]}7", formula='=IFERROR(FILTER(IF(ISNUMBER(tblPriceHist[MA60]),tblPriceHist[MA60],NA()),tblPriceHist[종목코드]=$C$7),NA())')
        put(ws, f"{hc[4]}7", formula='=IFERROR(FILTER(tblPriceHist[거래량],tblPriceHist[종목코드]=$C$7),NA())')
        set_nf(ws.Range(f"{hc[0]}7:{hc[0]}300"), NF["mmdd"])
        ws.Range(f"{hc[0]}:{hc[4]}").EntireColumn.Hidden = True
        add_names(self.wb, {
            "차트_일자": f"=OFFSET(시세판!${hc[0]}$7,0,0,MAX(1,COUNT(시세판!${hc[0]}$7:${hc[0]}$400)),1)",
            "차트_종가": f"=OFFSET(시세판!${hc[1]}$7,0,0,MAX(1,COUNT(시세판!${hc[0]}$7:${hc[0]}$400)),1)",
            "차트_MA20": f"=OFFSET(시세판!${hc[2]}$7,0,0,MAX(1,COUNT(시세판!${hc[0]}$7:${hc[0]}$400)),1)",
            "차트_MA60": f"=OFFSET(시세판!${hc[3]}$7,0,0,MAX(1,COUNT(시세판!${hc[0]}$7:${hc[0]}$400)),1)",
            "차트_거래량": f"=OFFSET(시세판!${hc[4]}$7,0,0,MAX(1,COUNT(시세판!${hc[0]}$7:${hc[0]}$400)),1)",
        })
        section(ws, "H6:R6", "가격 · 이동평균 · 거래량 (최근 약 100영업일)", "코드를 바꾸면 즉시 반영", "R6")
        shp, ch = chart(ws, "H7:R23", XL_LINE)
        n = self.wb.Name
        series(ch, "종가", f"='{n}'!차트_일자", f"='{n}'!차트_종가", color=C["navy"], weight=2)
        series(ch, "MA20", f"='{n}'!차트_일자", f"='{n}'!차트_MA20", color=C["up"], weight=1.25)
        series(ch, "MA60", f"='{n}'!차트_일자", f"='{n}'!차트_MA60", color=C["down"], weight=1.25)
        # 거래량은 마지막에 보조축 막대로 (첫 계열을 보조축에 두면 모두 주축으로 합쳐짐)
        series(ch, "거래량(우)", f"='{n}'!차트_일자", f"='{n}'!차트_거래량", kind=XL_COLUMN_CLUSTERED, fill_color="#D5DCE6",
               axis=XL_SECONDARY)
        style_axes(ch, y_nf="#,##0", x_nf="mm/dd", y2_nf="#,##0,,\"M\"", legend=XL_LEGEND_TOP)
        section(ws, "B25:R25", "시세판 (보유 → 관심 순)", "필터/정렬 가능 · 신호: 52주고가근접·거래량급증·골든크로스·RSI·관심가·수급 · "
                "추정가집계는 개장일 장중(09:00~15:30)만 · 신용/공매도·이벤트는 [전체] 갱신값", "R25")
        freeze(ws, 27, 4)
        self.say("시세판 시트 완료")

    # -------------------------------------------------------------------------------------------------------
    def build_trade_sheets(self):
        ws = self.ws["매매일지"]
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "TRADE JOURNAL", "매매 기록 = 포트폴리오·NAV·성과의 원천 데이터 | 매수/매도할 때마다 한 줄씩 입력", "B2:P3")
        fill(ws.Range("B4:P4"), "navy2")
        self.nav_links(ws, 4)
        notes = [
            "① 일자·종목코드(6자리)·구분(매수/매도)·수량·단가만 입력하면 됩니다. 종목명·금액·확인은 자동입니다.",
            "② 수수료/세금을 비워두면 [설정]의 수수료율·거래세율로 계산(원 미만 절사). 대회 체결내역의 실제 금액을 알면 직접 입력하세요.",
            "③ 목표가/손절가를 적으면 [포트폴리오]·[대시보드] 알림에 쓰입니다(비우면 설정의 기본 %). 매매근거는 인터뷰 때 강력한 자료가 됩니다.",
            "④ 같은 날 여러 건은 실제 체결 순서대로 위→아래로 입력. 표 바로 아래 행에 입력하면 표가 자동으로 확장됩니다.",
            "⑤ 입력 후 [데이터] > [모두 새로 고침] (Ctrl+Alt+F5) — 새 종목의 시세·NAV·위험이 다시 계산됩니다.",
        ]
        for i, t in enumerate(notes):
            put(ws, f"B{6 + i}", t, size=9, color=C["text"])
        fill(ws.Range("B6:P10"), "card")
        put(ws, "P6", "※ [샘플] 행은 예시입니다 — 대회 전 삭제", size=9, color=C["up"], bold=True, h=XL_RIGHT)
        lo = self.lo["tblTrades"]
        body = lo.DataBodyRange
        r0 = body.Row
        chk_col = col_letter(lo.ListColumns("확인").Range.Column)
        cf_expr(ws.Range(f"{chk_col}{r0}:{chk_col}{r0 + 500}"), f'=LEFT({chk_col}{r0},1)="⚠"', font=C["up"], bold=True)
        side_col = col_letter(lo.ListColumns("구분").Range.Column)
        cf_expr(ws.Range(f"{side_col}{r0}:{side_col}{r0 + 500}"), f'={side_col}{r0}="매수"', font=C["up"], bold=True)
        cf_expr(ws.Range(f"{side_col}{r0}:{side_col}{r0 + 500}"), f'={side_col}{r0}="매도"', font=C["down"], bold=True)
        freeze(ws, 11, 3)

        ws = self.ws["매매분석"]
        sheet_setup(ws, bg=False, zoom=85, tab="#E8890C", widths={"A": 1.5})
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "TRADE ANALYTICS", "거래별 실현손익(이동평균 원가) · 승률 · 손익비 · 보유기간 · 비용", "B2:Z3",
                  right_formula=self.HEADER_RIGHT, right_cell="Z2")
        fill(ws.Range("B4:Z4"), "navy2")
        self.nav_links(ws, 4)
        sells = 'FILTER(tblTradeLog[실현수익률],tblTradeLog[구분]="매도")'
        spans = ["B6:D9", "E6:G9", "H6:J9", "K6:M9", "N6:P9", "Q6:S9", "T6:V9", "W6:Y9"]
        card(ws, spans[0], "거래 수", '=COUNTA(tblTradeLog[순번])', '="매수 "&COUNTIF(tblTradeLog[구분],"매수")&" · 매도 "&COUNTIF(tblTradeLog[구분],"매도")', nf="0\"건\"")
        card(ws, spans[1], "승률 (매도 기준)", '=IFERROR(COUNTIF(tblTradeLog[결과],"승")/COUNTIF(tblTradeLog[구분],"매도"),"-")',
             '="승 "&COUNTIF(tblTradeLog[결과],"승")&" / 패 "&COUNTIF(tblTradeLog[결과],"패")', nf=NF["pct1"])
        card(ws, spans[2], "평균 수익 (승)", f'=IFERROR(AVERAGE(FILTER({sells},{sells}>0)),"-")', '="이익 매도의 평균 수익률"', nf=NF["pct_pl"])
        card(ws, spans[3], "평균 손실 (패)", f'=IFERROR(AVERAGE(FILTER({sells},{sells}<0)),"-")', '="손실 매도의 평균 수익률"', nf=NF["pct_pl"])
        card(ws, spans[4], "손익비", f'=IFERROR(AVERAGE(FILTER({sells},{sells}>0))/ABS(AVERAGE(FILTER({sells},{sells}<0))),"-")',
             '="평균수익/평균손실 (>1.5 권장)"', nf=NF["ratio2"])
        card(ws, spans[5], "실현손익 합계", '=SUM(tblTradeLog[실현손익])', '="수수료·세금 차감 후"', nf=NF["krw_pl"])
        card(ws, spans[6], "거래비용", '=SUM(tblTradeLog[수수료])+SUM(tblTradeLog[세금])',
             '="수수료 "&TEXT(SUM(tblTradeLog[수수료]),"#,##0")&" · 세금 "&TEXT(SUM(tblTradeLog[세금]),"#,##0")', nf=NF["krw"])
        card(ws, spans[7], "평균 보유기간", '=IFERROR(AVERAGE(FILTER(tblTradeLog[보유일수],tblTradeLog[구분]="매도")),"-")',
             '="매도 거래 기준 (달력일)"', nf="0.0\"일\"")
        section(ws, "B11:Z11", "매매 원장 (최신순)", "평균단가는 매수수수료 포함 원가 → 실현+평가손익 = 순자산-초기자금", "Z11")
        lo = self.lo["tblTradeLog"]
        for c, (w, nf) in {"순번": (5, "0"), "일자": (10, NF["date"]), "종목코드": (8, None), "종목명": (14, None), "섹터": (11, None),
                           "구분": (6, None), "수량": (7, NF["num0"]), "단가": (10, NF["krw"]), "금액": (12, NF["krw"]),
                           "수수료": (8, NF["krw"]), "세금": (8, NF["krw"]), "현금흐름": (12, NF["krw_pl"]), "보유수량": (8, NF["num0"]),
                           "평균단가": (11, NF["krw"]), "실현손익": (11, NF["krw_pl"]), "실현수익률": (9, NF["pct_pl"]),
                           "종목누적실현": (12, NF["krw_pl"]), "보유일수": (7, "0"), "결과": (5, None), "전략": (7, None),
                           "매매근거": (28, None), "목표가": (10, NF["krw"]), "손절가": (10, NF["krw"]), "메모": (16, None),
                           "확인": (14, None)}.items():
            set_col_format(lo, c, nf=nf, width=w)
        body = lo.DataBodyRange
        if body is not None:
            r0 = body.Row
            rc = col_letter(lo.ListColumns("결과").Range.Column)
            rr = ws.Range(f"{rc}{r0}:{rc}{r0 + 500}")
            cf_expr(rr, f'={rc}{r0}="승"', font=C["up"], bold=True)
            cf_expr(rr, f'={rc}{r0}="패"', font=C["down"], bold=True)
            sc = col_letter(lo.ListColumns("구분").Range.Column)
            sr = ws.Range(f"{sc}{r0}:{sc}{r0 + 500}")
            cf_expr(sr, f'={sc}{r0}="매수"', font=C["up"])
            cf_expr(sr, f'={sc}{r0}="매도"', font=C["down"])
        freeze(ws, 12, 5)
        self.say("매매일지·매매분석 시트 완료")

    # -------------------------------------------------------------------------------------------------------
    def build_universe_sheet(self):
        ws = self.ws["종목DB"]
        sheet_setup(ws, bg=False, zoom=85, tab="#6B7785", widths={"A": 1.5})
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "UNIVERSE", "KOSPI·KOSDAQ 개별 종목 마스터 (KIS 종목정보 파일, ETF·ETN·리츠·펀드 제외) | 대회편입 Y = 대회 종목: "
                  "2026-09-30 기준 시총 1,000억·5일 평균 거래대금 25억 이상 보통주(고정 명단) ± [설정] 수정표 | 섹터 = [설정] 섹터 기준",
                  "B2:Y3", right_formula='="종목 수 "&TEXT(COUNTA(tblUniverse[종목코드]),"#,##0")&" · 대회 종목 "&'
                                         'TEXT(COUNTIF(tblUniverse[대회편입],"Y"),"#,##0")', right_cell="Y2")
        fill(ws.Range("B4:Y4"), "navy2")
        self.nav_links(ws, 4)
        section(ws, "B6:Y6", "종목 검색", "이름 일부 또는 코드 입력 → 최대 10건", "Y6")
        put(ws, "B7", "검색어 ▶", bold=True, color=C["navy"])
        s = ws.Range("C7")
        s.Value = "삼성"
        style(s, bold=True, size=11, fill=C["input"])
        border(s, "warn")
        heads = ["종목코드", "종목명", "시장", "섹터", "규모", "지수편입", "유의사항", "대회편입", "기준가", "시가총액억",
                 "매출액억", "영업이익억", "영업이익률", "ROE"]
        for i, h in enumerate(heads):
            put(ws, ws.Cells(8, 2 + i).Address, h, bold=True, size=8, color=C["muted"])
        bottom_line(ws.Range(ws.Cells(8, 2), ws.Cells(8, 1 + len(heads))))
        cond = '(ISNUMBER(SEARCH($C$7,tblUniverse[종목명]))+ISNUMBER(SEARCH($C$7,tblUniverse[종목코드])))>0'
        cols = ",".join(f'IF(tblUniverse[{h}]="","",tblUniverse[{h}])' for h in heads)
        put(ws, "B9", formula=f'=IF($C$7="","",IFERROR(TAKE(FILTER(HSTACK({cols}),{cond}),10),"검색 결과 없음"))', size=9)
        set_nf(ws.Range("J9:M18"), NF["num0"])
        set_nf(ws.Range("N9:O18"), NF["pct1"])
        fill(ws.Range("B8:O18"), "card")
        section(ws, "B20:Y20", "전체 종목 (시가총액순) — 필터 버튼으로 시장·섹터·대회편입·대테마·유의사항 걸러보기 "
                               "| 대분류·NICS 업종·NICS 세부 = NICS(VALUESearch, 없으면 KRX업종) · 분류출처 · 대회비고")
        lo = self.lo["tblUniverse"]
        for c, (w, nf) in {"종목코드": (10, None), "종목명": (17, None), "시장": (9, None), "섹터": (16, None), "규모": (7, None),
                           "지수편입": (11, None), "유의사항": (16, None), "대회편입": (9, None), "기준가": (11, NF["krw"]),
                           "시가총액억": (12, NF["num0"]), "매출액억": (11, NF["num0"]), "영업이익억": (11, NF["num0"]),
                           "영업이익률": (11, NF["pct1"]), "ROE": (8, NF["pct1"]), "재무기준": (10, None), "상장일": (11, NF["date"]),
                           "상장주식수": (14, NF["num0"]), "대분류": (11, None), "NICS 업종": (16, None), "NICS 세부": (18, None),
                           "대테마": (12, None), "세부테마": (16, None), "분류출처": (9, None), "대회비고": (16, None),
                           "KRX업종": (12, None), "우선주": (8, None), "SPAC": (8, None),
                           "거래정지": (9, None), "관리종목": (9, None), "시장경고": (9, None), "종류": (7, None)}.items():
            set_col_format(lo, c, nf=nf, width=w)
        freeze(ws, 21, 3)
        self.say("종목DB 시트 완료")

    # -------------------------------------------------------------------------------------------------------
    def build_settings_sheet(self):
        ws = self.ws["설정"]
        ws.Rows(1).RowHeight = 6
        # 제목 막대는 ⑤ 수정표 열(W:AD)까지
        bar_end = col_letter(ws.Range(OVERRIDE_ANCHOR).Column + len(OVERRIDE_HEADERS) - 1)
        title_bar(ws, "SETTINGS", "노란 칸만 수정하세요 · 표 아래 행에 입력하면 표가 자동 확장됩니다 · ⑤ 수정표는 오른쪽 끝", f"B2:{bar_end}3")
        fill(ws.Range(f"B4:{bar_end}4"), "navy2")
        self.nav_links(ws, 4)
        put(ws, "B5", "① 기본 설정", bold=True, size=11, color=C["navy"])
        put(ws, "G5", "② 관심종목 (시세판·알림 대상)", bold=True, size=11, color=C["navy"])
        put(ws, "N5", "③ 해외·환율·금리 (N=지수, X=환율, I=금리)", bold=True, size=11, color=C["navy"])
        put(ws, "T5", "④ 휴장일 (D-day 계산)", bold=True, size=11, color=C["navy"])
        # 새 설정 키 입력 도우미(목록) — 키 이름으로 행을 찾음(행 위치가 바뀌어도 맞게)
        lo = self.lo["tblSettings"]
        keys = [str(k).strip() if k is not None else "" for (k,) in lo.ListColumns("키").DataBodyRange.Value]
        for key, choices in (("sector_basis", "대분류,업종,세부,대테마"), ("force_weekly", "Y,N")):
            if key in keys:
                validation_list(lo.ListColumns("값").DataBodyRange.Cells(keys.index(key) + 1, 1), choices)
        # 기본 설정 표가 새 키로 B6:E35까지 길어졌으므로 기존 안내 문구는 표 아래로(B28·B29에서 이동).
        # 이관한 사용자 키가 더 있으면 표가 더 길어지므로 표 끝 + 빈 행 1개 아래(최소 SETTINGS_NOTE_ROW)에 둔다 — 표 바로 아래
        # 행에 글자를 쓰면 Excel이 표를 자동으로 늘려 안내 문구가 설정 행이 된다(2026-10-01 이관 왕복 시험 실측: 30키 → 33키)
        n = max(SETTINGS_NOTE_ROW, lo.Range.Row + lo.Range.Rows.Count + 1)
        put(ws, f"B{n}", "※ 앱키·시크릿은 이 통합문서에 저장되지 않습니다. Power Query가 위 경로의 kis_devlp.yaml(저장소 샘플코드와 같은 파일)을 직접 읽습니다.",
            size=9, color=C["muted"])
        put(ws, f"B{n + 1}", "※ 휴장일 목록은 참고용입니다. 한국거래소(KRX) 휴장일 공지로 확인 후 필요하면 수정하세요.", size=9, color=C["muted"])
        ov_col = "".join(ch for ch in OVERRIDE_TITLE_CELL if ch.isalpha())
        hyperlink(ws, f"B{n + 2}", f"'설정'!{OVERRIDE_TITLE_CELL}", f"› ⑤ 수정표(대회편입·분류 고치기)는 오른쪽 {ov_col}열에 있습니다")
        style(ws.Range(f"B{n + 2}"), color=C["accent"], size=9)
        # ⑤ 수정표 — 표는 build_inputs가 OVERRIDE_ANCHOR에 만듦. 무시되는 행이 있으면 제목 줄 오른쪽에 경고
        t = ws.Range(OVERRIDE_TITLE_CELL)
        put(ws, t.Address, "⑤ 수정표 (대회편입·분류 — 항상 우선)", bold=True, size=11, color=C["navy"])
        put(ws, OVERRIDE_WARN_CELL, formula=F_OVERRIDE_WARN, bold=True, size=10, color=C["up"])
        put(ws, ws.Cells(t.Row + 1, t.Column).Address,
            "종목코드(6자리)와 바꿀 칸만 입력 · 빈칸 = 수정 없음 · 대회편입: 추가(대회 종목으로) / 제외(대회 종목에서 뺌) · "
            "같은 종목을 여러 행에 쓰면 칸마다 아래쪽 값 우선", size=9, color=C["text"])
        put(ws, ws.Cells(t.Row + 2, t.Column).Address,
            "잘못된 종목코드·대회편입 값, 종목DB에 없는 코드는 무시(위에 경고) · 반영: [모두 새로 고침] → 종목DB·순위★ "
            "(포트폴리오·시세판의 섹터·대테마는 한 번 더) · 대회종목 페이지는 [시세]/[전체]", size=9, color=C["muted"])
        freeze(ws, 6)

    # -------------------------------------------------------------------------------------------------------
    def build_guide(self):
        ws = self.ws["가이드"]
        sheet_setup(ws, zoom=90, tab="#6B7785", widths={"A": 2, "B": 4, "C": 110})
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "GUIDE", "처음 설정 · 버튼과 페이지 · 매일 루틴 · 수정표 · 지표 정의 · 데이터 한계 · 문제 해결", "B2:C3")
        fill(ws.Range("B4:C4"), "navy2")
        self.nav_links(ws, 4)
        lines = [
            ("h", "1. 처음 한 번만"),
            ("t", "① 저장소 README 3.5절대로 ~/KIS/config/kis_devlp.yaml 에 실전투자 앱키(my_app)·시크릿(my_sec)을 입력합니다. (시세·순위 API는 실전 앱키 필요)"),
            ("t", "② [설정] 시트에서 대회 시작일·종료일·초기자금·수수료율·거래세율·벤치마크를 대회 규정에 맞게 수정합니다. "
                  "VALUESearch 파일 경로(vs_path)와 섹터 기준(sector_basis)도 확인합니다."),
            ("t", "③ [매매일지]의 [샘플] 행과 [설정]의 샘플 관심종목을 지우고 내 종목으로 바꿉니다."),
            ("t", "④ 파일을 열 때 노란 '보안 경고' 줄이 뜨면 [콘텐츠 사용]을 누릅니다 — 매크로(버튼)와 데이터 연결을 허용하는 단계라 "
                  "누르지 않으면 버튼이 동작하지 않습니다. '웹 콘텐츠 액세스' 창이 뜨면 [익명] → [연결]."),
            ("t", "⑤ [데이터] > [모두 새로 고침] (Ctrl+Alt+F5, 약 1분 15초) 뒤 [대회종목] 시트의 [전체]를 한 번 눌러 가격 이력·수급·목표주가 등을 처음 채웁니다"
                  "(처음 한 번은 오래 걸립니다). 첫 실행이나 토큰 재발급 직후 오류가 보이면 한 번 더 새로 고침합니다."),
            ("h", "2. 버튼과 페이지"),
            ("t", "[모두 새로 고침] (Ctrl+Alt+F5) — 가벼운 데이터: 지수·시장 수급·업종 등락·해외지표·주도주 순위·시세판(외인·기관 추정가집계 포함)·"
                  "보유/NAV/위험·뉴스·증시 자금. 버튼 전용 데이터는 건드리지 않습니다. (약 72~84초)"),
            ("t", "[시세] ([대회종목] 시트) — 대회 종목·보유·관심의 현재가를 받아 가격 이력의 마지막 세션 봉을 갱신하고 기간 등락·테마 집계를 다시 "
                  "계산합니다. 장중 수시로 누릅니다. (약 30~37초, 통합문서를 연 뒤 첫 실행은 약 45초)"),
            ("t", "[전체] ([대회종목] 시트) — [시세]의 모든 것 + 분류(VALUESearch) · 가격 이력 보정 · 종목 수급 · 목표주가 · KIS 추정과 스냅샷 · "
                  "이벤트 · 주간 항목(분기 실적·신용/공매도/대차, 갱신 시기가 됐을 때) · KRX 업종. 하루 한 번 장 마감 후. "
                  "(평일 약 4~5분 — 통합문서를 연 뒤 첫 실행은 준비·대기로 약 1.5~3분 추가, 주간 항목이 도는 날은 10분 이상, 처음 한 번은 약 15분)"),
            ("t", "[업종] ([업종] 시트) — KRX 업종지수·업종별 투자자 수급과 테마 집계만 갱신합니다. (약 20초)"),
            ("t", "[조회] ([종목분석] 시트) — 종목코드 칸의 한 종목만 자세히(투자자 120세션·체결금액별 매매비중·매물대·추정가집계·신용/공매도·"
                  "목표주가·추정실적·분기 실적·뉴스·이벤트) 받아 옵니다. 대회 종목이 아니어도 됩니다. (약 1분)"),
            ("t", "버튼을 누르면 상태 표시줄에 진행 단계가 보이고, 끝나면 단계별 요약 창이 뜹니다. 버튼 옆 '최근 조회'에 끝난 시각, '상태'에 "
                  "정상 / 일부 오류 n건 / 실패(이전 데이터 표시 중)가 남습니다. 실행 중 Esc를 누르면 진행 중인 단계 뒤에서 멈춥니다."),
            ("t", "새 페이지: [대회종목] 대회 종목 전체(Q.Pack형 3색·최근 20일 줄무늬, +/− 열 묶음) · [업종] KRX 업종·테마 집계 · "
                  "[종목분석] 한 종목 상세 · [뉴스·이벤트] 이벤트 캘린더와 보유·관심 뉴스 100건."),
            ("t", "기존 화면: [시장]·[대시보드] 주도주는 시장 전체 개별종목이며 ★ = 대회 종목(회색 = 대회 밖), [시장]에 증시 자금 동향 추가, "
                  "[시세판]에 대테마·추정가집계·신용잔고율·공매도 비중·다음 이벤트 열 추가, [종목DB]에 NICS·테마·대회비고 열 추가."),
            ("h", "3. 매일 루틴 (실제 운용역처럼)"),
            ("t", "08:30 장 전 — [모두 새로 고침] → [시장] 해외지수·환율·금리·증시 자금 확인 → [대시보드] 알림(손절/목표/비중) 확인 → 오늘 매매 계획 메모. "
                  "(장 시작 전 [시세]는 새 봉을 만들지 않고 마지막 세션 봉만 갱신합니다)"),
            ("t", "장중 — 필요할 때 [모두 새로 고침](시세판·순위·추정가집계) + [시세](대회종목 페이지). [시세판] 신호·다음 이벤트, "
                  "[시장] 주도주(★)·수급 확인."),
            ("t", "매매 직후 — [매매일지]에 한 줄 기록(매매근거 필수). 새로 고침하면 포트폴리오·NAV·위험이 즉시 갱신됩니다."),
            ("t", "장 마감 후 — 15:40 이후에 [전체](종목별 수급 API가 15:40 전에는 오늘 날짜를 받지 않음) → [성과] 샤프·MDD·초과수익 점검 → "
                  "[리스크] 위험기여 상위 종목 점검 → [대회종목]·[업종]으로 내일 계획."),
            ("h", "4. 수정표 ([설정] ⑤) — 대회편입·분류 고치기"),
            ("t", "종목코드(6자리)를 쓰고 바꿀 칸만 채웁니다. 빈칸 = 수정 없음. 대회편입: 추가 = 대회 종목으로 취급, 제외 = 대회 종목에서 뺌 "
                  "(고정 명단은 그대로 두고 모든 화면에서 같은 판정)."),
            ("t", "대테마·세부테마·NICS 대분류·NICS 업종·NICS 세부에 값을 쓰면 기본 분류(VALUESearch NICS·기본 테마표)보다 항상 우선합니다. "
                  "같은 종목을 여러 행에 쓰면 칸마다 아래쪽 행의 값이 우선합니다."),
            ("t", "잘못된 종목코드(6자리 영숫자 아님)·잘못된 대회편입 값·종목DB에 없는 코드는 무시되고 수정표 위에 '⚠ 수정표 n행 무시됨'이 뜹니다."),
            ("t", "반영: [모두 새로 고침] → 종목DB·주도주 ★. 포트폴리오·리스크·시세판의 섹터·대테마는 종목DB를 읽으므로 [모두 새로 고침]을 한 번 더 "
                  "합니다. 대회종목 페이지는 [시세]/[전체] — 새로 추가한 종목의 가격 이력은 [전체]에서 채웁니다."),
            ("t", "섹터 기준: [설정] sector_basis = 대분류 / 업종 / 세부(기본) / 대테마. 포트폴리오 섹터 비중·리스크 섹터 집중도·대시보드 섹터 차트·"
                  "시세판·순위의 '섹터'가 이 기준을 따릅니다(값이 없으면 세부 → 업종 → 대분류 → KRX 업종)."),
            ("h", "5. 지표 정의 (대회 평가 대비)"),
            ("t", "NAV(t) = 초기자금 + 누적 현금흐름 + Σ 보유수량×종가(수정주가). 외부 입출금이 없다는 가정의 시간가중수익률과 같습니다."),
            ("t", "샤프지수 = (일간수익률 평균 − 무위험수익률/252) ÷ 일간수익률 표준편차 × √252. 대회가 rf=0을 쓰면 'rf=0' 값을 보세요."),
            ("t", "MDD = 고점 대비 최대 하락률. 알파 = 젠센 알파(연율화). 정보비율 = 연 초과수익 ÷ 추적오차."),
            ("t", "사전 변동성·위험기여 = 현재 비중을 최근 60영업일 수익률에 적용한 가상 포트폴리오 기준. 위험기여비중 합계 = 100%."),
            ("t", "샤프를 높이는 법: 같은 기대수익이면 변동성을 줄인다 → 위험기여가 비중보다 큰 종목 축소, 상관 낮은 섹터 분산, 손절 규칙 준수."),
            ("h", "6. 데이터 구조와 한계"),
            ("t", "Power Query 원본: excel_dashboard/powerquery/*.pq (저장소 examples_llm 샘플의 URL·tr_id·파라미터를 그대로 M으로 옮김)."),
            ("t", "T_Token만 토큰을 발급합니다(1분당 1회·발급 시 알림톡). 나머지 쿼리는 캐시된 토큰만 읽습니다. 유효시간 3시간 미만일 때만 재발급."),
            ("t", "호출 제한(초당 건수) 초과 시 각 호출이 자동으로 최대 5회 재시도합니다. 관심·보유 종목이 많을수록 [모두 새로 고침]이 느려집니다"
                  "(종목당 약 4~6회 호출: 시세·수급·일봉·뉴스·장중 추정가집계)."),
            ("t", "조회에 실패하면 오류 창 대신 직전 데이터를 그대로 두고 상태 열에 ‘이전 데이터(갱신 실패: 사유)’를 적습니다 → 머리글 경고를 확인하세요."),
            ("t", "시장·순위·수급·지수 API는 모의투자 도메인에서 지원되지 않는 경우가 많아 실전 도메인(prod)을 사용합니다. 주문은 전혀 하지 않습니다(조회 전용)."),
            ("t", "VALUESearch 파일 갱신: VALUESearch에서 수집기업 목록을 다시 내보내 같은 형식(시트 Sheet2, 1행 머리글, 열 '종목코드'·"
                  "'691300.NICS 산업분류' 등)으로 [설정] vs_path 위치에 덮어씁니다(기본값 = 저장소 루트의 수집기업_valuesearch.xlsx). "
                  "그다음 [전체](분류 갱신) → [모두 새로 고침](종목DB 반영). 파일을 못 읽거나 열 이름이 다르면 직전 분류를 유지하고 상태에 사유를 남깁니다."),
            ("t", "데이터 한계: KIS 추정실적(Fwd PER)은 대회 종목의 약 21%, 목표주가는 약 58%만 있습니다 — 빈칸은 오류가 아닙니다. "
                  "KIS 추정은 한국투자증권 리서치 자체 추정이며 컨센서스가 아닙니다."),
            ("t", "당일 전용 데이터: 외인·기관 추정가집계(장중 09:30·10:00·11:20·13:20·14:30 입력 — 시세판은 개장일 09:00~15:30에만 표시), "
                  "체결금액별 매매비중·당일 매물대([조회])는 지난 날짜를 받을 수 없습니다."),
            ("t", "늦게 오는 데이터: 신용잔고는 약 2영업일 늦게 공표되고, 증시 자금은 1~2일 늦습니다(각 표의 기준일 확인). 분기 실적은 KIS 재무 자료 "
                  "갱신이 늦어 약 50종목은 최근 분기가 늦게 반영됩니다. 주식형 펀드 잔고는 평가액이라 주가 등락을 따라 움직입니다."),
            ("h", "7. 문제 해결"),
            ("t", "“유효한 접근토큰이 없습니다” → 모두 새로 고침 한 번 더. 계속되면 [데이터] > [쿼리 및 연결] > T_Token 우클릭 > 새로 고침."),
            ("t", "“EGW00133” (토큰 1분당 1회) → 1분 후 다시. “EGW00201” (초당 건수 초과) → 자동 재시도, 계속되면 [설정] 종목별 수급 조회를 N으로."),
            ("t", "“Formula.Firewall” → [데이터] > [데이터 가져오기] > [쿼리 옵션] > 현재 통합 문서 > 개인정보 > '개인 정보 수준 무시' 선택."),
            ("t", "“kis_devlp.yaml에 필수 항목이 없습니다” → [설정]의 경로 확인, 파일에 my_app / my_sec / prod 항목이 있어야 합니다."),
            ("t", "통합문서를 다른 사람에게 보낼 때 → 숨김 시트 _sys의 토큰 표 내용을 지우고 보내세요(앱키·시크릿은 원래 포함되지 않음)."),
            ("t", "버튼을 눌러도 아무 일이 없음 → 파일을 열 때 노란 보안 경고 줄의 [콘텐츠 사용]을 눌렀는지 확인하고, 닫았다 다시 열어 [콘텐츠 사용]. "
                  "빨간 줄 '이 파일의 원본을 신뢰할 수 없어 매크로를 차단했습니다'가 뜨면 파일을 닫고 탐색기 [속성]에서 [차단 해제]를 체크합니다. "
                  "계속 막히면 [파일] > [옵션] > [보안 센터] > [보안 센터 설정] > [신뢰할 수 있는 위치]에 이 폴더를 추가합니다."),
            ("t", "버튼의 상태 칸에 '실패: … 이전 데이터 표시 중' → 요약 창(또는 상태 칸)의 단계·사유를 확인하고 같은 버튼을 다시 누릅니다. "
                  "토큰 문제면 [모두 새로 고침] 후 다시. [전체]가 장 마감 전에 수급 단계에서 실패하면 15:40 이후에 다시 누릅니다."),
            ("t", "수정표 위에 '⚠ 수정표 n행 무시됨' → 그 행의 종목코드(6자리)·대회편입(추가/제외/빈칸) 값과 종목DB에 있는 코드인지 확인합니다."),
        ]
        r = 6
        for kind, text in lines:
            if kind == "h":
                r += 1
                put(ws, f"B{r}", text, bold=True, size=12, color=C["navy"])
                border(ws.Range(f"B{r}:C{r}"), "navy2", XL_MEDIUM, edges=(XL_EDGE_BOTTOM,))
            else:
                put(ws, f"C{r}", text, size=10, wrap=True)
            r += 1
        fill(ws.Range(f"B6:C{r}"), "card")

    # -------------------------------------------------------------------------------------------------------
    # -------------------------------------------------------------------------------------------------------
    def build_pages(self):
        """새 페이지(업종·대회종목·종목분석·뉴스·이벤트): 표를 적재·새로 고친 뒤 각 모듈의 build(builder). 종목분석은 tblContest(정적 표)가 있어야
        대회 종목 드롭다운 이름(대회코드목록)을 만든다 — build_static_tables가 먼저 만든다."""
        for mod in PAGE_MODULES:
            mod.build(self)

    def add_vba(self):
        """VBA 모듈을 코드 텍스트로 넣는다(파일 가져오기 금지 — 이 PC에서 CP949로 읽혀 한글이 깨짐). 넣은 뒤 다시 읽어
        .bas(Attribute 줄 제외)와 같은지 확인한다(공백 끝 차이만 무시)."""
        for path in VBA_FILES:
            comp = add_vba_module(self.wb, path)
            with open(path, encoding="utf-8-sig") as fh:
                _, kept = _vba_text(fh.read())
            expected = [ln.rstrip() for _, ln in kept]
            cm = comp.CodeModule
            n = int(cm.CountOfLines)
            actual = [ln.rstrip() for ln in str(cm.Lines(1, n)).split("\r\n")] if n else []
            while actual and not actual[-1]:
                actual.pop()
            if actual == expected:
                self.say(f"VBA 모듈 {comp.Name}: {len(actual)}줄 넣음 (원본 {os.path.basename(path)}와 같음 확인)")
            else:
                diff = next((i for i, (x, y) in enumerate(zip(actual, expected)) if x != y), min(len(actual), len(expected)))
                self.say(f"⚠ VBA 모듈 {comp.Name}: 넣은 코드가 원본과 다름(줄 수 {len(actual)} vs {len(expected)}, 첫 차이 {diff + 1}행)")

    def check_contract(self):
        """빌드 결과 계약 점검: 이름 정의(최근 조회·상태·분석코드·대회코드목록)와 버튼 도형의 매크로 연결.
        빠진 것이 있으면 실패(결과물을 출력 위치로 옮기지 않음)."""
        problems = []
        names = {}
        for nm in self.wb.Names:
            names[str(nm.Name)] = str(nm.RefersTo)
        for nm in REQUIRED_NAMES:
            if nm not in names:
                problems.append(f"이름 정의 없음: {nm}")
            elif "#REF!" in names[nm]:
                problems.append(f"이름 정의가 깨짐: {nm} = {names[nm]}")
        for sheet, shape, macro in REQUIRED_BUTTONS:
            try:
                act = str(self.ws[sheet].Shapes(shape).OnAction)
            except Exception:  # noqa: BLE001
                problems.append(f"버튼 없음: {sheet}!{shape}")
                continue
            if not act.endswith(macro):
                problems.append(f"버튼 매크로 연결 다름: {sheet}!{shape} → {act!r} (기대 {macro})")
        if problems:
            raise RuntimeError("빌드 결과 점검 실패: " + "; ".join(problems))
        self.say(f"점검: 이름 정의 {len(REQUIRED_NAMES)}개 · 버튼 {len(REQUIRED_BUTTONS)}개(매크로 연결) 확인")

    def _count_rows(self, lo, cols=None) -> int:
        """표의 데이터 행 수(cols 열이 모두 빈 행은 세지 않음 — 빈 입력표의 자리 행)."""
        body = lo.DataBodyRange
        if body is None:
            return 0
        names = [str(c.Name) for c in lo.ListColumns]
        idx = [names.index(c) for c in (cols or names) if c in names]
        vals = body.Value
        if not isinstance(vals, tuple):
            vals = ((vals,),)
        return sum(1 for row in vals if any(not migrate.is_blank(row[j]) for j in idx))

    def report_counts(self):
        """이관 행 수 대조: 원본(파일에서 읽은 행) → 새 통합문서(지금 표에 있는 행). 입력표가 다르면 빌드 실패."""
        if self.plan is None and not any(c for _, c, _ in self.seeds.values()):
            self.say("이관 없음 — 새로 만든 통합문서" + (" (샘플 포함)" if self.sample else ""))
            return
        lines, bad = [], []
        if self.plan is not None:
            for name, cols in migrate.INPUT_TABLES.items():
                t = self.plan.tables.get(name)
                new = self._count_rows(self.lo[name], cols)
                if t is None:
                    lines.append(f"  {name:<13} 원본에 표 없음 → {new}행 (기본값)")
                    continue
                ok = new == t.source_rows
                note = (" · " + " · ".join(t.notes)) if t.notes else ""
                lines.append(f"  {name:<13} {t.source_rows:>7} → {new:>7}  {'✓' if ok else '✗ 불일치'}{note}")
                if not ok:
                    bad.append(name)
            info = self.settings_info
            kv = self.lo["tblSettings"].ListColumns("키").DataBodyRange.Value
            kv = kv if isinstance(kv, tuple) else ((kv,),)
            keys_new = [str(v[0]).strip() for v in kv if v[0] is not None]
            lost = [k for k, *_ in self.plan.settings if k not in keys_new]
            written = len(info["kept"]) + len(info["added"]) + len(info.get("extra", []))
            extra = f" · 원본에만 있던 키 {len(info['extra'])}개 유지" if info.get("extra") else ""
            mark = "✓" if not lost and len(keys_new) == written else (
                "✗ 빠진 키 " + ", ".join(lost) if lost else f"✗ 표 행 {len(keys_new)}개 ≠ 넣은 키 {written}개")
            lines.append(f"  {'tblSettings':<13} {info['source']:>5}키 → {len(keys_new):>5}키  {mark}"
                         f" (사용자 값 유지 {len(info['kept'])} · 새 키 기본값 {len(info['added'])}{extra})")
            if lost or len(keys_new) != written:
                bad.append("tblSettings")
        for seed, (own, n_seed, origin) in self.seeds.items():
            had = self.plan is not None and own in self.plan.tables
            lo_own = self.lo.get(own)
            n_own = 0 if lo_own is None or lo_own.DataBodyRange is None else int(lo_own.ListRows.Count)
            if not n_seed and not had:
                lines.append(f"  {own:<13} 원본에 없음 → {n_own}행")
                continue
            mark = "✓" if n_own == n_seed else "△ (중복·빈 행 정리 또는 상태 행)"
            lines.append(f"  {own:<13} {n_seed:>7} → {n_own:>7}  {mark} ({origin or '통합문서'} → {seed} → {own})")
        if self.token is not None:
            lines.append(f"  {'tblToken':<13} 이관(남은 {self.token.minutes_left()}분) → 상태 '{self._token_status()}'")
        else:
            lines.append(f"  {'tblToken':<13} 이관 없음 → 상태 '{self._token_status()}'")
        self.say("이관 행 수 대조 (원본 → 새 통합문서):")
        for ln in lines:
            self.say(ln)
        if bad:
            raise RuntimeError("이관 행 수 불일치: " + ", ".join(bad) + " — 결과물을 출력 위치로 옮기지 않습니다(원본·백업은 그대로)")

    def finish(self):
        for t in self.temp_rows:
            try:
                self.lo[t].ListRows(1).Delete()
            except Exception as e:
                self.say(f"임시 행 삭제 실패: {t} {e}")
        for s in HIDDEN_SHEETS:
            self.ws[s].Visible = XL_SHEET_HIDDEN
        self.xl.CalculateFull()
        self.ws["대시보드"].Activate()
        self.ws["대시보드"].Range("A1").Select()
        try:
            self.wb.EnableAutoRecover = True     # create_workbook에서 끈 자동 복구를 사용자 통합문서에는 다시 켬
        except Exception:  # noqa: BLE001
            pass
        self.wb.Save()
        self.say(f"저장 완료(작업 위치): {self.out}")


# ---------------------------------------------------------------------------------------------------------------
# 실행 (이관 원본 선택 → 계획 읽기 → 토큰 확인 → 백업 → 빌드 → 결과 옮기기 → .xlsx 원본 옮기기)
# ---------------------------------------------------------------------------------------------------------------
def _cleanup_work(work: str) -> None:
    """작업 폴더의 빌드 결과(빌더 소유 — 토큰 사본이 들어 있을 수 있음)를 지우고 빈 작업 폴더도 지운다."""
    try:
        if os.path.isfile(work):
            os.remove(work)
        d = os.path.dirname(work)
        if os.path.isdir(d) and not os.listdir(d):
            os.rmdir(d)
    except OSError:
        pass


def resume_refusal(a) -> str | None:
    """숨은 개발용 옵션 --resume을 거절할 사유를 돌려준다(실행해도 되면 None).

    --resume은 이관 없이 체크포인트 파일로 다시 만든 뒤 출력 파일을 바꾸므로(덮기 전 백업은 함), 체크포인트 이후에 출력 통합문서에
    입력한 데이터(매매일지·설정·스냅샷 등)가 결과물에서 빠질 수 있다. 그래서 --out을 함께 주고 그 경로가 기본 통합문서
    (excel_dashboard/KIS_PM_Dashboard.xlsm·.xlsx — 확장자와 관계없이 같은 이름이면 .xlsm으로 바뀌어 같은 파일이 됨)가 아닐 때만
    허용한다. git 워크트리에서 실행하면 주 저장소의 기본 통합문서도 거절한다.

    Args:
        a: argparse 결과(resume·out 속성; out은 주지 않았으면 None).

    Returns:
        거절 사유(한국어 한 문장) 또는 None.

    Example:
        resume_refusal(argparse.Namespace(resume="cp.xlsm", out=None))   # → 사유 문자열
    """
    if not getattr(a, "resume", None):
        return None
    hint = "--out으로 기본 통합문서가 아닌 스크래치 경로(예: --out D:/scratch/KIS_PM_Dashboard.xlsm)를 함께 주세요"
    if getattr(a, "out", None) is None:
        return f"--resume(개발용 체크포인트 재개)은 이관 없이 다시 만들어 출력 파일을 바꾸므로 기본 통합문서에는 쓰지 않습니다 — {hint}"
    stem = os.path.normcase(os.path.splitext(os.path.realpath(os.path.abspath(a.out)))[0])
    # 이 체크아웃의 기본 통합문서와, git 워크트리에서 실행할 때 주 저장소의 기본 통합문서(실제 통합문서) 둘 다 거절
    defaults = {DEFAULT_OUT, os.path.join(main_repo_root(), os.path.basename(HERE), os.path.basename(DEFAULT_OUT))}
    if stem in {os.path.normcase(os.path.splitext(os.path.realpath(p))[0]) for p in defaults}:
        return (f"--resume(개발용 체크포인트 재개)은 기본 통합문서({DEFAULT_OUT} · .xlsx)를 출력으로 쓸 수 없습니다"
                f"(체크포인트 이후 입력한 데이터가 빠질 수 있음) — {hint}")
    return None


def run_build(a) -> int:
    """명령줄 인수로 빌드 전체를 실행하고 종료 코드를 돌려준다: 0 성공 · 1 빌드 실패 · 2 토큰 갱신 필요 · 3 이관·설정 문제(시작 전 중단)."""
    log: list[str] = []

    def say(msg: str) -> None:
        print(msg, flush=True)
        log.append(msg)

    t0 = time.time()
    refusal = resume_refusal(a)
    if refusal:
        say(f"[중단] {refusal} (Excel을 띄우지 않고 멈춤)")
        return 3
    out = os.path.abspath(a.out if a.out is not None else DEFAULT_OUT)
    if not out.lower().endswith(".xlsm"):
        new_out = os.path.splitext(out)[0] + ".xlsm"
        say(f"⚠ 결과물은 매크로 포함 형식(.xlsm)이라 출력 확장자를 바꿉니다: {new_out}")
        out = new_out
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bdir = migrate.backup_dir_for(out)
    hcsv = migrate.history_csv_path(out)
    work = os.path.join(os.path.dirname(out), BUILD_TMP_DIR, os.path.basename(out))
    log_path = os.path.join(bdir, f"{os.path.splitext(os.path.basename(out))[0]}_{stamp}_build.log")

    def finish_log(code: int) -> int:
        try:
            os.makedirs(bdir, exist_ok=True)
            with open(log_path, "w", encoding="utf-8") as fh:
                fh.write("\n".join(log) + "\n")
            print(f"빌드 로그: {log_path}", flush=True)
        except OSError:
            pass
        return code

    # 1) 이관 원본: --migrate-from > 출력 위치 .xlsm > 같은 위치 .xlsx. --no-migrate면 그 파일은 토큰 캐시에만 씀
    try:
        if a.resume:
            source, token_src = None, None
        elif a.no_migrate:
            source, token_src = None, migrate.select_source(out, a.migrate_from)
        else:
            source = migrate.select_source(out, a.migrate_from)
            token_src = source
    except migrate.MigrationError as e:
        say(f"[중단] {e}")
        return 3
    # 2) 잠금 확인 — 원본은 옮기거나 덮을 수 있고 출력은 덮으므로 Excel에서 열려 있으면 시작하지 않음
    for p in sorted({x for x in (source, out) if x and os.path.exists(x)}):
        if migrate.is_locked(p):
            say(f"[중단] 파일이 다른 프로그램(Excel)에서 열려 있습니다 — 닫고 다시 실행하세요: {p}")
            return 3
        if os.path.exists(migrate.owner_file(p)):
            say(f"⚠ Excel 소유자 파일(~$)이 있습니다. 이 파일을 열어 둔 Excel이 있다면 저장하지 않은 변경은 이관되지 않습니다: {p}")
    # 3) 이관 계획 (Excel로 열지 않고 파일을 직접 읽음 — 토큰 포함)
    plan, token, token_note, est_snap = None, None, "", None
    if source:
        say(f"이관 원본: {source}")
        try:
            plan = migrate.read_plan(source, SHEETS, hcsv)
        except migrate.MigrationError as e:
            say(f"[중단] {e}")
            say("  이관 없이 새로 만들려면 --no-migrate (기존 파일은 백업 폴더에 복사해 두고, 스냅샷은 history\\est_snap.csv에서 복원)")
            return 3
        token, token_note = plan.token, plan.token_note
        found = [f"{n} {t.source_rows}행" + (f"({t.origin})" if t.origin != "통합문서" else "") for n, t in plan.tables.items()]
        say(f"  읽은 표: tblSettings {plan.settings_count}키 · " + " · ".join(found))
        for s in plan.unknown_sheets:
            say(f"⚠ 사용자 추가 시트 '{s}'는 이관하지 않습니다(백업본에 남아 있음)")
        for w in plan.warnings:
            say(f"⚠ {w}")
    elif not a.resume:
        say("이관 안 함(--no-migrate) — 입력표를 새로 만듭니다" if a.no_migrate else "이관 원본 없음 — 처음 빌드")
        if token_src:
            token, token_note = migrate.read_token_only(token_src)
            say(f"  토큰 캐시 원천: {token_src} ({token_note})")
        try:
            est_snap = migrate.read_est_snap_csv(hcsv)
        except (OSError, csv.Error, UnicodeDecodeError) as e:
            say(f"⚠ 이력 CSV를 읽지 못함({type(e).__name__}): {hcsv}")
        if est_snap is not None:
            say(f"  tblEstSnap: 이력 CSV에서 복원 {len(est_snap.rows)}행 ← {hcsv}")
    # 4) 토큰 확인 (남은 시간이 기준 미만이면 Excel을 띄우지 않고 멈춤 — 갱신은 원본을 열어 T_Token으로)
    if token is not None:
        left = token.minutes_left()
        if left < TOKEN_MIN_MINUTES:
            say(f"[중단] 토큰 갱신 필요: 토큰 남은 시간 {left}분 (기준 {TOKEN_MIN_MINUTES}분). 원본 통합문서를 Excel로 열어 "
                f"[데이터] > [모두 새로 고침]으로 토큰을 갱신·저장한 뒤 다시 실행하세요(Excel을 띄우지 않고 멈춤).")
            return 2
        say(f"토큰 캐시: 남은 {left}분 (기준 {TOKEN_MIN_MINUTES}분 이상 — 새 통합문서에 옮겨 재사용, 값은 표시하지 않음)")
    elif plan is not None:
        say(f"[중단] 토큰 갱신 필요: 이관 원본에 쓸 수 있는 토큰이 없습니다({token_note}). 원본 통합문서를 Excel로 열어 "
            f"[데이터] > [모두 새로 고침]으로 토큰을 받은 뒤 다시 실행하세요.")
        return 2
    # 5) 설정 파일: --cfg를 주면 그 경로, 이관이면 원본 설정의 cfg_path, 아니면 기본 경로
    cfg_explicit = a.cfg is not None
    if cfg_explicit:
        cfg = os.path.abspath(a.cfg)
    elif plan is not None:
        v = next((s[2] for s in plan.settings if s[0] == "cfg_path"), None)
        cfg = str(v).strip().strip('"') if not migrate.is_blank(v) else DEFAULT_CFG
    else:
        cfg = DEFAULT_CFG
    if not os.path.exists(cfg):
        say(f"[중단] KIS 설정 파일이 없습니다: {cfg}  (--cfg 로 경로 지정)")
        return 3
    # 6) 백업 — 원본(복사), 그리고 덮어쓸 출력 파일이 원본과 다르면 그것도
    backups: dict[str, str] = {}
    try:
        if source:
            backups[source] = migrate.backup_copy(source, bdir, stamp)
            say(f"백업: {source} → {backups[source]}")
        if os.path.exists(out) and (source is None or os.path.normcase(out) != os.path.normcase(source)):
            backups[out] = migrate.backup_copy(out, bdir, stamp)
            say(f"백업(덮어쓸 출력 파일): {out} → {backups[out]}")
    except (OSError, migrate.MigrationError) as e:
        say(f"[중단] 백업 실패: {e}")
        return finish_log(3)
    # 7) 빌드 (작업 폴더에서) — 실패하면 출력 위치의 파일은 그대로
    sample = plan is None and not a.empty and not a.resume
    b = Builder(out, cfg, sample, a.visible, a.seed_token, work_path=work, plan=plan, token=token, est_snap=est_snap,
                vbom_lock=a.vbom_lock, log=log, cfg_explicit=cfg_explicit)
    try:
        b.run(resume=a.resume, checkpoint=a.checkpoint)
    except Exception as e:  # noqa: BLE001 — 어떤 실패든 사유를 남기고 작업 파일 정리
        import traceback
        say(f"[실패] 빌드 중 오류: {type(e).__name__}: {e}")
        say(traceback.format_exc())
        _cleanup_work(work)
        say("출력 위치의 기존 파일·이관 원본은 바뀌지 않았습니다" + (f" (백업: {', '.join(backups.values())})" if backups else ""))
        return finish_log(1)
    # 8) 결과를 출력 위치로 (덮는 파일은 6)에서 백업함)
    try:
        os.replace(work, out)
    except OSError as e:
        say(f"[실패] 결과물을 출력 위치로 옮기지 못했습니다({e}) — 결과물은 {work}에 있습니다")
        return finish_log(1)
    _cleanup_work(work)
    say(f"출력: {out} ({os.path.getsize(out) / 1e6:.1f} MB)")
    # 9) .xlsx 원본은 백업 폴더로 옮김(살아 있는 통합문서는 .xlsm 하나만 남게 — 다음 빌드가 오래된 파일에서 이관하지 않도록)
    if source and source.lower().endswith(".xlsx") and os.path.normcase(source) != os.path.normcase(out):
        try:
            say(migrate.move_into_backup(source, backups.get(source, ""), bdir, stamp))
        except OSError as e:
            say(f"⚠ .xlsx 원본을 백업 폴더로 옮기지 못했습니다({e}) — 직접 옮기세요: {source}")
    say(f"레지스트리: {b.vbom_message}")
    say(f"완료 ({time.time() - t0:.0f}s)")
    return finish_log(0)


def main():
    ap = argparse.ArgumentParser(description="KIS PM 일일 대시보드(.xlsm) 생성 — 기존 통합문서에서 자동 이관")
    ap.add_argument("--out", default=None, help="출력 .xlsm 경로 (기본 excel_dashboard/KIS_PM_Dashboard.xlsm)")
    ap.add_argument("--cfg", default=None, help="kis_devlp.yaml 경로 (기본: 이관한 설정의 cfg_path, 없으면 ~/KIS/config/kis_devlp.yaml)")
    ap.add_argument("--empty", action="store_true", help="(이관 원본이 없을 때) 샘플 매매·관심종목 없이 생성")
    ap.add_argument("--visible", action="store_true")
    ap.add_argument("--migrate-from", help="이관 원본 통합문서(.xlsm/.xlsx). 없으면 출력 위치의 .xlsm → 같은 위치의 .xlsx")
    ap.add_argument("--no-migrate", action="store_true", help="이관 없이 새로 생성(덮어쓸 파일은 백업). 이력 CSV의 스냅샷은 복원")
    ap.add_argument("--checkpoint", help=argparse.SUPPRESS)
    ap.add_argument("--resume", help=argparse.SUPPRESS)
    ap.add_argument("--seed-token", help=argparse.SUPPRESS)
    ap.add_argument("--vbom-lock", help=argparse.SUPPRESS)     # 개발용: AccessVBOM 잠금 파일(병렬 작업 상호 배제)
    sys.exit(run_build(ap.parse_args()))


if __name__ == "__main__":
    main()
