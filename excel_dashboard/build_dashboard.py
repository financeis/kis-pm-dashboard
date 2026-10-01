"""KIS PM 일일 대시보드(Excel + Power Query) 생성 스크립트.

    python excel_dashboard/build_dashboard.py            # 샘플 매매일지 포함, 데이터 새로 고침까지 수행
    python excel_dashboard/build_dashboard.py --empty    # 매매일지·관심종목을 비운 상태로 생성
    python excel_dashboard/build_dashboard.py --cfg D:/my/kis_devlp.yaml   # 설정 파일 경로 지정

- Power Query(M) 원본은 excel_dashboard/powerquery/*.pq 에 있으며, 이 스크립트가 통합문서에 그대로 넣습니다.
- 앱키/시크릿은 통합문서에 저장하지 않고 ~/KIS/config/kis_devlp.yaml 을 Power Query가 직접 읽습니다.
- Windows + Microsoft 365 Excel 필요 (동적 배열 함수: LET, FILTER, SORTBY, TAKE, VSTACK, HSTACK, XLOOKUP).
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xl_helpers import (  # noqa: E402
    C, NF, FONT, XL_CENTER, XL_LEFT, XL_RIGHT, XL_TOP, XL_VCENTER, XL_SHEET_HIDDEN, XL_LINE, XL_AREA,
    XL_COLUMN_CLUSTERED, XL_BAR_CLUSTERED, XL_LEGEND_TOP, XL_LEGEND_BOTTOM, XL_SECONDARY, XL_EDGE_BOTTOM,
    XL_MEDIUM, add_calc_column, add_names, add_query, border, bottom_line, card, cf_databar, cf_expr, cf_heat, chart,
    ensure_table_style, excel_pid, excel_serial, fill, freeze, hyperlink, load_query, new_excel, put, quit_excel, refresh, rgb,
    section, series, set_col_format,
    set_nf, set_palette, sheet_setup, style, style_axes, title_bar, validation_list, write_table,
)

HERE = os.path.dirname(os.path.abspath(__file__))
PQ_DIR = os.path.join(HERE, "powerquery")
DEFAULT_OUT = os.path.join(HERE, "KIS_PM_Dashboard.xlsx")
DEFAULT_CFG = os.path.join(os.path.expanduser("~"), "KIS", "config", "kis_devlp.yaml")

SHEETS = ["대시보드", "시장", "포트폴리오", "리스크", "성과", "시세판", "매매일지", "매매분석", "종목DB", "설정", "가이드", "_data", "_sys"]

# 쿼리 → (시트, 위치, 표이름). 순서 = 빌드 시 새로 고침 순서
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

QUERY_DESC = {
    "T_Token": "KIS 접근토큰 캐시(자기참조). 남은 유효시간 3시간 미만일 때만 재발급",
    "T_Universe": "KOSPI·KOSDAQ 종목 마스터(인증 불필요)",
    "T_IndexNow": "국내 지수 현재가·시장폭", "T_IndexHist": "국내 지수 일별", "T_Sector": "업종 등락",
    "T_Flow": "시장별 투자자 순매수", "T_Global": "해외지수·환율·금리", "T_Rank": "순위(거래대금·등락률·수급)",
    "T_Trades": "매매 원장", "T_Holdings": "보유종목", "T_Positions": "종목별 손익(청산 포함)",
    "T_Quote": "보유·관심 시세판", "T_PriceHist": "일봉+기술지표", "T_NAV": "일별 순자산·벤치마크",
    "T_Risk": "사전 위험(변동성·베타·위험기여·VaR)",
}

# ---------------------------------------------------------------------------------------------------------------
# 입력표 기본값
# ---------------------------------------------------------------------------------------------------------------
def settings_rows(cfg_path: str, sample: bool):
    start = dt.date(2026, 9, 1) if sample else dt.date.today()
    end = dt.date(2026, 10, 30) if sample else dt.date.today() + dt.timedelta(days=61)
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
F_TRADE_CHK = ('=IF([@종목코드]="","",LET(e,XLOOKUP(' + F_TRADE_CODE + ',tblUniverse[종목코드],tblUniverse[대회편입],"X"),'
               'IF(e="X","⚠ 종목DB에 없는 코드",IF(e="N","⚠ 거래정지·SPAC 등",IF(OR([@구분]="매수",[@구분]="매도"),"✓","⚠ 구분 확인")))))')


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
def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Builder:
    def __init__(self, out_path: str, cfg_path: str, sample: bool, visible: bool, seed_token: str | None = None):
        self.out = os.path.abspath(out_path)
        self.cfg_path = cfg_path
        self.sample = sample
        self.xl = new_excel(visible)
        self.pid = excel_pid(self.xl)
        self.seed_token = seed_token
        self.temp_rows = []   # 빈 PQ 표에 임시로 넣은 행 (계산열·서식 정의용, 저장 직전 삭제)
        self.wb = None
        self.ws = {}
        self.lo = {}
        self.log = []

    # -------------------------------------------------------------------------------------------------------
    def say(self, msg):
        print(msg, flush=True)
        self.log.append(msg)

    def run(self, resume: str | None = None, checkpoint: str | None = None):
        try:
            if resume:
                self.open_checkpoint(resume)
            else:
                self.create_workbook()
                self.build_inputs()
                self.add_queries()
                self.load_all()
                self.refresh_all()      # 표의 열 구조는 첫 새로 고침 후에 생기므로 수식보다 먼저 실행
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
            self.finish()
        finally:
            try:
                if self.wb is not None:
                    self.wb.Close(False)
            except Exception:
                pass
            self.wb, self.ws, self.lo = None, {}, {}
            quit_excel(self.xl, self.pid)

    # -------------------------------------------------------------------------------------------------------
    def create_workbook(self):
        if os.path.exists(self.out):
            os.remove(self.out)
        wb = self.xl.Workbooks.Add()
        self.wb = wb
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
        # 먼저 저장해 두어야 차트의 이름 참조가 최종 파일명을 가리킴
        wb.SaveAs(self.out, 51)
        self.say(f"통합문서 생성: {self.out}")

    def open_checkpoint(self, path: str):
        """개발용: 새로 고침까지 끝난 체크포인트 파일을 열어 이후 단계만 다시 실행."""
        import shutil
        shutil.copyfile(path, self.out)
        self.wb = self.xl.Workbooks.Open(self.out)
        self.table_style = "PM Navy" if ensure_table_style(self.wb) is not None else "TableStyleLight9"
        for ws in self.wb.Worksheets:
            self.ws[ws.Name] = ws
            for lo in ws.ListObjects:
                self.lo[lo.Name] = lo
        self.say(f"체크포인트 열기: {path} (표 {len(self.lo)}개)")

    # -------------------------------------------------------------------------------------------------------
    def build_inputs(self):
        ws = self.ws["설정"]
        sheet_setup(ws, bg=False, zoom=90, tab="#6B7785",
                    widths={"A": 2, "B": 15, "C": 20, "D": 44, "E": 58, "F": 3, "G": 10, "H": 16, "I": 12, "J": 10,
                            "K": 10, "L": 44, "M": 3, "N": 8, "O": 16, "P": 9, "Q": 11, "R": 7, "S": 3, "T": 12, "U": 20})
        lo = write_table(ws, 6, 2, ["키", "항목", "값", "설명"], settings_rows(self.cfg_path, self.sample), "tblSettings",
                         style_name=self.table_style)
        self.lo["tblSettings"] = lo
        rows = settings_rows(self.cfg_path, self.sample)
        for i, (k, _, v, _) in enumerate(rows):
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
        validation_list(lo.DataBodyRange.Cells(5, 3), "KOSPI,KOSDAQ,혼합")
        validation_list(lo.DataBodyRange.Cells(18, 3), "전체,코스피,코스닥")
        validation_list(lo.DataBodyRange.Cells(19, 3), "Y,N")

        watch = SAMPLE_WATCH if self.sample else []
        wrows = [[c, None, g, a, t, p] for (c, g, a, t, p) in watch]
        lo = write_table(ws, 6, 7, ["종목코드", "종목명", "그룹", "관심가", "목표가", "투자포인트"], wrows, "tblWatch",
                         text_cols=("종목코드",), style_name=self.table_style)
        for c in ("관심가", "목표가"):
            set_nf(lo.ListColumns(c).DataBodyRange, NF["krw"])
        fill(lo.ListColumns("종목코드").DataBodyRange, "input")
        self.lo["tblWatch"] = lo

        lo = write_table(ws, 6, 14, ["구분", "이름", "시장코드", "심볼", "순서"], MACRO_ROWS, "tblMacro",
                         style_name=self.table_style)
        self.lo["tblMacro"] = lo
        lo = write_table(ws, 6, 20, ["휴장일", "설명"], HOLIDAYS, "tblHolidays", date_cols=("휴장일",),
                         style_name=self.table_style)
        self.lo["tblHolidays"] = lo

        # 매매일지
        ws = self.ws["매매일지"]
        sheet_setup(ws, bg=False, zoom=90, tab="#E8890C",
                    widths={"A": 2, "B": 11, "C": 9, "D": 16, "E": 6, "F": 8, "G": 11, "H": 13, "I": 9, "J": 9,
                            "K": 8, "L": 36, "M": 11, "N": 11, "O": 18, "P": 22})
        trades = sample_trades() if self.sample else []
        lo = write_table(ws, 11, 2, TRADE_HEADERS, trades, "tblTrades", text_cols=("종목코드",), date_cols=("일자",),
                         style_name=self.table_style)
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
        self.say("입력표 생성 완료 (설정·관심종목·해외지표·휴장일·매매일지)")

    # -------------------------------------------------------------------------------------------------------
    def add_queries(self):
        files = sorted(glob.glob(os.path.join(PQ_DIR, "*.pq")))
        for f in files:
            name = os.path.splitext(os.path.basename(f))[0]
            with open(f, encoding="utf-8") as fh:
                add_query(self.wb, name, fh.read(), QUERY_DESC.get(name, ""))
        self.say(f"Power Query {len(files)}개 추가")

    def load_all(self):
        for q, sheet, cell, tname in LOADS:
            st = "TableStyleLight1" if sheet.startswith("_") else self.table_style
            lo = load_query(self.ws[sheet], q, cell, tname, background=True, style_name=st)
            self.lo[tname] = lo
        # 토큰 쿼리: 파일 열 때 + 30분마다 자동 확인(유효하면 재사용, 알림톡 없음)
        qt = self.lo["tblToken"].QueryTable
        qt.RefreshOnFileOpen = True
        qt.RefreshPeriod = 30
        self.say("쿼리 → 표 로드 설정 완료")

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

    def refresh_all(self):
        self.say("데이터 새로 고침 (KIS API 호출)…")
        t0 = time.time()
        if self.seed_token:
            self.seed_token_cache()
        failed = []
        for q, sheet, cell, tname in LOADS:
            ok, msg, n = refresh(self.lo[tname], q)
            self.say(("  ✓ " if ok else "  ✗ ") + msg)
            if not ok:
                failed.append((q, tname))
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
        # 숨김 시트
        for s in ("_data", "_sys"):
            ws = self.ws[s]
            ws.Cells.Font.Name = FONT
            ws.Cells.Font.Size = 9
            put(ws, "B1", "※ Power Query 원본 데이터 (수정하지 마세요)", size=9, color=C["muted"])
        self.ws["_sys"].Range("C:C").ColumnWidth = 12
        put(self.ws["_sys"], "B1", "※ KIS 접근토큰 캐시 — 통합문서를 외부에 공유할 때는 이 표의 내용을 지우세요(24시간 유효)",
            size=9, color=C["up"])

    # -------------------------------------------------------------------------------------------------------
    def nav_links(self, ws, row: int, start_col: int = 2):
        items = [("대시보드", "대시보드"), ("시장", "시장"), ("포트폴리오", "포트폴리오"), ("리스크", "리스크"), ("성과", "성과"),
                 ("시세판", "시세판"), ("매매일지", "매매일지"), ("매매분석", "매매분석"), ("종목DB", "종목DB"), ("설정", "설정"),
                 ("가이드", "가이드")]
        c = start_col
        for label, sheet in items:
            cell = ws.Cells(row, c).Address
            hyperlink(ws, cell, f"'{sheet}'!A1", "› " + label)
            c += 2

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
        # 시장 주도주 3종
        blocks = [("B", "거래대금 상위", "거래대금상위", "tblRank[거래대금억]", NF["eok"], "거래대금"),
                  ("J", "외국인 순매수 상위", "외국인순매수", "tblRank[지표]", NF["eok_pl"], "순매수"),
                  ("R", "상승률 상위 (ETF·SPAC 제외)", "상승률상위", "tblRank[거래대금억]", NF["eok"], "거래대금")]
        for col, title, lst, extra, nf_extra, extra_h in blocks:
            c0 = ws.Range(f"{col}49").Column
            section(ws, ws.Range(ws.Cells(49, c0), ws.Cells(49, c0 + 7)).Address, title)
            for off, h in ((0, "#"), (1, "종목"), (4, "현재가"), (5, "등락률"), (6, extra_h)):
                put(ws, ws.Cells(50, c0 + off).Address, h, bold=True, size=8, color=C["muted"])
            bottom_line(ws.Range(ws.Cells(50, c0), ws.Cells(50, c0 + 7)))
            cond = f'tblRank[목록]="{lst}"'
            refs = [(0, "tblRank[순위]", "0"), (1, "tblRank[종목명]", None), (4, "tblRank[현재가]", NF["krw"]),
                    (5, "tblRank[등락률]", NF["pct_arrow"]), (6, extra, nf_extra)]
            for off, ref, nf in refs:
                cell = ws.Cells(51, c0 + off)
                put(ws, cell.Address, formula=f'=IFERROR(TAKE(FILTER({ref},{cond}),10),"")', size=9)
                set_nf(ws.Range(ws.Cells(51, c0 + off), ws.Cells(60, c0 + off)), nf or "General")
                ws.Range(ws.Cells(51, c0 + off), ws.Cells(60, c0 + off)).Font.Size = 9
            fill(ws.Range(ws.Cells(50, c0), ws.Cells(60, c0 + 7)), "card")
            style(ws.Range(ws.Cells(51, c0), ws.Cells(60, c0)), color=C["muted"], h=XL_CENTER)
            # 보유·관심 종목 강조
            first = ws.Cells(51, c0 + 1).Address.replace("$", "")
            cf_expr(ws.Range(ws.Cells(51, c0 + 1), ws.Cells(60, c0 + 3)),
                    f'=COUNTIF(tblQuote_종목명,{first})>0', font=C["accent"], bold=True)
        put(ws, "B62", "※ 파란 굵은 글씨 = 보유·관심 종목 | 상승=빨강·하락=파랑 (국내 관례) | 데이터: 한국투자증권 Open API",
            size=8, color=C["muted"])
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
        # 주도주 순위 6블록 (2행 × 3)
        blocks = [
            ("거래대금상위", "거래대금 상위", "tblRank[거래대금억]", NF["eok"], "거래대금"),
            ("상승률상위", "상승률 상위", "tblRank[거래대금억]", NF["eok"], "거래대금"),
            ("하락률상위", "하락률 상위", "tblRank[거래대금억]", NF["eok"], "거래대금"),
            ("외국인순매수", "외국인 순매수 (가집계)", "tblRank[지표]", NF["eok_pl"], "순매수"),
            ("기관순매수", "기관 순매수 (가집계)", "tblRank[지표]", NF["eok_pl"], "순매수"),
            ("거래량급증", "거래량 급증", "tblRank[지표]", '0%', "증가율"),
        ]
        for bi, (lst, title, extra, nf_extra, extra_h) in enumerate(blocks):
            row0 = 60 if bi < 3 else 78
            c0 = 2 + (bi % 3) * 8
            section(ws, ws.Range(ws.Cells(row0, c0), ws.Cells(row0, c0 + 7)).Address, title)
            for off, h in ((0, "#"), (1, "종목"), (3, "섹터"), (5, "현재가"), (6, "등락률"), (7, extra_h)):
                put(ws, ws.Cells(row0 + 1, c0 + off).Address, h, bold=True, size=8, color=C["muted"])
            bottom_line(ws.Range(ws.Cells(row0 + 1, c0), ws.Cells(row0 + 1, c0 + 7)))
            cond = f'tblRank[목록]="{lst}"'
            for off, ref, nf in ((0, "tblRank[순위]", "0"), (1, "tblRank[종목명]", None), (3, "tblRank[섹터]", None),
                                 (5, "tblRank[현재가]", NF["krw"]), (6, "tblRank[등락률]", NF["pct_arrow"]), (7, extra, nf_extra)):
                put(ws, ws.Cells(row0 + 2, c0 + off).Address, formula=f'=IFERROR(TAKE(FILTER({ref},{cond}),15),"")', size=9)
                rr = ws.Range(ws.Cells(row0 + 2, c0 + off), ws.Cells(row0 + 16, c0 + off))
                set_nf(rr, nf or "General")
                rr.Font.Size = 9
            fill(ws.Range(ws.Cells(row0 + 1, c0), ws.Cells(row0 + 16, c0 + 7)), "card")
            style(ws.Range(ws.Cells(row0 + 2, c0), ws.Cells(row0 + 16, c0)), color=C["muted"], h=XL_CENTER)
            style(ws.Range(ws.Cells(row0 + 2, c0 + 3), ws.Cells(row0 + 16, c0 + 3)), color=C["muted"], size=8)
            first = ws.Cells(row0 + 2, c0 + 1).Address.replace("$", "")
            cf_expr(ws.Range(ws.Cells(row0 + 2, c0 + 1), ws.Cells(row0 + 16, c0 + 2)),
                    f'=COUNTIF(tblQuote_종목명,{first})>0', font=C["accent"], bold=True)
        put(ws, "B95", "※ 외국인/기관 순매수는 장중 가집계(09:30·11:20·13:20·14:30 입력) 기준 | 파란 굵은 글씨 = 보유·관심 종목",
            size=8, color=C["muted"])
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

        # 섹터 비중 (동적 배열) T12~
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
        title_bar(ws, "QUOTE BOARD", "보유·관심 종목 시세 · 밸류에이션 · 52주 · 수급 · 기술적 신호 | 관심종목은 [설정] 시트에서 추가",
                  "B2:BD3", right_formula=self.HEADER_RIGHT, right_cell="R2", right2_formula=self.HEADER_STATUS, right2_cell="R3")
        fill(ws.Range("B4:BD4"), "navy2")
        self.nav_links(ws, 4)
        lo = self.lo["tblQuote"]
        widths = {"구분": 8, "종목코드": 9, "종목명": 15, "시장": 9, "섹터": 12, "현재가": 11, "전일대비": 9,
                  "등락률": 9, "시가": 10, "고가": 10, "저가": 10, "전일종가": 10, "거래량": 12, "거래대금억": 10,
                  "전일대비거래량": 9, "시가총액억": 12, "PER": 7, "PBR": 6, "EPS": 9, "BPS": 9, "외국인소진율": 8,
                  "고52주": 10, "고52주일": 10, "고52주대비": 9, "저52주": 10, "저52주일": 10, "저52주대비": 9,
                  "외국인당일주": 11, "프로그램당일주": 11, "외국인5일억": 10, "기관5일억": 10, "개인5일억": 10,
                  "외국인20일억": 10, "기관20일억": 10, "상한가": 10, "하한가": 10, "유의": 10, "DB유의": 10,
                  "상태": 6, "조회시각": 15, "거래량비율": 8, "MA20": 10, "MA60": 10, "추세": 6, "RSI14": 7,
                  "수익률20일": 9, "수익률60일": 9, "변동성20": 8, "교차": 10, "관심가": 10, "목표가": 10,
                  "투자포인트": 24, "신호": 34}
        for c, w in widths.items():
            set_col_format(lo, c, width=w)
        nfs = {"현재가": NF["krw"], "전일대비": "+#,##0;-#,##0;0", "등락률": NF["pct_arrow"], "시가": NF["krw"], "고가": NF["krw"],
               "저가": NF["krw"], "전일종가": NF["krw"], "거래량": NF["num0"], "거래대금억": NF["num0"], "전일대비거래량": NF["pct1"],
               "시가총액억": NF["num0"], "PER": NF["num1"], "PBR": NF["ratio2"], "EPS": NF["num0"], "BPS": NF["num0"],
               "외국인소진율": NF["pct1"], "고52주": NF["krw"], "고52주일": NF["date"], "고52주대비": NF["pct1"],
               "저52주": NF["krw"], "저52주일": NF["date"], "저52주대비": NF["pct1"], "외국인당일주": "+#,##0;-#,##0;0",
               "프로그램당일주": "+#,##0;-#,##0;0", "외국인5일억": NF["eok_pl"], "기관5일억": NF["eok_pl"], "개인5일억": NF["eok_pl"],
               "외국인20일억": NF["eok_pl"], "기관20일억": NF["eok_pl"], "상한가": NF["krw"], "하한가": NF["krw"], "조회시각": NF["dt"]}
        for c, nf in nfs.items():
            set_col_format(lo, c, nf=nf)
        set_col_format(lo, "종목명", bold=True)
        style(lo.ListColumns("신호").DataBodyRange, color=C["warn"], size=9)
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
        # 차트 도우미 (숨김 열)
        base = 60  # BH 열 근처 (표 오른쪽)
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
        section(ws, "B25:R25", "시세판 (보유 → 관심 순)", "행을 클릭해 필터/정렬 가능 · 신호: 52주고가근접·거래량급증·골든크로스·RSI·관심가·수급", "R25")
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
        title_bar(ws, "UNIVERSE", "KOSPI·KOSDAQ 개별 종목 마스터 (KIS 종목정보 파일) — ETF·ETN·리츠·펀드 제외 | 대회편입 N = 거래정지·정리매매·SPAC",
                  "B2:Y3", right_formula='="종목 수 "&TEXT(COUNTA(tblUniverse[종목코드]),"#,##0")', right_cell="Y2")
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
        section(ws, "B20:Y20", "전체 종목 (시가총액순) — 필터 버튼으로 시장·섹터·유의사항 걸러보기")
        lo = self.lo["tblUniverse"]
        for c, (w, nf) in {"종목코드": (10, None), "종목명": (17, None), "시장": (9, None), "섹터": (13, None), "규모": (7, None),
                           "지수편입": (11, None), "유의사항": (16, None), "대회편입": (9, None), "기준가": (11, NF["krw"]),
                           "시가총액억": (12, NF["num0"]), "매출액억": (11, NF["num0"]), "영업이익억": (11, NF["num0"]),
                           "영업이익률": (11, NF["pct1"]), "ROE": (8, NF["pct1"]), "재무기준": (10, None), "상장일": (11, NF["date"]),
                           "상장주식수": (14, NF["num0"]), "대분류": (9, None), "우선주": (8, None), "SPAC": (8, None),
                           "거래정지": (9, None), "관리종목": (9, None), "시장경고": (9, None), "종류": (7, None)}.items():
            set_col_format(lo, c, nf=nf, width=w)
        freeze(ws, 21, 3)
        self.say("종목DB 시트 완료")

    # -------------------------------------------------------------------------------------------------------
    def build_settings_sheet(self):
        ws = self.ws["설정"]
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "SETTINGS", "노란 칸만 수정하세요 · 표 아래 행에 입력하면 표가 자동 확장됩니다", "B2:U3")
        fill(ws.Range("B4:U4"), "navy2")
        self.nav_links(ws, 4)
        put(ws, "B5", "① 기본 설정", bold=True, size=11, color=C["navy"])
        put(ws, "G5", "② 관심종목 (시세판·알림 대상)", bold=True, size=11, color=C["navy"])
        put(ws, "N5", "③ 해외·환율·금리 (N=지수, X=환율, I=금리)", bold=True, size=11, color=C["navy"])
        put(ws, "T5", "④ 휴장일 (D-day 계산)", bold=True, size=11, color=C["navy"])
        put(ws, "B28", "※ 앱키·시크릿은 이 통합문서에 저장되지 않습니다. Power Query가 위 경로의 kis_devlp.yaml(저장소 샘플코드와 같은 파일)을 직접 읽습니다.",
            size=9, color=C["muted"])
        put(ws, "B29", "※ 휴장일 목록은 참고용입니다. 한국거래소(KRX) 휴장일 공지로 확인 후 필요하면 수정하세요.", size=9, color=C["muted"])
        freeze(ws, 6)

    # -------------------------------------------------------------------------------------------------------
    def build_guide(self):
        ws = self.ws["가이드"]
        sheet_setup(ws, zoom=90, tab="#6B7785", widths={"A": 2, "B": 4, "C": 110})
        ws.Rows(1).RowHeight = 6
        title_bar(ws, "GUIDE", "처음 설정 · 매일 루틴 · 지표 정의 · 문제 해결", "B2:C3")
        fill(ws.Range("B4:C4"), "navy2")
        self.nav_links(ws, 4)
        lines = [
            ("h", "1. 처음 한 번만"),
            ("t", "① 저장소 README 3.5절대로 ~/KIS/config/kis_devlp.yaml 에 실전투자 앱키(my_app)·시크릿(my_sec)을 입력합니다. (시세·순위 API는 실전 앱키 필요)"),
            ("t", "② [설정] 시트에서 대회 시작일·종료일·초기자금·수수료율·거래세율·벤치마크를 대회 규정에 맞게 수정합니다."),
            ("t", "③ [매매일지]의 [샘플] 행과 [설정]의 샘플 관심종목을 지우고 내 종목으로 바꿉니다."),
            ("t", "④ 파일을 열 때 노란 '보안 경고' 줄이 뜨면 [콘텐츠 사용]을 누릅니다. '웹 콘텐츠 액세스' 창이 뜨면 [익명] → [연결]."),
            ("t", "⑤ [데이터] > [모두 새로 고침] (Ctrl+Alt+F5). 첫 실행이나 토큰 재발급 직후 오류가 보이면 한 번 더 새로 고침합니다."),
            ("h", "2. 매일 루틴 (실제 운용역처럼)"),
            ("t", "08:30 장 전 — 모두 새로 고침 → [시장] 해외지수·환율·금리 확인 → [대시보드] 알림(손절/목표/비중) 확인 → 오늘 매매 계획 메모."),
            ("t", "장중 — 필요할 때마다 새로 고침(약 20~40초). [시세판] 신호(거래량급증·신고가근접·골든크로스)와 [시장] 주도주·수급 확인."),
            ("t", "매매 직후 — [매매일지]에 한 줄 기록(매매근거 필수). 새로 고침하면 포트폴리오·NAV·위험이 즉시 갱신됩니다."),
            ("t", "15:40 장 마감 후 — 새로 고침 → [성과] 샤프·MDD·초과수익 점검 → [리스크] 위험기여 상위 종목 점검 → 내일 계획."),
            ("h", "3. 지표 정의 (대회 평가 대비)"),
            ("t", "NAV(t) = 초기자금 + 누적 현금흐름 + Σ 보유수량×종가(수정주가). 외부 입출금이 없다는 가정의 시간가중수익률과 같습니다."),
            ("t", "샤프지수 = (일간수익률 평균 − 무위험수익률/252) ÷ 일간수익률 표준편차 × √252. 대회가 rf=0을 쓰면 'rf=0' 값을 보세요."),
            ("t", "MDD = 고점 대비 최대 하락률. 알파 = 젠센 알파(연율화). 정보비율 = 연 초과수익 ÷ 추적오차."),
            ("t", "사전 변동성·위험기여 = 현재 비중을 최근 60영업일 수익률에 적용한 가상 포트폴리오 기준. 위험기여비중 합계 = 100%."),
            ("t", "샤프를 높이는 법: 같은 기대수익이면 변동성을 줄인다 → 위험기여가 비중보다 큰 종목 축소, 상관 낮은 섹터 분산, 손절 규칙 준수."),
            ("h", "4. 데이터 구조"),
            ("t", "Power Query 원본: excel_dashboard/powerquery/*.pq (저장소 examples_llm 샘플의 URL·tr_id·파라미터를 그대로 M으로 옮김)."),
            ("t", "T_Token만 토큰을 발급합니다(1분당 1회·발급 시 알림톡). 나머지 쿼리는 캐시된 토큰만 읽습니다. 유효시간 3시간 미만일 때만 재발급."),
            ("t", "호출 제한(초당 건수) 초과 시 각 호출이 자동으로 최대 5회 재시도합니다. 관심종목이 많을수록 새로 고침이 느려집니다(종목당 약 4회 호출)."),
            ("t", "조회에 실패하면 오류 창 대신 직전 데이터를 그대로 두고 상태 열에 ‘이전 데이터(갱신 실패: 사유)’를 적습니다 → 머리글 경고를 확인하세요."),
            ("t", "시장·순위·수급·지수 API는 모의투자 도메인에서 지원되지 않는 경우가 많아 실전 도메인(prod)을 사용합니다. 주문은 전혀 하지 않습니다(조회 전용)."),
            ("h", "5. 문제 해결"),
            ("t", "“유효한 접근토큰이 없습니다” → 모두 새로 고침 한 번 더. 계속되면 [데이터] > [쿼리 및 연결] > T_Token 우클릭 > 새로 고침."),
            ("t", "“EGW00133” (토큰 1분당 1회) → 1분 후 다시. “EGW00201” (초당 건수 초과) → 자동 재시도, 계속되면 [설정] 종목별 수급 조회를 N으로."),
            ("t", "“Formula.Firewall” → [데이터] > [데이터 가져오기] > [쿼리 옵션] > 현재 통합 문서 > 개인정보 > '개인 정보 수준 무시' 선택."),
            ("t", "“kis_devlp.yaml에 필수 항목이 없습니다” → [설정]의 경로 확인, 파일에 my_app / my_sec / prod 항목이 있어야 합니다."),
            ("t", "통합문서를 다른 사람에게 보낼 때 → 숨김 시트 _sys의 토큰 표 내용을 지우고 보내세요(앱키·시크릿은 원래 포함되지 않음)."),
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
    def finish(self):
        for t in self.temp_rows:
            try:
                self.lo[t].ListRows(1).Delete()
            except Exception as e:
                self.say(f"임시 행 삭제 실패: {t} {e}")
        for s in ("_data", "_sys"):
            self.ws[s].Visible = XL_SHEET_HIDDEN
        self.xl.CalculateFull()
        self.ws["대시보드"].Activate()
        self.ws["대시보드"].Range("A1").Select()
        self.wb.Save()
        self.say(f"저장 완료: {self.out}")


def main():
    ap = argparse.ArgumentParser(description="KIS PM 일일 대시보드 생성")
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--cfg", default=DEFAULT_CFG, help="kis_devlp.yaml 경로")
    ap.add_argument("--empty", action="store_true", help="샘플 매매·관심종목 없이 생성")
    ap.add_argument("--visible", action="store_true")
    ap.add_argument("--checkpoint", help=argparse.SUPPRESS)
    ap.add_argument("--resume", help=argparse.SUPPRESS)
    ap.add_argument("--seed-token", help=argparse.SUPPRESS)
    a = ap.parse_args()
    if not os.path.exists(a.cfg):
        sys.exit(f"[오류] KIS 설정 파일이 없습니다: {a.cfg}  (--cfg 로 경로 지정)")
    Builder(a.out, a.cfg, sample=not a.empty, visible=a.visible, seed_token=a.seed_token).run(
        resume=a.resume, checkpoint=a.checkpoint)


if __name__ == "__main__":
    main()
