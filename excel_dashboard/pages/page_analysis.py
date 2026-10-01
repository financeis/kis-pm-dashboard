# -*- coding: utf-8 -*-
"""종목분석 페이지 빌더 (spec R17·R20, 2026-10-01 T19).

시트 `종목분석` 한 장에 종목 하나를 깊이 보는 화면을 그린다. 맨 위(틀 고정 9행)에 종목코드 입력 칸·[조회] 버튼·최근 조회/상태·종목 정보 카드가 있고,
아래로 구역 1~8(숫자 표)이, 오른쪽 끝(AA열~)에 구역 9·10(글자가 긴 표: 뉴스·이벤트)이 놓인다.
표는 모두 [조회] 때 Power Query(쿼리 T_A_* → 표 tblA_*)가 채우는 표라서, 이 모듈은 값을 쓰지 않고 서식·조건부 서식·상태 줄·차트·이름만 입힌다.

인터페이스 (배선 작업 T23이 지켜야 할 것)
  SHEET   시트 이름. 빌더의 SHEETS에 넣고 builder.ws[SHEET]로 접근할 수 있어야 한다.
  LOADS   (쿼리, 표 이름, 머리글 왼쪽 위 셀) 13개 — 이 위치에 표를 적재한다. 표마다 아래쪽에 최대 행 수만큼 빈 행이 남도록 배치했다
          (투자자 120·매물대 100·신용공매도 60·증권사 목록 45·뉴스 40·이벤트 30 …). 표 아래·옆의 빈 칸에 다른 것을 두지 말 것.
  prepare_tables(builder)  표 13개를 적재한 **직후, 첫 새로 고침 전에** 부른다. 표의 새로 고침 방식을 덮어쓰기(QueryTable.RefreshStyle = 0)로 바꾼다.
          이 시트는 표들이 같은 열 아래위로 쌓여 있어서, 기본값(셀 삽입·삭제)으로 한 번이라도 새로 고치면 — 특히 0행 결과(빈 표)에서 —
          Excel이 셀을 지우며 아래 표들을 위로 끌어올려 LOADS 위치가 어긋나고 뉴스·이벤트 표는 새로 고침 자체가 실패한다(실측).
          `xl_helpers.load_query`는 기본값 1(삽입·삭제)을 쓰므로 배선에서 반드시 이 함수(또는 같은 설정)를 먼저 불러야 한다. VBA도 되돌리지 말 것.
  build(builder)  표를 적재·새로 고친 뒤(빈 표여도 된다) 부른다. builder.wb / ws / lo / table_style / say / nav_links / HEADER_RIGHT / HEADER_STATUS만 쓴다.
          표가 없으면 경고만 남기고 그 구역의 서식·수식·차트를 건너뛴다(시트·구역 제목은 그대로 그린다). 표 머리글이 LOADS 셀과 다른 곳에 있으면
          (위 설정을 빠뜨려 밀린 경우) 경고를 남긴다. 여러 번 불러도 겹치지 않는다(조건부 서식·도형·링크·차트 이름을 지우고 다시 만든다).
  통합문서 범위 이름: 분석코드(입력 칸 D6, 텍스트 서식) · 분석_최근조회(M6, 날짜+시각 서식) · 분석_상태(P6, 글자) — VBA RunAnalysis가 읽고 쓴다.
          대회코드목록 = tblContest[종목코드] (입력 칸 드롭다운 원천 — 목록 밖 코드도 입력할 수 있다. tblContest가 있어야 만들어지며, 없으면 드롭다운만 건너뜀).
  시트 범위 이름 `분석차트_…` 13개(숨김): 투자자·신용공매도·분기 실적 차트가 읽는 동적 범위. 숨김 열 BD:BF에 매물대·목표가 추이 차트 도우미 수식. 지우지 말 것.
  버튼: 도형 `btn_RunAnalysis` (OnAction = RunAnalysis).

설계 이유
  - 구역 1~8은 같은 열 눈금(B~Y)을 쓰는 아래위 쌓기라서 표마다 첫 열(종목코드)이 B열에 맞는다. 열 너비는 표 정의에서 열마다 최댓값으로 정한다.
    같은 줄에 나란히 놓은 표(3-a/3-b, 6-a/6-b)와 표 옆 차트는 남는 폭을 쓴다.
  - 뉴스 제목·이벤트 상세는 60자 넘게 길어 숫자 눈금에 넣으면 열이 망가지므로 별도 구역(AA 이후)에 둔다. 이벤트의 상세 열(5번째)과 뉴스의 제목 열(3번째)이
    같은 열(AE)에 오도록 뉴스를 두 칸 오른쪽에서 시작한다.
  - 10행 아래는 모두 높이 16으로 같게 해서(머리글도 한 줄) 왼쪽 표와 오른쪽 표의 행이 어긋나 보이지 않게 한다. 긴 머리글은 칸에 맞춰 줄여 보인다.
  - 열 서식·조건부 서식은 표 데이터 범위가 아니라 표 아래 예약 칸까지 시트 칸에 직접 입힌다 — 빈 표(0행)에서도 되고, 덮어쓰기로 늘어난 행도 같은 모양이다.
  - 차트는 빈 표에서도 만들어지고 표 크기가 바뀌어도 따라가야 한다. 표 열 범위를 직접 걸면 0행에서는 만들 수 없고, 1행에서 만든 계열은 표가 늘어도
    따라 늘지 않았다(실측). 그래서 행 수가 바뀌는 표는 데이터 칸 수를 세는 동적 범위 이름으로, 행 수가 정해진 표는 그 칸 범위로 건다.
  - 가로 막대 차트(구역 3-b)의 '현재가 구간' 색은 숨김 열의 도우미 수식 두 줄(그 구간만 값, 나머지 0)로 나눠 새로 고침 뒤에도 따라간다.
"""
from __future__ import annotations

import contextlib
import dataclasses
from typing import Any, Optional

import pywintypes

from xl_helpers import (
    C, NF, XL_BAR_CLUSTERED, XL_CENTER, XL_COLUMN_CLUSTERED, XL_EDGE_BOTTOM, XL_EDGE_LEFT, XL_LEFT, XL_LEGEND_TOP, XL_LINE,
    XL_LINE_MARKERS, XL_RIGHT, XL_SECONDARY, add_button, add_names, border, cf_databar, cf_expr, chart, fill,
    freeze, put, rgb, section, series, set_nf, sheet_setup, style, style_axes, title_bar, validation_list,
)

__all__ = ["SHEET", "LOADS", "prepare_tables", "build"]

SHEET = "종목분석"

# ------------------------------------------------------------------------------------------------ 이름·상수
NAME_CODE = "분석코드"
NAME_TIME = "분석_최근조회"
NAME_STATUS = "분석_상태"
NAME_CONTEST_LIST = "대회코드목록"
CONTEST_LIST_REF = "=tblContest[종목코드]"
DEFAULT_CODE = "005930"
MACRO_NAME = "RunAnalysis"
# 이름이 가리키는 칸(분석_최근조회 M6:N6, 분석_상태 P6:W6은 병합 — VBA는 첫 칸에 쓴다)
INPUT_CELL, TIME_CELL, STATUS_CELL = "D6", "M6", "P6"

XL_OVERWRITE_CELLS = 0   # QueryTable.RefreshStyle: 셀을 끼워 넣거나 지우지 않고 덮어씀

BAND_LAST_COL = "AH"     # 제목 띠·메뉴 줄 오른쪽 끝(뉴스·이벤트 구역 끝)
RIGHT_TEXT_COL = "Y"     # 머리글 오른쪽 문구(기준 시각·상태)를 붙일 열 = 숫자 구역 오른쪽 끝(첫 화면에서 보이게)
FROZEN_ROWS = 9          # 틀 고정: 제목·메뉴·입력·정보 카드
BODY_FIRST_ROW, BODY_LAST_ROW, BODY_ROW_HEIGHT = 10, 445, 16
HELPER_COLS = ("BD", "BE", "BF")   # 숨김 열: 차트 도우미

NF_MDHM = "mm-dd hh:mm"
NF_INT = "#,##0"
NF_WON = '#,##0"원"'
NF_SIGNED = "[Color10]+#,##0;[Color11]-#,##0;0"           # 부호 색(상승 빨강·하락 파랑)
NF_NEG_BLUE = "#,##0;[Color11]-#,##0;0"                   # 음수만 파랑(실적 금액)
NF_DEC1 = "#,##0.0"
NF_DEC1_SIGNED = "[Color10]+#,##0.0;[Color11]-#,##0.0;0.0"
NF_PCT1 = "0.0%"
NF_PCT2 = "0.00%"
NF_PCT1_SIGNED = "[Color10]+0.0%;[Color11]-0.0%;0.0%"
NF_PCT2_SIGNED = NF["pct_pl"]
NF_MULT = '0.0"배"'
NF_DDAY = '"D-"0;"D+"0;"D-day"'   # 양수 = 남은 날(D-3), 음수 = 지난 날(D+2), 0 = D-day (조회 시점 기준 저장값)

COLOR_BAR = "#9DB4D8"       # 매물대 막대(현재가 구간 외)
COLOR_FOREIGN, COLOR_INST, COLOR_PERSON = "#1F6FEB", "#E8890C", "#7A8699"


def _c(width: float, nf: Optional[str] = None, h: Optional[int] = None, **text_style: Any) -> dict:
    """열 서식 한 개: width(글자 수), nf(숫자 서식), h(가로 정렬), bold·color·size·italic."""
    return {"w": width, "nf": nf, "h": h, **text_style}


_CODE = _c(6.5, h=XL_CENTER, color=C["muted"], size=8)
_STATUS = _c(7, h=XL_LEFT, color=C["muted"], size=8)
_STAMP = _c(11, nf=NF_MDHM, h=XL_CENTER, color=C["muted"], size=8)
_DATE = _c(10.5, nf=NF["date"], h=XL_CENTER)


@dataclasses.dataclass(frozen=True)
class _Tbl:
    """페이지가 알고 있는 표 하나: 쿼리·표 이름·머리글 왼쪽 위 셀·예약 행 수·열별 서식(쿼리 열 순서 그대로)."""

    query: str
    name: str
    cell: str
    rows: int
    cols: dict


# 열 순서 = T_A_*.pq의 Types 순서. 너비는 같은 열 번호에 놓인 모든 표 중 최댓값이 시트 열 너비가 된다.
INFO = _Tbl("T_A_Info", "tblA_Info", "AA92", 2, {
    "종목코드": _CODE, "종목명": _c(10.5, bold=True), "시장": _c(7.5, h=XL_CENTER), "대회편입": _c(14, h=XL_CENTER),
    "대회비고": _c(66), "NICS 대분류": _c(11), "NICS 업종": _c(11), "NICS 세부": _c(11),
    "대테마": _c(12), "세부테마": _c(12), "분류출처": _c(8, h=XL_CENTER), "KRX업종": _c(12), "KSIC 세분류": _c(24),
    "주요상품": _c(30), "기업집단": _c(10), "현재가": _c(9, NF_INT), "전일대비": _c(9, NF_SIGNED), "등락률": _c(8, NF_PCT2_SIGNED),
    "거래대금억": _c(9, NF_INT), "시가총액억": _c(11, NF_INT), "PER": _c(7, NF_DEC1), "PBR": _c(7, "0.00"),
    "외국인소진율": _c(9, NF_PCT1), "고52주": _c(9, NF_INT), "저52주": _c(9, NF_INT), "유의": _c(9),
    "상태": _c(12, h=XL_LEFT, color=C["muted"], size=8), "조회시각": _STAMP})
INVESTOR = _Tbl("T_A_Investor", "tblA_Investor", "B34", 120, {
    "종목코드": _CODE, "일자": _DATE, "종가": _c(8.5, NF_INT, bold=True), "등락률": _c(8, NF_PCT2_SIGNED),
    "거래대금억": _c(9, NF_INT), "외국인억": _c(8.5, NF_SIGNED), "기관계억": _c(8.5, NF_SIGNED), "연기금억": _c(8, NF_SIGNED),
    "투신억": _c(8, NF_SIGNED), "사모억": _c(8, NF_SIGNED), "금융투자억": _c(9, NF_SIGNED), "보험억": _c(8, NF_SIGNED),
    "은행억": _c(8, NF_SIGNED), "기타법인억": _c(9, NF_SIGNED), "개인억": _c(8.5, NF_SIGNED),
    "외국인누적억": _c(10.5, NF_SIGNED, bold=True), "기관계누적억": _c(10.5, NF_SIGNED, bold=True),
    "개인누적억": _c(10, NF_SIGNED, bold=True), "상태": _STATUS, "조회시각": _STAMP})
TRADESIZE = _Tbl("T_A_TradeSize", "tblA_TradeSize", "B159", 8, {
    "종목코드": _CODE, "기준일": _DATE, "순서": _c(5, "0", XL_CENTER, color=C["muted"]), "구간": _c(11, h=XL_CENTER, bold=True),
    "매수거래량": _c(10, NF_INT), "매도거래량": _c(10, NF_INT), "순매수거래량": _c(10, NF_SIGNED, bold=True),
    "매수비율": _c(8, NF_PCT1), "매도비율": _c(8, NF_PCT1), "순매수비율": _c(9, NF_PCT1_SIGNED, bold=True),
    "매수건수": _c(9, NF_INT), "매도건수": _c(9, NF_INT), "순매수건수": _c(9, NF_SIGNED), "평균가": _c(9, NF_INT),
    "상태": _STATUS, "조회시각": _STAMP})
PBAR = _Tbl("T_A_PbarToday", "tblA_PbarToday", "B176", 100, {
    "종목코드": _CODE, "기준일": _DATE, "가격": _c(9, NF_INT, bold=True), "거래량": _c(10, NF_INT), "비중": _c(8, NF_PCT2),
    "순위": _c(5.5, "0", XL_CENTER, color=C["muted"]), "현재가구간": _c(9, h=XL_CENTER, bold=True, color=C["accent"]),
    "상태": _STATUS, "조회시각": _STAMP})
PROFILE = _Tbl("T_A_Profile", "tblA_Profile", "L176", 20, {
    "종목코드": _CODE, "구간": _c(5.5, "0", XL_CENTER, color=C["muted"]), "구간하단": _c(9, NF_INT), "구간상단": _c(9, NF_INT),
    "가격대": _c(13.5, h=XL_CENTER, size=9), "거래량": _c(11, NF_INT), "비중": _c(8, NF_PCT1, bold=True),
    "현재가구간": _c(9, h=XL_CENTER, bold=True, color=C["accent"]), "현재가": _c(9, NF_INT), "시작일": _DATE, "종료일": _DATE,
    "세션수": _c(6, "0", XL_CENTER, color=C["muted"]), "상태": _STATUS, "조회시각": _STAMP})
ESTIMATE = _Tbl("T_A_Estimate", "tblA_Estimate", "B282", 5, {
    "종목코드": _CODE, "기준일": _DATE, "순서": _c(5, "0", XL_CENTER, color=C["muted"]), "입력시점": _c(8, h=XL_CENTER, bold=True),
    "외국인수량": _c(10, NF_SIGNED), "기관수량": _c(10, NF_SIGNED), "합산수량": _c(10, NF_SIGNED, bold=True),
    "외국인≈억": _c(9, NF_DEC1_SIGNED), "기관≈억": _c(9, NF_DEC1_SIGNED), "합산≈억": _c(9, NF_DEC1_SIGNED, bold=True),
    "현재가": _c(9, NF_INT), "상태": _STATUS, "조회시각": _STAMP})
CSL = _Tbl("T_A_CSL", "tblA_CSL", "B299", 60, {
    "종목코드": _CODE, "일자": _DATE, "종가": _c(8.5, NF_INT), "신용잔고율": _c(9, NF_PCT2), "신용잔고주수": _c(11, NF_INT),
    "공매도수량": _c(10, NF_INT), "공매도거래대금억": _c(10, NF_DEC1), "공매도비중": _c(9, NF_PCT1, bold=True),
    "대차잔고주수": _c(11, NF_INT), "대차잔고금액억": _c(11, NF_INT), "대차증감주수": _c(10, NF_SIGNED),
    "상태": _STATUS, "조회시각": _STAMP})
TARGETS = _Tbl("T_A_Targets", "tblA_Targets", "B364", 45, {
    "종목코드": _CODE, "증권사": _c(12, bold=True), "일자": _DATE, "투자의견": _c(11, h=XL_CENTER),
    "직전투자의견": _c(11, h=XL_CENTER, color=C["muted"]), "목표가": _c(9, NF_INT, bold=True), "목표가일자": _DATE,
    "직전목표가": _c(9, NF_INT, color=C["muted"]), "목표가변화율": _c(9, NF_PCT1_SIGNED), "상태": _STATUS, "조회시각": _STAMP})
TARGET_TREND = _Tbl("T_A_TargetTrend", "tblA_TargetTrend", "N364", 13, {
    "종목코드": _CODE, "기준일": _DATE, "구분": _c(6, h=XL_CENTER), "목표가평균": _c(10, NF_INT, bold=True),
    "증권사수": _c(7, "0", XL_CENTER), "최고목표가": _c(10, NF_INT), "최저목표가": _c(10, NF_INT), "상태": _STATUS, "조회시각": _STAMP})
KISEST = _Tbl("T_A_KisEst", "tblA_KisEst", "B414", 8, {
    "종목코드": _CODE, "연도": _c(9, h=XL_CENTER, bold=True), "결산년월": _c(9, h=XL_CENTER, color=C["muted"]),
    "추정여부": _c(6, h=XL_CENTER, bold=True, color=C["warn"]), "매출억": _c(10, NF_NEG_BLUE),
    "매출증가율": _c(8.5, NF_PCT1_SIGNED), "영업이익억": _c(10, NF_NEG_BLUE), "영업이익증가율": _c(9, NF_PCT1_SIGNED),
    "순이익억": _c(10, NF_NEG_BLUE), "순이익증가율": _c(9, NF_PCT1_SIGNED), "EBITDA억": _c(10, NF_NEG_BLUE),
    "EPS": _c(8, NF_NEG_BLUE), "EPS증가율": _c(8.5, NF_PCT1_SIGNED), "PER": _c(7.5, NF_MULT), "KIS_PER": _c(7.5, NF_MULT),
    "EV_EBITDA": _c(8.5, "0.0"), "ROE": _c(7.5, NF_PCT1), "부채비율": _c(8, NF_PCT1), "이자보상배율": _c(8.5, "0.0"),
    "추정일": _DATE, "KIS의견": _c(8, h=XL_CENTER), "현재가": _c(9, NF_INT), "상태": _STATUS, "조회시각": _STAMP})
FIN = _Tbl("T_A_Fin", "tblA_Fin", "B427", 10, {
    "종목코드": _CODE, "결산년월": _c(9, h=XL_CENTER, color=C["muted"]), "분기": _c(8, h=XL_CENTER, bold=True),
    "매출억": _c(10, NF_NEG_BLUE), "영업이익억": _c(10, NF_NEG_BLUE), "순이익억": _c(10, NF_NEG_BLUE),
    "영업이익률": _c(8.5, NF_PCT1), "매출YoY": _c(8.5, NF_PCT1_SIGNED), "영업이익YoY": _c(9.5, NF_PCT1_SIGNED),
    "순이익YoY": _c(9, NF_PCT1_SIGNED), "상태": _STATUS, "조회시각": _STAMP})
NEWS = _Tbl("T_A_News", "tblA_News", "AC12", 40, {
    "종목코드": _c(7.5, h=XL_CENTER, color=C["muted"], size=8), "일시": _c(14, nf=NF_MDHM, h=XL_CENTER),
    "제목": _c(66), "출처": _c(11, color=C["muted"], size=8), "상태": _c(11, h=XL_LEFT, color=C["muted"], size=8), "조회시각": _STAMP})
EVENTS = _Tbl("T_A_Events", "tblA_Events", "AA57", 30, {
    "종목코드": _CODE, "일자": _DATE, "D-day": _c(7.5, NF_DDAY, XL_CENTER, bold=True), "종류": _c(14, h=XL_CENTER),
    "상세": _c(66), "상태": _c(11, h=XL_LEFT, color=C["muted"], size=8), "조회시각": _STAMP})

TABLES = [INFO, INVESTOR, TRADESIZE, PBAR, PROFILE, ESTIMATE, CSL, TARGETS, TARGET_TREND, KISEST, FIN, NEWS, EVENTS]
LOADS = [(t.query, t.name, t.cell) for t in TABLES]

# 구역 제목·상태 줄·차트 위치 (행은 머리글 행 기준으로 위쪽에 둔다)
SECTION_LINKS = [   # (바로가기 글자, 이동할 셀)
    ("1 가격·수급", "B10"), ("2 체결금액", "B157"), ("3 매물대", "B174"), ("4 추정가집계", "B280"), ("5 신용·공매도", "B297"),
    ("6 목표주가", "B362"), ("7 KIS 추정", "B412"), ("8 분기 실적", "B425"), ("9 뉴스·공시", "AC10"), ("10 이벤트", "AA55"),
]
CHART_PRICE = "C12:U32"
CHART_TRADESIZE = "R159:Y171"
CHART_PROFILE = "L198:Y223"
CHART_ESTIMATE = "O282:Y295"
CHART_CSL = "P299:Y322"
CHART_TREND = "N379:Y399"
CHART_FIN = "O427:Y440"
HELPER_PROFILE_ROW = 177          # 숨김 열 BD·BE: 매물대 도우미(20행, 표 데이터 첫 행과 같은 행) — BD175 = 차트 제목 글자
HELPER_TREND_ROW = 365            # 숨김 열 BF: 목표가 추이 차트의 현재가 선(13행)


# ------------------------------------------------------------------------------------------------ 셀 주소 도우미
def _col_letter(n: int) -> str:
    """열 번호(1부터) → 문자. 예: 27 → 'AA'."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _col_number(letters: str) -> int:
    """열 문자 → 번호. 예: 'AA' → 27."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _split_cell(cell: str) -> tuple[str, int]:
    """'AC12' → ('AC', 12)."""
    i = 0
    while cell[i].isalpha():
        i += 1
    return cell[:i], int(cell[i:])


def _absolute(cell: str) -> str:
    """'D6' → '$D$6'."""
    letters, row = _split_cell(cell)
    return f"${letters}${row}"


def _tbl_left(t: _Tbl) -> int:
    """표 머리글 왼쪽 열 번호(LOADS 셀 기준)."""
    return _col_number(_split_cell(t.cell)[0])


def _tbl_header_row(t: _Tbl) -> int:
    """표 머리글 행 번호(LOADS 셀 기준)."""
    return _split_cell(t.cell)[1]


def _lane_widths() -> dict[int, float]:
    """열 번호 → 너비: 같은 열 번호에 놓인 모든 표의 열 너비 중 최댓값(구역이 같은 열 눈금을 나눠 쓰므로)."""
    widths: dict[int, float] = {}
    for t in TABLES:
        left = _tbl_left(t)
        for k, spec in enumerate(t.cols.values()):
            widths[left + k] = max(widths.get(left + k, 0.0), float(spec["w"]))
    # 입력 칸(굵은 13pt 코드)과 긴 머리글(공매도거래대금억·영업이익증가율·순이익증가율·이자보상배율)이 줄어들지 않게 최소 너비
    for letter, w in (("D", 11.0), ("H", 12.5), ("I", 11.0), ("K", 10.5), ("T", 10.0)):
        widths[_col_number(letter)] = max(widths.get(_col_number(letter), 0.0), w)
    widths[_col_number("A")] = 1.5
    widths[_col_number("Z")] = 2.5      # 숫자 구역과 뉴스·이벤트 구역 사이 간격
    widths[_col_number("AI")] = max(widths.get(35, 0.0), 9.0)
    return widths


# ------------------------------------------------------------------------------------------------ 표 위치
@dataclasses.dataclass
class _Geo:
    """표의 시트 위치: 머리글 행, 왼쪽·오른쪽 열 번호, 열 이름 → 열 번호."""

    header_row: int
    left: int
    right: int
    cols: dict[str, int]

    @property
    def first_row(self) -> int:
        return self.header_row + 1


def _geometry(lo: Any) -> _Geo:
    """살아 있는 표(ListObject)의 위치를 읽는다. 머리글은 0행 표에도 있으므로 빈 표에서도 된다."""
    hdr = lo.HeaderRowRange
    left = int(hdr.Column)
    cols = {str(c.Name): int(c.Range.Column) for c in lo.ListColumns}
    return _Geo(int(hdr.Row), left, left + len(cols) - 1, cols)


def _fallback_geo(t: _Tbl) -> _Geo:
    """표가 없을 때 구역 제목만 그리기 위한 위치(LOADS 셀과 정의된 열 순서 기준)."""
    left = _tbl_left(t)
    return _Geo(_tbl_header_row(t), left, left + len(t.cols) - 1, {n: left + k for k, n in enumerate(t.cols)})


def _sheet(builder: Any) -> Any:
    """builder.ws[SHEET]를 돌려준다. 시트가 없으면 경고만 남기고 None."""
    try:
        return builder.ws[SHEET]
    except KeyError:
        builder.say(f"[{SHEET}] 경고: 시트가 없어 페이지를 만들지 않습니다(빌더 SHEETS에 '{SHEET}' 추가 필요)")
        return None


def _table(builder: Any, name: str) -> Any:
    """builder.lo[표 이름]을 돌려준다. 없으면 경고만 남기고 None(그 구역은 건너뜀)."""
    try:
        return builder.lo[name]
    except KeyError:
        builder.say(f"[{SHEET}] 경고: 표 {name}이(가) 없어 해당 구역의 서식·수식·차트를 건너뜁니다(LOADS 위치에 적재한 뒤 호출해야 합니다)")
        return None


def _com_ok(builder: Any, what: str, fn: Any) -> bool:
    """COM 호출 하나를 실행하고 실패하면 경고만 남긴다(그 부분만 건너뜀). 성공 여부를 돌려준다."""
    try:
        fn()
        return True
    except pywintypes.com_error as e:
        detail = e.excepinfo[2] if getattr(e, "excepinfo", None) and len(e.excepinfo) > 2 and e.excepinfo[2] else e.strerror
        builder.say(f"[{SHEET}] 경고: {what} — 건너뜀 ({detail})")
        return False


def _overwrite_on_refresh(lo: Any) -> None:
    """쿼리 새로 고침이 표 열 안의 셀을 끼워 넣거나 지우지 않고 덮어쓰게 한다(xlOverwriteCells).

    기본값(셀 삽입·삭제)에서는 표가 0·1행에서 늘어날 때 Excel이 조건부 서식 규칙을 새 행에 이어 주지 않았다(T20 실측).
    이 시트는 표마다 아래에 예약 행이 비어 있어 덮어써도 안전하다. 정적 표(쿼리 아님)는 QueryTable이 없어 건너뛴다.
    """
    try:
        lo.QueryTable.RefreshStyle = XL_OVERWRITE_CELLS
    except pywintypes.com_error:
        pass


def prepare_tables(builder: Any) -> None:
    """LOADS의 표 13개(있는 것만)를 새로 고침 방식 `덮어쓰기`로 바꾼다. 표 적재 직후·첫 새로 고침 전에 부른다.

    Args:
        builder: builder.lo[표이름]으로 ListObject를 돌려주는 객체(Builder 또는 하네스 대역).

    Example:
        for query, table, cell in page_analysis.LOADS:
            ...load_query(...)           # 표 적재
        page_analysis.prepare_tables(builder)
        ...refresh_all()                  # 그 뒤에 새로 고침
    """
    for t in TABLES:
        try:
            lo = builder.lo[t.name]
        except KeyError:
            continue
        _overwrite_on_refresh(lo)


def _check_position(builder: Any, lo: Any, t: _Tbl) -> None:
    """표 머리글이 LOADS 셀에 있는지 확인한다(다르면 구역 제목·차트가 표와 어긋남)."""
    try:
        hdr = lo.HeaderRowRange
        actual = f"{_col_letter(int(hdr.Column))}{int(hdr.Row)}"
    except pywintypes.com_error:
        return
    if actual != t.cell:
        builder.say(f"[{SHEET}] 경고: {t.name}의 머리글이 {actual}에 있습니다(LOADS 위치 {t.cell}) — 표를 적재한 뒤 첫 새로 고침 전에 "
                    f"prepare_tables(RefreshStyle = 0)를 부르지 않아 표가 밀린 것으로 보입니다. 구역 제목·차트가 어긋납니다")


# ------------------------------------------------------------------------------------------------ 수식 조각
def _status_formula(t: str, key: str, ok_text: str, *, numeric_key: bool = True, none_hint: str = "") -> str:
    """구역 상태 줄 수식. 우선순위: 이전 데이터(갱신 실패) > 오류 > 데이터 행 없음 > 정상.

    `데이터 없음`·`추정 없음`은 정상 값이므로 빨강이 아니라 회색 `○`로 보이고, 일부 행만 그 상태인 표(추정가집계의 아직 입력되지 않은
    시점)는 데이터가 한 행이라도 있으면 정상으로 본다.

    Args:
        t: 표 이름(tblA_…).
        key: 데이터 행 수를 세는 열 이름(값이 있으면 데이터 행으로 침). numeric_key=False면 글자 열(COUNTA).
        ok_text: 정상일 때 `● 정상` 뒤에 붙일 수식 조각(`&`로 이어 쓸 글자 수식 한 덩어리, 없으면 "").
        none_hint: 데이터 행이 없을 때(`데이터 없음`·`추정 없음`) 덧붙일 설명.
    """
    n = f"COUNT({t}[{key}])" if numeric_key else f"COUNTA({t}[{key}])"
    ok = f'&{ok_text}' if ok_text else ""
    hint = f'&"{none_hint}"' if none_hint else ""
    return (f'=IFERROR(LET(z_e,COUNTIF({t}[상태],"오류*"),z_p,COUNTIF({t}[상태],"이전*"),z_n,{n},'
            f'z_x,COUNTIF({t}[상태],"데이터 없음")+COUNTIF({t}[상태],"추정 없음"),'
            f'IF(z_p>0,"▲ "&INDEX({t}[상태],MATCH("이전*",{t}[상태],0))&" — 직전 값을 표시 중입니다",'
            f'IF(z_e>0,"▲ "&INDEX({t}[상태],MATCH("오류*",{t}[상태],0)),'
            f'IF(z_n=0,IF(z_x>0,"○ "&IF(COUNTIF({t}[상태],"추정 없음")>0,"추정 없음","데이터 없음"){hint},'
            f'"○ 데이터 없음 — 종목코드를 입력하고 [조회]를 누르세요"),'
            f'"● 정상"{ok})))),"")')


def _note_formula(t: str, date_col: str, *, today_tag: bool = False, fmt: str = "yyyy-mm-dd", label: str = "기준",
                  stamp: bool = True) -> str:
    """구역 제목 줄 오른쪽 안내: 당일 구역은 `당일 mm-dd 기준 · 조회 hh:mm`, 그 밖은 `<라벨> <최근 날짜> · 조회 mm-dd hh:mm`. 값이 없으면 빈 글자."""
    if today_tag:
        return (f'=IFERROR(IF(COUNT({t}[{date_col}])=0,"당일 데이터",'
                f'"당일 "&TEXT(MAX({t}[{date_col}]),"mm-dd")&" 기준 · 조회 "&TEXT(MAX({t}[조회시각]),"hh:mm")),"")')
    tail = f'&" · 조회 "&TEXT(MAX({t}[조회시각]),"mm-dd hh:mm")&" · [조회]로 갱신"' if stamp else ""
    return f'=IFERROR(IF(COUNT({t}[{date_col}])=0,"","{label} "&TEXT(MAX({t}[{date_col}]),"{fmt}"){tail}),"")'


def _ok(*parts: str) -> str:
    """상태 줄 정상 문구 조각들을 ` · `로 잇는 수식 조각(글자 리터럴은 따옴표로, 수식은 그대로 넘긴다)."""
    return "&".join(parts)


def _lit(s: str) -> str:
    """글자를 수식 글자 리터럴로(따옴표 이스케이프)."""
    return '"' + s.replace('"', '""') + '"'


def _price_gap(avg_expr: str, have_info: bool) -> str:
    """평균 목표가 문구 `507,045원 (현재가 대비 +83.7%)` 수식 조각. 현재가가 없거나 tblA_Info가 없으면 괴리율은 뺀다."""
    if not have_info:
        return f'IFERROR(TEXT({avg_expr},"#,##0")&"원","-")'
    return (f'IFERROR(LET(z_a,{avg_expr},z_p,N(INDEX(tblA_Info[현재가],1)),TEXT(z_a,"#,##0")&"원"'
            f'&IF(z_p>0," (현재가 대비 "&TEXT(z_a/z_p-1,"+0.0%;-0.0%")&")","")),"-")')


# ------------------------------------------------------------------------------------------------ 상단 띠·입력·정보 카드
def _header_band(builder: Any, ws: Any) -> None:
    """제목 띠(2~3행) + 메뉴 줄(4행) + 구역 바로가기 줄(5행)."""
    subtitle = ("종목 하나를 골라 [조회] — 가격·수급·매물대·추정·목표가·실적·뉴스를 한 화면에서 (그 종목만 KIS에서 조회, 약 30초) | "
                "'당일' 표시 구역은 장중 값이 계속 변합니다")
    span = f"B2:{BAND_LAST_COL}3"
    right = getattr(builder, "HEADER_RIGHT", None)
    status = getattr(builder, "HEADER_STATUS", None)
    try:
        title_bar(ws, "STOCK ANALYSIS", subtitle, span, right_formula=right, right_cell=f"{RIGHT_TEXT_COL}2",
                  right2_formula=status, right2_cell=f"{RIGHT_TEXT_COL}3")
    except pywintypes.com_error:
        builder.say(f"[{SHEET}] 경고: 머리글 기준 시각·상태 수식을 넣지 못했습니다(참조 표 없음) — 제목만 표시")
        title_bar(ws, "STOCK ANALYSIS", subtitle, span)
    fill(ws.Range(f"B4:{BAND_LAST_COL}4"), "navy2")
    builder.nav_links(ws, 4)
    # 구역 바로가기: 이 시트 안의 이동 링크(2칸 간격으로 D열부터)
    put(ws, "B5", "구역 ▸", size=9, color=C["muted"], bold=True)
    col = _col_number("D")
    for text, target in SECTION_LINKS:
        cell = ws.Cells(5, col)
        cell.Value = text
        ws.Hyperlinks.Add(cell, "", f"'{SHEET}'!{target}", "", text)   # 위치 인수: Anchor, Address, SubAddress, ScreenTip, TextToDisplay
        style(cell, color=C["accent"], size=9, bold=True)
        cell.Font.Underline = False
        col += 2


def _norm_code_expr() -> str:
    """입력 칸 값을 쿼리와 같은 규칙(앞뒤 공백 제거·대문자·숫자만 6자리 미만이면 앞 0 채움)으로 정규화하는 수식 조각."""
    return (f'LET(z_c,UPPER(TRIM({NAME_CODE}&"")),'
            f'IF(AND(LEN(z_c)>0,LEN(z_c)<6,ISNUMBER(--z_c)),RIGHT("000000"&z_c,6),z_c))')


def _define_names(builder: Any) -> bool:
    """입력·결과 칸의 이름 정의와 대회 코드 목록 이름. 반환: 대회코드목록을 만들었는지."""
    sheet_ref = f"'{SHEET}'!"
    add_names(builder.wb, {
        NAME_CODE: f"={sheet_ref}{_absolute(INPUT_CELL)}",
        NAME_TIME: f"={sheet_ref}{_absolute(TIME_CELL)}",
        NAME_STATUS: f"={sheet_ref}{_absolute(STATUS_CELL)}",
    })
    try:
        add_names(builder.wb, {NAME_CONTEST_LIST: CONTEST_LIST_REF})
        return True
    except RuntimeError:
        builder.say(f"[{SHEET}] 경고: 이름 {NAME_CONTEST_LIST} = {CONTEST_LIST_REF}을 만들지 못했습니다(tblContest 없음) — 종목코드 드롭다운 없이 둡니다")
        return False


def _input_row(builder: Any, ws: Any, have_info: bool, have_list: bool) -> None:
    """6행: 종목코드 입력 칸(텍스트 서식·선행 0 유지)·[조회] 버튼·안내·최근 조회 시각·상태."""
    ws.Range("B6:C6").Merge()
    put(ws, "B6", "종목코드 ▶", bold=True, size=11, color=C["navy"], h=XL_RIGHT)
    sel = ws.Range(INPUT_CELL)
    set_nf(sel, NF["text"])
    if sel.Value is None or str(sel.Value).strip() == "":
        sel.Value = DEFAULT_CODE
    style(sel, bold=True, size=13, fill=C["input"], h=XL_CENTER)
    border(sel, "warn")
    with contextlib.suppress(pywintypes.com_error):
        sel.Errors(3).Ignore = True   # 숫자가 텍스트로 저장됨(초록 삼각형) 표시 끄기 — 선행 0을 지키려는 의도적 서식
    if have_list:
        _com_ok(builder, "종목코드 드롭다운", lambda: validation_list(sel, f"={NAME_CONTEST_LIST}", show_error=False))
        with contextlib.suppress(pywintypes.com_error):
            sel.Validation.ShowInput = False   # 선택할 때마다 말풍선이 정보 카드를 가리지 않게
    _com_ok(builder, "[조회] 버튼", lambda: add_button(ws, "E6:F6", "조회", MACRO_NAME, size=11))

    ws.Range("G6:K6").Merge()
    hint = "코드 입력 → [조회] (약 30초)"
    if have_info:
        # 우선순위: 입력 없음 > 조회 결과 오류·직전 값 표시 중 > 입력 코드와 표시 중인 종목이 다름(대회 명단에 있으면 이름도) > 평소 안내
        name_of = (f'IFERROR(XLOOKUP(z_in,tblContest[종목코드],tblContest[종목명],""),"")' if have_list else '""')
        formula = (f'=IFERROR(LET(z_in,{_norm_code_expr()},z_k,INDEX(tblA_Info[종목코드],1)&"",z_s,INDEX(tblA_Info[상태],1)&"",'
                   f'z_nm,{name_of},'
                   f'IF(z_in="","▲ 종목코드를 입력하세요",IF(OR(LEFT(z_s,2)="오류",LEFT(z_s,2)="이전"),"▲ "&z_s,'
                   f'IF(AND(z_k<>"",z_in<>z_k),"▲ 입력 "&z_in&IF(z_nm="",""," "&z_nm)&" — [조회]를 누르세요 (표시 중: "&z_k&")","{hint}")))),"{hint}")')
        if not _com_ok(builder, "입력 안내 수식", lambda: put(ws, "G6", formula=formula, size=9, color=C["muted"], h=XL_LEFT, indent=1,
                                                               wrap=True)):
            put(ws, "G6", hint, size=9, color=C["muted"], h=XL_LEFT, indent=1)
    else:
        put(ws, "G6", hint, size=9, color=C["muted"], h=XL_LEFT, indent=1)
    cf_expr(ws.Range("G6"), '=LEFT($G$6,1)="▲"', font=C["up"], bold=True)

    put(ws, "L6", "최근 조회", size=9, color=C["muted"], h=XL_RIGHT)
    ws.Range("M6:N6").Merge()
    put(ws, TIME_CELL, nf=NF["dt"], bold=True, size=10, color=C["navy"], h=XL_LEFT)
    put(ws, "O6", "상태", size=9, color=C["muted"], h=XL_RIGHT)
    ws.Range("P6:W6").Merge()
    if ws.Range(STATUS_CELL).Value is None:
        ws.Range(STATUS_CELL).Value = "조회 전"
    put(ws, STATUS_CELL, bold=True, size=10, color=C["muted"], h=XL_LEFT)
    status, ref = ws.Range(STATUS_CELL), _absolute(STATUS_CELL)
    cf_expr(status, f'=LEFT({ref},2)="실패"', font=C["up"], bold=True)
    cf_expr(status, f'=LEFT({ref},2)="일부"', font=C["warn"], bold=True)
    cf_expr(status, f'={ref}="정상"', font=C["good"], bold=True)
    bottom = ws.Range("B6:Y6")
    border(bottom, "line", edges=(XL_EDGE_BOTTOM,))


def _card_specs() -> list[tuple]:
    """정보 카드 정의: (라벨 글자 또는 `=`수식, 라벨 범위, 값 범위, 값 수식, 숫자 서식, 글자 크기, 줄바꿈).

    값은 모두 tblA_Info 1행에서 읽고 IFERROR로 감싸 표가 비어도 빈 칸이 된다(0이 아니라 빈 글자).
    """
    t = "tblA_Info"

    def one(col: str) -> str:
        return f"INDEX({t}[{col}],1)"

    name = (f'=IFERROR(LET(z_n,{one("종목명")}&"",z_c,{one("종목코드")}&"",z_s,{one("상태")}&"",'
            f'IF(LEFT(z_s,2)="오류","⚠ 종목코드 확인 "&z_c,IF(z_n&z_c="","",z_n&"   "&z_c))),"")')
    name_label = (f'=IFERROR(IF(COUNT({t}[조회시각])=0,"종목","종목  ·  조회 "&TEXT(MAX({t}[조회시각]),"mm-dd hh:mm")),"종목")')
    market = (f'=IFERROR(LET(z_m,{one("시장")}&"",z_f,{one("대회편입")}&"",'
              f'TEXTJOIN("  ·  ",TRUE,z_m,IF(z_f="Y","대회 편입",IF(z_f="N","대회 외","")))),"")')
    nics = (f'=IFERROR(LET(z_a,{one("NICS 대분류")}&"",z_b,{one("NICS 업종")}&"",z_c,{one("NICS 세부")}&"",z_s,{one("분류출처")}&"",'
            f'TEXTJOIN(" › ",TRUE,z_a,IF(z_b=z_a,"",z_b),IF(z_c=z_b,"",z_c))&IF(z_s="KRX","  (KRX 업종 대체)","")),"")')
    theme = (f'=IFERROR(LET(z_a,{one("대테마")}&"",z_b,{one("세부테마")}&"",z_k,{one("종목코드")}&"",'
             f'IF(z_a&z_b="",IF(z_k="","","-"),TEXTJOIN(" › ",TRUE,z_a,z_b))),"")')
    price = f'=IFERROR(IF({one("현재가")}="","",{one("현재가")}),"")'
    price_label = (f'=IFERROR(IF(OR({one("고52주")}="",{one("저52주")}=""),"현재가",'
                   f'"현재가  ·  52주 "&TEXT({one("저52주")},"#,##0")&"~"&TEXT({one("고52주")},"#,##0")),"현재가")')
    change = (f'=IFERROR(LET(z_c,{one("전일대비")},z_r,{one("등락률")},IF(OR(z_c="",z_r=""),"",'
              f'IF(z_c>0,"▲ ",IF(z_c<0,"▼ ",""))&TEXT(ABS(z_c),"#,##0")&"   ("&TEXT(z_r,"+0.00%;-0.00%;0.00%")&")")),"")')
    cap = (f'=IFERROR(LET(z_v,{one("시가총액억")},IF(z_v="","",TEXT(z_v,"#,##0")&"억"'
           f'&IF(z_v>=10000,"  ("&TEXT(z_v/10000,"#,##0.0")&"조)",""))),"")')
    turnover = f'=IFERROR(IF({one("거래대금억")}="","",TEXT({one("거래대금억")},"#,##0")&"억"),"")'
    per_pbr = (f'=IFERROR(LET(z_a,{one("PER")},z_b,{one("PBR")},z_k,{one("종목코드")}&"",IF(z_k="","",'
               f'IF(z_a="","-",TEXT(z_a,"0.0"))&" / "&IF(z_b="","-",TEXT(z_b,"0.00")))),"")')
    return [
        (name_label, "B7:E7", "B8:E8", name, None, 13, False),
        ("시장 · 대회", "F7:G7", "F8:G8", market, None, 10, False),
        ("NICS 분류", "H7:K7", "H8:K8", nics, None, 9, True),
        ("테마", "L7:N7", "L8:N8", theme, None, 9, True),
        (price_label, "O7:P7", "O8:P8", price, NF_WON, 13, False),
        ("전일대비  ·  등락률", "Q7:R7", "Q8:R8", change, None, 11, False),
        ("시가총액", "S7:U7", "S8:U8", cap, None, 10, False),
        ("거래대금", "V7:W7", "V8:W8", turnover, None, 10, False),
        ("PER / PBR", "X7:Y7", "X8:Y8", per_pbr, None, 10, False),
    ]


def _info_cards(builder: Any, ws: Any) -> None:
    """7~8행 정보 카드: 라벨(7행) + 값(8행). 표가 비어도 빈 칸으로 보인다."""
    for label, lrng, vrng, formula, nf, size, wrap in _card_specs():
        lab, val = ws.Range(lrng), ws.Range(vrng)
        if lab.Count > 1:
            lab.Merge()
        if val.Count > 1:
            val.Merge()
        lab_addr, first = lab.Cells(1, 1).Address, val.Cells(1, 1).Address
        if label.startswith("="):
            _com_ok(builder, f"정보 카드 라벨 {lab_addr}", lambda: put(ws, lab_addr, formula=label, size=8, color=C["muted"], h=XL_LEFT, indent=1))
        else:
            put(ws, lab_addr, label, size=8, color=C["muted"], h=XL_LEFT, indent=1)
        lab.ShrinkToFit = True      # 긴 라벨(52주 범위·조회 시각)이 옆 카드로 넘치지 않게
        ok = _com_ok(builder, f"정보 카드 '{label[:8]}'", lambda f=formula, a=first, n=nf, s=size, w=wrap: put(
            ws, a, formula=f, nf=n, bold=True, size=s, color=C["navy"], h=XL_LEFT, indent=1, wrap=w))
        if ok:
            border(ws.Range(f"{lab_addr}:{val.Cells(val.Rows.Count, val.Columns.Count).Address}"), "line", edges=(XL_EDGE_LEFT,))
    cf_expr(ws.Range("F8"), '=ISNUMBER(SEARCH("대회 편입",$F$8))', font=C["good"], bold=True)
    cf_expr(ws.Range("F8"), '=ISNUMBER(SEARCH("대회 외",$F$8))', font=C["warn"], bold=True)
    cf_expr(ws.Range("B8"), '=LEFT($B$8,1)="⚠"', font=C["up"], bold=True)
    cf_expr(ws.Range("Q8"), '=LEFT($Q$8,1)="▲"', font=C["up"])      # 상승 빨강(국내 관례)
    cf_expr(ws.Range("Q8"), '=LEFT($Q$8,1)="▼"', font=C["down"])    # 하락 파랑


# ------------------------------------------------------------------------------------------------ 구역 제목·상태 줄
@dataclasses.dataclass(frozen=True)
class _Sec:
    """구역 제목 줄 하나: 제목 셀 위치, 밑줄이 그어질 열 범위, 연결된 표, 안내(오른쪽)·상태 수식."""

    text: str
    first: str
    last: str
    row: int
    table: str
    note: Optional[str] = None
    status: Optional[str] = None
    today: bool = False   # 당일 구역 (안내 칸을 주황 강조)


def _sections(have_info: bool = True) -> list[_Sec]:
    """구역 제목 줄 13개(구역 1~10, 3·6은 둘씩, 종목 정보 원천표)의 위치·안내·상태 수식.

    Args:
        have_info: tblA_Info가 있는가(없으면 목표가 문구에서 현재가 대비 괴리율을 뺀다).
    """
    t = {k: v.name for k, v in (("inv", INVESTOR), ("ts", TRADESIZE), ("pb", PBAR), ("pf", PROFILE), ("es", ESTIMATE), ("csl", CSL),
                                ("tg", TARGETS), ("tr", TARGET_TREND), ("ke", KISEST), ("fin", FIN), ("nw", NEWS), ("ev", EVENTS),
                                ("info", INFO))}
    inv, ts, pb, pf, es, csl, tg, tr, ke, fin, nw, ev, info = (t[k] for k in
                                                              ("inv", "ts", "pb", "pf", "es", "csl", "tg", "tr", "ke", "fin", "nw", "ev", "info"))
    return [
        _Sec("1. 가격·투자자 일별 — 최근 120 완료 세션 (순매수 억원)", "B", "U", 10, inv,
             _note_formula(inv, "일자"),
             _status_formula(inv, "일자", _ok(_lit(" · "), f'COUNT({inv}[일자])', _lit("세션 "),
                                              f'TEXT(MIN({inv}[일자]),"yyyy-mm-dd")', _lit(" ~ "),
                                              f'TEXT(MAX({inv}[일자]),"yyyy-mm-dd")',
                                              _lit(" · 완료 세션만(15:40 전에는 어제까지) · 누적 = 첫 세션부터 합산")),
                             none_hint=" — KIS가 일별 수급을 주지 않은 종목입니다")),
        _Sec("2. 체결금액별 매매비중 — 당일", "B", "Q", 157, ts, _note_formula(ts, "기준일", today_tag=True),
             _status_formula(ts, "매수거래량", _lit(" · 구간 = 체결 1건의 금액 크기 · 장중에는 계속 변하는 값(비율 = 종목 당일 전체 거래량 대비)"),
                             none_hint=" — 거래량이 없습니다(장 시작 전이거나 거래 없음)"),
             today=True),
        _Sec("3-a. 당일 매물대 — KIS 가격대별 거래량·비중", "B", "J", 174, pb, _note_formula(pb, "기준일", today_tag=True),
             _status_formula(pb, "가격", _ok(_lit(" · 가격대 "), f'COUNT({pb}[가격])', _lit("개 · 비중 합 "),
                                            f'TEXT(SUM({pb}[비중]),"0.0%")', _lit(" (가격대 100개 초과분은 잘림)")),
                             none_hint=" — 거래량이 없습니다(장 시작 전이거나 거래 없음)"),
             today=True),
        _Sec("3-b. 최근 N세션 매물대 — 일봉 거래량을 20구간에 배분", "L", "Y", 174, pf, _note_formula(pf, "종료일"),
             _status_formula(pf, "구간", _ok(_lit(" · 최근 "), f'INDEX({pf}[세션수],1)', _lit("세션 "),
                                           f'TEXT(INDEX({pf}[시작일],1),"yyyy-mm-dd")', _lit(" ~ "),
                                           f'TEXT(INDEX({pf}[종료일],1),"mm-dd")', _lit(" · 현재가 "),
                                           f'TEXT(INDEX({pf}[현재가],1),"#,##0")', _lit("원은 "),
                                           f'IFERROR(INDEX({pf}[가격대],MATCH("Y",{pf}[현재가구간],0)),"-")', _lit(" 구간")))),
        _Sec("4. 외인·기관 추정가집계 — 당일 (입력 시점별 누적)", "B", "N", 280, es,
             _note_formula(es, "기준일", today_tag=True),
             _status_formula(es, "외국인수량",
                             _ok(_lit(" · 입력된 시점 "), f'COUNT({es}[합산수량])', _lit("/5"),
                                 _lit(" · 최신 "), f'IFERROR(XLOOKUP(TRUE,ISNUMBER({es}[합산수량]),{es}[입력시점],"-",0,-1),"-")',
                                 _lit(" 합산 ≈ "), f'IFERROR(TEXT(XLOOKUP(TRUE,ISNUMBER({es}[합산≈억]),{es}[합산≈억],0,0,-1),"+#,##0.0;-#,##0.0"),"-")',
                                 _lit("억(수량 × 현재가) · 수량은 장 시작부터의 누적")),
                             none_hint=" — 아직 입력된 시점이 없습니다(09:30 이전이거나 휴장일, 정상)"),
             today=True),
        _Sec("5. 신용·공매도·대차 일별 추이 — 최근 60세션", "B", "N", 297, csl, _note_formula(csl, "일자"),
             _status_formula(csl, "일자", _ok(_lit(" · "), f'COUNT({csl}[일자])', _lit("세션 "),
                                              f'TEXT(MIN({csl}[일자]),"yyyy-mm-dd")', _lit(" ~ "),
                                              f'TEXT(MAX({csl}[일자]),"yyyy-mm-dd")',
                                              _lit(" · 신용잔고는 결제일 기준이라 최근 2세션이 비어 있음(정상)")),
                             none_hint=" — 신용·공매도·대차 자료가 없습니다")),
        _Sec("6-a. 목표주가 — 증권사별 최근 의견·목표가", "B", "L", 362, tg, _note_formula(tg, "일자", label="최근 의견"),
             _status_formula(tg, "증권사", _ok(_lit(" · 증권사 "), f'COUNTA({tg}[증권사])', _lit("곳 · 목표가 평균 "),
                                              _price_gap(f'AVERAGE({tg}[목표가])', have_info)),
                             numeric_key=False,
                             none_hint=" — 최근 6개월 안에 의견을 낸 증권사가 없습니다(이 종목군의 약 42%)")),
        _Sec("6-b. 월말 목표주가 컨센서스 추이", "N", "V", 362, tr, _note_formula(tr, "기준일"),
             _status_formula(tr, "목표가평균", _ok(_lit(" · 현재 "),
                                                  _price_gap(f'INDEX({tr}[목표가평균],MATCH("현재",{tr}[구분],0))', have_info),
                                                  _lit(" · 증권사 "),
                                                  f'IFERROR(INDEX({tr}[증권사수],MATCH("현재",{tr}[구분],0))&"곳","-")'),
                             none_hint=" — 목표가 자료가 없습니다(이 종목군의 약 42%, 정상)")),
        _Sec("7. KIS 추정실적 — 한국투자증권 리서치 자체 추정 (억원)", "B", "Y", 412, ke, _note_formula(ke, "추정일", label="추정일"),
             _status_formula(ke, "매출억", _ok(_lit(" · 연도 "), f'COUNTA({ke}[연도])', _lit("개(E = 추정) · 추정일 "),
                                              f'IFERROR(IF(INDEX({ke}[추정일],1)="","-",TEXT(INDEX({ke}[추정일],1),"yyyy-mm-dd")),"-")', _lit(" · KIS 의견 "),
                                              f'IFERROR(INDEX({ke}[KIS의견],1)&"","-")',
                                              _lit(" · PER = 현재가 ÷ EPS, KIS_PER = 추정일 주가 기준")),
                             none_hint=" — KIS가 추정을 제공하지 않는 종목입니다(이 종목군의 약 79%, 정상)")),
        _Sec("8. 분기 실적 — 최근 8분기 (분기 단독값 · YoY)", "B", "M", 425, fin,
             f'=IFERROR(IF(COUNTA({fin}[분기])=0,"","최근 분기 "&INDEX({fin}[분기],ROWS({fin}[분기]))&" · 조회 "&TEXT(MAX({fin}[조회시각]),"mm-dd hh:mm")&" · [조회]로 갱신"),"")',
             _status_formula(fin, "매출억", _ok(_lit(" · "), f'COUNTA({fin}[분기])', _lit("분기 "),
                                               f'INDEX({fin}[분기],1)', _lit(" ~ "), f'INDEX({fin}[분기],ROWS({fin}[분기]))',
                                               _lit(" · 분기 단독 = 누적 − 직전 분기 · YoY = 전년 동분기 대비")),
                             none_hint=" — KIS 재무 자료가 없는 종목입니다")),
        _Sec("9. 뉴스·공시 제목 — 최근 40건", "AC", BAND_LAST_COL, 10, nw, _note_formula(nw, "일시", fmt="mm-dd hh:mm", label="최신"),
             _status_formula(nw, "일시", _ok(_lit(" · "), f'COUNT({nw}[일시])', _lit("건 · 최신순")),
                             none_hint=" — 이 종목의 뉴스·공시가 없습니다")),
        _Sec("10. 이 종목의 이벤트 — 오늘 −7일 ~ +60일 (예탁원 일정)", "AA", "AG", 55, ev,
             f'=IFERROR(IF(COUNT({ev}[조회시각])=0,"","조회 "&TEXT(MAX({ev}[조회시각]),"mm-dd hh:mm")&" · [조회]로 갱신"),"")',
             _status_formula(ev, "일자", _ok(_lit(" · "), f'COUNT({ev}[일자])', _lit("건 · D-day는 조회 시점 기준 저장값")),
                             none_hint=" — 기간 안에 이벤트가 없습니다(정상)")),
        _Sec("종목 정보 — 원천표 (상단 정보 카드가 이 표를 읽습니다)", "AA", BAND_LAST_COL, 90, info, None,
             _status_formula(info, "현재가", _lit(" · 1행")), False),
    ]


def _draw_section(builder: Any, ws: Any, sec: _Sec, have_table: bool) -> None:
    """구역 제목(밑줄) + 오른쪽 안내 칸 + 바로 아래 상태 줄. 표가 없으면 제목만 그린다.

    당일 구역(2·3-a·4)의 안내는 제목 줄 오른쪽 끝 3칸을 합친 주황 칩(`당일 mm-dd 기준 · 조회 hh:mm`)으로 보여 눈에 띄게 한다.
    """
    chip = None
    if have_table and sec.note and sec.today:
        last = _col_number(sec.last)
        chip = ws.Range(f"{_col_letter(last - 2)}{sec.row}:{sec.last}{sec.row}")
        chip.Merge()
    section(ws, f"{sec.first}{sec.row}:{sec.last}{sec.row}", sec.text)
    if not have_table:
        return
    if sec.note:
        addr = chip.Cells(1, 1).Address if chip is not None else f"{sec.last}{sec.row}"
        ok = _com_ok(builder, f"'{sec.text}' 안내 수식", lambda: put(ws, addr, formula=sec.note, size=8,
                                                                    color=C["warn"] if sec.today else C["muted"],
                                                                    bold=sec.today, h=XL_CENTER if sec.today else XL_RIGHT))
        if ok and chip is not None:
            chip.Interior.Color = rgb(C["warn_l"])
    if sec.status:
        addr = f"{sec.first}{sec.row + 1}"
        if _com_ok(builder, f"'{sec.text}' 상태 줄", lambda: put(ws, addr, formula=sec.status, size=9, color=C["muted"], h=XL_LEFT)):
            cf_expr(ws.Range(addr), f'=LEFT({sec.first}{sec.row + 1},1)="▲"', font=C["up"], bold=True)
            cf_expr(ws.Range(addr), f'=LEFT({sec.first}{sec.row + 1},1)="●"', font=C["good"])


# ------------------------------------------------------------------------------------------------ 표 서식·조건부 서식
def _block(ws: Any, geo: _Geo, rows: int, col: str) -> Any:
    """표 열의 예약 칸 범위(데이터 첫 행부터 rows칸 — 표 아래 빈 칸까지)."""
    c = geo.cols[col]
    return ws.Range(ws.Cells(geo.first_row, c), ws.Cells(geo.first_row + rows - 1, c))


def _format_table(builder: Any, ws: Any, lo: Any, t: _Tbl, style_name: Optional[str]) -> _Geo:
    """표 스타일·새로 고침 방식·머리글·예약 행 전체의 열 서식. 기대한 열이 없으면 경고만 남기고 있는 열만 입힌다.

    열 서식은 표 데이터 범위(0행이면 없음)가 아니라 표 아래 예약 칸까지 시트 칸에 직접 입힌다 — 빈 표에서도 되고, 표가 늘어나도 같은 모양이다.
    """
    _check_position(builder, lo, t)
    if style_name:
        _com_ok(builder, f"{t.name} 표 스타일 '{style_name}'", lambda: setattr(lo, "TableStyle", style_name))
    _overwrite_on_refresh(lo)
    with contextlib.suppress(pywintypes.com_error):
        lo.ShowAutoFilterDropDown = False      # 머리글 필터 화살표가 좁은 칸의 글자를 가리지 않게(이 페이지의 표는 보기 전용)
    geo = _geometry(lo)
    missing = [n for n in t.cols if n not in geo.cols]
    if missing:
        builder.say(f"[{SHEET}] 경고: {t.name}에 없는 열 {', '.join(missing)} — 그 열의 서식은 건너뜁니다")
    hdr = lo.HeaderRowRange
    style(hdr, size=9, bold=True, h=XL_CENTER)
    hdr.ShrinkToFit = True
    # 예약 행 전체(표 아래 빈 행 포함)에 글자 크기를 먼저 맞춘 뒤 열마다 서식
    style(ws.Range(ws.Cells(geo.first_row, geo.left), ws.Cells(geo.first_row + t.rows - 1, geo.right)), size=9)
    for name, spec in t.cols.items():
        if name not in geo.cols:
            continue
        blk = _block(ws, geo, t.rows, name)
        if spec.get("nf"):
            set_nf(blk, spec["nf"])
        keys = {k: spec[k] for k in ("bold", "color", "size", "italic", "h") if k in spec and spec[k] is not None}
        if keys:
            style(blk, **keys)
    return geo


def _status_cf(ws: Any, geo: _Geo, rows: int) -> None:
    """상태 열: 오류·이전 데이터는 빨강 굵게."""
    if "상태" not in geo.cols:
        return
    letter = _col_letter(geo.cols["상태"])
    r0 = geo.first_row
    cf_expr(_block(ws, geo, rows, "상태"), f'=OR(LEFT(${letter}{r0},2)="오류",LEFT(${letter}{r0},2)="이전")', font=C["up"], bold=True)


def _row_highlight(ws: Any, geo: _Geo, rows: int, flag_col: str, value: str, *, fill_color: str) -> None:
    """flag_col 값이 value인 행 전체(표 폭)를 칠한다."""
    if flag_col not in geo.cols:
        return
    letter = _col_letter(geo.cols[flag_col])
    r0 = geo.first_row
    whole = ws.Range(ws.Cells(r0, geo.left), ws.Cells(r0 + rows - 1, geo.right))
    cf_expr(whole, f'=${letter}{r0}="{value}"', fill_color=fill_color, bold=True)


def _bar_cf(builder: Any, ws: Any, geo: _Geo, rows: int, col: str, color: str = "#BBD2F5") -> None:
    """비율 열 안의 막대(데이터 막대). 0부터 시작."""
    if col not in geo.cols:
        return

    def add() -> None:
        db = cf_databar(_block(ws, geo, rows, col), color)
        with contextlib.suppress(pywintypes.com_error):
            db.MinPoint.Modify(0, 0)   # xlConditionValueNumber, 0 — 막대가 0부터 비례하도록

    _com_ok(builder, f"{col} 데이터 막대", add)


def _table_cf(builder: Any, ws: Any, geo: _Geo, t: _Tbl) -> None:
    """표마다 조건부 서식: 공통(상태 열 빨강) + 표별(현재가 구간 행·데이터 막대·추정 행·이벤트 D-day)."""
    _status_cf(ws, geo, t.rows)
    if t is PBAR:
        _row_highlight(ws, geo, t.rows, "현재가구간", "Y", fill_color=C["accent_l"])
        _bar_cf(builder, ws, geo, t.rows, "비중")
    elif t is PROFILE:
        _row_highlight(ws, geo, t.rows, "현재가구간", "Y", fill_color="#FFE9C7")
        _bar_cf(builder, ws, geo, t.rows, "비중")
    elif t is TRADESIZE:
        _bar_cf(builder, ws, geo, t.rows, "매수비율", "#F4B7B5")
        _bar_cf(builder, ws, geo, t.rows, "매도비율", "#B5CDF0")
    elif t is KISEST and "추정여부" in geo.cols:
        letter = _col_letter(geo.cols["추정여부"])
        r0 = geo.first_row
        cf_expr(ws.Range(ws.Cells(r0, geo.left), ws.Cells(r0 + t.rows - 1, geo.right)), f'=${letter}{r0}="E"', fill_color=C["warn_l"])
    elif t is EVENTS and "D-day" in geo.cols:
        letter = _col_letter(geo.cols["D-day"])
        r0 = geo.first_row
        whole = ws.Range(ws.Cells(r0, geo.left), ws.Cells(r0 + t.rows - 1, geo.right))
        cf_expr(whole, f'=AND(ISNUMBER(${letter}{r0}),${letter}{r0}<0)', font="#9AA5B1")
        cf_expr(whole, f'=AND(ISNUMBER(${letter}{r0}),${letter}{r0}>=0,${letter}{r0}<=7)', fill_color=C["warn_l"], bold=True)


# ------------------------------------------------------------------------------------------------ 숨김 도우미 (차트용)
def _row_number_expr(col: str, first_row: int) -> str:
    """도우미 열 안에서 이 칸이 몇 번째 줄인지(1부터) 세는 수식 조각: ROWS(BD$177:BD177)."""
    return f"ROWS({col}${first_row}:{col}{first_row})"


def _write_helpers(builder: Any, ws: Any, have_profile: bool, have_trend: bool) -> None:
    """숨김 열 BD:BF — 매물대 차트 두 계열(현재가 구간만 / 나머지)과 제목, 목표가 추이 차트의 현재가 선.

    표 열을 직접 가리키지 않고 INDEX(표[열], 몇 번째)로 읽으므로 표 크기가 바뀌어도 맞고, 행이 모자라면 0(막대 없음) 또는 #N/A(선 끊김)이 된다.
    """
    first, last = HELPER_COLS[0], HELPER_COLS[-1]
    ws.Range(f"{first}1").Value = "차트 도우미 (숨김 — 삭제하지 마세요)"
    if have_profile:
        pf = PROFILE.name
        title = (f'=IFERROR("최근 "&INDEX({pf}[세션수],1)&"세션 매물대 — 가격 구간별 거래량 비중 (주황 = 현재가 "'
                 f'&TEXT(INDEX({pf}[현재가],1),"#,##0")&"원이 속한 구간)","N일 매물대 — 가격 구간별 거래량 비중 (주황 = 현재가가 속한 구간)")')
        _com_ok(builder, "매물대 차트 제목 수식", lambda: setattr(ws.Range(f"{first}{HELPER_PROFILE_ROW - 2}"), "Formula2", title))
        r1, r2 = HELPER_PROFILE_ROW, HELPER_PROFILE_ROW + PROFILE.rows - 1
        for col, mark_value in ((HELPER_COLS[0], 0), (HELPER_COLS[1], 1)):
            n = _row_number_expr(col, r1)
            share = f"N(INDEX({pf}[비중],{n}))"
            on_mark = f'INDEX({pf}[현재가구간],{n})="Y"'
            # 첫 열 = 현재가 구간이 아닌 칸만 값, 둘째 열 = 현재가 구간 칸만 값
            formula = f"=IFERROR(IF({on_mark},{share if mark_value else 0},{0 if mark_value else share}),0)"
            _com_ok(builder, f"매물대 차트 도우미 {col}", lambda c=col, f=formula: setattr(ws.Range(f"{c}{r1}:{c}{r2}"), "Formula2", f))
    if have_trend:
        col = HELPER_COLS[2]
        r1, r2 = HELPER_TREND_ROW, HELPER_TREND_ROW + TARGET_TREND.rows - 1
        n = _row_number_expr(col, r1)
        f = f'=IFERROR(IF(INDEX({TARGET_TREND.name}[목표가평균],{n})="",NA(),INDEX(tblA_Info[현재가],1)),NA())'
        _com_ok(builder, "목표가 추이 도우미", lambda: setattr(ws.Range(f"{col}{r1}:{col}{r2}"), "Formula2", f))
    ws.Range(f"{first}:{last}").EntireColumn.Hidden = True


# ------------------------------------------------------------------------------------------------ 차트
# 차트 계열이 읽는 범위
#   - 행 수가 늘 정해진 표(체결금액 8·매물대 20·추정가집계 5·목표가 추이 13)는 그 칸 범위를 그대로 쓴다. 데이터가 없어도(1행) 같은 칸이라 안전하다.
#   - 행 수가 바뀌는 표(투자자 ≤120·신용공매도 ≤60·분기 실적 ≤8)는 `분석차트_…` 시트 범위 이름(동적 범위: 데이터 행 수만큼)으로 읽는다.
#     표 열 범위를 직접 걸면 빈 표(0행)에서는 차트를 만들 수 없고, 1행에서 만든 계열은 표가 늘어도 따라 늘지 않았다(실측).
#     이름은 데이터 칸 수를 세는 수식이라 0행 → N행 → 적은 행 어디서나 길이가 맞고, 표를 덮어쓰기로 새로 고치므로 위치도 바뀌지 않는다.
CHART_NAME_PREFIX = "분석차트_"      # 시세판 차트의 통합문서 범위 이름(차트_일자 …)과 겹치지 않게 접두어를 다르게


def _col_range(ws: Any, geo: _Geo, col: str, rows: int) -> Any:
    """표 열의 예약 칸 범위(데이터 첫 행부터 rows칸)."""
    letter = _col_letter(geo.cols[col])
    return ws.Range(f"{letter}{geo.first_row}:{letter}{geo.first_row + rows - 1}")


def _dynamic_range(geo: _Geo, rows: int, col: str, key_col: str, text_key: bool = False) -> str:
    """데이터 행 수만큼만 잡는 범위 수식. 예: ='종목분석'!$D$35:INDEX('종목분석'!$D$35:$D$154,MAX(1,COUNT('종목분석'!$C$35:$C$154))).

    key_col의 값이 있는 칸 수(숫자 열은 COUNT, 글자 열은 COUNTA)를 길이로 쓰고 최소 1칸을 보장한다(빈 표에서도 범위가 유효하도록).
    """
    sh = f"'{SHEET}'!"
    c, k = _col_letter(geo.cols[col]), _col_letter(geo.cols[key_col])
    r1, r2 = geo.first_row, geo.first_row + rows - 1
    count = "COUNTA" if text_key else "COUNT"
    return f"={sh}${c}${r1}:INDEX({sh}${c}${r1}:${c}${r2},MAX(1,{count}({sh}${k}${r1}:${k}${r2})))"


def _dynamic_refs(ws: Any, geo: _Geo, rows: int, prefix: str, key_col: str, cols: dict, *, text_key: bool = False) -> dict:
    """동적 범위 이름을 시트 범위로 만들고, 열 이름 → 계열 참조(`='시트'!이름`)를 돌려준다. cols = {이름 꼬리: 열 이름}."""
    refs: dict = {}
    for tail, col in cols.items():
        name = f"{CHART_NAME_PREFIX}{prefix}_{tail}"
        with contextlib.suppress(pywintypes.com_error):
            ws.Names(name).Delete()
        nm = ws.Names.Add(Name=name, RefersTo=_dynamic_range(geo, rows, col, key_col, text_key))
        with contextlib.suppress(pywintypes.com_error):
            nm.Visible = False      # 이름 관리자에 잡다한 이름이 늘지 않게(차트는 숨긴 이름도 읽는다)
        refs[col] = f"='{SHEET}'!{name}"
    return refs


def _need(builder: Any, geo: _Geo, what: str, cols: list) -> bool:
    """차트에 필요한 열이 표에 다 있는지. 없으면 경고만 남기고 False(그 차트는 건너뜀)."""
    miss = [c for c in cols if c not in geo.cols]
    if miss:
        builder.say(f"[{SHEET}] 경고: {what} 차트에 필요한 열 {', '.join(miss)}이(가) 없어 차트를 건너뜁니다")
        return False
    return True


def _axis_title(ch: Any, which: int, group: int, text: str) -> None:
    """차트 축 제목(작은 회색 글자). which = 1 가로·2 세로, group = 1 주축·2 보조축."""
    with contextlib.suppress(pywintypes.com_error):
        ax = ch.Axes(which, group)
        ax.HasTitle = True
        ax.AxisTitle.Text = text
        ax.AxisTitle.Font.Size = 8
        ax.AxisTitle.Font.Bold = False
        ax.AxisTitle.Font.Color = rgb(C["muted"])


def _chart_price_flow(builder: Any, ws: Any, geo: _Geo, t: _Tbl) -> None:
    """구역 1 차트: 종가(주축) + 외국인·기관계·개인 누적 순매수(보조축). 행 수가 바뀌므로 동적 범위 이름."""
    cols = ["일자", "종가", "외국인누적억", "기관계누적억", "개인누적억"]
    if not _need(builder, geo, "가격·누적 순매수", cols):
        return
    ref = _dynamic_refs(ws, geo, t.rows, "투자자", "일자", {"일자": "일자", "종가": "종가", "외국인": "외국인누적억",
                                                          "기관계": "기관계누적억", "개인": "개인누적억"})
    _, ch = chart(ws, CHART_PRICE, XL_LINE, "종가(왼쪽 축)와 외국인·기관계·개인 누적 순매수(오른쪽 축, 억원)")
    x = ref["일자"]
    series(ch, "종가", x, ref["종가"], color=C["navy"], weight=2.25)
    series(ch, "외국인 누적", x, ref["외국인누적억"], color=COLOR_FOREIGN, weight=1.5, axis=XL_SECONDARY)
    series(ch, "기관계 누적", x, ref["기관계누적억"], color=COLOR_INST, weight=1.5, axis=XL_SECONDARY)
    series(ch, "개인 누적", x, ref["개인누적억"], color=COLOR_PERSON, weight=1.25, axis=XL_SECONDARY)
    style_axes(ch, y_nf="#,##0", x_nf="mm/dd", y2_nf="#,##0", legend=XL_LEGEND_TOP)
    _axis_title(ch, 2, 1, "종가(원)")
    _axis_title(ch, 2, 2, "누적 순매수(억원)")


def _chart_profile(builder: Any, ws: Any, geo: _Geo, t: _Tbl) -> None:
    """구역 3-b 차트: 가격 구간별 거래량 비중 가로 막대(20구간 고정). 현재가 구간은 색이 다른 둘째 계열(숨김 도우미)."""
    if not _need(builder, geo, "N일 매물대", ["가격대", "비중", "현재가구간"]):
        return
    _, ch = chart(ws, CHART_PROFILE, XL_BAR_CLUSTERED, "N일 매물대 — 가격 구간별 거래량 비중 (주황 = 현재가가 속한 구간)")
    with contextlib.suppress(pywintypes.com_error):       # 제목을 숨김 셀에 연결(세션 수·현재가가 새로 고침을 따라감)
        ch.ChartTitle.Formula = f"='{SHEET}'!${HELPER_COLS[0]}${HELPER_PROFILE_ROW - 2}"
    cat = _col_range(ws, geo, "가격대", t.rows)
    r1, r2 = HELPER_PROFILE_ROW, HELPER_PROFILE_ROW + t.rows - 1
    s1 = series(ch, "거래량 비중", cat, ws.Range(f"{HELPER_COLS[0]}{r1}:{HELPER_COLS[0]}{r2}"), fill_color=COLOR_BAR)
    s2 = series(ch, "현재가 구간", cat, ws.Range(f"{HELPER_COLS[1]}{r1}:{HELPER_COLS[1]}{r2}"), fill_color=C["warn"])
    with contextlib.suppress(pywintypes.com_error):
        grp = ch.ChartGroups(1)
        grp.Overlap = 100
        grp.GapWidth = 35
    for s in (s1, s2):
        with contextlib.suppress(pywintypes.com_error):
            s.HasDataLabels = True
            dl = s.DataLabels()
            dl.NumberFormat = "0.0%;;"          # 0은 라벨을 숨김(그 구간에 값이 없는 계열)
            dl.Font.Size = 8
            dl.Position = 2                     # xlLabelPositionOutsideEnd
    style_axes(ch, y_nf="0%", category=False, legend=XL_LEGEND_TOP)
    with contextlib.suppress(pywintypes.com_error):
        ch.Axes(1).TickLabels.Font.Size = 8      # 가격대 글자


def _chart_csl(builder: Any, ws: Any, geo: _Geo, t: _Tbl) -> None:
    """구역 5 차트: 공매도 비중(막대, 왼쪽) + 신용잔고율(선, 오른쪽). 행 수가 바뀌므로 동적 범위 이름."""
    if not _need(builder, geo, "신용·공매도", ["일자", "공매도비중", "신용잔고율"]):
        return
    ref = _dynamic_refs(ws, geo, t.rows, "신용", "일자", {"일자": "일자", "공매도": "공매도비중", "잔고율": "신용잔고율"})
    _, ch = chart(ws, CHART_CSL, XL_COLUMN_CLUSTERED, "공매도 비중(막대, 왼쪽)과 신용잔고율(선, 오른쪽)")
    x = ref["일자"]
    series(ch, "공매도 비중", x, ref["공매도비중"], fill_color="#F2C48D")
    series(ch, "신용잔고율", x, ref["신용잔고율"], color=C["navy"], weight=2, kind=XL_LINE, axis=XL_SECONDARY)
    style_axes(ch, y_nf="0%", x_nf="mm/dd", y2_nf="0.0%", legend=XL_LEGEND_TOP)


def _chart_trend(builder: Any, ws: Any, geo: _Geo, t: _Tbl) -> None:
    """구역 6-b 차트: 월말 평균 목표가와 최고·최저, 현재가(점선). 13칸 고정."""
    if not _need(builder, geo, "목표가 추이", ["기준일", "목표가평균", "최고목표가", "최저목표가"]):
        return
    _, ch = chart(ws, CHART_TREND, XL_LINE_MARKERS, "월말 평균 목표가와 최고·최저 (회색 점선 = 현재가)")
    x = _col_range(ws, geo, "기준일", t.rows)
    series(ch, "평균 목표가", x, _col_range(ws, geo, "목표가평균", t.rows), color=C["navy"], weight=2.25)
    hi = series(ch, "최고", x, _col_range(ws, geo, "최고목표가", t.rows), color=C["up"], weight=1)
    low = series(ch, "최저", x, _col_range(ws, geo, "최저목표가", t.rows), color=C["down"], weight=1)
    r1, r2 = HELPER_TREND_ROW, HELPER_TREND_ROW + t.rows - 1
    cur = series(ch, "현재가", x, ws.Range(f"{HELPER_COLS[2]}{r1}:{HELPER_COLS[2]}{r2}"), color="#8A94A3", weight=1.25)
    for s in (hi, low, cur):
        with contextlib.suppress(pywintypes.com_error):
            s.MarkerStyle = -4142        # xlMarkerStyleNone — 평균 목표가만 점을 찍는다
    with contextlib.suppress(pywintypes.com_error):
        cur.Format.Line.DashStyle = 4    # msoLineDash
    style_axes(ch, y_nf="#,##0", x_nf="yy.mm", legend=XL_LEGEND_TOP)


def _chart_tradesize(builder: Any, ws: Any, geo: _Geo, t: _Tbl) -> None:
    """구역 2 차트: 체결금액 구간별 매수·매도 비율(막대) + 순매수 비율(선). 8칸 고정."""
    if not _need(builder, geo, "체결금액별", ["구간", "매수비율", "매도비율", "순매수비율"]):
        return
    _, ch = chart(ws, CHART_TRADESIZE, XL_COLUMN_CLUSTERED, "체결금액 구간별 매수·매도 비율 (선 = 순매수 비율)")
    x = _col_range(ws, geo, "구간", t.rows)
    series(ch, "매수 비율", x, _col_range(ws, geo, "매수비율", t.rows), fill_color="#E58A87")
    series(ch, "매도 비율", x, _col_range(ws, geo, "매도비율", t.rows), fill_color="#7FA6DB")
    series(ch, "순매수 비율", x, _col_range(ws, geo, "순매수비율", t.rows), color=C["navy"], weight=2, kind=XL_LINE_MARKERS)
    style_axes(ch, y_nf="0%", category=False, legend=XL_LEGEND_TOP)


def _chart_estimate(builder: Any, ws: Any, geo: _Geo, t: _Tbl) -> None:
    """구역 4 차트: 입력 시점별 외국인·기관 추정 순매수(≈억원) 막대 + 합산 선. 5칸 고정."""
    if not _need(builder, geo, "추정가집계", ["입력시점", "외국인≈억", "기관≈억", "합산≈억"]):
        return
    _, ch = chart(ws, CHART_ESTIMATE, XL_COLUMN_CLUSTERED, "입력 시점별 누적 추정 순매수 (≈억원)")
    x = _col_range(ws, geo, "입력시점", t.rows)
    series(ch, "외국인", x, _col_range(ws, geo, "외국인≈억", t.rows), fill_color=COLOR_FOREIGN)
    series(ch, "기관", x, _col_range(ws, geo, "기관≈억", t.rows), fill_color=COLOR_INST)
    series(ch, "합산", x, _col_range(ws, geo, "합산≈억", t.rows), color=C["navy"], weight=2, kind=XL_LINE_MARKERS)
    style_axes(ch, y_nf="#,##0", category=False, legend=XL_LEGEND_TOP)


def _chart_fin(builder: Any, ws: Any, geo: _Geo, t: _Tbl) -> None:
    """구역 8 차트: 분기별 매출·영업이익·순이익(막대, 억원) + 영업이익률(선, 오른쪽). 분기 수가 바뀌므로 동적 범위 이름."""
    if not _need(builder, geo, "분기 실적", ["분기", "매출억", "영업이익억", "순이익억", "영업이익률"]):
        return
    ref = _dynamic_refs(ws, geo, t.rows, "실적", "분기", {"분기": "분기", "매출": "매출억", "영업이익": "영업이익억", "순이익": "순이익억",
                                                       "이익률": "영업이익률"}, text_key=True)
    _, ch = chart(ws, CHART_FIN, XL_COLUMN_CLUSTERED, "분기별 매출·영업이익·순이익 (억원)과 영업이익률")
    x = ref["분기"]
    series(ch, "매출", x, ref["매출억"], fill_color="#C9D3E3")
    series(ch, "영업이익", x, ref["영업이익억"], fill_color=C["navy"])
    series(ch, "순이익", x, ref["순이익억"], fill_color=C["accent"])
    series(ch, "영업이익률", x, ref["영업이익률"], color=C["warn"], weight=2, kind=XL_LINE_MARKERS, axis=XL_SECONDARY)
    style_axes(ch, y_nf="#,##0", category=False, y2_nf="0%", legend=XL_LEGEND_TOP)


# ------------------------------------------------------------------------------------------------ 시트 준비·해제
def _prepare_sheet(ws: Any) -> None:
    """다시 부르면 겹치지 않게 이 시트 전용 장식(조건부 서식·도형·링크·병합)을 지우고 열 너비·행 높이를 정한다."""
    sheet_setup(ws, bg=False, zoom=85, tab=C["accent"])
    ws.Cells.FormatConditions.Delete()
    for shp in list(ws.Shapes):
        shp.Delete()
    with contextlib.suppress(pywintypes.com_error):
        ws.Hyperlinks.Delete()
    # 위쪽 카드와 구역 제목 줄의 당일 칩(병합)을 다시 만든다. 표가 든 범위에 UnMerge를 걸면 Excel이 오류를 내므로(실측) 병합이 있는 칸만 푼다
    ws.Range(f"A1:{BAND_LAST_COL}{FROZEN_ROWS}").UnMerge()
    for sec in _sections():
        ws.Range(f"{sec.first}{sec.row}:{sec.last}{sec.row}").UnMerge()
    for nm in list(ws.Names):      # 이전 빌드의 차트 범위 이름(다시 만든다). 표 적재가 만든 ExternalData_ 이름은 건드리지 않는다
        if nm.Name.split("!")[-1].startswith(CHART_NAME_PREFIX):
            nm.Delete()
    for n, w in _lane_widths().items():
        ws.Columns(n).ColumnWidth = w
    for r, h in ((1, 6), (2, 28), (3, 16), (4, 16), (5, 16), (6, 28), (7, 14), (8, 28), (9, 6)):
        ws.Rows(r).RowHeight = h
    ws.Range(f"{BODY_FIRST_ROW}:{BODY_LAST_ROW}").RowHeight = BODY_ROW_HEIGHT


def build(builder: Any) -> None:
    """`종목분석` 시트를 만든다(표 13개는 LOADS 위치에 이미 적재돼 있어야 함).

    Args:
        builder: build_dashboard.Builder 또는 하네스 대역(StandInBuilder). wb·ws·lo·table_style·say·nav_links·
            HEADER_RIGHT·HEADER_STATUS를 쓴다.

    Raises:
        없음 — 시트나 표가 없으면 경고만 남기고 가능한 부분까지 그린다(COM 서식 오류는 대부분 경고로 처리, 그 밖은 그대로 올라옴).

    Example:
        import pages.page_analysis as p; p.build(builder)
    """
    ws = _sheet(builder)
    if ws is None:
        return
    los = {t.name: _table(builder, t.name) for t in TABLES}
    style_name = getattr(builder, "table_style", None)

    _prepare_sheet(ws)
    have_list = _define_names(builder)
    _header_band(builder, ws)
    _input_row(builder, ws, have_info=los[INFO.name] is not None, have_list=have_list)
    if los[INFO.name] is not None:
        _info_cards(builder, ws)

    geos: dict[str, _Geo] = {}
    for t in TABLES:
        lo = los[t.name]
        geos[t.name] = _format_table(builder, ws, lo, t, style_name) if lo is not None else _fallback_geo(t)
        if lo is not None:
            _table_cf(builder, ws, geos[t.name], t)
    for sec in _sections(have_info=los[INFO.name] is not None):
        _draw_section(builder, ws, sec, los[sec.table] is not None)

    _write_helpers(builder, ws, los[PROFILE.name] is not None, los[TARGET_TREND.name] is not None and los[INFO.name] is not None)
    charts = ((INVESTOR, _chart_price_flow), (TRADESIZE, _chart_tradesize), (PROFILE, _chart_profile),
              (ESTIMATE, _chart_estimate), (CSL, _chart_csl), (TARGET_TREND, _chart_trend), (FIN, _chart_fin))
    for t, fn in charts:
        if los[t.name] is not None:
            _com_ok(builder, f"{t.name} 차트", lambda f=fn, t=t: f(builder, ws, geos[t.name], t))

    ws.Activate()
    win = ws.Application.ActiveWindow
    win.ScrollRow = 1          # 조건부 서식을 만들며 아래쪽 칸을 선택했으므로 맨 위로 돌려야 틀 고정이 1~9행에 걸린다
    win.ScrollColumn = 1
    freeze(ws, FROZEN_ROWS)
    ws.Range(INPUT_CELL).Select()
    builder.say("종목분석 시트 완료")
