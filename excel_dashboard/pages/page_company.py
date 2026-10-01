"""대회종목 페이지 빌더 — Q.Pack 'Company' 시트형 (spec R16, 2026-10-01 T17).

시트 `대회종목` 한 장에 `tblCompany`(쿼리 T_Company, 103열, 행 = 페이지 행)를 Q.Pack처럼 보여 준다.
표는 Power Query가 채우므로 이 모듈은 값을 쓰지 않고 서식·조건부 서식·열 묶음·상단(버튼·최근 조회 칸·범례)만 만든다.

시트 배치 (행)
   2~3  제목 띠(제목·부제, 오른쪽에 기준 시각·데이터 상태)        4  메뉴 줄(builder.nav_links)
   6    [시세] 버튼 · 시세 최근 조회 · `시세_최근조회`(D6:F6 병합) · `시세_상태`(G6) · 오른쪽 O6 = 가격 기준일·종목 수·
        (상태 열의 첫 '이전 데이터'/'오류' 사유 — 상태 열은 기본으로 접혀 있으므로)
   7    [전체] 버튼 · 전체 최근 조회 · `전체_최근조회`(D7:F7 병합) · `전체_상태`(G7) · 오른쪽 O7 = 색 범례
   9~11 숨김 행 — Q.Pack 기준점(낮음·중앙·높음). 색을 칠하는 열마다 그 열 바로 위 칸, 줄무늬는 D-19 열 위 칸
   12   열 묶음 머리글(16개 묶음 이름, 묶음마다 병합)              13  표 머리글 (B13 = LOADS 셀)   14~ 표 본문

인터페이스 (배선 작업 T23이 지켜야 할 것)
  SHEET           시트 이름. 빌더 SHEETS에 넣고 builder.ws[SHEET]로 접근할 수 있어야 한다.
  LOADS           [(쿼리, 표 이름, 머리글 왼쪽 위 셀)] = [("T_Company", "tblCompany", "B13")]. 다른 위치에 적재하면
                  build()는 경고만 남기고 서식을 입히지 않는다(위쪽 12행이 이 페이지의 머리 부분).
  prepare_tables  표를 적재한 직후·첫 새로 고침 전에 부른다 — QueryTable 새로 고침 방식을 덮어쓰기(RefreshStyle 0)로,
                  서식 유지·열 너비 고정을 켠다. build()도 같은 설정을 다시 건다. 이 값을 되돌리지 말 것.
  build(builder)  builder.wb / ws / lo / table_style / say / nav_links / HEADER_RIGHT / HEADER_STATUS만 쓴다.
                  표의 행을 더하거나 지우지 않는다(ListRows.Add/Delete 없음 — 0행 표도 그대로 둔다).
                  표가 없으면 경고만 남기고 상단(버튼·이름 정의·범례)만 만든다.
  이름 정의       통합문서 범위 `시세_최근조회`=D6, `시세_상태`=G6, `전체_최근조회`=D7, `전체_상태`=G7 (대회종목 시트).
                  VBA(mod_refresh)가 *_최근조회에 종료 시각, *_상태에 `정상`/`일부 오류 n건`/`실패: …`을 쓴다.
  버튼            도형 btn_RefreshQuick(글자 "시세", 매크로 RefreshQuick), btn_RefreshFull("전체", RefreshFull) — B6·B7.

설계 이유
  - 서식·조건부 서식은 표 열(DataBodyRange)이 아니라 첫 데이터 행부터 BODY_ROWS행의 고정 시트 범위에 건다. 0행 표에는
    본문 범위가 없고(임시 행 추가 금지), 덮어쓰기 새로 고침은 셀을 끼워 넣지 않으므로 고정 범위의 숫자 서식·규칙이
    표가 0 → 19 → 528행으로 늘고 줄어도 그대로 남는다(실측). 빈 칸에는 칠할 값이 없어 표 아래는 보이지 않는다.
  - Q.Pack 3색은 열마다 따로, 기준점 = 유니버스 = 대회 행만의 최소·중앙값·최대(xl_helpers.cf_qpack 기준점 셀 방식).
    유니버스 밖 행은 기준 계산에서 빠지고 같은 척도로 칠해진다(기준점 밖이면 끝 색). 공란은 칠하지 않는다.
    최근 20일 줄무늬는 20칸 전체를 한 척도로(대회 행 20칸 값 전체의 5·50·95 백분위, QPACK_STRIP_PCTS), 숫자는 숨긴다.
  - 규칙 우선순위: Q.Pack 척도(맨 앞, 헬퍼가 SetFirstPriority) > 보유 행 옅은 파랑 등 행 규칙. 보유 행에서도 색 칸은
    Q.Pack 색이 보이고 공란·다른 열만 보유색이 된다. 규칙은 약 50개(열마다 1개 + 행 규칙) — 전체 너비 규칙을 많이 두면
    새로 고침 중 Excel이 멈췄다(뉴스·이벤트 페이지 실측).
  - 열 묶음(+/−): 같은 수준의 묶음이 붙어 있으면 Excel이 하나로 합치므로, 접는 묶음마다 바로 왼쪽에 보이는 열(요약 열)을
    둔다. spec이 통째로 접으라는 실적·밸류 변화·신용/공매도/대차·이벤트·분류 세부·상태 묶음은 서로 붙어 있어 각 묶음의
    첫 열(최근분기·PER변화3M·신용잔고율·다음이벤트·NICS 대분류·유의)을 요약 열로 남기고 나머지를 접는다.
  - 메뉴 줄: nav_links는 2열 간격으로 링크를 놓는데 이 시트는 접히는 열·폭 2.5의 줄무늬 열이 있어 링크가 숨거나 잘린다.
    nav_links가 만든 링크(목록은 빌더 소관)를 읽어, 링크마다 글자 폭만큼의 '접히지 않는 열'을 병합한 칸에 다시 놓는다
    (누를 수 있는 범위 = 글자 전체, 메뉴 15개도 줄무늬 열까지 안에 들어감).
  - 표의 열별 필터 단추는 숨길 수 없다(표에서는 VisibleDropDown이 무시됨 — 실측). 줄무늬 머리글은 6pt 글자.
"""
from __future__ import annotations

import dataclasses
from typing import Any

import pywintypes

from xl_helpers import (
    C, NF, QPACK_HIGH, QPACK_LOW, QPACK_MID, QPACK_STRIP_PCTS, XL_CENTER, XL_LEFT, XL_RIGHT, add_button,
    add_names, cf_expr, cf_qpack, cf_qpack_block, fill, group_columns, hyperlink, put, rgb, set_nf, sheet_setup, style,
    title_bar,
)

__all__ = ["SHEET", "LOADS", "prepare_tables", "build"]

SHEET = "대회종목"
TABLE = "tblCompany"
QUERY = "T_Company"
HEADER_ROW = 13
TABLE_CELL = "B13"
LOADS = [(QUERY, TABLE, TABLE_CELL)]

NAV_ROW = 4
CTRL_ROWS = (6, 7)               # [시세] 줄, [전체] 줄
ANCHOR_ROWS = (9, 10, 11)        # 숨김 — Q.Pack 기준점 낮음·중앙·높음 (같은 열 바로 위)
BAND_ROW = 12                    # 열 묶음 머리글
BODY_ROWS = 1200                 # 서식·조건부 서식을 거는 본문 행 수(표가 이보다 길면 표 길이 + 여유)
STRIP_WIDTH = 2.5
XL_OVERWRITE_CELLS = 0           # QueryTable.RefreshStyle: 셀을 끼워 넣거나 지우지 않고 덮어씀

# ------------------------------------------------------------------------------------------------ 열 계약 (T_Company)
PERIODS = ["1D", "1W", "1M", "3M", "6M", "YTD", "1Y"]
STRIP = [f"D-{k}" for k in range(19, 0, -1)] + ["D0"]
COLUMNS = (["종목코드", "종목명", "시장", "구분", "유니버스", "NICS 업종", "대테마",
            "시가총액억", "거래대금60일억", "거래량60일천주", "현재가", "고52주대비", "가격기준일"]
           + PERIODS + STRIP
           + ["외국인1D%", "외국인1W%", "외국인1M%", "외국인1D억", "외국인1W억", "외국인1M억",
              "기관1D%", "기관1W%", "기관1M%", "기관1D억", "기관1W억", "기관1M억", "수급기준일",
              "후행PER", "PBR",
              "FwdPER", "FwdPER변화1W", "FwdPER변화1M", "FwdEPS", "추정일", "FY1", "FY1_EPS", "FY2", "FY2_EPS",
              "목표가평균", "괴리율", "증권사수", "목표가변화1M", "목표가변화3M", "최근의견일", "최근증권사", "최근의견",
              "최근분기", "매출YoY", "영업이익YoY", "순이익YoY", "실적상태", "ROE",
              "PER변화3M", "PER변화6M", "PER변화YTD", "PER변화1Y", "PBR변화3M", "PBR변화6M", "PBR변화YTD", "PBR변화1Y",
              "신용잔고율", "신용잔고율1M변화", "공매도비중5일", "대차잔고1M변화율", "신용기준일",
              "다음이벤트", "다음이벤트일",
              "NICS 대분류", "NICS 세부", "세부테마", "KSIC 세분류", "주요상품", "기업집단", "분류출처",
              "유의", "상태", "조회시각"])

# 열 묶음 머리글 (이름, 첫 열, 끝 열) — spec R16 순서
BANDS = [
    ("기본", "종목코드", "대테마"),
    ("규모/유동성", "시가총액억", "거래량60일천주"),
    ("가격", "현재가", "가격기준일"),
    ("Price change (%)", "1D", "1Y"),
    ("최근 20일", "D-19", "D0"),
    ("외국인 순매수(시총 대비 %)", "외국인1D%", "외국인1M억"),
    ("기관 순매수(시총 대비 %)", "기관1D%", "수급기준일"),
    ("밸류", "후행PER", "PBR"),
    ("Fwd PER(KIS 추정)", "FwdPER", "FY2_EPS"),
    ("목표주가", "목표가평균", "최근의견"),
    ("실적", "최근분기", "ROE"),
    ("밸류 변화", "PER변화3M", "PBR변화1Y"),
    ("신용/공매도/대차", "신용잔고율", "신용기준일"),
    ("이벤트", "다음이벤트", "다음이벤트일"),
    ("분류 세부", "NICS 대분류", "분류출처"),
    ("상태", "유의", "조회시각"),
]

# 기본으로 접어 두는 묶음 (첫 열, 끝 열) — 각 묶음 바로 왼쪽 열은 보이는 요약 열([+]/[−] 단추 자리)
COLLAPSED_GROUPS = [
    ("거래량60일천주", "거래량60일천주"),     # 유동성 상세: 60일 평균 거래량
    ("고52주대비", "가격기준일"),             # 52주 고가 대비 + 가격 기준일(상단에 표시)
    ("외국인1D억", "외국인1M억"),             # 수급 억원 금액
    ("기관1D억", "수급기준일"),               # 수급 억원 금액 + 수급 기준일
    ("PBR", "PBR"),                           # 밸류 변화 묶음의 PBR 수준
    ("FwdEPS", "FY2_EPS"),                    # KIS 추정 상세
    ("목표가변화1M", "최근의견"),             # 목표주가 상세
    ("매출YoY", "ROE"),                       # 실적 (요약 열 최근분기)
    ("PER변화6M", "PBR변화1Y"),               # 밸류 변화 (요약 열 PER변화3M)
    ("신용잔고율1M변화", "신용기준일"),       # 신용/공매도/대차 (요약 열 신용잔고율)
    ("다음이벤트일", "다음이벤트일"),         # 이벤트 (요약 열 다음이벤트)
    ("NICS 세부", "분류출처"),                # 분류 세부 (요약 열 NICS 대분류)
    ("상태", "조회시각"),                     # 상태 (요약 열 유의)
]

# Q.Pack 3색을 칠하는 변화·비율 열 (spec R16: 34열) — 수준 값(현재가·시가총액·거래대금·PER·PBR 등)은 색 없음
QPACK_COLUMNS = (PERIODS + ["고52주대비", "외국인1D%", "외국인1W%", "외국인1M%", "기관1D%", "기관1W%", "기관1M%",
                            "매출YoY", "영업이익YoY", "순이익YoY",
                            "PER변화3M", "PER변화6M", "PER변화YTD", "PER변화1Y",
                            "PBR변화3M", "PBR변화6M", "PBR변화YTD", "PBR변화1Y",
                            "괴리율", "목표가변화1M", "목표가변화3M", "FwdPER변화1W", "FwdPER변화1M",
                            "신용잔고율", "신용잔고율1M변화", "공매도비중5일", "대차잔고1M변화율"])

PCT1_SIGNED = "+0.0%;-0.0%;0.0%"
PCT2_SIGNED = "+0.00%;-0.00%;0.00%"
NUM_SIGNED = "+#,##0;-#,##0;0"
NF_TIME = "mm-dd hh:mm"

# 열 서식: 너비·숫자 서식·가로 정렬·글꼴. 금액 억원(소수 없음)·가격 원·비율 %(부호)·PER 0.0·PBR 0.00·날짜 yyyy-mm-dd.
# 등락 열 글꼴은 중립색(색은 Q.Pack 칸 채우기로만).
_W, _NF, _H = "width", "nf", "h"
SPECS: dict[str, dict] = {
    "종목코드": {_W: 7.5, _H: XL_CENTER, "color": C["muted"]},
    "종목명": {_W: 15, "bold": True},
    "시장": {_W: 7.5, _H: XL_CENTER, "color": C["muted"]},
    "구분": {_W: 8, _H: XL_CENTER},
    "유니버스": {_W: 9.5, _H: XL_CENTER, "color": C["muted"]},
    "NICS 업종": {_W: 17},
    "대테마": {_W: 11},
    "시가총액억": {_W: 10, _NF: NF["num0"]},
    "거래대금60일억": {_W: 9, _NF: NF["num0"]},
    "거래량60일천주": {_W: 9, _NF: NF["num0"]},
    "현재가": {_W: 9.5, _NF: NF["krw"]},
    "고52주대비": {_W: 7.5, _NF: NF["pct1"]},
    "가격기준일": {_W: 10, _NF: NF["date"], _H: XL_CENTER},
    **{p: {_W: 7, _NF: PCT1_SIGNED} for p in PERIODS},
    **{f"{who}{p}%": {_W: 7.5, _NF: PCT2_SIGNED} for who in ("외국인", "기관") for p in ("1D", "1W", "1M")},
    **{f"{who}{p}억": {_W: 8.5, _NF: NUM_SIGNED} for who in ("외국인", "기관") for p in ("1D", "1W", "1M")},
    "수급기준일": {_W: 10, _NF: NF["date"], _H: XL_CENTER},
    "후행PER": {_W: 7, _NF: "0.0"},
    "PBR": {_W: 6.5, _NF: "0.00"},
    "FwdPER": {_W: 7, _NF: "0.0"},
    "FwdPER변화1W": {_W: 7.5, _NF: PCT1_SIGNED},
    "FwdPER변화1M": {_W: 7.5, _NF: PCT1_SIGNED},
    "FwdEPS": {_W: 9, _NF: NF["num0"]},
    "추정일": {_W: 10, _NF: NF["date"], _H: XL_CENTER},
    "FY1": {_W: 8.5, _H: XL_CENTER},
    "FY1_EPS": {_W: 9, _NF: NF["num0"]},
    "FY2": {_W: 8.5, _H: XL_CENTER},
    "FY2_EPS": {_W: 9, _NF: NF["num0"]},
    "목표가평균": {_W: 10, _NF: NF["krw"]},
    "괴리율": {_W: 7.5, _NF: PCT1_SIGNED},
    "증권사수": {_W: 7, _NF: "0", _H: XL_CENTER},
    "목표가변화1M": {_W: 7.5, _NF: PCT1_SIGNED},
    "목표가변화3M": {_W: 7.5, _NF: PCT1_SIGNED},
    "최근의견일": {_W: 10, _NF: NF["date"], _H: XL_CENTER},
    "최근증권사": {_W: 9},
    "최근의견": {_W: 10, _H: XL_CENTER},
    "최근분기": {_W: 8, _H: XL_CENTER},
    "매출YoY": {_W: 8.5, _NF: PCT1_SIGNED},
    "영업이익YoY": {_W: 8.5, _NF: PCT1_SIGNED},
    "순이익YoY": {_W: 8.5, _NF: PCT1_SIGNED},
    "실적상태": {_W: 8, _H: XL_CENTER},
    "ROE": {_W: 7, _NF: NF["pct1"]},
    **{f"{k}변화{p}": {_W: 7.5, _NF: PCT1_SIGNED} for k in ("PER", "PBR") for p in ("3M", "6M", "YTD", "1Y")},
    "신용잔고율": {_W: 8.5, _NF: NF["pct"]},
    "신용잔고율1M변화": {_W: 8, _NF: PCT2_SIGNED},
    "공매도비중5일": {_W: 8, _NF: NF["pct1"]},
    "대차잔고1M변화율": {_W: 8, _NF: PCT1_SIGNED},
    "신용기준일": {_W: 10, _NF: NF["date"], _H: XL_CENTER},
    "다음이벤트": {_W: 20, "color": C["navy"]},
    "다음이벤트일": {_W: 10, _NF: NF["date"], _H: XL_CENTER},
    "NICS 대분류": {_W: 10},
    "NICS 세부": {_W: 16},
    "세부테마": {_W: 12},
    "KSIC 세분류": {_W: 20},
    "주요상품": {_W: 24},
    "기업집단": {_W: 12},
    "분류출처": {_W: 7, _H: XL_CENTER, "color": C["muted"]},
    "유의": {_W: 9, "color": C["up"], "bold": True},
    "상태": {_W: 12, "color": C["muted"]},
    "조회시각": {_W: 11, _NF: NF_TIME, _H: XL_CENTER, "color": C["muted"]},
}
for _s in STRIP:
    SPECS[_s] = {_W: STRIP_WIDTH}

HELD_FILL = "#F3F8FF"        # 시세판·뉴스 페이지의 보유 행 강조와 같은 색
PAST_FONT = "#9AA5B1"
BAND_FILLS = ("#1B3358", "#2B4A78")
BAND_LINE = "#FFFFFF"
SEPARATOR = "#8FA3BF"        # 대회 → 유니버스 밖 경계선
NAV_GAP = 1.0                # 메뉴 링크 글자 폭에 더하는 여백(열 너비 단위)

# 상단 배치에 쓰는 열(표 열 이름 기준 — 모두 접히지 않는 열)
BTN_COL, LABEL_COL, TIME_FIRST, TIME_LAST, STATUS_COL = "종목코드", "종목명", "시장", "유니버스", "NICS 업종"
INFO_COL, LEGEND_COLS, NOTE_COL, RIGHT_COL = "1D", ("1D", "1W", "1M"), "3M", "1Y"
BUTTONS = [  # (행, 글자, 매크로, 라벨, 시각 이름, 상태 이름, 색)
    (CTRL_ROWS[0], "시세", "RefreshQuick", "시세 최근 조회", "시세_최근조회", "시세_상태", C["accent"]),
    (CTRL_ROWS[1], "전체", "RefreshFull", "전체 최근 조회", "전체_최근조회", "전체_상태", C["navy2"]),
]
STATUS_INITIAL = "아직 실행 안 함"


# ------------------------------------------------------------------------------------------------ 위치
def _letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


@dataclasses.dataclass
class _Geo:
    """표의 시트 위치. cols는 실제 표 열(표가 계약과 다르면 있는 열만) 또는 표가 없을 때 계약 순서로 계산한 위치."""

    header_row: int
    cols: dict[str, int]
    exact: bool          # 표 열이 계약(COLUMNS)과 이름·순서까지 같음 → 묶음·머리띠·줄무늬 가능
    have: set[str]       # 서식을 입힐 실제 표 열
    nrows: int

    @property
    def first_row(self) -> int:
        return self.header_row + 1

    @property
    def last_body_row(self) -> int:
        return self.first_row + max(BODY_ROWS, self.nrows + 200) - 1

    @property
    def left(self) -> int:
        return min(self.cols.values())

    @property
    def right(self) -> int:
        return max(self.cols.values())


def _contract_cols(cell: str = TABLE_CELL) -> dict[str, int]:
    from_col = _col_index(cell)
    return {name: from_col + j for j, name in enumerate(COLUMNS)}


def _col_index(cell: str) -> int:
    letters = "".join(ch for ch in cell if ch.isalpha()).upper()
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _geometry(lo: Any) -> _Geo:
    hdr = lo.HeaderRowRange
    names = [str(c.Name) for c in lo.ListColumns]
    actual = {str(c.Name): int(c.Range.Column) for c in lo.ListColumns}
    exact = names == COLUMNS
    nrows = 0 if lo.DataBodyRange is None else int(lo.ListRows.Count)
    have = {n for n in COLUMNS if n in actual}
    cols = actual if exact else {**_contract_cols(), **{n: actual[n] for n in have}}
    return _Geo(int(hdr.Row), cols, exact, have, nrows)


def _hidden_by_default(geo: _Geo) -> set[int]:
    out: set[int] = set()
    for first, last in COLLAPSED_GROUPS:
        if first in geo.cols and last in geo.cols:
            out.update(range(geo.cols[first], geo.cols[last] + 1))
    return out


# ------------------------------------------------------------------------------------------------ 표·시트
def _sheet(builder: Any) -> Any:
    try:
        return builder.ws[SHEET]
    except KeyError:
        builder.say(f"[{SHEET}] 경고: 시트가 없어 페이지를 만들지 않습니다(빌더 SHEETS에 '{SHEET}' 추가 필요)")
        return None


def _table(builder: Any) -> Any:
    try:
        return builder.lo[TABLE]
    except KeyError:
        builder.say(f"[{SHEET}] 경고: 표 {TABLE}이(가) 없어 표 서식·색·열 묶음을 건너뜁니다"
                    f"(LOADS 위치 {TABLE_CELL}에 {QUERY}를 적재한 뒤 호출해야 합니다)")
        return None


def _overwrite_on_refresh(lo: Any) -> None:
    """새로 고침이 셀을 끼워 넣거나 지우지 않고 덮어쓰게(xlOverwriteCells), 서식 유지·열 너비 고정.

    기본값(셀 삽입·삭제)에서는 표가 0·1행에서 늘어날 때 조건부 서식이 새 행에 이어지지 않았고, 0행 새로 고침이 셀을 지워
    아래쪽이 밀렸다(실측). 열 너비 자동 맞춤은 폭 2.5의 줄무늬 열을 넓혀 버린다. 정적 표(쿼리 아님)는 건너뛴다.
    """
    try:
        qt = lo.QueryTable
    except pywintypes.com_error:
        return
    for attr, value in (("RefreshStyle", XL_OVERWRITE_CELLS), ("PreserveFormatting", True), ("AdjustColumnWidth", False)):
        try:
            setattr(qt, attr, value)
        except pywintypes.com_error:
            pass


def prepare_tables(builder: Any) -> None:
    """표를 적재한 직후·첫 새로 고침 전에 부른다(배선 T23): tblCompany의 새로 고침 방식을 덮어쓰기로 바꾼다.

    Args:
        builder: build_dashboard.Builder 또는 하네스 대역. lo·say만 쓴다.

    Example:
        page_company.prepare_tables(builder)   # load_query 뒤, refresh 전
    """
    lo = _table(builder)
    if lo is not None:
        _overwrite_on_refresh(lo)


def _reset_sheet(ws: Any, hr: int, last_col: int) -> None:
    """다시 불러도 겹치지 않게: 조건부 서식·열 묶음·숨김·머리 부분(표 위 행)의 병합·내용·링크를 지운다(이 시트 전용)."""
    ws.Cells.FormatConditions.Delete()
    ws.Columns.Hidden = False
    # 열 묶음 지우기: 이 시트(표가 있음)에서는 Range.ClearOutline이 모든 형태로 거부됨(실측) → 열마다 Ungroup
    for c in range(1, last_col + 3):
        col = ws.Columns(c)
        for _ in range(8):                   # 묶음 수준은 최대 8
            if int(col.OutlineLevel) <= 1:
                break
            col.Ungroup()
    top = ws.Range(ws.Cells(1, 1), ws.Cells(hr - 1, last_col + 2))
    top.EntireRow.Hidden = False
    try:
        top.Hyperlinks.Delete()
    except pywintypes.com_error:
        pass
    top.UnMerge()
    top.Clear()


# ------------------------------------------------------------------------------------------------ 열 서식
def _body(ws: Any, geo: _Geo, first: str, last: str | None = None) -> Any:
    return ws.Range(ws.Cells(geo.first_row, geo.cols[first]), ws.Cells(geo.last_body_row, geo.cols[last or first]))


def _set_widths(ws: Any, geo: _Geo) -> None:
    """열 너비. 표가 있으면 실제 표 열만(계약과 다른 표에서 엉뚱한 열을 바꾸지 않게), 없으면 계약 위치(상단 배치용)."""
    names = geo.have if geo.have else set(geo.cols)
    for name, sp in SPECS.items():
        if name in names and _W in sp:
            ws.Columns(geo.cols[name]).ColumnWidth = sp[_W]


def _format_columns(builder: Any, ws: Any, geo: _Geo) -> None:
    """본문 고정 범위의 숫자 서식·정렬·글꼴과 머리글 행. 계약에 없는 열 이름이 빠졌으면 경고하고 있는 열만 입힌다."""
    missing = [n for n in COLUMNS if n not in geo.have]
    if missing:
        builder.say(f"[{SHEET}] 경고: {TABLE}에 없는 열 {len(missing)}개({', '.join(missing[:6])}…) — 그 열의 서식·색은 건너뜁니다")
    if not geo.have:
        return
    whole = ws.Range(ws.Cells(geo.first_row, geo.left), ws.Cells(geo.last_body_row, geo.right))
    whole.Font.Size = 9
    for name in COLUMNS:
        if name not in geo.have:
            continue
        sp = SPECS.get(name, {})
        rng = _body(ws, geo, name)
        if _NF in sp:
            set_nf(rng, sp[_NF])
        if _H in sp:
            rng.HorizontalAlignment = sp[_H]
        if any(k in sp for k in ("bold", "color")):
            style(rng, bold=sp.get("bold"), color=sp.get("color"))
    hdr = ws.Range(ws.Cells(geo.header_row, geo.left), ws.Cells(geo.header_row, geo.right))
    style(hdr, size=9, h=XL_CENTER, wrap=True)
    ws.Rows(geo.header_row).RowHeight = 30
    if geo.exact:
        strip_hdr = ws.Range(ws.Cells(geo.header_row, geo.cols[STRIP[0]]), ws.Cells(geo.header_row, geo.cols[STRIP[-1]]))
        style(strip_hdr, size=6, bold=False, wrap=False)
        for _, first, _ in BANDS[1:]:          # 묶음 경계(머리글 행) 흰 세로선
            b = ws.Cells(geo.header_row, geo.cols[first]).Borders(7)   # xlEdgeLeft
            b.LineStyle = 1
            b.Weight = 2
            b.Color = rgb(BAND_LINE)


# ------------------------------------------------------------------------------------------------ 상단
def _text_width(text: str, size: float = 9.0) -> float:
    """글자 폭 추정(열 너비 단위) — 한글 약 1.75, 영문·숫자 약 0.95, 공백 0.5 (10pt 기준)."""
    w = 0.0
    for ch in text:
        o = ord(ch)
        if 0xAC00 <= o <= 0xD7A3 or 0x3130 <= o <= 0x318F or 0x4E00 <= o <= 0x9FFF:
            w += 1.75
        elif ch == " ":
            w += 0.5
        elif o < 128:
            w += 0.95
        else:
            w += 1.0
    return w * size / 10.0


def _nav_row(builder: Any, ws: Any, geo: _Geo, last_col: int) -> None:
    """메뉴 줄: builder.nav_links가 만든 링크(목록은 빌더 소관)를 읽어, 링크마다 글자 폭만큼의 접히지 않는 열을 병합한 칸에
    차례로 다시 놓는다. 자리가 모자라면 남은 링크를 빼고 경고한다."""
    row = NAV_ROW
    band = ws.Range(ws.Cells(row, 2), ws.Cells(row, last_col))
    fill(band, "navy2")
    ws.Rows(row).RowHeight = 16
    try:
        builder.nav_links(ws, row)
    except Exception as e:  # noqa: BLE001 — 메뉴는 부가 기능: 실패해도 페이지는 계속
        builder.say(f"[{SHEET}] 경고: 메뉴 줄을 만들지 못했습니다({type(e).__name__})")
        return
    links = []
    for hl in list(ws.Hyperlinks):
        try:
            cell = hl.Range
            if int(cell.Row) != row:
                continue
            text = str(cell.Value if cell.Value is not None else hl.TextToDisplay)
            links.append((int(cell.Column), text, str(hl.SubAddress or ""), str(hl.Address or "")))
        except pywintypes.com_error:
            continue
    if not links:
        return
    links.sort()
    row_rng = ws.Range(ws.Cells(row, 1), ws.Cells(row, last_col))
    row_rng.Hyperlinks.Delete()
    row_rng.ClearContents()
    fill(band, "navy2")
    hidden = _hidden_by_default(geo)
    cols = [c for c in range(2, last_col + 1) if c not in hidden]
    widths: dict[int, float] = {}

    def width(c: int) -> float:
        if c not in widths:
            widths[c] = float(ws.Columns(c).ColumnWidth)
        return widths[c]

    k, placed = 0, 0
    for _, text, sub, addr in links:
        need, j, acc = _text_width(text) + NAV_GAP, k, 0.0
        while j < len(cols) and acc < need:
            acc += width(cols[j])
            j += 1
        if acc < need:
            break
        target = ws.Cells(row, cols[k])
        if cols[j - 1] > cols[k]:   # 링크 한 칸 = 글자 폭만큼의 보이는 열 병합(누를 수 있는 범위 = 글자 전체, 줄무늬 열도 사용)
            ws.Range(target, ws.Cells(row, cols[j - 1])).Merge()
        if sub:
            hyperlink(ws, target.Address, sub, text)
        else:
            target.Value = text
            ws.Hyperlinks.Add(target, addr, "", "", text)
            style(target, color="#B8C4D6", size=9)
            target.Font.Underline = False
        target.HorizontalAlignment = XL_LEFT
        placed += 1
        k = j
    if placed < len(links):
        builder.say(f"[{SHEET}] 경고: 메뉴 링크 {len(links) - placed}개를 놓을 자리가 없어 뺐습니다")


def _title(builder: Any, ws: Any, geo: _Geo, last_col: int) -> None:
    subtitle = ("대회 종목(시가총액 큰 순) 아래에 유니버스 밖 보유·관심 종목 | 색 = Q.Pack 3색(열마다 대회 종목 기준) · "
                "열 위 [+]로 세부 열 펼치기 · [시세]·[전체] 버튼으로 갱신")
    span = f"B2:{_letter(last_col)}3"
    rc = _letter(geo.cols[RIGHT_COL])
    right = getattr(builder, "HEADER_RIGHT", None)
    status = getattr(builder, "HEADER_STATUS", None)
    try:
        title_bar(ws, "대회종목 (Q.Pack형)", subtitle, span, right_formula=right, right_cell=f"{rc}2",
                  right2_formula=status, right2_cell=f"{rc}3")
    except pywintypes.com_error:
        builder.say(f"[{SHEET}] 경고: 머리글 기준 시각·상태 수식을 넣지 못했습니다(참조 표 없음) — 제목만 표시")
        for a in (f"{rc}2", f"{rc}3"):
            ws.Range(a).ClearContents()
        title_bar(ws, "대회종목 (Q.Pack형)", subtitle, span)


def _info_formula(t: str = TABLE) -> str:
    return (f'=LET(cp_d,IF(COUNT({t}[가격기준일])=0,"-",TEXT(MAX({t}[가격기준일]),"yyyy-mm-dd")),'
            f'cp_c,COUNTIFS({t}[유니버스],"대회"),cp_o,COUNTIFS({t}[유니버스],"유니버스 밖"),'
            f'cp_e,COUNTIF({t}[상태],"오류*"),cp_p,COUNTIF({t}[상태],"이전*"),'
            f'"가격 기준일 "&cp_d&"   ·   대회 "&cp_c&"종목   ·   유니버스 밖 "&cp_o&"종목"'
            f'&IF(cp_p>0,"   ▲ "&LEFT(INDEX({t}[상태],MATCH("이전*",{t}[상태],0)),80),'
            f'IF(cp_e>0,"   ▲ 오류 "&cp_e&"행 — "&LEFT(INDEX({t}[상태],MATCH("오류*",{t}[상태],0)),60),"")))')


def _controls(builder: Any, ws: Any, geo: _Geo, has_table: bool) -> None:
    """[시세]·[전체] 버튼, 최근 조회 시각·상태 칸(이름 정의), 가격 기준일·종목 수, 색 범례."""
    for r in CTRL_ROWS:
        ws.Rows(r).RowHeight = 22
    ws.Rows(CTRL_ROWS[0] - 1).RowHeight = 6
    ws.Rows(CTRL_ROWS[1] + 1).RowHeight = 6
    t0, t1, sc = geo.cols[TIME_FIRST], geo.cols[TIME_LAST], geo.cols[STATUS_COL]
    names = {}
    for row, caption, macro, label, time_name, status_name, color in BUTTONS:
        add_button(ws, ws.Cells(row, geo.cols[BTN_COL]).Address, caption, macro, fill_color=color, size=10)
        put(ws, ws.Cells(row, geo.cols[LABEL_COL]).Address, label, size=9, color=C["muted"], h=XL_RIGHT)
        tcell = ws.Range(ws.Cells(row, t0), ws.Cells(row, t1))
        tcell.Merge()
        style(tcell, nf=NF["dt"], bold=True, size=10, color=C["text"], h=XL_CENTER)
        put(ws, ws.Cells(row, sc).Address, STATUS_INITIAL, size=9, color=C["muted"], h=XL_LEFT, indent=1)
        names[time_name] = f"='{SHEET}'!${_letter(t0)}${row}"
        names[status_name] = f"='{SHEET}'!${_letter(sc)}${row}"
    add_names(builder.wb, names)
    st = ws.Range(ws.Cells(CTRL_ROWS[0], sc), ws.Cells(CTRL_ROWS[1], sc))
    a = st.Cells(1, 1).Address.replace("$", "")
    cf_expr(st, f'=LEFT({a},2)="실패"', font=C["up"], bold=True)
    cf_expr(st, f'=LEFT({a},2)="일부"', font=C["warn"], bold=True)
    cf_expr(st, f'={a}="정상"', font=C["good"], bold=True)
    # 가격 기준일·종목 수·경고 (O6), 색 범례 (O7~)
    info = ws.Cells(CTRL_ROWS[0], geo.cols[INFO_COL])
    if has_table:
        try:
            put(ws, info.Address, formula=_info_formula(), bold=True, size=10, color=C["navy"], h=XL_LEFT)
            cf_expr(info, f'=ISNUMBER(SEARCH("▲",{info.Address.replace("$", "")}))', font=C["up"])
        except pywintypes.com_error:
            builder.say(f"[{SHEET}] 경고: 가격 기준일·종목 수 수식을 넣지 못했습니다(표 열 구성 확인)")
    else:
        put(ws, info.Address, f"표 {TABLE} 없음 — 빌더 배선 확인", bold=True, size=10, color=C["up"], h=XL_LEFT)
    for name, (text, color) in zip(LEGEND_COLS, (("낮음", QPACK_LOW), ("중앙", QPACK_MID), ("높음", QPACK_HIGH))):
        put(ws, ws.Cells(CTRL_ROWS[1], geo.cols[name]).Address, text, size=8, bold=True, color=C["text"], fill=color,
            h=XL_CENTER)
    pct = "·".join(str(p) for p in QPACK_STRIP_PCTS)
    note = ("← 열마다 대회 종목의 최소·중앙값·최대 기준(유니버스 밖 행은 기준에서 빼고 같은 척도로 칠함) · "
            f"최근 20일 = 20칸 전체를 한 척도로({pct} 백분위) · 빈칸 = 데이터 없음 · 옅은 파랑 행 = 보유")
    put(ws, ws.Cells(CTRL_ROWS[1], geo.cols[NOTE_COL]).Address, note, size=8, color=C["muted"], h=XL_LEFT, indent=1)


def _band_row(ws: Any, geo: _Geo) -> None:
    """열 묶음 머리글: 묶음마다 병합·줄바꿈·가운데(접히면 보이는 폭에 맞춰 두 줄로). 첫 묶음은 왼쪽 정렬(틀 고정 경계)."""
    r = geo.header_row - 1
    ws.Rows(r).RowHeight = 26
    for i, (label, first, last) in enumerate(BANDS):
        rng = ws.Range(ws.Cells(r, geo.cols[first]), ws.Cells(r, geo.cols[last]))
        if rng.Columns.Count > 1:
            rng.Merge()
        rng.Cells(1, 1).Value = label
        style(rng, bold=True, size=9, color="#FFFFFF", fill=BAND_FILLS[i % 2], h=XL_LEFT if i == 0 else XL_CENTER,
              wrap=True, indent=1 if i == 0 else None)
        b = rng.Borders(7)   # xlEdgeLeft
        b.LineStyle = 1
        b.Weight = 2
        b.Color = rgb(BAND_LINE)


# ------------------------------------------------------------------------------------------------ 조건부 서식
def _cf_border_top(rng: Any, formula: str, color: str) -> Any:
    fc = cf_expr(rng, formula)
    b = fc.Borders(-4160)   # xlTop
    b.LineStyle = 1
    b.Color = rgb(color)
    return fc


def _row_rules(builder: Any, ws: Any, geo: _Geo) -> None:
    """행 규칙(Q.Pack 척도보다 먼저 추가 → 척도가 앞 순위로 올라가 색 칸에서는 Q.Pack이 보임)."""
    need = [c for c in ("구분", "유니버스", "종목명") if c not in geo.have]
    if need:
        builder.say(f"[{SHEET}] 경고: {TABLE}에 열 {', '.join(need)}이(가) 없어 보유 강조·경계선을 건너뜁니다")
        return
    r0 = geo.first_row
    whole = ws.Range(ws.Cells(r0, geo.left), ws.Cells(geo.last_body_row, geo.right))
    g = f"${_letter(geo.cols['구분'])}{r0}"
    u = f"${_letter(geo.cols['유니버스'])}{r0}"
    u_prev = f"${_letter(geo.cols['유니버스'])}{r0 - 1}"
    cf_expr(whole, f'=LEFT({g},2)="보유"', fill_color=HELD_FILL)
    cf_expr(_body(ws, geo, "종목명"), f'=LEFT({g},2)="보유"', font=C["accent"], bold=True)
    cf_expr(_body(ws, geo, "유니버스"), f'={u}="유니버스 밖"', font=C["warn"], bold=True)
    _cf_border_top(whole, f'=AND({u}="유니버스 밖",{u_prev}<>"유니버스 밖")', SEPARATOR)
    if "상태" in geo.have:
        s = f"${_letter(geo.cols['상태'])}{r0}"
        cf_expr(_body(ws, geo, "상태"), f'=OR(LEFT({s},2)="오류",LEFT({s},2)="이전")', font=C["up"], bold=True)
    if "다음이벤트" in geo.have and "다음이벤트일" in geo.have:
        d = f"${_letter(geo.cols['다음이벤트일'])}{r0}"
        ev = _body(ws, geo, "다음이벤트")
        cf_expr(ev, f"=AND(ISNUMBER({d}),{d}<TODAY())", font=PAST_FONT)               # 지난 이벤트(갱신 전) 회색
        cf_expr(ev, f"=AND(ISNUMBER({d}),{d}>=TODAY(),{d}-TODAY()<=7)", font=C["accent"], bold=True)   # 7일 이내


def _qpack(builder: Any, ws: Any, geo: _Geo) -> None:
    """Q.Pack 3색: 열마다 기준점 = 유니버스 = 대회 행만의 최소·중앙값·최대(숨김 행), 줄무늬는 20칸 한 척도."""
    if "유니버스" not in geo.have:
        builder.say(f"[{SHEET}] 경고: {TABLE}에 유니버스 열이 없어 Q.Pack 색을 건너뜁니다")
        return
    # 기준 범위는 머리글 포함([#All]) — 머리글 칸은 글자라 ISNUMBER로 빠지고, 0·1행 표에서도 참조가 2칸 이상의 배열이 된다.
    # (본문만 쓰면 0·1행 표에서 참조가 칸 하나라 IF가 FALSE 한 값을 돌려주고, COUNT·MIN이 그것을 0으로 세어 기준점이
    #  ""가 아니라 0이 됐다 — 실측)
    cond = f'{TABLE}[[#All],[유니버스]]="대회"'
    rows = [geo.header_row - 4 + k for k in range(3)]      # 기본 배치에서 ANCHOR_ROWS
    for k, text in enumerate(("Q.Pack 기준점 낮음 — 열: 대회 종목 최소 · 줄무늬: 5백분위",
                              "Q.Pack 기준점 중앙 — 열: 대회 종목 중앙값 · 줄무늬: 50백분위",
                              "Q.Pack 기준점 높음 — 열: 대회 종목 최대 · 줄무늬: 95백분위")):
        put(ws, ws.Cells(rows[k], geo.left).Address, text, size=8, color=C["muted"])
    done = 0
    for name in QPACK_COLUMNS:
        if name not in geo.have:
            continue
        c = geo.cols[name]
        anchors = [ws.Cells(r, c) for r in rows]
        cf_qpack(_body(ws, geo, name), anchors, ref=f"{TABLE}[[#All],[{name}]]", cond=cond)
        nf = SPECS.get(name, {}).get(_NF)
        if nf:
            set_nf(ws.Range(anchors[0], anchors[-1]), nf)
        done += 1
    if geo.exact:
        c0 = geo.cols[STRIP[0]]
        anchors = [ws.Cells(r, c0) for r in rows]
        cf_qpack_block(_body(ws, geo, STRIP[0], STRIP[-1]), anchors,
                       ref=f"{TABLE}[[#All],[{STRIP[0]}]:[{STRIP[-1]}]]", cond=cond, pcts=QPACK_STRIP_PCTS,
                       hide_values=True, width=STRIP_WIDTH)
        set_nf(ws.Range(anchors[0], anchors[-1]), "0.00%")
    else:
        builder.say(f"[{SHEET}] 경고: 표 열 순서가 계약과 달라 최근 20일 줄무늬 색을 건너뜁니다")
    ws.Range(ws.Cells(rows[0], 1), ws.Cells(rows[-1], 1)).EntireRow.Hidden = True
    builder.say(f"[{SHEET}] Q.Pack 색: 열 {done}개 + 줄무늬" + ("" if geo.exact else " 생략"))


def _groups(ws: Any, geo: _Geo) -> None:
    """열 묶음(+/−)을 만들고 기본으로 접는다. 묶음마다 바로 왼쪽의 보이는 열이 요약 열(단추 자리)."""
    for first, last in COLLAPSED_GROUPS:
        group_columns(ws, f"{_letter(geo.cols[first])}:{_letter(geo.cols[last])}", collapsed=True, summary_left=True)


def _freeze(ws: Any, row: int, col: int) -> None:
    """(row, col) 칸의 위·왼쪽을 고정. 창을 맨 위·왼쪽으로 돌린 뒤 칸을 골라 고정해야 숨김 행·선택 위치와 무관하게 맞는다."""
    ws.Activate()
    win = ws.Application.ActiveWindow
    win.FreezePanes = False
    win.ScrollRow = 1
    win.ScrollColumn = 1
    ws.Cells(row, col).Select()
    win.FreezePanes = True


# ------------------------------------------------------------------------------------------------ 진입점
def build(builder: Any) -> None:
    """`대회종목` 시트를 만든다(tblCompany는 LOADS 위치에 적재돼 있어야 함 — 0행이어도 됨).

    Args:
        builder: build_dashboard.Builder 또는 하네스 대역(StandInBuilder). wb·ws·lo·table_style·say·nav_links·
            HEADER_RIGHT·HEADER_STATUS를 쓴다.

    Raises:
        없음 — 시트·표가 없거나 열 구성이 다르면 경고만 남기고 가능한 부분까지 만든다(COM 서식 오류는 그대로 올라옴).

    Example:
        import pages.page_company as p; p.prepare_tables(builder); ...새로 고침...; p.build(builder)
    """
    ws = _sheet(builder)
    if ws is None:
        return
    lo = _table(builder)
    if lo is not None:
        geo = _geometry(lo)
        if lo.Range.Cells(1, 1).Address.replace("$", "") != TABLE_CELL:
            builder.say(f"[{SHEET}] 경고: {TABLE}이(가) {lo.Range.Cells(1, 1).Address.replace('$', '')}에 있어 "
                        f"페이지 서식을 건너뜁니다(LOADS 위치 {TABLE_CELL}에 적재해야 함)")
            _overwrite_on_refresh(lo)
            return
    else:
        geo = _Geo(HEADER_ROW, _contract_cols(), False, set(), 0)
    last_col = geo.right + 1

    _reset_sheet(ws, geo.header_row, last_col)
    sheet_setup(ws, bg=False, zoom=85, tab=C["good"], widths={"A": 1.5})
    ws.Rows(1).RowHeight = 6
    if lo is not None:
        style_name = getattr(builder, "table_style", None)
        if style_name:
            try:
                lo.TableStyle = style_name
            except pywintypes.com_error:
                builder.say(f"[{SHEET}] 경고: 표 스타일 '{style_name}'을 적용하지 못했습니다")
        _overwrite_on_refresh(lo)
        try:
            lo.ShowAutoFilter = True
        except pywintypes.com_error:
            pass
    _set_widths(ws, geo)
    _title(builder, ws, geo, last_col)
    _nav_row(builder, ws, geo, last_col)
    _controls(builder, ws, geo, lo is not None)
    if lo is not None:
        _format_columns(builder, ws, geo)
        if geo.exact:
            _band_row(ws, geo)
        _row_rules(builder, ws, geo)
        _qpack(builder, ws, geo)
        if geo.exact:
            _groups(ws, geo)
        else:
            builder.say(f"[{SHEET}] 경고: 표 열이 계약(103열·순서)과 달라 열 묶음 머리글·접기를 건너뜁니다")
    _freeze(ws, geo.first_row, geo.cols[LABEL_COL] + 1)
    ws.Range("A1").Select()
    builder.say(f"{SHEET} 시트 완료" + ("" if lo is not None else " (표 없음 — 상단만)"))
