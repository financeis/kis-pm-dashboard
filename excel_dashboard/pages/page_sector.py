"""업종·테마 페이지 빌더 (화면과 [업종] 버튼, 2026-10-01).

시트 `업종` 한 장에 표 두 개를 위아래로 쌓는다(위 A, 아래 B). 둘 다 Power Query가 채우는 표라서 이 모듈은 값을 쓰지 않고
서식·조건부 서식·안내 줄·버튼만 입힌다.
  A. KRX 업종지수와 업종별 수급 `tblSectorKRX` (쿼리 T_SectorKRX, 46행 안팎) — 열마다 Q.Pack 3색
  B. 테마 집계 `tblThemeAgg` (쿼리 T_ThemeAgg, 대테마 행 + 그 아래 세부테마 행, 100행 안팎) — 열마다 Q.Pack 3색

인터페이스 (빌더 배선이 지켜야 할 것)
  SHEET           시트 이름. 빌더의 SHEETS에 넣어 builder.ws[SHEET]로 접근할 수 있어야 한다.
  LOADS           (쿼리, 표 이름, 머리글 왼쪽 위 셀) — 이 위치에 표를 적재한 뒤 prepare_tables → (첫 새로 고침) → build 순서로 부른다.
                  A는 24열(B~Y)이고 머리글(15행) 아래 56행(KOSPI 0005~0030 + KOSDAQ 1006~1033 코드 범위의 최댓값 54 + 여유 2)을,
                  B는 22열(B~W)이고 머리글(80행) 아래 150행을 비워 둔다. A가 다 차지 않아 생기는 빈 행(46행이면 12행)은 의도한 여유다.
  prepare_tables(builder)  두 표의 `QueryTable.RefreshStyle`을 0(xlOverwriteCells)으로 바꾼다. 표를 적재한 직후, 첫 새로 고침 전에
                  부른다(기본값 '셀 삽입·삭제'에서는 0행 새로 고침이 셀을 지워 아래 표를 끌어올릴 수 있다고 다른 페이지 작업에서 실측 보고됨). build도 한 번 더 부르지만 순서는 이쪽이 먼저.
  build(builder)  builder.wb / ws / lo / table_style / say / nav_links / HEADER_RIGHT / HEADER_STATUS만 쓴다.
                  표가 없으면 경고만 남기고 그 블록의 서식·색·안내 줄은 건너뛴다(제목 띠·버튼·이름 정의·구역 제목은 그대로 그린다).
                  여러 번 불러도 규칙·도형·이름이 겹치지 않는다.
  표 행 추가·삭제 금지  이 두 표에는 ListRows.Add/Delete(빌더의 빈 표 임시 행 `ensure_rows`·`finish` 포함)를 쓰지 않는다. 표 행을 지우면
                  표 열(B~Y)의 셀이 위로 당겨져 아래 블록 전체가 따라 올라온다(실측: A 표에서 22행을 지우자 B 표가 22행 올라와 A 표를 다시
                  채우는 새로 고침이 실패). 0행 표에도 열 서식이 들어가도록 표 아래 예약 범위에 서식을 직접 건다.
  이름 정의       통합문서 범위 `업종_최근조회`(D7 — 매크로가 종료 시각을 씀, 날짜·시각 서식)와 바로 오른쪽 `업종_상태`(E7). 둘 다 처음엔 "-".
                  배선에서 같은 이름을 다시 정의하지 않는다.
  버튼            도형 `btn_RefreshSector`(B6:C7, 글자 "업종"), OnAction = RefreshSector (매크로는 vba/mod_refresh.bas).
  표 구조         쿼리가 낸 그대로 둔다(열 추가·이동 없음). 계층(들여쓰기·굵게)·색·안내 줄은 모두 서식과 시트 수식이다.

설계 이유
  - 두 표를 위아래로 쌓아 각 표를 한 화면 너비(85% 확대에서 1920px 안팎)에 담았다. 열 너비는 같은 열 번호에서 두 표가 요구하는 값 중 큰 쪽.
    틀 고정은 위 7행(제목·메뉴·버튼·최근 조회·범례 — 어디서든 버튼과 상태가 보임)과 A~D열(두 표의 이름 열). 표 머리글은 블록마다
    위치가 달라 고정하지 않는다.
  - Q.Pack 기준점 3행(최솟값·중앙값·최댓값)을 색을 칠하는 열 바로 위에 보이게 둔다. 척도가 눈에 보이고 셀 수식이라 표가 커져도 따라간다.
    A는 `구분 = 업종` 행만, B는 `단계 = 세부테마` 행만 기준 계산에 쓴다(대분류·대테마 행은 합계라 극단값이 되기 쉽다). 합계 행도
    같은 척도로 칠하되 기준에서는 뺀다. 수준별로 따로 칠하는 방식은 색 척도가 연속 범위 전체에 걸려서 두 수준이 한 줄씩 번갈아 나오는
    B에서는 만들 수 없다. 대테마 값은 그 세부테마 값들의 시가총액 가중 평균이라 세부테마 최솟값~최댓값 안에 들어와 같은 척도로도 읽힌다.
    조건에 맞는 숫자가 없으면(빈 표·전부 오류) 기준점 칸은 빈 글자이고 아무 칸도 칠하지 않는다.
  - 행 단위 모양(대분류·대테마 행 굵게·옅은 바탕·위 구분선, 세부테마 행의 반복 대테마 이름 흐리게, 오류·이전 데이터 표시)은 조건부 서식이다.
    행이 새로 고침으로 바뀌어도, 사용자가 정렬해도 값만 보고 따라간다(정렬 후 시험함). 세부테마 들여쓰기는 그 열 전체의 정적 들여쓰기
    (대테마 행은 빈 칸이라 안 보임). 규칙은 색 척도 31개 + 행·상태 규칙 15개 = 46개로, 전체 너비 규칙은 5개뿐이다.
  - 색과 행 서식의 적용 범위는 표 높이가 아니라 예약 범위(고정 행 수)다. 새로 고침을 '덮어쓰기'로 바꿔서 표가 늘고 줄어도 규칙 범위가 그대로다
    (0행 → 46·96행 → 24·73행 → 46·96행으로 시험: 규칙 목록 불변, 아래 블록 불변, 새 행에 서식 유지).
  - 숫자는 Q.Pack 배경 위에서 읽혀야 해서 빨강·파랑 글자 서식 대신 부호(+/-)만 붙인 검은 글자. 숫자 칸에 IndentLevel을 쓰면 인쇄·PDF
    렌더링에서 글자 폭 계산이 달라 열이 넉넉해도 '####'로 넘쳤다(실측) — 오른쪽 여백은 숫자 서식의 `_)`로 둔다.
  - 쿼리가 0행을 돌려준 표는 구조적 참조가 '빈 행 1개'로 계산된다(COUNT=0, ROWS=1 — 오류 아님). 그래서 안내·상태 줄 수식은 빈 표에서도 동작한다.
"""
from __future__ import annotations

import dataclasses
import re
from typing import Any

import pywintypes

from xl_helpers import (
    C, NF, QPACK_HIGH, QPACK_LOW, QPACK_MID, XL_CENTER, XL_CONTINUOUS, XL_EDGE_LEFT, XL_EDGE_TOP, XL_LEFT, XL_RIGHT, XL_THIN,
    add_button, add_names, cf_expr, cf_qpack, fill, freeze, hyperlink, put, rgb, section, sheet_setup, style, title_bar,
)

__all__ = ["SHEET", "LOADS", "build", "prepare_tables"]

SHEET = "업종"
SECTOR_TABLE = "tblSectorKRX"
THEME_TABLE = "tblThemeAgg"

# 표 위치 (머리글 행). 머리글 위로 6행(구역 제목·상태 줄·기준점 3행·묶음 머리글)을 쓴다.
SECTOR_HEADER_ROW = 15
SECTOR_ROWS = 56                                   # A 예약 행 수 (코드 범위 최대 26 + 28 = 54, 여유 2)
THEME_ROWS = 150                                   # B 예약 행 수
THEME_HEADER_ROW = SECTOR_HEADER_ROW + SECTOR_ROWS + 3 + 6   # A 예약(16~71행) 뒤 빈 행 2개 → B 구역 제목 74행 → 머리글 80행
SECTOR_CELL = f"B{SECTOR_HEADER_ROW}"
THEME_CELL = f"B{THEME_HEADER_ROW}"
LOADS = [("T_SectorKRX", SECTOR_TABLE, SECTOR_CELL), ("T_ThemeAgg", THEME_TABLE, THEME_CELL)]

SECTOR_COLS = ["시장", "코드", "업종명", "구분", "기준일", "지수", "1D", "1W", "1M", "3M", "6M", "YTD", "1Y",
               "외국인1D억", "외국인1W억", "외국인1M억", "기관1D억", "기관1W억", "기관1M억", "개인1D억", "개인1W억", "개인1M억",
               "상태", "조회시각"]
THEME_COLS = ["단계", "대테마", "세부테마", "종목수", "시가총액억", "1D", "1W", "1M", "3M", "6M", "YTD", "1Y", "상승비율1D", "상승비율1W",
              "외국인1D%", "외국인1W%", "외국인1M%", "기관1D%", "기관1W%", "기관1M%", "상태", "조회시각"]
PERIODS = ["1D", "1W", "1M", "3M", "6M", "YTD", "1Y"]
SECTOR_FLOWS = ["외국인1D억", "외국인1W억", "외국인1M억", "기관1D억", "기관1W억", "기관1M억", "개인1D억", "개인1W억", "개인1M억"]
THEME_RATIOS = ["상승비율1D", "상승비율1W"]
THEME_FLOWS = ["외국인1D%", "외국인1W%", "외국인1M%", "기관1D%", "기관1W%", "기관1M%"]

# Q.Pack 기준 행 조건 — 합계 성격의 행(대분류·대테마)은 기준점 계산에서 뺀다
SECTOR_COND = f'{SECTOR_TABLE}[구분]="업종"'
THEME_COND = f'{THEME_TABLE}[단계]="세부테마"'

NAME_RECENT = "업종_최근조회"
NAME_STATUS = "업종_상태"
RECENT_CELL = "D7"
STATUS_CELL = "E7"
BUTTON_SPAN = "B6:C7"
BUTTON_CAPTION = "업종"
BUTTON_MACRO = "RefreshSector"

XL_OVERWRITE_CELLS = 0       # QueryTable.RefreshStyle: 셀을 끼워 넣거나 지우지 않고 덮어씀
XL_CENTER_ACROSS = 7         # 선택 영역의 가운데 맞춤 (병합 없이 묶음 이름을 가운데에)
XL_TOP = -4160               # 조건부 서식 테두리: 위쪽

FREEZE_ROWS = 7              # 제목·메뉴·버튼 줄 (어디서든 버튼과 상태가 보임)
FREEZE_COLS = 4              # A~D: 두 표의 이름 열(시장·코드·업종명 / 단계·대테마·세부테마)이 가로 스크롤에도 남음
BAND_LAST_COL = "AF"         # 제목 띠·메뉴 줄 오른쪽 끝 (메뉴 링크가 늘어도 글자가 띠 밖으로 나가지 않게 넉넉히)
DATA_ROW_HEIGHT = 16.5

# 숫자 서식: 부호만 붙인 검은 글자(Q.Pack 배경 위라 색 코드는 쓰지 않음). 끝의 _) 는 오른쪽 여백 — 들여쓰기(IndentLevel)는 인쇄·PDF 렌더링에서
# 글자 폭 계산이 달라 '####'로 넘치는 일이 있어(실측) 숫자 칸에는 쓰지 않는다.
NF_PCT = "+0.00%_);-0.00%_);0.00%_)"
NF_EOK = "+#,##0_);-#,##0_);0_)"
NF_RATIO = "0.0%_)"
NF_INDEX = "#,##0.00_)"
NF_INT = "#,##0_)"
NF_MDHM = "mm-dd hh:mm"

# 열 서식: width = 글자 수 (AutoFilter 단추 몫 포함), nf = 숫자 서식, h = 가로 정렬, 나머지는 글꼴. 같은 열 번호는 두 표 중 큰 너비를 쓴다.
_RET = dict(width=9.5, nf=NF_PCT, h=XL_RIGHT)
_FLOW = dict(width=13, nf=NF_EOK, h=XL_RIGHT)
SECTOR_SPEC: dict[str, dict] = {
    "시장": dict(width=8.5, h=XL_CENTER),
    "코드": dict(width=6, h=XL_CENTER, color=C["muted"], size=9),
    "업종명": dict(width=22, h=XL_LEFT),
    "구분": dict(width=8, h=XL_CENTER, color=C["muted"], size=9),
    "기준일": dict(width=11.5, nf=NF["date"], h=XL_CENTER, color=C["muted"], size=9),
    "지수": dict(width=12, nf=NF_INDEX, h=XL_RIGHT),
    **{p: dict(_RET) for p in PERIODS},
    **{f: dict(_FLOW) for f in SECTOR_FLOWS},
    "상태": dict(width=9, h=XL_LEFT, color=C["muted"], size=9),
    "조회시각": dict(width=11.5, nf=NF_MDHM, h=XL_CENTER, color=C["muted"], size=9),
}
THEME_SPEC: dict[str, dict] = {
    "단계": dict(width=9, h=XL_CENTER, color=C["muted"], size=9),
    "대테마": dict(width=12, h=XL_LEFT),
    "세부테마": dict(width=22, h=XL_LEFT, indent=1),
    "종목수": dict(width=9, nf=NF["num0"], h=XL_CENTER),
    "시가총액억": dict(width=12.5, nf=NF_INT, h=XL_RIGHT),
    **{p: dict(_RET) for p in PERIODS},
    **{r: dict(width=13, nf=NF_RATIO, h=XL_RIGHT) for r in THEME_RATIOS},
    **{f: dict(width=13, nf=NF_PCT, h=XL_RIGHT) for f in THEME_FLOWS},
    "상태": dict(width=9, h=XL_LEFT, color=C["muted"], size=9),
    "조회시각": dict(width=11.5, nf=NF_MDHM, h=XL_CENTER, color=C["muted"], size=9),
}

# 열 묶음 머리글: (이름, 첫 열, 마지막 열). 표 머리글 위 한 줄.
SECTOR_GROUPS = [("KRX 업종지수", "시장", "지수"), ("기간 등락 (지수 기준)", "1D", "1Y"),
                 ("외국인 순매수 (억원)", "외국인1D억", "외국인1M억"), ("기관계 순매수 (억원)", "기관1D억", "기관1M억"),
                 ("개인 순매수 (억원)", "개인1D억", "개인1M억"), ("조회", "상태", "조회시각")]
THEME_GROUPS = [("테마 (대회 종목)", "단계", "시가총액억"), ("기간 등락 (시가총액 가중)", "1D", "1Y"),
                ("상승 종목 비율", "상승비율1D", "상승비율1W"), ("외국인 순매수 (테마 시가총액 대비)", "외국인1D%", "외국인1M%"),
                ("기관 순매수 (테마 시가총액 대비)", "기관1D%", "기관1M%"), ("조회", "상태", "조회시각")]

GROUP_FILLS = ("#DCE6F5", "#E7EDF8")     # 묶음 머리글 바탕 (인접 묶음끼리 번갈아)
ROW_GROUP_FILL = "#EAEFF7"               # 대분류·대테마 행 바탕 (Q.Pack 색 칸은 색이 우선)
MUTED_ERR = "#8A94A3"
DIM_TEXT = "#9AA5B1"
SEPARATOR = "#B8C2D1"


# ------------------------------------------------------------------------------------------------ 위치·도우미
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


def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _geometry(lo: Any) -> _Geo:
    hdr = lo.HeaderRowRange
    left = int(hdr.Column)
    cols = {str(c.Name): int(c.Range.Column) for c in lo.ListColumns}
    return _Geo(int(hdr.Row), left, left + len(cols) - 1, cols)


def _fallback_geo(ws: Any, cell: str, names: list[str]) -> _Geo:
    """표가 없을 때 구역 제목·묶음 머리글만 그리기 위한 위치(LOADS의 셀 + 기대한 열 이름 순서)."""
    rng = ws.Range(cell)
    left = int(rng.Column)
    return _Geo(int(rng.Row), left, left + len(names) - 1, {n: left + i for i, n in enumerate(names)})


def _sheet(builder: Any) -> Any:
    try:
        return builder.ws[SHEET]
    except KeyError:
        builder.say(f"[{SHEET}] 경고: 시트가 없어 페이지를 만들지 않습니다(빌더 SHEETS에 '{SHEET}' 추가 필요)")
        return None


def _table(builder: Any, name: str) -> Any:
    try:
        return builder.lo[name]
    except KeyError:
        builder.say(f"[{SHEET}] 경고: 표 {name}이(가) 없어 해당 블록의 서식·색·안내 줄을 건너뜁니다"
                    f"(LOADS 위치에 적재한 뒤 호출해야 합니다)")
        return None


def _ref(table: str, col: str) -> str:
    """구조적 참조. 글자·숫자·밑줄 말고 다른 글자(%, 공백 등)가 든 열 이름은 이중 대괄호."""
    return f"{table}[{col}]" if re.fullmatch(r"[0-9A-Za-z_가-힣]+", col) else f"{table}[[{col}]]"


def _overwrite_on_refresh(lo: Any) -> None:
    """쿼리 새로 고침이 표 열 안의 셀을 끼워 넣거나 지우지 않고 덮어쓰게 한다(xlOverwriteCells).

    기본값(셀 삽입·삭제)에서는 행 수가 바뀔 때 Excel이 조건부 서식 범위를 쪼개거나 복사하고(표가 0·1행에서 늘면 열별 규칙이 새 행에
    이어지지 않음 — 실측), 0행 새로 고침은 셀을 지워 아래에 둔 표를 끌어올릴 수 있다(보고됨). 이 시트는 표 아래가 비어 있게 예약해 두었으므로 덮어써도 안전하다.
    정적 표(쿼리 아님)는 QueryTable이 없어 건너뛴다.
    """
    try:
        lo.QueryTable.RefreshStyle = XL_OVERWRITE_CELLS
    except pywintypes.com_error:
        pass


def prepare_tables(builder: Any) -> None:
    """두 표의 새로 고침 방식을 '덮어쓰기'(RefreshStyle 0)로 바꾼다 — 표를 적재한 직후, 첫 새로 고침 전에 부른다.

    Args:
        builder: build_dashboard.Builder 또는 하네스 대역. builder.lo[표이름]만 쓴다.

    Raises:
        없음 — 표가 없으면 경고만 남기고 있는 표만 처리한다.

    Example:
        import pages.page_sector as p; p.prepare_tables(builder)
    """
    for name in (SECTOR_TABLE, THEME_TABLE):
        lo = _table(builder, name)
        if lo is not None:
            _overwrite_on_refresh(lo)


def _block(ws: Any, geo: _Geo, n: int, first_col: str | None = None, last_col: str | None = None) -> Any:
    """표 데이터 첫 행부터 예약 n행 범위. first_col·last_col = 열 이름(없으면 표 전체 열)."""
    c_left = geo.cols[first_col] if first_col else geo.left
    c_right = geo.cols[last_col] if last_col else geo.right
    return ws.Range(ws.Cells(geo.first_row, c_left), ws.Cells(geo.first_row + n - 1, c_right))


def _cf(rng: Any, formula: str, *, font: str | None = None, fill_color: str | None = None, bold: bool | None = None,
        italic: bool | None = None, stop: bool = False, top: str | None = None) -> Any:
    """조건부 서식 한 개 추가. 먼저 추가한 규칙이 우선순위가 높다(COM은 항상 맨 뒤에 붙임)."""
    fc = cf_expr(rng, formula, font=font, fill_color=fill_color, bold=bold, stop=stop)
    if italic is not None:
        fc.Font.Italic = italic
    if top is not None:
        b = fc.Borders(XL_TOP)
        b.LineStyle = XL_CONTINUOUS
        b.Color = rgb(top)
    return fc


def _edge(rng: Any, edge: int, hex_color: str, weight: int = XL_THIN) -> None:
    b = rng.Borders(edge)
    b.LineStyle = XL_CONTINUOUS
    b.Weight = weight
    b.Color = rgb(hex_color)


# ------------------------------------------------------------------------------------------------ 수식 (구역 안내·상태 줄)
def _sector_note(t: str = SECTOR_TABLE) -> str:
    return (f'=LET(a_n,COUNT({t}[지수]),'
            f'IF(a_n=0,"데이터 없음 · 장중: 지수·등락은 실시간, 수급은 직전 완료 세션까지",'
            f'"기준일 "&TEXT(MAX({t}[기준일]),"yyyy-mm-dd")&" · 업종 "&COUNTIF({t}[구분],"업종")&"개 + 대분류 "&COUNTIF({t}[구분],"대분류")&"개"'
            f'&" · 조회 "&TEXT(MAX({t}[조회시각]),"mm-dd hh:mm")&" · 장중: 지수·등락은 실시간, 수급은 직전 완료 세션까지"))')


def _sector_status(t: str = SECTOR_TABLE) -> str:
    st = f"{t}[상태]"
    return (f'=LET(k_e,COUNTIF({st},"오류*"),k_p,COUNTIF({st},"이전*"),k_n,COUNT({t}[지수]),'
            f'IF(k_p>0,"▲ "&INDEX({st},MATCH("이전*",{st},0)),'
            f'IF(k_e>0,"▲ 업종 "&k_e&"개 조회 오류 — "&SUBSTITUTE(INDEX({st},MATCH("오류*",{st},0)),"오류: ",""),'
            f'IF(k_n=0,"○ 데이터 없음 — [업종] 버튼으로 불러오세요",'
            f'"● 정상 — 색은 열마다 따로 · 대분류 행(굵게·옅은 바탕)은 색 기준에서 빼고 같은 척도로 표시"))))')


def _theme_note(t: str = THEME_TABLE) -> str:
    return (f'=LET(m_b,COUNTIF({t}[단계],"대테마"),m_s,COUNTIF({t}[단계],"세부테마"),'
            f'IF(m_b=0,"데이터 없음 · 등락은 시가총액 가중 평균, 수급은 테마 시가총액 대비 %",'
            f'"대회 종목 "&TEXT(SUMIFS({t}[종목수],{t}[단계],"대테마"),"#,##0")&"개 · 대테마 "&m_b&" · 세부테마 "&m_s'
            f'&" · 계산 "&TEXT(MAX({t}[조회시각]),"mm-dd hh:mm")&" · 등락은 시가총액 가중, 수급은 테마 시가총액 대비 %"))')


def _theme_status(t: str = THEME_TABLE) -> str:
    st = f"{t}[상태]"
    return (f'=LET(m_e,COUNTIF({st},"오류*"),m_p,COUNTIF({st},"이전*"),m_b,COUNTIF({t}[단계],"대테마"),m_d,COUNTIF({st},"데이터 없음"),'
            f'IF(m_p>0,"▲ "&INDEX({st},MATCH("이전*",{st},0)),'
            f'IF(m_e>0,"▲ "&INDEX({st},MATCH("오류*",{st},0)),'
            f'IF(m_b=0,"○ 데이터 없음 — [시세]·[전체]로 대회종목 표를 채운 뒤 [업종]을 누르세요",'
            f'"● 정상 — 대테마(굵게) 아래 세부테마 · 색은 열마다 따로(세부테마 행 기준, 대테마 행도 같은 척도)"'
            f'&IF(m_d>0," · 시세 없는 묶음 "&m_d&"개","")))))')


# ------------------------------------------------------------------------------------------------ 제목 띠·조작 줄
def _header_band(builder: Any, ws: Any, right_col: str) -> None:
    """제목 띠 + 메뉴 줄. 머리글 수식이 참조하는 표(tblIndexNow 등)가 없으면 수식 없이 그린다."""
    subtitle = ("A. KRX 업종지수·업종별 투자자 순매수 · B. 대회 종목 테마 집계(대테마 → 세부테마) | [업종] 버튼으로 둘 다 갱신 | "
                "장중: 지수·등락은 실시간, 수급은 직전 완료 세션까지")
    span = f"B2:{BAND_LAST_COL}3"
    right = getattr(builder, "HEADER_RIGHT", None)
    status = getattr(builder, "HEADER_STATUS", None)
    try:
        title_bar(ws, "업종·테마", subtitle, span, right_formula=right, right_cell=f"{right_col}2",
                  right2_formula=status, right2_cell=f"{right_col}3")
    except pywintypes.com_error:
        builder.say(f"[{SHEET}] 경고: 머리글 기준 시각·상태 수식을 넣지 못했습니다(참조 표 없음) — 제목만 표시")
        title_bar(ws, "업종·테마", subtitle, span)
    fill(ws.Range(f"B4:{BAND_LAST_COL}4"), "navy2")
    builder.nav_links(ws, 4)
    ws.Rows(4).RowHeight = 16
    ws.Rows(5).RowHeight = 8


def _control_strip(builder: Any, ws: Any, jump_target: str) -> None:
    """6~7행: [업종] 버튼 자리 · 최근 조회(이름 정의) · 상태(이름 정의) · Q.Pack 범례 · B 구역으로 이동 링크(B 구역 제목 줄에는 되돌아오는 링크).

    버튼 도형은 열 너비·행 높이가 모두 정해진 뒤에 만든다(도형은 만들 때 셀 위치로 자리를 잡고 이후 따라가지 않음).
    """
    ws.Rows(6).RowHeight = 17
    ws.Rows(7).RowHeight = 17
    ws.Rows(8).RowHeight = 8
    put(ws, "D6", "최근 조회", size=8, color=C["muted"], h=XL_LEFT)
    put(ws, RECENT_CELL, "-", nf=NF["dt"], size=10, bold=True, color=C["navy"], h=XL_LEFT)
    put(ws, "E6", "상태", size=8, color=C["muted"], h=XL_LEFT)
    put(ws, STATUS_CELL, "-", size=10, bold=True, color=C["text"], h=XL_LEFT)
    st = ws.Range(STATUS_CELL)
    _cf(st, f'=LEFT({STATUS_CELL},2)="실패"', font=C["up"], bold=True)
    _cf(st, f'=LEFT({STATUS_CELL},2)="일부"', font=C["warn"], bold=True)
    _cf(st, f'={STATUS_CELL}="정상"', font=C["good"], bold=True)
    add_names(builder.wb, {NAME_RECENT: f"='{SHEET}'!${RECENT_CELL[0]}${RECENT_CELL[1:]}",
                           NAME_STATUS: f"='{SHEET}'!${STATUS_CELL[0]}${STATUS_CELL[1:]}"})
    # Q.Pack 범례 (낮음 → 중앙값 → 높음)
    put(ws, "K6", "색 범례", size=8, color=C["muted"], h=XL_RIGHT)
    for addr, label, color in (("L6", "낮음", QPACK_LOW), ("M6", "중앙값", QPACK_MID), ("N6", "높음", QPACK_HIGH)):
        put(ws, addr, label, size=9, bold=True, color=C["text"], fill=color, h=XL_CENTER)
    put(ws, "L7", "열마다 따로 칠함 · 기준점 = 그 열의 최솟값·중앙값·최댓값(A는 업종 행, B는 세부테마 행) · 높음 = 빨강",
        size=8, color=C["muted"], h=XL_LEFT)
    if jump_target:
        hyperlink(ws, "T6", jump_target, "▼ B. 테마 집계")
        style(ws.Range("T6"), color=C["accent"], size=9, bold=True)
        ws.Range("T6").Font.Underline = True


def _back_link(ws: Any, geo_b: _Geo, target: str) -> None:
    """B 구역 제목 줄(제목 글자 오른쪽 빈 칸)에 A 구역으로 돌아가는 링크 — 아래로 내려온 뒤 위로 올라갈 때 쓴다."""
    addr = ws.Cells(geo_b.header_row - 6, geo_b.left + 4).Address
    hyperlink(ws, addr, target, "▲ A. 업종")
    style(ws.Range(addr), color=C["accent"], size=9, bold=True)
    ws.Range(addr).Font.Underline = True


# ------------------------------------------------------------------------------------------------ 블록 (표 하나)
def _format_columns(builder: Any, ws: Any, lo: Any, geo: _Geo, spec: dict, rows: int) -> None:
    """열 서식을 표 아래 예약 범위(데이터 첫 행부터 rows행)에 직접 건다. 기대한 열이 없으면 경고만 남기고 있는 열만."""
    missing = [n for n in spec if n not in geo.cols]
    if missing:
        builder.say(f"[{SHEET}] 경고: {lo.Name}에 없는 열 {', '.join(missing)} — 그 열의 서식은 건너뜁니다")
    for name, sp in spec.items():
        if name not in geo.cols:
            continue
        rng = _block(ws, geo, rows, name, name)
        style(rng, nf=sp.get("nf"), bold=sp.get("bold"), color=sp.get("color"), size=sp.get("size"), h=sp.get("h"),
              indent=sp.get("indent"))
        hdr = ws.Cells(geo.header_row, geo.cols[name])
        hdr.HorizontalAlignment = XL_LEFT if sp.get("h") == XL_LEFT else XL_CENTER
    ws.Rows(f"{geo.first_row}:{geo.first_row + rows - 1}").RowHeight = DATA_ROW_HEIGHT
    ws.Rows(geo.header_row).RowHeight = 22


def _group_header(ws: Any, geo: _Geo, groups: list[tuple[str, str, str]]) -> None:
    """표 머리글 바로 위 한 줄: 열 묶음 이름(선택 영역 가운데 — 병합 없음), 인접 묶음끼리 바탕을 번갈아."""
    row = geo.header_row - 1
    ws.Rows(row).RowHeight = 17
    for i, (label, c1, c2) in enumerate(groups):
        if c1 not in geo.cols or c2 not in geo.cols:
            continue
        rng = ws.Range(ws.Cells(row, geo.cols[c1]), ws.Cells(row, geo.cols[c2]))
        ws.Cells(row, geo.cols[c1]).Value = label
        style(rng, bold=True, size=9, color=C["navy"], fill=GROUP_FILLS[i % 2], h=XL_CENTER_ACROSS)
        _edge(rng, XL_EDGE_LEFT, "#FFFFFF", 2)


def _anchor_labels(ws: Any, geo: _Geo, top: int) -> None:
    """기준점 3행의 이름 칸(표 첫 열 — 오른쪽 빈 칸으로 글자가 이어짐)."""
    for k, label in enumerate(("▼ 색 기준 최솟값 (낮음 · 초록)", "● 색 기준 중앙값 (노랑)", "▲ 색 기준 최댓값 (높음 · 빨강)")):
        c = ws.Cells(top + k, geo.left)
        c.Value = label
        style(c, size=8, color=C["muted"], h=XL_LEFT)
        ws.Rows(top + k).RowHeight = 13.5
    _edge(ws.Range(ws.Cells(top, geo.left), ws.Cells(top, geo.right)), XL_EDGE_TOP, C["line"])


def _blank_when_no_numbers(anchors: Any, ref: str, cond: str) -> None:
    """조건에 맞는 숫자가 하나도 없으면 기준점 칸을 빈 글자("")로 — 색 척도는 빈 기준점이면 아무 칸도 칠하지 않는다.

    헬퍼 수식은 행이 하나뿐인 표(쿼리가 0행을 돌려준 빈 표는 빈 행 1개로 계산됨)에서 조건에 안 맞으면 COUNT가 FALSE 한 개를
    숫자로 세어 0을 내므로(실측), 빈 표에서 최소·중앙·최대가 모두 0으로 보이지 않게 앞에 조건 확인을 덧붙인다.
    """
    for k in range(1, 4):
        cell = anchors.Cells(k, 1)
        body = str(cell.Formula2)[1:]
        cell.Formula2 = f'=IF(SUMPRODUCT(ISNUMBER({ref})*({cond}))=0,"",{body})'


def _qpack_columns(builder: Any, ws: Any, geo: _Geo, table: str, names: list[str], cond: str, anchor_top: int, rows: int,
                   spec: dict) -> None:
    """열마다 Q.Pack 3색: 기준점 3행(anchor_top부터) 수식 + 색 척도. 척도는 맨 앞 우선순위(SetFirstPriority)."""
    missing = [n for n in names if n not in geo.cols]
    if missing:
        builder.say(f"[{SHEET}] 경고: {table}에 없는 열 {', '.join(missing)} — 그 열의 색은 건너뜁니다")
    for name in names:
        if name not in geo.cols:
            continue
        c = geo.cols[name]
        anchors = ws.Range(ws.Cells(anchor_top, c), ws.Cells(anchor_top + 2, c))
        ref = _ref(table, name)
        cf_qpack(_block(ws, geo, rows, name, name), anchors, ref=ref, cond=cond)
        _blank_when_no_numbers(anchors, ref, cond)
        style(anchors, nf=spec[name].get("nf"), size=8, color=C["muted"], h=XL_RIGHT)


def _section_block(builder: Any, ws: Any, geo: _Geo, title: str, note: str | None, status: str | None) -> None:
    """구역 제목(밑줄) + 오른쪽 끝 안내 수식 + 바로 아래 상태 줄(수식, 색은 조건부 서식)."""
    sec = geo.header_row - 6
    stat = geo.header_row - 5
    lc, rc = _col_letter(geo.left), _col_letter(geo.right)
    section(ws, f"{lc}{sec}:{rc}{sec}", title)
    ws.Rows(sec).RowHeight = 20
    ws.Rows(stat).RowHeight = 16
    if note:
        try:
            put(ws, f"{rc}{sec}", formula=note, size=8, color=C["muted"], h=XL_RIGHT)
        except pywintypes.com_error:
            builder.say(f"[{SHEET}] 경고: '{title}' 구역의 안내 수식을 넣지 못했습니다(표 열 구성 확인)")
    if status:
        addr = f"{lc}{stat}"
        try:
            put(ws, addr, formula=status, size=9, color=C["muted"], h=XL_LEFT)
            cell = ws.Range(addr)
            _cf(cell, f'=LEFT({addr},1)="▲"', font=C["up"], bold=True)
            _cf(cell, f'=LEFT({addr},1)="●"', font=C["good"])
        except pywintypes.com_error:
            builder.say(f"[{SHEET}] 경고: '{title}' 구역의 상태 줄 수식을 넣지 못했습니다(표 열 구성 확인)")


def _sector_rows_cf(builder: Any, ws: Any, geo: _Geo) -> None:
    """A 표의 행 단위 조건부 서식 — 위에서부터 우선순위가 높다 (Q.Pack 색 척도는 이와 별개로 항상 맨 앞).

    상태 경고 > 오류 행 회색 > 대분류 행 굵게·옅은 바탕 > 시장이 바뀌는 첫 행의 위 구분선.
    """
    need = [c for c in ("시장", "구분", "지수", "상태") if c not in geo.cols]
    if need:
        builder.say(f"[{SHEET}] 경고: {SECTOR_TABLE}에 열 {', '.join(need)}이(가) 없어 행 서식을 건너뜁니다")
        return
    r0 = geo.first_row
    mk, kind, px, st = (f"${_col_letter(geo.cols[n])}{r0}" for n in ("시장", "구분", "지수", "상태"))
    mk_prev, px_prev = f"${_col_letter(geo.cols['시장'])}{r0 - 1}", f"${_col_letter(geo.cols['지수'])}{r0 - 1}"
    whole = _block(ws, geo, SECTOR_ROWS)
    _cf(_block(ws, geo, SECTOR_ROWS, "상태", "상태"), f'=OR(LEFT({st},2)="오류",LEFT({st},2)="이전")', font=C["up"], bold=True)
    _cf(whole, f'=AND(LEFT({st},2)="오류",NOT(ISNUMBER({px})))', font=MUTED_ERR, italic=True, stop=True)
    _cf(whole, f'={kind}="대분류"', fill_color=ROW_GROUP_FILL, bold=True)
    _cf(whole, f"=AND(ISNUMBER({px_prev}),{mk}<>{mk_prev})", top=SEPARATOR)


def _theme_rows_cf(builder: Any, ws: Any, geo: _Geo) -> None:
    """B 표의 행 단위 조건부 서식: 상태 경고 > 오류 행 회색 > 대테마 행(굵게·옅은 바탕·위 구분선) > 세부테마 행의 대테마 이름 흐리게."""
    need = [c for c in ("단계", "대테마", "종목수", "상태") if c not in geo.cols]
    if need:
        builder.say(f"[{SHEET}] 경고: {THEME_TABLE}에 열 {', '.join(need)}이(가) 없어 행 서식을 건너뜁니다")
        return
    r0 = geo.first_row
    lv, cnt, st = (f"${_col_letter(geo.cols[n])}{r0}" for n in ("단계", "종목수", "상태"))
    whole = _block(ws, geo, THEME_ROWS)
    _cf(_block(ws, geo, THEME_ROWS, "상태", "상태"), f'=OR(LEFT({st},2)="오류",LEFT({st},2)="이전")', font=C["up"], bold=True)
    _cf(whole, f'=AND(LEFT({st},2)="오류",NOT(ISNUMBER({cnt})))', font=MUTED_ERR, italic=True, stop=True)
    _cf(whole, f'={lv}="대테마"', fill_color=ROW_GROUP_FILL, bold=True, top=SEPARATOR)
    _cf(_block(ws, geo, THEME_ROWS, "대테마", "대테마"), f'={lv}="세부테마"', font=DIM_TEXT)


def _block_setup(builder: Any, ws: Any, lo: Any | None, geo: _Geo, *, kind: str) -> None:
    """블록 하나(A 또는 B)의 서식·안내·색을 모두 입힌다. kind = 'sector' | 'theme'."""
    sector = kind == "sector"
    table = SECTOR_TABLE if sector else THEME_TABLE
    spec = SECTOR_SPEC if sector else THEME_SPEC
    rows = SECTOR_ROWS if sector else THEME_ROWS
    title = "A. KRX 업종지수 · 업종별 수급" if sector else "B. 테마 집계 (대회 종목)"
    groups = SECTOR_GROUPS if sector else THEME_GROUPS
    anchor_top = geo.header_row - 4
    has = lo is not None
    _section_block(builder, ws, geo, title,
                   (_sector_note() if sector else _theme_note()) if has else None,
                   (_sector_status() if sector else _theme_status()) if has else None)
    _group_header(ws, geo, groups)
    if not has:
        return
    style_name = getattr(builder, "table_style", None)
    if style_name:
        try:
            lo.TableStyle = style_name
        except pywintypes.com_error:
            builder.say(f"[{SHEET}] 경고: 표 스타일 '{style_name}'을 적용하지 못했습니다({lo.Name})")
    _overwrite_on_refresh(lo)
    try:
        lo.ShowAutoFilter = True     # 정렬·필터 (색·행 서식이 모두 값 기준이라 정렬해도 따라감)
    except pywintypes.com_error:
        pass
    _format_columns(builder, ws, lo, geo, spec, rows)
    _anchor_labels(ws, geo, anchor_top)
    names = (PERIODS + SECTOR_FLOWS) if sector else (PERIODS + THEME_RATIOS + THEME_FLOWS)
    _qpack_columns(builder, ws, geo, table, names, SECTOR_COND if sector else THEME_COND, anchor_top, rows, spec)
    (_sector_rows_cf if sector else _theme_rows_cf)(builder, ws, geo)


def _apply_widths(ws: Any, geos_specs: list[tuple[_Geo, dict]]) -> None:
    """열 너비: 같은 열 번호에서 두 표가 요구하는 값 중 큰 쪽."""
    need: dict[int, float] = {}
    for geo, spec in geos_specs:
        for name, sp in spec.items():
            if name in geo.cols and "width" in sp:
                need[geo.cols[name]] = max(need.get(geo.cols[name], 0.0), float(sp["width"]))
    for col, w in need.items():
        ws.Columns(col).ColumnWidth = w


def build(builder: Any) -> None:
    """`업종` 시트를 만든다(표는 LOADS 위치에 이미 적재돼 있어야 함).

    Args:
        builder: build_dashboard.Builder 또는 하네스 대역(StandInBuilder). wb·ws·lo·table_style·say·nav_links·
            HEADER_RIGHT·HEADER_STATUS를 쓴다.

    Raises:
        없음 — 시트나 표가 없으면 경고만 남기고 가능한 부분까지 그린다(COM 서식 오류는 그대로 올라옴).

    Example:
        import pages.page_sector as p; p.build(builder)
    """
    ws = _sheet(builder)
    if ws is None:
        return
    lo_a = _table(builder, SECTOR_TABLE)
    lo_b = _table(builder, THEME_TABLE)

    sheet_setup(ws, bg=False, zoom=85, tab=C["accent"], widths={"A": 1.5})
    ws.Cells.FormatConditions.Delete()   # 다시 부르면 규칙이 겹치지 않게(이 시트는 이 모듈 전용)
    try:
        ws.Hyperlinks.Delete()
    except pywintypes.com_error:
        pass
    ws.Rows(1).RowHeight = 6

    geo_a = _geometry(lo_a) if lo_a is not None else _fallback_geo(ws, SECTOR_CELL, SECTOR_COLS)
    geo_b = _geometry(lo_b) if lo_b is not None else _fallback_geo(ws, THEME_CELL, THEME_COLS)
    for geo, cell, name in ((geo_a, SECTOR_CELL, SECTOR_TABLE), (geo_b, THEME_CELL, THEME_TABLE)):
        want = ws.Range(cell)
        if (geo.header_row, geo.left) != (int(want.Row), int(want.Column)):
            builder.say(f"[{SHEET}] 경고: {name}이(가) LOADS 위치({cell})가 아니라 {_col_letter(geo.left)}{geo.header_row}에 있습니다 — "
                        f"제목 띠·두 표 사이 간격이 어긋날 수 있으니 LOADS 위치에 적재하세요")

    _apply_widths(ws, [(geo_a, SECTOR_SPEC), (geo_b, THEME_SPEC)])
    _header_band(builder, ws, _col_letter(geo_a.right))
    _control_strip(builder, ws, f"'{SHEET}'!B{geo_b.header_row - 6}")
    _block_setup(builder, ws, lo_a, geo_a, kind="sector")
    _block_setup(builder, ws, lo_b, geo_b, kind="theme")
    _back_link(ws, geo_b, f"'{SHEET}'!B{geo_a.header_row - 6}")
    add_button(ws, BUTTON_SPAN, BUTTON_CAPTION, BUTTON_MACRO)   # 열 너비·행 높이가 정해진 뒤에

    ws.Activate()
    win = ws.Application.ActiveWindow
    win.ScrollRow = 1                # 조건부 서식 헬퍼가 셀을 선택하며 화면을 옮겼을 수 있으니 맨 위·왼쪽에서 고정
    win.ScrollColumn = 1
    freeze(ws, FREEZE_ROWS, FREEZE_COLS)
    ws.Range("A1").Select()
    builder.say("업종 시트 완료")
