"""뉴스·이벤트 페이지 빌더 (spec R18, 2026-10-01 T20).

시트 `뉴스·이벤트` 한 장에 두 표를 나란히 둔다.
  - 왼쪽  이벤트 캘린더 `tblEvents` (쿼리 T_Events, 일자순) — [전체] 버튼에서만 갱신. 보유 종목 행 강조·D-day·종류 색 태그.
  - 오른쪽 뉴스·공시 헤드라인 `tblNews` (쿼리 T_News, 최신순 최근 100건) — [모두 새로 고침]에서 갱신.
두 표는 Power Query가 채우는 표라서 이 모듈은 값을 쓰지 않고 서식·조건부 서식·상태 줄만 입힌다.

인터페이스 (배선 작업 T23이 지켜야 할 것)
  SHEET   시트 이름. 빌더의 SHEETS에 넣고 builder.ws[SHEET]로 접근할 수 있어야 한다.
  LOADS   (쿼리, 표 이름, 머리글 왼쪽 위 셀) — 이 위치에 표를 적재한 뒤 build(builder)를 부른다.
          이벤트 표는 9열(B~J), 가운데 K열은 빈 간격, 뉴스 표는 7열(L~R). 표 위쪽 6~7행에 구역 제목·상태 줄을 둔다(머리글은 8행).
  build(builder)  builder.wb / ws / lo / table_style / say / nav_links / HEADER_RIGHT / HEADER_STATUS만 쓴다.
                  표가 없으면 경고만 남기고 그 블록은 건너뛴다(시트·구역 제목은 그대로 그린다).

설계 이유
  - 상단 8행(제목·메뉴·구역 제목·상태 줄·표 머리글)을 틀 고정해, 아래로 스크롤해도 오류·갱신 경로 안내가 보인다.
  - 오류·이전 데이터 행의 사유는 상태 열이 좁아 잘리므로 상태 줄(수식)이 첫 사유를 대신 보여 준다.
  - D-day 열의 값은 쿼리가 돌던 날 기준이라 날짜가 바뀌면 틀린다(이벤트는 [전체]에서만 갱신). 일자에서 다시 계산한
    글자("D-3" 등)를 조건부 서식의 숫자 서식으로 덧입혀 항상 오늘 기준으로 보이게 한다(표 구조는 그대로).
  - 강조·회색·색 태그는 모두 일자·보유·종류 열 값과 TODAY()로 판정하므로 행이 바뀌어도 따라간다.
  - 두 표는 새로 고침 방식을 '덮어쓰기'(xlOverwriteCells)로 바꾼다. 기본값(셀 삽입·삭제)에서는 표가 0·1행에서 늘어날 때
    Excel이 열별 조건부 서식 규칙을 새 행에 이어 주지 않았고(실측), 규칙이 100개 넘는 전체 너비 방식은 Excel이 멈췄다.
    표 아래가 비어 있고 두 표가 나란히 있어 덮어써도 안전하다. 배선·VBA에서 이 값을 되돌리지 말 것.
"""
from __future__ import annotations

import contextlib
import dataclasses
from typing import Any, Iterator

import pywintypes

from xl_helpers import (
    C, NF, XL_CENTER, XL_EXPRESSION, XL_LEFT, XL_RIGHT, cf_expr, col_range, fill, freeze, put, rgb, section, set_col_format,
    sheet_setup, style, title_bar,
)

__all__ = ["SHEET", "LOADS", "build"]

SHEET = "뉴스·이벤트"
EVENTS_TABLE = "tblEvents"
NEWS_TABLE = "tblNews"
EVENTS_CELL = "B8"   # 이벤트 표 머리글 왼쪽 위 (9열: B~J)
NEWS_CELL = "L8"     # 뉴스 표 머리글 왼쪽 위 (7열: L~R)
LOADS = [("T_Events", EVENTS_TABLE, EVENTS_CELL), ("T_News", NEWS_TABLE, NEWS_CELL)]

EVENTS_COLS = ["일자", "D-day", "종목코드", "종목명", "종류", "상세", "보유", "상태", "조회시각"]
NEWS_COLS = ["일시", "종목코드", "종목명", "제목", "출처", "상태", "조회시각"]

XL_OVERWRITE_CELLS = 0   # QueryTable.RefreshStyle: 셀을 끼워 넣거나 지우지 않고 덮어씀

NF_MDHM = "mm-dd hh:mm"
NF_DDAY = '"D-"0;"D+"0;"D-day"'   # 양수 = 남은 날(D-3), 음수 = 지난 날(D+2), 0 = D-day

BAND_LAST_COL = "AF"     # 제목 띠·메뉴 줄 오른쪽 끝(메뉴 링크가 늘어도 글자가 띠 밖으로 나가지 않게 넉넉히)
RIGHT_TEXT_COL = "R"     # 머리글 오른쪽 문구(기준 시각·상태)를 붙일 열 = 뉴스 표 마지막 열
ROWS_EVENTS = 1500       # 조건부 서식 적용 행 수 — 표가 늘어나도(덮어쓰기) 규칙 범위를 넘지 않게 표보다 넉넉히
ROWS_NEWS = 600
DDAY_BACK, DDAY_AHEAD = 14, 90   # D-day 글자 서식을 덧입히는 범위(일). 밖의 행은 쿼리가 계산한 값 서식을 그대로 씀

# 열 서식: 너비(글자 수)·숫자 서식·가로 정렬·글꼴. 상세·제목은 줄바꿈 없이 두므로(두 표가 같은 행을 쓴다) 넓게 잡는다.
EVENTS_SPEC = {
    "일자": dict(width=11, nf=NF["date"], h=XL_CENTER),
    "D-day": dict(width=7, nf=NF_DDAY, h=XL_CENTER),
    "종목코드": dict(width=9, h=XL_CENTER, color=C["muted"], size=9),
    "종목명": dict(width=16, bold=True),
    "종류": dict(width=15, h=XL_CENTER),
    "상세": dict(width=58),
    "보유": dict(width=6, h=XL_CENTER, bold=True),
    "상태": dict(width=9, h=XL_LEFT, color=C["muted"], size=9),   # 긴 사유는 왼쪽부터 잘리게(이전 데이터·오류가 앞에서 보임)
    "조회시각": dict(width=11, nf=NF_MDHM, h=XL_CENTER, color=C["muted"], size=9),
}
NEWS_SPEC = {
    "일시": dict(width=11, nf=NF_MDHM, h=XL_CENTER),
    "종목코드": dict(width=9, h=XL_CENTER, color=C["muted"], size=9),
    "종목명": dict(width=15, bold=True),
    "제목": dict(width=80),
    "출처": dict(width=12, color=C["muted"], size=9),
    "상태": dict(width=9, h=XL_LEFT, color=C["muted"], size=9),
    "조회시각": dict(width=11, nf=NF_MDHM, h=XL_CENTER, color=C["muted"], size=9),
}

# 종류 색 태그 (채움, 글자) — 공급 늘어나는 것(전환·유상)은 따뜻한 색, 호재성(무상)은 초록, 일정성은 중간색
TAG_STYLES = {
    "CB·BW 전환 상장": ("#FFF1DB", "#B45309"),
    "신규상장": ("#DCE9FD", "#1A56B0"),
    "유상증자": ("#FDECEC", "#B91C1C"),
    "무상증자": ("#E6F4EA", "#1E7A34"),
    "합병·분할": ("#EFE8FB", "#5B35A8"),
    "배당 기준일": ("#FFF8DB", "#8A6A00"),
    "주주총회": ("#EEF1F5", "#4A5568"),
}
HELD_FILL = "#F3F8FF"      # 시세판의 보유 행 강조와 같은 색
PAST_FONT = "#9AA5B1"
MUTED_ERR = "#8A94A3"
SEPARATOR = "#C5CCD8"


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


def _fallback_geo(ws: Any, cell: str, ncols: int) -> _Geo:
    """표가 없을 때 구역 제목만 그리기 위한 위치(LOADS의 셀 기준)."""
    rng = ws.Range(cell)
    left = int(rng.Column)
    return _Geo(int(rng.Row), left, left + ncols - 1, {})


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
        builder.say(f"[{SHEET}] 경고: 표 {name}이(가) 없어 해당 블록의 서식·상태 줄을 건너뜁니다"
                    f"(LOADS 위치에 적재한 뒤 호출해야 합니다)")
        return None


def _overwrite_on_refresh(lo: Any) -> None:
    """쿼리 새로 고침이 표 열 안의 셀을 끼워 넣거나 지우지 않고 덮어쓰게 한다(xlOverwriteCells).

    기본값(셀 삽입·삭제)에서는 행 수가 바뀔 때 Excel이 조건부 서식 범위를 쪼개거나 복사하는데, 표가 0·1행에서
    늘어나면 열별 규칙(보유 칩·종류 태그 등)이 새 행에 적용되지 않고 오래된 규칙이 쌓였다(실측). 이 시트는 표 아래가
    비어 있고 두 표가 나란히라 덮어써도 안전하다. 정적 표(쿼리 아님)는 QueryTable이 없어 건너뛴다.
    """
    try:
        lo.QueryTable.RefreshStyle = XL_OVERWRITE_CELLS
    except pywintypes.com_error:
        pass


@contextlib.contextmanager
def _with_body(lo: Any) -> Iterator[None]:
    """0행인 쿼리 표는 열 서식을 입힐 본문이 없으므로 임시 행을 두고 서식을 정한 뒤 되돌린다.
    서식은 행이 다시 채워져도 표 정의에 남는다(빌더 ensure_rows와 같은 기법)."""
    added = False
    if lo.DataBodyRange is None:
        try:
            lo.ListRows.Add()
            added = True
        except pywintypes.com_error:
            pass
    try:
        yield
    finally:
        if added:
            try:
                lo.ListRows(1).Delete()
            except pywintypes.com_error:
                pass


def _format_columns(builder: Any, lo: Any, specs: dict) -> None:
    """열 너비·숫자 서식·정렬·글꼴. 기대한 열이 없으면(쿼리 열 변경) 경고만 남기고 있는 열만 입힌다."""
    have = {str(c.Name) for c in lo.ListColumns}
    missing = [n for n in specs if n not in have]
    if missing:
        builder.say(f"[{SHEET}] 경고: {lo.Name}에 없는 열 {', '.join(missing)} — 그 열의 서식은 건너뜁니다")
    for name, sp in specs.items():
        if name not in have:
            continue
        set_col_format(lo, name, nf=sp.get("nf"), width=sp.get("width"), h=sp.get("h"))
        body = col_range(lo, name)
        if body is not None and any(k in sp for k in ("bold", "color", "size")):
            style(body, bold=sp.get("bold"), color=sp.get("color"), size=sp.get("size"))


def _cf(rng: Any, formula: str, *, font: str | None = None, fill_color: str | None = None, bold: bool | None = None,
        italic: bool | None = None, stop: bool = False, top: str | None = None) -> Any:
    """조건부 서식 한 개 추가. 먼저 추가한 규칙이 우선순위가 높다(COM은 항상 맨 뒤에 붙임)."""
    fc = cf_expr(rng, formula, font=font, fill_color=fill_color, bold=bold, stop=stop)
    if italic is not None:
        fc.Font.Italic = italic
    if top is not None:
        b = fc.Borders(-4160)   # xlTop
        b.LineStyle = 1
        b.Color = rgb(top)
    return fc


def _rows(ws: Any, geo: _Geo, n: int, first_col: str | None = None, last_col: str | None = None) -> Any:
    """표 데이터 첫 행부터 n행 범위. first_col·last_col = 열 이름(없으면 표 전체 열)."""
    c_left = geo.cols[first_col] if first_col else geo.left
    c_right = geo.cols[last_col] if last_col else geo.right
    return ws.Range(ws.Cells(geo.first_row, c_left), ws.Cells(geo.first_row + n - 1, c_right))


def _events_cf(builder: Any, ws: Any, geo: _Geo) -> None:
    """이벤트 표 조건부 서식 — 위에서부터 우선순위가 높다.

    상태 경고 > 오류 행 회색 > D-day 글자 > 보유 칩 > 종류 태그(지난 일정 제외) > 보유 행 > 지난 일정 회색
    > 오늘 > 7일 이내 > 날짜가 바뀌는 행의 위 구분선. 지난 일정도 보유 강조·보유 종목명은 남긴다.
    """
    need = [c for c in EVENTS_COLS if c not in geo.cols]
    if need:
        builder.say(f"[{SHEET}] 경고: {EVENTS_TABLE}에 열 {', '.join(need)}이(가) 없어 조건부 서식을 건너뜁니다")
        return
    r0 = geo.first_row
    letter = {c: _col_letter(geo.cols[c]) for c in EVENTS_COLS}
    d, s, h, t = (f"${letter[x]}{r0}" for x in ("일자", "상태", "보유", "종류"))
    d_prev = f"${letter['일자']}{r0 - 1}"
    is_date = f"ISNUMBER({d})"
    is_past = f"AND({is_date},{d}<TODAY())"
    whole = _rows(ws, geo, ROWS_EVENTS)
    # 1) 상태 칸: 오류·이전 데이터는 빨강 굵게 (오류 행 회색 규칙보다 먼저라 색이 남는다)
    _cf(_rows(ws, geo, ROWS_EVENTS, "상태", "상태"), f'=OR(LEFT({s},2)="오류",LEFT({s},2)="이전")',
        font=C["up"], bold=True)
    # 2) 오류 행(일자 없음): 회색 기울임 — 아래 규칙은 적용하지 않음
    _cf(whole, f'=AND(NOT({is_date}),LEFT({s},2)="오류")', font=MUTED_ERR, italic=True, stop=True)
    # 3) D-day 열: 저장된 값 대신 일자에서 오늘 기준으로 다시 계산한 글자("D-3", "D-day", "D+2")를 덧입힌다
    ddr = _rows(ws, geo, ROWS_EVENTS, "D-day", "D-day")
    ws.Activate()
    ddr.Cells(1, 1).Select()
    for n in range(-DDAY_BACK, DDAY_AHEAD + 1):
        text = "D-day" if n == 0 else (f"D-{n}" if n > 0 else f"D+{-n}")
        fc = ddr.FormatConditions.Add(XL_EXPRESSION, 3, f"=AND({is_date},{d}-TODAY()={n})")
        fc.NumberFormat = f'"{text}";"{text}";"{text}"'   # 세 구역 모두 같은 글자 — 저장된 값이 음수여도 '-'가 붙지 않게
        fc.StopIfTrue = False
    # 4) 보유 칩
    _cf(_rows(ws, geo, ROWS_EVENTS, "보유", "보유"), f'={h}="Y"', font=C["white"], fill_color=C["accent"], bold=True)
    # 5) 종류 색 태그 (지난 일정은 회색으로 둠)
    tag_rng = _rows(ws, geo, ROWS_EVENTS, "종류", "종류")
    for name, (bg, fg) in TAG_STYLES.items():
        _cf(tag_rng, f'=AND({t}="{name}",NOT({is_past}))', font=fg, fill_color=bg)
    # 6) 보유 행 옅은 파랑 + 보유 종목명 파랑 굵게
    _cf(whole, f'={h}="Y"', fill_color=HELD_FILL)
    _cf(_rows(ws, geo, ROWS_EVENTS, "종목명", "종목명"), f'={h}="Y"', font=C["accent"], bold=True)
    # 7) 지난 일정 회색
    _cf(whole, f"={is_past}", font=PAST_FONT)
    # 8) 오늘: 행 연한 주황 + 굵게, D-day·일자 칸은 빨강
    _cf(_rows(ws, geo, ROWS_EVENTS, "일자", "D-day"), f"=AND({is_date},{d}=TODAY())", font=C["up"], bold=True)
    _cf(whole, f"=AND({is_date},{d}=TODAY())", fill_color=C["warn_l"], bold=True)
    # 9) 내일~7일 이내: 일자·D-day 파랑 굵게
    _cf(_rows(ws, geo, ROWS_EVENTS, "일자", "D-day"), f"=AND({is_date},{d}>TODAY(),{d}<=TODAY()+7)",
        font=C["accent"], bold=True)
    # 10) 날짜가 바뀌는 첫 행 위에 가는 선 (머리글 바로 아래 첫 행은 제외 — 위 칸이 날짜가 아님)
    _cf(whole, f"=AND({is_date},ISNUMBER({d_prev}),{d}<>{d_prev})", top=SEPARATOR)


def _news_cf(builder: Any, ws: Any, geo: _Geo) -> None:
    """뉴스 표 조건부 서식: 상태 경고 > 오류 행(일시 없음) 회색 > 날짜가 바뀌는 행의 위 구분선."""
    need = [c for c in NEWS_COLS if c not in geo.cols]
    if need:
        builder.say(f"[{SHEET}] 경고: {NEWS_TABLE}에 열 {', '.join(need)}이(가) 없어 조건부 서식을 건너뜁니다")
        return
    r0 = geo.first_row
    letter = {c: _col_letter(geo.cols[c]) for c in NEWS_COLS}
    d, s = f"${letter['일시']}{r0}", f"${letter['상태']}{r0}"
    d_prev = f"${letter['일시']}{r0 - 1}"
    whole = _rows(ws, geo, ROWS_NEWS)
    _cf(_rows(ws, geo, ROWS_NEWS, "상태", "상태"), f'=OR(LEFT({s},2)="오류",LEFT({s},2)="이전")',
        font=C["up"], bold=True)
    _cf(whole, f'=AND(NOT(ISNUMBER({d})),LEFT({s},2)="오류")', font=MUTED_ERR, italic=True, stop=True)
    _cf(whole, f"=AND(ISNUMBER({d}),ISNUMBER({d_prev}),INT({d})<>INT({d_prev}))", top=SEPARATOR)


# ------------------------------------------------------------------------------------------------ 수식 (구역 안내·상태 줄)
def _events_note(t: str = EVENTS_TABLE) -> str:
    return (f'=LET(ev_n,COUNT({t}[일자]),ev_h,COUNTIF({t}[보유],"Y"),'
            f'ev_r,IF(ev_n=0,"",TEXT(MIN({t}[일자]),"mm-dd")&"~"&TEXT(MAX({t}[일자]),"mm-dd")),'
            f'ev_t,IF(COUNT({t}[조회시각])=0,"-",TEXT(MAX({t}[조회시각]),"mm-dd hh:mm")),'
            f'"오늘 "&TEXT(TODAY(),"mm-dd(aaa)")&" · 일정 "&ev_n&"건"&IF(ev_r="","","("&ev_r&")")&" · 그중 보유 종목 "&ev_h&"건"'
            f'&" · 조회 "&ev_t&" · [전체] 버튼으로 갱신")')


def _events_status(t: str = EVENTS_TABLE) -> str:
    return (f'=LET(ev_e,COUNTIF({t}[상태],"오류*"),ev_p,COUNTIF({t}[상태],"이전*"),ev_n,COUNT({t}[일자]),'
            f'IF(ev_p>0,"▲ "&INDEX({t}[상태],MATCH("이전*",{t}[상태],0)),'
            f'IF(ev_e>0,"▲ 조회 오류 "&ev_e&"건 — "&INDEX({t}[상태],MATCH("오류*",{t}[상태],0))'
            f'&IF(ev_e>1," 외 "&(ev_e-1)&"건",""),'
            f'IF(ev_n=0,"○ 데이터 없음 — [전체] 버튼으로 불러오세요",'
            f'"● 정상 — 보유 종목 = 연한 파랑 행 · 지난 일정 = 회색 · 오늘 = 주황 행 · 7일 이내 = 파랑 굵은 일자"))))')


def _news_note(t: str = NEWS_TABLE) -> str:
    return (f'=LET(nw_n,COUNT({t}[일시]),'
            f'nw_k,IFERROR(ROWS(UNIQUE(FILTER({t}[종목코드],ISNUMBER({t}[일시])))),0),'
            f'nw_t,IF(COUNT({t}[조회시각])=0,"-",TEXT(MAX({t}[조회시각]),"mm-dd hh:mm")),'
            f'"뉴스 "&nw_n&"건 · "&nw_k&"종목 · 조회 "&nw_t&" · [모두 새로 고침]으로 갱신")')


def _news_status(t: str = NEWS_TABLE) -> str:
    names = f'IF({t}[종목명]="",{t}[종목코드],{t}[종목명])'
    return (f'=LET(nw_e,COUNTIF({t}[상태],"오류*"),nw_p,COUNTIF({t}[상태],"이전*"),nw_n,COUNT({t}[일시]),'
            f'IF(nw_p>0,"▲ "&INDEX({t}[상태],MATCH("이전*",{t}[상태],0)),'
            f'IF(nw_e>0,"▲ 조회 오류 "&nw_e&"종목("&TEXTJOIN(", ",TRUE,TAKE(FILTER({names},LEFT({t}[상태],2)="오류"),4))'
            f'&IF(nw_e>4," 외 "&(nw_e-4)&"종목","")&") — "&INDEX({t}[상태],MATCH("오류*",{t}[상태],0)),'
            f'IF(nw_n=0,"○ 데이터 없음 — [모두 새로 고침]으로 불러오세요",'
            f'"● 정상 — 최신순 · 같은 기사는 종목마다 한 줄"))))')


def _section_block(builder: Any, ws: Any, geo: _Geo, title: str, note: str | None, status: str | None) -> None:
    """구역 제목(밑줄) + 오른쪽 끝 안내 수식 + 바로 아래 상태 줄(수식, 색은 조건부 서식).

    수식이 참조하는 열이 표에 없으면(쿼리 열 변경) 그 수식만 건너뛰고 경고를 남긴다.
    """
    sec = geo.header_row - 2
    stat = geo.header_row - 1
    lc, rc = _col_letter(geo.left), _col_letter(geo.right)
    section(ws, f"{lc}{sec}:{rc}{sec}", title)
    ws.Rows(sec).RowHeight = 20
    ws.Rows(stat).RowHeight = 16
    ws.Rows(geo.header_row).RowHeight = 20
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
            _cf(cell, f'=LEFT({lc}{stat},1)="▲"', font=C["up"], bold=True)
            _cf(cell, f'=LEFT({lc}{stat},1)="●"', font=C["good"])
        except pywintypes.com_error:
            builder.say(f"[{SHEET}] 경고: '{title}' 구역의 상태 줄 수식을 넣지 못했습니다(표 열 구성 확인)")


def _header_band(builder: Any, ws: Any) -> None:
    """제목 띠 + 메뉴 줄. 머리글 수식이 참조하는 표(tblIndexNow 등)가 없으면 수식 없이 그린다."""
    subtitle = ("이벤트 = 대회·보유·관심 종목의 예탁원 일정(CB·BW 전환·신규상장·증자·합병분할·배당·주주총회) · "
                "뉴스 = 보유·관심 종목 제목 | 이벤트는 [전체], 뉴스는 [모두 새로 고침]에서 갱신")
    span = f"B2:{BAND_LAST_COL}3"
    right = getattr(builder, "HEADER_RIGHT", None)
    status = getattr(builder, "HEADER_STATUS", None)
    try:
        title_bar(ws, "NEWS & EVENTS", subtitle, span, right_formula=right, right_cell=f"{RIGHT_TEXT_COL}2",
                  right2_formula=status, right2_cell=f"{RIGHT_TEXT_COL}3")
    except pywintypes.com_error:
        builder.say(f"[{SHEET}] 경고: 머리글 기준 시각·상태 수식을 넣지 못했습니다(참조 표 없음) — 제목만 표시")
        title_bar(ws, "NEWS & EVENTS", subtitle, span)
    fill(ws.Range(f"B4:{BAND_LAST_COL}4"), "navy2")
    builder.nav_links(ws, 4)
    ws.Rows(4).RowHeight = 16
    ws.Rows(5).RowHeight = 8


def build(builder: Any) -> None:
    """`뉴스·이벤트` 시트를 만든다(표는 LOADS 위치에 이미 적재돼 있어야 함).

    Args:
        builder: build_dashboard.Builder 또는 하네스 대역(StandInBuilder). wb·ws·lo·table_style·say·nav_links·
            HEADER_RIGHT·HEADER_STATUS를 쓴다.

    Raises:
        없음 — 시트나 표가 없으면 경고만 남기고 가능한 부분까지 그린다(COM 서식 오류는 그대로 올라옴).

    Example:
        import pages.page_news_events as p; p.build(builder)
    """
    ws = _sheet(builder)
    if ws is None:
        return
    lo_ev = _table(builder, EVENTS_TABLE)
    lo_nw = _table(builder, NEWS_TABLE)
    style_name = getattr(builder, "table_style", None)

    sheet_setup(ws, bg=False, zoom=85, tab=C["gold"], widths={"A": 1.5})
    ws.Cells.FormatConditions.Delete()   # 다시 부르면 규칙이 겹치지 않게(이 시트는 이 모듈 전용)
    ws.Rows(1).RowHeight = 6
    _header_band(builder, ws)

    geo_ev = _geometry(lo_ev) if lo_ev is not None else _fallback_geo(ws, EVENTS_CELL, len(EVENTS_COLS))
    geo_nw = _geometry(lo_nw) if lo_nw is not None else _fallback_geo(ws, NEWS_CELL, len(NEWS_COLS))
    _section_block(builder, ws, geo_ev, "이벤트 캘린더", _events_note() if lo_ev is not None else None,
                   _events_status() if lo_ev is not None else None)
    _section_block(builder, ws, geo_nw, "뉴스·공시 헤드라인", _news_note() if lo_nw is not None else None,
                   _news_status() if lo_nw is not None else None)
    ws.Columns(_col_letter(geo_ev.right + 1)).ColumnWidth = 3   # 두 표 사이 간격

    for lo, geo, specs, cf in ((lo_ev, geo_ev, EVENTS_SPEC, _events_cf), (lo_nw, geo_nw, NEWS_SPEC, _news_cf)):
        if lo is None:
            continue
        if style_name:
            try:
                lo.TableStyle = style_name
            except pywintypes.com_error:
                builder.say(f"[{SHEET}] 경고: 표 스타일 '{style_name}'을 적용하지 못했습니다({lo.Name})")
        _overwrite_on_refresh(lo)
        with _with_body(lo):
            _format_columns(builder, lo, specs)
        cf(builder, ws, geo)

    freeze(ws, max(geo_ev.header_row, geo_nw.header_row))
    ws.Activate()
    ws.Range("A1").Select()
    builder.say("뉴스·이벤트 시트 완료")
