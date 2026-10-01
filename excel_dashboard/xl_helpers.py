"""Excel COM 헬퍼 (build_dashboard.py 전용).

- 항상 새 Excel 인스턴스(DispatchEx)를 띄워 작업하고, 사용자가 열어 둔 Excel에는 손대지 않습니다.
- 수식은 Formula2(동적 배열)로 입력합니다. 함수명·구분자는 영문/쉼표 기준입니다.
- 확장 페이지용: Q.Pack 3색(기준점 셀 방식), 열 묶음(+/−), 도형 버튼, VBA 코드 텍스트 삽입,
  AccessVBOM 임시 켜기(반드시 원래 상태로 복원), .xlsm 저장.
"""
from __future__ import annotations

import contextlib
import ctypes
import datetime as dt
import os
import re
import time
import types
import winreg

import pythoncom
import win32com.client

# ---------------------------------------------------------------- 상수 (Excel enum)
XL_SRC_EXTERNAL = 0
XL_SRC_RANGE = 1
XL_YES = 1
XL_CMD_SQL = 2
XL_INSERT_DELETE_CELLS = 1
XL_LEFT, XL_CENTER, XL_RIGHT = -4131, -4108, -4152
XL_TOP, XL_VCENTER, XL_BOTTOM = -4160, -4108, -4107
XL_CONTINUOUS = 1
XL_THIN, XL_MEDIUM = 2, -4138
XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT = 7, 8, 9, 10
XL_INSIDE_V, XL_INSIDE_H = 11, 12
XL_SHEET_HIDDEN, XL_SHEET_VERY_HIDDEN = 0, 2
XL_CELL_VALUE, XL_EXPRESSION = 1, 2
XL_GREATER_EQUAL, XL_LESS_EQUAL, XL_LESS, XL_GREATER = 7, 8, 6, 5
XL_VALIDATE_LIST = 3
# 차트
XL_LINE, XL_LINE_MARKERS, XL_AREA = 4, 65, 1
XL_COLUMN_CLUSTERED, XL_BAR_CLUSTERED, XL_COLUMN_STACKED = 51, 57, 52
XL_DOUGHNUT, XL_XY_SCATTER = -4120, -4169
XL_CATEGORY, XL_VALUE = 1, 2
XL_PRIMARY, XL_SECONDARY = 1, 2
XL_CATEGORY_SCALE = 2
XL_LEGEND_TOP, XL_LEGEND_BOTTOM, XL_LEGEND_RIGHT = -4160, -4107, -4152
XL_TICK_LABEL_LOW = -4134
# 조건부 서식 색 척도 기준 형식 (XlConditionValueTypes)
XL_COND_VALUE_FORMULA = 4
# 열 묶음 요약 열 위치 (XlSummaryColumn)
XL_SUMMARY_ON_LEFT, XL_SUMMARY_ON_RIGHT = -4131, -4152
# 도형·버튼
MSO_SHAPE_ROUNDED_RECTANGLE = 5
MSO_ALIGN_CENTER, MSO_ANCHOR_MIDDLE = 2, 3
MSO_TRUE, MSO_FALSE = -1, 0
XL_MOVE_AND_SIZE, XL_MOVE, XL_FREE_FLOATING = 1, 2, 3
# VBA·저장 형식
VBEXT_CT_STD_MODULE = 1
XL_OPENXML_WORKBOOK_MACRO_ENABLED = 52   # .xlsm
# VBA 프로젝트 접근 신뢰 설정 (HKCU 아래). 빌드 동안만 1로 켜고 원래 상태(값 없음 포함)로 되돌린다.
VBOM_KEY = r"Software\Microsoft\Office\16.0\Excel\Security"
VBOM_VALUE = "AccessVBOM"


def rgb(hex_color: str) -> int:
    """'#RRGGBB' → Excel BGR 정수."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536


# ---------------------------------------------------------------- 팔레트 (디자인 토큰)
C = {
    "navy": "#0F203D", "navy2": "#1B3358", "accent": "#1F6FEB", "accent_l": "#DCE9FD",
    "bg": "#F4F6F9", "card": "#FFFFFF", "line": "#D9DEE7", "line2": "#E9EDF3",
    "text": "#1F2933", "muted": "#6B7785", "white": "#FFFFFF",
    "up": "#D9322F", "down": "#1864C8", "up_l": "#FDECEC", "down_l": "#E8F1FD",
    "good": "#2F9E44", "good_l": "#E6F4EA", "warn": "#E8890C", "warn_l": "#FFF4E0",
    "gold": "#B7791F", "gray_l": "#F1F3F5", "input": "#FFFBE6",
}

FONT = "맑은 고딕"

# 숫자 서식 (Color10=상승 빨강, Color11=하락 파랑 — 통합문서 팔레트에서 지정)
NF = {
    "krw": "#,##0",
    "krw_pl": "[Color10]+#,##0;[Color11]-#,##0;0",
    "pct": "0.00%",
    "pct1": "0.0%",
    "pct_pl": "[Color10]+0.00%;[Color11]-0.00%;0.00%",
    "pct_arrow": "[Color10]▲0.00%;[Color11]▼0.00%;0.00%",
    "num2": "#,##0.00",
    "num1": "#,##0.0",
    "num0": "#,##0",
    "eok": "#,##0\"억\"",
    "eok_pl": "[Color10]+#,##0\"억\";[Color11]-#,##0\"억\";0\"억\"",
    "date": "yyyy-mm-dd",
    "mmdd": "mm/dd",
    "dt": "yyyy-mm-dd hh:mm",
    "ratio2": "0.00",
    "x": "0.0\"배\"",
    "bp_pl": "[Color10]+0.0\"bp\";[Color11]-0.0\"bp\";0.0\"bp\"",
    "text": "@",
}

# Q.Pack 3색: 낮음 초록 → 중앙값 노랑 → 높음 빨강
QPACK_LOW, QPACK_MID, QPACK_HIGH = "#63BE7B", "#FFEB84", "#F8696B"
QPACK_POINTS = ("min", "median", "max")   # 열별 색의 기준점(최소·중앙값·최대 — docs/business-rules.md의 '색 비교 기준')
QPACK_STRIP_PCTS = (5, 50, 95)            # 최근 20일 줄무늬 블록의 기준 백분위 (조정 가능)


# ---------------------------------------------------------------- Excel 인스턴스
def new_excel(visible: bool = False):
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")  # 기존 Excel과 분리된 새 프로세스
    xl.Visible = visible
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    xl.EnableEvents = False
    return xl


def excel_pid(xl) -> int | None:
    try:
        import win32process
        return win32process.GetWindowThreadProcessId(xl.Hwnd)[1]
    except Exception:
        return None


def quit_excel(xl, pid: int | None = None, wait: float = 15.0):
    """이 스크립트가 띄운 Excel만 종료. Quit 후에도 남아 있으면(COM 참조 잔존 등) 그 PID만 강제 종료."""
    pid = pid or excel_pid(xl)
    try:
        xl.Quit()
    except Exception:
        pass
    if not pid:
        return
    import subprocess
    t0 = time.time()
    while time.time() - t0 < wait:
        r = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True)
        if str(pid) not in r.stdout:
            return
        time.sleep(0.5)
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)


def set_palette(wb, index: int, hex_color: str):
    """통합문서 56색 팔레트 지정 (숫자서식 [ColorN]에 사용). Colors는 인덱스 속성이라 Invoke로 설정."""
    dispid = wb._oleobj_.GetIDsOfNames("Colors")
    wb._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, False, index, rgb(hex_color))


def excel_serial(d) -> float:
    """date/datetime → Excel 일련번호 (pywin32의 시간대 변환을 피하기 위해 Value2로 입력)."""
    if isinstance(d, dt.datetime):
        return (d - dt.datetime(1899, 12, 30)).total_seconds() / 86400.0
    if isinstance(d, dt.date):
        return float((d - dt.date(1899, 12, 30)).days)
    raise TypeError(d)


# ---------------------------------------------------------------- 셀 쓰기/서식
def put(ws, addr, value=None, *, formula=None, nf=None, bold=None, size=None, color=None, fill=None,
        h=None, v=None, italic=None, wrap=None, font=None, indent=None):
    r = ws.Range(addr)
    if formula is not None:
        r.Formula2 = formula
    elif value is not None:
        if isinstance(value, (dt.date, dt.datetime)):
            r.Value2 = excel_serial(value)
            if nf is None:
                nf = NF["date"]
        else:
            r.Value = value
    style(r, nf=nf, bold=bold, size=size, color=color, fill=fill, h=h, v=v, italic=italic, wrap=wrap,
          font=font, indent=indent)
    return r


_LOCAL_COLOR = [("[Color", "[색"), ("[Red]", "[빨강]"), ("[Blue]", "[파랑]"), ("General", "G/표준")]


def set_nf(r, nf: str):
    """숫자서식 지정. 한국어 Excel은 색 코드도 현지화([색10])를 요구하므로 실패 시 변환해 재시도."""
    try:
        r.NumberFormat = nf
    except Exception:
        loc = nf
        for a, b in _LOCAL_COLOR:
            loc = loc.replace(a, b)
        r.NumberFormatLocal = loc


def style(r, *, nf=None, bold=None, size=None, color=None, fill=None, h=None, v=None, italic=None,
          wrap=None, font=None, indent=None):
    if nf is not None:
        set_nf(r, nf)
    f = r.Font
    if font is not None:
        f.Name = font
    if bold is not None:
        f.Bold = bold
    if italic is not None:
        f.Italic = italic
    if size is not None:
        f.Size = size
    if color is not None:
        f.Color = rgb(color)
    if fill is not None:
        r.Interior.Color = rgb(fill)
    if h is not None:
        r.HorizontalAlignment = h
    if v is not None:
        r.VerticalAlignment = v
    if wrap is not None:
        r.WrapText = wrap
    if indent is not None:
        r.IndentLevel = indent
    return r


def border(r, color="line", weight=XL_THIN, edges=(XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT)):
    for e in edges:
        b = r.Borders(e)
        b.LineStyle = XL_CONTINUOUS
        b.Weight = weight
        b.Color = rgb(C[color])


def bottom_line(r, color="line", weight=XL_THIN):
    border(r, color, weight, edges=(XL_EDGE_BOTTOM,))


def fill(r, color):
    r.Interior.Color = rgb(C[color] if color in C else color)


# ---------------------------------------------------------------- 표 (ListObject)
def write_table(ws, top: int, left: int, headers, rows, name: str, *, text_cols=(), date_cols=(), style_name="TableStyleLight1"):
    """입력용 엑셀 표 생성. text_cols/date_cols: 헤더명 목록."""
    for j, hname in enumerate(headers):
        ws.Cells(top, left + j).Value = hname
    n = max(1, len(rows))
    for j, hname in enumerate(headers):
        col = ws.Range(ws.Cells(top + 1, left + j), ws.Cells(top + n, left + j))
        if hname in text_cols:
            col.NumberFormat = "@"
        if hname in date_cols:
            set_nf(col, NF["date"])
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = ws.Cells(top + 1 + i, left + j)
            if val is None:
                continue
            if isinstance(val, str) and val.startswith("="):
                c.Formula2 = val
            elif isinstance(val, (dt.date, dt.datetime)):
                c.Value2 = excel_serial(val)
            else:
                c.Value = val
    rng = ws.Range(ws.Cells(top, left), ws.Cells(top + n, left + len(headers) - 1))
    lo = ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES)
    lo.Name = name
    lo.TableStyle = style_name
    return lo


def ensure_table_style(wb, name: str = "PM Navy"):
    """네이비 머리글 + 옅은 줄무늬 표 스타일 생성 (실패 시 None)."""
    try:
        return wb.TableStyles(name)
    except Exception:
        pass
    try:
        ts = wb.TableStyles.Add(name)
        ts.ShowAsAvailableTableStyle = True
        hdr = ts.TableStyleElements(1)            # xlHeaderRow
        hdr.Interior.Color = rgb(C["navy"])
        hdr.Font.Color = rgb("#FFFFFF")
        hdr.Font.Bold = True
        whole = ts.TableStyleElements(0)          # xlWholeTable
        b = whole.Borders(12)                     # xlInsideHorizontal
        b.Color = rgb(C["line2"])
        b.LineStyle = XL_CONTINUOUS
        b2 = whole.Borders(XL_EDGE_BOTTOM)
        b2.Color = rgb(C["line"])
        b2.LineStyle = XL_CONTINUOUS
        s1 = ts.TableStyleElements(5)             # xlRowStripe1
        s1.Interior.Color = rgb("#F7F9FC")
        return ts
    except Exception:
        return None


def add_query(wb, name: str, m_code: str, description: str = ""):
    for q in wb.Queries:
        if q.Name == name:
            q.Formula = m_code
            return q
    return wb.Queries.Add(name, m_code, description)


def load_query(ws, query: str, cell: str, table_name: str, *, background=True, style_name="TableStyleLight1"):
    """Power Query를 시트의 표로 로드 (새로 고침은 호출자가 수행)."""
    conn = ("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
            f"Location={query};Extended Properties=\"\"")
    lo = ws.ListObjects.Add(XL_SRC_EXTERNAL, conn, None, XL_YES, ws.Range(cell))
    qt = lo.QueryTable
    qt.CommandType = XL_CMD_SQL
    qt.CommandText = f"SELECT * FROM [{query}]"
    qt.RowNumbers = False
    qt.FillAdjacentFormulas = False
    qt.PreserveFormatting = True
    qt.RefreshOnFileOpen = False
    qt.BackgroundQuery = background
    qt.RefreshStyle = XL_INSERT_DELETE_CELLS
    qt.SavePassword = False
    qt.SaveData = True
    qt.AdjustColumnWidth = False
    qt.RefreshPeriod = 0
    qt.PreserveColumnInfo = True
    lo.Name = table_name
    lo.TableStyle = style_name
    return lo


def refresh(lo, label=""):
    """표를 동기 새로 고침하고 (성공, 메시지, 행수) 반환."""
    t0 = time.time()
    try:
        lo.QueryTable.Refresh(False)
        n = lo.ListRows.Count
        return True, f"{label or lo.Name}: {n}행 ({time.time() - t0:.1f}s)", n
    except Exception as e:  # pywintypes.com_error
        msg = getattr(e, "excepinfo", None)
        msg = msg[2] if msg and len(msg) > 2 else str(e)
        return False, f"{label or lo.Name}: 오류 → {msg}", 0


def add_calc_column(lo, name: str, formula: str, nf=None, width=None, after: str | None = None):
    """계산열 추가 (구조적 참조 수식). after=열이름 이면 그 열 바로 뒤에 삽입 (새로 고침 후에도 위치 유지)."""
    if after:
        col = lo.ListColumns.Add(lo.ListColumns(after).Index + 1)
    else:
        col = lo.ListColumns.Add()
    col.Name = name
    rng = col.DataBodyRange
    if rng is not None:
        try:
            rng.Formula2 = formula
        except Exception as e:
            raise RuntimeError(f"계산열 수식 오류: {lo.Name}[{name}] {formula}") from e
        if nf:
            set_nf(rng, nf)
    if width:
        col.Range.ColumnWidth = width
    return col


def col_range(lo, name: str):
    return lo.ListColumns(name).DataBodyRange


def set_col_format(lo, name: str, nf=None, width=None, h=None, bold=None, color=None):
    c = lo.ListColumns(name)
    if nf is not None and c.DataBodyRange is not None:
        set_nf(c.DataBodyRange, nf)
    if width is not None:
        c.Range.ColumnWidth = width
    if h is not None:
        c.Range.HorizontalAlignment = h
    if (bold is not None or color is not None) and c.DataBodyRange is not None:
        style(c.DataBodyRange, bold=bold, color=color)


# ---------------------------------------------------------------- 시트 레이아웃
def sheet_setup(ws, *, zoom=85, gridlines=False, tab=None, bg=True, widths=None, default_width=None):
    ws.Activate()
    win = ws.Application.ActiveWindow
    win.DisplayGridlines = gridlines
    win.Zoom = zoom
    if tab:
        ws.Tab.Color = rgb(C[tab] if tab in C else tab)
    if default_width:
        ws.Cells.ColumnWidth = default_width
    if widths:
        for colspec, w in widths.items():
            ws.Columns(colspec).ColumnWidth = w
    if bg:
        ws.Cells.Interior.Color = rgb(C["bg"])
    ws.Cells.Font.Name = FONT
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = XL_VCENTER


def freeze(ws, row: int, col: int = 0):
    ws.Activate()
    win = ws.Application.ActiveWindow
    win.FreezePanes = False
    win.SplitRow = row
    win.SplitColumn = col
    win.FreezePanes = True


def _set_summary_column(ws, value: int):
    """열 묶음 요약 위치 설정. 활성 셀이 표(ListObject) 안이면 Excel이 설정을 거부하므로,
    그때는 사용 범위 바로 아래 A열(어떤 표에도 속하지 않는 셀)을 잠시 선택하고 다시 시도한다."""
    outline = ws.Outline
    if int(outline.SummaryColumn) == value:
        return
    try:
        outline.SummaryColumn = value
        return
    except Exception:
        pass
    ws.Activate()
    prev = ws.Application.ActiveCell
    used = ws.UsedRange
    ws.Cells(used.Row + used.Rows.Count, 1).Select()
    outline.SummaryColumn = value
    try:
        prev.Select()
    except Exception:
        pass


def group_columns(ws, cols: str, *, collapsed: bool = True, summary_left: bool = True):
    """열 묶음(+/−) 만들기. 기본은 접힌 상태(열 숨김)로 두고, 사용자가 [+]로 펼친다.

    Args:
        ws: 대상 시트.
        cols: 묶을 열 주소. 예: "X:Z", "AC" (한 열).
        collapsed: True면 만들자마자 접는다(요약 열의 ShowDetail=False, [−]를 누른 것과 같음).
        summary_left: True면 요약 열(+/− 단추가 붙는 열)이 묶음 바로 왼쪽 열.
            머리글·핵심 열이 왼쪽에 있고 세부 열이 오른쪽으로 이어지는 표에 맞다.
            시트 전체 설정(Outline.SummaryColumn)이므로 한 시트에서는 같은 값으로 쓴다.

    Returns:
        묶은 열 전체(Range).

    주의:
        - 같은 수준의 묶음끼리 바로 붙어 있으면 Excel이 하나의 묶음으로 합친다. 따로 접고 펴려면
          묶음 사이에 묶지 않은 열(요약 열)을 하나 둔다.
        - summary_left=True인데 묶음이 A열에서 시작하면 요약 열이 없으므로 그냥 숨긴다.

    Example:
        group_columns(ws, "AC:AI")                    # 분류 세부 7열, 기본 접힘
        group_columns(ws, "BA:BC", collapsed=False)   # 펼친 채로 묶기만
    """
    if ":" not in cols:
        cols = f"{cols}:{cols}"
    rng = ws.Range(cols).EntireColumn
    _set_summary_column(ws, XL_SUMMARY_ON_LEFT if summary_left else XL_SUMMARY_ON_RIGHT)
    rng.Group()
    if collapsed:
        summary = rng.Column - 1 if summary_left else rng.Column + rng.Columns.Count
        try:
            if summary < 1:
                raise ValueError("요약 열 없음")
            ws.Columns(summary).ShowDetail = False
        except Exception:
            rng.Hidden = True
    return rng


def title_bar(ws, title: str, subtitle: str, span: str, *, right_formula=None, right_cell=None, right2_formula=None,
              right2_cell=None):
    """상단 네이비 제목 바 (2~3행)."""
    r = ws.Range(span)
    fill(r, "navy")
    first = r.Cells(1, 1)
    put(ws, first.Address, title, bold=True, size=16, color=C["white"])
    put(ws, r.Cells(2, 1).Address, subtitle, size=9, color="#B8C4D6")
    ws.Rows(r.Row).RowHeight = 30
    ws.Rows(r.Row + 1).RowHeight = 18
    if right_formula and right_cell:
        put(ws, right_cell, formula=right_formula, size=9, color=C["white"], h=XL_RIGHT)
    if right2_formula and right2_cell:
        put(ws, right2_cell, formula=right2_formula, size=9, color="#FFD580", h=XL_RIGHT, bold=True)


def section(ws, addr_span: str, text: str, note: str | None = None, note_cell: str | None = None):
    """섹션 제목 (굵은 글씨 + 하단 강조선)."""
    r = ws.Range(addr_span)
    put(ws, r.Cells(1, 1).Address, text, bold=True, size=11, color=C["navy"])
    border(r, "navy2", XL_MEDIUM, edges=(XL_EDGE_BOTTOM,))
    if note and note_cell:
        put(ws, note_cell, note, size=8, color=C["muted"], h=XL_RIGHT)


def card(ws, span: str, label: str, value_formula: str, sub_formula: str | None, *, nf=None, value_size=16,
         value_color=None, sub_nf=None):
    """KPI 카드: span = 'B6:D9' (4행: 라벨/값/값/보조)."""
    r = ws.Range(span)
    fill(r, "card")
    border(r, "line")
    top = r.Row
    left = r.Column
    ncols = r.Columns.Count
    lab = ws.Cells(top, left)
    put(ws, lab.Address, label, size=9, color=C["muted"], indent=1)
    val = ws.Range(ws.Cells(top + 1, left), ws.Cells(top + 2, left + ncols - 1))
    val.Merge()
    put(ws, val.Address, formula=value_formula, nf=nf, bold=True, size=value_size,
        color=value_color or C["text"], h=XL_LEFT, v=XL_VCENTER, indent=1)
    if sub_formula:
        sub = ws.Range(ws.Cells(top + 3, left), ws.Cells(top + 3, left + ncols - 1))
        sub.Merge()
        put(ws, sub.Address, formula=sub_formula, nf=sub_nf, size=8, color=C["muted"], h=XL_LEFT, indent=1)
    return r


def add_names(wb, names: dict[str, str]):
    for n, ref in names.items():
        try:
            wb.Names(n).Delete()
        except Exception:
            pass
        try:
            wb.Names.Add(Name=n, RefersTo=ref)
        except Exception as e:
            raise RuntimeError(f"이름 정의 실패: {n} = {ref}") from e


def validation_list(r, source: str, show_error=True):
    dv = r.Validation
    dv.Delete()
    dv.Add(XL_VALIDATE_LIST, 1, 1, source)
    dv.IgnoreBlank = True
    dv.InCellDropdown = True
    dv.ShowError = show_error


def hyperlink(ws, cell: str, target: str, text: str):
    r = ws.Range(cell)
    r.Value = text
    ws.Hyperlinks.Add(r, "", target, "", text)   # 위치 인수: Anchor, Address, SubAddress, ScreenTip, TextToDisplay
    style(r, color="#B8C4D6", size=9)
    r.Font.Underline = False


# ---------------------------------------------------------------- 조건부 서식
def cf_expr(r, formula: str, *, font=None, fill_color=None, bold=None, stop=False):
    """수식 기반 조건부 서식. 상대참조는 '활성 셀' 기준으로 해석되므로 범위의 첫 셀을 먼저 선택."""
    ws = r.Worksheet
    ws.Activate()
    r.Cells(1, 1).Select()
    fc = r.FormatConditions.Add(XL_EXPRESSION, 3, formula)  # (Type, Operator(무시됨), Formula1) — 위치 인수만 동작
    if font:
        fc.Font.Color = rgb(font)
    if bold is not None:
        fc.Font.Bold = bold
    if fill_color:
        fc.Interior.Color = rgb(fill_color)
    fc.StopIfTrue = stop
    return fc


def cf_databar(r, color="#1F6FEB"):
    db = r.FormatConditions.AddDatabar()
    db.BarColor.Color = rgb(color)
    try:
        db.BarFillType = 1  # gradient=1, solid=0
    except Exception:
        pass
    return db


def cf_heat(r):
    """0 기준 3색 스케일: 하락=파랑, 0=흰색, 상승=빨강 (국내 관례)."""
    cs = r.FormatConditions.AddColorScale(3)
    cs.ColorScaleCriteria(1).Type = 1  # lowest value
    cs.ColorScaleCriteria(1).FormatColor.Color = rgb("#5B8FD9")
    cs.ColorScaleCriteria(2).Type = 0  # number
    cs.ColorScaleCriteria(2).Value = 0
    cs.ColorScaleCriteria(2).FormatColor.Color = rgb("#FFFFFF")
    cs.ColorScaleCriteria(3).Type = 2  # highest value
    cs.ColorScaleCriteria(3).FormatColor.Color = rgb("#E86A67")
    return cs


# ---------------------------------------------------------------- Q.Pack 3색 (기준점 셀 방식)
# 색을 칠할 '적용 범위'와 기준점을 계산할 '기준 범위'를 분리한다. 기준점(최소·중앙·최대 또는 백분위)은
# 시트의 셀 3개에 수식으로 두고, 색 척도의 세 기준을 그 셀 참조로 건다. 조건부 서식 수식에는 구조적 참조·
# FILTER를 쓸 수 없지만 일반 셀 수식에는 쓸 수 있으므로, 표가 커지거나 줄어도 기준점이 따라간다.
def qpack_anchor_formula(ref: str, cond: str | None = None, point: str | float = "median") -> str:
    """기준 범위에서 기준점 하나를 계산하는 셀 수식(Formula2용)을 만든다.

    Args:
        ref: 기준 범위. 구조적 참조를 권장한다(예: 'tblCompany[1D]', 'tblCompany[[D-19]:[D0]]').
            '-'·공백 등이 든 열 이름은 tbl[[이름]] 형식으로 쓴다.
        cond: 행 조건(예: 'tblCompany[유니버스]="대회"'). ref와 행 수가 같아야 하며, ref가 여러 열이면
            열 방향으로 퍼져 같은 행 전체에 적용된다. None이면 모든 행.
        point: 'min' / 'median' / 'max' 또는 백분위 숫자 0~100 (PERCENTILE.INC).

    Returns:
        예: '=LET(q_v,IF(ISNUMBER(tblT[1D])*(tblT[유니버스]="대회"),tblT[1D]),IF(COUNT(q_v),MIN(q_v),""))'
        숫자 칸만 센다(공란·""·글자·오류 제외). 조건에 맞는 숫자가 없으면 ""를 돌려주며,
        기준점이 ""이면 그 색 척도는 아무 칸도 칠하지 않는다.

    Example:
        qpack_anchor_formula("tblCompany[1D]", 'tblCompany[유니버스]="대회"', "min")
        qpack_anchor_formula("tblCompany[[D-19]:[D0]]", 'tblCompany[유니버스]="대회"', 95)
    """
    if point == "min":
        agg = "MIN(q_v)"
    elif point == "median":
        agg = "MEDIAN(q_v)"
    elif point == "max":
        agg = "MAX(q_v)"
    elif isinstance(point, (int, float)) and not isinstance(point, bool) and 0 <= point <= 100:
        agg = f"PERCENTILE.INC(q_v,{point / 100:g})"
    else:
        raise ValueError(f"기준점 종류 오류: {point!r} ('min'/'median'/'max' 또는 0~100)")
    mask = f"ISNUMBER({ref})" if cond is None else f"ISNUMBER({ref})*({cond})"
    return f'=LET(q_v,IF({mask},{ref}),IF(COUNT(q_v),{agg},""))'


def _resolve_range(ws, addr: str):
    """'X2:X4' 또는 "'시트'!B2:D2" 주소를 Range로 (시트 이름이 있으면 같은 통합문서의 그 시트)."""
    if "!" in addr:
        sheet, a = addr.rsplit("!", 1)
        sheet = sheet.strip()
        if sheet.startswith("'") and sheet.endswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        return ws.Parent.Worksheets(sheet).Range(a)
    return ws.Range(addr)


def _anchor_cells(ws, anchors) -> list:
    """기준점 셀 3개(최소·중앙·최대 순)를 Range 목록으로. anchors = Range / 주소 문자열 / 셀 3개 목록."""
    if isinstance(anchors, str):
        anchors = _resolve_range(ws, anchors)
    if isinstance(anchors, (list, tuple)):
        cells = [_resolve_range(ws, a) if isinstance(a, str) else a for a in anchors]
        if any(c.Count != 1 for c in cells):
            raise ValueError("기준점 목록의 각 항목은 셀 하나여야 합니다")
    else:
        cells = [anchors.Cells(i) for i in range(1, anchors.Count + 1)]
    if len(cells) != 3:
        raise ValueError(f"기준점 셀은 3개여야 합니다(받은 수 {len(cells)})")
    return cells


def _cf_ref(cell, target_ws) -> str:
    """조건부 서식 기준에 넣을 절대 참조 수식. 다른 시트면 시트 이름을 붙인다(같은 통합문서만)."""
    sh = cell.Worksheet
    if sh.Parent.Name != target_ws.Parent.Name:
        raise ValueError("기준점 셀은 적용 범위와 같은 통합문서에 있어야 합니다")
    if sh.Name == target_ws.Name:
        return "=" + cell.Address
    return "='" + sh.Name.replace("'", "''") + "'!" + cell.Address


def write_qpack_anchors(ws, anchors, ref: str, cond: str | None = None, points=QPACK_POINTS) -> list:
    """기준점 셀 3개에 기준점 수식을 써 넣는다 (보통은 cf_qpack(ref=...)가 대신 호출).

    Args:
        ws: anchors가 주소 문자열일 때 기준이 되는 시트.
        anchors: 기준점 셀 3개 — Range('X2:X4'), 주소 문자열, 또는 셀 3개 목록. 순서 = 낮음·중앙·높음.
        ref, cond: qpack_anchor_formula와 같음.
        points: 세 기준점 종류. 기본 ('min','median','max'), 줄무늬는 (5, 50, 95) 등.

    Returns:
        기준점 셀 Range 3개의 목록.

    Example:
        write_qpack_anchors(ws, "M4:M6", "tblCompany[1D]", 'tblCompany[유니버스]="대회"')
    """
    if len(points) != 3:
        raise ValueError("points는 3개여야 합니다")
    cells = _anchor_cells(ws, anchors)
    for c, p in zip(cells, points):
        c.Formula2 = qpack_anchor_formula(ref, cond, p)
    return cells


def _require_range(r):
    if r is None:
        raise ValueError("적용 범위가 없습니다(행이 없는 표의 DataBodyRange?)")


def _color_scale3(r, cells, colors):
    """3색 척도의 세 기준을 '수식 = 기준점 셀 참조'로 건다."""
    ws = r.Worksheet
    try:  # 한국어 Excel 규칙: 대상 첫 셀을 선택한 뒤 추가 (참조가 절대라 결과에는 영향 없음)
        ws.Activate()
        r.Cells(1, 1).Select()
    except Exception:
        pass
    cs = r.FormatConditions.AddColorScale(3)
    for k, (cell, color) in enumerate(zip(cells, colors), start=1):
        crit = cs.ColorScaleCriteria(k)
        crit.Type = XL_COND_VALUE_FORMULA
        crit.Value = _cf_ref(cell, ws)
        crit.FormatColor.Color = rgb(color)
    # COM으로 추가한 규칙은 우선순위가 맨 뒤라, 먼저 걸어 둔 채우기 규칙(예: 보유 행 강조)이 색을 가린다.
    # 맨 앞으로 올려 겹치는 칸에서는 항상 Q.Pack 색이 보이게 한다.
    cs.SetFirstPriority()
    return cs


def cf_qpack(r, anchors, *, ref: str | None = None, cond: str | None = None, reverse: bool = False):
    """① 열별 Q.Pack 3색: 낮음 #63BE7B(초록) → 중앙값 #FFEB84(노랑) → 높음 #F8696B(빨강).

    색의 세 기준은 기준점 셀(anchors)을 참조한다. ref를 주면 그 '기준 범위'(·cond 행 조건)에서
    최소·중앙값·최대를 계산하는 수식을 anchors에 먼저 써 넣고, 생략하면 anchors의 기존 값을 그대로 쓴다.
    기준 범위 밖의 행(예: 유니버스 밖)도 같은 척도로 칠해지며, 기준점보다 작거나 크면 끝 색이 된다.
    공란·""·글자 칸은 칠하지 않는다. 규칙은 맨 앞 우선순위로 올려 다른 채우기 규칙보다 먼저 보이게 한다.
    표 열 전체(DataBodyRange)에 걸면 쿼리 새로 고침으로 행 수가 바뀌어도 서식 범위가 따라간다.

    Args:
        r: 적용 범위(예: col_range(lo, "1D") — 표 열 전체).
        anchors: 기준점 셀 3개(낮음·중앙·높음 순). Range, 주소('M4:M6', "'_calc'!B2:D2"), 셀 3개 목록.
            다른 시트(숨김 시트 가능)에 두어도 된다 — 그때 ref는 구조적 참조나 시트 이름이 붙은 주소로.
        ref: 기준 범위(구조적 참조 권장). None이면 anchors를 쓰지 않음.
        cond: 기준 행 조건(예: 'tblCompany[유니버스]="대회"').
        reverse: True면 낮음 빨강·높음 초록(값이 낮을수록 좋은 열).

    Returns:
        ColorScale 조건부 서식 객체.

    Example:
        cf_qpack(col_range(lo, "1D"), "M4:M6", ref="tblCompany[1D]", cond='tblCompany[유니버스]="대회"')
    """
    _require_range(r)
    cells = _anchor_cells(r.Worksheet, anchors)
    if ref is not None:
        write_qpack_anchors(r.Worksheet, cells, ref, cond, QPACK_POINTS)
    colors = (QPACK_HIGH, QPACK_MID, QPACK_LOW) if reverse else (QPACK_LOW, QPACK_MID, QPACK_HIGH)
    return _color_scale3(r, cells, colors)


def cf_qpack_block(r, anchors, *, ref: str | None = None, cond: str | None = None, pcts=QPACK_STRIP_PCTS,
                   reverse: bool = False, hide_values: bool = False, width: float | None = None):
    """② 2차원 블록(최근 20일 줄무늬 등) 전체를 하나의 Q.Pack 척도로 칠한다.

    기준점 = 블록의 기준 행(cond) 전체 숫자의 백분위(기본 5·50·95 — 극단값에 색이 몰리지 않게).
    ref를 주면 anchors에 백분위 수식을 써 넣고, 생략하면 anchors의 기존 값을 쓴다.

    Args:
        r: 적용 블록(예: 표의 D-19~D0 20열 범위).
        anchors: 기준점 셀 3개(낮음·중앙·높음 순).
        ref: 블록과 같은 모양의 기준 범위(예: 'tblCompany[[D-19]:[D0]]').
        cond: 기준 행 조건(열 방향으로 퍼져 적용).
        pcts: 세 백분위(0~100, 오름차순). 기본 (5, 50, 95).
        reverse: True면 낮음 빨강·높음 초록.
        hide_values: True면 숫자를 숨긴다(숫자 서식 ';;;', 색만 보임).
        width: 열 너비(좁은 칸). None이면 그대로.

    Returns:
        ColorScale 조건부 서식 객체.

    Example:
        cf_qpack_block(strip, "AB4:AB6", ref="tblCompany[[D-19]:[D0]]",
                       cond='tblCompany[유니버스]="대회"', hide_values=True, width=2.5)
    """
    if len(pcts) != 3 or not (0 <= pcts[0] < pcts[1] < pcts[2] <= 100):
        raise ValueError(f"pcts는 0~100 사이 오름차순 3개여야 합니다: {pcts!r}")
    _require_range(r)
    cells = _anchor_cells(r.Worksheet, anchors)
    if ref is not None:
        write_qpack_anchors(r.Worksheet, cells, ref, cond, tuple(pcts))
    colors = (QPACK_HIGH, QPACK_MID, QPACK_LOW) if reverse else (QPACK_LOW, QPACK_MID, QPACK_HIGH)
    cs = _color_scale3(r, cells, colors)
    if hide_values:
        set_nf(r, ";;;")
    if width is not None:
        r.EntireColumn.ColumnWidth = width
    return cs


# ---------------------------------------------------------------- 차트
def chart(ws, span: str, ctype: int, title: str | None = None):
    r = ws.Range(span)
    shp = ws.Shapes.AddChart2(-1, ctype, r.Left, r.Top, r.Width, r.Height)
    ch = shp.Chart
    while ch.SeriesCollection().Count > 0:
        ch.SeriesCollection(1).Delete()
    ch.PlotVisibleOnly = False  # 숨김 열(차트 도우미)의 데이터도 표시
    try:
        ch.DisplayBlanksAs = 1  # xlNotPlotted: 빈 셀은 0이 아니라 끊어서 표시
    except Exception:
        pass
    ch.ChartArea.Format.Line.Visible = False
    ch.ChartArea.Format.Fill.ForeColor.RGB = rgb("#FFFFFF")
    ch.ChartArea.Font.Name = FONT
    ch.ChartArea.Font.Size = 8
    ch.ChartArea.Font.Color = rgb(C["muted"])
    if title:
        ch.HasTitle = True
        ch.ChartTitle.Text = title
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True
        ch.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(C["navy"])
        ch.ChartTitle.Left = 6
    else:
        ch.HasTitle = False
    return shp, ch


def series(ch, name, x_ref, y_ref, *, color=None, weight=None, kind=None, axis=XL_PRIMARY, fill_color=None,
           transparency=None):
    s = ch.SeriesCollection().NewSeries()
    s.Name = name
    s.Values = y_ref
    if x_ref is not None:
        s.XValues = x_ref
    if kind is not None:
        s.ChartType = kind
    if axis == XL_SECONDARY:
        s.AxisGroup = XL_SECONDARY
    if color:
        try:
            s.Format.Line.Visible = True
            s.Format.Line.ForeColor.RGB = rgb(color)
        except Exception:
            pass
    if weight:
        try:
            s.Format.Line.Weight = weight
        except Exception:
            pass
    if fill_color:
        s.Format.Fill.Visible = True
        s.Format.Fill.Solid()
        s.Format.Fill.ForeColor.RGB = rgb(fill_color)
        if transparency is not None:
            s.Format.Fill.Transparency = transparency
    return s


def style_axes(ch, *, y_nf=None, x_nf=None, category=True, gridlines=True, y2_nf=None, legend=XL_LEGEND_TOP):
    try:
        ax = ch.Axes(XL_CATEGORY)
        if category:
            ax.CategoryType = XL_CATEGORY_SCALE
        if x_nf:
            ax.TickLabels.NumberFormat = x_nf
        ax.TickLabelPosition = XL_TICK_LABEL_LOW
        ax.Format.Line.ForeColor.RGB = rgb(C["line"])
    except Exception:
        pass
    try:
        ay = ch.Axes(XL_VALUE)
        if y_nf:
            ay.TickLabels.NumberFormat = y_nf
        ay.HasMajorGridlines = gridlines
        if gridlines:
            ay.MajorGridlines.Format.Line.ForeColor.RGB = rgb(C["line2"])
        ay.Format.Line.Visible = False
    except Exception:
        pass
    if y2_nf:
        try:
            ch.Axes(XL_VALUE, XL_SECONDARY).TickLabels.NumberFormat = y2_nf
        except Exception:
            pass
    if legend is None:
        ch.HasLegend = False
    else:
        ch.HasLegend = True
        ch.Legend.Position = legend
        ch.Legend.Font.Size = 8


# ---------------------------------------------------------------- 버튼 (도형 + 매크로)
def add_button(ws, span: str, caption: str, macro: str, *, name: str | None = None, fill_color: str = C["accent"],
               font_color: str = "#FFFFFF", size: float = 10, bold: bool = True, placement: int = XL_FREE_FLOATING):
    """둥근 사각형 도형 버튼을 만들고 매크로를 연결한다(OnAction). 같은 이름이 있으면 지우고 다시 만든다.

    매크로가 아직 없어도 연결할 수 있다(빌더는 페이지를 먼저 만들고 VBA는 나중에 넣는다).
    .xlsm으로 저장하고 다시 열어도 연결이 유지된다.

    Args:
        ws: 대상 시트.
        span: 버튼이 차지할 셀 범위(예: "M2:N3") — 범위 크기에 맞춘다.
        caption: 버튼 글자(한글 가능).
        macro: 매크로 이름(예: "RefreshQuick").
        name: 도형 이름. 기본 "btn_<매크로>".
        fill_color, font_color, size, bold: 모양.
        placement: 기본 XL_FREE_FLOATING — 열 묶음 접기·열 숨김에도 크기·위치가 변하지 않음.

    Returns:
        Shape 객체.

    Example:
        add_button(ws, "M2:N3", "시세", "RefreshQuick")
    """
    r = ws.Range(span)
    name = name or f"btn_{macro}"
    for s in list(ws.Shapes):
        if s.Name == name:
            s.Delete()
    inset = 1.5
    shp = ws.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, r.Left + inset, r.Top + inset,
                             max(r.Width - 2 * inset, 12), max(r.Height - 2 * inset, 10))
    shp.Name = name
    shp.Fill.Visible = MSO_TRUE
    shp.Fill.Solid()
    shp.Fill.ForeColor.RGB = rgb(fill_color)
    shp.Line.Visible = MSO_FALSE
    tf = shp.TextFrame2
    tf.MarginLeft = tf.MarginRight = 2
    tf.MarginTop = tf.MarginBottom = 0
    tf.VerticalAnchor = MSO_ANCHOR_MIDDLE
    tf.WordWrap = MSO_FALSE
    tr = tf.TextRange
    tr.Text = caption
    tr.ParagraphFormat.Alignment = MSO_ALIGN_CENTER
    f = tr.Font
    f.Name = FONT
    f.NameFarEast = FONT
    f.Size = size
    f.Bold = MSO_TRUE if bold else MSO_FALSE
    f.Fill.ForeColor.RGB = rgb(font_color)
    shp.OnAction = macro
    shp.Placement = placement
    shp.AlternativeText = caption
    return shp


# ---------------------------------------------------------------- VBA 모듈 삽입 (코드 텍스트)
_VBA_ATTRIBUTE = re.compile(r"^\s*Attribute\s+([\w.]+)\s*=\s*(.*?)\s*$", re.IGNORECASE)


def _vba_text(text: str) -> tuple[str | None, list[tuple[int, str]]]:
    """내보낸 .bas 텍스트에서 Attribute 줄(머리줄·프로시저 속성)을 떼고 (VB_Name, [(원래 줄 번호, 줄)]) 반환."""
    vb_name = None
    kept = []
    for no, ln in enumerate(text.replace("\r\n", "\n").replace("\r", "\n").split("\n"), start=1):
        m = _VBA_ATTRIBUTE.match(ln)
        if m:
            if m.group(1).lower() == "vb_name":
                vb_name = m.group(2).strip('"')
            continue
        kept.append((no, ln))
    while kept and not kept[-1][1].strip():
        kept.pop()
    return vb_name, kept


def _check_vba_codepage(kept: list[tuple[int, str]], name: str):
    """VBA 코드는 시스템 ANSI 코드 페이지(이 PC는 cp949)로 저장된다. 표현할 수 없는 문자는 '?'로
    조용히 바뀌므로(시험: 'a—b✓c⚠d한' → 'a?b?c?d한') 넣기 전에 막는다. 줄 번호는 원본 파일 기준."""
    enc = f"cp{ctypes.windll.kernel32.GetACP()}"
    bad = []
    for no, ln in kept:
        for ch in ln:
            if ord(ch) < 128:
                continue
            try:
                ch.encode(enc)
            except UnicodeEncodeError:
                bad.append(f"{no}행 U+{ord(ch):04X} -> ChrW(&H{ord(ch):X})")
            except LookupError:
                return
    if bad:
        raise ValueError(f"VBA 모듈 {name}: 코드 페이지 {enc}로 표현할 수 없는 문자가 있어 '?'로 깨집니다. "
                         f"문자열은 ChrW()로 바꾸세요: " + ", ".join(bad[:10]))


def add_vba_module(wb, bas_path: str | None = None, *, code: str | None = None, name: str | None = None):
    """표준 모듈에 VBA 코드 텍스트를 직접 넣는다(VBComponents.Import 금지 — 이 PC에서 CP949로 읽혀 한글이 깨짐).

    .bas(UTF-8, BOM 있어도 됨)를 읽어 'Attribute ...' 줄을 떼고 CodeModule.AddFromString으로 넣는다.
    같은 이름의 표준 모듈이 있으면 내용을 바꾼다(다시 빌드해도 중복 없음).
    AccessVBOM이 켜진 상태에서 시작한 Excel 인스턴스가 필요하다(access_vbom() 참고).

    Args:
        wb: 대상 통합문서.
        bas_path: .bas 파일 경로(UTF-8). code와 둘 중 하나만.
        code: 코드 문자열(바로 넣을 때).
        name: 모듈 이름. 기본은 'Attribute VB_Name' 값, 없으면 파일 이름.

    Returns:
        VBComponent 객체.

    Raises:
        ValueError: 인수 오류, 또는 코드 페이지로 표현할 수 없는 문자가 있음(ChrW로 바꿀 것).
        RuntimeError: VBA 프로젝트에 접근할 수 없음(AccessVBOM 꺼짐 등).

    Example:
        with access_vbom(log=say):
            xl = new_excel()
            ...
            add_vba_module(wb, os.path.join(HERE, "vba", "modButtons.bas"))
    """
    if (bas_path is None) == (code is None):
        raise ValueError("bas_path 또는 code 중 하나만 주세요")
    if bas_path is not None:
        with open(bas_path, encoding="utf-8-sig") as f:
            text = f.read()
    else:
        text = code
    vb_name, kept = _vba_text(text)
    name = name or vb_name or (os.path.splitext(os.path.basename(bas_path))[0] if bas_path else None)
    if not name:
        raise ValueError("모듈 이름을 알 수 없습니다(name 인수 또는 Attribute VB_Name)")
    _check_vba_codepage(kept, name)
    body = "\r\n".join(ln for _, ln in kept) + "\r\n"
    try:
        comps = wb.VBProject.VBComponents
    except Exception as e:
        raise RuntimeError("VBA 프로젝트에 접근할 수 없습니다(AccessVBOM 꺼짐). access_vbom() 안에서, "
                           "그 안에서 새로 띄운 Excel 인스턴스로 호출하세요.") from e
    comp = None
    for c in comps:
        if c.Name.lower() == name.lower():
            if c.Type != VBEXT_CT_STD_MODULE:
                raise ValueError(f"같은 이름의 표준 모듈이 아닌 구성 요소가 있습니다: {c.Name}")
            comp = c
            break
    if comp is None:
        comp = comps.Add(VBEXT_CT_STD_MODULE)
        comp.Name = name
    cm = comp.CodeModule
    if cm.CountOfLines:
        cm.DeleteLines(1, cm.CountOfLines)   # 'Option Explicit' 자동 삽입분 등
    cm.AddFromString(body)
    return comp


# ---------------------------------------------------------------- AccessVBOM 임시 켜기 (반드시 복원)
@contextlib.contextmanager
def file_lock(path: str, owner: str = "build", *, timeout: float = 900.0, stale_after: float = 1800.0,
              poll: float = 2.0, log=None):
    """잠금 파일(O_CREAT|O_EXCL)로 같은 자원(예: AccessVBOM)을 쓰는 작업끼리 상호 배제한다.

    다른 작업의 잠금이 있으면 poll초마다 다시 시도하며 timeout초까지 기다린다(TimeoutError).
    잠금 파일이 stale_after초보다 오래됐으면 지우지 않고 RuntimeError로 멈춘다(남의 잠금은 절대 지우지 않음).
    끝나면 자기가 쓴 잠금 파일만 지운다.

    Example:
        with file_lock(r"...\\locks\\vbom.lock", "build_dashboard", log=say):
            ...
    """
    token = f"{owner} pid={os.getpid()} {dt.datetime.now():%Y-%m-%d %H:%M:%S}"
    t0 = time.time()
    told = False
    while True:
        try:
            fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY | getattr(os, "O_BINARY", 0))
            break
        except FileExistsError:
            try:
                age = time.time() - os.path.getmtime(path)
            except FileNotFoundError:
                continue  # 방금 풀림
            if age > stale_after:
                raise RuntimeError(f"잠금 파일이 {age / 60:.0f}분째 남아 있습니다(다른 작업의 비정상 종료 의심). "
                                   f"지우지 않고 멈춥니다: {path}")
            if time.time() - t0 > timeout:
                raise TimeoutError(f"잠금 대기 {timeout:.0f}초 초과: {path}")
            if log and not told:
                log(f"잠금 대기 중: {path}")
                told = True
            time.sleep(poll)
    try:
        os.write(fd, (token + "\n").encode("utf-8"))
    finally:
        os.close(fd)
    try:
        yield path
    finally:
        try:
            with open(path, encoding="utf-8") as f:
                mine = f.read().strip() == token
        except FileNotFoundError:
            mine = False
        if mine:
            os.remove(path)
        elif log:
            log(f"잠금 파일이 바뀌어 있어 지우지 않음: {path}")


def _vbom_read(key_path: str):
    """(키 있음, 값 있음, 값, 형식)."""
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
    except FileNotFoundError:
        return False, False, None, None
    with k:
        try:
            v, t = winreg.QueryValueEx(k, VBOM_VALUE)
            return True, True, v, t
        except FileNotFoundError:
            return True, False, None, None


def _missing_keys(key_path: str) -> list[str]:
    """key_path와 그 상위 키 중 지금 없는 것(깊은 것부터) — 켜면서 새로 만들게 될 키."""
    parts = key_path.split("\\")
    missing = []
    for i in range(len(parts), 0, -1):
        p = "\\".join(parts[:i])
        try:
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CURRENT_USER, p, 0, winreg.KEY_READ))
            break
        except FileNotFoundError:
            missing.append(p)
    return missing


def _vbom_desc(st) -> str:
    if not st.key_existed:
        return "키 없음"
    if not st.had_value:
        return "값 없음"
    return f"값 {st.value}"


def _vbom_restore(st) -> str:
    """원래 상태로 되돌리고 다시 읽어 확인. 실패하면 RuntimeError(수동 복구 안내)."""
    desc = _vbom_desc(st)
    hk = winreg.HKEY_CURRENT_USER
    try:
        if st.had_value:
            with winreg.CreateKeyEx(hk, st.key_path, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, VBOM_VALUE, 0, st.value_type, st.value)
        else:
            try:
                with winreg.OpenKey(hk, st.key_path, 0, winreg.KEY_SET_VALUE) as k:
                    winreg.DeleteValue(k, VBOM_VALUE)
            except FileNotFoundError:
                pass
            for p in st.created_keys:            # 켜면서 만든 키(깊은 것부터): 비어 있을 때만 지움
                try:
                    with winreg.OpenKey(hk, p, 0, winreg.KEY_READ) as k:
                        n_sub, n_val, _ = winreg.QueryInfoKey(k)
                except FileNotFoundError:
                    continue
                if n_sub or n_val:
                    break
                winreg.DeleteKey(hk, p)
    except OSError as e:
        raise RuntimeError(f"AccessVBOM 복원 실패: HKCU\\{st.key_path}\\{VBOM_VALUE} 를 직접 '{desc}' 상태로 "
                           f"되돌리세요 ({e})") from e
    key_now, had_now, v_now, t_now = _vbom_read(st.key_path)
    same = had_now == st.had_value and (not st.had_value or (v_now == st.value and t_now == st.value_type))
    if not same:
        raise RuntimeError(f"AccessVBOM 복원 확인 실패: 원래 '{desc}', 지금 값 있음={had_now} 값={v_now}")
    tail = " (Security 키는 다른 항목이 생겨 남겨 둠)" if (not st.key_existed and key_now) else ""
    return f"AccessVBOM 원래 상태로 복원 확인: {desc}{tail}"


def _log_safe(log, msg: str):
    if log is None:
        return
    try:
        log(msg)
    except Exception:
        pass


@contextlib.contextmanager
def access_vbom(log=print, *, lock_path: str | None = None, owner: str = "build", key_path: str = VBOM_KEY):
    """AccessVBOM(VBA 프로젝트 개체 모델 접근 신뢰)을 블록 동안만 1로 켜고, 성공·실패와 관계없이
    원래 상태(값 / 값 없음 / 키 없음)로 되돌린 뒤 다시 읽어 확인한다.

    Excel은 이 값을 인스턴스가 시작될 때 한 번 읽는다(시험: 켜기 전에 띄운 인스턴스는 켠 뒤에도 접근 불가,
    켠 동안 띄운 인스턴스는 복원 뒤에도 접근 가능, Excel 종료 때 값을 되쓰지 않음). 따라서 VBA를 넣을
    Excel 인스턴스는 반드시 이 블록 안에서 새로 띄운다. 빌더는 Excel 시작~종료 전체를 이 블록으로 감싼다.

    Args:
        log: 로그 함수(예: builder.say). 켤 때와 복원 결과(한국어 한 줄)를 남긴다.
        lock_path: 주면 file_lock으로 같은 값을 만지는 다른 작업과 상호 배제(개발 중 병렬 작업용).
        owner: 잠금 파일에 적을 작업 이름.
        key_path: HKCU 아래 키 경로(시험용으로만 바꿈).

    Yields:
        상태 객체(SimpleNamespace): key_existed, had_value, value, value_type, message.
        message는 블록이 끝난 뒤 복원 결과 문구가 채워진다(빌드 로그용).

    Raises:
        RuntimeError: 복원 또는 복원 확인 실패(수동 복구 안내 포함). 블록 안의 예외는 복원 뒤 그대로 전달.

    Example:
        with access_vbom(log=self.say) as vbom:
            xl = new_excel()
            ...                      # 통합문서 만들기, add_vba_module, save_xlsm
            quit_excel(xl, pid)
        # vbom.message == "AccessVBOM 원래 상태로 복원 확인: 값 없음"
    """
    lock = file_lock(lock_path, owner, log=log) if lock_path else contextlib.nullcontext()
    with lock:
        key_existed, had, value, vtype = _vbom_read(key_path)
        st = types.SimpleNamespace(key_path=key_path, key_existed=key_existed, had_value=had, value=value,
                                   value_type=vtype, created_keys=_missing_keys(key_path), message="")
        try:
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, VBOM_VALUE, 0, winreg.REG_DWORD, 1)
            _log_safe(log, f"AccessVBOM 임시 켜기(1), 원래 상태: {_vbom_desc(st)}")
            yield st
        finally:
            try:
                st.message = _vbom_restore(st)
            except Exception as e:
                st.message = str(e)
                _log_safe(log, st.message)
                raise
            _log_safe(log, st.message)


# ---------------------------------------------------------------- .xlsm 저장
def save_xlsm(wb, path: str) -> str:
    """매크로 포함 통합문서(.xlsm, FileFormat 52)로 저장하고 절대 경로를 반환한다.

    VBA가 든 통합문서를 .xlsx(51)로 저장하면 매크로가 조용히 빠지므로(DisplayAlerts=False일 때) 반드시 이걸 쓴다.
    같은 경로의 파일은 덮어쓴다(DisplayAlerts=False). 그 파일이 다른 Excel에서 열려 있으면 실패한다.

    Example:
        save_xlsm(wb, os.path.join(HERE, "KIS_PM_Dashboard.xlsm"))
    """
    path = os.path.abspath(path)
    if not path.lower().endswith(".xlsm"):
        raise ValueError(f".xlsm 경로가 아닙니다: {path}")
    wb.SaveAs(path, XL_OPENXML_WORKBOOK_MACRO_ENABLED)
    return path
