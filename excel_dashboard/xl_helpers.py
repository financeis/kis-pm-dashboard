"""Excel COM 헬퍼 (build_dashboard.py 전용).

- 항상 새 Excel 인스턴스(DispatchEx)를 띄워 작업하고, 사용자가 열어 둔 Excel에는 손대지 않습니다.
- 수식은 Formula2(동적 배열)로 입력합니다. 함수명·구분자는 영문/쉼표 기준입니다.
"""
from __future__ import annotations

import datetime as dt
import time

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
