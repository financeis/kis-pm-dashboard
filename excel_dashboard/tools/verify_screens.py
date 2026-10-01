# -*- coding: utf-8 -*-
"""색·화면 검사: Q.Pack 기준점·색 규칙·접기 묶음(파일 검사)과 페이지 PNG 렌더링·'####' 칸 검사(Excel).

파일 검사 (--workbook, Excel 없이 xlsx_tables로 저장값·XML을 읽음)
- 열별 기준점: [대회종목] 숨김 9~11행의 기준점 칸(낮음·중앙·높음)이 `tblCompany`의 **유니버스 = 대회** 행만의 최소·중앙값·최대와
  같은지 파이썬으로 다시 계산해 대조하고, 수식에 유니버스 조건이 들어 있는지 확인한다. 유니버스 밖 행이 있으면 그 행을 넣었을 때의
  값과 달라지는 열 수도 보고한다(기준에서 빠졌다는 증거).
- 줄무늬: D-19…D0 20열 전체(대회 행)를 하나의 척도로 — 기준점 = 20칸 값 전체의 5·50·95 백분위(PERCENTILE.INC)와 같은지,
  색 규칙 하나가 20열 블록 전체에 걸려 있는지.
- 색 규칙: 열마다 3색 척도(#63BE7B → #FFEB84 → #F8696B)가 그 열 위 기준점 칸을 참조하는지, 수준 값 열에는 색이 없는지.
- 접기 묶음: 시트 XML의 열 숨김·개요 수준으로 기본 펼침 열(아래 SPEC_VISIBLE)이 보이고 접는 묶음이 접혀 있는지.

Excel 검사 (--render-copy, 반드시 **복사본** 경로 — 열기 전 토큰 관문)
- 시트별 PNG 렌더링(범위 → PDF → PNG, 클립보드 미사용): 대회종목·업종·종목분석·뉴스·이벤트·시장·시세판·대시보드.
- '####' 칸: 확대/축소 85%인 창에서 숫자 칸의 표시 글자(Range.Text)가 '#'로만 된 칸을 시트별로 찾는다(열 너비 부족).
  --screenshot을 주면 자기 Excel 창을 보이게 해 85% 화면을 PrintWindow로 캡처한다(그 창 내용만 — 다른 창은 찍히지 않음).
  참고(2026-10-01 실측): PDF 렌더링(한 쪽 맞춤 인쇄 배율)에서는 [시장] 지수 카드·[대시보드] 시장 스트립 숫자가 '####'로 보이지만
  실제 화면(85%·100%)의 Range.Text와 창 캡처에서는 모두 제대로 보인다 — 인쇄 배율의 글꼴 폭 차이.

    python excel_dashboard/tools/verify_screens.py --workbook <저장한.xlsm> [--json 결과.json]
    python excel_dashboard/tools/verify_screens.py --render-copy <복사본.xlsm> --out <PNG 폴더> --token-source <원천.xlsm> [--screenshot]
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import statistics
import sys
import time
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
sys.path.insert(0, DASH)
sys.path.insert(0, HERE)
from xlsx_tables import Workbook, _M, _split_ref  # noqa: E402

SHEET = "대회종목"
TABLE = "tblCompany"
ANCHOR_ROWS = (9, 10, 11)
COLORS = ("63BE7B", "FFEB84", "F8696B")
PERIODS = ["1D", "1W", "1M", "3M", "6M", "YTD", "1Y"]
STRIP = [f"D-{k}" for k in range(19, 0, -1)] + ["D0"]
QPACK_COLUMNS = (PERIODS + ["고52주대비", "외국인1D%", "외국인1W%", "외국인1M%", "기관1D%", "기관1W%", "기관1M%",
                            "매출YoY", "영업이익YoY", "순이익YoY", "PER변화3M", "PER변화6M", "PER변화YTD", "PER변화1Y",
                            "PBR변화3M", "PBR변화6M", "PBR변화YTD", "PBR변화1Y", "괴리율", "목표가변화1M", "목표가변화3M",
                            "FwdPER변화1W", "FwdPER변화1M", "신용잔고율", "신용잔고율1M변화", "공매도비중5일", "대차잔고1M변화율"])
LEVEL_COLUMNS = ["현재가", "시가총액억", "거래대금60일억", "후행PER", "PBR", "FwdPER", "목표가평균"]   # 색이 없어야 하는 수준 값
# 대회종목 표의 기본 펼침 열(접지 않고 보이는 열)
SPEC_VISIBLE = (["종목코드", "종목명", "시장", "구분", "유니버스", "NICS 업종", "대테마", "시가총액억", "거래대금60일억", "현재가"]
                + PERIODS + STRIP + ["외국인1D%", "외국인1W%", "외국인1M%", "기관1D%", "기관1W%", "기관1M%",
                                     "후행PER", "FwdPER", "FwdPER변화1W", "FwdPER변화1M", "목표가평균", "괴리율", "증권사수"])
# 대회종목 표에서 기본으로 접어 둘 열
SPEC_FOLDED = (["NICS 대분류", "NICS 세부", "세부테마", "KSIC 세분류", "주요상품", "기업집단", "분류출처",
                "거래량60일천주", "고52주대비", "외국인1D억", "외국인1W억", "외국인1M억", "기관1D억", "기관1W억", "기관1M억",
                "최근분기", "매출YoY", "영업이익YoY", "순이익YoY", "실적상태", "ROE",
                "PBR", "PER변화3M", "PER변화6M", "PER변화YTD", "PER변화1Y", "PBR변화3M", "PBR변화6M", "PBR변화YTD", "PBR변화1Y",
                "목표가변화1M", "목표가변화3M", "최근의견일", "최근증권사", "최근의견",
                "FwdEPS", "추정일", "FY1", "FY1_EPS", "FY2", "FY2_EPS",
                "신용잔고율", "신용잔고율1M변화", "공매도비중5일", "대차잔고1M변화율", "신용기준일",
                "다음이벤트", "유의", "상태", "조회시각"])
RENDER = [("대회종목", "B2:BT60"), ("대회종목", "BT12:DA60"), ("업종", None), ("종목분석", None), ("뉴스·이벤트", None),
          ("시장", None), ("시세판", None), ("대시보드", None)]
HASH_SHEETS = ["대시보드", "시장", "업종", "대회종목", "종목분석", "뉴스·이벤트", "포트폴리오", "리스크", "성과", "시세판"]


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def pct_inc(xs: list[float], p: float) -> float:
    """Excel PERCENTILE.INC."""
    v = sorted(xs)
    h = (len(v) - 1) * p
    lo = int(h)
    return v[lo] if lo + 1 >= len(v) else v[lo] + (h - lo) * (v[lo + 1] - v[lo])


def isnum(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def file_checks(path: str) -> dict:
    with Workbook(path) as wb:
        info = wb.tables()[TABLE]
        cols, rows = wb.read_table(TABLE)
        first = info["ref"].split(":")[0]
        c0, hdr_row = _split_ref(first)
        part = wb.sheets[SHEET]
        xml = wb.zip.read(part)
        # 기준점 칸 값·수식
        cells, forms = {}, {}
        for _ev, el in ET.iterparse(io.BytesIO(xml), events=("end",)):
            if el.tag == _M + "c":
                ref = el.get("r")
                col, row = _split_ref(ref)
                if row in ANCHOR_ROWS:
                    cells[ref] = wb._cell_value(el)
                    f = el.find(_M + "f")
                    forms[ref] = f.text if f is not None else None
    letter = {name: col_letter(c0 + i) for i, name in enumerate(cols)}
    contest = [dict(zip(cols, r)) for r in rows if dict(zip(cols, r)).get("유니버스") == "대회"]
    outside = [dict(zip(cols, r)) for r in rows if dict(zip(cols, r)).get("유니버스") == "유니버스 밖"]
    res = {"rows": len(rows), "contest_rows": len(contest), "outside_rows": len(outside), "anchors": [], "anchor_bad": [],
           "outside_effect_cols": []}

    def close(a, b):
        return (a in (None, "") and b in (None, "")) or (isnum(a) and isnum(b) and abs(a - b) <= 1e-9 * max(1, abs(a), abs(b)))
    for name in QPACK_COLUMNS:
        L = letter[name]
        xs = [r[name] for r in contest if isnum(r.get(name))]
        exp = (min(xs), statistics.median(xs), max(xs)) if xs else ("", "", "")
        got = tuple(cells.get(f"{L}{k}") for k in ANCHOR_ROWS)
        # 파일에는 구조적 참조가 tblCompany[[#All],[열]] 형식으로 저장될 수 있다(머리글 칸은 ISNUMBER·조건에서 빠져 결과 같음)
        cond_ok = all(forms.get(f"{L}{k}") and re.search(r'\[유니버스\]\]?="대회"', forms[f"{L}{k}"])
                      and f"[{name}]" in forms[f"{L}{k}"] for k in ANCHOR_ROWS)
        ok = all(close(a, b) for a, b in zip(got, exp)) and cond_ok
        res["anchors"].append({"col": name, "letter": L, "expected": exp, "got": got, "cond_in_formula": cond_ok, "ok": ok})
        if not ok:
            res["anchor_bad"].append(name)
        if outside:
            ys = xs + [r[name] for r in outside if isnum(r.get(name))]
            if ys and xs and (min(ys), statistics.median(ys), max(ys)) != exp:
                res["outside_effect_cols"].append(name)
    # 줄무늬 블록
    Ls = letter[STRIP[0]]
    xs = [r[c] for r in contest for c in STRIP if isnum(r.get(c))]
    exp = tuple(pct_inc(xs, p) for p in (0.05, 0.5, 0.95)) if xs else ("", "", "")
    got = tuple(cells.get(f"{Ls}{k}") for k in ANCHOR_ROWS)
    f9 = forms.get(f"{Ls}9") or ""
    res["strip"] = {"letter": Ls, "values": len(xs), "expected": exp, "got": got,
                    "formula_block": bool(re.search(r"\[D-19\]:\[D0\]", f9)) and bool(re.search(r'\[유니버스\]\]?="대회"', f9))
                    and "PERCENTILE.INC" in f9,
                    "ok": all(close(a, b) for a, b in zip(got, exp))}
    # 색 규칙
    root = ET.fromstring(xml)
    scales = []
    for cf in root.iter(_M + "conditionalFormatting"):
        sq = cf.get("sqref", "")
        for rule in cf.findall(_M + "cfRule"):
            csn = rule.find(_M + "colorScale")
            if csn is None:
                continue
            vals = [v.get("val") for v in csn.findall(_M + "cfvo")]
            cols_ = [(c.get("rgb") or "")[-6:].upper() for c in csn.findall(_M + "color")]
            scales.append({"sqref": sq, "vals": vals, "colors": cols_})
    body_top = hdr_row + 1
    rule_bad, has_rule = [], {}
    for name in QPACK_COLUMNS:
        L = letter[name]
        want = [f"${L}${k}" for k in ANCHOR_ROWS]
        hit = [s for s in scales if any(re.fullmatch(rf"{L}{body_top}:{L}\d+", part) for part in s["sqref"].split())
               and [v.replace("=", "") for v in s["vals"]] == want]
        has_rule[name] = bool(hit)
        if not hit or tuple(hit[0]["colors"]) != COLORS:
            rule_bad.append(name)
    strip_want = [f"${Ls}${k}" for k in ANCHOR_ROWS]
    Le = letter[STRIP[-1]]
    strip_rules = [s for s in scales if any(re.fullmatch(rf"{Ls}{body_top}:{Le}\d+", p) for p in s["sqref"].split())
                   and [v.replace("=", "") for v in s["vals"]] == strip_want and tuple(s["colors"]) == COLORS]
    level_colored = [n for n in LEVEL_COLUMNS
                     if any(re.fullmatch(rf"{letter[n]}{body_top}:{letter[n]}\d+", p) for s in scales for p in s["sqref"].split())]
    res["color_rules"] = {"color_scales": len(scales), "qpack_columns": len(QPACK_COLUMNS), "missing_or_wrong": rule_bad,
                          "strip_single_rule": len(strip_rules) == 1, "level_columns_colored": level_colored}
    # 접기 묶음
    hidden, levels = set(), {}
    for c in root.iter(_M + "col"):
        a, b = int(c.get("min")), int(c.get("max"))
        for i in range(a, b + 1):
            if c.get("hidden") in ("1", "true"):
                hidden.add(i)
            if c.get("outlineLevel"):
                levels[i] = int(c.get("outlineLevel"))
    idx = {name: c0 + i for i, name in enumerate(cols)}
    res["fold"] = {
        "spec_visible_hidden": [n for n in SPEC_VISIBLE if idx[n] in hidden],
        "spec_folded_visible": [n for n in SPEC_FOLDED if idx[n] not in hidden],
        "spec_folded_hidden": sum(1 for n in SPEC_FOLDED if idx[n] in hidden),
        "grouped_columns": sum(1 for n in cols if levels.get(idx[n])),
    }
    return res


# ------------------------------------------------------------------------------------------------- Excel(복사본)
def grab_window(hwnd: int, png: str) -> None:
    """창 하나의 내용만 PNG로(PrintWindow, PW_RENDERFULLCONTENT) — 다른 프로그램 창은 찍히지 않는다."""
    import win32gui
    import win32ui
    from ctypes import windll
    from PIL import Image
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    w, h = right - left, bottom - top
    hdc = win32gui.GetWindowDC(hwnd)
    mfc = win32ui.CreateDCFromHandle(hdc)
    mem = mfc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc, w, h)
    mem.SelectObject(bmp)
    windll.user32.PrintWindow(hwnd, mem.GetSafeHdc(), 2)
    info = bmp.GetInfo()
    Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), bmp.GetBitmapBits(True), "raw", "BGRX", 0, 1).save(png)
    win32gui.DeleteObject(bmp.GetHandle())
    mem.DeleteDC()
    mfc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hdc)


def render_copy(path: str, out_dir: str, token_source: str | None, screenshot: bool,
                hash_sheets: list[str] | None = None) -> dict:
    import fitz
    from verify_common import close_workbook, open_workbook
    os.makedirs(out_dir, exist_ok=True)
    real = os.path.normcase(os.path.abspath(token_source or ""))
    if os.path.normcase(os.path.abspath(path)) == real:
        raise SystemExit("렌더링·창 조작은 복사본에서만 합니다(토큰 원천 = 실제 통합문서 경로를 줬음).")
    xl, wb, pid, _info = open_workbook(path, visible=screenshot, token_source=token_source)
    res = {"png": [], "hash_cells": {}}
    try:
        xl.ScreenUpdating = True
        for sheet, rng_addr in RENDER:
            ws = wb.Worksheets(sheet)
            rng = ws.Range(rng_addr) if rng_addr else ws.UsedRange
            png = os.path.join(out_dir, f"{sheet}{'_' + rng_addr.replace(':', '-') if rng_addr else ''}.png")
            pdf = png[:-4] + ".__r.pdf"
            ps = ws.PageSetup
            try:
                xl.PrintCommunication = False
            except Exception:  # noqa: BLE001
                pass
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            ps.Orientation = 2 if rng.Width >= rng.Height else 1
            m = xl.InchesToPoints(0.2)
            ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = m
            try:
                xl.PrintCommunication = True
            except Exception:  # noqa: BLE001
                pass
            if os.path.exists(pdf):
                os.remove(pdf)
            rng.ExportAsFixedFormat(0, pdf, 0, False, True)
            doc = fitz.open(pdf)
            try:
                doc[0].get_pixmap(dpi=170).save(png)
            finally:
                doc.close()
                os.remove(pdf)
            res["png"].append(png)
        # '####' 칸 (85% 창)
        for sheet in (hash_sheets if hash_sheets is not None else HASH_SHEETS):
            ws = wb.Worksheets(sheet)
            ws.Activate()
            xl.ActiveWindow.Zoom = 85
            ur = ws.UsedRange
            vals = ur.Value2
            r0, c0 = ur.Row, ur.Column
            bad = []
            for i, row in enumerate(vals or ()):
                for j, v in enumerate(row if isinstance(row, tuple) else (row,)):
                    if isinstance(v, float) or (isinstance(v, int) and not isinstance(v, bool)):
                        cell = ws.Cells(r0 + i, c0 + j)
                        if cell.EntireColumn.Hidden or cell.EntireRow.Hidden:
                            continue
                        t = str(cell.Text)
                        if t and set(t) <= {"#"}:
                            bad.append(cell.Address.replace("$", ""))
            if bad:
                res["hash_cells"][sheet] = bad
        if screenshot:
            # 자기 Excel 창만 캡처(PrintWindow) — 화면 캡처(ImageGrab)는 앞에 있는 다른 창이 찍힐 수 있어 쓰지 않는다
            xl.WindowState = -4137          # 최대화
            for sheet in ("시장", "대시보드"):
                ws = wb.Worksheets(sheet)
                ws.Activate()
                xl.ActiveWindow.Zoom = 85
                xl.ActiveWindow.ScrollRow = 1
                xl.ActiveWindow.ScrollColumn = 1
                time.sleep(1.5)
                p = os.path.join(out_dir, f"화면85_{sheet}.png")
                grab_window(xl.Hwnd, p)
                res["png"].append(p)
    finally:
        close_workbook(xl, wb, pid, save=False)
    return res


def main() -> int:
    ap = argparse.ArgumentParser(description="V6 색·화면 검사")
    ap.add_argument("--workbook", help="파일 검사할 통합문서(저장된 값)")
    ap.add_argument("--render-copy", help="렌더링·#### 검사할 복사본")
    ap.add_argument("--out", default=None, help="PNG 폴더")
    ap.add_argument("--token-source", default=None)
    ap.add_argument("--screenshot", action="store_true", help="자기 Excel 창(85%)을 PrintWindow로 캡처 — 시장·대시보드")
    ap.add_argument("--hash-sheets", nargs="*", default=None,
                    help="'####' 검사할 시트(기본: 화면 시트 10개 — 칸마다 COM 호출이라 수 분 걸림)")
    ap.add_argument("--json", default=None)
    a = ap.parse_args()
    out = {}
    if a.workbook:
        r = file_checks(a.workbook)
        out["file"] = r
        print(f"[대회종목] 행 {r['rows']} (대회 {r['contest_rows']} · 유니버스 밖 {r['outside_rows']})")
        print(f"  열별 기준점 {len(r['anchors'])}열: 불일치 {r['anchor_bad'] or '없음'} (모두 수식에 유니버스=대회 조건: "
              f"{all(x['cond_in_formula'] for x in r['anchors'])})")
        if r["outside_rows"]:
            print(f"  유니버스 밖 행을 넣으면 기준점이 달라지는 열 {len(r['outside_effect_cols'])}개 → 기준에서 빠져 있음: {r['outside_effect_cols'][:10]}")
        s = r["strip"]
        print(f"  줄무늬: 값 {s['values']}개, 기준점 5·50·95 백분위 기대 {tuple(round(x, 6) if isnum(x) else x for x in s['expected'])} "
              f"/ 칸 {tuple(round(x, 6) if isnum(x) else x for x in s['got'])} → {'일치' if s['ok'] else '불일치'}, 블록 수식 {s['formula_block']}")
        c = r["color_rules"]
        print(f"  색 규칙: 3색 척도 {c['color_scales']}개, 34열 중 누락·오류 {c['missing_or_wrong'] or '없음'}, 줄무늬 단일 규칙 {c['strip_single_rule']},"
              f" 수준 값 열 색 {c['level_columns_colored'] or '없음'}")
        f = r["fold"]
        print(f"  접기: 기본 펼침 열 중 숨김 {f['spec_visible_hidden'] or '없음'} · 접을 열 {len(SPEC_FOLDED)}개 중 숨김 {f['spec_folded_hidden']}"
              f" · 보이는 요약 열 {f['spec_folded_visible']}")
    if a.render_copy:
        r = render_copy(a.render_copy, a.out or os.path.dirname(a.render_copy), a.token_source, a.screenshot, a.hash_sheets)
        out["render"] = r
        print(f"PNG {len(r['png'])}장: {r['png']}")
        print(f"'####' 칸(85%): {r['hash_cells'] or '없음'}")
    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump(out, fh, ensure_ascii=False, indent=1, default=str)
    return 0


if __name__ == "__main__":
    sys.exit(main())
