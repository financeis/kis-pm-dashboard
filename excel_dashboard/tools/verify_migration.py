# -*- coding: utf-8 -*-
"""이관 왕복 검사 (T24 — spec §8 V7, spec R22): 복사본에 사용자 데이터를 넣고 다시 빌드해 모두 보존되는지.

순서 (모두 **복사본·스크래치 폴더**에서 — 사용자 통합문서는 읽기만 함)
1. 원본(.xlsm)을 작업 폴더로 복사 → 자기 Excel로 열어(토큰 관문) 사용자 데이터를 넣는다:
   매매일지 1행(전략 'V7시험'), 관심종목 1행(대회 종목이 아닌 코드), 설정값 변경(max_weight), 수정표 2행(테마 변경·대회편입 추가),
   스냅샷(tblEstSnap) 한 칸에 표시값(FwdPER = 12.3456), 사용자 추가 시트 '내 메모'.
   그리고 **.xlsx로 다른 이름 저장**(VBA 없는 형식) → 빌드의 기본 이관 원본 규칙(출력 위치에 .xlsm이 없으면 같은 이름 .xlsx)을 시험.
2. `build_dashboard.py --out <작업 폴더>/KIS_PM_Dashboard.xlsm`(이관 원본 자동 선택)을 실행하고 종료 코드·로그·시간을 기록.
3. 새 .xlsm을 파일 파싱으로 읽어 넣은 값이 모두 있는지, 로그에 사용자 추가 시트 경고가 있는지, .xlsx 원본이 backup 폴더로
   옮겨졌는지(원래 자리에 없음), AccessVBOM이 원래 상태(값 없음)인지 확인.

    python excel_dashboard/tools/verify_migration.py --source <실제 통합문서.xlsm> --work <스크래치 폴더> \\
        --token-source <원천.xlsm> [--vbom-lock <잠금 파일>] [--json 결과.json]
종료 코드: 0 통과 / 1 보존 실패·빌드 실패 / 2 토큰 갱신 필요
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import json
import os
import shutil
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
sys.path.insert(0, DASH)
sys.path.insert(0, HERE)
from verify_common import (  # noqa: E402
    TokenGateError, find_table, log, open_workbook, quit_instance, com_retry, wait_idle,
)
from xlsx_tables import Workbook  # noqa: E402

SENTINEL_PER = 12.3456
USER_SHEET = "내 메모"


def vbom_state() -> str:
    import winreg
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Excel\Security")
    except OSError:
        return "키 없음"
    try:
        v, _t = winreg.QueryValueEx(k, "AccessVBOM")
        return f"값 {v}"
    except OSError:
        return "값 없음"
    finally:
        winreg.CloseKey(k)


def add_row(lo, values: dict) -> None:
    """정적 입력표 끝에 행을 더하고 열 이름으로 값을 넣는다(날짜는 Value2 일련번호)."""
    hdr = [str(h) for h in lo.HeaderRowRange.Value2[0]]
    # 마지막 행이 완전히 비었으면 그 행을 쓰고(빈 표의 빈 행), 아니면 새 행
    use = None
    if lo.ListRows.Count:
        last = lo.ListRows(lo.ListRows.Count).Range
        if all(v in (None, "") for v in last.Value2[0]):
            use = last
    if use is None:
        use = lo.ListRows.Add().Range
    for k, v in values.items():
        cell = use.Cells(1, hdr.index(k) + 1)
        if isinstance(v, dt.date):
            cell.Value2 = (dt.datetime(v.year, v.month, v.day) - dt.datetime(1899, 12, 30)).days
        elif isinstance(v, str) and k == "종목코드":
            cell.NumberFormat = "@"
            cell.Value2 = v
        else:
            cell.Value2 = v


def prepare(src: str, xlsx_out: str, token_source: str) -> dict:
    """복사본에 사용자 데이터를 넣고 .xlsx로 저장. 넣은 값(기대값)을 돌려준다."""
    xl, wb, pid, _ = open_workbook(src, token_source=token_source)
    exp = {}
    try:
        uni = find_table(wb, "tblUniverse")
        hdr = [str(h) for h in uni.HeaderRowRange.Value2[0]]
        vals = uni.DataBodyRange.Value2
        ci, ei, ki = hdr.index("종목코드"), hdr.index("대회편입"), hdr.index("종류") if "종류" in hdr else None
        outside = [r[ci] for r in vals if r[ei] == "N" and (ki is None or r[ki] == "ST")]
        out_code = sorted(str(c) for c in outside if c)[len(outside) // 2]
        exp["outside_code"] = out_code
        add_row(find_table(wb, "tblTrades"), {"일자": dt.date(2026, 9, 30), "종목코드": "000270", "구분": "매수", "수량": 5,
                                              "단가": 100000, "전략": "V7시험", "매매근거": "이관 왕복 시험 행"})
        add_row(find_table(wb, "tblWatch"), {"종목코드": out_code, "그룹": "V7", "투자포인트": "이관 왕복 시험(유니버스 밖)"})
        st = find_table(wb, "tblSettings")
        h = [str(x) for x in st.HeaderRowRange.Value2[0]]
        for r in range(1, st.ListRows.Count + 1):
            row = st.ListRows(r).Range
            if row.Cells(1, h.index("키") + 1).Value2 == "max_weight":
                row.Cells(1, h.index("값") + 1).Value2 = 0.25
        exp["max_weight"] = 0.25
        ov = find_table(wb, "tblOverride")
        add_row(ov, {"종목코드": "005380", "대테마": "V7테마", "세부테마": "V7세부", "메모": "테마 변경 시험"})
        add_row(ov, {"종목코드": out_code, "대회편입": "추가", "메모": "대회편입 추가 시험"})
        snap = find_table(wb, "tblEstSnap")
        if snap is not None and snap.ListRows.Count > 0 and snap.DataBodyRange is not None:
            sh = [str(x) for x in snap.HeaderRowRange.Value2[0]]
            row = snap.ListRows(1).Range
            exp["snap_key"] = (row.Cells(1, sh.index("일자") + 1).Value2, row.Cells(1, sh.index("종목코드") + 1).Value2)
            row.Cells(1, sh.index("FwdPER") + 1).Value2 = SENTINEL_PER
            exp["snap_rows"] = snap.ListRows.Count
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = USER_SHEET
        ws.Range("A1").Value2 = "V7 사용자 추가 시트(이관하지 않고 경고해야 함)"
        exp["trades"] = find_table(wb, "tblTrades").ListRows.Count
        exp["watch"] = find_table(wb, "tblWatch").ListRows.Count
        wait_idle(wb, timeout=300)
        com_retry(lambda: wb.SaveAs(xlsx_out, 51))          # xlOpenXMLWorkbook (.xlsx, VBA 없음)
        log(f"복사본에 사용자 데이터 넣고 .xlsx로 저장: {xlsx_out}")
        com_retry(lambda: wb.Close(False))
    finally:
        quit_instance(xl, pid)
    return exp


def main() -> int:
    ap = argparse.ArgumentParser(description="V7 이관 왕복 검사(복사본) — 토큰 값은 출력하지 않음")
    ap.add_argument("--source", required=True, help="원본 통합문서(.xlsm) — 복사만 함")
    ap.add_argument("--work", required=True, help="작업 폴더(스크래치)")
    ap.add_argument("--token-source", required=True)
    ap.add_argument("--vbom-lock", default=None)
    ap.add_argument("--json", default=None)
    a = ap.parse_args()
    work = os.path.abspath(a.work)
    if os.path.exists(work) and os.listdir(work):
        print(f"작업 폴더가 비어 있지 않습니다: {work}")
        return 1
    os.makedirs(work, exist_ok=True)
    src_copy = os.path.join(work, "_src", "KIS_PM_Dashboard.xlsm")
    os.makedirs(os.path.dirname(src_copy))
    shutil.copy2(a.source, src_copy)
    xlsx = os.path.join(work, "KIS_PM_Dashboard.xlsx")
    out = os.path.join(work, "KIS_PM_Dashboard.xlsm")
    res = {"vbom_before": vbom_state()}
    try:
        exp = prepare(src_copy, xlsx, a.token_source)
    except TokenGateError as e:
        print(f"멈춤: {e}")
        return 2
    res["expected"] = exp
    os.remove(src_copy)                       # 토큰이 든 중간 복사본 정리(.xlsx는 빌드의 이관 원본)
    os.rmdir(os.path.dirname(src_copy))
    cmd = [sys.executable, os.path.join(DASH, "build_dashboard.py"), "--out", out]
    if a.vbom_lock:
        cmd += ["--vbom-lock", a.vbom_lock]
    log("빌드 시작: build_dashboard.py --out <작업 폴더>/KIS_PM_Dashboard.xlsm (이관 원본 자동 = 같은 이름 .xlsx)")
    t0 = time.time()
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", env=env)
    res["build_seconds"] = round(time.time() - t0, 1)
    res["build_exit"] = p.returncode
    with open(os.path.join(work, "build_stdout.log"), "w", encoding="utf-8") as fh:
        fh.write(p.stdout + "\n--- stderr ---\n" + p.stderr)
    logs = glob.glob(os.path.join(work, "backup", "*_build.log"))
    text = p.stdout
    for lg in logs:
        with open(lg, encoding="utf-8", errors="replace") as fh:
            text += fh.read()
    res["user_sheet_warning"] = f"사용자 추가 시트 '{USER_SHEET}'" in text
    res["xlsx_moved"] = not os.path.exists(xlsx) and any(f.endswith(".xlsx") for f in os.listdir(os.path.join(work, "backup")))
    res["backup_files"] = sorted(os.listdir(os.path.join(work, "backup"))) if os.path.isdir(os.path.join(work, "backup")) else []
    res["vbom_after"] = vbom_state()
    checks = {}
    if p.returncode == 0 and os.path.exists(out):
        with Workbook(out) as wb:
            def recs(n):
                c, r = wb.read_table(n)
                return [dict(zip(c, x)) for x in r if any(v not in (None, "") for v in x)]
            tr = recs("tblTrades")
            checks["trades_rows"] = (len(tr), exp["trades"])
            checks["trade_row"] = any(r.get("전략") == "V7시험" and r.get("종목코드") == "000270" for r in tr)
            wa = recs("tblWatch")
            checks["watch_rows"] = (len(wa), exp["watch"])
            checks["watch_row"] = any(r.get("종목코드") == exp["outside_code"] and r.get("그룹") == "V7" for r in wa)
            st = {r["키"]: r["값"] for r in recs("tblSettings")}
            checks["setting_max_weight"] = st.get("max_weight") == exp["max_weight"]
            ov = recs("tblOverride")
            checks["override_rows"] = (sum(1 for r in ov if r.get("메모") in ("테마 변경 시험", "대회편입 추가 시험")), 2)
            if "snap_key" in exp:
                sn = recs("tblEstSnap")
                checks["snap_rows"] = (len(sn), exp["snap_rows"])
                checks["snap_sentinel"] = any(abs((r.get("FwdPER") or 0) - SENTINEL_PER) < 1e-9 for r in sn)
            sheets = list(wb.sheets)
            checks["user_sheet_absent"] = USER_SHEET not in sheets
            # 이관 직후 build 모드 새로 고침으로 대회종목·종목DB에도 반영됐는지(수정표 효과)
            co = recs("tblCompany")
            checks["override_theme_applied"] = any(r.get("종목코드") == "005380" and r.get("대테마") == "V7테마" for r in co)
            checks["override_add_applied"] = any(r.get("종목코드") == exp["outside_code"] and r.get("유니버스") == "대회" for r in co)
    res["checks"] = checks

    def good(v):
        return v is True or (isinstance(v, tuple) and v[0] == v[1])
    ok = (p.returncode == 0 and res["user_sheet_warning"] and res["xlsx_moved"] and res["vbom_after"] == res["vbom_before"]
          and checks and all(good(v) for v in checks.values()))
    res["pass"] = ok
    print(f"V7 이관 왕복: 빌드 종료 코드 {p.returncode} ({res['build_seconds']}초), 사용자 시트 경고 {res['user_sheet_warning']}, "
          f".xlsx 백업 폴더로 이동 {res['xlsx_moved']}, AccessVBOM {res['vbom_before']} → {res['vbom_after']}")
    for k, v in checks.items():
        print(f"  {'✓' if good(v) else '✗'} {k}: {v}")
    print(f"  판정: {'통과' if ok else '실패'}")
    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump(res, fh, ensure_ascii=False, indent=1, default=str)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
