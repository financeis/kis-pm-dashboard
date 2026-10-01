# -*- coding: utf-8 -*-
"""버튼 매크로·모두 새로 고침 통합 검증 (T24 — spec §8 V3·V4, R20 소요 시간 실측).

대시보드 통합문서(.xlsm)를 **자기 Excel 인스턴스(DispatchEx, 숨김)**로 열어 버튼 매크로를 COM으로 실행하고
(`Application.Run` — 사용자가 버튼을 누른 것과 같은 VBA 경로), 단계마다 소요 시간·`…_최근조회`/`…_상태` 칸·
`tblRunCtl`의 `last_summary`·표별 행 수와 `상태` 요약을 JSON으로 남깁니다. `refreshall` 단계는 [모두 새로 고침]을
실행해 시간을 재고, 버튼 전용 표(spec R20)의 행 수·조회시각이 바뀌지 않았는지(= 모두 새로 고침에서 빠짐) 확인합니다.

안전 규칙 (plan T24 '토큰 예외 규칙'·하드 게이트)
- 열기 직전과 매크로·모두 새로 고침 실행 직전마다 토큰 원천을 파일 파싱으로 다시 확인합니다(기준 210분 미만이면 멈춤,
  종료 코드 2 — 오케스트레이터가 갱신). 토큰 값은 읽지도 출력하지도 않습니다(토큰 표는 expires·status·checked만).
- 통합문서는 한 번에 하나만 엽니다. 자기 PID만 종료하며, 매크로가 한도 시간을 넘기면 자기 Excel만 강제 종료합니다
  (저장 전이면 디스크의 파일은 그대로).
- `CalculateUntilAsyncQueriesDone`은 쓰지 않습니다(교착). 새로 고침 완료는 `QueryTable.Refreshing` 폴링으로 확인합니다.
- 실패 경로·시험 시계(now_override)·설정 변경 단계는 **복사본**에서만 쓰세요(실제 통합문서에서는 버튼과 저장만).

    python excel_dashboard/tools/verify_buttons.py --workbook <경로.xlsm> --token-source <원천.xlsm> \\
        --steps full save quick sector analysis=005930 save --json <스크래치>/run1.json

단계 (적힌 순서대로)
    full | quick | sector | analysis[=종목코드]   매크로 RefreshFull·RefreshQuick·RefreshSector·RunAnalysis 실행
    refreshall                                  [모두 새로 고침] 실행·시간 측정·버튼 표 불변 확인
    call:함수이름                                VBA 공개 함수를 불러 반환 글자를 기록(예: call:RefreshedTables — 예열 확인)
    override:종목코드=값                         (복사본 시험용) 수정표 tblOverride에 행 추가 — 값은 대회편입(추가/제외)
    emulate:quick | emulate:sector              (복사본 측정용) 매크로의 예열 없이 같은 단계 순서로 mode·started를 쓰고 표를
                                                동기 새로 고침만 해 시간 측정 — 첫 실행 예열(0번)의 비용을 따로 재기 위함
    save                                        통합문서 저장(새로 고침이 모두 끝난 뒤)
    snapshot                                    표 요약만 기록
    runctl:키=값 | setting:키=값                 tblRunCtl·tblSettings 값 바꾸기(복사본 시험용; 값 비우기는 '키=')
                                                값이 yyyy-mm-dd hh:mm 형식이면 일시로 넣음. '키=@bad-path'는 없는 경로,
                                                '키=@restore'는 이 실행에서 처음 바꾸기 전 값으로 되돌림(값은 출력 안 함)
종료 코드: 0 모든 단계 실행(판정은 JSON 참고) / 1 단계 실행 오류 / 2 토큰 갱신 필요 / 3 인자 오류
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from verify_common import (  # noqa: E402
    BUTTON_TABLES, TokenGateError, Watchdog, close_workbook, com_retry, default_token_source, find_table, first_errors,
    get_key_value, iter_tables, log, name_values, open_workbook, query_tables, runctl, set_key_value, table_summaries,
    token_gate, wait_idle,
)

BAD_PATH = r"C:\T24_없는_폴더\kis_devlp.yaml"     # 실패 경로 시험용(존재하지 않는 설정 파일)

MACROS = {
    "full": ("RefreshFull", "전체", 90 * 60),
    "quick": ("RefreshQuick", "시세", 15 * 60),
    "sector": ("RefreshSector", "업종", 15 * 60),
    "analysis": ("RunAnalysis", "분석", 20 * 60),
}
# 매크로별로 결과를 볼 표(단계 순서) — 요약 JSON에 넣을 범위
STEP_TABLES = {
    "full": ["tblClass", "tblSessions", "tblPxStore", "tblPxMetrics", "tblFlowU", "tblTarget", "tblEst", "tblEstSnap",
             "tblEvents", "tblFin", "tblCSL", "tblCompany", "tblThemeAgg", "tblSectorKRX"],
    "quick": ["tblSessions", "tblPxStore", "tblPxMetrics", "tblCompany", "tblThemeAgg"],
    "sector": ["tblSectorKRX", "tblThemeAgg"],
    "analysis": None,     # tblA_* 전부
}


def parse_value(v: str):
    """단계 값 표기: 빈 값 → None, yyyy-mm-dd hh:mm[:ss] → 일시, 그 밖 → 글자."""
    v = v.strip()
    if v == "":
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return dt.datetime.strptime(v, fmt)
        except ValueError:
            pass
    return v


def is_analysis(name: str) -> bool:
    return name.lower().startswith("tbla_") and not name.lower().endswith("seed")


def run_macro(xl, wb, pid, step: str, arg: str | None, token_source: str) -> dict:
    macro, label, limit = MACROS[step]
    token_gate(token_source, what=f"[{label}] 실행 전")
    res = {"step": step, "macro": macro}
    if step == "analysis" and arg:
        rng = wb.Names("분석코드").RefersToRange
        if str(rng.NumberFormat) != "@":
            rng.NumberFormat = "@"
        rng.Value2 = arg
        res["code"] = arg
    wait_idle(wb, timeout=600)
    before = name_values(wb)
    log(f"▶ {macro} 실행 ([{label}])")
    t0 = time.time()
    with Watchdog(macro, pid, limit) as wd:
        try:
            com_retry(lambda: xl.Run(macro), timeout=60)
            res["ok"] = True
        except Exception as e:  # noqa: BLE001 — 매크로 밖으로 오류가 나오면 기록하고 계속
            res["ok"] = False
            res["error"] = str(e)[:300]
    res["seconds"] = round(time.time() - t0, 1)
    if wd.killed:
        res["ok"] = False
        res["error"] = "한도 시간 초과로 강제 종료"
        return res
    after = name_values(wb)
    rc = runctl(wb)
    res["names_before"] = {k: v for k, v in before.items() if k.startswith(label)}
    res["names_after"] = {k: v for k, v in after.items() if k.startswith(label)}
    res["mode_after"] = rc.get("mode")
    res["started"] = rc.get("started")
    res["last_summary"] = rc.get("last_summary")
    tabs = STEP_TABLES[step]
    if tabs is None:
        tabs = [lo.Name for _ws, lo in iter_tables(wb) if is_analysis(lo.Name)]
    res["tables"] = table_summaries(wb, tabs)
    errs = {}
    for t in tabs:
        st = res["tables"].get(t, {}).get("status") or {}
        if st.get("오류") or st.get("이전"):
            errs[t] = first_errors(wb, t)
    res["first_errors"] = errs
    log(f"■ {macro} 끝: {res['seconds']}초, 상태 칸 = {res['names_after'].get(label + '_상태')!r}, mode = {res['mode_after']}")
    for line in str(res["last_summary"] or "").splitlines():
        log(f"    {line}")
    return res


EMULATE = {"quick": ["tblToken", "tblSessions", "tblPxStore", "tblPxMetrics", "tblCompany", "tblThemeAgg"],
           "sector": ["tblToken", "tblSectorKRX", "tblThemeAgg"]}


def emulate(xl, wb, pid, mode: str, token_source: str) -> dict:
    """매크로 RunButton의 본 단계만(예열 없이) 흉내 낸다: mode·started 쓰기 → 표마다 BackgroundQuery 끄고 Refresh(False) → mode=build."""
    token_gate(token_source, what=f"[{mode} 흉내] 실행 전")
    wait_idle(wb, timeout=600)
    res = {"step": f"emulate:{mode}", "tables": {}}
    t0 = time.time()
    set_key_value(wb, "tblRunCtl", "mode", mode)
    set_key_value(wb, "tblRunCtl", "started", dt.datetime.now())
    with Watchdog(f"emulate:{mode}", pid, 15 * 60):
        for t in EMULATE[mode]:
            lo = find_table(wb, t)
            qt = lo.QueryTable
            bg = qt.BackgroundQuery
            t1 = time.time()
            try:
                qt.BackgroundQuery = False
                com_retry(lambda: qt.Refresh(False), timeout=60)
                ok = True
            except Exception as e:  # noqa: BLE001
                ok = str(e)[:120]
            finally:
                qt.BackgroundQuery = bg
            res["tables"][t] = {"seconds": round(time.time() - t1, 1), "ok": ok}
    set_key_value(wb, "tblRunCtl", "mode", "build")
    res["seconds"] = round(time.time() - t0, 1)
    res["summary"] = table_summaries(wb, EMULATE[mode])
    log(f"■ emulate:{mode} 끝: {res['seconds']}초 " + ", ".join(f"{k} {v['seconds']}s" for k, v in res["tables"].items()))
    return res


def add_override(wb, code: str, value: str) -> None:
    """수정표(tblOverride) 끝에 '종목코드·대회편입·메모' 행을 더한다(마지막 행이 비었으면 그 행에)."""
    lo = find_table(wb, "tblOverride")
    hdr = [str(h) for h in lo.HeaderRowRange.Value2[0]]
    last = lo.ListRows(lo.ListRows.Count).Range if lo.ListRows.Count else None
    row = last if last is not None and all(v in (None, "") for v in last.Value2[0]) else lo.ListRows.Add().Range
    c = row.Cells(1, hdr.index("종목코드") + 1)
    c.NumberFormat = "@"
    c.Value2 = code
    row.Cells(1, hdr.index("대회편입") + 1).Value2 = value
    row.Cells(1, hdr.index("메모") + 1).Value2 = "T24 시험(복사본)"


def connection_flags(wb) -> dict:
    """표별 '모두 새로 고침 시 이 연결 새로 고침'·'파일 열 때 새로 고침' 설정."""
    out = {}
    for _ws, lo in iter_tables(wb):
        if lo.SourceType == 1:      # xlSrcRange — 정적 표
            continue
        try:
            qt = lo.QueryTable
            out[lo.Name] = {"refresh_all": bool(qt.WorkbookConnection.RefreshWithRefreshAll),
                            "on_open": bool(qt.RefreshOnFileOpen), "period": int(qt.RefreshPeriod)}
        except Exception as e:  # noqa: BLE001
            out[lo.Name] = {"error": str(e)[:120]}
    return out


def run_refresh_all(xl, wb, pid, token_source: str) -> dict:
    token_gate(token_source, what="[모두 새로 고침] 실행 전")
    qts = query_tables(wb)
    wait_idle(wb, timeout=600, qts=qts)
    flags = connection_flags(wb)
    before = table_summaries(wb)
    log("▶ 모두 새로 고침(RefreshAll)")
    t0 = time.time()
    com_retry(lambda: wb.RefreshAll(), timeout=60)
    t_call = time.time() - t0
    with Watchdog("RefreshAll", pid, 30 * 60):
        idle = wait_idle(wb, timeout=30 * 60, settle=3.0, qts=qts)
    secs = time.time() - t0 - 3.0          # 마지막 '조용한' 확인 3초는 빼고 기록
    time.sleep(2.0)                      # 새로 고침 직후의 백그라운드 재평가가 끝날 틈
    after = table_summaries(wb)
    excluded = sorted(t for t, f in flags.items() if f.get("refresh_all") is False)
    included = sorted(t for t, f in flags.items() if f.get("refresh_all") is True)
    expected_ex = sorted([t for t in flags if t in BUTTON_TABLES or is_analysis(t)])
    unchanged, changed = [], []
    for t in expected_ex:
        b, a = before.get(t, {}), after.get(t, {})
        same = (b.get("rows") == a.get("rows") and b.get("ts_max") == a.get("ts_max") and b.get("ts_min") == a.get("ts_min"))
        (unchanged if same else changed).append(t)
    refreshed, stale = [], []
    for t in included:
        b, a = before.get(t, {}), after.get(t, {})
        if "ts_max" in a:
            (refreshed if a.get("ts_max") != b.get("ts_max") else stale).append(t)
    errs = {t: first_errors(wb, t) for t in included if (after.get(t, {}).get("status") or {}).get("이전")
            or (after.get(t, {}).get("status") or {}).get("오류")}
    res = {"step": "refreshall", "seconds": round(secs, 1), "call_seconds": round(t_call, 1),
           "excluded": excluded, "included": included, "expected_excluded": expected_ex,
           "exclusion_matches": excluded == expected_ex,
           "button_tables_unchanged": unchanged, "button_tables_changed": changed,
           "general_refreshed": refreshed, "general_not_refreshed": stale,
           "flags": flags, "before": {t: before[t] for t in before if t in included + expected_ex},
           "after": {t: after[t] for t in after if t in included + expected_ex}, "errors": errs}
    log(f"■ 모두 새로 고침 끝: {res['seconds']}초 · 제외 {len(excluded)}개(기대와 일치={res['exclusion_matches']}) · "
        f"버튼 표 불변 {len(unchanged)}/{len(expected_ex)} · 조회시각 갱신된 일반 표 {len(refreshed)}개")
    if changed:
        log(f"  ✗ 바뀐 버튼 표: {changed}")
    return res


def main() -> int:
    ap = argparse.ArgumentParser(description="버튼 매크로·모두 새로 고침 통합 검증 (토큰 값은 출력하지 않음)")
    ap.add_argument("--workbook", required=True, help="열 통합문서(.xlsm)")
    ap.add_argument("--token-source", default=None, help="토큰 원천 통합문서(기본: 주 저장소 KIS_PM_Dashboard.xlsm)")
    ap.add_argument("--steps", nargs="+", required=True, help="단계 목록(머리 주석 참고)")
    ap.add_argument("--json", default=None, help="결과 JSON 경로(저장소 밖)")
    ap.add_argument("--visible", action="store_true", help="자기 Excel 창을 보이게(사람이 지켜볼 때만)")
    a = ap.parse_args()

    for s in a.steps:
        head = s.split("=", 1)[0].split(":", 1)[0]
        if head not in MACROS and head not in ("save", "snapshot", "refreshall", "runctl", "setting", "emulate", "override", "call"):
            print(f"알 수 없는 단계: {s}")
            return 3
    # 실패 경로·시험 시계·설정 변경·수정표 추가·흉내 측정은 복사본에서만(사용자 통합문서 = 토큰 원천에서는 거부)
    copy_only = [s for s in a.steps if s.split(":", 1)[0] in ("runctl", "setting", "override", "emulate")]
    src = a.token_source or default_token_source()
    if copy_only and os.path.normcase(os.path.abspath(a.workbook)) == os.path.normcase(os.path.abspath(src)):
        print(f"복사본 전용 단계는 실제 통합문서(토큰 원천)에서 실행하지 않습니다: {copy_only}")
        return 3
    out = {"workbook": a.workbook, "started": dt.datetime.now().isoformat(sep=" ", timespec="seconds"), "steps": []}

    def dump():
        if a.json:
            with open(a.json, "w", encoding="utf-8") as fh:
                json.dump(out, fh, ensure_ascii=False, indent=1, default=str)

    try:
        xl, wb, pid, info = open_workbook(a.workbook, visible=a.visible, token_source=a.token_source)
    except TokenGateError as e:
        log(f"멈춤: {e}")
        out["blocked"] = str(e)
        dump()
        return 2
    out["open"] = info
    out["initial_names"] = name_values(wb)
    out["initial_runctl"] = {k: v for k, v in runctl(wb).items() if k != "last_summary"}
    rc = 0
    saved = False
    saved_vals: dict = {}
    try:
        for s in a.steps:
            if s == "save":
                wait_idle(wb, timeout=600)
                t0 = time.time()
                com_retry(lambda: wb.Save())
                saved = True
                out["steps"].append({"step": "save", "seconds": round(time.time() - t0, 1)})
                log(f"저장 완료({time.time() - t0:.1f}초)")
            elif s == "snapshot":
                out["steps"].append({"step": "snapshot", "tables": table_summaries(wb), "names": name_values(wb)})
            elif s == "refreshall":
                out["steps"].append(run_refresh_all(xl, wb, pid, a.token_source))
            elif s.startswith("call:"):
                fn = s.split(":", 1)[1].strip()
                val = com_retry(lambda: xl.Run(fn), timeout=60)
                out["steps"].append({"step": s, "value": val})
                log(f"VBA {fn}() = {val!r}")
            elif s.startswith("override:"):
                code, val = (x.strip() for x in s.split(":", 1)[1].split("=", 1))
                add_override(wb, code, val)
                out["steps"].append({"step": s})
                log(f"수정표 행 추가: {code} 대회편입={val}")
            elif s.startswith("emulate:"):
                out["steps"].append(emulate(xl, wb, pid, s.split(":", 1)[1], a.token_source))
            elif s.startswith("runctl:") or s.startswith("setting:"):
                kind, kv = s.split(":", 1)
                k, v = (x.strip() for x in kv.split("=", 1))
                table = "tblRunCtl" if kind == "runctl" else "tblSettings"
                if (table, k) not in saved_vals:
                    saved_vals[(table, k)] = get_key_value(wb, table, k)
                if v == "@restore":
                    val = saved_vals[(table, k)]
                    val = None if val in (None, "") else val
                elif v == "@bad-path":
                    val = BAD_PATH
                else:
                    val = parse_value(v)
                set_key_value(wb, table, k, val)
                out["steps"].append({"step": s})
                log(f"값 변경: {kind} {k} = {'(원래 값으로)' if v == '@restore' else repr(v)}")
            else:
                step, _, arg = s.partition("=")
                r = run_macro(xl, wb, pid, step, arg or None, a.token_source)
                out["steps"].append(r)
                if not r.get("ok"):
                    rc = 1
                    if "강제 종료" in str(r.get("error")):
                        dump()
                        return 1
            dump()
    except TokenGateError as e:
        log(f"멈춤: {e}")
        out["blocked"] = str(e)
        rc = 2
    except Exception as e:  # noqa: BLE001
        import traceback
        log(f"단계 실행 오류: {e}")
        log(traceback.format_exc()[-1500:])
        out["error"] = str(e)[:500]
        rc = 1
    finally:
        try:
            out["final_names"] = name_values(wb)
        except Exception:  # noqa: BLE001 — Excel이 이미 끝났으면 생략
            out["final_names"] = None
        out["finished"] = dt.datetime.now().isoformat(sep=" ", timespec="seconds")
        dump()
        try:
            close_workbook(xl, wb, pid, save=False)
        except Exception as e:  # noqa: BLE001
            log(f"닫기 중 오류(무시하고 자기 Excel 종료): {str(e)[:120]}")
        log(f"닫음(저장 단계 실행={saved})")
    return rc


if __name__ == "__main__":
    sys.exit(main())
