# -*- coding: utf-8 -*-
"""기존 기능 회귀 검사: NAV·성과 지표 독립 재계산, 시트 오류 셀 검사.

통합문서는 Excel로 열지 않고 파일을 직접 읽습니다(`xlsx_tables` — 저장된 계산 결과 값).

1. NAV 재구성: [매매일지] tblTrades 입력(수수료·세금 빈칸이면 설정 요율로 원 미만 절사 — 기존 규칙)과 KIS 일봉 종가(수정주가,
   FHKST03010100)로 매일 NAV = 초기자금 + 누적 현금흐름 + Σ 보유수량 × 종가(종가가 없으면 마지막 매매 단가)를 KOSPI 세션
   (FHPTJ04040000 0001 날짜)마다 다시 계산해 tblNAV 순자산과 대조한다(기준 행 = 대회 시작 직전 세션).
2. 성과 지표: 재구성한 NAV의 일간수익률로 누적 수익률·연율화 변동성·샤프(rf 적용/0)·소르티노·MDD·현재 낙폭·연율화 수익률·칼마·
   베타·알파·상관·추적오차·정보비율·승률·BM 대비 승률을 계산해 [성과] 시트 지표 칸과 대조한다(무위험수익률은 [성과]의 표시값을 입력으로 씀).
3. 손익 항등식: 실현손익 합(tblTradeLog) + 평가손익 합(tblHoldings) = 최신 순자산 − 초기자금(README 5장).
4. 오류 셀: 모든 시트의 셀 중 오류 값(#N/A·#VALUE! 등 — 파일의 t="e" 칸)을 시트별로 센다. [시세판]의 숨김 도우미 열에 있는
   의도된 #N/A는 따로 보고한다(기존 검증과 같은 예외).

    python excel_dashboard/tools/verify_regress.py --workbook <저장한.xlsm> --token-source <원천.xlsm> [--json 결과.json]
종료 코드: 0 통과 / 1 불일치·오류 셀 / 2 토큰 갱신 필요
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import math
import os
import statistics
import sys
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
sys.path.insert(0, DASH)
sys.path.insert(0, HERE)
from kis_dev import KisClient, TokenError  # noqa: E402
from xlsx_tables import Workbook, _M  # noqa: E402

# [성과] 지표 칸: B열 이름 → 재계산 키
PERF_LABELS = {
    "누적 수익률": "cum", "연율화 수익률": "cagr", "연율화 변동성": "vol", "샤프지수 (rf 적용)": "sharpe",
    "샤프지수 (rf=0)": "sharpe0", "소르티노": "sortino", "최대낙폭 (MDD)": "mdd", "현재 낙폭": "curdd",
    "칼마 비율": "calmar", "베타 (vs BM)": "beta", "알파 (연, 젠센)": "alpha", "상관계수 (vs BM)": "corr",
    "추적오차 (연)": "te", "정보비율": "ir", "일간 승률": "hit", "BM 대비 승률": "hitbm", "순자산 (NAV)": "nav",
    "무위험수익률(연)": "rf", "초과 수익률": "excess",
}


def sheet_cells(wb: Workbook, sheet: str) -> dict[str, tuple]:
    """시트의 모든 칸 {주소: (값, 형식 t)} — 값은 xlsx_tables 규칙으로 변환."""
    out = {}
    part = wb.sheets[sheet]
    with wb.zip.open(part) as fh:
        for _ev, el in ET.iterparse(fh, events=("end",)):
            if el.tag == _M + "c":
                out[el.get("r")] = (wb._cell_value(el), el.get("t", "n"))
            elif el.tag == _M + "row":
                el.clear()
    return out


def error_cells(wb: Workbook) -> dict[str, list[str]]:
    """시트별 오류 값 칸 주소 목록."""
    out = {}
    for sheet, part in wb.sheets.items():
        bad = []
        with wb.zip.open(part) as fh:
            for _ev, el in ET.iterparse(fh, events=("end",)):
                if el.tag == _M + "c" and el.get("t") == "e":
                    v = el.find(_M + "v")
                    bad.append(f"{el.get('r')}={v.text if v is not None else '?'}")
                elif el.tag == _M + "row":
                    el.clear()
        if bad:
            out[sheet] = bad
    return out


def records(wb: Workbook, name: str) -> list[dict]:
    cols, rows = wb.read_table(name)
    return [dict(zip(cols, r)) for r in rows]


def num(v):
    if v is None or v == "" or v == "-":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def d_of(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def ymd(s):
    s = str(s or "").strip()
    return dt.date(int(s[:4]), int(s[4:6]), int(s[6:8])) if len(s) == 8 and s.isdigit() else None


def main() -> int:
    ap = argparse.ArgumentParser(description="V9 NAV·성과 독립 재계산과 오류 셀 검사 (토큰 값은 출력하지 않음)")
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--token-source", required=True)
    ap.add_argument("--json", default=None)
    a = ap.parse_args()

    with Workbook(a.workbook) as wb:
        trades = records(wb, "tblTrades")
        settings = {r["키"]: r["값"] for r in records(wb, "tblSettings") if r.get("키")}
        nav_wb = records(wb, "tblNAV")
        tradelog = records(wb, "tblTradeLog")
        holdings = records(wb, "tblHoldings")
        perf_cells = sheet_cells(wb, "성과")
        errors = error_cells(wb)

    init = num(settings.get("init_capital")) or 1e8
    start = d_of(settings.get("start_date"))
    fee_rate = num(settings.get("fee_rate")) or 0.00015
    tax_rate = num(settings.get("tax_rate")) or 0.002
    tr = []
    for r in trades:
        d, code, side = d_of(r.get("일자")), str(r.get("종목코드") or "").strip(), str(r.get("구분") or "").strip()
        q, p = num(r.get("수량")), num(r.get("단가"))
        if not (d and code and side in ("매수", "매도") and q and p):
            continue
        amt = q * p
        fee = num(r.get("수수료"))
        fee = math.floor(amt * fee_rate) if fee is None else fee
        tax = 0.0 if side == "매수" else (num(r.get("세금")) if num(r.get("세금")) is not None else math.floor(amt * tax_rate))
        cf = -(amt + fee) if side == "매수" else amt - fee - tax
        tr.append({"d": d, "code": code, "q": q if side == "매수" else -q, "p": p, "cf": cf, "amt": amt})
    tr.sort(key=lambda x: x["d"])
    first = min([x["d"] for x in tr] + ([start] if start else []))

    try:
        kis = KisClient(token_path=a.token_source)
    except TokenError as e:
        print(f"멈춤: {e}")
        return 2
    today = max(d_of(r.get("일자")) for r in nav_wb if d_of(r.get("일자")))
    look = first - dt.timedelta(days=20)

    def market(code, iscd1):
        d = today.strftime("%Y%m%d")
        b, _ = kis.get("/uapi/domestic-stock/v1/quotations/inquire-investor-daily-by-market", "FHPTJ04040000",
                       {"FID_COND_MRKT_DIV_CODE": "U", "FID_INPUT_ISCD": code, "FID_INPUT_DATE_1": d,
                        "FID_INPUT_ISCD_1": iscd1, "FID_INPUT_DATE_2": d, "FID_INPUT_ISCD_2": code})
        return {ymd(x.get("stck_bsop_date")): float(x["bstp_nmix_prpr"]) for x in b.get("output") or []
                if ymd(x.get("stck_bsop_date")) and x.get("bstp_nmix_prpr")}
    kospi = market("0001", "KSP")
    days_all = sorted(d for d in kospi if d <= today)
    prior = [d for d in days_all if d < first]
    base = prior[-1] if prior else first - dt.timedelta(days=1)
    grid = [base] + [d for d in days_all if d >= first]
    closes = {}
    for code in sorted({x["code"] for x in tr}):
        b, _ = kis.get("/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice", "FHKST03010100",
                       {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": code, "FID_INPUT_DATE_1": look.strftime("%Y%m%d"),
                        "FID_INPUT_DATE_2": today.strftime("%Y%m%d"), "FID_PERIOD_DIV_CODE": "D", "FID_ORG_ADJ_PRC": "0"})
        closes[code] = {ymd(x["stck_bsop_date"]): float(x["stck_clpr"]) for x in b.get("output2") or []
                        if ymd(x.get("stck_bsop_date")) and x.get("stck_clpr")}
    nav, cashes = [], []
    for d in grid:
        cash = init + sum(x["cf"] for x in tr if x["d"] <= d)
        mv = 0.0
        for code in closes:
            q = sum(x["q"] for x in tr if x["code"] == code and x["d"] <= d)
            if q == 0:
                continue
            px = closes[code].get(d)
            if px is None:     # 그날 종가가 없으면 직전 종가(기존 쿼리는 날짜 격자에 맞춰 채움) → 그것도 없으면 마지막 매매 단가
                prev = [k for k in closes[code] if k <= d]
                px = closes[code][max(prev)] if prev else [x["p"] for x in tr if x["code"] == code and x["d"] <= d][-1]
            mv += q * px
        nav.append(cash + mv)
        cashes.append(cash)
    wb_nav = {d_of(r.get("일자")): num(r.get("순자산")) for r in nav_wb if d_of(r.get("일자"))}
    nav_cmp = [(d, v, wb_nav.get(d)) for d, v in zip(grid, nav)]
    nav_bad = [(d.isoformat(), v, w) for d, v, w in nav_cmp if w is None or abs(v - w) > 0.5]

    # 성과 지표
    r = [nav[i] / nav[i - 1] - 1 for i in range(1, len(nav))]
    kc = [kospi.get(d) for d in grid]
    b = [kc[i] / kc[i - 1] - 1 for i in range(1, len(kc))]
    labels = {}
    for ref, (v, _t) in perf_cells.items():
        if ref.startswith("B") and isinstance(v, str) and v.strip() in PERF_LABELS:
            labels[PERF_LABELS[v.strip()]] = perf_cells.get("C" + ref[1:], (None, None))[0]
    rf_y = num(labels.get("rf")) or 0.0
    rf = rf_y / 252
    n = len(r)
    sd = statistics.stdev(r)
    mean = statistics.fmean(r)
    peak, mdd = -1e300, 0.0
    dds = []
    for v in nav:
        peak = max(peak, v)
        dds.append(v / peak - 1)
    mdd = min(dds)
    cum = nav[-1] / init - 1
    cagr = (1 + cum) ** (252 / n) - 1
    downside = [min(x - rf, 0) if x < rf else 0 for x in r]
    ddv = math.sqrt(sum(x * x for x in downside) / n)
    mb = statistics.fmean(b)
    cov = sum((x - mean) * (y - mb) for x, y in zip(r, b)) / (n - 1)
    beta = cov / statistics.variance(b)
    ex = [x - y for x, y in zip(r, b)]
    bmcum = 1.0
    for y in b:
        bmcum *= 1 + y
    mine = {
        "cum": cum, "cagr": cagr, "vol": sd * math.sqrt(252), "sharpe": (mean - rf) / sd * math.sqrt(252),
        "sharpe0": mean / sd * math.sqrt(252), "sortino": (mean - rf) / ddv * math.sqrt(252) if ddv else None,
        "mdd": mdd, "curdd": dds[-1], "calmar": cagr / abs(mdd) if mdd else None, "beta": beta,
        "alpha": (mean - rf - beta * (mb - rf)) * 252, "corr": statistics.correlation(r, b),
        "te": statistics.stdev(ex) * math.sqrt(252), "ir": statistics.fmean(ex) * 252 / (statistics.stdev(ex) * math.sqrt(252)),
        "hit": sum(1 for x in r if x > 0) / n, "hitbm": sum(1 for x, y in zip(r, b) if x > y) / n, "nav": nav[-1],
        "excess": cum - (bmcum - 1),
    }
    perf_rows = []
    for k, v in mine.items():
        w = num(labels.get(k))
        ok = (v is None and w is None) or (v is not None and w is not None and abs(v - w) <= max(1e-6 * abs(w), 1e-9 if k != "nav" else 0.5))
        perf_rows.append({"metric": k, "recalc": v, "workbook": labels.get(k), "ok": ok})

    # 손익 항등식
    realized = sum(num(x.get("실현손익")) or 0 for x in tradelog)
    unreal = sum(num(x.get("평가손익")) or 0 for x in holdings)
    identity = {"realized": realized, "unrealized": unreal, "nav_minus_init": (wb_nav.get(grid[-1]) or 0) - init}
    identity["ok"] = abs(realized + unreal - identity["nav_minus_init"]) <= 1.0

    # 오류 셀 — 시세판 숨김 도우미의 의도된 #N/A는 따로
    intended, unexpected = {}, {}
    for sheet, cells in errors.items():
        if sheet == "시세판":
            intended[sheet] = cells
        else:
            unexpected[sheet] = cells

    print(f"V9 NAV 재구성: {len(grid)}일(기준 {grid[0]} ~ {grid[-1]}), 불일치 {len(nav_bad)}일, 최신 NAV 재계산 {nav[-1]:,.0f} / 통합문서 {wb_nav.get(grid[-1]) or 0:,.0f}")
    for row in perf_rows:
        print(f"  {'✓' if row['ok'] else '✗'} {row['metric']}: 재계산 {row['recalc']} / 통합문서 {row['workbook']}")
    print(f"  손익 항등식: 실현 {realized:,.0f} + 평가 {unreal:,.0f} = {realized + unreal:,.0f} vs 순자산-초기자금 {identity['nav_minus_init']:,.0f} → {'일치' if identity['ok'] else '불일치'}")
    print(f"  오류 셀(시세판 외): {({k: len(v) for k, v in unexpected.items()}) or '없음'} · 시세판(숨김 도우미 예외 확인용): {({k: len(v) for k, v in intended.items()}) or '없음'}")
    for k, v in unexpected.items():
        print(f"    {k}: {v[:12]}")
    if intended:
        print(f"    시세판: {intended['시세판'][:12]}")
    ok = not nav_bad and all(x["ok"] for x in perf_rows) and identity["ok"] and not unexpected
    print(f"  KIS 호출 {kis.calls}회 · 판정: {'통과' if ok else '확인 필요'}")
    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump({"nav_days": len(grid), "nav_mismatch": nav_bad, "perf": perf_rows, "identity": identity,
                       "errors_unexpected": unexpected, "errors_quote_board": intended, "kis_calls": kis.calls},
                      fh, ensure_ascii=False, indent=1, default=str)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
