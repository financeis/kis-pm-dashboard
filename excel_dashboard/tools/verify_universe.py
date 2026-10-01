# -*- coding: utf-8 -*-
"""대회 종목 명단 경계값 독립 재계산 (T24 — spec §8 V2, spec R1).

고정 명단 `data/contest_universe_20260930.csv`의 종목 수를 보고하고, 기준(시가총액 ≥ 1,000억 · 2026-09-22·23·28·29·30
5세션 평균 거래대금 ≥ 25억, 둘 다 '이상')의 **경계 부근 양쪽** 종목을 원천 일봉으로 다시 판정해 명단 포함 여부와 대조합니다.

- 명단 안쪽 표본: CSV 값 기준 여유(min(시총/1,000억, 평균거래대금/25억))가 가장 작은 종목들.
- 명단 바깥 표본: KIS 종목 마스터(인증 불필요)의 개별종목(ST·FS·DR, 우선주·SPAC 제외) 중 명단에 없고 마스터 시총이 0.8 × 1,000억
  이상(또는 공란)인 종목을 모두 일봉으로 판정한 뒤, 기준에 가장 가까운 종목들. (마스터 파싱은 make_contest_universe의 함수를 쓰지만
  판정 값 — 9/30 종가·거래대금·주식수 — 은 여기서 따로 받아 계산한다.)
- 판정 원천: 국내주식기간별시세 FHKST03010100(원주가 FID_ORG_ADJ_PRC=1, 기간 9/22~9/30) — 9/30 종가, 세션별 거래대금(acml_tr_pbmn),
  응답 output1.lstn_stcn(조회 시점 상장주식수). 9/30 뒤 상장 변동이 있으면 lstn_stcn이 9/30 주식수와 다를 수 있어 CSV 주식수와 비교해 보고한다.
- 상장 전 세션(봉 없음 + 마스터 상장일 이후가 아님)은 평균에서 뺀다. 상장 중 거래 없는 세션은 0원.

    python excel_dashboard/tools/verify_universe.py --token-source <원천.xlsm> [--n-in 18] [--n-out 18] [--json 결과.json] [--csv 표본.csv]
종료 코드: 0 표본 전부 일치 / 1 불일치 / 2 토큰 갱신 필요
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
sys.path.insert(0, DASH)
sys.path.insert(0, HERE)
from kis_dev import KisClient, TokenError  # noqa: E402
import make_contest_universe as mcu  # noqa: E402

CSV_PATH = os.path.join(DASH, "data", "contest_universe_20260930.csv")
SESSIONS = ["20260922", "20260923", "20260928", "20260929", "20260930"]
BASE = "20260930"
MCAP_MIN = 100_000_000_000          # 1,000억 원
TURN_MIN = 2_500_000_000            # 25억 원


def fetch(kis: KisClient, code: str) -> tuple[int | None, dict[str, dict]]:
    body, _ = kis.get("/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice", "FHKST03010100",
                      {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": code, "FID_INPUT_DATE_1": SESSIONS[0],
                       "FID_INPUT_DATE_2": BASE, "FID_PERIOD_DIV_CODE": "D", "FID_ORG_ADJ_PRC": "1"})
    o1 = body.get("output1") or {}
    bars = {}
    for r in body.get("output2") or []:
        d = str(r.get("stck_bsop_date") or "").strip()
        if d:
            bars[d] = {"close": int(float(r.get("stck_clpr") or 0)), "turnover": int(float(r.get("acml_tr_pbmn") or 0))}
    shares = o1.get("lstn_stcn")
    return (int(float(shares)) if shares not in (None, "") else None), bars


def decide(bars: dict[str, dict], shares: int | None, list_date: str | None) -> dict:
    sess = [s for s in SESSIONS if not (list_date and s < list_date)]
    turnover = [bars[s]["turnover"] if s in bars else 0 for s in sess]
    priced = sorted(d for d, b in bars.items() if d <= BASE and b["close"] > 0)
    close = bars[BASE]["close"] if BASE in bars and bars[BASE]["close"] > 0 else (bars[priced[-1]]["close"] if priced else None)
    mcap = close * shares if close and shares else None
    avg = sum(turnover) / len(turnover) if turnover else None
    ok = mcap is not None and mcap >= MCAP_MIN and bool(turnover) and sum(turnover) >= TURN_MIN * len(turnover)
    return {"close": close, "shares": shares, "mcap_eok": mcap / 1e8 if mcap else None, "avg_eok": avg / 1e8 if avg is not None else None,
            "days": len(turnover), "included": ok}


def margin(mcap_eok: float | None, avg_eok: float | None) -> float:
    """기준에 대한 여유(1 = 경계). 판정 쪽(포함/제외)에 상관없이 경계까지의 거리를 잰다."""
    if mcap_eok is None or avg_eok is None:
        return 99.0
    a, b = mcap_eok / 1000, avg_eok / 25
    if a >= 1 and b >= 1:
        return min(a, b) - 1                   # 포함: 둘 중 가까운 쪽까지
    return max(1 - a if a < 1 else 0, 1 - b if b < 1 else 0)   # 제외: 미달한 쪽의 부족분


def main() -> int:
    ap = argparse.ArgumentParser(description="V2 대회 종목 명단 경계값 독립 재계산 (토큰 값은 출력하지 않음)")
    ap.add_argument("--token-source", required=True)
    ap.add_argument("--n-in", type=int, default=18)
    ap.add_argument("--n-out", type=int, default=18)
    ap.add_argument("--json", default=None)
    ap.add_argument("--csv", default=None)
    a = ap.parse_args()

    with open(CSV_PATH, encoding="utf-8-sig") as fh:
        listed = {r["종목코드"]: r for r in csv.DictReader(fh)}
    by_mkt = {}
    for r in listed.values():
        by_mkt[r["시장"]] = by_mkt.get(r["시장"], 0) + 1
    print(f"명단 종목 수: {len(listed)} ({', '.join(f'{k} {v}' for k, v in sorted(by_mkt.items()))})")
    try:
        kis = KisClient(token_path=a.token_source)
    except TokenError as e:
        print(f"멈춤: {e}")
        return 2
    cands, _counts = mcu.load_candidates()
    master = {r.code: r for r in cands}

    inside = sorted(listed.values(), key=lambda r: margin(float(r["시가총액_0930_억"]), float(r["평균거래대금_5일_억"])))[:a.n_in]
    pool_out = [r for r in cands if r.code not in listed and (r.mcap_eok is None or r.mcap_eok <= 0 or r.mcap_eok >= 800)]
    print(f"바깥 후보(마스터 시총 ≥ 800억 또는 공란, 명단 밖): {len(pool_out)}종목 → 모두 일봉 판정")
    results = []
    for r in inside:
        m = master.get(r["종목코드"])
        shares_now, bars = fetch(kis, r["종목코드"])
        d = decide(bars, shares_now, m.list_date if m else None)
        csv_shares = int(r["상장주식수"])
        d2 = decide(bars, csv_shares, m.list_date if m else None)
        results.append({"side": "명단 안", "code": r["종목코드"], "name": r["종목명"], "in_list": True,
                        "recalc_included": d["included"], "recalc_included_csv_shares": d2["included"],
                        "mcap_eok": d["mcap_eok"], "mcap_eok_csv_shares": d2["mcap_eok"], "csv_mcap_eok": float(r["시가총액_0930_억"]),
                        "avg_eok": d["avg_eok"], "csv_avg_eok": float(r["평균거래대금_5일_억"]), "days": d["days"],
                        "csv_days": int(r["거래일수"]), "shares_now": shares_now, "csv_shares": csv_shares,
                        "margin": margin(d2["mcap_eok"], d["avg_eok"])})
    outs = []
    for m in pool_out:
        shares_now, bars = fetch(kis, m.code)
        d = decide(bars, shares_now, m.list_date)
        outs.append({"side": "명단 밖", "code": m.code, "name": m.name, "in_list": False, "recalc_included": d["included"],
                     "recalc_included_csv_shares": None, "mcap_eok": d["mcap_eok"], "avg_eok": d["avg_eok"], "days": d["days"],
                     "shares_now": shares_now, "margin": margin(d["mcap_eok"], d["avg_eok"])})
    wrong_out = [o for o in outs if o["recalc_included"]]          # 명단 밖인데 기준 충족 → 명단 누락 의심
    outs.sort(key=lambda o: o["margin"])
    results += outs[:a.n_out]

    bad = [x for x in results if x["recalc_included"] != x["in_list"]
           and not (x["in_list"] and x.get("recalc_included_csv_shares"))]
    value_bad = [x for x in results if x["in_list"] and (
        abs((x["mcap_eok_csv_shares"] or 0) - x["csv_mcap_eok"]) > 0.01 or abs((x["avg_eok"] or 0) - x["csv_avg_eok"]) > 0.01
        or x["days"] != x["csv_days"])]
    print(f"표본 {len(results)}종목(명단 안 {sum(1 for x in results if x['in_list'])} · 밖 {sum(1 for x in results if not x['in_list'])}),"
          f" 판정 불일치 {len(bad)}, 명단 안 값(시총·평균거래대금·거래일수) 불일치 {len(value_bad)}, KIS 호출 {kis.calls}회")
    print(f"바깥 후보 전체 {len(outs)}종목 중 기준 충족(명단 누락 의심) {len(wrong_out)}종목")
    for x in results:
        flag = "✓" if x not in bad else "✗"
        print(f"  {flag} {x['side']} {x['code']} {x['name'][:10]:<10} 시총 {x['mcap_eok'] or 0:>10,.1f}억 · 5일평균 {x['avg_eok'] or 0:>8,.2f}억"
              f" · {x['days']}일 → 재판정 {'포함' if x['recalc_included'] else '제외'}"
              + (f" (CSV 주식수로 {'포함' if x['recalc_included_csv_shares'] else '제외'}, 주식수 {'같음' if x['shares_now'] == x['csv_shares'] else '다름'})"
                 if x["in_list"] else ""))
    for x in wrong_out:
        print(f"  ✗ 명단 밖 기준 충족: {x['code']} {x['name']} 시총 {x['mcap_eok']:.1f}억 · 평균 {x['avg_eok']:.2f}억")
    if a.csv:
        with open(a.csv, "w", encoding="utf-8-sig", newline="") as fh:
            keys = sorted({k for x in results for k in x})
            wr = csv.DictWriter(fh, fieldnames=keys)
            wr.writeheader()
            wr.writerows(results)
    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump({"listed": len(listed), "by_market": by_mkt, "sample": results, "mismatch": bad, "value_mismatch": value_bad,
                       "outside_pool": len(outs), "outside_meeting_rule": wrong_out, "kis_calls": kis.calls},
                      fh, ensure_ascii=False, indent=1, default=str)
    return 0 if not bad and not value_bad and not wrong_out else 1


if __name__ == "__main__":
    sys.exit(main())
