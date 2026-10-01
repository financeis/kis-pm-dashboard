# -*- coding: utf-8 -*-
"""대회종목·업종·테마 값 독립 재계산.

[전체]를 돌려 저장한 통합문서의 `tblCompany`(대회종목 표)·`tblSectorKRX`·`tblThemeAgg` 값을, Power Query와 다른 경로 —
**KIS 원천 응답을 파이썬으로 다시 받아 업무 규칙(docs/business-rules.md) 정의대로 계산** — 로 재계산해 칸마다 대조합니다. 통합문서는 Excel로 열지 않고
파일을 직접 읽습니다(`xlsx_tables`). KIS 호출은 조회 전용 `kis_dev.KisClient`(토큰은 원천 통합문서 파일 파싱, 비밀값 비출력).

재계산 항목 (규칙은 docs/business-rules.md)
- 가격: 일봉(수정주가, FHKST03010100)으로 1D·1W·1M·3M·6M·YTD·1Y(세션 1·5·21·63·126·250 전, YTD = 직전 연도 마지막 세션),
  최근 20일 줄무늬(D-19…D0 일간 등락), 60 완료 세션 평균 거래대금(억)·거래량(천 주), 52주(250세션) 고가 대비, 가격기준일·현재가.
  세션 달력 = KOSPI(0001) 일별 날짜(FHPTJ04040000) — docs/business-rules.md의 '용어'.
  시가총액(억) = 현재가 × 상장주식수(주식현재가 시세 FHKST01010100의 lstn_stcn) — 통합문서는 종목 마스터의 상장주식수를 쓰므로 0.5% 허용.
- 수급: 종목별 투자자매매동향(일별, FHPTJ04160001) 외국인·기관계 최근 1·5·21 완료 세션 합(백만원 ÷ 100 = 억원), 시총 대비 비율.
- 실적·후행 밸류: 손익계산서·재무비율(분기, 누적 → 분기 단독), 후행 PER·PBR(최근 4개 분기 TTM EPS·최근 분기 BPS — 공개기준일과
  무관), PER·PBR 변화 3M·6M·YTD·1Y(분자·분모 모두 공시 지연(1~3분기 + 45일, 4분기 + 90일)으로 '그날 알려진' TTM EPS·BPS),
  최근 분기·YoY 3개·실적 상태·ROE.
- 목표주가: 종목투자의견(FHKST663300C0, 날짜 분할로 기간 전체) → 기준일 D·D−1개월·D−3개월 컨센서스(증권사별 6개월 내 마지막 목표가 평균),
  괴리율·증권사 수·목표가 변화 1M·3M·최근 의견.
- KIS 추정: 종목추정실적(HHKST668300C0) EPS ÷ 10·연도는 output4 dt 라벨, Fwd EPS = FY1 × r/12 + FY2 × (12 − r)/12,
  Fwd PER과 변화 1W·1M(스냅샷 tblEstSnap의 [B−3세션, B] 행 → 없으면 추정일 < B일 때 기준일 가중치로 재구성 → 그 밖 공란).
- 신용·공매도·대차: 신용잔고 일별(FHPST04760000)·공매도 일별(FHPST04830000)·대차 일별(HHPST074500C0) → 신용잔고율·1M 변화(%p),
  공매도 거래대금 비중 5 완료 세션 평균, 대차잔고(주수) 1M 변화율, 기준일.
- KRX 업종: 업종별 FHPTJ04040000 → 지수·기간 등락·외국인/기관/개인 순매수 1D·1W·1M(완료 세션, 억원).
- 테마 집계: 통합문서 tblCompany의 대회 행으로 대테마·세부테마별 시가총액 가중 등락·상승 비율·수급 %를 다시 계산.

    python excel_dashboard/tools/verify_recalc.py --workbook <저장한.xlsm> --token-source <원천.xlsm> \\
        [--codes 005930 000660 …] [--n 12] [--sectors 0013 0021 1012] [--json 결과.json] [--csv 대조표.csv]
종료 코드: 0 전부 일치 / 1 불일치 있음 / 2 토큰 갱신 필요 / 3 인자·파일 오류
"""
from __future__ import annotations

import argparse
import calendar
import csv
import datetime as dt
import json
import math
import os
import sys
from collections import defaultdict
from typing import Any, Optional

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
sys.path.insert(0, DASH)
sys.path.insert(0, HERE)
from kis_dev import KisClient, KisError, TokenError  # noqa: E402
from xlsx_tables import Workbook  # noqa: E402

PERIODS = [("1D", 1), ("1W", 5), ("1M", 21), ("3M", 63), ("6M", 126), ("1Y", 250)]
STRIP = [f"D-{i}" if i else "D0" for i in range(19, -1, -1)]
LAG_Q, LAG_Y = 45, 90
TARGET_MONTHS = 6


# ------------------------------------------------------------------------------------------------- 값 도우미
def num(v: Any) -> Optional[float]:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip().replace(",", "")
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def ymd(v: Any) -> Optional[dt.date]:
    s = str(v or "").strip()
    if len(s) == 8 and s.isdigit():
        try:
            return dt.date(int(s[:4]), int(s[4:6]), int(s[6:]))
        except ValueError:
            return None
    return None


def as_date(v: Any) -> Optional[dt.date]:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return ymd(v)


def add_months(d: dt.date, n: int) -> dt.date:
    """달 더하기(없는 날은 그 달 마지막 날 — Power Query Date.AddMonths와 같은 규칙)."""
    m = d.month - 1 + n
    y = d.year + m // 12
    m = m % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def month_end(yyyymm: str) -> dt.date:
    y, m = int(yyyymm[:4]), int(yyyymm[4:6])
    return dt.date(y, m, calendar.monthrange(y, m)[1])


def ratio(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None or b == 0:
        return None
    return a / b - 1


# ------------------------------------------------------------------------------------------------- 원천 조회
class Source:
    """KIS 원천 응답(조회 전용). 같은 요청은 한 번만 보낸다."""

    def __init__(self, kis: KisClient, today: dt.date):
        self.kis = kis
        self.today = today
        self._cache: dict = {}

    def get(self, path: str, tr: str, params: dict) -> dict:
        key = (path, tr, tuple(sorted(params.items())))
        if key not in self._cache:
            body, _ = self.kis.get(path, tr, params)
            self._cache[key] = body
        return self._cache[key]

    def calendar(self) -> list[dt.date]:
        """KOSPI(0001) 일별 날짜(오름차순) — 세션 달력."""
        out: set[dt.date] = set()
        end = self.today
        for _ in range(3):
            d = end.strftime("%Y%m%d")
            b = self.get("/uapi/domestic-stock/v1/quotations/inquire-investor-daily-by-market", "FHPTJ04040000",
                         {"FID_COND_MRKT_DIV_CODE": "U", "FID_INPUT_ISCD": "0001", "FID_INPUT_DATE_1": d,
                          "FID_INPUT_ISCD_1": "KSP", "FID_INPUT_DATE_2": d, "FID_INPUT_ISCD_2": "0001"})
            ds = [ymd(r.get("stck_bsop_date")) for r in b.get("output") or []]
            ds = [x for x in ds if x]
            if not ds:
                break
            out.update(ds)
            if len(out) >= 262:
                break
            end = min(ds) - dt.timedelta(days=1)
        return sorted(out)

    def daily(self, code: str, start: dt.date) -> list[dict]:
        """일봉(수정주가) start ~ 오늘, 오름차순."""
        rows: dict[dt.date, dict] = {}
        end = self.today
        for _ in range(8):
            b = self.get("/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice", "FHKST03010100",
                         {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": code, "FID_INPUT_DATE_1": start.strftime("%Y%m%d"),
                          "FID_INPUT_DATE_2": end.strftime("%Y%m%d"), "FID_PERIOD_DIV_CODE": "D", "FID_ORG_ADJ_PRC": "0"})
            got = [r for r in b.get("output2") or [] if ymd(r.get("stck_bsop_date"))]
            for r in got:
                d = ymd(r["stck_bsop_date"])
                c = num(r.get("stck_clpr"))
                if d and c and d not in rows:
                    rows[d] = {"d": d, "o": num(r.get("stck_oprc")), "h": num(r.get("stck_hgpr")), "l": num(r.get("stck_lwpr")),
                               "c": c, "v": num(r.get("acml_vol")), "amt": num(r.get("acml_tr_pbmn"))}
            if len(got) < 100:
                break
            mn = min(ymd(r["stck_bsop_date"]) for r in got)
            if mn <= start:
                break
            end = mn - dt.timedelta(days=1)
        return [rows[d] for d in sorted(rows) if d >= start]

    def price(self, code: str) -> dict:
        b = self.get("/uapi/domestic-stock/v1/quotations/inquire-price", "FHKST01010100",
                     {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": code})
        return b.get("output") or {}

    def investor(self, code: str, day: dt.date) -> list[dict]:
        b = self.get("/uapi/domestic-stock/v1/quotations/investor-trade-by-stock-daily", "FHPTJ04160001",
                     {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": code, "FID_INPUT_DATE_1": day.strftime("%Y%m%d"),
                      "FID_ORG_ADJ_PRC": "", "FID_ETC_CLS_CODE": ""})
        return b.get("output2") or []

    def fin(self, code: str) -> tuple[list[dict], list[dict]]:
        q = {"FID_DIV_CLS_CODE": "1", "fid_cond_mrkt_div_code": "J", "fid_input_iscd": code}
        a = self.get("/uapi/domestic-stock/v1/finance/income-statement", "FHKST66430200", q)
        b = self.get("/uapi/domestic-stock/v1/finance/financial-ratio", "FHKST66430300", q)
        return a.get("output") or [], b.get("output") or []

    def opinions(self, code: str, start: dt.date, end: dt.date) -> list[dict]:
        """기간 [start, end]의 의견 전체(최신순). 한 응답 100건 제한 → 가장 오래된 날짜로 끝을 당겨 이어 받음."""
        out: list[dict] = []
        d2 = end
        for _ in range(10):
            b = self.get("/uapi/domestic-stock/v1/quotations/invest-opinion", "FHKST663300C0",
                         {"FID_COND_MRKT_DIV_CODE": "J", "FID_COND_SCR_DIV_CODE": "16633", "FID_INPUT_ISCD": code,
                          "FID_INPUT_DATE_1": start.strftime("%Y%m%d"), "FID_INPUT_DATE_2": d2.strftime("%Y%m%d")})
            rows = [r for r in (b.get("output") or []) if ymd(r.get("stck_bsop_date"))]
            if len(rows) < 100:
                out.extend(rows)
                break
            oldest = min(ymd(r["stck_bsop_date"]) for r in rows)
            out.extend(r for r in rows if ymd(r["stck_bsop_date"]) > oldest)
            if oldest <= start or oldest == d2:
                out.extend(r for r in rows if ymd(r["stck_bsop_date"]) == oldest)
                break
            d2 = oldest
        return [r for r in out if start <= ymd(r["stck_bsop_date"]) <= end]

    def estimate(self, code: str) -> dict:
        return self.get("/uapi/domestic-stock/v1/quotations/estimate-perform", "HHKST668300C0", {"SHT_CD": code})

    def csl(self, code: str) -> tuple[list, list, list]:
        t = self.today
        cr = self.get("/uapi/domestic-stock/v1/quotations/daily-credit-balance", "FHPST04760000",
                      {"FID_COND_MRKT_DIV_CODE": "J", "FID_COND_SCR_DIV_CODE": "20476", "FID_INPUT_ISCD": code,
                       "FID_INPUT_DATE_1": t.strftime("%Y%m%d")}).get("output") or []
        sh = self.get("/uapi/domestic-stock/v1/quotations/daily-short-sale", "FHPST04830000",
                      {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": code,
                       "FID_INPUT_DATE_1": (t - dt.timedelta(days=30)).strftime("%Y%m%d"),
                       "FID_INPUT_DATE_2": t.strftime("%Y%m%d")}).get("output2") or []
        ln = self.get("/uapi/domestic-stock/v1/quotations/daily-loan-trans", "HHPST074500C0",
                      {"MRKT_DIV_CLS_CODE": "3", "MKSC_SHRN_ISCD": code,
                       "START_DATE": (t - dt.timedelta(days=60)).strftime("%Y%m%d"),
                       "END_DATE": t.strftime("%Y%m%d"), "CTS": ""}).get("output1") or []
        return cr, sh, ln

    def sector(self, code: str, iscd1: str) -> list[dict]:
        d = self.today.strftime("%Y%m%d")
        b = self.get("/uapi/domestic-stock/v1/quotations/inquire-investor-daily-by-market", "FHPTJ04040000",
                     {"FID_COND_MRKT_DIV_CODE": "U", "FID_INPUT_ISCD": code, "FID_INPUT_DATE_1": d,
                      "FID_INPUT_ISCD_1": iscd1, "FID_INPUT_DATE_2": d, "FID_INPUT_ISCD_2": code})
        return b.get("output") or []


# ------------------------------------------------------------------------------------------------- 재계산 (업무 규칙 정의)
def px_metrics(bars: list[dict], cal: list[dt.date], now: dt.datetime) -> dict:
    """가격 지표. 세션 위치는 달력(KOSPI 날짜) 기준, 봉은 달력 날짜만 쓴다."""
    pos = {d: i for i, d in enumerate(cal)}
    by = {b["d"]: b for b in bars if b["d"] in pos}
    if not by:
        return {}
    p0 = max(pos[d] for d in by)
    d0 = cal[p0]
    cur = by[d0]["c"]

    def close_at(p: int) -> Optional[float]:
        if p < 0:
            return None
        b = by.get(cal[p])
        return b["c"] if b else None

    out: dict[str, Any] = {"가격기준일": d0, "현재가": cur}
    for k, n in PERIODS:
        out[k] = ratio(cur, close_at(p0 - n))
    prev_year = [p for p, d in enumerate(cal) if d.year < d0.year]
    out["YTD"] = ratio(cur, close_at(max(prev_year))) if prev_year else None
    last = len(cal) - 1
    for i, name in enumerate(STRIP):           # D-19 … D0
        q = last - (19 - i)
        out[name] = ratio(close_at(q), close_at(q - 1))
    in_progress = cal[last] == now.date() and now.time() < dt.time(15, 30)
    last_done = last - 1 if in_progress else last
    win = [cal[p] for p in range(max(0, last_done - 59), last_done + 1)]
    amts = [by[d]["amt"] / 1e8 for d in win if d in by and by[d]["amt"] is not None]
    vols = [by[d]["v"] / 1000 for d in win if d in by and by[d]["v"] is not None]
    out["거래대금60일억"] = sum(amts) / len(amts) if amts else None
    out["거래량60일천주"] = sum(vols) / len(vols) if vols else None
    hw = [cal[p] for p in range(max(0, p0 - 249), p0 + 1)]
    highs = [max(x for x in (by[d]["h"], by[d]["c"]) if x is not None) for d in hw if d in by]
    out["고52주대비"] = ratio(cur, max(highs)) if highs else None
    return out


def flow_metrics(rows: list[dict], today: dt.date, now: dt.datetime) -> dict:
    """종목 수급: 완료 세션 순매수 합(억원). 시세가 없는 행(상장 전)·상장일(순매수 0)·장 마감 전 오늘 행은 세지 않는다."""
    recs = []
    seen = set()
    for r in rows:
        d = ymd(r.get("stck_bsop_date"))
        if not d or d in seen:
            continue
        seen.add(d)
        recs.append({"d": d, "px": num(r.get("stck_clpr")), "f": num(r.get("frgn_ntby_tr_pbmn")),
                     "o": num(r.get("orgn_ntby_tr_pbmn")), "p": num(r.get("prsn_ntby_tr_pbmn"))})
    recs.sort(key=lambda x: x["d"], reverse=True)
    done = now.time() >= dt.time(15, 30)
    ses = [x for x in recs if x["px"] and x["px"] > 0 and x["d"] <= today
           and (x["d"] < today or (done and any(x[k] for k in ("f", "o", "p"))))]
    out: dict[str, Any] = {"수급기준일": ses[0]["d"] if ses else None}
    for key, col in (("f", "외국인"), ("o", "기관")):
        for lab, n in (("1D", 1), ("1W", 5), ("1M", 21)):
            v = [x[key] for x in ses[:n]]
            out[f"{col}{lab}억"] = sum(v) / 100 if len(v) == n and None not in v else None
    return out


def standalone(inc: list[dict], rat: list[dict]) -> dict[str, dict]:
    """분기 누적 → 분기 단독(12월 결산 등 회계연도는 누적 매출이 줄어드는 달로 판정)."""
    def blank(v):
        x = num(v)
        return None if x is None or x == 0 or x == 99.99 else x
    q: dict[str, dict] = {}
    for r in inc:
        ym = str(r.get("stac_yymm") or "").strip()
        if len(ym) == 6:
            q.setdefault(ym, {}).update({"sales": blank(r.get("sale_account")), "op": blank(r.get("bsop_prti")),
                                         "ni": blank(r.get("thtr_ntin"))})
    for r in rat:
        ym = str(r.get("stac_yymm") or "").strip()
        if len(ym) == 6:
            q.setdefault(ym, {}).update({"eps": blank(r.get("eps")), "bps": blank(r.get("bps")), "roe": blank(r.get("roe_val"))})
    keys = sorted(q)
    # 1분기 월 판정: 3개월 간격 두 행에서 누적 매출이 줄면 뒤 행의 월이 1분기 월
    votes: dict[int, int] = defaultdict(int)
    for a, b in zip(keys, keys[1:]):
        if month_diff(a, b) == 3:
            sa, sb = q[a].get("sales"), q[b].get("sales")
            if sa is not None and sb is not None and sb < sa:
                votes[int(b[4:])] += 1
    q1 = max(votes, key=lambda m: (votes[m], m)) if votes else (3 if all(int(k[4:]) in (3, 6, 9, 12) for k in keys) else None)
    fy_end = ((q1 - 3 - 1) % 12) + 1 if q1 else None
    out = {}
    for k in keys:
        m = int(k[4:])
        rec = dict(q[k])
        rec["q4"] = fy_end is not None and m == fy_end
        if fy_end is None:
            for f in ("sales", "op", "ni", "eps"):
                rec[f + "_q"] = None
        else:
            first = ((m - q1) % 12) == 0
            prev = shift_ym(k, -3)
            for f in ("sales", "op", "ni", "eps"):
                cum = q[k].get(f)
                if first:
                    rec[f + "_q"] = cum
                else:
                    pc = q.get(prev, {}).get(f)
                    rec[f + "_q"] = cum - pc if cum is not None and pc is not None else None
        rec["pub"] = month_end(k) + dt.timedelta(days=LAG_Y if rec["q4"] else LAG_Q)
        rec["fqn"] = (((m - q1) % 12) // 3 + 1) if q1 else None
        out[k] = rec
    return out


def month_diff(a: str, b: str) -> int:
    return (int(b[:4]) * 12 + int(b[4:])) - (int(a[:4]) * 12 + int(a[4:]))


def shift_ym(k: str, n: int) -> str:
    t = int(k[:4]) * 12 + int(k[4:]) - 1 + n
    return f"{t // 12:04d}{t % 12 + 1:02d}"


def ttm_eps(q: dict[str, dict], d: dt.date) -> Optional[float]:
    """날 d에 알려진(공개기준일 ≤ d) 가장 최근 분기와 그 앞 3개 분기의 단독 EPS 합 — PER 변화(분자·분모)용."""
    known = sorted([k for k, r in q.items() if r["pub"] <= d and any(r.get(f) is not None for f in ("eps", "bps", "sales"))])
    if not known:
        return None
    k0 = known[-1]
    ks = [k0, shift_ym(k0, -3), shift_ym(k0, -6), shift_ym(k0, -9)]
    vals = [q.get(k, {}).get("eps_q") for k in ks]
    if any(v is None for v in vals) or any(q[k]["pub"] > d for k in ks if k in q):
        return None
    return sum(vals)


def bps_at(q: dict[str, dict], d: dt.date) -> Optional[float]:
    """날 d에 알려진 가장 최근 분기의 BPS — PBR 변화(분자·분모)용."""
    known = sorted([k for k, r in q.items() if r["pub"] <= d and any(r.get(f) is not None for f in ("eps", "bps", "sales"))])
    return q[known[-1]].get("bps") if known else None


def latest_quarter(q: dict[str, dict]) -> Optional[str]:
    """값이 하나라도 있는 가장 최근 분기(결산년월) — 실적 저장소의 첫 행과 같은 분기. 없으면 None."""
    filled = sorted(k for k, r in q.items() if any(r.get(f) is not None for f in ("sales", "op", "ni", "eps", "bps", "roe")))
    return filled[-1] if filled else None


def ttm_eps_latest(q: dict[str, dict]) -> Optional[float]:
    """후행 PER용 TTM EPS: 최근 분기와 그 앞 3개 분기(3개월 간격)의 단독 EPS 합, 공개기준일과 무관. 하나라도 없으면 None."""
    k0 = latest_quarter(q)
    if k0 is None:
        return None
    vals = [q.get(k, {}).get("eps_q") for k in (k0, shift_ym(k0, -3), shift_ym(k0, -6), shift_ym(k0, -9))]
    return None if any(v is None for v in vals) else sum(vals)


def bps_latest(q: dict[str, dict]) -> Optional[float]:
    """PBR용 BPS: 최근 분기의 BPS(공개기준일과 무관)."""
    k0 = latest_quarter(q)
    return q[k0].get("bps") if k0 is not None else None


def consensus(ops: list[dict], asof: dt.date, months: int = TARGET_MONTHS) -> tuple[Optional[float], Optional[int]]:
    """증권사별 (asof − months, asof] 마지막 목표가(0·공란 제외) 평균. 응답은 최신순 — 같은 날은 앞선 행이 최근."""
    start = add_months(asof, -months)
    last: dict[str, float] = {}
    best_date: dict[str, dt.date] = {}
    for r in ops:                                  # 최신순으로 처음 본 값이 그 증권사의 마지막 목표가
        d = ymd(r.get("stck_bsop_date"))
        tp = num(r.get("hts_goal_prc"))
        br = str(r.get("mbcr_name") or "").strip()
        if not d or not br or tp is None or tp <= 0 or not (start < d <= asof):
            continue
        if br not in last or d > best_date[br]:
            last[br] = tp
            best_date[br] = d
    if not last:
        return None, None
    return sum(last.values()) / len(last), len(last)


def fwd_eps(fy1: Optional[str], e1: Optional[float], e2: Optional[float], asof: dt.date) -> Optional[float]:
    digits = "".join(ch for ch in str(fy1 or "") if ch.isdigit())
    if len(digits) < 6 or e1 is None:
        return None
    y, m = int(digits[:4]), int(digits[4:6])
    r = (y * 12 + m) - (asof.year * 12 + asof.month) + 1
    if r < 1 or r > 12:
        return None
    return e1 if e2 is None else e1 * r / 12 + e2 * (12 - r) / 12


def est_parse(body: dict, today: dt.date) -> dict:
    o1 = body.get("output1") or {}
    if isinstance(o1, list):
        o1 = o1[0] if o1 else {}
    o3 = [r for r in (body.get("output3") or []) if isinstance(r, dict)]
    o4 = [r for r in (body.get("output4") or []) if isinstance(r, dict)]
    eps_row = o3[1] if len(o3) >= 2 else None
    cands = []
    for i, r in enumerate(o4):
        lab = str(r.get("dt") or "").strip()
        dig = "".join(ch for ch in lab if ch.isdigit())
        if len(dig) < 6:
            continue
        raw = num(eps_row.get(f"data{i + 1}")) if eps_row else None
        cands.append((dig[:6], lab, raw / 10 if raw is not None else None))
    cur = f"{today.year:04d}{today.month:02d}"
    ahead = sorted([c for c in cands if c[0] >= cur])
    out = {"추정일": ymd(o1.get("estdate")), "FY1": None, "FY1_EPS": None, "FY2": None, "FY2_EPS": None}
    if ahead and ahead[0][2] is not None:
        out["FY1"], out["FY1_EPS"] = ahead[0][1], ahead[0][2]
        later = [c for c in ahead if c[0] > ahead[0][0]]
        if later and later[0][2] is not None:
            out["FY2"], out["FY2_EPS"] = later[0][1], later[0][2]
    return out


def csl_metrics(cr: list, sh: list, ln: list, today: dt.date) -> dict:
    """신용·공매도·대차: 신용잔고율(최신)·1M(21세션 전) 변화 %p, 공매도 비중 = 완료 세션 5개 평균(공매도 거래대금 ÷ 거래대금), 대차잔고 주수 1M 변화율."""
    def uniq(rows, key):
        seen, out = set(), []
        for r in rows:
            d = ymd(r.get(key))
            if d and d not in seen:
                seen.add(d)
                out.append((d, r))
        return sorted(out, key=lambda x: x[0], reverse=True)
    c = [(d, num(r.get("whol_loan_rmnd_rate"))) for d, r in uniq(cr, "deal_date")]
    c = [(d, v) for d, v in c if v is not None]
    s = [(d, num(r.get("ssts_tr_pbmn")), num(r.get("acml_tr_pbmn"))) for d, r in uniq(sh, "stck_bsop_date") if d < today]
    s = [(d, a, b) for d, a, b in s if b and b > 0 and a is not None]
    loan = [(d, num(r.get("rmnd_stcn"))) for d, r in uniq(ln, "bsop_date") if d < today]
    loan = [(d, v) for d, v in loan if v is not None]
    out: dict[str, Any] = {}
    out["신용잔고율"] = c[0][1] / 100 if c else None
    out["신용잔고율1M변화"] = (c[0][1] - c[21][1]) / 100 if len(c) >= 22 else None
    out["공매도비중5일"] = sum(a / b for _, a, b in s[:5]) / 5 if len(s) >= 5 else None
    out["대차잔고1M변화율"] = (loan[0][1] / loan[21][1] - 1) if len(loan) >= 22 and loan[21][1] else None
    dates = []
    if out["신용잔고율"] is not None:
        dates.append(c[0][0])
    if out["공매도비중5일"] is not None:
        dates.append(s[0][0])
    if out["대차잔고1M변화율"] is not None:
        dates.append(loan[0][0])
    out["신용기준일"] = min(dates) if dates else None
    return out


def sector_metrics(rows: list[dict], today: dt.date, now: dt.datetime) -> dict:
    seen, recs = set(), []
    for r in rows:
        d = ymd(r.get("stck_bsop_date"))
        if d and d not in seen:
            seen.add(d)
            recs.append({"d": d, "px": num(r.get("bstp_nmix_prpr")), "f": num(r.get("frgn_ntby_tr_pbmn")),
                         "o": num(r.get("orgn_ntby_tr_pbmn")), "p": num(r.get("prsn_ntby_tr_pbmn"))})
    recs.sort(key=lambda x: x["d"], reverse=True)
    px = [x for x in recs if x["px"] and x["px"] > 0 and not (x["d"] == today and now.time() < dt.time(9, 0))]
    fl = [x for x in recs if not (x["d"] == today and now.time() < dt.time(15, 30))]
    if not px:
        return {}
    cur = px[0]
    out: dict[str, Any] = {"기준일": cur["d"], "지수": cur["px"]}
    for k, n in PERIODS:
        out[k] = cur["px"] / px[n]["px"] - 1 if len(px) > n else None
    prev = [x for x in px if x["d"].year < cur["d"].year]
    out["YTD"] = cur["px"] / prev[0]["px"] - 1 if prev else None
    for key, col in (("f", "외국인"), ("o", "기관"), ("p", "개인")):
        for lab, n in (("1D", 1), ("1W", 5), ("1M", 21)):
            v = [x[key] for x in fl[:n]]
            out[f"{col}{lab}억"] = sum(v) / 100 if len(v) == n and None not in v else None
    return out


# ------------------------------------------------------------------------------------------------- 대조
class Report:
    def __init__(self):
        self.rows: list[dict] = []

    def add(self, group: str, key: str, field: str, mine: Any, wb: Any, rel: float = 1e-6, note: str = "") -> bool:
        ok = same(mine, wb, rel)
        self.rows.append({"group": group, "key": key, "field": field, "mine": fmt(mine), "workbook": fmt(wb),
                          "ok": ok, "rel": rel, "note": note})
        return ok

    def summary(self) -> dict:
        by = defaultdict(lambda: [0, 0])
        for r in self.rows:
            by[r["group"]][0] += 1
            by[r["group"]][1] += 0 if r["ok"] else 1
        return {g: {"checked": n, "mismatch": m} for g, (n, m) in by.items()}


def same(a: Any, b: Any, rel: float) -> bool:
    if a in (None, "") and b in (None, ""):
        return True
    if a in (None, "") or b in (None, ""):
        return False
    if isinstance(a, dt.date) or isinstance(b, dt.date):
        return as_date(a) == as_date(b)
    fa, fb = num(a), num(b)
    if fa is not None and fb is not None and not isinstance(a, str):
        return abs(fa - fb) <= max(rel * max(abs(fa), abs(fb)), 1e-9)
    return str(a).strip() == str(b).strip()


def fmt(v: Any) -> Any:
    if isinstance(v, dt.datetime):
        return v.isoformat(sep=" ", timespec="seconds")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, float):
        return round(v, 10)
    return v


# ------------------------------------------------------------------------------------------------- 통합문서 읽기
def read_tables(path: str, names: list[str]) -> dict[str, list[dict]]:
    out = {}
    with Workbook(path) as wb:
        for n in names:
            try:
                cols, rows = wb.read_table(n)
            except KeyError:
                out[n] = []
                continue
            out[n] = [dict(zip(cols, r)) for r in rows]
    return out


def pick_codes(company: list[dict], n: int) -> list[str]:
    """표본: 시총 상위 4 + KIS 추정(Fwd EPS) 있는 종목 3 + KOSDAQ 목표가 있는 종목 2 + 후행 PER 빈 종목 1 + 목표가 빈 종목 1 + 나머지 시총순."""
    rows = [r for r in company if r.get("유니버스") == "대회" and r.get("종목코드")]
    rows.sort(key=lambda r: -(num(r.get("시가총액억")) or 0))
    picked: list[str] = []

    def take(cands, k):
        for r in cands:
            if k <= 0:
                break
            if r["종목코드"] not in picked:
                picked.append(r["종목코드"])
                k -= 1
    take(rows, 4)
    take([r for r in rows if num(r.get("FwdEPS")) is not None and r.get("FY2_EPS") not in (None, "")], 3)
    take([r for r in rows if r.get("시장") == "KOSDAQ" and num(r.get("목표가평균")) is not None], 2)
    take([r for r in rows if num(r.get("후행PER")) is None and num(r.get("현재가")) is not None], 1)
    take([r for r in rows if num(r.get("목표가평균")) is None and num(r.get("현재가")) is not None], 1)
    take(rows[len(rows) // 3:], max(0, n - len(picked)))
    return picked[:max(n, len(picked))]


def main() -> int:
    ap = argparse.ArgumentParser(description="V5 독립 재계산(KIS 원천 → spec 정의) — 토큰 값은 출력하지 않음")
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--token-source", required=True)
    ap.add_argument("--codes", nargs="*", default=None)
    ap.add_argument("--n", type=int, default=12)
    ap.add_argument("--sectors", nargs="*", default=None, help="업종코드(기본: 0013 0021 + KOSDAQ 첫 업종)")
    ap.add_argument("--themes", nargs="*", default=None, help="대테마 이름(기본: 시총 1위 대테마 + 다른 대테마의 세부테마 1개)")
    ap.add_argument("--now", default=None, help="재계산 기준 시각(기본: 통합문서 tblCompany 조회시각)")
    ap.add_argument("--json", default=None)
    ap.add_argument("--csv", default=None)
    a = ap.parse_args()

    T = read_tables(a.workbook, ["tblCompany", "tblEstSnap", "tblSectorKRX", "tblThemeAgg", "tblFin", "tblEst", "tblSessions"])
    company = {r["종목코드"]: r for r in T["tblCompany"] if r.get("종목코드")}
    if not company:
        print("tblCompany가 비어 있습니다.")
        return 3
    stamps = [r.get("조회시각") for r in T["tblCompany"] if isinstance(r.get("조회시각"), dt.datetime)]
    now = dt.datetime.fromisoformat(a.now) if a.now else max(stamps)
    today = now.date()
    try:
        kis = KisClient(token_path=a.token_source)
    except TokenError as e:
        print(f"멈춤: {e}")
        return 2
    src = Source(kis, today)
    rep = Report()
    codes = a.codes or pick_codes(T["tblCompany"], a.n)
    print(f"V5 재계산 기준 시각 {now:%Y-%m-%d %H:%M} · 표본 {len(codes)}종목: {' '.join(codes)}")

    cal = src.calendar()
    sess_wb = sorted(as_date(r.get("일자")) for r in T["tblSessions"] if as_date(r.get("일자")))
    rep.add("세션 달력", "KOSPI", "최근 260세션 날짜 일치", ",".join(d.isoformat() for d in cal[-260:]),
            ",".join(d.isoformat() for d in sess_wb[-260:]))
    snaps = defaultdict(list)
    for r in T["tblEstSnap"]:
        if r.get("종목코드") and as_date(r.get("일자")):
            snaps[r["종목코드"]].append(r)

    for code in codes:
        w = company.get(code)
        if not w:
            rep.add("표본", code, "tblCompany 행", "있음", None)
            continue
        start = cal[max(0, len(cal) - 262)] - dt.timedelta(days=7)
        bars = src.daily(code, start)
        pm = px_metrics(bars, cal, now)
        pinfo = src.price(code)
        shares = num(pinfo.get("lstn_stcn"))
        cur = pm.get("현재가")
        mcap = cur * shares / 1e8 if cur and shares else None
        for f in ["가격기준일", "현재가", "1D", "1W", "1M", "3M", "6M", "YTD", "1Y"] + STRIP + ["거래대금60일억", "거래량60일천주", "고52주대비"]:
            rep.add("가격", code, f, pm.get(f), w.get(f))
        rep.add("가격", code, "시가총액억", mcap, w.get("시가총액억"), rel=5e-3, note="상장주식수: lstn_stcn vs 종목 마스터")

        fl = flow_metrics(src.investor(code, today), today, now)
        wcap = num(w.get("시가총액억"))
        for f in ["수급기준일", "외국인1D억", "외국인1W억", "외국인1M억", "기관1D억", "기관1W억", "기관1M억"]:
            rep.add("수급", code, f, fl.get(f), w.get(f))
        for col in ("외국인", "기관"):
            for lab in ("1D", "1W", "1M"):
                v = fl.get(f"{col}{lab}억")
                rep.add("수급", code, f"{col}{lab}%", v / mcap if v is not None and mcap else None, w.get(f"{col}{lab}%"),
                        rel=5e-3, note="분모 = 시가총액(상장주식수 출처 차이 허용)")

        inc, rat = src.fin(code)
        q = standalone(inc, rat)
        pdate = pm.get("가격기준일") or today
        # 후행 PER·PBR = 최근 분기 기준(공개기준일과 무관), PER·PBR 변화의 분자 = 오늘 알려진 값(공시 지연 규칙)
        ttm_l, bps_l = ttm_eps_latest(q), bps_latest(q)
        ttm0, bps0 = ttm_eps(q, today), bps_at(q, today)
        per = cur / ttm_l if cur and ttm_l and ttm_l > 0 else None
        pbr = cur / bps_l if cur and bps_l and bps_l > 0 else None
        rep.add("실적·밸류", code, "후행PER", per, w.get("후행PER"))
        rep.add("실적·밸류", code, "PBR", pbr, w.get("PBR"))
        pos = {d: i for i, d in enumerate(cal)}
        p0 = pos.get(pdate)
        prev_year = [i for i, d in enumerate(cal) if d.year < pdate.year]
        base_pos = {"3M": p0 - 63 if p0 is not None else None, "6M": p0 - 126 if p0 is not None else None,
                    "YTD": max(prev_year) if prev_year else None, "1Y": p0 - 250 if p0 is not None else None}
        closes = {b["d"]: b["c"] for b in bars}
        for k, bp in base_pos.items():
            bd = cal[bp] if bp is not None and bp >= 0 else None
            bc = closes.get(bd) if bd else None
            te, bb = (ttm_eps(q, bd), bps_at(q, bd)) if bd else (None, None)
            pv = ((cur / ttm0) / (bc / te) - 1) if (cur and ttm0 and ttm0 > 0 and bc and te and te > 0) else None
            bv = ((cur / bps0) / (bc / bb) - 1) if (cur and bps0 and bps0 > 0 and bc and bb and bb > 0) else None
            rep.add("실적·밸류", code, f"PER변화{k}", pv, w.get(f"PER변화{k}"), note=f"기준 {bd}")
            rep.add("실적·밸류", code, f"PBR변화{k}", bv, w.get(f"PBR변화{k}"), note=f"기준 {bd}")
        filled = sorted(k for k, r in q.items() if any(r.get(f) is not None for f in ("sales", "op", "ni", "eps", "bps", "roe")))
        if filled:
            k0 = filled[-1]
            r0 = q[k0]
            ly = q.get(shift_ym(k0, -12), {})
            label = f"{k0[:4]}.Q{r0['fqn']}" if r0.get("fqn") else f"{k0[:4]}.{k0[4:]}"
            rep.add("실적·밸류", code, "최근분기", label, w.get("최근분기"))
            for f, col in (("sales", "매출YoY"), ("op", "영업이익YoY"), ("ni", "순이익YoY")):
                a1, a0 = r0.get(f + "_q"), ly.get(f + "_q")
                rep.add("실적·밸류", code, col, (a1 / a0 - 1) if a1 is not None and a0 is not None and a0 > 0 else None, w.get(col))
            a1, a0 = r0.get("op_q"), ly.get("op_q")
            st = None
            if a1 is not None and a0 is not None:
                st = "흑자전환" if a0 <= 0 < a1 else "적자전환" if a0 > 0 >= a1 else "적자지속" if a1 <= 0 else "흑자지속"
            rep.add("실적·밸류", code, "실적상태", st, w.get("실적상태"))
            rep.add("실적·밸류", code, "ROE", r0.get("roe") / 100 if r0.get("roe") is not None else None, w.get("ROE"))

        D = today
        ops = src.opinions(code, add_months(add_months(D, -3), -TARGET_MONTHS), D)
        avg0, n0 = consensus(ops, D)
        avg1, _ = consensus(ops, add_months(D, -1))
        avg3, _ = consensus(ops, add_months(D, -3))
        rep.add("목표주가", code, "목표가평균", avg0, w.get("목표가평균"))
        rep.add("목표주가", code, "증권사수", n0, w.get("증권사수"))
        rep.add("목표주가", code, "괴리율", ratio(avg0, cur), w.get("괴리율"))
        rep.add("목표주가", code, "목표가변화1M", ratio(avg0, avg1) if avg1 and avg1 > 0 else None, w.get("목표가변화1M"))
        rep.add("목표주가", code, "목표가변화3M", ratio(avg0, avg3) if avg3 and avg3 > 0 else None, w.get("목표가변화3M"))
        top = sorted([(ymd(r["stck_bsop_date"]), -i, r) for i, r in enumerate(ops)], key=lambda x: (x[0], x[1]), reverse=True)
        if top:
            r = top[0][2]
            rep.add("목표주가", code, "최근의견일", ymd(r["stck_bsop_date"]), w.get("최근의견일"))
            rep.add("목표주가", code, "최근증권사", str(r.get("mbcr_name") or "").strip() or None, w.get("최근증권사"))
            rep.add("목표주가", code, "최근의견", str(r.get("invt_opnn") or "").strip() or None, w.get("최근의견"))

        e = est_parse(src.estimate(code), today)
        fe = fwd_eps(e["FY1"], e["FY1_EPS"], e["FY2_EPS"], today)
        fper = cur / fe if cur and fe and fe > 0 else None
        for f in ("추정일", "FY1", "FY1_EPS", "FY2_EPS"):
            rep.add("KIS 추정", code, f, e[f], w.get(f))
        rep.add("KIS 추정", code, "FY2", e["FY2"] if e["FY2_EPS"] is not None else ("FY1만" if e["FY1_EPS"] is not None else None),
                w.get("FY2"))
        rep.add("KIS 추정", code, "FwdEPS", fe, w.get("FwdEPS"))
        rep.add("KIS 추정", code, "FwdPER", fper, w.get("FwdPER"))
        for lab, n in (("1W", 5), ("1M", 21)):
            bp = p0 - n if p0 is not None else None
            bd = cal[bp] if bp is not None and bp >= 0 else None
            base = None
            how = "없음"
            if bd:
                lo = cal[max(0, bp - 3)]
                cand = [s for s in snaps.get(code, []) if lo <= as_date(s["일자"]) <= bd]
                if cand:
                    s = max(cand, key=lambda s: as_date(s["일자"]))
                    base = num(s.get("FwdPER"))
                    how = f"스냅샷 {as_date(s['일자'])}"
                elif e["추정일"] and e["추정일"] < bd:
                    bfe = fwd_eps(e["FY1"], e["FY1_EPS"], e["FY2_EPS"], bd)
                    bc = closes.get(bd)
                    base = bc / bfe if bc and bfe and bfe > 0 else None
                    how = f"재구성 {bd}"
            v = fper / base - 1 if fper and base and base > 0 else None
            rep.add("KIS 추정", code, f"FwdPER변화{lab}", v, w.get(f"FwdPER변화{lab}"), note=how)

        cm = csl_metrics(*src.csl(code), today)
        for f in ("신용잔고율", "신용잔고율1M변화", "공매도비중5일", "대차잔고1M변화율", "신용기준일"):
            rep.add("신용·공매도·대차", code, f, cm.get(f), w.get(f))

    # KRX 업종
    sec_rows = {r["코드"]: r for r in T["tblSectorKRX"] if r.get("코드")}
    sectors = a.sectors or (["0013", "0021"] + [c for c, r in sorted(sec_rows.items()) if r.get("시장") == "KOSDAQ" and r.get("구분") == "업종"][:1])
    for sc in sectors:
        w = sec_rows.get(sc)
        if not w:
            rep.add("KRX 업종", sc, "tblSectorKRX 행", "있음", None)
            continue
        m = sector_metrics(src.sector(sc, "KSQ" if w.get("시장") == "KOSDAQ" else "KSP"), today, now)
        for f in ["기준일", "지수", "1D", "1W", "1M", "3M", "6M", "YTD", "1Y"] + [f"{c}{p}억" for c in ("외국인", "기관", "개인") for p in ("1D", "1W", "1M")]:
            rep.add("KRX 업종", f"{sc} {w.get('업종명')}", f, m.get(f), w.get(f))

    # 테마 집계 — 통합문서 tblCompany 대회 행으로 다시 계산
    rows = [r for r in T["tblCompany"] if r.get("유니버스") == "대회" and r.get("대테마")]
    agg = {(r.get("단계"), r.get("대테마"), r.get("세부테마") or None): r for r in T["tblThemeAgg"]}
    caps = defaultdict(float)
    for r in rows:
        caps[r["대테마"]] += num(r.get("시가총액억")) or 0
    big = sorted(caps, key=lambda k: -caps[k])
    themes = a.themes or big[:1]
    subs = []
    if not a.themes and len(big) > 1:
        sub_caps = defaultdict(float)
        for r in rows:
            if r["대테마"] == big[1] and r.get("세부테마"):
                sub_caps[r["세부테마"]] += num(r.get("시가총액억")) or 0
        if sub_caps:
            subs = [(big[1], max(sub_caps, key=lambda k: sub_caps[k]))]
    groups = [("대테마", t, None) for t in themes] + [("세부테마", t, s) for t, s in subs]
    for lvl, t, s in groups:
        members = [r for r in rows if r["대테마"] == t and (s is None or r.get("세부테마") == s)]
        w = agg.get((lvl, t, s))
        key = f"{t}" + (f" / {s}" if s else "")
        if not w:
            rep.add("테마", key, "tblThemeAgg 행", "있음", None)
            continue
        rep.add("테마", key, "종목수", len(members), w.get("종목수"))
        capv = [num(r.get("시가총액억")) for r in members if num(r.get("시가총액억")) is not None]
        rep.add("테마", key, "시가총액억", sum(capv) if capv else None, w.get("시가총액억"))
        for p in ["1D", "1W", "1M", "3M", "6M", "YTD", "1Y"]:
            pairs = [(num(r.get("시가총액억")), num(r.get(p))) for r in members]
            pairs = [(c, x) for c, x in pairs if c is not None and x is not None]
            tot = sum(c for c, _ in pairs)
            rep.add("테마", key, p, sum(c * x for c, x in pairs) / tot if pairs and tot else None, w.get(p))
        for p in ("1D", "1W"):
            xs = [num(r.get(p)) for r in members if num(r.get(p)) is not None]
            rep.add("테마", key, f"상승비율{p}", (sum(1 for x in xs if x > 0) / len(xs)) if xs else None, w.get(f"상승비율{p}"))
        for col in ("외국인", "기관"):
            for p in ("1D", "1W", "1M"):
                pairs = [(num(r.get("시가총액억")), num(r.get(f"{col}{p}억"))) for r in members]
                pairs = [(c, x) for c, x in pairs if c is not None and x is not None]
                tot = sum(c for c, _ in pairs)
                rep.add("테마", key, f"{col}{p}%", sum(x for _, x in pairs) / tot if pairs and tot else None, w.get(f"{col}{p}%"))

    summ = rep.summary()
    bad = [r for r in rep.rows if not r["ok"]]
    print(f"KIS 호출 {kis.calls}회")
    for g, s in summ.items():
        print(f"  {g}: 대조 {s['checked']}칸, 불일치 {s['mismatch']}")
    for r in bad[:80]:
        print(f"  ✗ [{r['group']}] {r['key']} {r['field']}: 재계산 {r['mine']!r} / 통합문서 {r['workbook']!r} {r['note']}")
    if a.csv:
        with open(a.csv, "w", encoding="utf-8-sig", newline="") as fh:
            wr = csv.DictWriter(fh, fieldnames=list(rep.rows[0].keys()))
            wr.writeheader()
            wr.writerows(rep.rows)
    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump({"now": now.isoformat(), "codes": codes, "sectors": sectors, "themes": [list(g) for g in groups],
                       "summary": summ, "mismatches": bad, "kis_calls": kis.calls}, fh, ensure_ascii=False, indent=1, default=str)
    return 0 if not bad else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KisError as e:
        print(f"KIS 오류: {e}")
        sys.exit(1)
