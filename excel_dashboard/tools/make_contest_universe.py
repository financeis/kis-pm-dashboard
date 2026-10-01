"""대회 종목군 고정 명단 생성기 — 기준일에 한 번 계산해 버전 관리되는 CSV로 고정합니다.

인코딩 UTF-8 · 작성 2026-10-01 · 출력 파일은 빌더가 통합문서의 정적 표 `tblContest`로 넣습니다.
명단은 자동으로 바뀌지 않습니다(쿼리·빌더는 이 CSV만 읽음). 고정 명단을 지키려고 이미 있는 출력 파일은 덮어쓰지 않습니다 —
재현·감사로 다시 돌릴 때는 --out에 다른 경로를 주고 비교하세요(정말 다시 만들 때만 --force). 나중에 다시 돌리면 그사이
상장폐지된 종목이 마스터에서 빠지는 등 결과가 달라질 수 있습니다.

규칙 (두 조건을 모두 만족, 두 기준 모두 경계값 포함 '이상'):
  1. 시가총액 = 기준일 종가 × 상장주식수 ≥ 시총 하한 (기본 1,000억원 = 100,000,000,000원)
  2. 세션 목록(기본 2026-09-22·23·28·29·30)의 KRX(시장구분 J) 일별 거래대금 평균 ≥ 거래대금 하한 (기본 25억원)
  - 대상: 기존 종목DB(powerquery/T_Universe.pq)와 같은 방식으로 파싱한 KIS 종목 마스터의 개별종목
    (종류 ST·FS·DR) 중 우선주·SPAC 제외. 리츠·ETF·ETN·펀드는 종류 단계에서 이미 빠집니다.
    FS(외국주권)·DR(주식예탁증서)은 포함하고 비고에 '외국기업'을 적습니다.
  - 마스터 상장일 전 세션은 평균에서 빼고(있는 세션만 평균) 비고에 '5일 미만'을 적습니다.
    상장 중인데 거래가 없던 세션(거래정지 등, 거래량 0 봉 또는 봉 없음)은 0원으로 평균에 넣습니다.
  - 판정은 원 단위 정수로 합니다: 시총 = 종가 × 주식수 ≥ 하한, 거래대금 합계 ≥ 하한 × 거래일수 (평균 ≥ 하한과 동치).

데이터 원천과 선택 이유:
  - 후보·종류·우선주·SPAC·상장일: KIS 종목정보 마스터 kospi_code.mst.zip / kosdaq_code.mst.zip (인증 불필요).
    고정폭 위치와 판정(종류 = 그룹코드, 우선주 = 우선주 구분 코드가 공란·0이 아님, SPAC = 'Y')은 T_Universe.pq와 같습니다.
  - 종가·거래대금: 국내주식기간별시세 FHKST03010100 (FID_COND_MRKT_DIV_CODE=J, 일봉, 기간 = 첫 세션 ~ 기준일).
    FID_ORG_ADJ_PRC=1(원주가)을 씁니다. 0(수정주가)은 기준일 뒤에 권리락·액면분할 등이 생기면 기준일 종가까지
    소급 보정하므로, 기준일에 실제 체결된 종가를 얻으려면 원주가여야 합니다. 거래대금(acml_tr_pbmn, 원)은 금액이라 보정과 무관합니다.
  - 상장주식수(기준일): 같은 응답 output1의 lstn_stcn(주 단위 정확값, 조회 시점 기준)에서 '기준일 다음 날 ~ 실행일'에
    상장된 추가 주식을 되돌린 값입니다. 마스터의 상장주수(천 주 단위 절사)와 마스터 시가총액(마스터 기준가 × 실행일 주식수)은
    기준일 주식수를 알려 주지 못하므로 쓰지 않습니다.
    · 기준일 뒤 상장 변동은 예탁원정보(상장정보일정) HHKDB669107C0을 날짜별로 조회합니다(1회 최대 100행, 연속 조회가
      동작하지 않아 하루 단위로 나눔. 하루 100행이면 잘림 위험으로 중단).
    · 주식수를 '더하는' 사유(유상증자·무상증자·주식배당·CB·BW 행사·STOCKOPTION행사·주식전환)는 그 수량을 빼고,
      상호변경은 무시합니다. 그 밖의 사유(액면분할·병합, 자본감소, 통일교체, 처음 보는 사유)가 기준일까지 상장된 조회 대상에
      있거나, 마지막 변동의 총발행주식수가 현재 lstn_stcn과 다르면 추측하지 않고 종료 코드 4로 멈춥니다(사람이 확인).
    · 예: 2026-10-01 실행 때 피노(033790)는 10-01 유상증자 2,691,066주 상장분을 빼 9/30 주식수 83,321,887주로 계산.
  - 일봉 생략: 마스터 시가총액(억, 직전 영업일 기준)이 시총 하한의 절반 미만인 종목은 일봉을 조회하지 않고 탈락시킵니다.
    마스터 시총이 0·공란인 종목(신규상장 등)은 생략하지 않습니다. 그 밖의 후보는 모두 일봉으로 판정합니다.

출력 CSV (UTF-8 BOM, 시가총액 내림차순): 종목코드, 종목명, 시장, 시가총액_0930_억, 평균거래대금_5일_억, 거래일수, 상장주식수, 비고
  - 종목코드는 6자리 텍스트(최근 상장 종목은 '0126Z0'처럼 영문 포함), 시장 = KOSPI/KOSDAQ.
  - 억원 값은 원 단위 값을 소수 둘째 자리에서 반올림(ROUND_HALF_UP). 판정은 반올림 전 원 단위 값으로 했습니다.
  - 비고는 '외국기업', '5일 미만'을 '; '로 잇습니다. 열 이름은 기준일 인자와 관계없이 고정(빌더·쿼리 인터페이스)입니다.

안전장치:
  - 조회 전용: kis_dev.KisClient(GET만, 초당 8회 이하, EGW00201 재시도, 주문·계좌 경로 거부)를 씁니다.
  - 토큰은 원천 통합문서 파일을 Excel 없이 직접 파싱해 메모리로만 씁니다. 남은 유효시간이 210분 미만이면
    KIS를 호출하지 않고 '토큰 갱신 필요'로 멈춥니다(종료 코드 2 — kis_dev와 같음, 인자 오류도 argparse 관례상 2).
    토큰·앱키·시크릿은 출력·파일에 쓰지 않습니다.
  - 일부 종목 조회가 재시도 뒤에도 실패하면 명단을 쓰지 않고 종료 코드 3으로 끝납니다(부분 명단 방지).
  - 기준일 주식수를 자동으로 되돌릴 수 없는 종목이 있으면 명단을 쓰지 않고 종료 코드 4로 끝납니다(위 상장주식수 참고).
  - 출력 파일이 이미 있으면 KIS를 호출하지 않고 종료 코드 5로 끝납니다(--force로만 덮어씀).

사용 예:
    python excel_dashboard/tools/make_contest_universe.py                     # 대회 규칙 기본값(docs/business-rules.md의 '대회 종목 명단', 출력 파일이 없을 때)
    python excel_dashboard/tools/make_contest_universe.py --base-date 20260930 ^
        --sessions 20260922,20260923,20260928,20260929,20260930 --mcap-min-eok 1000 --turnover-min-eok 25 ^
        --out %TEMP%/contest_check.csv --audit %TEMP%/contest_audit.csv       # 재현·감사 (고정 명단과 비교)
    (토큰 원천: --token-source > 환경변수 KIS_TOKEN_SOURCE > excel_dashboard/KIS_PM_Dashboard.xlsm·.xlsx)
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import os
import sys
import time
import zipfile
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Optional

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from kis_dev import KisClient, KisError, TokenError  # noqa: E402

# ---------------------------------------------------------------- 기본값 (대회 규칙: docs/business-rules.md의 '대회 종목 명단')
DEFAULT_BASE_DATE = "20260930"
DEFAULT_SESSIONS = "20260922,20260923,20260928,20260929,20260930"
DEFAULT_MCAP_MIN_EOK = "1000"
DEFAULT_TURNOVER_MIN_EOK = "25"

EOK_WON = Decimal(100_000_000)          # 1억원
MASTER_SKIP_RATIO = Decimal("0.5")      # 마스터 시총이 하한의 이 비율 미만이면 일봉 조회 생략
RETRY_ROUNDS = 3                        # 실패 종목 재시도 회차
MAX_ROWS_PER_CALL = 100                 # KIS 1회 최대 행 수(FHKST03010100·HHKDB669107C0) — 이만큼 오면 잘림 위험

CSV_COLUMNS = ["종목코드", "종목명", "시장", "시가총액_0930_억", "평균거래대금_5일_억", "거래일수", "상장주식수", "비고"]
AUDIT_COLUMNS = ["종목코드", "종목명", "시장", "종류", "상장일", "마스터시총억", "마스터기준가", "마스터상장주식수",
                 "일봉조회", "기준일종가", "종가일자", "상장주식수", "시가총액원", "거래일수", "거래대금합계원",
                 "세션별거래대금", "상장전봉수", "시총충족", "거래대금충족", "편입", "비고", "사유",
                 "현재상장주식수", "주식수보정"]

STOCK_KINDS = ("ST", "FS", "DR")        # 개별종목 (T_Universe와 같음)
FOREIGN_KINDS = ("FS", "DR")            # 비고 '외국기업'
NOTE_FOREIGN = "외국기업"
NOTE_SHORT = "5일 미만"

# ---------------------------------------------------------------- KIS 종목 마스터 (T_Universe.pq와 같은 위치)
MASTER_URL = "https://new.real.download.dws.co.kr/common/master/"
# 각 줄 = 앞부분(단축코드 9 + 표준코드 12 + 한글명) + 뒷부분 고정폭(tail 글자). 위치 = (뒷부분 안의 시작, 길이)
MASTER_LAYOUTS: dict[str, dict] = {
    "KOSPI": {"file": "kospi_code.mst.zip", "tail": 227, "kind": (0, 2), "spac": (29, 1), "base": (41, 9),
              "list_date": (105, 8), "shares": (113, 15), "pref": (158, 1), "mcap": (212, 9)},
    "KOSDAQ": {"file": "kosdaq_code.mst.zip", "tail": 221, "kind": (0, 2), "spac": (24, 1), "base": (36, 9),
               "list_date": (100, 8), "shares": (108, 15), "pref": (153, 1), "mcap": (206, 9)},
}

CHART_PATH = "/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice"
CHART_TR_ID = "FHKST03010100"

# 예탁원정보(상장정보일정): 기준일 뒤 상장 변동을 되돌려 기준일 상장주식수를 구하는 데 씀
LISTING_PATH = "/uapi/domestic-stock/v1/ksdinfo/list-info"
LISTING_TR_ID = "HHKDB669107C0"
ADDITIVE_LISTING_TYPES = frozenset({"유상증자", "무상증자", "주식배당", "국내CB행사", "해외CB행사", "국내BW행사",
                                    "해외BW행사", "STOCKOPTION행사", "주식전환"})   # 기존 주식수에 더해지는 사유
NEUTRAL_LISTING_TYPES = frozenset({"상호변경"})                                      # 주식수 변화 없음


@dataclass
class MasterRow:
    """마스터 한 줄에서 명단 판정에 필요한 값."""
    code: str
    name: str
    market: str
    kind: str
    preferred: bool
    spac: bool
    list_date: Optional[str]            # YYYYMMDD, 형식이 틀리면 None
    mcap_eok: Optional[int]             # 직전 영업일 기준 시가총액(억)
    base_price: Optional[int]           # 마스터 기준가
    shares_master: Optional[int]        # 상장주수(천 주 단위 절사) × 1000


@dataclass
class ListingEvent:
    """예탁원 상장정보일정 한 줄 (기준일 뒤 상장 변동)."""
    list_date: str
    code: str
    issue_type: str
    quantity: int                       # 이번 상장 주식수
    total: int                          # 상장 뒤 총발행주식수


@dataclass
class Judgment:
    """종목 한 개의 판정 결과 (금액은 모두 원 단위 정수)."""
    row: MasterRow
    fetched: bool
    close: Optional[int] = None
    close_date: Optional[str] = None
    shares: Optional[int] = None        # 기준일 상장주식수 (시총 계산에 쓴 값)
    shares_now: Optional[int] = None    # 조회 시점 lstn_stcn
    share_note: str = ""                # 기준일 주식수로 되돌린 내역 (감사용)
    mcap: Optional[int] = None
    session_turnover: dict[str, int] = field(default_factory=dict)   # 상장 중 세션 → 거래대금
    pre_listing_bars: int = 0
    ok_mcap: bool = False
    ok_turnover: bool = False
    reason: str = ""

    @property
    def days(self) -> int:
        return len(self.session_turnover)

    @property
    def turnover_sum(self) -> int:
        return sum(self.session_turnover.values())

    @property
    def included(self) -> bool:
        return self.ok_mcap and self.ok_turnover

    def notes(self, n_sessions: int) -> list[str]:
        out = []
        if self.row.kind in FOREIGN_KINDS:
            out.append(NOTE_FOREIGN)
        if self.fetched and 0 < self.days < n_sessions:      # 기준일까지 미상장(0일)은 명단 밖이라 적지 않음
            out.append(NOTE_SHORT)
        return out


# ---------------------------------------------------------------- 인자
def parse_ymd(text: str) -> str:
    """'YYYYMMDD' 형식 검사 후 그대로 반환."""
    try:
        dt.datetime.strptime(text, "%Y%m%d")
    except ValueError:
        raise argparse.ArgumentTypeError(f"날짜 형식은 YYYYMMDD 입니다: {text}") from None
    return text


def eok_to_won(text: str) -> int:
    """억원 문자열 → 원 단위 정수 (예: '25' → 2,500,000,000). 원 단위로 나누어떨어지지 않거나 0 이하이면 오류."""
    try:
        won = Decimal(text) * EOK_WON
    except InvalidOperation:
        raise argparse.ArgumentTypeError(f"숫자가 아닙니다: {text}") from None
    if not won.is_finite() or won <= 0 or won != won.to_integral_value():
        raise argparse.ArgumentTypeError(f"억원 값은 0보다 크고 원 단위로 나누어떨어져야 합니다: {text}")
    return int(won)


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="대회 종목군 고정 명단 생성 (spec R1, 조회 전용)")
    ap.add_argument("--base-date", type=parse_ymd, default=DEFAULT_BASE_DATE,
                    help="시가총액 기준일 = 세션 목록의 마지막 날 (기본 %(default)s)")
    ap.add_argument("--sessions", default=DEFAULT_SESSIONS,
                    help="거래대금 평균 세션(쉼표 구분 YYYYMMDD, 기본 %(default)s)")
    ap.add_argument("--mcap-min-eok", default=DEFAULT_MCAP_MIN_EOK, help="시가총액 하한(억원, 이상) 기본 %(default)s")
    ap.add_argument("--turnover-min-eok", default=DEFAULT_TURNOVER_MIN_EOK,
                    help="평균 거래대금 하한(억원, 이상) 기본 %(default)s")
    ap.add_argument("--out", help="명단 CSV 경로 (기본 excel_dashboard/data/contest_universe_<기준일>.csv)")
    ap.add_argument("--audit", help="전체 후보 판정 내역 CSV 경로 (선택, 감사·검증용)")
    ap.add_argument("--token-source", help="토큰 원천 통합문서 경로 (기본: 환경변수 KIS_TOKEN_SOURCE 등)")
    ap.add_argument("--force", action="store_true", help="이미 있는 명단 파일을 덮어씀 (고정 명단 — 꼭 필요할 때만)")
    a = ap.parse_args(argv)
    try:
        sessions = sorted({parse_ymd(s.strip()) for s in a.sessions.split(",") if s.strip()})
        a.mcap_min_won = eok_to_won(a.mcap_min_eok)
        a.turnover_min_won = eok_to_won(a.turnover_min_eok)
    except argparse.ArgumentTypeError as e:
        ap.error(str(e))
    if not sessions:
        ap.error("세션 목록이 비었습니다.")
    if sessions[-1] != a.base_date:
        ap.error(f"기준일({a.base_date})은 세션 목록의 마지막 날({sessions[-1]})이어야 합니다.")
    a.session_list = sessions
    a.out = a.out or os.path.join(DASH, "data", f"contest_universe_{a.base_date}.csv")
    return a


# ---------------------------------------------------------------- 마스터
def to_int(text: object) -> Optional[int]:
    """숫자 문자열 → int (공란·형식 오류는 None)."""
    s = str(text if text is not None else "").strip()
    if not s:
        return None
    try:
        return int(Decimal(s))
    except (InvalidOperation, ValueError):
        return None


def download_master_lines(file_name: str) -> list[str]:
    """마스터 zip(파일 1개)을 내려받아 cp949로 풀고 줄 목록을 돌려줌 (줄 끝 CR 제거)."""
    resp = requests.get(MASTER_URL + file_name, timeout=120)
    resp.raise_for_status()
    with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
        data = zf.read(zf.namelist()[0])
    text = data.decode("cp949", errors="replace")
    return [line.rstrip("\r") for line in text.split("\n")]


def parse_master(lines: list[str], market: str) -> list[MasterRow]:
    """T_Universe.pq의 ParseFile과 같은 규칙으로 줄을 해석 (뒷부분 고정폭 + 앞부분 코드·이름)."""
    layout = MASTER_LAYOUTS[market]
    tail = layout["tail"]
    rows: list[MasterRow] = []
    for line in lines:
        if len(line) <= tail + 21:                       # T_Universe: Text.Length(_) > Tail + 21
            continue
        head, body = line[:-tail], line[-tail:]

        def fld(key: str) -> str:
            start, length = layout[key]
            return body[start:start + length].strip()

        list_date = fld("list_date")
        shares_k = to_int(fld("shares"))
        pref = fld("pref")
        rows.append(MasterRow(
            code=head[:9].strip(),
            name=head[21:].strip(),
            market=market,
            kind=fld("kind"),
            preferred=pref not in ("", "0"),
            spac=fld("spac") == "Y",
            list_date=list_date if len(list_date) == 8 and list_date.isdigit() else None,
            mcap_eok=to_int(fld("mcap")),
            base_price=to_int(fld("base")),
            shares_master=None if shares_k is None else shares_k * 1000,
        ))
    return rows


def load_candidates() -> tuple[list[MasterRow], dict[str, int]]:
    """두 시장 마스터 → 후보(종류 ST·FS·DR, 우선주·SPAC 제외)와 시장별 마스터 줄 수."""
    candidates: list[MasterRow] = []
    counts: dict[str, int] = {}
    for market, layout in MASTER_LAYOUTS.items():
        rows = parse_master(download_master_lines(layout["file"]), market)
        counts[market] = len(rows)
        candidates += [r for r in rows if r.kind in STOCK_KINDS and not r.preferred and not r.spac]
    codes = [r.code for r in candidates]
    if len(codes) != len(set(codes)):
        raise RuntimeError("마스터 후보에 중복 종목코드가 있습니다.")
    return candidates, counts


# ---------------------------------------------------------------- 기준일 뒤 상장 변동 (기준일 주식수 복원)
def fetch_listing_events(kis: KisClient, base_date: str, run_date: str) -> dict[str, list[ListingEvent]]:
    """기준일 다음 날 ~ 실행일의 상장 변동을 종목별로 모음 (HHKDB669107C0, 하루씩 조회).

    연속 조회(tr_cont)가 같은 페이지를 되풀이하므로 날짜를 하루 단위로 나눕니다. 응답은 최대 100행이고
    빈 행으로 채워 오므로, 실제 행이 100개면 잘렸을 수 있어 중단합니다.

    Raises:
        KisError: 조회 실패(재시도 3회 뒤) 또는 하루 100행(잘림 위험).
    """
    events: dict[str, list[ListingEvent]] = {}
    day = dt.datetime.strptime(base_date, "%Y%m%d").date() + dt.timedelta(days=1)
    last = dt.datetime.strptime(run_date, "%Y%m%d").date()
    while day <= last:
        ymd = day.strftime("%Y%m%d")
        params = {"SHT_CD": "", "T_DT": ymd, "F_DT": ymd, "CTS": ""}
        for attempt in range(3):
            try:
                body, _cont = kis.get(LISTING_PATH, LISTING_TR_ID, params)
                break
            except KisError:
                if attempt == 2:
                    raise
                time.sleep(2.0)
        rows = [r for r in (body.get("output1") or []) if str(r.get("sht_cd") or "").strip()]
        if len(rows) >= MAX_ROWS_PER_CALL:
            raise KisError(f"{LISTING_TR_ID} {ymd} 상장 변동이 {len(rows)}행 — 잘렸을 수 있어 중단합니다.")
        for r in rows:
            ev = ListingEvent(list_date=str(r.get("list_dt") or "").strip(), code=str(r["sht_cd"]).strip(),
                              issue_type=str(r.get("issue_type") or "").strip(),
                              quantity=to_int(r.get("issue_stk_qty")) or 0, total=to_int(r.get("tot_issue_stk_qty")) or 0)
            events.setdefault(ev.code, []).append(ev)
        day += dt.timedelta(days=1)
    return events


def shares_on_base_date(current: Optional[int], events: list[ListingEvent]) -> tuple[Optional[int], str]:
    """조회 시점 주식수에서 기준일 뒤 상장 변동을 되돌린 기준일 주식수와 내역. 되돌릴 수 없으면 (None, 사유)."""
    if not events or current is None:
        return current, ""
    unknown = sorted({e.issue_type for e in events} - ADDITIVE_LISTING_TYPES - NEUTRAL_LISTING_TYPES)
    if unknown:
        return None, f"확인 필요: 기준일 뒤 주식수 변동 사유 {', '.join(unknown)}"
    latest = max(e.list_date for e in events)
    totals = {e.total for e in events if e.list_date == latest}
    if totals != {current}:
        return None, f"확인 필요: {latest} 총발행주식수 {sorted(totals)} ≠ 현재 상장주식수 {current}"
    added = sum(e.quantity for e in events if e.issue_type in ADDITIVE_LISTING_TYPES)
    detail = ", ".join(f"{e.list_date} {e.issue_type} {e.quantity}" for e in sorted(events, key=lambda e: e.list_date))
    return current - added, f"기준일 뒤 상장 {added}주 제외 ({detail})"


# ---------------------------------------------------------------- 일봉 조회·판정
def fetch_daily(kis: KisClient, code: str, date_from: str, date_to: str) -> tuple[Optional[int], dict[str, dict]]:
    """FHKST03010100 원주가 일봉 → (상장주식수 lstn_stcn, {일자: {close, turnover, volume}}).

    Raises:
        KisError: rt_cd != '0', 응답 행이 잘렸을 가능성(100행 이상), 같은 일자 중복.
    """
    params = {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": code,
              "FID_INPUT_DATE_1": date_from, "FID_INPUT_DATE_2": date_to,
              "FID_PERIOD_DIV_CODE": "D", "FID_ORG_ADJ_PRC": "1"}
    body, _cont = kis.get(CHART_PATH, CHART_TR_ID, params, check=False)
    if str(body.get("rt_cd")) != "0":
        raise KisError(f"{CHART_TR_ID} {code} {body.get('msg_cd')} {body.get('msg1')}")
    out1 = body.get("output1") or {}
    rows = [r for r in (body.get("output2") or []) if str(r.get("stck_bsop_date") or "").strip()]
    if len(rows) >= MAX_ROWS_PER_CALL:
        raise KisError(f"{CHART_TR_ID} {code} 응답이 {len(rows)}행 — 기간이 길어 잘렸을 수 있습니다.")
    bars: dict[str, dict] = {}
    for r in rows:
        day = str(r["stck_bsop_date"]).strip()
        if day in bars:
            raise KisError(f"{CHART_TR_ID} {code} 같은 일자 봉이 중복되었습니다: {day}")
        bars[day] = {"close": to_int(r.get("stck_clpr")), "turnover": to_int(r.get("acml_tr_pbmn")) or 0,
                     "volume": to_int(r.get("acml_vol")) or 0}
    return to_int(out1.get("lstn_stcn")), bars


def judge(row: MasterRow, shares: Optional[int], bars: dict[str, dict], sessions: list[str], base_date: str,
          mcap_min_won: int, turnover_min_won: int) -> Judgment:
    """한 종목 판정. 상장일 전 세션은 제외, 상장 중 봉이 없거나 거래가 없던 세션은 0원."""
    j = Judgment(row=row, fetched=True, shares=shares)
    for s in sessions:
        if row.list_date is not None and s < row.list_date:
            j.pre_listing_bars += 1 if s in bars else 0
            continue
        j.session_turnover[s] = bars[s]["turnover"] if s in bars else 0
    # 기준일 종가: 기준일 봉(거래정지일도 직전 가격의 0거래 봉이 옴). 봉이 없으면 기간 안 직전 봉의 종가.
    priced = sorted(d for d, b in bars.items() if d <= base_date and (b["close"] or 0) > 0)
    if priced:
        j.close_date = base_date if base_date in priced else priced[-1]
        j.close = bars[j.close_date]["close"]
    if j.close and shares:
        j.mcap = j.close * shares
    j.ok_mcap = j.mcap is not None and j.mcap >= mcap_min_won
    j.ok_turnover = j.days > 0 and j.turnover_sum >= turnover_min_won * j.days
    if j.days == 0:
        j.reason = "기준일까지 미상장"
    elif j.mcap is None:
        j.reason = "기준일 종가 또는 상장주식수 없음"
    elif not j.ok_mcap and not j.ok_turnover:
        j.reason = "시총·거래대금 미달"
    elif not j.ok_mcap:
        j.reason = "시총 미달"
    elif not j.ok_turnover:
        j.reason = "거래대금 미달"
    return j


def needs_fetch(row: MasterRow, mcap_min_won: int) -> bool:
    """마스터 시총이 하한의 절반 미만(0·공란 제외)이면 일봉 생략."""
    if row.mcap_eok is None or row.mcap_eok <= 0:
        return True
    return Decimal(row.mcap_eok) * EOK_WON >= Decimal(mcap_min_won) * MASTER_SKIP_RATIO


def run_fetches(kis: KisClient, targets: list[MasterRow], sessions: list[str], base_date: str,
                mcap_min_won: int, turnover_min_won: int,
                events: dict[str, list[ListingEvent]]) -> tuple[dict[str, Judgment], list[str]]:
    """대상 종목을 모두 조회·판정. 실패 종목은 회차를 바꿔 재시도하고, 끝까지 실패한 코드 목록을 함께 반환.

    events: fetch_listing_events 결과. 조회 시점 주식수를 기준일 주식수로 되돌리는 데 씀.
    """
    done: dict[str, Judgment] = {}
    pending = list(targets)
    total = len(targets)
    for round_no in range(1, RETRY_ROUNDS + 2):
        failed: list[MasterRow] = []
        for i, row in enumerate(pending, 1):
            try:
                shares_now, bars = fetch_daily(kis, row.code, sessions[0], base_date)
            except KisError:
                failed.append(row)
                continue
            # 기준일 뒤에 처음 상장한 종목(세션 0일)은 명단 밖이므로 주식수를 되돌리지 않음
            listed_by_base = row.list_date is None or row.list_date <= base_date
            shares, note = (shares_on_base_date(shares_now, events.get(row.code, [])) if listed_by_base
                            else (shares_now, ""))
            j = judge(row, shares, bars, sessions, base_date, mcap_min_won, turnover_min_won)
            j.shares_now, j.share_note = shares_now, note
            if j.days > 0 and j.close and not shares_now:  # 상장 중인데 주식수가 비면 응답 이상 → 재시도
                failed.append(row)
                continue
            done[row.code] = j
            if round_no == 1 and (i % 200 == 0 or i == len(pending)):
                print(f"  일봉 조회 {i}/{total} (누적 호출 {kis.calls}회, 실패 {len(failed)})", flush=True)
        if not failed:
            return done, []
        if round_no <= RETRY_ROUNDS:
            print(f"  재시도 {round_no}회차: {len(failed)}종목", flush=True)
            time.sleep(2.0 * round_no)
        pending = failed
    return done, [r.code for r in pending]


# ---------------------------------------------------------------- 출력
def eok_text(won: Decimal) -> str:
    """원 → 억원 소수 둘째 자리 반올림 문자열."""
    return str((won / EOK_WON).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def write_universe_csv(path: str, judgments: list[Judgment], n_sessions: int) -> int:
    """편입 종목을 시가총액 내림차순으로 UTF-8(BOM) CSV에 씀. 임시 파일에 쓴 뒤 교체. 반환 = 행 수."""
    rows = sorted((j for j in judgments if j.included), key=lambda j: (-(j.mcap or 0), j.row.code))
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(CSV_COLUMNS)
        for j in rows:
            w.writerow([j.row.code, j.row.name, j.row.market,
                        eok_text(Decimal(j.mcap)), eok_text(Decimal(j.turnover_sum) / j.days),
                        j.days, j.shares, "; ".join(j.notes(n_sessions))])
    os.replace(tmp, path)
    return len(rows)


def write_audit_csv(path: str, judgments: list[Judgment], n_sessions: int) -> None:
    """전체 후보(일봉 생략 포함)의 판정 근거를 원 단위로 기록 (감사·검증용)."""
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(AUDIT_COLUMNS)
        for j in sorted(judgments, key=lambda j: (-(j.mcap or 0), -(j.row.mcap_eok or 0), j.row.code)):
            r = j.row
            yn = (lambda b: "Y" if b else "N")
            w.writerow([r.code, r.name, r.market, r.kind, r.list_date or "",
                        "" if r.mcap_eok is None else r.mcap_eok, "" if r.base_price is None else r.base_price,
                        "" if r.shares_master is None else r.shares_master,
                        yn(j.fetched), "" if j.close is None else j.close, j.close_date or "",
                        "" if j.shares is None else j.shares, "" if j.mcap is None else j.mcap,
                        j.days if j.fetched else "", j.turnover_sum if j.fetched else "",
                        ";".join(f"{d}:{v}" for d, v in j.session_turnover.items()),
                        j.pre_listing_bars if j.fetched else "",
                        yn(j.ok_mcap), yn(j.ok_turnover), yn(j.included),
                        "; ".join(j.notes(n_sessions)), j.reason,
                        "" if j.shares_now is None else j.shares_now, j.share_note])


def print_summary(judgments: list[Judgment], n_sessions: int, base_date: str) -> None:
    """편입·탈락 건수와 점검이 필요한 이상 징후 건수만 출력 (종목별 값은 감사 파일)."""
    inc = [j for j in judgments if j.included]
    fetched = [j for j in judgments if j.fetched]
    by_mkt = {m: sum(1 for j in inc if j.row.market == m) for m in MASTER_LAYOUTS}
    print(f"편입 {len(inc)}종목 (KOSPI {by_mkt['KOSPI']}, KOSDAQ {by_mkt['KOSDAQ']}) | "
          f"외국기업 {sum(1 for j in inc if NOTE_FOREIGN in j.notes(n_sessions))}, "
          f"5일 미만 {sum(1 for j in inc if NOTE_SHORT in j.notes(n_sessions))}")
    print(f"탈락(일봉 조회분) {len(fetched) - len(inc)}: 시총만 충족 {sum(1 for j in fetched if j.ok_mcap and not j.ok_turnover)}, "
          f"거래대금만 충족 {sum(1 for j in fetched if j.ok_turnover and not j.ok_mcap)}, "
          f"둘 다 미달 {sum(1 for j in fetched if not j.ok_mcap and not j.ok_turnover and j.days > 0)}, "
          f"미상장 {sum(1 for j in fetched if j.days == 0)}")
    # 점검용 이상 징후 (0이 정상)
    no_base_bar = sum(1 for j in fetched if j.days > 0 and j.close_date and j.close_date != base_date)
    pre_bars = sum(1 for j in fetched if j.pre_listing_bars > 0)
    no_list_date = sum(1 for j in judgments if j.row.list_date is None)
    print(f"점검: 기준일 봉 없음(직전 종가 사용) {no_base_bar}, 상장일 전 봉 존재 {pre_bars}, 상장일 형식 오류 {no_list_date}")
    adjusted = [j for j in fetched if j.days > 0 and j.share_note and j.shares is not None]
    print(f"기준일 주식수 복원(기준일 뒤 상장분 제외): {len(adjusted)}종목 "
          f"(그중 편입 {sum(1 for j in adjusted if j.included)}) — 내역은 감사 파일 '주식수보정'")
    # 마스터 시총(억)은 마스터 기준가 × 실행일 주식수라 기준일 주식수 검증용은 아님 — 큰 차이(가격 기준 변화 등)만 신호로 봄
    diff = [j for j in fetched if j.mcap and j.row.mcap_eok and not j.share_note
            and abs(Decimal(j.mcap) / EOK_WON - j.row.mcap_eok) > max(Decimal(1), Decimal(j.row.mcap_eok) / 200)]
    print(f"점검: 주식수 복원 없이 마스터 시총과 0.5%·1억 넘게 다른 종목 {len(diff)} (감사 파일의 마스터시총억·시가총액원 비교)")


def main(argv: Optional[list[str]] = None) -> int:
    a = parse_args(argv)
    t0 = time.time()
    print(f"기준일 {a.base_date} | 세션 {','.join(a.session_list)} | 시총 ≥ {a.mcap_min_eok}억 | "
          f"평균 거래대금 ≥ {a.turnover_min_eok}억", flush=True)
    if os.path.exists(a.out) and not a.force:
        print(f"중단: 명단 파일이 이미 있습니다(고정 명단은 덮어쓰지 않음): {a.out}\n"
              f"  재현·감사는 --out에 다른 경로를 주고 비교하세요. 정말 다시 만들 때만 --force.")
        return 5
    try:
        kis = KisClient(token_path=a.token_source)        # 남은 시간 210분 미만이면 호출 없이 TokenError
    except TokenError as e:
        print(f"중단: {e}")
        return 2
    candidates, counts = load_candidates()
    print(f"마스터 KOSPI {counts['KOSPI']}줄, KOSDAQ {counts['KOSDAQ']}줄 → 후보(ST·FS·DR, 우선주·SPAC 제외) "
          f"{len(candidates)} (KOSPI {sum(1 for r in candidates if r.market == 'KOSPI')}, "
          f"KOSDAQ {sum(1 for r in candidates if r.market == 'KOSDAQ')})", flush=True)
    targets = [r for r in candidates if needs_fetch(r, a.mcap_min_won)]
    skipped = [Judgment(row=r, fetched=False, reason="마스터 시총 < 하한 절반(일봉 생략)")
               for r in candidates if not needs_fetch(r, a.mcap_min_won)]
    print(f"일봉 조회 {len(targets)}종목, 생략 {len(skipped)}종목 (마스터 시총 < 하한의 절반)", flush=True)
    run_date = dt.date.today().strftime("%Y%m%d")
    try:
        events = fetch_listing_events(kis, a.base_date, run_date)
    except KisError as e:
        print(f"중단: 기준일 뒤 상장 변동을 확인하지 못했습니다 → 명단을 쓰지 않습니다: {e}")
        return 3
    print(f"기준일 뒤 상장 변동({a.base_date} 다음 날 ~ {run_date}): {sum(len(v) for v in events.values())}건, "
          f"{len(events)}종목", flush=True)
    done, failed = run_fetches(kis, targets, a.session_list, a.base_date, a.mcap_min_won, a.turnover_min_won, events)
    if failed:
        print(f"중단: 재시도 뒤에도 조회 실패 {len(failed)}종목 → 명단을 쓰지 않습니다: {', '.join(failed[:20])}")
        return 3
    review = [j for j in done.values() if j.days > 0 and j.shares is None and j.share_note.startswith("확인 필요")]
    if review:
        print(f"중단: 기준일 주식수를 자동으로 되돌릴 수 없는 종목 {len(review)}개 → 명단을 쓰지 않습니다.")
        for j in review[:20]:
            print(f"  {j.row.code} {j.row.name}: {j.share_note}")
        return 4
    judgments = list(done.values()) + skipped
    n_rows = write_universe_csv(a.out, judgments, len(a.session_list))
    if a.audit:
        write_audit_csv(a.audit, judgments, len(a.session_list))
    print_summary(judgments, len(a.session_list), a.base_date)
    print(f"저장: {a.out} ({n_rows}행){' | 감사: ' + a.audit if a.audit else ''}")
    print(f"KIS 호출 {kis.calls}회, 소요 {time.time() - t0:.0f}초")
    return 0


if __name__ == "__main__":
    sys.exit(main())
