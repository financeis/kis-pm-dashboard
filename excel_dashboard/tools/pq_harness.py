# -*- coding: utf-8 -*-
"""Power Query 단위 시험 하네스 (2026-10-01) — 전체 빌드 없이 `.pq` 쿼리를 실제 KIS API로 시험합니다.

새(숨김) Excel 인스턴스에 지정한 쿼리와 그 의존 쿼리(`fn*` 등)를 넣고, 정적 입력표와 토큰 표를 만든 뒤
쿼리를 **지정한 순서로** 표에 적재·동기 새로 고침하고 결과(행 수·열·앞부분·`상태` 요약·오류)를 보여 줍니다.
단계 사이에 `tblRunCtl`·`tblSettings`·이름 정의·적재된 표의 셀 값을 바꿀 수 있고, 페이지 빌더(`pages/*.py`)
시험용으로 임시 통합문서를 저장해 열어 둔 채 대역 빌더 객체와 PNG 렌더링을 넘겨줍니다.
`build_dashboard.py`(기본 입력표·대역 빌더 속성), `xl_helpers.py`(COM 헬퍼), `tools/kis_dev.py`(토큰 파싱)를 가져다 씁니다.

안전 규칙 (하드 게이트 — 하네스가 강제함)
--------------------------------------------
- 토큰: 원천 통합문서의 `_sys!tblToken`을 Excel로 열지 않고 파일을 직접 파싱해(`kis_dev.require_token`) 메모리에서
  시험 통합문서의 정적 표 `_sys!tblToken`으로만 복사합니다. 출력·JSON·CSV·로그에는 남은 시간(분)만 나옵니다.
- 시험 통합문서에는 `T_Token` 쿼리를 절대 넣지 않습니다. 요청하거나 의존(참조)하면 Excel을 띄우기 전에 거부합니다(종료 코드 3).
- 시작할 때 원천 토큰의 남은 유효시간이 기준(기본 210분 = 3시간 30분) 미만이면 Excel 기동·API 호출 없이
  "토큰 갱신 필요"로 멈춥니다(종료 코드 2). 갱신은 오케스트레이터만 합니다.
- Excel은 항상 별도 인스턴스(`xl_helpers.new_excel` = DispatchEx)로 띄우고 자기 PID만 종료합니다. 파이썬이 비정상 종료돼도
  작업 개체(Job, KILL_ON_JOB_CLOSE)에 묶은 자기 Excel만 함께 정리됩니다. `CalculateUntilAsyncQueriesDone`은 쓰지 않고
  `QueryTable.Refreshing`을 폴링합니다.
- 결과 파일(저장 통합문서·CSV·PNG·JSON)은 git 저장소 밖(스크래치 폴더)에만 씁니다. 저장소 안 경로는 거부합니다.
- 토큰 표(`tblToken`)는 읽기·덤프·렌더링을 거부하고, 다른 출력에 토큰 문자열이 섞이면 `<토큰 가림>`으로 바꿉니다.
- KIS 호출 한도(앱키당 초당 약 20건)를 병렬 작업이 나눠 쓰므로 표본(명단·관심종목)을 작게 유지하세요.

토큰 원천 경로: 인자(`token_source`/`--token-source`) > 환경변수 `KIS_TOKEN_SOURCE` > 이 폴더 위의 통합문서(`kis_dev` 기본)
> 주 저장소(git 공통 디렉터리 기준)의 `excel_dashboard/KIS_PM_Dashboard.xlsm`·`.xlsx`. 워크트리에는 통합문서가 없으므로
보통 `KIS_TOKEN_SOURCE`를 주 저장소 통합문서의 절대 경로로 지정합니다.

CLI
---
    python excel_dashboard/tools/pq_harness.py --queries T_IndexNow --load T_IndexNow=tblIndexNow --show 5
    python excel_dashboard/tools/pq_harness.py --pq-dir D:/scratch/pq --queries T_PxStore \\
        --static tblContest=D:/scratch/contest30.csv --runctl mode=full --refresh tblPxStore \\
        --runctl mode=quick --refresh tblPxStore --dump tblPxStore=D:/scratch/px.csv --json D:/scratch/run.json
    python excel_dashboard/tools/pq_harness.py --list-queries

  쿼리 준비(순서 무관)
    --queries 이름|경로.pq …   쿼리 추가(의존 쿼리 자동 포함). 경로를 주면 그 파일을 씁니다.
    --m 이름=M식               인라인 쿼리 추가 (예: --m "Q_Rows=fnPageRows()") — 함수 시험용
    --load 쿼리[=표][@시트!셀]   쿼리를 표로 적재. 표 기본값: build_dashboard.LOADS에 있으면 빌더와 같은 이름
                               (예: T_Trades → tblTradeLog), 아니면 T_X → tblX, T_A_X → tblA_X, 그 밖 X → tblX.
                               시트 기본값 = 표 이름, 셀 기본값 = B2. --load가 하나도 없으면 --queries/--m 중
                               이름이 fn으로 시작하지 않는 쿼리를 모두 기본 이름으로 적재합니다.
    --pq-dir 폴더              추가 검색 폴더(반복 가능). 지정 순서대로 기본 폴더(excel_dashboard/powerquery)보다 우선.
  단계(적힌 순서대로 실행 — 새로 고침 사이에 값을 바꿀 때 사용)
    --settings 키=값 …         tblSettings 값 변경·키 추가
    --runctl 키=값 …           tblRunCtl 값 변경·키 추가 (mode/started/now_override)
    --static 표=경로.csv|json[@시트!셀]   정적 표 만들기/교체 (tblContest·tblThemeBase·tblOverride·*Seed 등)
    --name 이름=값             이름 정의 만들기/값 변경 (예: --name 분석코드=005930)
    --set "표[열=값,…].열=값"   적재된 표에서 조건에 맞는 행의 셀 값 수정 (분할 모사 등)
    --refresh 표|쿼리|all      동기 새로 고침. 단계에 --refresh가 하나도 없으면 마지막에 적재한 표 전부를 적재 순서대로.
  값 표기: 빈 값 → 빈칸, 'abc → 텍스트 abc, 0으로 시작하는 숫자(005930) → 텍스트, 정수·소수 → 숫자,
           2026-09-30 → 날짜, 2026-09-30 08:30[:00] → 일시, now → 현재 일시, 그 밖 → 텍스트.
  결과·저장
    --show N (기본 5)  적재 표마다 앞 N행 출력 (0이면 생략)      --dump 표=경로.csv  CSV 저장(UTF-8 BOM, 반복 가능)
    --dump-dir 폴더    적재 표 전부를 <표>.csv로 저장           --json 경로        결과 요약 JSON(기계 판독용)
    --render 시트[!범위]=경로.png  범위를 PNG로 렌더링          --save 경로.xlsx|.xlsm  시험 통합문서 저장
    --keep-open        끝난 뒤 Excel을 보이게 하고 열어 둠(사람이 직접 확인·종료할 때만; 자동화는 Python API 사용)
  기타
    --min-minutes N (기본 210)  --token-source 경로  --no-token(토큰 표 없이 — KIS 호출 불가, 로컬 쿼리·페이지 시험용)
    --cfg 경로(kis_devlp.yaml, 기본 ~/KIS/config/kis_devlp.yaml)  --sample(빌더 샘플 관심종목·매매일지·대회 기간)
    --visible  --timeout 초(표당, 기본 900)  --no-diagnose  --allow-repo-output  --list-queries
  종료 코드: 0 성공 / 1 새로 고침 실패(시간 초과 포함)가 하나 이상 / 2 토큰 갱신 필요·원천 읽기 실패(Excel 미기동) /
            3 설정·인자 오류(쿼리 파일 없음·T_Token 요청·참조는 Excel 기동 전에 거부) / 4 예기치 않은 오류.
  `상태` 열의 `오류:` 행은 데이터 수준 문제라 종료 코드에 반영하지 않고 요약에만 표시합니다.

Python API (스크래치 스크립트에서)
---------------------------------
    import datetime as dt, sys; sys.path.insert(0, r"<워크트리>/excel_dashboard/tools")
    from pq_harness import Harness                          # 가져오면 excel_dashboard도 sys.path에 들어감

    with Harness(pq_dirs=[r"<스크래치>/pq"]) as h:          # 토큰 확인 → Excel → 기본 입력표·토큰 표
        h.static("tblContest", r"<스크래치>/contest30.csv")   # CSV/JSON/rows로 정적 표 만들기·교체
        h.settings(hist_days=60)                              # tblSettings 값 변경·키 추가
        h.runctl(mode="full")                                 # tblRunCtl 값 변경
        h.define_name("분석코드", "005930")                   # 이름 정의(셀 참조) 만들기·값 변경
        h.load("T_PxStore")                                   # → tblPxStore (의존 쿼리 자동 추가)
        r = h.refresh("tblPxStore")                           # RefreshResult(ok, rows, columns, seconds, status, error…)
        print(h.report("tblPxStore", head=10))
        h.set_cells("tblPxStore", {"종목코드": "005930", "일자": dt.date(2026, 9, 30)}, {"종가": 1})
        h.runctl(mode="quick"); h.refresh("tblPxStore")
        data = h.table("tblPxStore")                          # TableData(columns, rows) — records()로 dict 목록
        h.dump("tblPxStore", r"<스크래치>/px.csv")
        h.save(r"<스크래치>/page_test.xlsx")                  # 저장해도 열린 채로 계속 사용
        b = h.builder()                                       # 페이지 빌더 대역: wb·ws[시트]·lo[표]·table_style·say()·
        from pages import page_company; page_company.build(b) #   nav_links()·HEADER_RIGHT·HEADER_STATUS (+ xl_helpers)
        h.render("대회종목", r"<스크래치>/company.png")        # 범위 → PDF → PNG (클립보드 미사용)
    # with 블록이 끝나면 자기 Excel만 종료(keep_open(detach=True)를 부르면 보이게 남김).
    # load()가 돌려준 ListObject 같은 COM 객체를 블록 밖까지 쥐고 있으면 Excel이 스스로 끝나지 않아 8초 뒤 강제 종료됩니다.

  Harness(pq_dirs=None, token_source=None, min_minutes=210, use_token=True, cfg_path=None, sample=False,
          visible=False, timeout=900.0, diagnose=True, allow_repo_output=False, quiet=False)
  - start()/close(force=False) 또는 with 문. start()가 토큰 기준 미달이면 Excel 없이 TokenGateError.
  - add_queries(*이름|경로) / add_query_text(이름, M식) / load(쿼리, table=None, sheet=None, cell="B2")
  - refresh(표|쿼리, timeout=None, sync=False) → RefreshResult / refresh_all(names=None) → list[RefreshResult]
  - static(표, 원천=None, rows=None, columns=None, sheet=None, cell=None, text_cols=(), date_cols=())
  - settings(dict|**키값) / runctl(dict|**키값) / define_name(이름, 값=None, refers_to=None, sheet=None, cell=None)
  - set_name(이름, 값) / set_cells(표, 조건 dict, 값 dict) → 수정 행 수 / delete_rows(표, 조건 dict) → 삭제 행 수
  - table(표, max_rows=None, dates="auto") → TableData / row_count(표) / status_summary(표) / status_errors(표, top=5)
  - report(표, head=5) → 문자열 / dump(표, 경로.csv) / queries() / tables() / default_table(쿼리) / sheet(이름)
  - save(경로) / keep_open(detach=False) / builder(sheets=()) → StandInBuilder / render(시트, png, cell_range=None, dpi=150)
  - summary(head=5) → JSON으로 쓸 수 있는 dict (토큰 없음). 오류: HarnessError(설정), TokenGateError(토큰 기준 미달)

동작 세부
---------
- 쿼리 검색: 추가 폴더(지정 순서) → 기본 폴더. 쿼리 M 본문에서 주석·문자열을 뺀 식별자 중 다른 `.pq` 파일 이름과 같은 것을
  의존 쿼리로 보고 먼저 넣습니다(재귀). `T_Token`이 걸리면 거부합니다.
- 기본 정적 표(각자 자기 이름의 시트 B2, 토큰 표만 `_sys`): tblSettings(`build_dashboard.settings_rows` 기본값),
  tblWatch·tblTrades(기본 빈 표, sample=True면 빌더 샘플), tblMacro, tblHolidays, tblRunCtl(mode=build, started·now_override 빈칸),
  tblContest·tblThemeBase·tblOverride(docs/contracts.md의 열을 갖춘 빈 표). 빈 표는 빌더처럼 빈 행 1개를 가집니다.
  `static()`으로 같은 이름을 주면 교체합니다. `*Seed` 표는 지정할 때만 만듭니다.
- 값 형식: 문자열 → 텍스트(셀 서식 @, 선행 0 유지), 숫자 → 숫자, 날짜·일시 → Excel 일련번호 + 날짜 서식(파워 쿼리에서 datetime).
  CSV는 열 이름이 `코드`로 끝나거나 0으로 시작하는 숫자가 하나라도 있는 열을 텍스트로 두고, 나머지 칸은 위 '값 표기'로 변환합니다.
  JSON은 값 형식을 그대로 쓰되 ISO 날짜 문자열은 날짜로 바꿉니다(텍스트 열 제외).
- 새로 고침: 백그라운드 새로 고침 → `QueryTable.Refreshing` 폴링(0.1초, 시간 초과 시 취소) → `AfterRefresh` 이벤트로
  성공·실패 확인. Excel은 백그라운드 완료를 타이머로 확인하고 그 간격이 1→5→20초로 늘어나므로(실측) 폴링마다 자기 Excel
  창에만 WM_TIMER를 보내 즉시 확인시킵니다(새로 고침 1건 약 0.2초 + 쿼리 실행 시간).
  실패하면(1차 시도가 60초 이내였을 때) 이벤트를 끄고 한 번 동기 재실행해 M 오류 메시지를 얻습니다(API 호출이 반복될 수 있음).
  `sync=True`는 처음부터 동기 새로 고침(시간 초과 없음). 오류가 나도 실행은 멈추지 않고 결과에 기록합니다.
- 결과 읽기: 파워 쿼리 표의 날짜는 Excel에 일련번호로 들어오므로 `table()`·덤프·출력은 이름이 일자/일/시각/일시/날짜로 끝나고
  값이 일련번호 범위인 열을 날짜로 바꿔 보여 줍니다(dates="none"이면 원값). 통합문서 값은 바꾸지 않습니다.
- 대역 빌더의 `ws[시트]`는 없는 시트를 만들어 줍니다. `HEADER_RIGHT`·`HEADER_STATUS`는 tblIndexNow·tblQuote·tblGlobal·
  tblRank·tblPriceHist를 참조하므로, 그 수식을 쓰는 페이지 시험은 해당 표를 적재(또는 정적 표로)해야 입력됩니다.
"""
from __future__ import annotations

import argparse
import csv
import dataclasses
import datetime as dt
import decimal
import gc
import glob
import json
import math
import os
import re
import subprocess
import sys
import time
from typing import Any, Callable, Iterable, Optional, Sequence

HERE: str = os.path.dirname(os.path.abspath(__file__))        # excel_dashboard/tools
DASH: str = os.path.dirname(HERE)                             # excel_dashboard
DEFAULT_PQ_DIR: str = os.path.join(DASH, "powerquery")
for _p in (HERE, DASH):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import pythoncom  # noqa: E402
import pywintypes  # noqa: E402

import kis_dev  # noqa: E402
from xl_helpers import (  # noqa: E402
    C, FONT, NF, XL_SHEET_HIDDEN, XL_SRC_RANGE, XL_YES, add_query, ensure_table_style, excel_pid, excel_serial,
    load_query, new_excel, quit_excel, set_nf, set_palette,
)

# ---------------------------------------------------------------- 상수
TOKEN_QUERY: str = "T_Token"
TOKEN_TABLE: str = "tblToken"
TOKEN_SHEET: str = "_sys"
TOKEN_COLUMNS: list[str] = ["env", "token", "expires", "issued", "key_sig", "status", "checked"]
NAMES_SHEET: str = "_names"
DATA_STYLE: str = "TableStyleLight1"
DEFAULT_CELL: str = "B2"
REDACTED: str = "<토큰 가림>"

# 입력 계약 표(docs/contracts.md) (기본은 빈 표 — 내용은 static()으로 CSV를 넣어 교체)
PLAN_EMPTY_TABLES: dict[str, list[str]] = {
    "tblContest": ["종목코드", "종목명", "시장", "시가총액_0930_억", "평균거래대금_5일_억", "거래일수", "상장주식수", "비고"],
    "tblThemeBase": ["종목코드", "종목명", "대테마", "세부테마", "근거"],
    "tblOverride": ["종목코드", "대회편입", "대테마", "세부테마", "NICS 대분류", "NICS 업종", "NICS 세부", "메모"],
}
RUNCTL_DEFAULT: list[tuple[str, Any]] = [("mode", "build"), ("started", None), ("now_override", None)]

EXIT_OK, EXIT_FAILED, EXIT_TOKEN, EXIT_CONFIG, EXIT_UNEXPECTED = 0, 1, 2, 3, 4

POLL_SECONDS: float = 0.1
QUIT_WAIT_SECONDS: float = 8.0                # Quit 뒤 정상 종료 대기(참조를 놓으면 실측 약 2.5초) — 넘으면 자기 프로세스만 강제 종료
DIAGNOSE_LIMIT_SECONDS: float = 60.0          # 1차 시도가 이보다 오래 걸린 실패는 재실행하지 않음(호출 반복 방지)
_BUSY_HRESULTS: set[int] = {-2147418111, -2147417846}   # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER

# QueryTable 이벤트(RefreshEvents) — 실행 중 형식 정보로 다시 확인하고, 실패하면 이 값 사용
_REFRESH_EVENTS_IID: str = "{0002441B-0000-0000-C000-000000000046}"
_DISPID_BEFORE_REFRESH: int = 1596
_DISPID_AFTER_REFRESH: int = 1597

# Excel 오류 값(Range.Value가 돌려주는 정수) → 표시 문자열
_XL_ERRORS: dict[int, str] = {
    -2146826281: "#DIV/0!", -2146826246: "#N/A", -2146826259: "#NAME?", -2146826288: "#NULL!",
    -2146826252: "#NUM!", -2146826265: "#REF!", -2146826273: "#VALUE!", -2146826238: "#SPILL!",
    -2146826236: "#CALC!", -2146826234: "#BUSY!", -2146826232: "#FIELD!", -2146826230: "#UNKNOWN!",
}

_RE_INT = re.compile(r"[+-]?\d+")
_RE_FLOAT = re.compile(r"[+-]?(\d+\.\d*|\.\d+|\d+)([eE][+-]?\d+)?")
_RE_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
_RE_DATETIME = re.compile(r"(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?")
_RE_DATE_COLUMN = re.compile(r"(일자|일|시각|일시|날짜|Date|date)$")


# ---------------------------------------------------------------- 예외·결과 형식
class HarnessError(RuntimeError):
    """하네스 사용 오류(설정·인자·안전 규칙 위반). CLI 종료 코드 3."""


class TokenGateError(HarnessError):
    """토큰 원천을 읽지 못했거나 남은 유효시간이 기준 미달 — Excel을 띄우지 않고 멈춤. CLI 종료 코드 2."""


@dataclasses.dataclass
class RefreshResult:
    """표 하나의 새로 고침 결과.

    Attributes:
        table: 표 이름 / query: 쿼리 이름 / ok: 새로 고침 성공 여부(쿼리 수준 — `상태` 열의 행 오류와는 별개)
        rows·columns: 새로 고침 뒤 행 수·열 이름 / seconds: 걸린 시간(재실행 포함)
        status: `상태` 열 요약 {"OK": n, "오류": n, "이전 데이터": n, 그 밖 값: n} (열이 없으면 빈 dict)
        error: 실패 사유(M 오류 메시지 등) / timed_out: 시간 초과로 취소됨
        confirmed: AfterRefresh 이벤트로 성공·실패를 확인함 / rerun: 오류 메시지를 얻으려고 동기 재실행함
        warnings: 참고 사항(참조 표 없음 등)
    """

    table: str
    query: str
    ok: bool
    rows: int = 0
    columns: list = dataclasses.field(default_factory=list)
    seconds: float = 0.0
    status: dict = dataclasses.field(default_factory=dict)
    error: Optional[str] = None
    timed_out: bool = False
    confirmed: bool = False
    rerun: bool = False
    warnings: list = dataclasses.field(default_factory=list)

    def to_dict(self) -> dict:
        d = dataclasses.asdict(self)
        d["seconds"] = round(self.seconds, 2)
        return d

    def line(self) -> str:
        """한 줄 요약 (콘솔 출력용)."""
        head = f"{self.table} ← {self.query}: "
        if self.ok:
            body = f"성공 {self.rows}행 × {len(self.columns)}열 ({self.seconds:.1f}s)"
            if self.status:
                body += " | 상태 " + ", ".join(f"{k} {v}" for k, v in self.status.items())
        else:
            body = f"실패 ({self.seconds:.1f}s) — {self.error or '사유 불명'}"
        if self.warnings:
            body += " | 참고: " + "; ".join(self.warnings)
        return head + body


@dataclasses.dataclass
class TableData:
    """표 내용(파이썬 값). rows는 행 목록(열 순서 = columns)."""

    name: str
    columns: list
    rows: list

    def __len__(self) -> int:
        return len(self.rows)

    def records(self) -> list[dict]:
        return [dict(zip(self.columns, r)) for r in self.rows]

    def column(self, name: str) -> list:
        j = self.columns.index(name)
        return [r[j] for r in self.rows]


# ---------------------------------------------------------------- M 코드 분석 (Excel 불필요)
_M_TOKEN = re.compile(
    r'(?P<line_comment>//[^\n]*)'
    r'|(?P<block_comment>/\*.*?\*/)'
    r'|(?P<qident>#"(?:[^"]|"")*")'
    r'|(?P<string>"(?:[^"]|"")*")'
    r'|(?P<ident>(?<![\w.])[^\W\d]\w*(?:\.[^\W\d]\w*)*)',
    re.S,
)


def m_identifiers(m_code: str) -> set[str]:
    """M 본문의 식별자 집합 (주석·문자열 안의 단어는 제외, `#"이름"` 형식 식별자는 포함)."""
    out: set[str] = set()
    for mt in _M_TOKEN.finditer(m_code):
        if mt.group("ident"):
            out.add(mt.group("ident"))
        elif mt.group("qident"):
            out.add(mt.group("qident")[2:-1].replace('""', '"'))
    return out


def m_workbook_refs(m_code: str) -> set[str]:
    """`Excel.CurrentWorkbook(){[Name = "표"]}`처럼 이름으로 읽는 통합문서 표·이름 정의 (주석 안은 제외)."""
    parts: list[str] = []
    pos = 0
    for mt in _M_TOKEN.finditer(m_code):
        if mt.group("line_comment") or mt.group("block_comment"):
            parts.append(m_code[pos:mt.start()])
            parts.append(" ")
            pos = mt.end()
    parts.append(m_code[pos:])
    return set(re.findall(r'Name\s*=\s*"([^"]+)"', "".join(parts)))


def default_table_name(query: str) -> str:
    """쿼리 → 표 이름 관례: T_X → tblX, T_A_X → tblA_X, 그 밖 X → tblX."""
    if query.startswith("T_"):
        return "tbl" + query[2:]
    return "tbl" + query


class QueryCatalog:
    """`.pq` 파일·인라인 쿼리 목록과 의존 관계 해석 (Excel 없이 동작 — 시작 전 검증용).

    Args:
        dirs: 검색 폴더 목록(앞쪽이 우선). 같은 이름의 파일이 여러 폴더에 있으면 앞쪽 폴더의 파일을 씁니다.
    """

    def __init__(self, dirs: Sequence[str]):
        self.dirs: list[str] = [os.path.abspath(d) for d in dirs]
        missing = [d for d in self.dirs if not os.path.isdir(d)]
        if missing:
            raise HarnessError(f"쿼리 검색 폴더가 없습니다: {', '.join(missing)}")
        self.files: dict[str, str] = {}
        for d in reversed(self.dirs):
            for f in sorted(glob.glob(os.path.join(d, "*.pq"))):
                self.files[os.path.splitext(os.path.basename(f))[0]] = f
        self.inline: dict[str, str] = {}
        self._texts: dict[str, str] = {}

    def register(self, item: str) -> str:
        """이름 또는 `.pq` 경로를 받아 쿼리 이름을 돌려줌 (경로면 그 파일을 같은 이름보다 우선 등록)."""
        if item.lower().endswith(".pq") or os.sep in item or "/" in item:
            path = os.path.abspath(item)
            if not os.path.isfile(path):
                raise HarnessError(f"쿼리 파일이 없습니다: {path}")
            name = os.path.splitext(os.path.basename(path))[0]
            self.files[name] = path
            self._texts.pop(name, None)
            return name
        return item

    def add_inline(self, name: str, m_code: str) -> str:
        if not re.fullmatch(r"[^\W\d]\w*", name):
            raise HarnessError(f"인라인 쿼리 이름 형식 오류: {name} (영문·한글·숫자·_ , 숫자로 시작 불가)")
        self.inline[name] = m_code
        self._texts.pop(name, None)
        return name

    def known(self) -> set[str]:
        return set(self.files) | set(self.inline)

    def text(self, name: str) -> str:
        if name in self._texts:
            return self._texts[name]
        if name in self.inline:
            txt = self.inline[name]
        elif name in self.files:
            with open(self.files[name], encoding="utf-8-sig") as fh:
                txt = fh.read()
        else:
            raise HarnessError(f"쿼리를 찾을 수 없습니다: {name} (검색 폴더: {', '.join(self.dirs)})")
        self._texts[name] = txt
        return txt

    def deps(self, name: str) -> set[str]:
        """직접 의존하는 쿼리 이름 (T_Token은 별도로 검사하므로 그대로 포함)."""
        ids = m_identifiers(self.text(name))
        return {i for i in ids if i != name and (i in self.known() or i == TOKEN_QUERY)}

    def resolve(self, names: Iterable[str]) -> list[str]:
        """의존 쿼리를 먼저 오도록 정렬한 전체 목록. T_Token을 요청·참조하면 HarnessError."""
        order: list[str] = []
        state: dict[str, int] = {}          # 1 = 방문 중, 2 = 완료

        def visit(n: str, chain: list[str]) -> None:
            if n == TOKEN_QUERY:
                path = " → ".join(chain + [n])
                raise HarnessError(f"T_Token은 시험 통합문서에 넣지 않습니다(토큰 발급 경로 차단): {path}")
            if state.get(n) == 2:
                return
            if state.get(n) == 1:
                return                      # 상호 참조: 이미 경로에 있음
            state[n] = 1
            for d in sorted(self.deps(n)):
                visit(d, chain + [n])
            state[n] = 2
            order.append(n)

        for n in names:
            visit(n, [])
        return order

    def refs(self, names: Iterable[str]) -> set[str]:
        """쿼리들(의존 포함)이 이름으로 읽는 통합문서 표·이름 정의."""
        out: set[str] = set()
        for n in self.resolve(names):
            out |= m_workbook_refs(self.text(n))
        return out


# ---------------------------------------------------------------- 값 변환
def coerce_text(s: Optional[str], *, cli: bool = True) -> Any:
    """문자열 표기 → 파이썬 값 ('값 표기' 규칙). cli=False(CSV 칸)는 `now`·작은따옴표 규칙을 쓰지 않음."""
    if s is None:
        return None
    t = s.strip()
    if t == "":
        return None
    if cli and t.startswith("'"):
        return t[1:]
    if cli and t.lower() == "now":
        return dt.datetime.now().replace(microsecond=0)
    if len(t) > 1 and t.isdigit() and t.startswith("0"):
        return t                                            # 005930 → 텍스트(선행 0 유지)
    if _RE_INT.fullmatch(t):
        return int(t)
    if _RE_FLOAT.fullmatch(t):
        return float(t)
    m = _RE_DATE.fullmatch(t)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return s
    m = _RE_DATETIME.fullmatch(t)
    if m:
        try:
            return dt.datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4)), int(m.group(5)),
                               int(m.group(6) or 0))
        except ValueError:
            return s
    return s


def _iso_to_date(s: str) -> Any:
    """JSON 문자열 중 ISO 날짜·일시만 날짜로 (그 밖은 그대로)."""
    t = s.strip()
    if _RE_DATE.fullmatch(t) or _RE_DATETIME.fullmatch(t):
        v = coerce_text(t, cli=False)
        if isinstance(v, (dt.date, dt.datetime)):
            return v
    return s


def _is_code_column(name: str) -> bool:
    return str(name).endswith("코드")


def _to_value2(v: Any) -> Any:
    """파이썬 값 → Range.Value2에 넣을 값 (날짜는 일련번호 — pywin32 시간대 변환 회피)."""
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (dt.datetime, dt.date)):
        return excel_serial(v)
    if isinstance(v, decimal.Decimal):
        return float(v)
    if isinstance(v, int):
        return float(v) if abs(v) >= 2 ** 31 else v
    if isinstance(v, float):
        return None if math.isnan(v) else v
    return v if isinstance(v, str) else str(v)


def _kind_of(v: Any) -> str:
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, dt.datetime):
        return "datetime"
    if isinstance(v, dt.date):
        return "date"
    if isinstance(v, (int, float, decimal.Decimal)):
        return "number"
    return "text"


def _column_kind(name: str, values: list, text_cols: Sequence[str], date_cols: Sequence[str]) -> str:
    """열 하나의 쓰기 방식: text / number / bool / date / datetime / mixed / blank."""
    if name in text_cols or _is_code_column(name):
        return "text"
    kinds = {_kind_of(v) for v in values if v is not None}
    if name in date_cols:
        return "datetime" if "datetime" in kinds else "date"
    if not kinds:
        return "blank"
    if kinds == {"date"}:
        return "date"
    if kinds <= {"date", "datetime"}:
        return "datetime"
    if len(kinds) == 1:
        return kinds.pop()
    return "mixed"


def _normalize_rows(columns: list, rows: list, kinds: dict, text_cols: Sequence[str], date_cols: Sequence[str]) -> list:
    """열 방식에 맞게 값 정리 (텍스트 열의 숫자 → 문자열, 날짜 열의 ISO 문자열 → 날짜)."""
    out = []
    for r in rows:
        nr = []
        for j, c in enumerate(columns):
            v = r[j] if j < len(r) else None
            k = kinds[c]
            if v is not None and k == "text" and not isinstance(v, str):
                if isinstance(v, float) and v.is_integer():
                    v = str(int(v))
                elif isinstance(v, (dt.date, dt.datetime)):
                    v = v.isoformat()
                else:
                    v = str(v)
            elif isinstance(v, str) and c in date_cols:
                v = _iso_to_date(v)
            nr.append(v)
        out.append(nr)
    return out


def _py_value(v: Any) -> Any:
    """COM이 돌려준 값 → 파이썬 값 (pywintypes 일시 → naive datetime, 오류 정수 → '#N/A' 등, Decimal → float)."""
    if v is None:
        return None
    if isinstance(v, (dt.datetime, pywintypes.TimeType)):
        # Excel 날짜에는 시간대가 없으므로 벽시계 값 그대로 naive datetime으로 (초 단위 반올림)
        base = dt.datetime(v.year, v.month, v.day, v.hour, v.minute, v.second)
        return base + dt.timedelta(seconds=1) if getattr(v, "microsecond", 0) >= 500000 else base
    if isinstance(v, decimal.Decimal):
        return float(v)
    if isinstance(v, int) and not isinstance(v, bool) and v in _XL_ERRORS:
        return _XL_ERRORS[v]
    return v


def _serial_to_dt(v: float) -> dt.datetime:
    return dt.datetime(1899, 12, 30) + dt.timedelta(milliseconds=round(v * 86400000))


def _auto_date_columns(columns: list, rows: list) -> list[int]:
    """이름이 일자/일/시각/일시/날짜로 끝나고 값이 모두 일련번호 범위(1954~2119년)인 열."""
    found = []
    for j, c in enumerate(columns):
        if not _RE_DATE_COLUMN.search(str(c)):
            continue
        vals = [r[j] for r in rows if r[j] is not None and r[j] != ""]
        if vals and all(isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 <= v <= 80000 for v in vals):
            found.append(j)
    return found


def _fmt(v: Any) -> str:
    """CSV·출력용 문자열."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if math.isnan(v):
            return ""
        if v.is_integer() and abs(v) < 1e15:
            return str(int(v))
        return format(v, ".15g")            # 0.17/100 같은 이진 잡음(0.0017000000000000001) 제거
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, dt.date):
        return v.isoformat()
    return str(v)


def _as_grid(v: Any) -> list[list]:
    """Range.Value 결과를 항상 2차원 목록으로 (셀 하나면 스칼라가 옴)."""
    if not isinstance(v, tuple):
        return [[v]]
    return [list(r) for r in v]


def _com_message(e: BaseException) -> str:
    """COM 오류에서 사람이 읽을 메시지 (Excel이 넘긴 M 오류 문구 우선)."""
    info = getattr(e, "excepinfo", None)
    if info and len(info) > 2 and info[2]:
        return str(info[2])
    args = getattr(e, "args", ())
    if len(args) > 1 and args[1]:
        return f"{args[1]} (hresult {args[0]})"
    return str(e)


def _com_retry(fn: Callable[[], Any], timeout: float = 30.0) -> Any:
    """Excel이 바쁠 때(RPC_E_CALL_REJECTED 등) 잠깐 기다렸다 다시 호출."""
    t0 = time.time()
    while True:
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult in _BUSY_HRESULTS and time.time() - t0 < timeout:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.1)
                continue
            raise


def _inside_git_repo(path: str) -> Optional[str]:
    """경로가 git 작업 트리 안이면 그 루트, 아니면 None (.git 파일·폴더를 위로 찾음)."""
    d = os.path.dirname(os.path.abspath(path))
    while True:
        if os.path.exists(os.path.join(d, ".git")):
            return d
        parent = os.path.dirname(d)
        if parent == d:
            return None
        d = parent


def _main_repo_dashboard() -> Optional[str]:
    """git 공통 디렉터리 기준 주 저장소의 excel_dashboard 폴더 (워크트리에서 토큰 원천을 찾을 때)."""
    try:
        r = subprocess.run(["git", "-C", DASH, "rev-parse", "--path-format=absolute", "--git-common-dir"],
                           capture_output=True, text=True, timeout=15)
    except (OSError, subprocess.SubprocessError):
        return None
    common = r.stdout.strip()
    if r.returncode != 0 or not common:
        return None
    return os.path.join(os.path.dirname(os.path.normpath(common)), "excel_dashboard")


def resolve_token_source(path: Optional[str] = None) -> str:
    """토큰 원천 경로: 인자 > KIS_TOKEN_SOURCE > kis_dev 기본(이 폴더 위) > 주 저장소 excel_dashboard (.xlsm > .xlsx)."""
    if path:
        return os.path.abspath(path)
    if os.environ.get("KIS_TOKEN_SOURCE"):
        return os.path.abspath(os.environ["KIS_TOKEN_SOURCE"])
    try:
        return kis_dev.token_source(None)
    except kis_dev.TokenError:
        pass
    main = _main_repo_dashboard()
    if main:
        for name in ("KIS_PM_Dashboard.xlsm", "KIS_PM_Dashboard.xlsx"):
            p = os.path.join(main, name)
            if os.path.exists(p):
                return p
    raise TokenGateError("토큰 원천 통합문서를 찾을 수 없습니다(KIS_TOKEN_SOURCE 또는 --token-source로 지정하세요).")


def _builder_module():
    """build_dashboard 모듈 (기본 입력표·대역 빌더 속성). 가져오기만 하며 빌드는 실행되지 않음(__main__ 보호)."""
    try:
        import build_dashboard
    except Exception as e:  # noqa: BLE001 — 어떤 가져오기 오류든 사용자에게 사유를 알림
        raise HarnessError(f"build_dashboard.py를 가져오지 못했습니다: {type(e).__name__}: {e}") from None
    return build_dashboard


def _bind_to_job(pid: int):
    """자기 Excel PID를 KILL_ON_JOB_CLOSE 작업 개체에 묶음 → 이 파이썬이 어떻게 끝나도 그 Excel만 함께 종료."""
    try:
        import win32api
        import win32con
        import win32job
        job = win32job.CreateJobObject(None, "")
        info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
        info["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
        h = win32api.OpenProcess(win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE, False, pid)
        try:
            win32job.AssignProcessToJobObject(job, h)
        finally:
            win32api.CloseHandle(h)
        return job
    except Exception:  # noqa: BLE001 — 실패해도 정상 종료 경로(quit_excel)는 동작
        return None


class _QuitDone:
    """이미 Quit을 부른 뒤 xl_helpers.quit_excel의 PID 대기·강제 종료만 쓰기 위한 대역 (Quit 호출은 무시)."""

    def Quit(self) -> None:  # noqa: N802 — Excel.Application.Quit과 같은 이름
        return None


def _open_process(pid: int):
    """자기 Excel 프로세스 핸들(대기·종료용). 핸들을 쥐고 있으면 PID가 다른 프로세스에 재사용되지 않음."""
    try:
        import win32api
        import win32con
        access = win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE | win32con.PROCESS_QUERY_LIMITED_INFORMATION
        return win32api.OpenProcess(access, False, pid)
    except Exception:  # noqa: BLE001 — 실패하면 quit_excel(PID 기반)로 대신함
        return None


def _wait_process(handle, seconds: float) -> bool:
    """프로세스가 끝날 때까지 최대 seconds초 대기. 끝났으면 True."""
    import win32event
    return win32event.WaitForSingleObject(handle, int(seconds * 1000)) == win32event.WAIT_OBJECT_0


def _release_job(job) -> None:
    """작업 개체의 KILL_ON_JOB_CLOSE를 끔 (Excel을 이 파이썬 종료 뒤에도 남길 때)."""
    if job is None:
        return
    import win32job
    info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
    info["BasicLimitInformation"]["LimitFlags"] &= ~win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)


# ---------------------------------------------------------------- 새로 고침 이벤트 수신기
class _RefreshSink:
    """QueryTable.AfterRefresh(Success) 수신 (makepy 없이 연결 지점으로 직접 연결)."""

    _public_methods_: list = []
    _com_interfaces_: list = []
    _dispid_to_func_: dict = {}

    def __init__(self) -> None:
        self.before: int = 0
        self.after: Optional[bool] = None

    def OnBeforeRefresh(self, cancel):  # noqa: N802 — COM 이벤트 이름
        self.before += 1
        return cancel

    def OnAfterRefresh(self, success):  # noqa: N802
        self.after = bool(success)


def _refresh_event_ids(qt) -> tuple[Any, int, int]:
    """RefreshEvents의 IID·DISPID를 형식 정보에서 찾음 (실패하면 알려진 값)."""
    iid = pywintypes.IID(_REFRESH_EVENTS_IID)
    before, after = _DISPID_BEFORE_REFRESH, _DISPID_AFTER_REFRESH
    try:
        tlb, _ = qt._oleobj_.GetTypeInfo().GetContainingTypeLib()
        for i in range(tlb.GetTypeInfoCount()):
            if tlb.GetDocumentation(i)[0] != "RefreshEvents":
                continue
            ti = tlb.GetTypeInfo(i)
            attr = ti.GetTypeAttr()
            iid = attr[0]
            for j in range(attr.cFuncs):
                fd = ti.GetFuncDesc(j)
                nm = ti.GetNames(fd.memid)[0]
                if nm == "BeforeRefresh":
                    before = fd.memid
                elif nm == "AfterRefresh":
                    after = fd.memid
            break
    except pywintypes.com_error:
        pass
    return iid, before, after


# ---------------------------------------------------------------- 대역 빌더
class _SheetDict(dict):
    """ws[시트명] — 없는 시트는 만들어 줌 (페이지 모듈이 자기 시트를 쓰도록)."""

    def __init__(self, harness: "Harness"):
        super().__init__()
        self._h = harness
        for ws in harness.wb.Worksheets:
            self[ws.Name] = ws

    def __missing__(self, key: str):
        ws = self._h.sheet(key)
        self[key] = ws
        self._h.say(f"[대역 빌더] 시트 자동 생성: {key}")
        return ws


class _TableDict(dict):
    """lo[표이름] — 대역 빌더를 만든 뒤 생긴 표도 통합문서에서 찾아 줌 (없으면 KeyError)."""

    def __init__(self, harness: "Harness"):
        super().__init__()
        self._h = harness
        for ws in harness.wb.Worksheets:
            for lo in ws.ListObjects:
                self[lo.Name] = lo

    def __missing__(self, key: str):
        lo = self._h._find_lo(key)
        if lo is None:
            raise KeyError(key)
        self[key] = lo
        return lo


class StandInBuilder:
    """페이지 빌더 대역 — `build(builder)`에 넘기는 최소 객체 (pages/AGENTS.md의 페이지 모듈 계약).

    속성: wb, xl, pid, ws(시트 dict, 없는 시트 자동 생성), lo(표 dict, 나중에 생긴 표도 찾음), table_style,
          HEADER_RIGHT, HEADER_STATUS, log, out(저장 경로)
    메서드: say(msg), nav_links(ws, row, start_col=2) — build_dashboard.Builder.nav_links를 그대로 사용.
    """

    def __init__(self, harness: "Harness"):
        bd = _builder_module()
        self.harness = harness
        self.xl = harness.xl
        self.wb = harness.wb
        self.pid = harness.pid
        self.out = harness.saved_path
        self.table_style = harness.table_style
        self.ws = _SheetDict(harness)
        self.lo = _TableDict(harness)
        self.log: list[str] = []
        self.HEADER_RIGHT = bd.Builder.HEADER_RIGHT
        self.HEADER_STATUS = bd.Builder.HEADER_STATUS
        self._nav = bd.Builder.nav_links

    def say(self, msg: str) -> None:
        print(msg, flush=True)
        self.log.append(msg)

    def nav_links(self, ws, row: int, start_col: int = 2) -> None:
        self._nav(self, ws, row, start_col)

    def release(self) -> None:
        """COM 참조를 놓음 (하네스 close()가 부름 — 이 객체를 계속 쥐고 있어도 Excel이 끝날 수 있게)."""
        self.ws.clear()
        self.lo.clear()
        self.wb = None
        self.xl = None


# ---------------------------------------------------------------- 하네스
class Harness:
    """PQ 단위 시험 하네스. 자세한 사용법은 모듈 설명 참고.

    Args:
        pq_dirs: 추가 쿼리 검색 폴더(앞쪽 우선, 기본 폴더 excel_dashboard/powerquery보다 우선).
        token_source: 토큰 원천 통합문서 경로(기본: resolve_token_source 규칙).
        min_minutes: 시작할 때 필요한 토큰 남은 시간(분, 기본 210 = 3시간 30분).
        use_token: False면 토큰 확인·토큰 표를 생략(KIS 호출 불가 — 로컬 쿼리·페이지 시험용).
        cfg_path: tblSettings의 cfg_path 값(기본 ~/KIS/config/kis_devlp.yaml).
        sample: True면 빌더 샘플(관심종목·매매일지·대회 기간)로 기본 입력표를 채움(기본 빈 표 — 호출 최소화).
        visible: Excel 창 표시. timeout: 표당 새로 고침 시간 제한(초).
        diagnose: 백그라운드 새로 고침이 실패하면 동기 재실행으로 오류 메시지를 얻음.
        allow_repo_output: True면 git 저장소 안 경로에도 결과 파일을 씀(기본 거부).
        quiet: True면 진행 메시지를 출력하지 않음(log에는 남음).
    """

    def __init__(self, *, pq_dirs: Optional[Sequence[str]] = None, token_source: Optional[str] = None,
                 min_minutes: int = kis_dev.MIN_MINUTES, use_token: bool = True, cfg_path: Optional[str] = None,
                 sample: bool = False, visible: bool = False, timeout: float = 900.0, diagnose: bool = True,
                 allow_repo_output: bool = False, quiet: bool = False):
        self.catalog = QueryCatalog(list(pq_dirs or []) + [DEFAULT_PQ_DIR])
        self.token_source_arg = token_source
        self.min_minutes = int(min_minutes)
        self.use_token = use_token
        self.cfg_path = cfg_path or kis_dev.DEFAULT_CFG
        self.sample = sample
        self.visible = visible
        self.timeout = float(timeout)
        self.diagnose = diagnose
        self.allow_repo_output = allow_repo_output
        self.quiet = quiet
        self.xl = None
        self.wb = None
        self.pid: Optional[int] = None
        self.table_style: str = "TableStyleLight9"
        self.excel_started: bool = False
        self.saved_path: Optional[str] = None
        self.token_info: dict = {"source": None, "minutes_left": None, "min_minutes": self.min_minutes,
                                 "copied": False, "used": use_token}
        self.timings: dict[str, float] = {}
        self.steps: list[dict] = []
        self.results: list[RefreshResult] = []
        self.log: list[str] = []
        self.loads: list[str] = []                 # 쿼리로 적재한 표 (적재 순서)
        self.query_of: dict[str, str] = {}         # 표 → 쿼리
        self.static_tables: list[str] = []
        self._token: Optional[str] = None          # 메모리 전용 (repr·출력 금지)
        self._token_expires: Optional[dt.datetime] = None
        self._hwnd: int = 0
        self._proc = None
        self._job = None
        self._builders: list = []
        self._keep: bool = False
        self._detached: bool = False
        self._names_row: int = 2
        self._t0: float = time.time()
        self._event_ids = None

    def __repr__(self) -> str:
        return f"<Harness pid={self.pid} tables={len(self.loads)} started={self.excel_started}>"

    # ------------------------------------------------------------ 진행 메시지
    def say(self, msg: str) -> None:
        msg = self._redact_text(msg)
        self.log.append(msg)
        if not self.quiet:
            print(msg, flush=True)

    def _redact_text(self, s: str) -> str:
        if self._token and isinstance(s, str) and self._token in s:
            return s.replace(self._token, REDACTED)
        return s

    def _redact(self, obj: Any) -> Any:
        if isinstance(obj, str):
            return self._redact_text(obj)
        if isinstance(obj, list):
            return [self._redact(x) for x in obj]
        if isinstance(obj, tuple):
            return tuple(self._redact(x) for x in obj)
        if isinstance(obj, dict):
            return {k: self._redact(v) for k, v in obj.items()}
        return obj

    def _step(self, kind: str, **info) -> None:
        self.steps.append(self._redact({"no": len(self.steps) + 1, "kind": kind, **info}))

    # ------------------------------------------------------------ 시작·종료
    def __enter__(self) -> "Harness":
        return self.start()

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def check_token(self) -> None:
        """토큰 원천의 남은 시간 확인 + 메모리로 읽기. 기준 미달이면 TokenGateError (Excel 기동 전)."""
        src = resolve_token_source(self.token_source_arg)
        self.token_info["source"] = src
        try:
            self.token_info["minutes_left"] = kis_dev.remaining_minutes(src)
        except kis_dev.TokenError as e:
            raise TokenGateError(str(e)) from None
        try:
            tok, exp = kis_dev.require_token(src, self.min_minutes)
        except kis_dev.TokenError as e:
            raise TokenGateError(str(e)) from None
        self._token, self._token_expires = tok, exp
        self.say(f"[하네스] 토큰 원천: {src} | 남은 시간: {self.token_info['minutes_left']}분 (기준 {self.min_minutes}분)")

    def start(self) -> "Harness":
        """토큰 확인 → Excel 기동 → 통합문서·기본 정적 표·토큰 표. 기준 미달이면 Excel 없이 TokenGateError."""
        if self.excel_started:
            return self
        self._t0 = time.time()
        if self.use_token:
            self.check_token()
            if not os.path.isfile(self.cfg_path):
                self.say(f"[하네스] 경고: KIS 설정 파일이 없습니다(tblSettings cfg_path): {self.cfg_path} — KIS 호출 쿼리는 설정 오류가 납니다.")
        else:
            self.say("[하네스] 토큰 없이 시작(--no-token): KIS를 호출하는 쿼리는 토큰 오류가 납니다.")
        t = time.time()
        self.xl = new_excel(self.visible)
        self.excel_started = True
        self.pid = excel_pid(self.xl)
        self._hwnd = int(self.xl.Hwnd)
        self._proc = _open_process(self.pid) if self.pid else None
        self._job = _bind_to_job(self.pid) if self.pid else None
        if self.visible:
            self.xl.ScreenUpdating = True
        self.timings["excel_start"] = round(time.time() - t, 2)
        self.say(f"[하네스] Excel 시작 (PID {self.pid}, {self.timings['excel_start']:.1f}s)"
                 + ("" if self._job else " — 작업 개체 연결 실패: 비정상 종료 시 정리 보장 안 됨"))
        try:
            t = time.time()
            self._setup_workbook()
            self.timings["workbook_setup"] = round(time.time() - t, 2)
            self.say(f"[하네스] 통합문서 준비: 정적 표 {len(self.static_tables)}개"
                     + (" + 토큰 표" if self.token_info["copied"] else "") + f" ({self.timings['workbook_setup']:.1f}s)")
        except BaseException:
            self.close(force=True)
            raise
        return self

    def close(self, force: bool = False) -> None:
        """자기 Excel 종료. keep_open() 뒤에는 force=True일 때만 종료."""
        if self.xl is None:
            return
        if self._keep and not force:
            self.say(f"[하네스] Excel을 열어 둡니다 (PID {self.pid})"
                     + (" — 이 파이썬이 끝나도 남습니다. 직접 닫으세요." if self._detached else ""))
            return
        t = time.time()
        pid = self.pid
        try:
            if self.wb is not None:
                self.wb.Close(False)
        except pywintypes.com_error:
            pass
        try:
            self.xl.Quit()
        except pywintypes.com_error:
            pass
        # 이 하네스·대역 빌더가 쥔 COM 참조를 모두 놓아야 Quit 뒤 Excel이 스스로 끝남
        # (참조를 쥔 채 기다리면 끝나지 않아 xl_helpers.quit_excel처럼 대기 시간 뒤 강제 종료하게 됨)
        for b in self._builders:
            b.release()
        self._builders = []
        self.wb = None
        self.xl = None
        gc.collect()
        how = "정상 종료"
        if self._proc is not None:
            if not _wait_process(self._proc, QUIT_WAIT_SECONDS):
                import win32api
                win32api.TerminateProcess(self._proc, 1)      # 자기 프로세스 핸들로만 종료
                _wait_process(self._proc, 5.0)
                how = "강제 종료(호출자가 쥔 COM 참조 등)"
            self._proc = None
        elif pid:
            quit_excel(_QuitDone(), pid, wait=QUIT_WAIT_SECONDS)   # 핸들이 없으면 PID 기준(자기 PID만)
            how = "PID 기준 종료"
        if self._job is not None:
            try:                                      # 자기 Excel이 띄운 자식(Mashup 컨테이너 등)만 남았으면 정리
                import win32job
                win32job.TerminateJobObject(self._job, 0)
            except Exception:  # noqa: BLE001 — 이미 모두 끝났으면 실패해도 무방
                pass
            self._job = None
        self.timings["excel_quit"] = round(time.time() - t, 2)
        self.timings["total"] = round(time.time() - self._t0, 2)
        self.say(f"[하네스] Excel 종료 (PID {pid}, {how} {self.timings['excel_quit']:.1f}s) | 총 {self.timings['total']:.1f}s")

    def keep_open(self, detach: bool = False) -> None:
        """close()에서 Excel을 닫지 않음. detach=True면 보이게 하고 이 파이썬 종료 뒤에도 남김(사람이 확인할 때)."""
        self._keep = True
        if detach and self.xl is not None:
            _release_job(self._job)
            self._detached = True
            self.xl.Visible = True
            self.xl.ScreenUpdating = True
            self.xl.DisplayAlerts = True
            self.xl.EnableEvents = True

    def _require(self) -> None:
        if self.wb is None:
            raise HarnessError("하네스가 시작되지 않았습니다(start() 또는 with 문).")

    # ------------------------------------------------------------ 통합문서·기본 표
    def _setup_workbook(self) -> None:
        bd = _builder_module()
        wb = self.xl.Workbooks.Add()
        self.wb = wb
        try:
            wb.EnableAutoRecover = False                 # 강제 종료돼도 사용자 Excel에 '복구된 문서'가 남지 않게(이 통합문서만)
        except pywintypes.com_error:
            pass
        set_palette(wb, 10, C["up"])                     # 숫자서식 [Color10] = 상승 빨강
        set_palette(wb, 11, C["down"])                   # [Color11] = 하락 파랑
        wb.Queries.FastCombine = True                    # 개인정보 수준 무시 (Formula.Firewall 방지, 빌더와 같음)
        self.table_style = "PM Navy" if ensure_table_style(wb) is not None else "TableStyleLight9"
        st = wb.Styles("Normal")
        st.Font.Name = FONT
        st.Font.Size = 10
        first = wb.Worksheets(1)
        first.Name = TOKEN_SHEET
        settings = [list(r) for r in bd.settings_rows(self.cfg_path, self.sample)]
        self._write_new_table("tblSettings", ["키", "항목", "값", "설명"], settings)
        watch = [[c, None, g, a, t, p] for (c, g, a, t, p) in bd.SAMPLE_WATCH] if self.sample else []
        self._write_new_table("tblWatch", ["종목코드", "종목명", "그룹", "관심가", "목표가", "투자포인트"], watch)
        self._write_new_table("tblTrades", list(bd.TRADE_HEADERS), bd.sample_trades() if self.sample else [])
        self._write_new_table("tblMacro", ["구분", "이름", "시장코드", "심볼", "순서"], [list(r) for r in bd.MACRO_ROWS])
        self._write_new_table("tblHolidays", ["휴장일", "설명"], [list(r) for r in bd.HOLIDAYS])
        self._write_new_table("tblRunCtl", ["키", "값"], [list(r) for r in RUNCTL_DEFAULT])
        for name, cols in PLAN_EMPTY_TABLES.items():
            self._write_new_table(name, cols, [])
        if self.use_token and self._token:
            self._write_token_table()

    def _write_token_table(self) -> None:
        """토큰 표 `_sys!tblToken` (T_Token 결과와 같은 열). 값은 메모리의 토큰만 — 어디에도 출력하지 않음.

        expires는 Excel 일련번호 + 일시 서식으로 써서 fnToken의 DateTime.From이 datetime으로 읽습니다.
        issued·key_sig는 시험 통합문서에 T_Token이 없으므로 비워 둡니다.
        """
        now = dt.datetime.now().replace(microsecond=0)
        row = ["prod", self._token, self._token_expires, None, None, "하네스 복사(원천 파일 파싱)", now]
        ws = self.sheet(TOKEN_SHEET)
        self._write_table_at(ws, 2, 2, TOKEN_TABLE, TOKEN_COLUMNS, [row],
                             {"env": "text", "token": "text", "expires": "datetime", "issued": "datetime",
                              "key_sig": "text", "status": "text", "checked": "datetime"})
        ws.Visible = XL_SHEET_HIDDEN
        self.token_info["copied"] = True

    def sheet(self, name: str):
        """시트 가져오기(없으면 맨 뒤에 만듦)."""
        self._require()
        for ws in self.wb.Worksheets:
            if ws.Name == name:
                return ws
        if not name or len(name) > 31 or re.search(r"[\[\]:*?/\\]", name):
            raise HarnessError(f"시트 이름 형식 오류: {name!r} (31자 이하, []:*?/\\ 불가)")
        # 늦은 바인딩에서 After= 키워드는 무시되고(활성 시트 앞에 들어감) (None, 마지막 시트) 위치 인수만 동작(실측)
        ws = self.wb.Worksheets.Add(None, self.wb.Worksheets(self.wb.Worksheets.Count))
        ws.Name = name
        return ws

    def _find_lo(self, name: str):
        self._require()
        for ws in self.wb.Worksheets:
            for lo in ws.ListObjects:
                if lo.Name.lower() == name.lower():
                    return lo
        return None

    def _lo(self, name: str):
        lo = self._find_lo(name)
        if lo is None:
            raise HarnessError(f"표를 찾을 수 없습니다: {name} (있는 표: {', '.join(self.tables())})")
        return lo

    # ------------------------------------------------------------ 표 쓰기 (정적 표)
    def _write_new_table(self, name: str, columns: list, rows: list, *, sheet: Optional[str] = None,
                         cell: Optional[str] = None, text_cols: Sequence[str] = (), date_cols: Sequence[str] = ()):
        kinds = {c: _column_kind(c, [r[j] if j < len(r) else None for r in rows], text_cols, date_cols)
                 for j, c in enumerate(columns)}
        rows = _normalize_rows(columns, rows, kinds, text_cols, date_cols)
        ws = self.sheet(sheet or name)
        rng = ws.Range(cell or DEFAULT_CELL)
        lo = self._write_table_at(ws, rng.Row, rng.Column, name, columns, rows, kinds)
        if name not in self.static_tables:
            self.static_tables.append(name)
        return lo

    def _write_table_at(self, ws, top: int, left: int, name: str, columns: list, rows: list, kinds: dict):
        """머리글 + 본문을 한 번에(Value2 덩어리) 쓰고 ListObject로 만듦. 빈 표는 빈 행 1개."""
        if not columns:
            raise HarnessError(f"열이 없는 표는 만들 수 없습니다: {name}")
        if len(set(columns)) != len(columns):
            raise HarnessError(f"열 이름이 중복됩니다: {name} {columns}")
        ncol = len(columns)
        nrow = max(1, len(rows))
        cells = ws.Cells
        hdr = ws.Range(cells(top, left), cells(top, left + ncol - 1))
        hdr.NumberFormat = "@"
        hdr.Value2 = (tuple(str(c) for c in columns),)
        for j, c in enumerate(columns):
            col = ws.Range(cells(top + 1, left + j), cells(top + nrow, left + j))
            if kinds[c] == "text":
                col.NumberFormat = "@"
            elif col.NumberFormat in (None, "@"):      # None = 칸마다 서식이 다름
                set_nf(col, "General")                 # 교체 전 표가 남긴 텍스트 서식 제거(숫자가 문자로 저장되지 않게)
        mixed_cells: list[tuple[int, int, Any]] = []
        grid: list[tuple] = []
        for i, r in enumerate(rows):
            line = []
            for j, c in enumerate(columns):
                v = r[j] if j < len(r) else None
                if kinds[c] == "mixed" and isinstance(v, (str, dt.date, dt.datetime)):
                    mixed_cells.append((i, j, v))      # 문자열·날짜는 칸별 서식이 필요해 따로 씀
                    line.append(None)
                else:
                    line.append(_to_value2(v))
            grid.append(tuple(line))
        chunk = max(1, 20000 // ncol)
        for s in range(0, len(grid), chunk):
            part = grid[s:s + chunk]
            ws.Range(cells(top + 1 + s, left), cells(top + s + len(part), left + ncol - 1)).Value2 = tuple(part)
        for i, j, v in mixed_cells:
            self._write_cell(cells(top + 1 + i, left + j), v)
        for j, c in enumerate(columns):
            if kinds[c] in ("date", "datetime"):
                set_nf(ws.Range(cells(top + 1, left + j), cells(top + nrow, left + j)),
                       NF["date"] if kinds[c] == "date" else NF["dt"])
        rng = ws.Range(cells(top, left), cells(top + nrow, left + ncol - 1))
        lo = ws.ListObjects.Add(XL_SRC_RANGE, rng, None, XL_YES)
        lo.Name = name
        lo.TableStyle = DATA_STYLE
        return lo

    def _write_cell(self, cell, value: Any, *, keep_number_format: bool = False) -> None:
        """셀 하나에 형식을 지켜 씀: 문자열 → 텍스트(@), 날짜 → 일련번호+날짜 서식, 숫자 → 숫자, None → 지움."""
        if value is None:
            cell.ClearContents()
            return
        if isinstance(value, str):
            cell.NumberFormat = "@"
            cell.Value2 = value
            return
        if isinstance(value, (dt.datetime, dt.date)):
            cell.Value2 = excel_serial(value)
            set_nf(cell, NF["dt"] if isinstance(value, dt.datetime) else NF["date"])
            return
        if not keep_number_format or cell.NumberFormat == "@":
            set_nf(cell, "General")
        cell.Value2 = _to_value2(value)

    # ------------------------------------------------------------ 정적 표 API
    def static(self, name: str, source: Optional[str] = None, *, rows: Optional[list] = None,
               columns: Optional[list] = None, sheet: Optional[str] = None, cell: Optional[str] = None,
               text_cols: Sequence[str] = (), date_cols: Sequence[str] = ()):
        """정적 표 만들기·교체 (CSV/JSON 파일 또는 rows). 같은 이름이 있으면 그 자리(또는 sheet/cell)에 새로 씀.

        Args:
            name: 표 이름 (예: tblContest, tblPxStoreSeed). tblToken·쿼리 적재 표는 불가.
            source: .csv(UTF-8, BOM 허용, 실패 시 cp949) 또는 .json(레코드 목록 또는 {"columns", "rows"}).
            rows/columns: 파이썬 값으로 직접 — dict 목록(열은 키 순서) 또는 리스트 목록(columns 필요).
            text_cols/date_cols: 강제로 텍스트/날짜로 쓸 열.
        Returns: ListObject
        """
        self._require()
        if name.lower() == TOKEN_TABLE.lower():
            raise HarnessError("tblToken은 하네스가 관리합니다(정적 표로 바꿀 수 없음).")
        if name in self.query_of:
            raise HarnessError(f"{name}은 쿼리 적재 표입니다(셀 수정은 set_cells 사용).")
        if source is not None:
            columns, rows, text_extra = self._read_source(source, text_cols)
            text_cols = tuple(text_cols) + tuple(text_extra)
        else:
            columns, rows = self._rows_from_python(rows or [], columns)
        old = self._find_lo(name)
        if old is not None:
            ws_old = old.Range.Worksheet
            sheet = sheet or ws_old.Name
            cell = cell or old.Range.Cells(1, 1).Address.replace("$", "")
            old.Delete()
        lo = self._write_new_table(name, list(columns), rows, sheet=sheet, cell=cell, text_cols=text_cols,
                                   date_cols=date_cols)
        n = 0 if not rows else len(rows)
        self._step("static", table=name, rows=n, columns=list(columns), source=source)
        self.say(f"[하네스] 정적 표 {name}: {n}행 × {len(columns)}열" + (f" ← {source}" if source else ""))
        return lo

    @staticmethod
    def _rows_from_python(rows: list, columns: Optional[list]) -> tuple[list, list]:
        if rows and isinstance(rows[0], dict):
            cols = list(columns or [])
            for r in rows:
                for k in r:
                    if k not in cols:
                        cols.append(k)
            return cols, [[r.get(c) for c in cols] for r in rows]
        if columns is None:
            raise HarnessError("리스트 행에는 columns가 필요합니다.")
        return list(columns), [list(r) for r in rows]

    def _read_source(self, source: str, text_cols: Sequence[str]) -> tuple[list, list, list]:
        path = os.path.abspath(source)
        if not os.path.isfile(path):
            raise HarnessError(f"정적 표 원천 파일이 없습니다: {path}")
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            return self._read_csv(path, text_cols)
        if ext == ".json":
            with open(path, encoding="utf-8-sig") as fh:
                data = json.load(fh)
            if isinstance(data, dict) and "columns" in data:
                cols, rows = list(data["columns"]), [list(r) for r in data.get("rows", [])]
            elif isinstance(data, dict) and "rows" in data:
                cols, rows = self._rows_from_python(data["rows"], None)
            elif isinstance(data, list):
                cols, rows = self._rows_from_python(data, None)
            else:
                raise HarnessError(f"JSON 형식을 알 수 없습니다(레코드 목록 또는 columns·rows): {path}")
            text_set = set(text_cols) | {c for c in cols if _is_code_column(c)}
            rows = [[(_iso_to_date(v) if isinstance(v, str) and cols[j] not in text_set else v)
                     for j, v in enumerate(r)] for r in rows]
            return cols, rows, []
        raise HarnessError(f"정적 표 원천은 .csv 또는 .json이어야 합니다: {path}")

    @staticmethod
    def _read_csv(path: str, text_cols: Sequence[str]) -> tuple[list, list, list]:
        raw: Optional[list] = None
        for enc in ("utf-8-sig", "cp949"):
            try:
                with open(path, encoding=enc, newline="") as fh:
                    raw = [r for r in csv.reader(fh) if any(x.strip() for x in r)]
                break
            except UnicodeDecodeError:
                continue
        if not raw:
            raise HarnessError(f"CSV가 비었거나 읽을 수 없습니다: {path}")
        cols = [c.strip() for c in raw[0]]
        body = [r + [""] * (len(cols) - len(r)) if len(r) < len(cols) else r[:len(cols)] for r in raw[1:]]
        text_extra = []
        for j, c in enumerate(cols):
            vals = [r[j].strip() for r in body if r[j].strip()]
            if c in text_cols or _is_code_column(c) or any(len(v) > 1 and v.isdigit() and v.startswith("0") for v in vals):
                text_extra.append(c)
        rows = [[(r[j] if r[j] != "" else None) if cols[j] in text_extra else coerce_text(r[j], cli=False)
                 for j in range(len(cols))] for r in body]
        return cols, rows, text_extra

    def _upsert_kv(self, table: str, values: dict) -> None:
        """키/값 표(tblSettings·tblRunCtl)의 값 변경·키 추가."""
        lo = self._lo(table)
        cols = [str(c) for c in _as_grid(lo.HeaderRowRange.Value)[0]]
        if "키" not in cols or "값" not in cols:
            raise HarnessError(f"{table}에 키/값 열이 없습니다: {cols}")
        jk, jv = cols.index("키") + 1, cols.index("값") + 1
        for key, value in values.items():
            body = lo.DataBodyRange
            keys = [str(r[0]).strip() if r[0] is not None else "" for r in _as_grid(body.Columns(jk).Value2)] \
                if body is not None else []
            existing = key in keys
            if existing:
                i = keys.index(key) + 1
            elif len(keys) == 1 and keys[0] == "":
                i = 1                                         # 빈 표의 빈 행 사용
            else:
                lo.ListRows.Add()                             # 새 행은 위 행 서식을 물려받으므로 아래에서 서식을 다시 정함
                i = lo.ListRows.Count
            if not existing:
                self._write_cell(lo.DataBodyRange.Cells(i, jk), key)
            # 기존 키는 숫자 서식(%, 원 등)을 유지, 새 키는 값 형식에 맞게 서식을 새로 정함
            self._write_cell(lo.DataBodyRange.Cells(i, jv), value, keep_number_format=existing)

    def settings(self, values: Optional[dict] = None, **kw: Any) -> None:
        """tblSettings 값 변경·키 추가 (예: settings(hist_days=60, vs_path=r"D:/수집기업_valuesearch.xlsx"))."""
        self._require()
        merged = dict(values or {}, **kw)
        self._upsert_kv("tblSettings", merged)
        self._step("settings", values={k: _fmt(v) for k, v in merged.items()})
        self.say("[하네스] tblSettings: " + ", ".join(f"{k}={_fmt(v)}" for k, v in merged.items()))

    def runctl(self, values: Optional[dict] = None, **kw: Any) -> None:
        """tblRunCtl 값 변경·키 추가 (예: runctl(mode="full", now_override=dt.datetime(2026, 10, 2, 8, 30)))."""
        self._require()
        merged = dict(values or {}, **kw)
        self._upsert_kv("tblRunCtl", merged)
        self._step("runctl", values={k: _fmt(v) for k, v in merged.items()})
        self.say("[하네스] tblRunCtl: " + ", ".join(f"{k}={_fmt(v)}" for k, v in merged.items()))

    # ------------------------------------------------------------ 이름 정의·셀 수정
    def define_name(self, name: str, value: Any = None, *, refers_to: Optional[str] = None,
                    sheet: Optional[str] = None, cell: Optional[str] = None) -> str:
        """이름 정의 만들기(이미 있으면 값만 변경). 값은 셀에 쓰고 이름이 그 셀을 가리킴(파워 쿼리에서 읽을 수 있게).

        Args:
            refers_to: 직접 참조식(예: "=종목분석!$C$5"). 주면 value·sheet·cell은 쓰지 않음.
            sheet/cell: 값을 쓸 위치(기본 `_names` 시트 C열 다음 행, B열에 이름 표시).
        Returns: 참조식
        """
        self._require()
        existing = self._name_obj(name)
        if refers_to is None and existing is not None and sheet is None and cell is None:
            self.set_name(name, value)
            return existing.RefersTo
        if existing is not None:
            existing.Delete()
        written = []                                     # 이름 정의가 실패하면 되돌릴 셀
        if refers_to is None:
            if sheet is None and cell is None:
                ws = self.sheet(NAMES_SHEET)
                r = self._names_row
                ws.Cells(r, 2).Value2 = name
                target = ws.Cells(r, 3)
                written = [ws.Cells(r, 2), target]
            else:
                ws = self.sheet(sheet or NAMES_SHEET)
                target = ws.Range(cell or "C2")
            self._write_cell(target, value)
            refers_to = f"='{ws.Name}'!{target.Address}"
        try:
            self.wb.Names.Add(name, refers_to)
        except pywintypes.com_error as e:
            for c in written:
                c.ClearContents()
            raise HarnessError(f"이름 정의 실패: {name} = {refers_to} — 셀 주소와 겹치는 이름(예: A1, AB12)은 쓸 수 없습니다 "
                               f"({_com_message(e)})") from None
        if written:
            self._names_row += 1
        self._step("name", name=name, refers_to=refers_to, value=_fmt(value))
        self.say(f"[하네스] 이름 정의 {name} = {refers_to}" + (f" (값 {_fmt(value)})" if value is not None else ""))
        return refers_to

    def _name_obj(self, name: str):
        for n in self.wb.Names:
            if n.Name == name:
                return n
        return None

    def set_name(self, name: str, value: Any) -> None:
        """이름 정의가 가리키는 셀의 값 변경 (단계 사이 분석코드 바꾸기 등)."""
        self._require()
        n = self._name_obj(name)
        if n is None:
            raise HarnessError(f"이름 정의가 없습니다: {name}")
        try:
            target = n.RefersToRange.Cells(1, 1)
        except pywintypes.com_error:
            raise HarnessError(f"이름 {name}이 셀을 가리키지 않습니다: {n.RefersTo}") from None
        self._write_cell(target, value)
        self._step("name", name=name, value=_fmt(value))
        self.say(f"[하네스] 이름 {name} 값 = {_fmt(value)}")

    def _match_rows(self, lo, match: dict) -> tuple[list, list[int]]:
        cols = [str(c) for c in _as_grid(lo.HeaderRowRange.Value)[0]]
        missing = [c for c in match if c not in cols]
        if missing:
            raise HarnessError(f"{lo.Name}에 없는 열: {missing} (열: {cols})")
        body = lo.DataBodyRange
        if body is None or lo.ListRows.Count == 0:
            return cols, []
        grid = _as_grid(body.Value2)
        idx = [cols.index(c) for c in match]
        hits = [i for i, r in enumerate(grid) if all(_loose_equal(r[j], match[c]) for j, c in zip(idx, match))]
        return cols, hits

    def set_cells(self, table: str, match: dict, values: dict) -> int:
        """적재된 표(또는 정적 표)에서 match 조건(열=값, 모두 일치)에 맞는 행의 셀을 values로 바꿈.

        다음 새로 고침에서 쿼리가 자기 표를 읽으면 바뀐 값이 보입니다(자기참조 저장소의 분할·수정주가 모사).
        Returns: 바꾼 행 수
        """
        self._require()
        if table.lower() == TOKEN_TABLE.lower():
            raise HarnessError("tblToken은 수정할 수 없습니다.")
        lo = self._lo(table)
        cols, hits = self._match_rows(lo, match)
        bad = [c for c in values if c not in cols]
        if bad:
            raise HarnessError(f"{table}에 없는 열: {bad} (열: {cols})")
        for i in hits:
            for c, v in values.items():
                self._write_cell(lo.DataBodyRange.Cells(i + 1, cols.index(c) + 1), v, keep_number_format=True)
        self._step("set_cells", table=table, match={k: _fmt(v) for k, v in match.items()},
                   values={k: _fmt(v) for k, v in values.items()}, rows=len(hits))
        self.say(f"[하네스] {table} 셀 수정 {len(hits)}행: " + ", ".join(f"{k}={_fmt(v)}" for k, v in values.items()))
        return len(hits)

    def delete_rows(self, table: str, match: dict) -> int:
        """match 조건에 맞는 행 삭제 (빠진 세션 모사 등). Returns: 삭제한 행 수."""
        self._require()
        if table.lower() == TOKEN_TABLE.lower():
            raise HarnessError("tblToken은 수정할 수 없습니다.")
        lo = self._lo(table)
        _cols, hits = self._match_rows(lo, match)
        for i in sorted(hits, reverse=True):
            lo.ListRows(i + 1).Delete()
        self._step("delete_rows", table=table, match={k: _fmt(v) for k, v in match.items()}, rows=len(hits))
        self.say(f"[하네스] {table} 행 삭제 {len(hits)}행")
        return len(hits)

    # ------------------------------------------------------------ 쿼리
    def add_queries(self, *items: str) -> list[str]:
        """쿼리 추가(이름 또는 .pq 경로). 의존 쿼리를 먼저 자동 추가. Returns: 이번에 새로 넣은 쿼리 이름(순서대로)."""
        self._require()
        names = [self.catalog.register(it) for it in items]
        order = self.catalog.resolve(names)
        present = set(self.queries())
        added: list[str] = []
        t = time.time()
        for n in order:
            if n in present and n not in names:
                continue
            add_query(self.wb, n, self.catalog.text(n), "하네스 시험")
            added.append(n)
        self.timings["query_add"] = round(self.timings.get("query_add", 0.0) + time.time() - t, 2)
        self._assert_no_token_query()
        if added:
            self._step("queries", added=added)
            self.say(f"[하네스] 쿼리 추가 {len(added)}개: {', '.join(added)} ({time.time() - t:.1f}s)")
        return added

    def add_query_text(self, name: str, m_code: str) -> list[str]:
        """인라인 M 쿼리 추가 (예: add_query_text("Q_Rows", "fnPageRows()") — 함수 결과를 표로 보기)."""
        self.catalog.add_inline(name, m_code)
        return self.add_queries(name)

    def queries(self) -> list[str]:
        """통합문서의 쿼리 이름 목록."""
        self._require()
        return [q.Name for q in self.wb.Queries]

    def tables(self) -> list[str]:
        """통합문서의 표 이름 목록 (토큰 표 포함 — 이름만)."""
        self._require()
        return [lo.Name for ws in self.wb.Worksheets for lo in ws.ListObjects]

    def _assert_no_token_query(self) -> None:
        if TOKEN_QUERY in self.queries():
            raise HarnessError("시험 통합문서에 T_Token 쿼리가 들어갔습니다 — 중단합니다.")

    def load(self, query: str, table: Optional[str] = None, sheet: Optional[str] = None, cell: Optional[str] = None):
        """쿼리를 표로 적재(새로 고침은 refresh). 쿼리가 없으면 의존 쿼리와 함께 추가.

        Args:
            table: 표 이름(기본: default_table — 빌더 LOADS 이름, 없으면 T_X → tblX, T_A_X → tblA_X).
            sheet: 기본 = 표 이름 시트. cell: 기본 B2.
        Returns: ListObject
        """
        self._require()
        name = self.catalog.register(query)
        if name == TOKEN_QUERY:
            raise HarnessError("T_Token은 적재할 수 없습니다(토큰 발급 경로 차단).")
        if name not in self.queries():
            self.add_queries(query)
        table = table or self.default_table(name)
        if table.lower() == TOKEN_TABLE.lower():
            raise HarnessError("tblToken 이름으로는 적재할 수 없습니다.")
        old = self._find_lo(table)
        if old is not None:
            if self.query_of.get(old.Name) == name:
                return old
            raise HarnessError(f"이미 있는 표 이름입니다: {table}")
        ws = self.sheet(sheet or table)
        lo = load_query(ws, name, cell or DEFAULT_CELL, table, background=True, style_name=DATA_STYLE)
        self.loads.append(table)
        self.query_of[table] = name
        self._step("load", query=name, table=table, sheet=ws.Name, cell=cell or DEFAULT_CELL)
        self.say(f"[하네스] 적재: {name} → {table} ({ws.Name}!{cell or DEFAULT_CELL})")
        return lo

    @staticmethod
    def default_table(query: str) -> str:
        """기본 표 이름: build_dashboard.LOADS에 있는 쿼리는 빌더와 같은 표(예: T_Trades → tblTradeLog), 그 밖은 관례."""
        try:
            for entry in _builder_module().LOADS:        # (쿼리, 시트, 셀, 표, …) — 뒤에 항목이 늘어도 동작
                if len(entry) >= 4 and entry[0] == query:
                    return str(entry[3])
        except (HarnessError, AttributeError, TypeError):
            pass
        return default_table_name(query)

    def _table_for(self, name: str) -> tuple[str, str]:
        """표 또는 쿼리 이름 → (표, 쿼리)."""
        for t, q in self.query_of.items():
            if t.lower() == name.lower() or q == name:
                return t, q
        raise HarnessError(f"적재된 표·쿼리가 아닙니다: {name} (적재: {', '.join(self.loads) or '없음'})")

    # ------------------------------------------------------------ 새로 고침
    def refresh(self, name: str, timeout: Optional[float] = None, sync: bool = False) -> RefreshResult:
        """표 하나를 동기 새로 고침 (백그라운드 + Refreshing 폴링 + AfterRefresh 확인). 오류는 결과에 기록.

        Args:
            name: 표 또는 쿼리 이름. timeout: 초(기본 하네스 설정). sync: True면 처음부터 동기(시간 초과 없음).
        Returns: RefreshResult
        """
        self._require()
        table, query = self._table_for(name)
        lo = self._lo(table)
        warnings: list[str] = []
        try:
            missing = sorted(r for r in self.catalog.refs([query]) if r != table and self._find_lo(r) is None
                             and self._name_obj(r) is None)
            if missing:
                warnings.append("통합문서에 없는 표·이름 참조: " + ", ".join(missing))
        except HarnessError:
            pass
        t0 = time.time()
        if sync:
            ok, err, confirmed, timed_out, rerun = self._refresh_sync(lo) + (False, False, False)
        else:
            ok, err, confirmed, timed_out, rerun, w2 = self._refresh_polled(lo, timeout or self.timeout)
            warnings += w2
        seconds = time.time() - t0
        cols = self._columns(lo)
        rows = 0 if not cols else lo.ListRows.Count
        status = self.status_summary(table) if ("상태" in cols and rows) else {}
        res = RefreshResult(table=table, query=query, ok=ok, rows=rows, columns=cols, seconds=seconds, status=status,
                            error=self._redact_text(err) if err else None, timed_out=timed_out, confirmed=confirmed,
                            rerun=rerun, warnings=warnings)
        self.results.append(res)
        self.timings["refresh_total"] = round(self.timings.get("refresh_total", 0.0) + seconds, 2)
        self.steps.append(self._redact({"no": len(self.steps) + 1, "kind": "refresh", **res.to_dict()}))
        self.say(f"[{len(self.results)}] 새로 고침 " + res.line())
        return res

    def refresh_all(self, names: Optional[Sequence[str]] = None) -> list[RefreshResult]:
        """여러 표를 지정 순서로 새로 고침 (기본: 적재한 표 전부, 적재 순서)."""
        return [self.refresh(n) for n in (names or list(self.loads))]

    def _refresh_sync(self, lo) -> tuple[bool, Optional[str]]:
        """동기 새로 고침(이벤트 끔 — Excel이 M 오류 메시지를 COM 예외로 돌려줌)."""
        prev = self.xl.EnableEvents
        try:
            self.xl.EnableEvents = False
            _com_retry(lambda: lo.QueryTable.Refresh(False))
            return True, None
        except pywintypes.com_error as e:
            return False, _com_message(e) or "새로 고침 실패(메시지 없음)"
        finally:
            self.xl.EnableEvents = prev

    def _advise(self, qt):
        if self._event_ids is None:
            self._event_ids = _refresh_event_ids(qt)
        iid, before, after = self._event_ids
        import win32com.server.util
        from win32com.server.policy import EventHandlerPolicy
        _RefreshSink._com_interfaces_ = [iid]
        _RefreshSink._dispid_to_func_ = {before: "OnBeforeRefresh", after: "OnAfterRefresh"}
        impl = _RefreshSink()
        sink = win32com.server.util.wrap(impl, usePolicy=EventHandlerPolicy)
        cp = qt._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer).FindConnectionPoint(iid)
        cookie = cp.Advise(sink)
        return impl, cp, cookie

    def _nudge(self) -> None:
        """자기 Excel 창에 WM_TIMER를 보내 백그라운드 새로 고침 완료를 바로 확인하게 함.

        Excel은 백그라운드 쿼리 완료를 타이머로 확인하는데, 유휴 상태가 길수록 간격이 1→5→20초로 늘어
        상수 쿼리도 20초씩 걸렸다(실측). 폴링마다 이 메시지를 보내면 약 0.2초에 끝난다. 대상은 이 하네스가 띄운
        Excel(xl.Hwnd)뿐이다.
        """
        if not self._hwnd:
            return
        try:
            import win32con
            import win32gui
            win32gui.PostMessage(self._hwnd, win32con.WM_TIMER, 0, 0)
        except Exception:  # noqa: BLE001 — 깨우기 실패는 지연만 생길 뿐 결과에는 영향 없음
            pass

    def _refresh_polled(self, lo, timeout: float):
        """백그라운드 새로 고침 → Refreshing 폴링(시간 초과 시 취소) → AfterRefresh로 성공 확인 → 실패면 동기 재실행."""
        qt = lo.QueryTable
        xl = self.xl
        warnings: list[str] = []
        impl, cp, cookie = None, None, None
        start_err: Optional[str] = None
        timed_out = False
        prev_events = xl.EnableEvents
        t0 = time.time()
        try:
            xl.EnableEvents = True                    # AfterRefresh 이벤트는 EnableEvents가 켜져 있어야 옴
            try:
                impl, cp, cookie = self._advise(qt)
            except (pywintypes.com_error, AttributeError, ImportError) as e:
                warnings.append(f"새로 고침 이벤트 연결 실패({type(e).__name__}) — 성공 여부는 표 상태로 추정")
            try:
                _com_retry(lambda: qt.Refresh(True))
            except pywintypes.com_error as e:
                start_err = _com_message(e) or "새로 고침을 시작하지 못했습니다"
            if start_err is None:
                deadline = t0 + timeout
                while True:
                    pythoncom.PumpWaitingMessages()
                    self._nudge()
                    if not _com_retry(lambda: qt.Refreshing, timeout=60):
                        break
                    if time.time() > deadline:
                        timed_out = True
                        try:
                            _com_retry(lambda: qt.CancelRefresh())
                        except pywintypes.com_error:
                            pass
                        t_c = time.time()
                        while time.time() - t_c < 15:
                            pythoncom.PumpWaitingMessages()
                            self._nudge()
                            try:
                                if not qt.Refreshing:
                                    break
                            except pywintypes.com_error:
                                pass
                            time.sleep(0.2)
                        break
                    time.sleep(POLL_SECONDS)
                if impl is not None and not timed_out:
                    t_e = time.time()
                    while impl.after is None and time.time() - t_e < 3.0:   # 이벤트가 조금 늦게 올 수 있음
                        pythoncom.PumpWaitingMessages()
                        self._nudge()
                        time.sleep(0.05)
        finally:
            if cp is not None and cookie is not None:
                try:
                    cp.Unadvise(cookie)
                except pywintypes.com_error:
                    pass
            try:
                xl.EnableEvents = prev_events
            except pywintypes.com_error:
                pass
        first_seconds = time.time() - t0
        confirmed = impl is not None and impl.after is not None
        if timed_out:
            return False, f"시간 초과({timeout:.0f}초) — 새로 고침을 취소했습니다", confirmed, True, False, warnings
        if start_err is not None:
            ok, err = False, start_err
        elif confirmed:
            ok, err = bool(impl.after), (None if impl.after else "새로 고침 실패")
        elif self._is_placeholder(lo):
            ok, err = False, "새로 고침 실패(표가 만들어지지 않음)"
        else:
            ok, err = True, None
            warnings.append("성공 여부를 이벤트로 확인하지 못함(표가 직전 값일 수 있음)")
        rerun = False
        if not ok and self.diagnose and (start_err is None or not start_err.strip()):
            if first_seconds <= DIAGNOSE_LIMIT_SECONDS:
                rerun = True
                ok2, err2 = self._refresh_sync(lo)
                if ok2:
                    ok, err = True, None
                    warnings.append("백그라운드 새로 고침 실패 → 동기 재실행 성공")
                else:
                    err = err2
            else:
                warnings.append(f"1차 시도가 {first_seconds:.0f}초 걸려 재실행 생략 — 메시지는 refresh(..., sync=True)로 확인")
        return ok, err, confirmed, False, rerun, warnings

    @staticmethod
    def _is_placeholder(lo) -> bool:
        """첫 새로 고침 전·실패 시 Excel이 넣는 자리표시 머리글('ExternalData_1: …')인지."""
        try:
            if lo.ListColumns.Count != 1:
                return False
            return "ExternalData_" in str(lo.HeaderRowRange.Cells(1, 1).Value)
        except pywintypes.com_error:
            return False

    def _columns(self, lo) -> list[str]:
        if self._is_placeholder(lo):
            return []
        return [str(c) for c in _as_grid(lo.HeaderRowRange.Value)[0]]

    # ------------------------------------------------------------ 결과 읽기
    def table(self, name: str, max_rows: Optional[int] = None, dates: Any = "auto") -> TableData:
        """표 내용(파이썬 값). tblToken은 거부.

        Args:
            max_rows: 앞에서 이만큼만. dates: "auto"(이름·값으로 날짜 열 추정해 변환), "none"(원값), 열 이름 목록.
        """
        self._require()
        if name.lower() == TOKEN_TABLE.lower():
            raise HarnessError("tblToken은 읽기·출력·덤프를 하지 않습니다(토큰 비출력).")
        lo = self._lo(name)
        cols = self._columns(lo)
        if not cols or lo.ListRows.Count == 0 or lo.DataBodyRange is None:
            return TableData(lo.Name, cols, [])
        body = lo.DataBodyRange
        if max_rows is not None and max_rows < lo.ListRows.Count:
            # Range.Resize(n)은 늦은 바인딩에서 Item(n)으로 해석되므로 범위를 직접 만듦
            body = body.Worksheet.Range(body.Cells(1, 1), body.Cells(max(1, max_rows), len(cols)))
        rows = [[_py_value(v) for v in r] for r in _as_grid(body.Value)]
        if dates == "auto":
            idx = _auto_date_columns(cols, rows)
        elif dates in (None, "none"):
            idx = []
        else:
            idx = [cols.index(c) for c in dates if c in cols]
        for j in idx:
            for r in rows:
                v = r[j]
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    d = _serial_to_dt(float(v))
                    r[j] = d.date() if float(v).is_integer() else d
        return TableData(lo.Name, cols, self._redact(rows))

    def row_count(self, name: str) -> int:
        """표의 데이터 행 수 (첫 새로 고침 전·실패한 자리표시 표는 0)."""
        lo = self._lo(name)
        return 0 if not self._columns(lo) else int(lo.ListRows.Count)

    def _status_values(self, name: str) -> Optional[list]:
        """`상태` 열 값만 읽음 (큰 표도 빠르게). 열이 없으면 None."""
        if name.lower() == TOKEN_TABLE.lower():
            raise HarnessError("tblToken은 읽기·출력·덤프를 하지 않습니다(토큰 비출력).")
        lo = self._lo(name)
        if "상태" not in self._columns(lo):
            return None
        if lo.ListRows.Count == 0:
            return []
        return [_py_value(r[0]) for r in _as_grid(lo.ListColumns("상태").DataBodyRange.Value)]

    def status_summary(self, name: str) -> dict:
        """`상태` 열 요약: {"OK": n, "오류": n, "이전 데이터": n, 그 밖 값: n} (많은 순). 열이 없으면 빈 dict."""
        vals = self._status_values(name)
        if vals is None:
            return {}
        counts: dict[str, int] = {}
        for v in vals:
            s = "(빈칸)" if v in (None, "") else str(v)
            key = "오류" if s.startswith("오류") else "이전 데이터" if s.startswith("이전 데이터") else s
            counts[key] = counts.get(key, 0) + 1
        return dict(sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])))

    def status_errors(self, name: str, top: int = 5) -> list[tuple[str, int]]:
        """`상태`가 오류·이전 데이터인 행의 메시지별 건수 (많은 순)."""
        counts: dict[str, int] = {}
        for v in self._status_values(name) or []:
            s = "" if v is None else str(v)
            if s.startswith("오류") or s.startswith("이전 데이터"):
                counts[s] = counts.get(s, 0) + 1
        return [(self._redact_text(k), n) for k, n in sorted(counts.items(), key=lambda kv: -kv[1])[:top]]

    def report(self, name: str, head: int = 5) -> str:
        """표 요약 문자열: 행·열 수, 상태 요약·오류 메시지, 열 이름, 앞 head행."""
        data = self.table(name, max_rows=max(1, head))
        nrows = self.row_count(name)
        lines = [f"== {data.name}: {nrows}행 × {len(data.columns)}열"
                 + (f" ← {self.query_of[data.name]}" if data.name in self.query_of else "")]
        if "상태" in data.columns:
            lines.append("   상태: " + ", ".join(f"{k} {v}" for k, v in self.status_summary(name).items()))
            for msg, n in self.status_errors(name, 3):
                lines.append(f"   · {n}건: {msg[:160]}")
        lines.append("   열: " + " | ".join(data.columns))
        for i, r in enumerate(data.rows[:max(0, head)], 1):
            lines.append(f"   {i:>3}: " + " | ".join(_fmt(v)[:24] for v in r))
        return self._redact_text("\n".join(lines))

    def dump(self, name: str, csv_path: str, dates: Any = "auto") -> str:
        """표를 CSV(UTF-8 BOM)로 저장. tblToken은 거부. Returns: 저장 경로."""
        if name.lower() == TOKEN_TABLE.lower():
            raise HarnessError("tblToken은 덤프하지 않습니다(토큰을 파일에 쓰지 않음).")
        path = self._out_path(csv_path)
        data = self.table(name, dates=dates)
        with open(path, "w", encoding="utf-8-sig", newline="") as fh:
            w = csv.writer(fh)
            w.writerow(data.columns)
            for r in data.rows:
                w.writerow([self._redact_text(_fmt(v)) for v in r])
        self._step("dump", table=data.name, path=path, rows=len(data))
        self.say(f"[하네스] CSV 저장: {data.name} {len(data)}행 → {path}")
        return path

    def _out_path(self, path: str) -> str:
        """결과 파일 경로 확인: git 저장소 안이면 거부(allow_repo_output 제외), 폴더 생성."""
        full = os.path.abspath(path)
        repo = _inside_git_repo(full)
        if repo and not self.allow_repo_output:
            raise HarnessError(f"git 저장소 안에는 결과 파일을 쓰지 않습니다: {full} (저장소 {repo}) — 스크래치 폴더를 쓰세요.")
        os.makedirs(os.path.dirname(full), exist_ok=True)
        return full

    # ------------------------------------------------------------ 저장·넘겨주기·렌더링
    def save(self, path: str) -> str:
        """시험 통합문서를 지정 경로에 저장(.xlsx=51, .xlsm=52). 저장 뒤에도 열린 채로 계속 쓸 수 있음.

        저장 파일에는 토큰 표가 들어 있으므로 스크래치 폴더에만 두고 커밋하지 마세요. Returns: 저장 경로.
        """
        self._require()
        full = self._out_path(path)
        ext = os.path.splitext(full)[1].lower()
        if ext not in (".xlsx", ".xlsm"):
            raise HarnessError(f"저장 확장자는 .xlsx 또는 .xlsm이어야 합니다: {full}")
        if os.path.exists(full):
            os.remove(full)
        self.wb.SaveAs(full, 52 if ext == ".xlsm" else 51)
        self.saved_path = full
        self._step("save", path=full)
        self.say(f"[하네스] 통합문서 저장: {full} (열린 채 유지)")
        return full

    def builder(self, sheets: Sequence[str] = ()) -> StandInBuilder:
        """페이지 빌더 대역 객체 (sheets에 준 시트는 미리 만듦)."""
        self._require()
        for s in sheets:
            self.sheet(s)
        b = StandInBuilder(self)
        self._builders.append(b)
        return b

    def render(self, sheet: str, png_path: str, cell_range: Optional[str] = None, dpi: int = 150,
               crop: bool = True) -> str:
        """시트 범위(기본 UsedRange)를 PNG로 렌더링: 한 쪽 맞춤 PDF(ExportAsFixedFormat) → pymupdf → PNG.

        클립보드를 쓰지 않아 병렬 작업·사용자 클립보드와 충돌하지 않습니다. `_sys`(토큰) 시트와 토큰이 보이는 범위는 거부.
        Returns: PNG 경로
        """
        self._require()
        import fitz
        ws = None
        for w in self.wb.Worksheets:
            if w.Name == sheet:
                ws = w
        if ws is None:
            raise HarnessError(f"시트가 없습니다: {sheet}")
        if ws.Name == TOKEN_SHEET or any(lo.Name.lower() == TOKEN_TABLE.lower() for lo in ws.ListObjects):
            raise HarnessError("토큰 표가 있는 시트는 렌더링하지 않습니다.")
        rng = ws.Range(cell_range) if cell_range else ws.UsedRange
        if self._token and self._token in str(rng.Value2):
            raise HarnessError("범위에 토큰 값이 보여 렌더링하지 않습니다.")
        png = self._out_path(png_path)
        pdf = os.path.splitext(png)[0] + ".__render.pdf"
        ps = ws.PageSetup
        saved = {}
        was_visible = ws.Visible
        try:
            if was_visible != -1:
                ws.Visible = -1
            for k in ("Zoom", "FitToPagesWide", "FitToPagesTall", "Orientation", "LeftMargin", "RightMargin",
                      "TopMargin", "BottomMargin"):
                try:
                    saved[k] = getattr(ps, k)
                except pywintypes.com_error:
                    pass
            self._set_print_communication(False)
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = 1
            ps.Orientation = 2 if rng.Width >= rng.Height else 1
            margin = self.xl.InchesToPoints(0.2)
            ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = margin
            self._set_print_communication(True)
            if os.path.exists(pdf):
                os.remove(pdf)
            rng.ExportAsFixedFormat(0, pdf, 0, False, True)
        finally:
            try:
                self._set_print_communication(False)
                for k, v in saved.items():
                    try:
                        setattr(ps, k, v)
                    except pywintypes.com_error:
                        pass
                self._set_print_communication(True)
            except pywintypes.com_error:
                pass
            if was_visible != -1:
                ws.Visible = was_visible
        doc = fitz.open(pdf)
        try:
            pix = doc[0].get_pixmap(dpi=dpi)
            pix.save(png)
        finally:
            doc.close()
            try:
                os.remove(pdf)
            except OSError:
                pass
        if crop:
            self._crop_png(png)
        self._step("render", sheet=sheet, range=cell_range or rng.Address, path=png)
        self.say(f"[하네스] PNG 렌더링: {sheet}!{cell_range or rng.Address.replace('$', '')} → {png}")
        return png

    def _set_print_communication(self, on: bool) -> None:
        try:
            self.xl.PrintCommunication = on
        except pywintypes.com_error:
            pass

    @staticmethod
    def _crop_png(png: str, pad: int = 12) -> None:
        """흰 여백 잘라내기."""
        from PIL import Image, ImageChops
        with Image.open(png) as im:
            rgb = im.convert("RGB")
        diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, (255, 255, 255)))
        box = diff.getbbox()
        if box:
            l, t, r, b = box
            rgb.crop((max(0, l - pad), max(0, t - pad), min(rgb.width, r + pad), min(rgb.height, b + pad))).save(png)

    # ------------------------------------------------------------ 요약
    def summary(self, head: int = 5) -> dict:
        """JSON으로 쓸 수 있는 실행 요약 (토큰 값 없음 — 남은 분만)."""
        tables: dict[str, dict] = {}
        if self.wb is not None:
            for name in self.tables():
                if name.lower() == TOKEN_TABLE.lower():
                    tables[name] = {"kind": "token", "rows": 1, "note": "토큰 표 — 값은 출력하지 않음"}
                    continue
                try:
                    data = self.table(name, max_rows=max(1, head))
                    lo = self._lo(name)
                    info = {"kind": "query" if name in self.query_of else "static", "query": self.query_of.get(name),
                            "sheet": lo.Range.Worksheet.Name, "rows": self.row_count(name), "columns": data.columns,
                            "head": [[_fmt(v) for v in r] for r in data.rows[:max(0, head)]]}
                    if "상태" in data.columns:
                        info["status"] = self.status_summary(name)
                    tables[name] = info
                except (HarnessError, pywintypes.com_error) as e:
                    tables[name] = {"error": str(e)}
        out = {
            "harness": "pq_harness", "version": 1,
            "token": dict(self.token_info),
            "excel_started": self.excel_started, "excel_pid": self.pid,
            "timings": dict(self.timings),
            "queries": self.queries() if self.wb is not None else [],
            "loads": [{"table": t, "query": self.query_of[t]} for t in self.loads],
            "steps": self.steps,
            "tables": tables,
            "saved_path": self.saved_path,
            "refresh_ok": all(r.ok for r in self.results),
            "refresh_failed": [r.table for r in self.results if not r.ok],
        }
        return self._redact(out)


def _loose_equal(cell: Any, wanted: Any) -> bool:
    """set_cells·delete_rows 조건 비교 (Value2 기준: 날짜는 일련번호, 텍스트 코드와 숫자는 느슨하게)."""
    if wanted is None:
        return cell is None or cell == ""
    if isinstance(wanted, (dt.datetime, dt.date)):
        if not isinstance(cell, (int, float)) or isinstance(cell, bool):
            return False
        ser = excel_serial(wanted)
        if isinstance(wanted, dt.datetime):
            return abs(cell - ser) < 1e-6
        return math.floor(cell) == ser
    if isinstance(wanted, bool):
        return cell is wanted
    if isinstance(wanted, (int, float)):
        if isinstance(cell, (int, float)) and not isinstance(cell, bool):
            return float(cell) == float(wanted)
        if isinstance(cell, str):
            try:
                return float(cell.strip()) == float(wanted)
            except ValueError:
                return False
        return False
    s = str(wanted)
    if isinstance(cell, str):
        return cell == s
    if isinstance(cell, (int, float)) and not isinstance(cell, bool):
        try:
            return float(s) == float(cell)
        except ValueError:
            return False
    return False


# ---------------------------------------------------------------- CLI
class _Step(argparse.Action):
    """순서가 있는 단계 옵션: (종류, 값)을 같은 목록에 등장 순서대로 쌓음."""

    def __call__(self, parser, namespace, values, option_string=None):
        steps = list(getattr(namespace, "steps", None) or [])
        steps.append((self.metavar_kind, values))
        namespace.steps = steps

    def __init__(self, option_strings, dest, kind: str = "", **kw):
        self.metavar_kind = kind
        super().__init__(option_strings, dest, **kw)


def build_arg_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="pq_harness.py", allow_abbrev=False, formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Power Query 단위 시험 하네스 — 전체 빌드 없이 .pq 쿼리를 실제 KIS API로 시험 (자세한 설명은 모듈 docstring)",
        epilog="예: python excel_dashboard/tools/pq_harness.py --queries T_IndexNow --load T_IndexNow=tblIndexNow --show 5")
    ap.set_defaults(steps=[])
    g = ap.add_argument_group("쿼리 준비")
    g.add_argument("--queries", nargs="+", action="extend", default=[], metavar="이름|경로.pq")
    g.add_argument("--m", action="append", default=[], metavar="이름=M식", help="인라인 쿼리 (함수 결과 보기 등)")
    g.add_argument("--load", action="append", default=[], metavar="쿼리[=표][@시트!셀]")
    g.add_argument("--pq-dir", action="append", default=[], metavar="폴더", help="추가 검색 폴더(기본 폴더보다 우선)")
    s = ap.add_argument_group("단계 (적힌 순서대로 실행)")
    s.add_argument("--settings", nargs="+", action=_Step, kind="settings", dest="steps", metavar="키=값",
                   help="tblSettings 값 변경·키 추가")
    s.add_argument("--runctl", nargs="+", action=_Step, kind="runctl", dest="steps", metavar="키=값",
                   help="tblRunCtl 값 변경 (mode=build|quick|full|sector|analysis, started, now_override)")
    s.add_argument("--static", action=_Step, kind="static", dest="steps", metavar="표=경로[@시트!셀]",
                   help="정적 표 만들기·교체 (.csv/.json)")
    s.add_argument("--name", action=_Step, kind="name", dest="steps", metavar="이름=값", help="이름 정의 만들기·값 변경")
    s.add_argument("--set", action=_Step, kind="set", dest="steps", metavar="표[열=값,…].열=값",
                   help="적재된 표의 셀 값 수정")
    s.add_argument("--refresh", action=_Step, kind="refresh", dest="steps", metavar="표|쿼리|all",
                   help="동기 새로 고침 (없으면 끝에 적재 표 전부)")
    o = ap.add_argument_group("결과·저장")
    o.add_argument("--show", type=int, default=5, metavar="N", help="적재 표마다 앞 N행 출력(0이면 생략)")
    o.add_argument("--dump", action="append", default=[], metavar="표=경로.csv")
    o.add_argument("--dump-dir", metavar="폴더")
    o.add_argument("--json", metavar="경로")
    o.add_argument("--render", action="append", default=[], metavar="시트[!범위]=경로.png")
    o.add_argument("--save", metavar="경로.xlsx")
    o.add_argument("--keep-open", action="store_true", help="끝난 뒤 Excel을 보이게 열어 둠(사람이 직접 닫음)")
    e = ap.add_argument_group("기타")
    e.add_argument("--min-minutes", type=int, default=kis_dev.MIN_MINUTES, metavar="분",
                   help="필요한 토큰 남은 시간(기본 210분) — 미달이면 Excel 없이 종료 코드 2")
    e.add_argument("--token-source", metavar="경로", help="토큰 원천 통합문서(파일 파싱만, Excel로 열지 않음)")
    e.add_argument("--no-token", action="store_true", help="토큰 표 없이(KIS 호출 불가 — 로컬 쿼리·페이지 시험)")
    e.add_argument("--cfg", metavar="kis_devlp.yaml", help="tblSettings cfg_path 값")
    e.add_argument("--sample", action="store_true", help="빌더 샘플 관심종목·매매일지·대회 기간으로 입력표 채움")
    e.add_argument("--visible", action="store_true", help="Excel 창 표시")
    e.add_argument("--timeout", type=float, default=900.0, metavar="초", help="표당 새로 고침 시간 제한(기본 900)")
    e.add_argument("--no-diagnose", action="store_true", help="실패 시 오류 메시지용 동기 재실행 안 함")
    e.add_argument("--allow-repo-output", action="store_true", help="git 저장소 안에도 결과 파일 쓰기 허용")
    e.add_argument("--list-queries", action="store_true", help="검색 폴더의 쿼리 목록만 출력")
    return ap


def _split_eq(spec: str, what: str) -> tuple[str, str]:
    if "=" not in spec:
        raise HarnessError(f"{what} 형식 오류(이름=값): {spec}")
    k, v = spec.split("=", 1)
    if not k.strip():
        raise HarnessError(f"{what} 형식 오류(이름이 비었음): {spec}")
    return k.strip(), v


def _parse_kv(items: Sequence[str], what: str) -> dict:
    out = {}
    for it in items:
        k, v = _split_eq(it, what)
        out[k] = coerce_text(v)
    return out


def _parse_location(spec: str) -> tuple[str, Optional[str], Optional[str]]:
    """'앞부분@시트!셀' → (앞부분, 시트, 셀). @ 뒤에 !가 없으면 위치 없음으로 봄."""
    if "@" in spec:
        head, loc = spec.rsplit("@", 1)
        if "!" in loc:
            sheet, cell = loc.rsplit("!", 1)
            return head, sheet.strip("'"), cell
    return spec, None, None


def _parse_set(spec: str) -> tuple[str, dict, str, Any]:
    m = re.fullmatch(r"\s*(?P<table>[^\[\]]+)\[(?P<match>[^\]]*)\]\.(?P<col>[^=]+)=(?P<val>.*)", spec)
    if not m:
        raise HarnessError(f'--set 형식 오류: {spec} (예: "tblPxStore[종목코드=005930,일자=2026-09-30].종가=1000")')
    match = _parse_kv([p for p in m.group("match").split(",") if p.strip()], "--set 조건")
    return m.group("table").strip(), match, m.group("col").strip(), coerce_text(m.group("val"))


def _write_json(path: Optional[str], data: dict, allow_repo: bool) -> None:
    if not path:
        return
    full = os.path.abspath(path)
    repo = _inside_git_repo(full)
    if repo and not allow_repo:
        print(f"[하네스] JSON을 저장소 안에 쓰지 않습니다: {full}", flush=True)
        return
    os.makedirs(os.path.dirname(full), exist_ok=True)
    with open(full, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=1, default=str)


def run_cli(argv: Optional[Sequence[str]] = None) -> int:
    """CLI 실행 → 종료 코드 (0 성공 / 1 새로 고침 실패 / 2 토큰 갱신 필요 / 3 설정 오류 / 4 예기치 않은 오류)."""
    ap = build_arg_parser()
    a = ap.parse_args(argv)
    try:
        h = Harness(pq_dirs=a.pq_dir, token_source=a.token_source, min_minutes=a.min_minutes, use_token=not a.no_token,
                    cfg_path=a.cfg, sample=a.sample, visible=a.visible or a.keep_open, timeout=a.timeout,
                    diagnose=not a.no_diagnose, allow_repo_output=a.allow_repo_output)
    except HarnessError as ex:
        print(f"[하네스] 설정 오류: {ex}", flush=True)
        _write_json(a.json, {"harness": "pq_harness", "version": 1, "excel_started": False, "excel_pid": None,
                             "exit_code": EXIT_CONFIG, "ok": False, "error": str(ex)}, a.allow_repo_output)
        return EXIT_CONFIG
    if a.list_queries:
        for n in sorted(h.catalog.files):
            note = "\t(하네스에서 사용 불가 — 토큰 발급 쿼리)" if n == TOKEN_QUERY else ""
            print(f"{n}\t{h.catalog.files[n]}{note}")
        return EXIT_OK
    code, err = EXIT_OK, None
    try:
        # 1) 쿼리 해석·검증 (Excel 기동 전: 파일 없음·T_Token 요청/참조는 여기서 거부)
        names = [h.catalog.register(q) for q in a.queries]
        for spec in a.m:
            k, v = _split_eq(spec, "--m")
            names.append(h.catalog.add_inline(k, v))
        loads = []
        for spec in a.load:
            head, sheet, cell = _parse_location(spec)
            q, _, t = head.partition("=")
            loads.append((h.catalog.register(q.strip()), t.strip() or None, sheet, cell))
        if not loads:
            loads = [(n, None, None, None) for n in names if not n.startswith("fn")]
        h.catalog.resolve(names + [q for q, _, _, _ in loads])
        # 2) 토큰 확인 → Excel
        h.start()
        if names:
            h.add_queries(*names)
        for q, t, sheet, cell in loads:
            h.load(q, t, sheet, cell)
        # 3) 단계
        steps = list(a.steps)
        if not any(k == "refresh" for k, _ in steps):
            steps.append(("refresh", "all"))
        for kind, val in steps:
            if kind == "settings":
                h.settings(_parse_kv(val, "--settings"))
            elif kind == "runctl":
                h.runctl(_parse_kv(val, "--runctl"))
            elif kind == "static":
                head, sheet, cell = _parse_location(val)
                tname, path = _split_eq(head, "--static")
                h.static(tname, path, sheet=sheet, cell=cell)
            elif kind == "name":
                k, v = _split_eq(val, "--name")
                h.define_name(k, coerce_text(v))
            elif kind == "set":
                tname, match, col, value = _parse_set(val)
                h.set_cells(tname, match, {col: value})
            elif kind == "refresh":
                if val == "all":
                    h.refresh_all()
                else:
                    h.refresh(val)
        # 4) 결과
        if a.show > 0:
            for t in h.loads:
                print(h.report(t, head=a.show), flush=True)
        for spec in a.dump:
            tname, path = _split_eq(spec, "--dump")
            h.dump(tname, path)
        if a.dump_dir:
            for t in h.loads:
                h.dump(t, os.path.join(a.dump_dir, t + ".csv"))
        for spec in a.render:
            left, path = _split_eq(spec, "--render")
            sheet, _, rng = left.partition("!")
            h.render(sheet, path, rng or None)
        if a.save:
            h.save(a.save)
        if a.keep_open:
            h.keep_open(detach=True)
        failed = [r for r in h.results if not r.ok]
        code = EXIT_FAILED if failed else EXIT_OK
        h.say(f"[하네스] 완료: 새로 고침 {len(h.results)}건 중 실패 {len(failed)}건"
              + (f" ({', '.join(r.table for r in failed)})" if failed else ""))
    except TokenGateError as ex:
        code, err = EXIT_TOKEN, str(ex)
        print(f"[하네스] 중단: {ex}", flush=True)
    except HarnessError as ex:
        code, err = EXIT_CONFIG, str(ex)
        print(f"[하네스] 설정 오류: {ex}", flush=True)
    except Exception as ex:  # noqa: BLE001 — CLI는 어떤 오류든 종료 코드와 JSON으로 남김
        code, err = EXIT_UNEXPECTED, f"{type(ex).__name__}: {_com_message(ex)}"
        print(f"[하네스] 예기치 않은 오류: {h._redact_text(err)}", flush=True)
    finally:
        try:
            summary = h.summary(head=a.show if a.show > 0 else 5)
        except Exception as ex:  # noqa: BLE001 — 요약 실패가 종료를 막지 않게
            summary = {"error": f"요약 실패: {type(ex).__name__}"}
        summary["exit_code"] = code
        summary["ok"] = code == EXIT_OK
        summary["error"] = h._redact_text(err) if err else None
        h.close()
        summary["timings"] = dict(h.timings)
        _write_json(a.json, summary, a.allow_repo_output)
    return code


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass
    sys.exit(run_cli())


if __name__ == "__main__":
    main()
