# -*- coding: utf-8 -*-
"""통합 검증 공용 도우미 — 토큰 관문, 자기 Excel 인스턴스 열기·닫기, COM 바쁨 재시도, 표 요약.

이 모듈은 `verify_*.py` 스크립트들이 함께 쓰는 작은 도구 모음입니다.

- 토큰 관문(`token_gate`): 통합문서를 Excel로 열기 직전과 버튼 매크로를 실행하기 직전마다 토큰 원천 통합문서를
  **Excel로 열지 않고 파일을 직접 파싱해**(`kis_dev.remaining_minutes`) 남은 유효시간을 다시 확인합니다. 기준(기본 210분 =
  3시간 30분) 미만이면 `TokenGateError`로 멈춥니다(갱신은 한 사람이 한 번에 — docs/security.md의 '개발 중 토큰 규칙').
  출력은 남은 분만이며 토큰 값은 어디에도 쓰지 않습니다.
- Excel: 항상 `DispatchEx`로 새 인스턴스를 띄우고(사용자가 연 Excel 창과 분리), 끝나면 자기 PID만 종료합니다.
  `Application.CalculateUntilAsyncQueriesDone`은 쓰지 않습니다(통합문서 표를 읽는 쿼리와 교착) — `QueryTable.Refreshing` 폴링만.
- COM 바쁨: 새로 고침·매크로 실행 중 Excel은 외부 호출을 거절할 수 있어(RPC_E_CALL_REJECTED) `com_retry`로 다시 시도합니다.
- 표 요약(`table_summaries`): 표마다 행 수, `상태` 열의 값 종류별 개수(OK·오류:·이전 데이터·데이터 없음·추정 없음),
  `조회시각` 최솟값·최댓값을 COM의 워크시트 함수로 셉니다(13만 행 저장소도 값 전체를 옮기지 않음). 토큰 표는 env·expires·
  status·checked만 읽습니다.

    from verify_common import token_gate, open_workbook, close_workbook, table_summaries
    left = token_gate(source)                       # 미달이면 TokenGateError
    xl, wb, pid = open_workbook(path)               # 파일 열 때 T_Token 새로 고침이 끝날 때까지 기다림
    print(table_summaries(wb)["tblPxStore"])
    close_workbook(xl, wb, pid, save=True)
"""
from __future__ import annotations

import datetime as dt
import os
import subprocess
import sys
import threading
import time
from typing import Any, Callable, Optional

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
for _p in (HERE, DASH):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import kis_dev  # noqa: E402

import pythoncom  # noqa: E402
import pywintypes  # noqa: E402
import win32com.client  # noqa: E402

TOKEN_MIN_MINUTES = 210            # 3시간 30분 — T_Token 재발급 기준(3시간)보다 길게(docs/security.md의 '개발 중 토큰 규칙')
RPC_BUSY = (-2147418111, -2147417846)   # RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER (Excel이 바빠 거절)
STATUS_KINDS = (("OK", "OK"), ("오류", "오류:*"), ("이전", "이전 데이터*"), ("없음", "데이터 없음"), ("추정없음", "추정 없음"))
# 버튼 전용 쿼리의 표(모두 새로 고침에서 빠져야 함). tblA_*는 이름으로 판정
BUTTON_TABLES = ("tblClass", "tblSessions", "tblPxStore", "tblPxMetrics", "tblFlowU", "tblFin", "tblTarget", "tblEst",
                 "tblEstSnap", "tblCSL", "tblEvents", "tblSectorKRX", "tblCompany", "tblThemeAgg")
NAMES = ("시세_최근조회", "시세_상태", "전체_최근조회", "전체_상태", "업종_최근조회", "업종_상태",
         "분석_최근조회", "분석_상태", "분석코드")


class TokenGateError(RuntimeError):
    """토큰 원천의 남은 시간이 기준 미달이거나 읽을 수 없음 — 호출·열기 없이 멈춤(토큰 갱신 필요)."""


def log(msg: str) -> None:
    """시각을 붙여 즉시 출력(파이프로 넘겨도 버퍼에 갇히지 않게)."""
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def default_token_source() -> str:
    """토큰 원천 경로: 환경변수 KIS_TOKEN_SOURCE > 이 폴더 위의 통합문서 > 주 저장소(git 공통 디렉터리 기준)의 통합문서.

    워크트리에는 통합문서가 없으므로 보통 주 저장소의 `excel_dashboard/KIS_PM_Dashboard.xlsm`이 선택됩니다.
    """
    try:
        return kis_dev.token_source()
    except kis_dev.TokenError:
        pass
    r = subprocess.run(["git", "-C", DASH, "rev-parse", "--git-common-dir"], capture_output=True, text=True)
    common = r.stdout.strip()
    if r.returncode == 0 and common:
        if not os.path.isabs(common):
            common = os.path.join(DASH, common)
        root = os.path.dirname(os.path.abspath(common))
        for name in ("KIS_PM_Dashboard.xlsm", "KIS_PM_Dashboard.xlsx"):
            p = os.path.join(root, "excel_dashboard", name)
            if os.path.exists(p):
                return p
    raise TokenGateError("토큰 원천 통합문서를 찾지 못했습니다(--token-source 또는 KIS_TOKEN_SOURCE로 지정).")


def token_gate(source: Optional[str] = None, min_minutes: int = TOKEN_MIN_MINUTES, what: str = "") -> int:
    """토큰 원천을 파일 파싱으로 다시 읽어 남은 분이 기준 이상인지 확인하고 남은 분을 돌려준다.

    Args:
        source: 토큰 원천 통합문서 경로(없으면 default_token_source()).
        min_minutes: 필요한 최소 남은 시간(분). 기본 210.
        what: 로그에 붙일 설명(예: "열기 전", "[전체] 실행 전").
    Raises:
        TokenGateError: 남은 시간이 기준 미달이거나 원천을 읽지 못함.
    """
    src = source or default_token_source()
    try:
        left = kis_dev.remaining_minutes(src)
    except kis_dev.TokenError as e:
        raise TokenGateError(f"토큰 원천을 읽지 못했습니다: {e}") from None
    if left < min_minutes:
        raise TokenGateError(f"토큰 갱신 필요: 남은 {left}분 (기준 {min_minutes}분)")
    log(f"토큰 관문 통과{(' — ' + what) if what else ''}: 남은 {left}분 (기준 {min_minutes}분, 값은 표시하지 않음)")
    return left


def com_retry(fn: Callable[[], Any], timeout: float = 180.0, pause: float = 0.5) -> Any:
    """Excel이 바빠 외부 COM 호출을 거절하면(RPC_E_CALL_REJECTED 등) 잠시 뒤 다시 시도한다."""
    t0 = time.time()
    while True:
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult in RPC_BUSY and time.time() - t0 < timeout:
                time.sleep(pause)
                continue
            raise


def serial_to_dt(v: Any) -> Optional[dt.datetime]:
    """Excel 일련번호(Value2) → naive datetime. pywin32의 시간대 변환을 피하려고 날짜는 항상 Value2로 읽는다."""
    if isinstance(v, (int, float)) and v > 1:
        return dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))
    return None


def excel_pid(xl) -> Optional[int]:
    try:
        import win32process
        return win32process.GetWindowThreadProcessId(xl.Hwnd)[1]
    except Exception:  # noqa: BLE001 — 창 핸들이 없으면 PID를 모름
        return None


def new_instance(visible: bool = False):
    """사용자 Excel과 분리된 새 Excel 프로세스(DispatchEx). 매크로는 자동화 기본값(낮은 보안)으로 실행된다."""
    pythoncom.CoInitialize()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = visible
    xl.DisplayAlerts = False
    try:
        xl.AutomationSecurity = 1          # msoAutomationSecurityLow — 자동화로 연 통합문서의 매크로 실행 허용
    except pywintypes.com_error:
        pass
    return xl, excel_pid(xl)


def iter_tables(wb):
    """(시트, ListObject) 목록."""
    out = []
    for ws in wb.Worksheets:
        for lo in ws.ListObjects:
            out.append((ws, lo))
    return out


def find_table(wb, name: str):
    for _ws, lo in iter_tables(wb):
        if lo.Name.lower() == name.lower():
            return lo
    return None


def query_tables(wb) -> list[tuple[str, Any]]:
    """(표 이름, QueryTable) 목록 — Excel이 한가할 때 한 번 모아 두고 폴링에 쓴다(새로 고침 중에는 목록 조회가 거절됨)."""
    out = []
    for _ws, lo in com_retry(lambda: iter_tables(wb)):
        try:
            if com_retry(lambda: lo.SourceType) != 1:           # 1 = xlSrcRange(정적 표)
                out.append((lo.Name, com_retry(lambda: lo.QueryTable)))
        except pywintypes.com_error:
            continue
    return out


def any_refreshing(wb, qts: Optional[list] = None) -> list[str]:
    """백그라운드 새로 고침 중인 표 이름 목록. Excel이 바빠 거절하면 그 표는 '새로 고침 중'으로 본다."""
    busy = []
    for name, qt in (qts if qts is not None else query_tables(wb)):
        try:
            if qt.Refreshing:
                busy.append(name)
        except pywintypes.com_error as e:
            if e.hresult in RPC_BUSY:
                busy.append(name + "(바쁨)")
            else:
                raise
    return busy


def wait_idle(wb, timeout: float = 600.0, settle: float = 2.0, qts: Optional[list] = None) -> float:
    """모든 표의 새로 고침이 끝날 때까지 기다린다(QueryTable.Refreshing 폴링, 바쁨 거절은 '진행 중'). 걸린 초를 돌려준다."""
    t0 = time.time()
    qts = qts if qts is not None else query_tables(wb)
    quiet_since = None
    while time.time() - t0 < timeout:
        busy = any_refreshing(wb, qts)
        if busy:
            quiet_since = None
        else:
            quiet_since = quiet_since or time.time()
            if time.time() - quiet_since >= settle:
                return time.time() - t0
        time.sleep(0.5)
    raise TimeoutError(f"새로 고침이 {timeout:.0f}초 안에 끝나지 않음: {any_refreshing(wb, qts)}")


def token_row(wb) -> dict:
    """tblToken에서 env=prod 행의 expires·status·checked만(토큰 값은 읽지 않음)."""
    lo = find_table(wb, "tblToken")
    if lo is None or lo.DataBodyRange is None:
        return {}
    hdr = [str(h) for h in lo.HeaderRowRange.Value2[0]]
    out = {}
    for r in range(1, lo.ListRows.Count + 1):
        row = lo.ListRows(r).Range
        env = row.Cells(1, hdr.index("env") + 1).Value2 if "env" in hdr else None
        if env != "prod":
            continue
        for k in ("expires", "checked", "issued"):
            if k in hdr:
                out[k] = serial_to_dt(row.Cells(1, hdr.index(k) + 1).Value2)
        if "status" in hdr:
            out["status"] = row.Cells(1, hdr.index("status") + 1).Value2
        if "token" in hdr:
            v = row.Cells(1, hdr.index("token") + 1).Value2
            out["token_len"] = len(str(v)) if v else 0
        break
    return out


def open_workbook(path: str, visible: bool = False, wait_token: bool = True, token_source: Optional[str] = None,
                  min_minutes: int = TOKEN_MIN_MINUTES):
    """토큰 관문 → 새 인스턴스로 통합문서 열기 → 파일 열 때 T_Token 새로 고침이 끝날 때까지 기다림.

    Returns: (xl, wb, pid, info) — info에 열기 시간·토큰 표 상태(값 제외).
    """
    token_gate(token_source, min_minutes, f"열기 전: {os.path.basename(path)}")
    xl, pid = new_instance(visible)
    log(f"Excel 시작(PID {pid}) → 열기: {path}")
    t0 = time.time()
    try:
        wb = com_retry(lambda: xl.Workbooks.Open(path, 0, False))
    except Exception:
        quit_instance(xl, pid)
        raise
    info = {"open_seconds": round(time.time() - t0, 1)}
    if wait_token:
        time.sleep(3.0)                    # 파일 열 때 새로 고침이 시작될 시간
        info["token_wait_seconds"] = round(wait_idle(wb, timeout=300), 1)
    tr = token_row(wb)
    info["token"] = {k: (v.isoformat(sep=" ", timespec="seconds") if isinstance(v, dt.datetime) else v) for k, v in tr.items()}
    log(f"열림({info['open_seconds']}초, 토큰 표 대기 {info.get('token_wait_seconds')}초) 토큰 표: {info['token']}")
    return xl, wb, pid, info


def quit_instance(xl, pid: Optional[int], wait: float = 20.0) -> None:
    """자기 Excel만 종료. Quit 뒤에도 남으면 그 PID만 강제 종료."""
    try:
        xl.Quit()
    except Exception:  # noqa: BLE001
        pass
    if not pid:
        return
    t0 = time.time()
    while time.time() - t0 < wait:
        r = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True)
        if str(pid) not in r.stdout:
            return
        time.sleep(0.5)
    subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
    log(f"Excel PID {pid} 강제 종료(Quit 뒤에도 남아 있었음)")


def close_workbook(xl, wb, pid: Optional[int], save: bool) -> None:
    """(선택) 저장 → 닫기 → 자기 Excel 종료."""
    try:
        if save:
            wait_idle(wb, timeout=600)
            com_retry(lambda: wb.Save())
            log("저장 완료")
        com_retry(lambda: wb.Close(False))
    finally:
        quit_instance(xl, pid)


def name_values(wb) -> dict:
    """버튼이 쓰는 이름 칸 값(시각은 일시 문자열)."""
    out = {}
    for n in NAMES:
        try:
            r = wb.Names(n).RefersToRange
            v = r.Value2
            d = serial_to_dt(v) if n.endswith("최근조회") else None
            out[n] = d.isoformat(sep=" ", timespec="seconds") if d else v
        except pywintypes.com_error:
            out[n] = "<이름 없음>"
    return out


def runctl(wb) -> dict:
    """tblRunCtl 키/값(started 등 시각은 문자열)."""
    lo = find_table(wb, "tblRunCtl")
    out = {}
    if lo is None or lo.DataBodyRange is None:
        return out
    for k, v in lo.DataBodyRange.Value2:
        if k is None:
            continue
        if k in ("started", "now_override"):
            d = serial_to_dt(v)
            v = d.isoformat(sep=" ", timespec="seconds") if d else v
        out[str(k)] = v
    return out


def get_key_value(wb, table: str, key: str) -> Any:
    """정적 키/값 표의 값 칸 원값(Value2 — 날짜는 일련번호). 키가 없으면 KeyError."""
    lo = find_table(wb, table)
    if lo is None:
        raise KeyError(table)
    hdr = [str(h) for h in lo.HeaderRowRange.Value2[0]]
    kc, vc = hdr.index("키") + 1, hdr.index("값") + 1
    for r in range(1, lo.ListRows.Count + 1):
        row = lo.ListRows(r).Range
        if str(row.Cells(1, kc).Value2 or "").strip() == key:
            return row.Cells(1, vc).Value2
    raise KeyError(f"{table}에 키 {key} 없음")


def set_key_value(wb, table: str, key: str, value: Any) -> None:
    """정적 키/값 표(tblRunCtl·tblSettings)의 값 칸을 바꾼다. 날짜·일시는 Value2(일련번호)로 넣는다."""
    lo = find_table(wb, table)
    if lo is None:
        raise KeyError(table)
    hdr = [str(h) for h in lo.HeaderRowRange.Value2[0]]
    kc = hdr.index("키") + 1
    vc = hdr.index("값") + 1
    for r in range(1, lo.ListRows.Count + 1):
        row = lo.ListRows(r).Range
        if str(row.Cells(1, kc).Value2 or "").strip() == key:
            cell = row.Cells(1, vc)
            if isinstance(value, (dt.datetime, dt.date)):
                base = dt.datetime(1899, 12, 30)
                vv = value if isinstance(value, dt.datetime) else dt.datetime(value.year, value.month, value.day)
                cell.Value2 = (vv - base).total_seconds() / 86400.0
                cell.NumberFormatLocal = "yyyy-mm-dd hh:mm"
            elif value is None:
                cell.ClearContents()
            else:
                cell.Value2 = value
            return
    raise KeyError(f"{table}에 키 {key} 없음")


def _summary_of(wb, ws, lo) -> dict:
    wf = wb.Application.WorksheetFunction
    name = lo.Name
    info = {"sheet": ws.Name, "rows": lo.ListRows.Count}
    body = lo.DataBodyRange
    if name.lower() == "tbltoken" or body is None:
        return info
    hdr = [str(h) for h in lo.HeaderRowRange.Value2[0]]
    if "상태" in hdr:
        col = lo.ListColumns("상태").DataBodyRange
        info["status"] = {k: int(wf.CountIf(col, crit)) for k, crit in STATUS_KINDS}
        info["status"]["빈칸"] = int(wf.CountBlank(col))
    if "조회시각" in hdr:
        col = lo.ListColumns("조회시각").DataBodyRange
        for key, val in (("ts_max", wf.Max(col)), ("ts_min", wf.Min(col))):
            d = serial_to_dt(val)
            info[key] = d.isoformat(sep=" ", timespec="seconds") if d else None
    return info


def table_summaries(wb, names: Optional[list[str]] = None) -> dict:
    """표마다 {rows, status{OK,오류,이전,없음,추정없음}, ts_min, ts_max, sheet}. 토큰 표는 행 수만. Excel이 바쁘면 다시 시도."""
    out = {}
    for ws, lo in com_retry(lambda: iter_tables(wb)):
        name = com_retry(lambda: lo.Name)
        if names and name not in names:
            continue
        out[name] = com_retry(lambda: _summary_of(wb, ws, lo))
    return out


def first_errors(wb, table: str, top: int = 5) -> list[str]:
    """표의 `상태`가 '오류:'·'이전 데이터'로 시작하는 앞쪽 몇 개(종목코드 포함)."""
    lo = find_table(wb, table)
    if lo is None or lo.DataBodyRange is None:
        return []
    hdr = [str(h) for h in lo.HeaderRowRange.Value2[0]]
    if "상태" not in hdr:
        return []
    si = hdr.index("상태")
    ci = hdr.index("종목코드") if "종목코드" in hdr else None
    vals = lo.DataBodyRange.Value2 if lo.ListRows.Count <= 20000 else None
    if vals is None:
        return ["(행이 많아 생략 — 파일 파싱으로 확인)"]
    out = []
    for row in vals:
        s = str(row[si] or "")
        if s.startswith("오류") or s.startswith("이전 데이터"):
            out.append((f"{row[ci]} " if ci is not None else "") + s[:160])
            if len(out) >= top:
                break
    return out


class Watchdog:
    """오래 걸리는 동기 COM 호출(매크로 실행)을 지켜보며 진행 시간을 찍고, 한도를 넘으면 자기 Excel PID만 강제 종료한다.

    강제 종료되면 저장하지 않은 변경은 사라지지만 디스크의 통합문서는 그대로다(실제 파일 보호).
    """

    def __init__(self, label: str, pid: Optional[int], limit_seconds: float, every: float = 60.0):
        self.label, self.pid, self.limit, self.every = label, pid, limit_seconds, every
        self._done = threading.Event()
        self.killed = False
        self._t = threading.Thread(target=self._run, daemon=True)

    def _run(self):
        t0 = time.time()
        while not self._done.wait(self.every):
            el = time.time() - t0
            log(f"  … {self.label} 진행 중 {el / 60:.1f}분")
            if el > self.limit and self.pid:
                self.killed = True
                log(f"  ✗ {self.label} 한도 {self.limit / 60:.0f}분 초과 — 자기 Excel(PID {self.pid}) 강제 종료")
                subprocess.run(["taskkill", "/PID", str(self.pid), "/F"], capture_output=True)
                return

    def __enter__(self):
        self._t.start()
        return self

    def __exit__(self, *exc):
        self._done.set()
        return False
