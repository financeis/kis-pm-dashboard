"""기존 대시보드 통합문서에서 사용자 데이터·이력·토큰 캐시를 읽어 '이관 계획'을 만든다 (2026-10-01, 규칙은 docs/business-rules.md의 '사용자 데이터 보존').

- 원본은 **Excel로 열지 않는다**: 열면 '파일 열 때 새로 고침'이 걸린 T_Token이 돌아 토큰이 새로 발급될 수 있다(알림톡).
  xlsx_tables(표준 라이브러리 zip/XML)로 파일을 직접 읽는다. 이 모듈은 COM을 쓰지 않는다.
- 결과 MigrationPlan을 빌더(build_dashboard.py)가 받아 새 통합문서에 쓴다:
    입력표(tblTrades·tblWatch·tblMacro·tblHolidays·tblOverride)는 그 표에 직접, tblSettings는 키별 병합,
    자기참조 저장소 표(tblPxStore·tblSessions·tblEstSnap·tblCSL·tblFin, 있으면 tblFlowU·tblTarget·tblEst·tblEvents·tblSectorKRX)는
    정적 표 `<표이름>Seed`로 — Power Query가 채우는 표에는 행을 직접 쓸 수 없어서, 자기참조 쿼리가 자기 표가 비었을 때
    Seed를 읽어 이어 쓰게 한다. tblToken은 메모리로만 넘겨 새 통합문서의 토큰 표에 넣는다(불필요한 재발급 방지).
- 토큰 값은 TokenRecord 안에만 있고 repr·로그·파일 어디에도 나오지 않는다(남은 분만 보고).
- 파일 작업 도우미(원본 선택·잠금 확인·백업 복사·.xlsx 원본 옮기기)도 여기 둔다. 어떤 파일도 백업 없이 지우지 않는다.

    from migrate import select_source, read_plan, backup_copy
    src = select_source(out_path, migrate_from=None)          # --migrate-from > 출력 위치 .xlsm > 같은 위치 .xlsx > None
    plan = read_plan(src, known_sheets=SHEETS, history_csv=history_csv_path(out_path))
"""
from __future__ import annotations

import csv
import dataclasses
import datetime as dt
import hashlib
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET

from xlsx_tables import Workbook


class MigrationError(RuntimeError):
    """원본을 고를 수 없거나 읽을 수 없음 — 빌드를 멈춘다(새로 만들려면 --no-migrate)."""


# ---------------------------------------------------------------------------------------------------------------
# 이관 대상 (표 이름·열은 docs/contracts.md의 표 계약 그대로)
# ---------------------------------------------------------------------------------------------------------------
# 입력표: 이관하는 열 = 사용자가 입력하는 열. 수식 열(빌더가 다시 넣음)은 FORMULA_COLUMNS — 원본 값(계산 결과)은 버린다.
# 입력 열에 사용자가 수식을 넣었으면 파일에 저장된 계산 값으로 옮긴다(파일 파싱은 수식을 다시 계산하지 않음).
INPUT_TABLES: dict[str, list[str]] = {
    "tblTrades": ["일자", "종목코드", "구분", "수량", "단가", "수수료", "세금", "전략", "매매근거", "목표가", "손절가", "메모"],
    "tblWatch": ["종목코드", "그룹", "관심가", "목표가", "투자포인트"],
    "tblMacro": ["구분", "이름", "시장코드", "심볼", "순서"],
    "tblHolidays": ["휴장일", "설명"],
    "tblOverride": ["종목코드", "대회편입", "대테마", "세부테마", "NICS 대분류", "NICS 업종", "NICS 세부", "메모"],
}
FORMULA_COLUMNS: dict[str, list[str]] = {
    "tblTrades": ["종목명", "금액", "확인"],
    "tblWatch": ["종목명"],
}
CODE_COLUMN = "종목코드"
SETTINGS_TABLE = "tblSettings"
SETTINGS_COLUMNS = ["키", "항목", "값", "설명"]
TOKEN_TABLE = "tblToken"
TOKEN_COLUMNS = ["env", "token", "expires", "issued", "key_sig", "status", "checked"]

# 자기참조 저장소 표 → 원본에 표가 없을 때 빈 Seed의 열 구성(각 쿼리의 Cols와 같음). 원본에 표가 있으면 원본 열 그대로 옮긴다.
STORE_COLUMNS: dict[str, list[str]] = {
    "tblSessions": ["일자", "오늘개장", "개장확인일", "상태", "조회시각"],
    "tblPxStore": ["종목코드", "일자", "시가", "고가", "저가", "종가", "거래량", "거래대금억", "상태", "조회시각"],
    "tblEstSnap": ["일자", "종목코드", "FwdEPS", "종가", "FwdPER", "추정일", "상태", "조회시각"],
    "tblCSL": ["종목코드", "기준일", "신용잔고율", "신용잔고율1M변화", "공매도비중5일", "대차잔고1M변화율", "상태", "조회시각"],
    "tblFin": ["종목코드", "결산년월", "공개기준일", "매출", "영업이익", "순이익", "EPS", "BPS", "ROE", "상태", "조회시각"],
}
REQUIRED_SEEDS = ("tblSessions", "tblPxStore", "tblEstSnap", "tblCSL", "tblFin")     # 늘 만드는 Seed (원본에 없으면 빈 표)
OPTIONAL_SEEDS = ("tblFlowU", "tblTarget", "tblEst", "tblEvents", "tblSectorKRX")    # 원본에 행이 있을 때만 만드는 Seed
STORE_TABLES = REQUIRED_SEEDS + OPTIONAL_SEEDS

# 없으면 대시보드 통합문서가 아니라고 보고 멈춘다(잘못 고른 파일을 이관한 뒤 백업 폴더로 옮겨 버리는 일 방지)
CORE_TABLES = ("tblSettings", "tblTrades")
OPTIONAL_INPUTS = ("tblWatch", "tblMacro", "tblHolidays")   # 없으면 경고 후 기본값(빌더)
HISTORY_DIR = "history"
HISTORY_CSV = "est_snap.csv"
BACKUP_DIR = "backup"


# ---------------------------------------------------------------------------------------------------------------
# 결과 형식
# ---------------------------------------------------------------------------------------------------------------
@dataclasses.dataclass
class TableData:
    """이관할 표 하나. rows는 빈 행(모든 칸 공란)을 뺀 행, 값 순서 = columns."""

    name: str
    columns: list[str]
    rows: list[list]
    origin: str = "통합문서"          # '통합문서' 또는 '이력 CSV'
    source_rows: int = 0               # 원본의 데이터 행 수(빈 행 제외) — 행 수 대조 보고용
    notes: list[str] = dataclasses.field(default_factory=list)


class TokenRecord:
    """원본 tblToken의 prod 행. 토큰 값은 .token으로만 꺼내며 repr·str에는 남은 시간만 나온다."""

    __slots__ = ("_token", "expires", "issued", "key_sig", "status")

    def __init__(self, token: str, expires: dt.datetime, issued: dt.datetime | None, key_sig: str | None, status: str):
        self._token = token
        self.expires = expires
        self.issued = issued
        self.key_sig = key_sig
        self.status = status

    @property
    def token(self) -> str:
        return self._token

    def minutes_left(self, now: dt.datetime | None = None) -> int:
        return int(((self.expires - (now or dt.datetime.now())).total_seconds()) // 60)

    def __repr__(self) -> str:
        return f"<TokenRecord 남은 {self.minutes_left()}분>"

    __str__ = __repr__


@dataclasses.dataclass
class MigrationPlan:
    """원본 한 개에서 읽은 이관 계획. 빌더가 이 값만 보고 새 통합문서를 채운다."""

    source: str
    tables: dict[str, TableData]                 # 입력표 + 저장소 표(원본에 있던 것, 0행 포함) + 이력 CSV로 복원한 tblEstSnap
    settings: list[tuple]                        # 원본 순서 (키, 항목, 값, 설명)
    token: TokenRecord | None
    token_note: str
    sheets: list[str]
    unknown_sheets: list[str]
    missing_tables: list[str]
    warnings: list[str]

    @property
    def settings_count(self) -> int:
        return len(self.settings)

    def rows(self, name: str) -> list[list] | None:
        t = self.tables.get(name)
        return None if t is None else t.rows


# ---------------------------------------------------------------------------------------------------------------
# 값 도우미
# ---------------------------------------------------------------------------------------------------------------
def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def norm_code(v):
    """종목코드는 텍스트로: 숫자로 저장된 코드(앞 0 빠짐)는 6자리로 채운다. 문자열은 그대로(영문 포함 코드 유지)."""
    if v is None:
        return None
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return f"{v:06d}" if 0 <= v < 1_000_000 else str(v)
    if isinstance(v, float):
        if v.is_integer() and 0 <= v < 1_000_000:
            return f"{int(v):06d}"
        return str(v)
    if isinstance(v, (dt.date, dt.datetime)):
        return v.isoformat()
    return str(v)


def _to_datetime(v, wb: Workbook | None = None) -> dt.datetime | None:
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if wb is not None:
            return wb.serial_to_datetime(float(v))
        return dt.datetime(1899, 12, 30) + dt.timedelta(milliseconds=round(float(v) * 86400000))
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(s, fmt)
            except ValueError:
                continue
    return None


# ---------------------------------------------------------------------------------------------------------------
# 파일 작업 (원본 선택·잠금·백업·옮기기)
# ---------------------------------------------------------------------------------------------------------------
def history_csv_path(out_path: str) -> str:
    """이력 CSV 위치: 통합문서와 같은 폴더의 history\\est_snap.csv (VBA [전체]가 쓰는 곳과 같음)."""
    return os.path.join(os.path.dirname(os.path.abspath(out_path)), HISTORY_DIR, HISTORY_CSV)


def backup_dir_for(out_path: str) -> str:
    """백업 폴더: 출력 통합문서와 같은 폴더의 backup\\ (기본 출력이면 excel_dashboard\\backup\\)."""
    return os.path.join(os.path.dirname(os.path.abspath(out_path)), BACKUP_DIR)


def candidate_sources(out_path: str) -> list[str]:
    """--migrate-from이 없을 때의 이관 원본 후보: 출력 위치의 .xlsm, 그다음 같은 위치·같은 이름의 .xlsx."""
    stem = os.path.splitext(os.path.abspath(out_path))[0]
    return [stem + ".xlsm", stem + ".xlsx"]


def select_source(out_path: str, migrate_from: str | None = None) -> str | None:
    """이관 원본 경로. --migrate-from이 있으면 그 파일(없으면 MigrationError), 없으면 후보 중 처음 있는 것, 없으면 None(처음 빌드)."""
    if migrate_from:
        p = os.path.abspath(migrate_from)
        if not os.path.isfile(p):
            raise MigrationError(f"--migrate-from 파일이 없습니다: {p}")
        return p
    for p in candidate_sources(out_path):
        if os.path.isfile(p):
            return p
    return None


def is_locked(path: str) -> bool:
    """다른 프로그램(보통 Excel)이 파일을 쓰기 잠금 중이면 True. 아무것도 쓰지 않고 열었다 닫기만 한다."""
    if not os.path.exists(path):
        return False
    try:
        with open(path, "r+b"):
            pass
        return False
    except PermissionError:
        return True
    except OSError:
        return True


def owner_file(path: str) -> str:
    """Excel이 파일을 열 때 같은 폴더에 만드는 소유자 파일(~$이름) 경로."""
    d, f = os.path.split(os.path.abspath(path))
    return os.path.join(d, "~$" + f)


def sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def _unique(path: str) -> str:
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    k = 1
    while os.path.exists(f"{stem}_{k}{ext}"):
        k += 1
    return f"{stem}_{k}{ext}"


def backup_copy(path: str, backup_dir: str, stamp: str) -> str:
    """backup\\<이름>_YYYYMMDD_HHMMSS.<확장자>로 복사(수정 시각 유지)하고 내용이 같은지 해시로 확인한 뒤 그 경로를 돌려준다.
    같은 이름이 이미 있으면 _1, _2 …를 붙인다(기존 백업을 덮지 않음)."""
    os.makedirs(backup_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    dst = _unique(os.path.join(backup_dir, f"{stem}_{stamp}{ext}"))
    shutil.copy2(path, dst)
    if sha256(path) != sha256(dst):
        raise MigrationError(f"백업 확인 실패(내용이 다름): {dst}")
    return dst


def move_into_backup(path: str, backup_path: str, backup_dir: str, stamp: str) -> str:
    """이관이 끝난 .xlsx 원본을 백업 폴더로 '옮긴다'. 빌드 전에 만든 백업 복사본과 내용이 같으면 그 복사본이 곧
    옮긴 파일이므로 원래 자리의 파일만 지운다. 빌드 중에 원본이 바뀌었으면(해시 다름) 새 이름으로 백업 폴더에 옮긴다.
    어느 경우든 백업 폴더에 같은 내용이 있는 것을 확인한 뒤에만 원래 자리에서 없앤다. 돌려주는 값 = 로그 한 줄."""
    if not os.path.isfile(path):
        return f"원본이 이미 없음(옮길 것 없음): {path}"
    if backup_path and os.path.isfile(backup_path) and sha256(path) == sha256(backup_path):
        os.remove(path)
        return f".xlsx 원본을 백업 폴더로 옮김: {backup_path} (빌드 전 복사본과 내용 같음 확인 → 원래 자리 파일 정리)"
    os.makedirs(backup_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    dst = _unique(os.path.join(backup_dir, f"{stem}_{stamp}_moved{ext}"))
    shutil.move(path, dst)
    return f".xlsx 원본을 백업 폴더로 옮김: {dst} (빌드 중 원본이 바뀌어 새 이름으로)"


# ---------------------------------------------------------------------------------------------------------------
# 읽기
# ---------------------------------------------------------------------------------------------------------------
def _open(path: str) -> Workbook:
    try:
        return Workbook(path)
    except PermissionError as e:
        raise MigrationError(f"원본을 읽을 수 없습니다(다른 프로그램이 잠금): {path}") from e
    except (zipfile.BadZipFile, KeyError, ET.ParseError, OSError, ValueError) as e:
        raise MigrationError(f"원본을 읽을 수 없습니다({type(e).__name__}: {e}): {path}") from e


def _find(tables: dict, name: str) -> str | None:
    return next((k for k in tables if k.lower() == name.lower()), None)


def _read(wb: Workbook, name: str) -> tuple[list[str], list[list]]:
    try:
        return wb.read_table(name)
    except (KeyError, ET.ParseError, ValueError, zipfile.BadZipFile) as e:
        raise MigrationError(f"원본 표 {name}을(를) 읽지 못했습니다({type(e).__name__}: {e})") from e


def _input_table(wb: Workbook, name: str, warnings: list[str]) -> TableData:
    cols, rows = _read(wb, name)
    want = INPUT_TABLES[name]
    known = set(want) | set(FORMULA_COLUMNS.get(name, []))
    idx = {c: i for i, c in enumerate(cols)}
    missing = [c for c in want if c not in idx]
    extra = [c for c in cols if c not in known]
    t = TableData(name, list(want), [])
    if missing:
        t.notes.append(f"원본에 없는 열(빈칸으로 둠): {', '.join(missing)}")
    if extra:
        warnings.append(f"{name}: 알려지지 않은 열 {', '.join(extra)}은(는) 이관하지 않음(백업본에 남음)")
    n_codes = 0
    for r in rows:
        vals = [r[idx[c]] if c in idx else None for c in want]
        if all(is_blank(v) for v in vals):
            continue
        if CODE_COLUMN in want:
            j = want.index(CODE_COLUMN)
            if vals[j] is not None and not isinstance(vals[j], str):
                n_codes += 1
            vals[j] = norm_code(vals[j])
        t.rows.append(vals)
    t.source_rows = len(t.rows)
    if n_codes:
        t.notes.append(f"숫자로 저장된 종목코드 {n_codes}개를 6자리 텍스트로 맞춤")
    return t


def _store_table(wb: Workbook, name: str) -> TableData:
    cols, rows = _read(wb, name)
    t = TableData(name, list(cols), [])
    ci = cols.index(CODE_COLUMN) if CODE_COLUMN in cols else -1
    for r in rows:
        if all(is_blank(v) for v in r):
            continue
        r = list(r)
        if ci >= 0 and r[ci] is not None and not isinstance(r[ci], str):
            r[ci] = norm_code(r[ci])
        t.rows.append(r)
    t.source_rows = len(t.rows)
    return t


def _settings(wb: Workbook) -> list[tuple]:
    cols, rows = _read(wb, SETTINGS_TABLE)
    idx = {c: i for i, c in enumerate(cols)}
    if "키" not in idx or "값" not in idx:
        raise MigrationError(f"원본 {SETTINGS_TABLE}에 키/값 열이 없습니다(열: {', '.join(cols)})")
    out, seen = [], set()
    for r in rows:
        key = r[idx["키"]]
        if is_blank(key):
            continue
        key = str(key).strip()
        if key in seen:              # 같은 키가 여러 행이면 첫 행(Power Query fnSettings와 같은 규칙)
            continue
        seen.add(key)
        out.append((key, r[idx["항목"]] if "항목" in idx else None, r[idx["값"]], r[idx["설명"]] if "설명" in idx else None))
    return out


def _token(wb: Workbook, tables: dict) -> tuple[TokenRecord | None, str]:
    key = _find(tables, TOKEN_TABLE)
    if key is None:
        return None, "원본에 tblToken 없음"
    try:
        cols, rows = wb.read_table(key)
    except (KeyError, ET.ParseError, ValueError) as e:
        return None, f"tblToken을 읽지 못함({type(e).__name__})"
    idx = {c: i for i, c in enumerate(cols)}
    if "token" not in idx or "expires" not in idx:
        return None, "tblToken에 token/expires 열 없음"

    def get(r, c):
        return r[idx[c]] if c in idx else None

    for r in rows:
        env = get(r, "env")
        tok = get(r, "token")
        if (env is not None and str(env).strip().lower() != "prod") or is_blank(tok):
            continue
        exp = _to_datetime(get(r, "expires"), wb)
        if exp is None:
            return None, "tblToken의 expires가 날짜가 아님"
        sig = get(r, "key_sig")
        sig = None if is_blank(sig) else (str(int(sig)) if isinstance(sig, float) and sig.is_integer() else str(sig).strip())
        return TokenRecord(str(tok).strip(), exp, _to_datetime(get(r, "issued"), wb), sig, str(get(r, "status") or "")), "prod 토큰 있음"
    return None, "tblToken에 prod 토큰 행 없음"


def read_est_snap_csv(path: str) -> TableData | None:
    """history\\est_snap.csv(VBA [전체]가 저장: UTF-8 BOM, 머리글 = 표 열 이름, 날짜 yyyy-mm-dd, 시각 yyyy-mm-dd hh:nn:ss,
    소수점 '.')를 tblEstSnap 행으로 읽는다. 파일이 없거나 비면 None."""
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8-sig", newline="") as fh:
        raw = [r for r in csv.reader(fh) if any(x.strip() for x in r)]
    if len(raw) < 2:
        return None
    cols = [c.strip() for c in raw[0]]
    date_re = re.compile(r"^\d{4}-\d{2}-\d{2}$")
    dt_re = re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?$")
    num_re = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")
    rows = []
    for r in raw[1:]:
        r = (r + [""] * len(cols))[:len(cols)]
        out = []
        for c, v in zip(cols, r):
            s = v.strip()
            if s == "":
                out.append(None)
            elif c == CODE_COLUMN or c == "상태":
                out.append(s)
            elif dt_re.match(s):
                out.append(_to_datetime(s))
            elif date_re.match(s):
                out.append(dt.datetime.strptime(s, "%Y-%m-%d"))
            elif num_re.match(s):
                f = float(s)
                out.append(int(f) if f.is_integer() and "." not in s and "e" not in s.lower() else f)
            else:
                out.append(s)
        if not all(is_blank(v) for v in out):
            rows.append(out)
    if not rows:
        return None
    t = TableData("tblEstSnap", cols, rows, origin="이력 CSV", source_rows=len(rows))
    t.notes.append(f"이력 CSV에서 복원: {path}")
    return t


def read_plan(path: str, known_sheets, history_csv: str | None = None) -> MigrationPlan:
    """원본 통합문서를 파일로 읽어 이관 계획을 만든다(Excel 미사용).

    Args:
        path: 원본(.xlsm/.xlsx).
        known_sheets: 빌더가 아는 시트 이름(이 밖의 시트는 사용자 추가 시트로 보고 경고 — 이관하지 않음).
        history_csv: 원본에 스냅샷 행이 없을 때 tblEstSnap을 복원할 이력 CSV(없으면 None).

    Raises:
        MigrationError: 파일을 읽을 수 없거나(zip 손상·잠금), 대시보드 통합문서의 핵심 표(tblSettings·tblTrades)가 없음.
    """
    wb = _open(path)
    try:
        try:
            tables = wb.tables()
        except (ET.ParseError, KeyError, ValueError, zipfile.BadZipFile) as e:
            raise MigrationError(f"원본의 표 목록을 읽지 못했습니다({type(e).__name__}: {e})") from e
        missing_core = [n for n in CORE_TABLES if _find(tables, n) is None]
        if missing_core:
            raise MigrationError(f"원본에 대시보드의 핵심 표({', '.join(missing_core)})가 없습니다 — 대시보드 통합문서가 맞는지 "
                                 f"확인하세요(새로 만들려면 --no-migrate): {path}")
        warnings: list[str] = []
        sheets = list(wb.sheets)
        known = set(known_sheets)
        unknown = [s for s in sheets if s not in known]
        out: dict[str, TableData] = {}
        missing: list[str] = []
        for name in INPUT_TABLES:
            key = _find(tables, name)
            if key is None:
                missing.append(name)
                if name in OPTIONAL_INPUTS:
                    warnings.append(f"원본에 {name}이(가) 없어 기본값으로 만듭니다")
                continue
            out[name] = _input_table(wb, key, warnings)
        for name in STORE_TABLES:
            key = _find(tables, name)
            t = _store_table(wb, key) if key is not None else None
            # 자기 표가 없거나 비었는데 지난 빌드의 Seed가 남아 있으면(빌더가 옮기지 못해 남겨 둔 경우) 그 행을 옮긴다
            seed_key = _find(tables, name + "Seed")
            if (t is None or not t.rows) and seed_key is not None:
                s = _store_table(wb, seed_key)
                if s.rows:
                    s.name = name
                    s.origin = f"통합문서({name}Seed)"
                    t = s
            if t is not None:
                out[name] = t
        settings = _settings(wb)
        token, token_note = _token(wb, tables)
    finally:
        wb.close()
    snap = out.get("tblEstSnap")
    if (snap is None or not snap.rows) and history_csv:
        try:
            csv_t = read_est_snap_csv(history_csv)
        except (OSError, csv.Error, UnicodeDecodeError) as e:
            csv_t = None
            warnings.append(f"이력 CSV를 읽지 못함({type(e).__name__}): {history_csv}")
        if csv_t is not None:
            csv_t.source_rows = len(csv_t.rows)
            out["tblEstSnap"] = csv_t
    return MigrationPlan(path, out, settings, token, token_note, sheets, unknown, missing, warnings)


def read_token_only(path: str) -> tuple[TokenRecord | None, str]:
    """--no-migrate일 때 토큰 캐시만 읽는다(원본이 손상됐으면 (None, 사유) — 빌드는 계속)."""
    try:
        wb = Workbook(path)
    except (PermissionError, zipfile.BadZipFile, KeyError, ET.ParseError, OSError, ValueError) as e:
        return None, f"토큰을 읽을 수 없음({type(e).__name__})"
    try:
        try:
            tables = wb.tables()
        except (ET.ParseError, KeyError, ValueError, zipfile.BadZipFile) as e:
            return None, f"토큰을 읽을 수 없음({type(e).__name__})"
        return _token(wb, tables)
    finally:
        wb.close()
