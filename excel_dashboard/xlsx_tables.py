"""통합문서(.xlsx/.xlsm)의 표(ListObject)를 Excel 없이 파일에서 직접 읽는 도구 (표준 라이브러리만 사용).

- Excel로 열면 '파일 열 때 새로 고침'이 걸린 쿼리(T_Token 등)가 돌아 토큰이 발급될 수 있으므로,
  토큰 확인·이관·검증은 이 모듈로 파일(zip/XML)을 직접 파싱합니다.
- 값 변환: 공유·인라인 문자열 → str, 논리값 → bool, 오류 → '#N/A' 같은 문자열, 숫자 → int/float,
  날짜 서식이 적용된 숫자 → datetime, 빈 셀 → None.
- 큰 표(가격 저장소 등)도 읽을 수 있게 시트 XML을 스트리밍으로 읽습니다.

    from xlsx_tables import list_tables, read_table, read_records, read_sheet
    cols, rows = read_table("KIS_PM_Dashboard.xlsx", "tblTrades")
    recs = read_records("KIS_PM_Dashboard.xlsx", "tblSettings")   # [{"키": ..., "값": ...}, ...]
    cols, rows = read_sheet("수집기업_valuesearch.xlsx", "Sheet2")   # 표가 아닌 일반 시트
"""
from __future__ import annotations

import datetime as dt
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_M = "{%s}" % _NS_MAIN
_RID = "{%s}id" % _NS_REL

# 기본 제공 날짜·시간 서식 번호 (한국어 등 동아시아 로캘의 27~36·50~58 포함)
_BUILTIN_DATE_IDS = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))


def _is_date_format(code: str) -> bool:
    """사용자 지정 서식 코드가 날짜·시간 서식인지 판정."""
    s = re.sub(r'"[^"]*"', "", code)                     # 따옴표 안 문자열
    s = re.sub(r"\\.", "", s)                            # 이스케이프 문자
    s = re.sub(r"\[(?![hms]+\])[^\]]*\]", "", s, flags=re.I)   # [색10]·[$-412] 등 (경과 시간 [h]·[mm]·[ss]는 유지)
    s = s.split(";")[0]
    return bool(re.search(r"[dmyhs]", s, re.I)) and s.strip().lower() != "general"


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _split_ref(ref: str) -> tuple[int, int]:
    """'AB12' → (열 번호, 행 번호)"""
    m = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", ref)
    if not m:
        raise ValueError(f"셀 주소 형식 오류: {ref}")
    return _col_index(m.group(1)), int(m.group(2))


class Workbook:
    """zip 패키지를 한 번 열어 시트·표·공유 문자열·날짜 스타일을 해석."""

    def __init__(self, path: str):
        self.path = path
        try:
            self.zip = zipfile.ZipFile(path)
        except PermissionError as e:
            raise PermissionError(f"통합문서를 읽을 수 없습니다(다른 프로그램이 잠금): {path}") from e
        self.names = set(self.zip.namelist())
        root = ET.fromstring(self.zip.read("xl/workbook.xml"))
        pr = root.find(_M + "workbookPr")
        self.date1904 = pr is not None and pr.get("date1904") in ("1", "true")
        rels = self._rels("xl/workbook.xml")
        self.sheets: dict[str, str] = {}          # 시트 이름 → 파트 경로
        for s in root.iter(_M + "sheet"):
            rel = rels.get(s.get(_RID))
            if rel:
                self.sheets[s.get("name")] = rel[1]
        self.shared = self._shared_strings()
        self.date_styles = self._date_styles()
        self._tables: dict[str, dict] | None = None

    def close(self):
        self.zip.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # ------------------------------------------------------------------ 패키지 해석
    def _rels(self, part: str) -> dict[str, tuple[str, str]]:
        """파트의 관계 파일 → {Id: (관계 종류, 대상 파트 경로)}"""
        d, f = posixpath.split(part)
        rp = posixpath.join(d, "_rels", f + ".rels")
        out: dict[str, tuple[str, str]] = {}
        if rp not in self.names:
            return out
        for r in ET.fromstring(self.zip.read(rp)):
            if r.get("TargetMode") == "External":
                continue
            tgt = r.get("Target", "")
            full = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(d, tgt))
            out[r.get("Id")] = (r.get("Type", ""), full)
        return out

    def _shared_strings(self) -> list[str]:
        part = "xl/sharedStrings.xml"
        if part not in self.names:
            return []
        out = []
        for si in ET.fromstring(self.zip.read(part)):
            out.append(self._rich_text(si))
        return out

    @staticmethod
    def _rich_text(node) -> str:
        """<si>/<is> 안의 텍스트 (윗주 rPh 제외)."""
        parts = []
        for ch in node:
            if ch.tag == _M + "t":
                parts.append(ch.text or "")
            elif ch.tag == _M + "r":
                t = ch.find(_M + "t")
                parts.append(t.text or "" if t is not None else "")
        return "".join(parts)

    def _date_styles(self) -> set[int]:
        part = "xl/styles.xml"
        if part not in self.names:
            return set()
        root = ET.fromstring(self.zip.read(part))
        custom = {}
        fmts = root.find(_M + "numFmts")
        if fmts is not None:
            for f in fmts:
                custom[int(f.get("numFmtId"))] = f.get("formatCode", "")
        out = set()
        xfs = root.find(_M + "cellXfs")
        if xfs is not None:
            for i, xf in enumerate(xfs):
                fid = int(xf.get("numFmtId", "0"))
                if fid in _BUILTIN_DATE_IDS or (fid in custom and _is_date_format(custom[fid])):
                    out.add(i)
        return out

    # ------------------------------------------------------------------ 표
    def tables(self) -> dict[str, dict]:
        """{표 이름: {sheet, part, ref, header_rows, totals_rows, columns}}"""
        if self._tables is not None:
            return self._tables
        out = {}
        for sheet, part in self.sheets.items():
            for typ, tpart in self._rels(part).values():
                if not typ.endswith("/table") or tpart not in self.names:
                    continue
                t = ET.fromstring(self.zip.read(tpart))
                name = t.get("displayName") or t.get("name")
                cols_node = t.find(_M + "tableColumns")
                cols = [c.get("name") for c in cols_node] if cols_node is not None else []
                out[name] = {
                    "sheet": sheet, "part": part, "ref": t.get("ref"),
                    "header_rows": int(t.get("headerRowCount", "1")),
                    "totals_rows": int(t.get("totalsRowCount", "0")),
                    "columns": cols,
                }
        self._tables = out
        return out

    def _cell_value(self, c):
        t = c.get("t", "n")
        if t == "inlineStr":
            node = c.find(_M + "is")
            return self._rich_text(node) if node is not None else None
        v = c.find(_M + "v")
        if v is None or v.text is None:
            return None
        x = v.text
        if t == "s":
            return self.shared[int(x)]
        if t == "str" or t == "e":
            return x
        if t == "b":
            return x == "1"
        if t == "d":
            return dt.datetime.fromisoformat(x)
        num = float(x)
        s = c.get("s")
        if s is not None and int(s) in self.date_styles:
            return self.serial_to_datetime(num)
        if re.fullmatch(r"-?\d+", x):
            return int(x)
        return num

    def serial_to_datetime(self, serial: float) -> dt.datetime:
        base = dt.datetime(1904, 1, 1) if self.date1904 else dt.datetime(1899, 12, 30)
        return base + dt.timedelta(milliseconds=round(serial * 86400000))

    def read_table(self, name: str) -> tuple[list[str], list[list]]:
        """표의 (열 이름 목록, 데이터 행 목록). 이름은 대소문자를 구분하지 않음."""
        tabs = self.tables()
        key = next((k for k in tabs if k.lower() == name.lower()), None)
        if key is None:
            raise KeyError(f"표를 찾을 수 없습니다: {name}")
        info = tabs[key]
        first, last = info["ref"].split(":") if ":" in info["ref"] else (info["ref"], info["ref"])
        c1, r1 = _split_ref(first)
        c2, r2 = _split_ref(last)
        top = r1 + info["header_rows"]
        bottom = r2 - info["totals_rows"]
        width = c2 - c1 + 1
        cols = info["columns"] or [f"열{i + 1}" for i in range(width)]
        rows: dict[int, list] = {}
        with self.zip.open(info["part"]) as fh:
            row_no = 0
            for _ev, el in ET.iterparse(fh, events=("end",)):
                if el.tag != _M + "row":
                    continue
                row_no = int(el.get("r")) if el.get("r") else row_no + 1
                if top <= row_no <= bottom:
                    vals = [None] * width
                    col_no = 0
                    for c in el.findall(_M + "c"):
                        col_no = _split_ref(c.get("r"))[0] if c.get("r") else col_no + 1
                        if c1 <= col_no <= c2:
                            vals[col_no - c1] = self._cell_value(c)
                    rows[row_no] = vals
                el.clear()
                if row_no > bottom:
                    break
        data = [rows.get(r, [None] * width) for r in range(top, bottom + 1)]
        return cols, data

    def read_sheet(self, sheet: str, header_row: int = 1) -> tuple[list[str], list[list]]:
        """표가 아닌 일반 시트를 머리글 행 기준으로 읽음 (예: VALUESearch 내보내기 파일).
        머리글이 빈 열은 버리고, 머리글 아래의 완전히 빈 행은 건너뜁니다."""
        part = self.sheets.get(sheet)
        if part is None:
            raise KeyError(f"시트를 찾을 수 없습니다: {sheet}")
        header: dict[int, str] = {}
        body: list[dict[int, object]] = []
        with self.zip.open(part) as fh:
            row_no = 0
            for _ev, el in ET.iterparse(fh, events=("end",)):
                if el.tag != _M + "row":
                    continue
                row_no = int(el.get("r")) if el.get("r") else row_no + 1
                if row_no >= header_row:
                    vals: dict[int, object] = {}
                    col_no = 0
                    for c in el.findall(_M + "c"):
                        col_no = _split_ref(c.get("r"))[0] if c.get("r") else col_no + 1
                        v = self._cell_value(c)
                        if v is not None and v != "":
                            vals[col_no] = v
                    if row_no == header_row:
                        header = {k: str(v).strip() for k, v in vals.items()}
                    elif vals:
                        body.append(vals)
                el.clear()
        order = sorted(header)
        return [header[k] for k in order], [[r.get(k) for k in order] for r in body]


def list_tables(path: str) -> dict[str, dict]:
    with Workbook(path) as wb:
        return wb.tables()


def read_table(path: str, name: str) -> tuple[list[str], list[list]]:
    with Workbook(path) as wb:
        return wb.read_table(name)


def read_records(path: str, name: str) -> list[dict]:
    cols, rows = read_table(path, name)
    return [dict(zip(cols, r)) for r in rows]


def read_sheet(path: str, sheet: str, header_row: int = 1) -> tuple[list[str], list[list]]:
    with Workbook(path) as wb:
        return wb.read_sheet(sheet, header_row)
