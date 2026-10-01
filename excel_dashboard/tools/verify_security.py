# -*- coding: utf-8 -*-
"""보안 검사 (T24 — spec §8 V8): 통합문서 안에 앱키·시크릿·계좌번호·HTS ID가 없고, 접근토큰은 `_sys` 시트에만 있는지.

검사 방법 (비밀값은 메모리에서만 쓰고 출력하지 않음 — 결과에는 항목 이름·길이·발견 위치(파트 이름)만 나옴)
1. 비밀값 목록: `~/KIS/config/kis_devlp.yaml`(또는 --cfg)의 앱키·시크릿(실전·모의), HTS ID, 계좌번호들(6자 미만 값은
   우연 일치가 많아 제외하고 보고). 접근토큰은 검사 대상 통합문서의 `tblToken`을 파일 파싱으로 읽는다(Excel로 열지 않음).
2. 통합문서 zip의 모든 파트를 풀어 UTF-8·UTF-16LE·CP949 바이트로 검색한다. Power Query 원본이 들어 있는 `customXml`의
   DataMashup(base64 → 이진 → 안쪽 zip)도 풀어서 안쪽 파트까지 검색한다. VBA(`vbaProject.bin`)도 이진 그대로 검색한다.
3. 접근토큰: 공유 문자열 표(sharedStrings)의 항목 번호를 찾아 그 번호를 참조하는 시트가 `_sys`뿐인지, 다른 파트(쿼리 정의·
   연결·VBA·다른 시트 인라인 문자열)에 토큰 문자열이 없는지 확인한다.
4. 함께 검사할 파일(--extra): 이력 CSV 등 통합문서 밖 파일에 비밀값·토큰이 없는지.
5. `.gitignore`: `git check-ignore`로 `.xlsm`·`backup/`·`history/` 경로가 커밋 대상에서 빠지는지.

    python excel_dashboard/tools/verify_security.py --workbook <경로.xlsm> [--extra <파일> …] [--json 결과.json]
종료 코드: 0 통과 / 1 발견(실패) / 3 인자·파일 오류
"""
from __future__ import annotations

import argparse
import base64
import io
import json
import os
import re
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
sys.path.insert(0, DASH)
sys.path.insert(0, HERE)
from xlsx_tables import Workbook  # noqa: E402

DEFAULT_CFG = os.path.expanduser("~/KIS/config/kis_devlp.yaml")
# kis_devlp.yaml에서 비밀로 취급하는 키(값이 비었거나 자리 표시 값이면 건너뜀)
SECRET_KEYS = ("my_app", "my_sec", "paper_app", "paper_sec", "my_htsid", "my_acct_stock", "my_acct_future",
               "my_paper_stock", "my_paper_future", "my_acct")
MIN_LEN = 6


def load_secrets(cfg_path: str) -> dict[str, str]:
    """설정 파일의 비밀값(메모리에서만). 반환: {키: 값}. 6자 미만 값은 따로 표시해 검색에서 뺀다."""
    import yaml
    with open(cfg_path, encoding="utf-8") as fh:
        cfg = yaml.safe_load(fh) or {}
    out = {}
    for k in SECRET_KEYS:
        v = str(cfg.get(k) or "").strip()
        if not v or any(ord(ch) > 127 for ch in v):      # 빈 값·한글 자리 표시 값(예: 설명 문구)은 비밀이 아님
            continue
        out[k] = v
    return out


def read_token(path: str) -> str | None:
    """검사 대상 통합문서의 tblToken에서 env=prod 토큰(메모리에서만)."""
    try:
        with Workbook(path) as wb:
            cols, rows = wb.read_table("tblToken")
    except (KeyError, OSError):
        return None
    for r in rows:
        rec = dict(zip(cols, r))
        if rec.get("env") == "prod" and rec.get("token"):
            return str(rec["token"])
    return None


def encodings(value: str) -> list[bytes]:
    out = [value.encode("utf-8"), value.encode("utf-16-le")]
    try:
        out.append(value.encode("cp949"))
    except UnicodeEncodeError:
        pass
    return list(dict.fromkeys(out))


def mashup_parts(blob: bytes) -> dict[str, bytes]:
    """customXml 파트 안의 DataMashup(base64)을 풀어 안쪽 zip 파트들을 돌려준다(없으면 빈 dict)."""
    m = re.search(rb"<DataMashup[^>]*>([^<]+)</DataMashup>", blob)
    if not m:
        try:
            text = blob.decode("utf-16")
        except UnicodeDecodeError:
            return {}
        m2 = re.search(r"<DataMashup[^>]*>([^<]+)</DataMashup>", text)
        if not m2:
            return {}
        raw = base64.b64decode(m2.group(1))
    else:
        raw = base64.b64decode(m.group(1))
    out = {"<DataMashup 이진>": raw}
    # 형식: 버전(4바이트) + 패키지 길이(4바이트, little endian) + 패키지(zip) + 권한·메타데이터…
    size = int.from_bytes(raw[4:8], "little") if len(raw) >= 8 else 0
    pkg = raw[8:8 + size] if size and raw[8:12] == b"PK\x03\x04" else b""
    if pkg:
        try:
            with zipfile.ZipFile(io.BytesIO(pkg)) as z:
                for n in z.namelist():
                    out["<DataMashup>/" + n] = z.read(n)
        except zipfile.BadZipFile:
            out["<DataMashup 패키지 해제 실패>"] = b""
    return out


def all_parts(path: str) -> dict[str, bytes]:
    parts = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            data = z.read(n)
            parts[n] = data
            if n.startswith("customXml/") and (b"DataMashup" in data or b"D\x00a\x00t\x00a\x00M" in data):
                parts.update({f"{n}::{k}": v for k, v in mashup_parts(data).items()})
    return parts


def token_locations(path: str, parts: dict[str, bytes], token: str) -> dict:
    """토큰 문자열이 있는 파트와, 공유 문자열 번호를 참조하는 시트."""
    found = [n for n, d in parts.items() if any(e in d for e in encodings(token))]
    res = {"parts_with_token": found, "sheets_referencing": []}
    ss = parts.get("xl/sharedStrings.xml")
    idx = None
    if ss and token.encode("utf-8") in ss:
        items = re.findall(rb"<si>(.*?)</si>", ss, flags=re.S)
        for i, it in enumerate(items):
            if token.encode("utf-8") in it:
                idx = i
                break
    res["shared_index_found"] = idx is not None
    # 시트 파트 → 이름 (workbook.xml + rels)
    wbx = parts.get("xl/workbook.xml", b"").decode("utf-8", "ignore")
    rels = parts.get("xl/_rels/workbook.xml.rels", b"").decode("utf-8", "ignore")
    rid_target = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid_target.update({a: b for b, a in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    part_name = {}
    for nm, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        tgt = rid_target.get(rid, "")
        tgt = tgt.lstrip("/")
        part_name[tgt if tgt.startswith("xl/") else "xl/" + tgt] = nm
    if idx is not None:
        pat = re.compile(rb'<c [^>]*t="s"[^>]*>(?:<f[^<]*</f>)?<v>' + str(idx).encode() + rb"</v>")
        for n, d in parts.items():
            if n.startswith("xl/worksheets/sheet") and pat.search(d):
                res["sheets_referencing"].append(part_name.get(n, n))
    res["inline_sheets"] = [part_name.get(n, n) for n in found if n.startswith("xl/worksheets/")]
    return res


def scan_file(path: str, secrets: dict[str, str]) -> dict[str, list[str]]:
    """일반 파일(이력 CSV 등)에서 비밀값 검색 → {키: [파일]}."""
    with open(path, "rb") as fh:
        data = fh.read()
    hits = {}
    for k, v in secrets.items():
        if len(v) >= MIN_LEN and any(e in data for e in encodings(v)):
            hits.setdefault(k, []).append(os.path.basename(path))
    return hits


def gitignore_check() -> dict[str, bool]:
    """대표 경로가 git에서 무시되는지(git check-ignore). True = 무시됨(커밋 대상 아님)."""
    samples = ["excel_dashboard/KIS_PM_Dashboard.xlsm", "excel_dashboard/backup/KIS_PM_Dashboard_20261001_000000.xlsm",
               "excel_dashboard/backup/KIS_PM_Dashboard_20261001_000000.xlsx", "excel_dashboard/history/est_snap.csv",
               "excel_dashboard/KIS_PM_Dashboard.xlsx", "excel_dashboard/.build_tmp/KIS_PM_Dashboard.xlsm"]
    root = subprocess.run(["git", "-C", DASH, "rev-parse", "--show-toplevel"], capture_output=True, text=True).stdout.strip()
    out = {}
    for s in samples:
        r = subprocess.run(["git", "-C", root, "check-ignore", "-q", s], capture_output=True)
        out[s] = r.returncode == 0
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="통합문서 비밀값·토큰 위치 검사 (값은 출력하지 않음)")
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--cfg", default=DEFAULT_CFG)
    ap.add_argument("--extra", nargs="*", default=[], help="함께 검사할 파일(이력 CSV 등)")
    ap.add_argument("--json", default=None)
    a = ap.parse_args()
    if not os.path.exists(a.workbook):
        print(f"파일 없음: {a.workbook}")
        return 3
    secrets = load_secrets(a.cfg)
    short = sorted(k for k, v in secrets.items() if len(v) < MIN_LEN)
    token = read_token(a.workbook)
    parts = all_parts(a.workbook)
    res = {"workbook": a.workbook, "parts": len(parts), "secret_keys": {k: len(v) for k, v in secrets.items()},
           "short_skipped": short, "secret_hits": {}, "token_present": token is not None}
    for k, v in secrets.items():
        if len(v) < MIN_LEN:
            continue
        hits = [n for n, d in parts.items() if any(e in d for e in encodings(v))]
        if hits:
            res["secret_hits"][k] = hits
    if token:
        res["token"] = token_locations(a.workbook, parts, token)
        tl = res["token"]
        allowed = {"xl/sharedStrings.xml"}
        bad_parts = [p for p in tl["parts_with_token"] if p not in allowed and not p.startswith("xl/worksheets/")]
        bad_sheets = [s for s in tl["sheets_referencing"] + tl["inline_sheets"] if s != "_sys"]
        res["token_ok"] = not bad_parts and not bad_sheets and ("_sys" in tl["sheets_referencing"] or "_sys" in tl["inline_sheets"])
        res["token_bad_parts"] = bad_parts
        res["token_bad_sheets"] = bad_sheets
    res["extra"] = {}
    for f in a.extra:
        if os.path.exists(f):
            h = scan_file(f, secrets)
            if token:
                with open(f, "rb") as fh:
                    if any(e in fh.read() for e in encodings(token)):
                        h.setdefault("token", []).append(os.path.basename(f))
            res["extra"][os.path.basename(f)] = h
    res["gitignore"] = gitignore_check()
    ok = (not res["secret_hits"] and res.get("token_ok", True) and all(res["gitignore"].values())
          and not any(res["extra"].values()))
    res["pass"] = ok
    print(f"V8 보안 검사: {os.path.basename(a.workbook)} — 파트 {res['parts']}개(DataMashup 안쪽 포함)")
    print(f"  비밀값 항목 {len(secrets)}개(길이만: {res['secret_keys']}), 6자 미만 제외: {short or '없음'}")
    print(f"  비밀값 발견: {res['secret_hits'] or '없음'}")
    if token:
        tl = res["token"]
        print(f"  토큰: 있는 파트 {tl['parts_with_token']} · 참조 시트 {tl['sheets_referencing'] + tl['inline_sheets']} "
              f"→ {'_sys에만 있음' if res['token_ok'] else '다른 곳에도 있음 ✗'}")
    else:
        print("  토큰: 통합문서에 토큰 없음")
    print(f"  함께 검사한 파일: {res['extra'] or '없음'}")
    print(f"  .gitignore(무시됨=True): {res['gitignore']}")
    print(f"  판정: {'통과' if ok else '실패'}")
    if a.json:
        with open(a.json, "w", encoding="utf-8") as fh:
            json.dump(res, fh, ensure_ascii=False, indent=1)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
