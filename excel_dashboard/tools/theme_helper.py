"""대회 종목 기본 테마표 작성 보조 도구 — 검증(check)과 분류용 참고 목록 추출(extract).

인코딩 UTF-8 · 작성 2026-10-01 · 대상 파일 excel_dashboard/data/themes_base.csv (빌더가 정적 표 `tblThemeBase`로 넣음)

기본 테마표 형식 (UTF-8 BOM, 대회 명단과 같은 행 순서 = 시가총액 내림차순):
    종목코드, 종목명, 대테마, 세부테마, 근거
  - 종목코드는 6자리 텍스트(앞자리 0 유지, 최근 상장 종목은 '0126Z0'처럼 영문 포함) — 숫자로 바꾸지 않습니다.
  - 대회 명단 전 종목에 대테마 1개 + 세부테마 1개, 세부테마는 항상 같은 대테마 아래, 근거 = 주요 제품·사업 한 줄.

하위 명령:
  check    기본 테마표를 검증합니다. 오류가 하나라도 있거나 파일이 없으면 종료 코드 1, 통과하면 0.
           오류: 인코딩(UTF-8 BOM)·머리글, 빈 칸·앞뒤 공백·줄바꿈, 종목코드 형식(6자리 영문 대문자·숫자)·중복,
                 명단과 종목코드 집합 불일치(누락·초과), 종목명 불일치, 행 순서 불일치, 세부테마가 둘 이상의 대테마 아래.
           안내(종료 코드에 영향 없음): 대테마·세부테마 수가 규모 기준(대테마 10~20, 세부테마 40~80) 밖,
                 1종목 세부테마 목록, 긴 근거.
  extract  대회 명단 + VALUESearch 내보내기(주요상품·KSIC·NICS·기업집단) + KIS 테마 마스터 소속을 합쳐
           분류 검토용 목록을 출력 폴더에 씁니다(커밋하지 않는 작업 자료):
             theme_extract.csv     종목별 한 줄 (UTF-8 BOM)
             theme_by_nics.txt     NICS 산업별로 묶은 읽기용 목록
             kis_theme_counts.txt  대회 종목이 속한 KIS 테마별 종목 수
           KIS 테마 마스터(theme_code.mst.zip, 인증 불필요)는 --theme-zip이 없으면 내려받아 출력 폴더에 저장합니다.

사용 예:
    python excel_dashboard/tools/theme_helper.py check
    python excel_dashboard/tools/theme_helper.py extract --out-dir %TEMP%/themes ^
        --valuesearch C:/path/수집기업_valuesearch.xlsx
"""
from __future__ import annotations

import argparse
import csv
import io
import os
import re
import sys
from collections import Counter, OrderedDict
from typing import Optional

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
REPO = os.path.dirname(DASH)

DEFAULT_CONTEST = os.path.join(DASH, "data", "contest_universe_20260930.csv")
DEFAULT_THEMES = os.path.join(DASH, "data", "themes_base.csv")

THEME_COLUMNS = ["종목코드", "종목명", "대테마", "세부테마", "근거"]
CODE_RE = re.compile(r"[0-9A-Z]{6}")
UTF8_BOM = b"\xef\xbb\xbf"

# 규모 기준 (대테마·세부테마 개수, 조정 가능 — 벗어나면 안내만 함)
MAJOR_RANGE = (10, 20)
MINOR_RANGE = (40, 80)
RATIONALE_WARN_LEN = 60          # 근거가 이보다 길면 안내 (한 줄 요약 유지)


# ---------------------------------------------------------------- 공통
def read_csv_bom(path: str) -> tuple[bool, list[str], list[list[str]]]:
    """CSV를 텍스트 그대로 읽음 → (BOM 여부, 머리글, 행 목록). 모든 값은 str (종목코드 앞자리 0 보존)."""
    with open(path, "rb") as fh:
        raw = fh.read()
    has_bom = raw.startswith(UTF8_BOM)
    text = raw[len(UTF8_BOM):].decode("utf-8") if has_bom else raw.decode("utf-8")
    rows = list(csv.reader(io.StringIO(text, newline="")))
    if not rows:
        return has_bom, [], []
    return has_bom, rows[0], [r for r in rows[1:] if any(c != "" for c in r)]


def load_contest(path: str) -> list[dict]:
    """대회 명단 → [{종목코드, 종목명, 시장, 시가총액_0930_억, ...}] (명단 순서 유지)."""
    _bom, header, rows = read_csv_bom(path)
    if "종목코드" not in header or "종목명" not in header:
        raise ValueError(f"대회 명단 머리글에 종목코드·종목명이 없습니다: {path}")
    return [dict(zip(header, r)) for r in rows]


# ---------------------------------------------------------------- check
def check_themes(themes_path: str, contest_path: str) -> int:
    """기본 테마표 검증 → 종료 코드 (0 통과, 1 오류)."""
    errors: list[str] = []
    notes: list[str] = []

    if not os.path.exists(themes_path):
        print(f"[오류] 기본 테마표가 없습니다: {themes_path}")
        return 1
    contest = load_contest(contest_path)
    contest_codes = [r["종목코드"] for r in contest]
    contest_name = {r["종목코드"]: r["종목명"] for r in contest}

    try:
        has_bom, header, rows = read_csv_bom(themes_path)
    except UnicodeDecodeError as e:
        print(f"[오류] UTF-8로 읽을 수 없습니다: {e}")
        return 1
    if not has_bom:
        errors.append("인코딩: UTF-8 BOM이 없습니다 (Excel에서 한글이 깨지지 않게 BOM 필요)")
    if header != THEME_COLUMNS:
        errors.append(f"머리글이 다릅니다: {header} (기대 {THEME_COLUMNS})")
        _report(errors, notes)
        return 1

    # 행 단위 검사
    seen: Counter = Counter()
    for i, r in enumerate(rows, start=2):           # 2 = 머리글 다음 줄 (Excel 행 번호와 같음)
        if len(r) != len(THEME_COLUMNS):
            errors.append(f"{i}행: 열 수 {len(r)} (기대 {len(THEME_COLUMNS)}) — 쉼표가 든 값은 따옴표로 감싸야 합니다")
            continue
        code, name, major, minor, why = r
        for col, val in zip(THEME_COLUMNS, r):
            if val.strip() == "":
                errors.append(f"{i}행 {code}: {col} 빈 칸")
            elif val != val.strip():
                errors.append(f"{i}행 {code}: {col} 앞뒤 공백 '{val}'")
            if "\n" in val or "\r" in val:
                errors.append(f"{i}행 {code}: {col}에 줄바꿈")
        if not CODE_RE.fullmatch(code):
            errors.append(f"{i}행: 종목코드 형식 오류 '{code}' (6자리 영문 대문자·숫자 텍스트)")
        seen[code] += 1
        if code in contest_name and name != contest_name[code]:
            errors.append(f"{i}행 {code}: 종목명 '{name}' ≠ 명단 '{contest_name[code]}'")
        if len(why) > RATIONALE_WARN_LEN:
            notes.append(f"근거가 깁니다({len(why)}자) {code} {name}: {why}")

    dups = sorted(c for c, n in seen.items() if n > 1)
    if dups:
        errors.append(f"종목코드 중복 {len(dups)}개: {', '.join(dups)}")
    missing = [c for c in contest_codes if c not in seen]
    extra = sorted(c for c in seen if c not in contest_name)
    if missing:
        errors.append(f"명단 종목 누락 {len(missing)}개: {', '.join(missing[:30])}{' …' if len(missing) > 30 else ''}")
    if extra:
        errors.append(f"명단에 없는 종목 {len(extra)}개: {', '.join(extra[:30])}{' …' if len(extra) > 30 else ''}")
    order = [r[0] for r in rows if r]
    if not missing and not extra and not dups and order != contest_codes:
        first = next(k for k, (a, b) in enumerate(zip(order, contest_codes)) if a != b)
        errors.append(f"행 순서가 명단(시가총액 내림차순)과 다릅니다: {first + 2}행 {order[first]} (명단 {contest_codes[first]})")

    # 계층: 세부테마 → 대테마 하나
    valid = [r for r in rows if len(r) == len(THEME_COLUMNS)]
    parents: dict[str, set] = {}
    for r in valid:
        parents.setdefault(r[3], set()).add(r[2])
    for minor, majors in sorted(parents.items()):
        if len(majors) > 1:
            errors.append(f"세부테마 '{minor}'가 대테마 여러 개 아래에 있습니다: {sorted(majors)}")
    major_names = {r[2] for r in valid}
    clash = sorted(m for m, ps in parents.items() if m in major_names and ps != {m})
    if clash:
        notes.append(f"다른 대테마와 이름이 같은 세부테마: {clash}")

    # 집계
    major_count = Counter(r[2] for r in valid)
    minor_count = Counter((r[2], r[3]) for r in valid)
    coverage = len([c for c in contest_codes if c in seen])
    print(f"커버리지: {coverage}/{len(contest_codes)} ({coverage / len(contest_codes):.1%})  행 수 {len(rows)}")
    print(f"대테마 {len(major_count)}개 · 세부테마 {len(parents)}개")
    for major, n in sorted(major_count.items(), key=lambda kv: (-kv[1], kv[0])):
        subs = sorted(((mi, k) for (ma, mi), k in minor_count.items() if ma == major), key=lambda kv: (-kv[1], kv[0]))
        print(f"  {major} {n}: " + ", ".join(f"{mi} {k}" for mi, k in subs))
    singles = sorted(f"{ma} > {mi}" for (ma, mi), k in minor_count.items() if k == 1)
    print(f"1종목 세부테마 {len(singles)}개" + (": " + "; ".join(singles) if singles else ""))
    lo, hi = MAJOR_RANGE
    if not lo <= len(major_count) <= hi:
        notes.append(f"대테마 수 {len(major_count)}개가 규모 기준 {lo}~{hi}개 밖입니다(조정 가능 기준 — 사유 확인)")
    lo, hi = MINOR_RANGE
    if not lo <= len(parents) <= hi:
        notes.append(f"세부테마 수 {len(parents)}개가 규모 기준 {lo}~{hi}개 밖입니다(조정 가능 기준 — 사유 확인)")

    _report(errors, notes)
    return 1 if errors else 0


def _report(errors: list[str], notes: list[str]) -> None:
    for n in notes:
        print(f"[안내] {n}")
    for e in errors:
        print(f"[오류] {e}")
    print("결과: " + (f"실패 (오류 {len(errors)}건)" if errors else "통과"))


# ---------------------------------------------------------------- extract
THEME_MASTER_URL = "https://new.real.download.dws.co.kr/common/master/theme_code.mst.zip"
VS_SHEET = "Sheet2"
VS_FIELDS = OrderedDict([            # 출력 열 이름 ← VALUESearch 머리글
    ("NICS", "691300.NICS 산업분류"),
    ("KSIC세분류", "691440.KSIC-세분류(11차)"),
    ("KSIC세세분류", "691450.KSIC-세세분류(11차)"),
    ("주식업종", "691240.주식업종"),
    ("기업집단", "691280.소속기업집단"),
    ("주요상품", "691230.주요상품"),
])
EXTRACT_COLUMNS = ["종목코드", "종목명", "시장", "시가총액억", "비고", "VS"] + list(VS_FIELDS) + ["KIS테마"]


def strip_code_prefix(text: str) -> str:
    """'N45301.반도체 및 반도체장비' → '반도체 및 반도체장비' (분류 코드 접두어 제거)."""
    return re.sub(r"^[A-Z0-9]+\.", "", text or "").strip()


def load_valuesearch(path: str) -> dict[str, dict]:
    """VALUESearch 내보내기 → {6자리 종목코드: {출력 열: 값}}. 종목코드 'A005930'·'A0126Z0'의 'A'를 뗍니다."""
    sys.path.insert(0, DASH)
    from xlsx_tables import read_sheet   # 표준 라이브러리만 쓰는 저장소 리더 (openpyxl 불필요)

    cols, rows = read_sheet(path, VS_SHEET)
    idx = {c: i for i, c in enumerate(cols)}
    missing = [h for h in ["종목코드", *VS_FIELDS.values()] if h not in idx]
    if missing:
        raise ValueError(f"VALUESearch 머리글이 없습니다: {missing}")
    out: dict[str, dict] = {}
    for r in rows:
        raw = str(r[idx["종목코드"]] or "").strip()
        code = raw[1:] if len(raw) == 7 and raw.startswith("A") else raw
        if not CODE_RE.fullmatch(code):
            continue
        rec = {}
        for col, head in VS_FIELDS.items():
            val = r[idx[head]]
            text = "" if val is None else re.sub(r"\s+", " ", str(val)).strip()
            rec[col] = text if col in ("주요상품", "기업집단") else strip_code_prefix(text)   # 분류 열만 코드 접두어 제거
        out[code] = rec
    return out


def load_theme_master(zip_path: str) -> dict[str, list[str]]:
    """KIS 테마 마스터 → {종목코드: [테마명, ...]}.
    한 줄 = 고정폭 52바이트(cp949): 테마코드 3 + 테마명(공백 채움) + 종목코드 6 + 공백 3 (stocks_info/theme_code.py와 같은 배치)."""
    import zipfile

    with zipfile.ZipFile(zip_path) as zf:
        data = zf.read(zf.namelist()[0])
    out: dict[str, list[str]] = {}
    for line in data.split(b"\n"):
        line = line.rstrip(b"\r")
        if len(line.strip()) == 0:
            continue
        code = line[-9:].decode("ascii", errors="replace").strip()
        name = line[3:-9].decode("cp949", errors="replace").strip()
        out.setdefault(code, [])
        if name not in out[code]:
            out[code].append(name)
    return out


def download_theme_master(out_dir: str) -> str:
    """theme_code.mst.zip 내려받기 (인증 불필요) → 저장 경로."""
    import urllib.request

    dest = os.path.join(out_dir, "theme_code.mst.zip")
    urllib.request.urlretrieve(THEME_MASTER_URL, dest)
    return dest


def extract(out_dir: str, contest_path: str, vs_path: str, theme_zip: Optional[str]) -> int:
    """분류 검토용 목록 작성 → 종료 코드."""
    os.makedirs(out_dir, exist_ok=True)
    if not os.path.exists(vs_path):
        print(f"[오류] VALUESearch 파일이 없습니다: {vs_path} (--valuesearch로 경로 지정)")
        return 1
    contest = load_contest(contest_path)
    vs = load_valuesearch(vs_path)
    zip_path = theme_zip or download_theme_master(out_dir)
    themes = load_theme_master(zip_path)

    recs = []
    for r in contest:
        code = r["종목코드"]
        v = vs.get(code)
        rec = {"종목코드": code, "종목명": r["종목명"], "시장": r.get("시장", ""),
               "시가총액억": r.get("시가총액_0930_억", ""), "비고": r.get("비고", ""), "VS": "Y" if v else "N"}
        rec.update(v or {k: "" for k in VS_FIELDS})
        rec["KIS테마"] = "|".join(themes.get(code, []))
        recs.append(rec)

    path_csv = os.path.join(out_dir, "theme_extract.csv")
    with open(path_csv, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=EXTRACT_COLUMNS)
        w.writeheader()
        w.writerows(recs)

    groups: dict[str, list[dict]] = OrderedDict()
    for rec in sorted(recs, key=lambda x: (x["NICS"] or "~" + x["VS"], -float(x["시가총액억"] or 0))):
        label = rec["NICS"] or ("(NICS 없음)" if rec["VS"] == "Y" else "(VALUESearch 없음)")
        groups.setdefault(label, []).append(rec)
    path_txt = os.path.join(out_dir, "theme_by_nics.txt")
    with open(path_txt, "w", encoding="utf-8") as fh:
        for nics, items in groups.items():
            fh.write(f"## {nics} ({len(items)})\n")
            for x in items:
                mcap = f"{float(x['시가총액억'] or 0):,.0f}억"
                fh.write(f"{x['종목코드']} {x['종목명']} [{x['시장']} {mcap}] KSIC={x['KSIC세세분류'] or x['KSIC세분류']}"
                         f" | 그룹={x['기업집단']} | 상품={x['주요상품']} | KIS={x['KIS테마']}\n")
            fh.write("\n")

    tc = Counter(t for rec in recs for t in rec["KIS테마"].split("|") if t)
    path_cnt = os.path.join(out_dir, "kis_theme_counts.txt")
    with open(path_cnt, "w", encoding="utf-8") as fh:
        for t, n in tc.most_common():
            fh.write(f"{n}\t{t}\n")

    no_vs = [f"{x['종목코드']} {x['종목명']}" for x in recs if x["VS"] == "N"]
    no_kis = sum(1 for x in recs if not x["KIS테마"])
    print(f"대회 종목 {len(recs)}개 · VALUESearch 매칭 {len(recs) - len(no_vs)}개 · KIS 테마 소속 없음 {no_kis}개")
    if no_vs:
        print("VALUESearch에 없는 종목: " + ", ".join(no_vs))
    print(f"출력: {path_csv}\n      {path_txt}\n      {path_cnt}")
    return 0


# ---------------------------------------------------------------- 인자·진입점
def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="대회 종목 기본 테마표(themes_base.csv) 검증·참고 목록 추출 (spec R3)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("check", help="기본 테마표 검증 (오류 시 종료 코드 1)")
    c.add_argument("--themes", default=DEFAULT_THEMES, help="기본 테마표 경로 (기본 %(default)s)")
    c.add_argument("--contest", default=DEFAULT_CONTEST, help="대회 명단 경로 (기본 %(default)s)")
    e = sub.add_parser("extract", help="분류 검토용 참고 목록 작성 (명단 + VALUESearch + KIS 테마 마스터)")
    e.add_argument("--out-dir", required=True, help="출력 폴더 (작업 자료 — 저장소 밖 임시 폴더 권장)")
    e.add_argument("--contest", default=DEFAULT_CONTEST, help="대회 명단 경로 (기본 %(default)s)")
    e.add_argument("--valuesearch", default=os.path.join(REPO, "수집기업_valuesearch.xlsx"),
                   help="VALUESearch 내보내기 경로 (읽기 전용, 기본 %(default)s)")
    e.add_argument("--theme-zip", help="KIS 테마 마스터 zip 경로 (없으면 내려받음)")
    return ap.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):        # Windows 콘솔(cp949)에서도 한글·기호가 깨지지 않게
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    args = parse_args(argv)
    if args.cmd == "check":
        return check_themes(args.themes, args.contest)
    if args.cmd == "extract":
        return extract(args.out_dir, args.contest, args.valuesearch, args.theme_zip)
    return 2


if __name__ == "__main__":
    sys.exit(main())
