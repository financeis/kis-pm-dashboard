"""개발·검증 전용 KIS 도우미 (조회 전용, 토큰 발급 없음).

- 토큰 원천 통합문서의 `_sys!tblToken`을 Excel로 열지 않고 파일에서 직접 읽습니다
  (Excel로 열면 '파일 열 때 새로 고침'으로 T_Token이 돌아 토큰이 발급될 수 있음).
- 토큰·앱키·시크릿·계좌번호는 어떤 출력·로그·파일에도 쓰지 않습니다. 확인은 남은 시간(분)만.
- 남은 유효시간이 MIN_MINUTES(3시간 30분) 미만이면 호출 없이 TokenError로 멈춥니다. 갱신은 오케스트레이터만 합니다.
- 토큰 원천 경로: 인자 > 환경변수 KIS_TOKEN_SOURCE > 이 폴더 위의 KIS_PM_Dashboard.xlsm > .xlsx
- 주문·계좌 경로(`/trading/`)와 `/uapi/` 밖의 경로(토큰 발급 등)는 호출을 거부합니다.

    python excel_dashboard/tools/kis_dev.py              # 토큰 원천 경로·남은 시간(분)·상태만 표시
    python excel_dashboard/tools/kis_dev.py --min 210    # 기준 미달이면 종료 코드 2

    from kis_dev import KisClient
    kis = KisClient()                                    # 기준 미달이면 TokenError
    body, cont = kis.get("/uapi/domestic-stock/v1/quotations/inquire-price", "FHKST01010100",
                         {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": "005930"})
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
DASH = os.path.dirname(HERE)
sys.path.insert(0, DASH)
from xlsx_tables import read_records  # noqa: E402

MIN_MINUTES = 210                     # 3시간 30분 (T_Token 재발급 기준 3시간보다 길게)
DEFAULT_CFG = os.path.expanduser("~/KIS/config/kis_devlp.yaml")


class TokenError(RuntimeError):
    """토큰 원천을 읽을 수 없거나 남은 시간이 기준 미달."""


class KisError(RuntimeError):
    """KIS 응답 오류 (rt_cd != '0')."""


def token_source(path: str | None = None) -> str:
    if path:
        return os.path.abspath(path)
    env = os.environ.get("KIS_TOKEN_SOURCE")
    if env:
        return os.path.abspath(env)
    for name in ("KIS_PM_Dashboard.xlsm", "KIS_PM_Dashboard.xlsx"):
        p = os.path.join(DASH, name)
        if os.path.exists(p):
            return p
    raise TokenError("토큰 원천 통합문서를 찾을 수 없습니다(환경변수 KIS_TOKEN_SOURCE로 지정하세요).")


def read_token(path: str | None = None) -> tuple[str, dt.datetime, str]:
    """(토큰, 만료 시각, 상태) — 토큰 값은 메모리에서만 쓰고 출력하지 마세요."""
    src = token_source(path)
    try:
        recs = read_records(src, "tblToken")
    except (KeyError, OSError) as e:
        raise TokenError(f"토큰 표를 읽지 못했습니다: {type(e).__name__}") from None
    for rec in recs:
        if rec.get("env") == "prod" and rec.get("token"):
            exp = rec.get("expires")
            if not isinstance(exp, dt.datetime):
                raise TokenError("tblToken의 expires가 날짜 형식이 아닙니다.")
            return str(rec["token"]), exp, str(rec.get("status") or "")
    raise TokenError("토큰 원천에 유효한 토큰 행이 없습니다.")


def remaining_minutes(path: str | None = None, now: dt.datetime | None = None) -> int:
    _tok, exp, _st = read_token(path)
    return int(((exp - (now or dt.datetime.now())).total_seconds()) // 60)


def require_token(path: str | None = None, min_minutes: int = MIN_MINUTES) -> tuple[str, dt.datetime]:
    """남은 시간이 기준 이상인 토큰을 돌려줌. 미달이면 호출 없이 TokenError."""
    tok, exp, _st = read_token(path)
    left = int(((exp - dt.datetime.now()).total_seconds()) // 60)
    if left < min_minutes:
        raise TokenError(f"토큰 갱신 필요: 남은 {left}분 (기준 {min_minutes}분) — 오케스트레이터에게 알리세요.")
    return tok, exp


def load_cfg(path: str | None = None) -> dict:
    import yaml
    with open(path or DEFAULT_CFG, encoding="utf-8") as fh:
        cfg = yaml.safe_load(fh) or {}
    missing = [k for k in ("my_app", "my_sec", "prod") if not str(cfg.get(k) or "").strip()]
    if missing:
        raise KisError(f"kis_devlp.yaml에 필수 항목이 없습니다: {', '.join(missing)}")
    return cfg


class KisClient:
    """조회 전용 GET 클라이언트. 초당 호출 수 제한·EGW00201 재시도 포함. 비밀값은 repr에도 나오지 않음."""

    def __init__(self, token_path: str | None = None, cfg_path: str | None = None,
                 rate_per_sec: float = 8.0, min_minutes: int = MIN_MINUTES):
        import requests
        self._token, self.expires = require_token(token_path, min_minutes)
        cfg = load_cfg(cfg_path)
        self._app = str(cfg["my_app"]).strip()
        self._sec = str(cfg["my_sec"]).strip()
        self.base = str(cfg["prod"]).strip().rstrip("/")
        self._interval = 1.0 / rate_per_sec
        self._last = 0.0
        self._s = requests.Session()
        self.calls = 0

    def __repr__(self):
        return f"<KisClient calls={self.calls}>"

    def _throttle(self):
        wait = self._interval - (time.monotonic() - self._last)
        if wait > 0:
            time.sleep(wait)
        self._last = time.monotonic()

    def get(self, path: str, tr_id: str, params: dict, tr_cont: str = "", check: bool = True) -> tuple[dict, str]:
        """(응답 본문, 응답 헤더 tr_cont). 연속 조회는 tr_cont='N'으로 다시 호출."""
        if not path.startswith("/uapi/") or "/trading/" in path:
            raise ValueError(f"조회 전용 도우미가 허용하지 않는 경로입니다: {path}")
        headers = {
            "content-type": "application/json; charset=utf-8",
            "authorization": "Bearer " + self._token,
            "appkey": self._app, "appsecret": self._sec,
            "tr_id": tr_id, "custtype": "P",
        }
        if tr_cont:
            headers["tr_cont"] = tr_cont
        body, cont, net_fail = {}, "", 0
        for attempt in range(6):
            self._throttle()
            try:
                r = self._s.get(self.base + path, headers=headers, params=params, timeout=20)
                body = r.json()
                cont = r.headers.get("tr_cont", "")
            except Exception as e:  # 네트워크·JSON 오류: 1회만 재시도 (메시지에 헤더를 담지 않음)
                net_fail += 1
                if net_fail <= 1:
                    continue
                raise KisError(f"{tr_id} 네트워크/응답 오류: {type(e).__name__}") from None
            self.calls += 1
            if body.get("msg_cd") == "EGW00201" and attempt < 5:
                time.sleep(0.3 * 2 ** attempt)
                continue
            break
        if check and str(body.get("rt_cd")) != "0":
            raise KisError(f"{tr_id} {body.get('msg_cd')} {body.get('msg1')}")
        return body, cont


def main():
    ap = argparse.ArgumentParser(description="토큰 원천 상태 확인 (토큰 값은 출력하지 않음)")
    ap.add_argument("--source", help="토큰 원천 통합문서 경로")
    ap.add_argument("--min", type=int, default=MIN_MINUTES, help="필요한 최소 남은 시간(분)")
    a = ap.parse_args()
    try:
        src = token_source(a.source)
        _tok, exp, status = read_token(src)
    except TokenError as e:
        print(f"토큰 확인 실패: {e}")
        sys.exit(2)
    left = int(((exp - dt.datetime.now()).total_seconds()) // 60)
    print(f"토큰 원천: {src} | 만료: {exp:%Y-%m-%d %H:%M} | 남은 시간: {left}분 | 상태: {status}")
    sys.exit(0 if left >= a.min else 2)


if __name__ == "__main__":
    main()
