// fnMultiQuote : 관심종목(멀티종목) 시세조회 [국내주식-205] FHKST11300006 — 종목 목록을 30종목씩 묶어 조회 (2026-10-01 T5)
//  - 가격 저장소(T_PxStore)의 [시세] 경로가 쓰는 범용 함수. 다른 쿼리도 현재가·기준가가 여러 종목 필요하면 이 함수를 쓴다
//    (종목당 1회인 inquire-price 대비 호출 수 1/30 — 대회 527종목 18회).
//  - 입력: 종목코드 목록. 앞뒤 공백 제거·대문자, 숫자만 6자리 미만이면 앞 0 채움(fnClassify와 같은 정규화).
//    6자리 [0-9A-Z]가 아니면 호출하지 않고 오류 행. 빈 값은 버리고, 중복은 처음 것만. 출력 행 순서 = 입력의 처음 등장 순서.
//    영문이 섞인 코드(0126Z0·0035S0 등 대회 10종목)도 그대로 조회된다(2026-10-01 실측: 대회 527종목 18회 호출, 527종목 모두 응답).
//  - 시장 구분 코드는 J(KRX 정규장). 응답 output[]의 inter_shrn_iscd(종목코드)로 입력과 짝짓는다(응답 순서에 의존하지 않음).
//  - 출력 열(형식 고정):
//      종목코드 text        입력(정규화) 코드
//      종목명 text          inter_kor_isnm
//      현재가 number        inter2_prpr (거래정지 종목은 직전 종가가 온다. 0·빈 값 → 빈칸)
//      전일대비 number      inter2_prdy_vrss (기준가 대비, 부호 포함)
//      등락률 number        prdy_ctrt ÷ 100 (소수, 기준가 대비)
//      시가·고가·저가 number inter2_oprc / inter2_hgpr / inter2_lwpr (거래가 없으면 0으로 와서 빈칸)
//      거래량 number        acml_vol (주, 거래 없으면 0)
//      거래대금억 number    acml_tr_pbmn(원) ÷ 1억
//      기준가 number        inter2_sdpr — 오늘 가격제한폭·전일대비의 기준이 되는 '수정 전일종가'(아래 실측)
//      전일종가 number      inter2_prdy_clpr — 실제(수정 전) 전일 종가
//      상한가·하한가 number inter2_mxpr / inter2_llam
//      상태 text            OK / 오류: <사유>
//      조회시각 datetime
//  - 기준가 실측(2026-10-01 장중): 대회 527종목 모두 inter2_sdpr = inter2_prdy_clpr(그날 권리락·병합 없음).
//    주식병합 재상장일인 007110 일신석재(액면 500→2,500원, 10/1 재상장)는 inter2_sdpr = 3,015, inter2_prdy_clpr = 603,
//    inter2_prpr = 2,825, inter2_prdy_vrss = −190(= 2,825 − 3,015), 상한가 3,915·하한가 2,115(3,015 기준 ±30% 호가 단위),
//    수정주가 일봉(FID_ORG_ADJ_PRC=0)의 9/30 종가 = 3,015, 원주가 일봉의 9/30 종가 = 603
//    → inter2_sdpr = 기준가 = 수정주가 기준 직전 종가, inter2_prdy_clpr = 수정 전 실제 직전 종가. 둘이 다르면 그날 권리·병합·분할 등으로
//    기준가가 조정된 것이다. 거래정지 종목(084680 이월드, 병합 정지 중)은 현재가 = 직전 종가, 시가·고가·저가 0, 거래량 0으로 온다.
//  - 실패: 묶음 호출이 실패하면 그 묶음 종목만 상태 "오류: 사유", 응답에 없거나 현재가가 없는 종목은 "오류: 시세 없음(…)".
//    오류를 올리지 않는다 — 전부 실패했을 때 직전 표를 유지할지는 호출한 쿼리가 판단한다. 묶음 응답은 버퍼로 강제 평가해
//    호출 오류가 그 묶음의 try 안에서 잡히게 한다(M 지연 평가).
//  - 사용 예) Q = Table.Buffer(fnMultiQuote(fnPageRows()[종목코드]))
(codes as list) as table =>
let
    Path = "/uapi/domestic-stock/v1/quotations/intstock-multprice",
    Now = DateTime.LocalNow(),
    Digits = {"0".."9"},
    CodeChars = {"0".."9", "A".."Z"},
    Txt = (v as any) as nullable text =>
        if v = null then null else let t = Text.Trim(try Text.From(v) otherwise "") in if t = "" then null else t,
    // 종목코드 정규화 — 6자리 [0-9A-Z]가 아니면 null
    Norm = (v as any) as nullable text =>
        let
            t0 = Txt(v),
            t1 = if t0 = null then null else Text.Upper(t0),
            t2 = if t1 <> null and Text.Length(t1) < 6 and Text.Select(t1, Digits) = t1 then Text.PadStart(t1, 6, "0") else t1
        in
            if t2 <> null and Text.Length(t2) = 6 and Text.Select(t2, CodeChars) = t2 then t2 else null,
    // {출력 코드, 유효 여부} — 무효 코드는 입력 글자 그대로 한 행
    Entries = List.Distinct(
        List.Select(List.Transform(codes, (v) => let n = Norm(v), t = Txt(v) in {if n = null then (if t = null then null else Text.Upper(t)) else n, n <> null}),
            each _{0} <> null),
        each _{0}),
    Valid = List.Transform(List.Select(Entries, each _{1}), each _{0}),
    Chunks = List.Split(Valid, 30),
    Fetch = (chunk as list) as list =>
        let
            names = List.Combine(List.Transform(List.Positions(chunk), (i) =>
                        {"FID_COND_MRKT_DIV_CODE_" & Text.From(i + 1), "FID_INPUT_ISCD_" & Text.From(i + 1)})),
            vals = List.Combine(List.Transform(chunk, (c) => {"J", c})),
            J = fnKisGet(Path, "FHKST11300006", Record.FromList(vals, names)),
            O = Record.FieldOrDefault(J, "output", {}),
            L = if O is list then O else if O is record then {O} else {}
        in
            List.Select(L, each _ is record),
    // 묶음마다 한 번만 호출되도록 목록을 버퍼로 고정하고, 응답 목록도 버퍼로 강제 평가(try가 호출 오류를 잡도록)
    Results = List.Buffer(List.Transform(Chunks, (c) => [codes = c, r = try List.Buffer(Fetch(c))])),
    Msg = (e as record) as text =>
        let m = try Text.Trim(Text.From(e[Message])) otherwise "" in if m = "" then "조회 실패" else Text.Start(m, 120),
    RespPairs = List.Combine(List.Transform(Results, (x) =>
        if x[r][HasError] then {}
        else List.Transform(x[r][Value], (o) => {Norm(Record.FieldOrDefault(o, "inter_shrn_iscd", null)), o}))),
    RespDistinct = List.Distinct(List.Select(RespPairs, each _{0} <> null), each _{0}),
    Resp = Record.FromList(List.Transform(RespDistinct, each _{1}), List.Transform(RespDistinct, each _{0})),
    ErrPairs = List.Combine(List.Transform(Results, (x) =>
        if x[r][HasError] then List.Transform(x[codes], (c) => {c, Msg(x[r][Error])}) else {})),
    Err = Record.FromList(List.Transform(ErrPairs, each _{1}), List.Transform(ErrPairs, each _{0})),
    N = (o as nullable record, k as text) as nullable number => if o = null then null else fnNum(Record.FieldOrDefault(o, k, null)),
    // 가격류는 0이 '값 없음'(거래정지 종목의 시가·고가·저가 등) → 빈칸
    Px = (v as nullable number) as nullable number => if v = null or v = 0 then null else v,
    Row = (e as list) as list =>
        let
            code = e{0},
            valid = e{1},
            err = if valid then Record.FieldOrDefault(Err, code, null) else null,
            o = if valid then Record.FieldOrDefault(Resp, code, null) else null,
            prpr = Px(N(o, "inter2_prpr")),
            ctrt = N(o, "prdy_ctrt"),
            amt = N(o, "acml_tr_pbmn"),
            st = if not valid then "오류: 종목코드 형식(6자리 영숫자)"
                 else if err <> null then "오류: " & err
                 else if o = null or prpr = null then "오류: 시세 없음(거래정지·상장폐지 확인)"
                 else "OK"
        in
            {code, if o = null then null else Txt(Record.FieldOrDefault(o, "inter_kor_isnm", null)),
             prpr, N(o, "inter2_prdy_vrss"), if ctrt = null then null else ctrt / 100,
             Px(N(o, "inter2_oprc")), Px(N(o, "inter2_hgpr")), Px(N(o, "inter2_lwpr")),
             N(o, "acml_vol"), if amt = null then null else amt / 100000000,
             Px(N(o, "inter2_sdpr")), Px(N(o, "inter2_prdy_clpr")), Px(N(o, "inter2_mxpr")), Px(N(o, "inter2_llam")),
             st, Now},
    Result = #table(
        type table [종목코드 = text, 종목명 = nullable text, 현재가 = nullable number, 전일대비 = nullable number,
                    등락률 = nullable number, 시가 = nullable number, 고가 = nullable number, 저가 = nullable number,
                    거래량 = nullable number, 거래대금억 = nullable number, 기준가 = nullable number, 전일종가 = nullable number,
                    상한가 = nullable number, 하한가 = nullable number, 상태 = text, 조회시각 = datetime],
        List.Transform(Entries, Row))
in
    Result
