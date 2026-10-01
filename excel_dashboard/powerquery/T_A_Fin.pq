// T_A_Fin : 종목분석 8구역 — 분기 실적 최근 8분기(매출·영업이익·순이익, YoY) (spec R17-8, R7 YoY 규칙) → tblA_Fin (2026-10-01 T15)
//  - 원천(호출 2회):
//      국내주식 손익계산서 FHKST66430200 /uapi/domestic-stock/v1/finance/income-statement (FID_DIV_CLS_CODE 1 = 분기)
//        → 최근 약 30분기, stac_yymm(결산년월) 내림차순. 값은 회계연도 누적(2026-10-01 실측: 005930 202603 133.9조, 202606 305.4조),
//          단위 억원(2023 연간 매출 2,589,355억 = KIS 추정 API·공시와 일치). 매출 sale_account, 영업이익 bsop_prti, 순이익 thtr_ntin(당기순이익).
//          빈 값 표시 ""·99.99(미제공 칸)·0.00(자료 없는 분기 — 모든 항목이 0.00)은 모두 빈칸(T_Fin과 같은 규칙 — 정확히 0억인 진짜 값도 빈칸).
//      주식현재가 시세 FHKST01010100 → stac_month(결산월; 없으면 12) — 회계연도의 첫 분기를 정하려고.
//  - 분기 단독값: 회계연도 첫 분기(결산월 + 3개월, 12월 결산이면 3월)는 누적 그대로, 그 밖은 누적 − 직전 분기(3개월 전) 누적.
//    직전 분기 행이 없으면 빈칸.
//  - YoY = 분기 단독값 ÷ 전년 동분기(12개월 전) 단독값 − 1 — 전년 값이 0 이하이거나 빈칸이면 빈칸(spec R7). 영업이익률 = 영업이익 ÷ 매출(매출 > 0).
//  - 열: 결산년월 "2026.06", 분기 "2026.Q2"(분기 말 달력 기준), 정렬 결산년월 오름차순(마지막 행 = 최근 분기).
//    행 = 값이 하나라도 있는 가장 최근 분기까지의 응답 분기 중 최근 8개(중간에 값이 전부 빈 분기도 행으로 남김 — T_Fin과 같음).
//    값이 있는 분기가 없으면(우선주·KIS 재무 자료 없는 종목·없는 코드) "데이터 없음" 1행.
//  - T_Fin(→ tblFin)은 결산월을 누적값이 줄어드는 곳으로 추정하지만, 여기서는 한 종목이라 현재가 시세의 결산월(stac_month)을 바로 쓴다.
let
    Types = {{"종목코드", type text}, {"결산년월", type text}, {"분기", type text}, {"매출억", type number}, {"영업이익억", type number},
             {"순이익억", type number}, {"영업이익률", type number}, {"매출YoY", type number}, {"영업이익YoY", type number},
             {"순이익YoY", type number}, {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Keep = 8,
            J = fnKisGet("/uapi/domestic-stock/v1/finance/income-statement", "FHKST66430200",
                    [FID_DIV_CLS_CODE = "1", fid_cond_mrkt_div_code = "J", fid_input_iscd = code]),
            O = Record.FieldOrDefault(J, "output", {}),
            L = List.Select(if O is list then O else if O is record then {O} else {}, each _ is record),
            P = let j = fnKisGet("/uapi/domestic-stock/v1/quotations/inquire-price", "FHKST01010100",
                        [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code]), o = Record.FieldOrDefault(j, "output", null)
                in if o is record then o else [],
            // 원값은 try 밖에서 읽음 — 시세 호출 실패가 '12월 결산' 가정으로 묻히지 않고 표 실패(직전 값 유지)로 가게
            FyeRaw = Record.FieldOrDefault(P, "stac_month", null),
            FyeIn = if FyeRaw = null then null else try Number.From(Text.Trim(Text.From(FyeRaw))) otherwise null,
            Fye = if FyeIn = null or FyeIn < 1 or FyeIn > 12 then 12 else FyeIn,
            FirstMonth = Number.Mod(Fye + 2, 12) + 1,
            Val = (r as record, k as text) as nullable number =>
                let v = fnNum(Record.FieldOrDefault(r, k, null)) in if v = null or v = 99.99 or v = 0 then null else v,
            // 누적 행: ym = 년*100+월
            Cum = List.Buffer(List.RemoveNulls(List.Transform(L, (r) =>
                let
                    t = Text.Trim(Text.From(Record.FieldOrDefault(r, "stac_yymm", ""))),
                    ym = if Text.Length(t) = 6 then try Number.From(t) otherwise null else null
                in
                    if ym = null then null
                    else [ym = ym, 매출 = Val(r, "sale_account"), 영업 = Val(r, "bsop_prti"), 순익 = Val(r, "thtr_ntin")]))),
            Keys = List.Transform(Cum, each Text.From([ym])),
            Distinct = List.Distinct(List.Zip({Keys, Cum}), each _{0}),
            ByYm = Record.FromList(List.Transform(Distinct, each _{1}), List.Transform(Distinct, each _{0})),
            Shift = (ym as number, months as number) as number =>
                let y = Number.IntegerDivide(ym, 100), m = Number.Mod(ym, 100), t = y * 12 + (m - 1) + months
                in Number.IntegerDivide(t, 12) * 100 + Number.Mod(t, 12) + 1,
            Sub = (a as nullable number, b as nullable number) as nullable number => if a = null or b = null then null else a - b,
            // 분기 단독값 레코드 (없으면 null)
            Discrete = (ym as number) as nullable record =>
                let
                    cur = Record.FieldOrDefault(ByYm, Text.From(ym), null),
                    prev = Record.FieldOrDefault(ByYm, Text.From(Shift(ym, -3)), null)
                in
                    if cur = null then null
                    else if Number.Mod(ym, 100) = FirstMonth then cur
                    else if prev = null then [ym = ym, 매출 = null, 영업 = null, 순익 = null]
                    else [ym = ym, 매출 = Sub(cur[매출], prev[매출]), 영업 = Sub(cur[영업], prev[영업]), 순익 = Sub(cur[순익], prev[순익])],
            YoY = (cur as nullable number, base as nullable number) as nullable number =>
                if cur = null or base = null or base <= 0 then null else cur / base - 1,
            // 값이 하나라도 있는 가장 최근 분기까지의 최근 8개 분기
            WithData = List.Transform(List.Select(Cum, each [매출] <> null or [영업] <> null or [순익] <> null), each [ym]),
            Anchor = if List.IsEmpty(WithData) then null else List.Max(WithData),
            Yms = if Anchor = null then {}
                  else List.LastN(List.Sort(List.Select(List.Distinct(List.Transform(Cum, each [ym])), each _ <= Anchor)), Keep),
            Rows = List.Transform(Yms, (ym) =>
                let
                    d = Discrete(ym),
                    b = Discrete(Shift(ym, -12)),
                    y = Number.IntegerDivide(ym, 100),
                    m = Number.Mod(ym, 100),
                    bv = (k as text) as nullable number => if b = null then null else Record.Field(b, k)
                in
                    {code, Text.From(y) & "." & Text.PadStart(Text.From(m), 2, "0"), Text.From(y) & ".Q" & Text.From(Number.RoundUp(m / 3)),
                     d[매출], d[영업], d[순익], if d[매출] = null or d[매출] <= 0 or d[영업] = null then null else d[영업] / d[매출],
                     YoY(d[매출], bv("매출")), YoY(d[영업], bv("영업")), YoY(d[순익], bv("순익")), "OK", now})
        in
            if List.IsEmpty(Yms) then fnAOneRow(Types, code, "데이터 없음", now)
            else #table(List.Transform(Types, each _{0}), Rows)
in
    fnARun("tblA_Fin", Types, Main)
