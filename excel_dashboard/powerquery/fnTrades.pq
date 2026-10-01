// fnTrades : [매매일지] tblTrades 입력값을 정리 (수수료·세금 기본값 적용, 현금흐름 계산)
//  - 수수료/세금 칸이 비어 있으면 [설정]의 수수료율·거래세율로 계산 (원 미만 절사)
//  - 현금흐름: 매수 = -(금액+수수료), 매도 = 금액-수수료-세금
() as table =>
let
    S = fnSettings(),
    FeeRate = try Number.From(S[fee_rate]) otherwise 0.00015,
    TaxRate = try Number.From(S[tax_rate]) otherwise 0.002,
    Raw = Excel.CurrentWorkbook(){[Name = "tblTrades"]}[Content],
    Col = (r as record, k as text) as any => Record.FieldOrDefault(r, k, null),
    NormCode = (v as any) as nullable text =>
        if v = null then null
        else let t = Text.Upper(Text.Trim(Text.From(v))) in
            if t = "" then null
            else if Text.Length(t) < 6 and (try Number.FromText(t) otherwise null) <> null then Text.PadStart(t, 6, "0")
            else t,
    NormSide = (v as any) as nullable text =>
        let t = if v = null then "" else Text.Upper(Text.Trim(Text.From(v))) in
            if List.Contains({"매수", "BUY", "B", "LONG"}, t) then "매수"
            else if List.Contains({"매도", "SELL", "S"}, t) then "매도" else null,
    ToNum = (v as any) as nullable number => if v = null or v = "" then null else try Number.From(v) otherwise null,
    RawRecs = List.Buffer(Table.ToRecords(Raw)),
    Recs = List.Transform(List.Positions(RawRecs), (i) =>
        let
            r = RawRecs{i},
            d = try Date.From(Col(r, "일자")) otherwise null,
            code = NormCode(Col(r, "종목코드")),
            side = NormSide(Col(r, "구분")),
            q = ToNum(Col(r, "수량")),
            p = ToNum(Col(r, "단가")),
            amt = if q = null or p = null then null else q * p,
            feeIn = ToNum(Col(r, "수수료")),
            taxIn = ToNum(Col(r, "세금")),
            fee = if feeIn <> null then feeIn else if amt = null then null else Number.RoundDown(amt * FeeRate),
            tax = if side <> "매도" then 0 else if taxIn <> null then taxIn else if amt = null then null else Number.RoundDown(amt * TaxRate),
            cf = if amt = null then null else if side = "매수" then -(amt + fee) else amt - fee - tax
        in
            [순번 = i + 1, 일자 = d, 종목코드 = code, 구분 = side, 수량 = q, 단가 = p, 금액 = amt, 수수료 = fee, 세금 = tax,
             현금흐름 = cf, 전략 = Col(r, "전략"), 매매근거 = Col(r, "매매근거"),
             목표가 = ToNum(Col(r, "목표가")), 손절가 = ToNum(Col(r, "손절가")), 메모 = Col(r, "메모")]),
    Valid = List.Select(Recs, each [일자] <> null and [종목코드] <> null and [구분] <> null and [수량] <> null and [수량] > 0 and [단가] <> null and [단가] > 0),
    T = Table.FromRecords(Valid, type table [순번 = Int64.Type, 일자 = date, 종목코드 = text, 구분 = text, 수량 = number, 단가 = number,
            금액 = number, 수수료 = number, 세금 = number, 현금흐름 = number, 전략 = any, 매매근거 = any, 목표가 = number, 손절가 = number, 메모 = any]),
    Sorted = Table.Sort(T, {{"일자", Order.Ascending}, {"순번", Order.Ascending}})
in
    Sorted
