// T_PriceHist : 보유·관심 종목 일봉 + 기술지표 (MA5/20/60, RSI14, 20일 변동성, 거래량 20일평균)
//  - 키 = 종목코드|역순 (역순 0 = 가장 최근 거래일)  → 시세판/포트폴리오에서 최신 지표를 XLOOKUP으로 조회
let
    Main =
        let
            S = fnSettings(),
            Today = Date.From(DateTime.LocalNow()),
            HistDays = try Number.From(S[hist_days]) otherwise 100,
            From = Date.AddDays(Today, -Number.RoundUp(HistDays * 1.5)),
            Codes = List.Buffer(fnCodes()[종목코드]),
            One = (code as text) as table =>
                let
                    r = try Table.Buffer(fnDailyChart(code, From, Today)),
                    T = if r[HasError] then null else r[Value],
                    Stats = if T = null or Table.RowCount(T) = 0 then null else fnSeriesStats(T),
                    N = if Stats = null then 0 else Table.RowCount(Stats),
                    WithCode = if Stats = null then null
                               else Table.AddIndexColumn(Table.AddColumn(Stats, "종목코드", each code, type text), "__i", 0, 1, Int64.Type),
                    WithRev = if WithCode = null then null
                              else Table.RemoveColumns(Table.AddColumn(WithCode, "역순", each N - 1 - [__i], Int64.Type), {"__i"})
                in
                    if WithRev = null then #table({"종목코드", "상태"}, {{code, if r[HasError] then "오류: " & r[Error][Message] else "데이터 없음"}})
                    else Table.AddColumn(WithRev, "상태", each "OK", type text),
            All0 = Table.Combine(List.Transform(Codes, One)),
            Errs = if Table.HasColumns(All0, "상태") then List.Select(All0[상태], each Text.StartsWith(Text.From(_), "오류")) else {},
            // 전부 실패하면 오류로 올려 직전 데이터를 유지(fnFallback)
            All = if Table.RowCount(All0) > 0 and List.Count(Errs) = Table.RowCount(All0) then error Text.Middle(Errs{0}, 4) else All0,
            Cols = {"종목코드", "일자", "역순", "시가", "고가", "저가", "종가", "거래량", "거래대금", "등락률", "MA5", "MA20", "MA60",
                    "이격도20", "RSI14", "변동성20", "거래량20평균", "상태"},
            Safe = if Table.RowCount(All) = 0 then #table(Cols, {}) else Table.SelectColumns(All, Cols, MissingField.UseNull),
            WithAmt = Table.TransformColumns(Safe, {{"거래대금", each if _ = null then null else _ / 100000000, type number}}),
            Renamed = Table.RenameColumns(WithAmt, {{"거래대금", "거래대금억"}}),
            Keyed = Table.AddColumn(Renamed, "키", each [종목코드] & "|" & (if [역순] = null then "x" else Text.From([역순])), type text),
            Typed = Table.TransformColumnTypes(Keyed, {{"종목코드", type text}, {"일자", type date}, {"역순", Int64.Type}, {"시가", type number},
                {"고가", type number}, {"저가", type number}, {"종가", type number}, {"거래량", type number}, {"등락률", type number},
                {"MA5", type number}, {"MA20", type number}, {"MA60", type number}, {"이격도20", type number}, {"RSI14", type number},
                {"변동성20", type number}, {"거래량20평균", type number}, {"상태", type text}})
        in
            Typed,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblPriceHist", {"종목코드", "일자", "역순", "시가", "고가", "저가", "종가", "거래량", "거래대금억", "등락률", "MA5", "MA20", "MA60", "이격도20", "RSI14", "변동성20", "거래량20평균", "상태", "키"}, Tried[Error][Message]) else Tried[Value]
in
    Final
