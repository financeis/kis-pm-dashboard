// T_IndexHist : KOSPI / KOSDAQ / KOSPI200 일별 지수 (벤치마크·시장 차트용)
//  - 조회기간: min(대회 시작일 - 45일, 오늘 - 가격이력일수×1.6일) ~ 오늘
let
    Main =
        let
            S = fnSettings(),
            Today = Date.From(DateTime.LocalNow()),
            Start = try Date.From(S[start_date]) otherwise Today,
            HistDays = try Number.From(S[hist_days]) otherwise 100,
            From = List.Min({Date.AddDays(Start, -45), Date.AddDays(Today, -Number.RoundUp(HistDays * 1.6))}),
            List0 = {{"KOSPI", "0001"}, {"KOSDAQ", "1001"}, {"KOSPI200", "2001"}},
            One = (name as text, code as text) as table =>
                let
                    T = fnIndexChart(code, From, Today),
                    WithStats = fnSeriesStats(Table.SelectColumns(T, {"일자", "시가", "고가", "저가", "종가", "거래량", "거래대금"})),
                    // 대회 시작 전일 종가 대비 누적수익률 (시작 전 구간은 null)
                    BaseClose = let prior = Table.SelectRows(WithStats, each [일자] < Start) in
                        if Table.RowCount(prior) = 0 then null else List.Last(prior[종가]),
                    AddCum = Table.AddColumn(WithStats, "대회누적", each
                        if BaseClose = null or [일자] < Start then null else [종가] / BaseClose - 1, type number),
                    AddName = Table.AddColumn(AddCum, "지수", each name, type text)
                in
                    AddName,
            All = Table.Combine(List.Transform(List0, each One(_{0}, _{1}))),
            WithPb = Table.TransformColumns(All, {{"거래대금", each if _ = null then null else _ / 100, type number}}),
            Renamed = Table.RenameColumns(WithPb, {{"거래대금", "거래대금억"}}),
            Final = Table.SelectColumns(Renamed, {"지수", "일자", "시가", "고가", "저가", "종가", "등락률", "대회누적",
                "MA5", "MA20", "MA60", "이격도20", "RSI14", "변동성20", "거래대금억"}),
            Keyed = Table.AddColumn(Final, "키", each [지수] & "|" & Date.ToText([일자], [Format = "yyyy-MM-dd", Culture = "en-US"]), type text)
        in
            Keyed,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblIndexHist", {"지수", "일자", "시가", "고가", "저가", "종가", "등락률", "대회누적", "MA5", "MA20", "MA60", "이격도20", "RSI14", "변동성20", "거래대금억", "키"}, Tried[Error][Message]) else Tried[Value]
in
    Final
