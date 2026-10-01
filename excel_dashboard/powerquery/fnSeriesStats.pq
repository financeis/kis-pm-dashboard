// fnSeriesStats : 일자 오름차순 가격표(일자, 종가, [거래량])에 기술지표 열을 추가
//  - 등락률(전일 대비), MA5/MA20/MA60, 이격도20(종가/MA20-1), RSI14(Wilder), 변동성20(연율화), 거래량20평균
(T as table) as table =>
let
    Sorted = Table.Sort(T, {{"일자", Order.Ascending}}),
    C = List.Buffer(Sorted[종가]),
    V = List.Buffer(if Table.HasColumns(Sorted, "거래량") then Sorted[거래량] else List.Repeat({null}, List.Count(C))),
    N = List.Count(C),
    Ret = List.Buffer(List.Transform({0 .. N - 1}, each
        if _ = 0 or C{_ - 1} = null or C{_ - 1} = 0 or C{_} = null then null else C{_} / C{_ - 1} - 1)),
    MA = (i as number, n as number) as nullable number =>
        if i < n - 1 then null else List.Average(List.Range(C, i - n + 1, n)),
    Vol20 = (i as number) as nullable number =>
        if i < 20 then null
        else let r = List.RemoveNulls(List.Range(Ret, i - 19, 20)) in
            if List.Count(r) < 15 then null else List.StandardDeviation(r) * Number.Sqrt(252),
    AvgVol20 = (i as number) as nullable number =>
        if i < 19 then null else List.Average(List.RemoveNulls(List.Range(V, i - 19, 20))),
    // RSI(14) — Wilder 평활
    Rn = 14,
    D = if N < 2 then {} else List.Transform({1 .. N - 1}, each C{_} - C{_ - 1}),
    G = List.Buffer(List.Transform(D, each if _ > 0 then _ else 0)),
    L = List.Buffer(List.Transform(D, each if _ < 0 then -_ else 0)),
    RsiTail =
        if N <= Rn then {}
        else List.Generate(
            () => [i = Rn, ag = List.Average(List.FirstN(G, Rn)), al = List.Average(List.FirstN(L, Rn))],
            each [i] < N,
            each [i = [i] + 1, ag = ([ag] * (Rn - 1) + G{[i]}) / Rn, al = ([al] * (Rn - 1) + L{[i]}) / Rn],
            each if [al] = 0 then 100 else 100 - 100 / (1 + [ag] / [al])),
    RSI = List.Buffer(List.Repeat({null}, List.Min({Rn, N})) & RsiTail),
    Idx = Table.AddIndexColumn(Sorted, "__i", 0, 1, Int64.Type),
    Out = Table.AddColumn(Idx, "__s", each
        let i = [__i], ma20 = MA(i, 20) in [
            등락률 = Ret{i},
            MA5 = MA(i, 5),
            MA20 = ma20,
            MA60 = MA(i, 60),
            이격도20 = if ma20 = null or ma20 = 0 then null else C{i} / ma20 - 1,
            RSI14 = RSI{i},
            변동성20 = Vol20(i),
            거래량20평균 = AvgVol20(i)]),
    Expanded = Table.ExpandRecordColumn(Out, "__s",
        {"등락률", "MA5", "MA20", "MA60", "이격도20", "RSI14", "변동성20", "거래량20평균"}),
    Typed = Table.TransformColumnTypes(Table.RemoveColumns(Expanded, {"__i"}), {
        {"등락률", type number}, {"MA5", type number}, {"MA20", type number}, {"MA60", type number},
        {"이격도20", type number}, {"RSI14", type number}, {"변동성20", type number}, {"거래량20평균", type number}})
in
    Typed
