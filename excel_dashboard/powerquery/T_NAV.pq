// T_NAV : 일별 순자산(NAV) 재구성 — 매매일지 + 일별 종가(수정주가) + 벤치마크 지수
//  - NAV(t) = 초기자금 + Σ현금흐름(≤t) + Σ 보유수량(t) × 종가(t)   (외부 입출금 없음 가정)
//  - 첫 행 '기준' = 대회 시작 직전 거래일 (NAV = 초기자금, 누적 0%)
//  - 벤치마크: [설정] benchmark = KOSPI / KOSDAQ / 혼합(KOSPI 비중 bm_kospi_weight, 일별 리밸런싱)
let
    Main =
        let
            S = fnSettings(),
            C0 = try Number.From(S[init_capital]) otherwise 100000000,
            StartSet = try Date.From(S[start_date]) otherwise null,
            Today = Date.From(DateTime.LocalNow()),
            BmName = try Text.Upper(Text.Trim(Text.From(S[benchmark]))) otherwise "KOSPI",
            WK = try Number.From(S[bm_kospi_weight]) otherwise 0.8,
            Tr = fnTrades(),
            TrRecs = List.Buffer(Table.ToRecords(Tr)),
            FirstTrade = if List.Count(TrRecs) = 0 then null else List.Min(Tr[일자]),
            PeriodStart = List.Min(List.RemoveNulls({StartSet, FirstTrade, Today})),
            Lookback = Date.AddDays(PeriodStart, -20),
            Kospi = fnIndexChart("0001", Lookback, Today),
            Kosdaq = fnIndexChart("1001", Lookback, Today),
            AllDays = List.Sort(List.Distinct(Kospi[일자])),
            Prior = List.Select(AllDays, each _ < PeriodStart),
            BaseDay = if List.Count(Prior) = 0 then Date.AddDays(PeriodStart, -1) else List.Last(Prior),
            Days = List.Select(AllDays, each _ >= PeriodStart),
            Grid = List.Buffer({BaseDay} & Days),
            N = List.Count(Grid),
            GridT = Table.FromColumns({Grid}, type table [일자 = date]),
            Align = (P as table) as list =>
                let
                    J = Table.NestedJoin(GridT, {"일자"}, Table.SelectColumns(P, {"일자", "종가"}), {"일자"}, "p", JoinKind.LeftOuter),
                    E = Table.Sort(Table.ExpandTableColumn(J, "p", {"종가"}), {{"일자", Order.Ascending}}),
                    F = Table.FillDown(E, {"종가"})
                in
                    List.Buffer(F[종가]),
            KC = Align(Kospi),
            QC = Align(Kosdaq),
            Codes = List.Buffer(List.Distinct(List.Transform(TrRecs, each [종목코드]))),
            PriceFor = (code as text) as list =>
                let P = try Table.Buffer(fnDailyChart(code, Lookback, Today)) otherwise #table(type table [일자 = date, 종가 = number], {})
                in Align(P),
            Closes = Record.FromList(List.Transform(Codes, PriceFor), Codes),
            Signed = (t as record) as number => if t[구분] = "매수" then t[수량] else -t[수량],
            QtyAt = (code as text, d as date) as number =>
                List.Sum({0} & List.Transform(List.Select(TrRecs, each [종목코드] = code and [일자] <= d), Signed)),
            LastTradePx = (code as text, d as date) as nullable number =>
                let L = List.Select(TrRecs, each [종목코드] = code and [일자] <= d) in if List.Count(L) = 0 then null else List.Last(L)[단가],
            CashAt = (d as date) as number => C0 + List.Sum({0} & List.Transform(List.Select(TrRecs, each [일자] <= d), each [현금흐름])),
            Base = List.Buffer(List.Transform(List.Positions(Grid), (i) =>
                let
                    d = Grid{i},
                    prev = if i = 0 then #date(1900, 1, 1) else Grid{i - 1},
                    pos = List.Transform(Codes, (c) =>
                        let q = QtyAt(c, d), px = Record.Field(Closes, c){i}, p2 = if px = null then LastTradePx(c, d) else px
                        in [q = q, v = if q = 0 or p2 = null then 0 else q * p2]),
                    mv = List.Sum({0} & List.Transform(pos, each [v])),
                    cash = CashAt(d),
                    traded = List.Sum({0} & List.Transform(List.Select(TrRecs, each [일자] > prev and [일자] <= d), each [금액]))
                in
                    [d = d, nav = cash + mv, cash = cash, mv = mv, traded = traded, n = List.Count(List.Select(pos, each [q] > 0))])),
            NAV = List.Buffer(List.Transform(Base, each [nav])),
            Ret = List.Buffer(List.Transform({0 .. N - 1}, each if _ = 0 or NAV{_ - 1} = 0 then null else NAV{_} / NAV{_ - 1} - 1)),
            KR = List.Buffer(List.Transform({0 .. N - 1}, each if _ = 0 or KC{_ - 1} = null or KC{_} = null then null else KC{_} / KC{_ - 1} - 1)),
            QR = List.Buffer(List.Transform({0 .. N - 1}, each if _ = 0 or QC{_ - 1} = null or QC{_} = null then null else QC{_} / QC{_ - 1} - 1)),
            BR = List.Buffer(List.Transform({0 .. N - 1}, each
                if _ = 0 then null
                else if BmName = "KOSDAQ" then QR{_}
                else if BmName = "혼합" or BmName = "MIX" then (if KR{_} = null or QR{_} = null then null else WK * KR{_} + (1 - WK) * QR{_})
                else KR{_})),
            BmCum = List.Buffer(List.Accumulate(List.Skip(BR, 1), {0}, (acc, r) => acc & {(1 + List.Last(acc)) * (1 + (if r = null then 0 else r)) - 1})),
            Peak = List.Buffer(List.Accumulate(NAV, {}, (acc, v) => acc & {if List.Count(acc) = 0 then v else List.Max({List.Last(acc), v})})),
            Window = (i as number, n as number) as list => List.RemoveNulls(List.Range(Ret, List.Max({1, i - n + 1}), List.Min({n, i}))),
            Vol20 = (i as number) as nullable number => let w = Window(i, 20) in if List.Count(w) < 5 then null else List.StandardDeviation(w) * Number.Sqrt(252),
            CumSharpe = (i as number) as nullable number =>
                let w = List.RemoveNulls(List.Range(Ret, 1, i)) in
                    if List.Count(w) < 5 then null
                    else let sd = List.StandardDeviation(w) in if sd = 0 then null else List.Average(w) / sd * Number.Sqrt(252),
            Rows = List.Transform({0 .. N - 1}, (i) =>
                let b = Base{i} in [
                    구분 = if i = 0 then "기준" else "거래일",
                    경과일 = i,
                    일자 = b[d],
                    순자산 = b[nav],
                    현금 = b[cash],
                    주식평가 = b[mv],
                    주식비중 = if b[nav] = 0 then null else b[mv] / b[nav],
                    일간손익 = if i = 0 then null else b[nav] - NAV{i - 1},
                    일간수익률 = Ret{i},
                    누적수익률 = b[nav] / C0 - 1,
                    KOSPI = KC{i}, KOSPI일간 = KR{i}, KOSPI누적 = if KC{0} = null or KC{i} = null then null else KC{i} / KC{0} - 1,
                    KOSDAQ = QC{i}, KOSDAQ일간 = QR{i}, KOSDAQ누적 = if QC{0} = null or QC{i} = null then null else QC{i} / QC{0} - 1,
                    BM일간 = BR{i},
                    BM누적 = BmCum{i},
                    초과일간 = if Ret{i} = null or BR{i} = null then null else Ret{i} - BR{i},
                    누적초과 = (b[nav] / C0 - 1) - BmCum{i},
                    고점 = Peak{i},
                    낙폭 = if Peak{i} = 0 then null else b[nav] / Peak{i} - 1,
                    변동성20 = if i = 0 then null else Vol20(i),
                    누적샤프 = if i = 0 then null else CumSharpe(i),
                    매매금액 = b[traded],
                    보유종목수 = b[n]]),
            T = Table.FromRecords(Rows),
            Typed = Table.TransformColumnTypes(T, {{"구분", type text}, {"경과일", Int64.Type}, {"일자", type date}, {"순자산", type number},
                {"현금", type number}, {"주식평가", type number}, {"주식비중", type number}, {"일간손익", type number}, {"일간수익률", type number},
                {"누적수익률", type number}, {"KOSPI", type number}, {"KOSPI일간", type number}, {"KOSPI누적", type number},
                {"KOSDAQ", type number}, {"KOSDAQ일간", type number}, {"KOSDAQ누적", type number}, {"BM일간", type number},
                {"BM누적", type number}, {"초과일간", type number}, {"누적초과", type number}, {"고점", type number}, {"낙폭", type number},
                {"변동성20", type number}, {"누적샤프", type number}, {"매매금액", type number}, {"보유종목수", Int64.Type}})
        in
            Typed,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblNAV", {"구분", "경과일", "일자", "순자산", "현금", "주식평가", "주식비중", "일간손익", "일간수익률", "누적수익률", "KOSPI", "KOSPI일간", "KOSPI누적", "KOSDAQ", "KOSDAQ일간", "KOSDAQ누적", "BM일간", "BM누적", "초과일간", "누적초과", "고점", "낙폭", "변동성20", "누적샤프", "매매금액", "보유종목수"}, Tried[Error][Message]) else Tried[Value]
in
    Final
