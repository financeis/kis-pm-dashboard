// T_Risk : 현재 비중 기준 사전(ex-ante) 위험 분석 — 최근 N영업일(기본 60) 일간수익률 사용
//  - 포트폴리오 가상수익률 r_p(t) = Σ w_i × r_i(t)  (현재 비중을 과거에 적용, 현금 수익률 0)
//  - 위험기여(CTR_i) = w_i × Cov(r_i, r_p) / σ_p  → Σ CTR_i = σ_p,  위험기여비중 = CTR_i / σ_p
//  - 베타 = Cov(r_i, r_KOSPI) / Var(r_KOSPI),  VaR95(1일) = 1.645 × σ_p(일) × 순자산
let
    Main =
        let
            S = fnSettings(),
            C0 = try Number.From(S[init_capital]) otherwise 100000000,
            RiskDays = try Number.From(S[risk_days]) otherwise 60,
            Today = Date.From(DateTime.LocalNow()),
            From = Date.AddDays(Today, -Number.RoundUp(RiskDays * 1.6 + 10)),
            Tr = fnTrades(),
            Cash = C0 + List.Sum({0} & Tr[현금흐름]),
            P = fnPositions(fnLedger(Tr)),
            H = Table.SelectRows(P, each [보유수량] > 0),
            U = fnUniverse(),
            Mkt = fnIndexChart("0001", From, Today),
            Dates = List.Buffer(List.LastN(List.Sort(Mkt[일자]), RiskDays + 1)),
            DT = Table.FromColumns({Dates}, type table [일자 = date]),
            Align = (Px as table) as list =>
                let
                    J = Table.NestedJoin(DT, {"일자"}, Table.SelectColumns(Px, {"일자", "종가"}), {"일자"}, "p", JoinKind.LeftOuter),
                    E = Table.Sort(Table.ExpandTableColumn(J, "p", {"종가"}), {{"일자", Order.Ascending}}),
                    F = Table.FillDown(E, {"종가"})
                in
                    List.Buffer(F[종가]),
            ToRet = (c as list) as list => List.Buffer(List.Transform({1 .. List.Count(c) - 1}, each
                if c{_ - 1} = null or c{_} = null or c{_ - 1} = 0 then 0 else c{_} / c{_ - 1} - 1)),
            MC = Align(Mkt),
            RM = ToRet(MC),
            Mean = (x as list) as number => List.Average(x),
            Cov = (x as list, y as list) as nullable number =>
                let n = List.Count(x), mx = Mean(x), my = Mean(y) in
                    if n < 3 then null else List.Sum(List.Transform({0 .. n - 1}, each (x{_} - mx) * (y{_} - my))) / (n - 1),
            Hold = List.Buffer(List.Transform(Table.ToRecords(H), (h) =>
                let
                    px = try Table.Buffer(fnDailyChart(h[종목코드], From, Today)) otherwise #table(type table [일자 = date, 종가 = number], {}),
                    c = Align(px),
                    last = List.Last(List.RemoveNulls(c), h[평균단가]),
                    first = List.First(List.RemoveNulls(c), null)
                in
                    h & [closes = c, ret = ToRet(c), 종가 = last, 평가금액 = h[보유수량] * last,
                         수익률N = if first = null or first = 0 then null else last / first - 1])),
            MV = List.Sum({0} & List.Transform(Hold, each [평가금액])),
            NAV = Cash + MV,
            W = List.Transform(Hold, each if NAV = 0 then 0 else [평가금액] / NAV),
            T = List.Count(RM),
            RP = List.Buffer(List.Transform({0 .. T - 1}, (t) => List.Sum({0} & List.Transform(List.Positions(Hold), (i) => W{i} * Hold{i}[ret]{t})))),
            VarP = Cov(RP, RP),
            SdP = if VarP = null or VarP <= 0 then null else Number.Sqrt(VarP),
            VarM = Cov(RM, RM),
            SortedRP = List.Buffer(List.Sort(RP)),
            HistVaR = if T < 20 then null else -SortedRP{Number.RoundDown(0.05 * (T - 1))} * NAV,
            Rows = List.Transform(List.Positions(Hold), (i) =>
                let
                    h = Hold{i}, r = h[ret],
                    vi = Cov(r, r),
                    sdi = if vi = null or vi <= 0 then null else Number.Sqrt(vi),
                    cip = Cov(r, RP),
                    cim = Cov(r, RM),
                    ctr = if SdP = null or cip = null then null else W{i} * cip / SdP
                in
                    [구분 = "종목", 종목코드 = h[종목코드], 비중 = W{i}, 평가금액 = h[평가금액],
                     변동성 = if sdi = null then null else sdi * Number.Sqrt(252),
                     베타 = if VarM = null or VarM = 0 or cim = null then null else cim / VarM,
                     상관계수 = if sdi = null or SdP = null or cip = null then null else cip / (sdi * SdP),
                     한계위험 = if SdP = null or cip = null then null else cip / SdP * Number.Sqrt(252),
                     위험기여 = if ctr = null then null else ctr * Number.Sqrt(252),
                     위험기여비중 = if ctr = null or SdP = null then null else ctr / SdP,
                     수익률N = h[수익률N], 관측일수 = T]),
            BetaP = List.Sum({0} & List.Transform(Rows, each if [베타] = null then 0 else [비중] * [베타])),
            CashRow = [구분 = "현금", 종목코드 = "CASH", 비중 = if NAV = 0 then null else Cash / NAV, 평가금액 = Cash, 변동성 = 0, 베타 = 0,
                       상관계수 = null, 한계위험 = 0, 위험기여 = 0, 위험기여비중 = 0, 수익률N = 0, 관측일수 = T],
            PortRow = [구분 = "포트폴리오", 종목코드 = "PORT", 비중 = 1, 평가금액 = NAV,
                       변동성 = if SdP = null then null else SdP * Number.Sqrt(252), 베타 = BetaP, 상관계수 = null,
                       한계위험 = null, 위험기여 = if SdP = null then null else SdP * Number.Sqrt(252), 위험기여비중 = if SdP = null then null else 1,
                       수익률N = null, 관측일수 = T,
                       VaR95 = if SdP = null then null else 1.645 * SdP * NAV, 과거VaR95 = HistVaR,
                       시장변동성 = if VarM = null then null else Number.Sqrt(VarM) * Number.Sqrt(252)],
            All = Rows & {CashRow, PortRow},
            Cols = {"구분", "종목코드", "비중", "평가금액", "변동성", "베타", "상관계수", "한계위험", "위험기여", "위험기여비중", "수익률N",
                    "관측일수", "VaR95", "과거VaR95", "시장변동성"},
            Tbl = Table.FromRecords(All, Cols, MissingField.UseNull),
            J = Table.NestedJoin(Tbl, {"종목코드"}, U, {"종목코드"}, "u", JoinKind.LeftOuter),
            E = Table.ExpandTableColumn(J, "u", {"종목명", "섹터"}),
            Named = Table.AddColumn(E, "이름", each if [구분] = "현금" then "현금" else if [구분] = "포트폴리오" then "포트폴리오 합계"
                        else if [종목명] = null then [종목코드] else [종목명], type text),
            Final = Table.SelectColumns(Named, {"구분", "종목코드", "이름", "섹터"} & List.Skip(Cols, 2)),
            Typed = Table.TransformColumnTypes(Final, {{"구분", type text}, {"종목코드", type text}, {"이름", type text}, {"섹터", type text},
                {"비중", type number}, {"평가금액", type number}, {"변동성", type number}, {"베타", type number}, {"상관계수", type number},
                {"한계위험", type number}, {"위험기여", type number}, {"위험기여비중", type number}, {"수익률N", type number},
                {"관측일수", Int64.Type}, {"VaR95", type number}, {"과거VaR95", type number}, {"시장변동성", type number}}),
            Sorted = Table.Sort(Typed, {{each if [구분] = "종목" then 0 else if [구분] = "현금" then 1 else 2, Order.Ascending}, {"위험기여비중", Order.Descending}})
        in
            Sorted,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblRisk", {"구분", "종목코드", "이름", "섹터", "비중", "평가금액", "변동성", "베타", "상관계수", "한계위험", "위험기여", "위험기여비중", "수익률N", "관측일수", "VaR95", "과거VaR95", "시장변동성"}, Tried[Error][Message]) else Tried[Value]
in
    Final
