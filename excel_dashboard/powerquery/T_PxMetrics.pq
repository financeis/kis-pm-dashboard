// T_PxMetrics : 페이지 행별 가격 지표 — 가격 저장소로 계산, KIS 호출 없음 (2026-10-01) → tblPxMetrics
//  - 행 = fnPageRows() 순서(대회 종목 → 유니버스 밖). 실행 모드와 무관하게 늘 있는 표로 계산한다(build에서도 계산, 호출 없음).
//    단 버튼 실행(quick·full) 중 tblRunCtl started 이후에 이미 계산한 표(자기 표 조회시각 ≥ started)면 자기 표를 그대로 돌려준다
//    — 새로 고침 뒤 Excel이 백그라운드에서 쿼리를 다시 평가할 때 같은 계산을 다음 단계와 겹쳐 반복하지 않게.
//  - 원천: tblPxStore(없거나 비면 tblPxStoreSeed), 세션 달력 = fnPxCalendar(tblSessions → tblSessionsSeed). 달력이 비면(재생성 직후 등)
//    저장소에 있는 날짜들을 달력으로 쓴다. 달력 밖 날짜의 봉은 계산에 쓰지 않는다.
//    '지금'은 fnPxRunCtl(시험용 now_override 반영) — 완료 세션 판정에만 쓴다. 조회시각은 실제 계산 시각.
//  - 열·정의 (금액 억원, 비율 소수, 없으면 빈칸):
//      가격기준일  이 종목의 가장 최근 봉 날짜(보통 마지막 세션 L — [시세] 시점 봉)
//      현재가      그 봉의 종가(장중이면 마지막 [시세] 시점 가격, 장 마감 후면 종가)
//      기준가      가격기준일 직전 세션의 종가(1D 기준)
//      1D·1W·1M·3M·6M·1Y  현재가 ÷ (가격기준일에서 1·5·21·63·126·250세션 전 종가) − 1. 그 세션 봉이 없으면 공란
//      YTD         현재가 ÷ (가격기준일 연도의 직전 연도 마지막 세션 종가) − 1
//      D-19 … D0   최근 20세션(D0 = 달력의 마지막 세션, 장중이면 오늘) 각각의 일간 등락률 = 그날 종가 ÷ 직전 세션 종가 − 1
//      거래대금60일억·거래량60일천주  최근 60 '완료' 세션(마지막 세션이 오늘이고 15:30 전이면 오늘은 진행 중이라 빼고 그 전 60세션)의
//                  평균(거래량 ÷ 1,000). 그 구간에 있는 봉만 평균(상장 60세션 미만 종목은 있는 날만; 거래정지일은 0으로 포함)
//      고52주대비  현재가 ÷ (가격기준일까지 최근 250세션의 최고가 — 고가·종가 중 큰 값) − 1
//      상장주식수  tblUniverse[상장주식수](종목 마스터, 천 주 단위 반올림 값)에 그 종목이 있으면 그 값, 없으면 tblContest[상장주식수](9/30 기준)
//      시가총액억  현재가 × 상장주식수 ÷ 1억 (어느 쪽이든 없으면 공란)
//      상태        OK / 데이터 없음(저장소에 봉이 없음 — 첫 [전체] 전) / 저장소의 그 종목 상태가 "오류:"이면 그 문구(가격이 최신이 아닐 수 있음).
//                  저장소 표 전체가 '이전 데이터(갱신 실패)'인 경우는 T_PxStore 단계가 실패로 보고하므로 여기로 옮기지 않는다.
//  - 기간 세션 수는 아래 Back·StripN·AvgN·HighN 상수로 조정한다.
//  - 성능(527종목 × 260세션 저장소): 종목별로 정렬된 목록을 한 번 만들고 달력 위치로 바로 꺼낸다(종목별 레코드 조회 대신).
//  - 실패: 직전 표 유지(fnFallback + 열 형식 다시 지정), 직전 표가 없으면 종목코드가 빈 1행에 "오류: 사유".
let
    Periods = {"1D", "1W", "1M", "3M", "6M", "YTD", "1Y"},
    Back = [#"1D" = 1, #"1W" = 5, #"1M" = 21, #"3M" = 63, #"6M" = 126, #"1Y" = 250],
    StripN = 20,
    AvgN = 60,
    HighN = 250,
    StripCols = List.Transform(List.Reverse(List.Numbers(0, StripN)), each if _ = 0 then "D0" else "D-" & Text.From(_)),
    Cols = {"종목코드", "가격기준일", "현재가", "기준가"} & Periods & StripCols
           & {"거래대금60일억", "거래량60일천주", "고52주대비", "상장주식수", "시가총액억", "상태", "조회시각"},
    Types = List.Transform(Cols, each {_,
        if _ = "종목코드" or _ = "상태" then type text
        else if _ = "가격기준일" then type date
        else if _ = "조회시각" then type datetime
        else type number}),
    Ctl = fnPxRunCtl(),
    // 버튼 실행(quick·full) 중 이미 이번 실행(started 이후)에 계산한 표면 자기 표를 그대로 — 새로 고침 뒤 Excel의 백그라운드
    // 재평가가 13만 행 계산(약 5초)을 다음 단계와 겹쳐 반복하지 않게. build(빌드·수동 새로 고침)에서는 늘 다시 계산한다.
    OwnM = try Excel.CurrentWorkbook(){[Name = "tblPxMetrics"]}[Content] otherwise null,
    OwnOk = OwnM is table and List.ContainsAll(Table.ColumnNames(OwnM), Cols),
    LastStamp = if not OwnOk then null
                else List.Max(List.RemoveNulls(List.Transform(Table.Column(OwnM, "조회시각"), each try DateTime.From(_) otherwise null)), null),
    Already = (Ctl[Mode] = "quick" or Ctl[Mode] = "full") and Ctl[Started] <> null and LastStamp <> null and LastStamp >= Ctl[Started],
    KeptOwn = if OwnOk then Table.TransformColumnTypes(Table.SelectColumns(OwnM, Cols), Types) else null,
    // 계산은 함수 단계 안에서만(맨 위 단계를 따로 평가해도 계산하지 않음)
    Compute = () as table =>
        let
            T = Ctl[Today],
            Tm = Ctl[Time],
            Now = DateTime.LocalNow(),
            Digits = {"0".."9"},
            CodeChars = {"0".."9", "A".."Z"},
            Txt = (v as any) as nullable text =>
                if v = null then null else let t = Text.Trim(try Text.From(v) otherwise "") in if t = "" then null else t,
            Code = (v as any) as nullable text =>
                let
                    t0 = Txt(v),
                    t1 = if t0 = null then null else Text.Upper(t0),
                    t2 = if t1 <> null and Text.Length(t1) < 6 and Text.Select(t1, Digits) = t1 then Text.PadStart(t1, 6, "0") else t1
                in
                    if t2 <> null and Text.Length(t2) = 6 and Text.Select(t2, CodeChars) = t2 then t2 else null,
            Read = (name as text) as any => try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null,
            StoreCols = {"종목코드", "일자", "고가", "종가", "거래량", "거래대금억", "상태"},
            StoreTypes = {{"종목코드", type text}, {"일자", type date}, {"고가", type number}, {"종가", type number},
                          {"거래량", type number}, {"거래대금억", type number}, {"상태", type text}},
            // 자기 저장소 표는 T_PxStore가 형식을 붙여 쓴 값이라 형식만 붙인다(13만 행 성능). 시드는 코드 정규화·오류 값 정리·중복 제거
            FromSheet = (t as any, seed as logical) as table =>
                if not (t is table) then #table(StoreCols, {})
                else
                    let
                        s = Table.SelectColumns(t, StoreCols, MissingField.UseNull),
                        n = if seed then Table.TransformColumns(s, {{"종목코드", Code}}) else s,
                        typed = Table.TransformColumnTypes(n, StoreTypes),
                        fixed = if seed then Table.ReplaceErrorValues(typed, List.Transform(StoreCols, each {_, null})) else typed,
                        ok = Table.SelectRows(fixed, each [종목코드] <> null)
                    in
                        if seed then Table.Distinct(ok, {"종목코드", "일자"}) else ok,
            Own = Table.Buffer(FromSheet(Read("tblPxStore"), false)),
            Store = if Table.RowCount(Own) > 0 then Own else Table.Buffer(FromSheet(Read("tblPxStoreSeed"), true)),
            // 저장소의 종목별 오류 상태(첫 행 문구) — 지표 행 상태로 보여 줌
            ErrRows = Table.Buffer(Table.Distinct(Table.SelectRows(Store, each [상태] <> null and [상태] <> "OK" and Text.StartsWith([상태], "오류")), {"종목코드"})),
            ErrOf = Record.FromList(ErrRows[상태], ErrRows[종목코드]),
            Bars = Table.SelectRows(Store, each [일자] <> null and [종가] <> null),

            // 세션 달력 (비면 저장소 날짜)
            CalSes = fnPxCalendar(Ctl[Now])[Dates],
            Cal = List.Buffer(if not List.IsEmpty(CalSes) then CalSes
                  else List.Sort(List.Select(List.Distinct(Bars[일자]), each _ < T or (_ = T and Tm >= #time(9, 0, 0))))),
            N = List.Count(Cal),
            L = if N = 0 then null else Cal{N - 1},
            InProgress = L = T and Tm < #time(15, 30, 0),
            LastDone = if InProgress then N - 2 else N - 1,
            // 달력 위치(__p)를 붙이고(달력 밖 날짜는 빠짐) 종목·위치 순으로 정렬 → 종목별 목록(오름차순)
            CalT = #table(type table [__d = date, __p = Int64.Type], List.Transform(List.Positions(Cal), each {Cal{_}, _})),
            Joined = Table.Join(Table.RemoveColumns(Bars, {"상태"}), "일자", CalT, "__d", JoinKind.Inner),
            Sorted = Table.Sort(Joined, {{"종목코드", Order.Ascending}, {"__p", Order.Ascending}}),
            // 표 값은 참조할 때마다 다시 계산되므로(M) 두 번 쓰는 표는 버퍼로 고정
            Grouped = Table.Buffer(Table.Group(Sorted, {"종목코드"}, {
                    {"p", each List.Buffer([__p]), type list}, {"cl", each List.Buffer([종가]), type list},
                    {"hi", each List.Buffer([고가]), type list}, {"vo", each List.Buffer([거래량]), type list},
                    {"am", each List.Buffer([거래대금억]), type list}}, GroupKind.Local)),
            GroupRec = Record.FromList(Table.ToRecords(Grouped), Grouped[종목코드]),
            // 연도별 마지막 세션 위치 (YTD 기준)
            Years = List.Distinct(List.Transform(Cal, Date.Year)),
            YearLast = Record.FromList(List.Transform(Years, (y) => List.Last(List.Select(List.Positions(Cal), each Date.Year(Cal{_}) = y))),
                                       List.Transform(Years, Text.From)),

            // 상장주식수: 종목 마스터(tblUniverse) → 대회 명단(tblContest)
            SharesMap = (name as text) as record =>
                let
                    t = Read(name),
                    ok = t is table and List.ContainsAll(Table.ColumnNames(t), {"종목코드", "상장주식수"}),
                    pairs = if not ok then {}
                            else List.Transform(Table.ToRecords(Table.SelectColumns(t, {"종목코드", "상장주식수"})),
                                    (r) => {Code(r[종목코드]), try Number.From(r[상장주식수]) otherwise null}),
                    good = List.Distinct(List.Select(pairs, each _{0} <> null and _{1} <> null and _{1} > 0), each _{0})
                in
                    Record.FromList(List.Transform(good, each _{1}), List.Transform(good, each _{0})),
            UnivShares = SharesMap("tblUniverse"),
            ContestShares = SharesMap("tblContest"),
            SharesOf = (c as text) as nullable number =>
                let u = Record.FieldOrDefault(UnivShares, c, null) in if u <> null then u else Record.FieldOrDefault(ContestShares, c, null),

            Ratio = (a as nullable number, b as nullable number) as nullable number =>
                if a = null or b = null or b = 0 then null else a / b - 1,
            Avg = (xs as list) as nullable number => let v = List.RemoveNulls(xs) in if List.IsEmpty(v) then null else List.Average(v),
            Row = (c as text) as list =>
                let
                    g = Record.FieldOrDefault(GroupRec, c, null),
                    shares = SharesOf(c),
                    status = Record.FieldOrDefault(ErrOf, c, if g = null then "데이터 없음" else "OK")
                in
                    if g = null then
                        {c, null, null, null} & List.Repeat({null}, List.Count(Periods) + StripN) & {null, null, null, shares, null, status, Now}
                    else
                        let
                            p = g[p],
                            pmin = p{0},
                            p0 = List.Last(p),
                            // 빈 세션 없이 이어진 목록(보통)은 위치로 바로 꺼내고, 빈 세션이 있으면 위치 → 값 레코드로 꺼냄
                            dense = p0 - pmin + 1 = List.Count(p),
                            keys = if dense then {} else List.Transform(p, Text.From),
                            // (중복 날짜가 섞여도 오류 없이 처음 값만 — 저장소 불변식상 없어야 함)
                            Rec = (xs as list) as record =>
                                let pairs = List.Distinct(List.Zip({keys, xs}), each _{0})
                                in Record.FromList(List.Transform(pairs, each _{1}), List.Transform(pairs, each _{0})),
                            Cl = if dense then null else Rec(g[cl]),
                            Hi = if dense then null else Rec(g[hi]),
                            Vo = if dense then null else Rec(g[vo]),
                            Am = if dense then null else Rec(g[am]),
                            At = (xs as list, rec as nullable record, pos as number) as any =>
                                if pos < pmin or pos > p0 then null
                                else if dense then xs{pos - pmin}
                                else Record.FieldOrDefault(rec, Text.From(pos), null),
                            Range = (xs as list, rec as nullable record, a as number, b as number) as list =>
                                let lo = List.Max({a, pmin}), hi = List.Min({b, p0})
                                in if lo > hi then {}
                                   else if dense then List.Range(xs, lo - pmin, hi - lo + 1)
                                   else List.Transform(List.Numbers(lo, hi - lo + 1), (q) => Record.FieldOrDefault(rec, Text.From(q), null)),
                            Close = (pos as number) as nullable number => At(g[cl], Cl, pos),
                            d0 = Cal{p0},
                            px = Close(p0),
                            yPos = Record.FieldOrDefault(YearLast, Text.From(Date.Year(d0) - 1), null),
                            periodVals = List.Transform(Periods, (k) =>
                                            if k = "YTD" then (if yPos = null then null else Ratio(px, Close(yPos)))
                                            else Ratio(px, Close(p0 - Record.Field(Back, k)))),
                            strip = List.Transform(List.Reverse(List.Numbers(0, StripN)), (i) =>
                                        let q = (N - 1) - i in Ratio(Close(q), Close(q - 1))),
                            amt60 = Avg(Range(g[am], Am, LastDone - AvgN + 1, LastDone)),
                            vol60 = let v = Avg(Range(g[vo], Vo, LastDone - AvgN + 1, LastDone)) in if v = null then null else v / 1000,
                            hi52 = List.Max(List.RemoveNulls(Range(g[hi], Hi, p0 - HighN + 1, p0) & Range(g[cl], Cl, p0 - HighN + 1, p0)), null),
                            mcap = if px = null or shares = null then null else px * shares / 100000000
                        in
                            {c, d0, px, Close(p0 - 1)} & periodVals & strip & {amt60, vol60, Ratio(px, hi52), shares, mcap, status, Now},
            // 종목 행마다 모든 칸을 강제 평가해 계산 오류를 그 행의 상태로(목록·표 버퍼는 칸 안 오류를 올리지 않고 빈칸으로 적재하므로)
            SafeRow = (c as text) as list =>
                let
                    r = try (let x = Row(c), n = List.Count(List.Select(x, each _ = null or _ <> null)) in if n = List.Count(Cols) then x else error "열 수 불일치"),
                    why = if r[HasError] then (try Text.Start(Text.From(r[Error][Message]), 100) otherwise "사유 불명") else null
                in
                    if r[HasError] then {c} & List.Repeat({null}, List.Count(Cols) - 3) & {"오류: 지표 계산 실패(" & why & ")", Now}
                    else r[Value],
            Pages = List.Buffer(List.Distinct(List.RemoveNulls(List.Transform(fnPageRows()[종목코드], Code)))),
            Out = Table.TransformColumnTypes(#table(Cols, List.Transform(Pages, SafeRow)), Types)
        in
            Out,
    Tried = try Table.Buffer(if Already then KeptOwn else Compute()),
    Reason = if not Tried[HasError] then null
             else let m = Tried[Error][Message] in if m = null or m = "" then "사유 불명" else m,
    KeptOnFail = if Tried[HasError] then fnFallback("tblPxMetrics", Cols, Reason) else null,
    Warn = #table(Cols, {List.Repeat({null}, List.Count(Cols) - 2) & {"오류: " & Text.Start(if Reason = null then "" else Reason, 120), DateTime.LocalNow()}}),
    Final =
        if not Tried[HasError] then Tried[Value]
        else Table.TransformColumnTypes(if Table.RowCount(KeptOnFail) > 0 then KeptOnFail else Warn, Types)
in
    Final
