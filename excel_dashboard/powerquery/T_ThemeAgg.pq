// T_ThemeAgg : 테마 집계 — 대회 종목만, 대테마별 행과 그 아래 세부테마별 행 (spec R15, 2026-10-01 T16) → tblThemeAgg
//   [KIS 호출 없음 — 어느 모드에서나 tblCompany(대회종목 표)로 계산]
//  - 원천: tblCompany. [시세]·[전체]에서 T_Company 바로 다음, [업종]에서는 T_SectorKRX 다음에 새로 고쳐 늘 최신 대회종목 표를 반영한다
//    (같은 새로 고침 안에서 다른 쿼리의 적재 표는 직전 결과로 읽히므로, 빌드도 T_Company → T_ThemeAgg 순서로 새로 고쳐야 한다).
//    tblCompany가 없거나 비어 있으면(빌드 직후 등) 열만 있는 빈 표(오류 아님).
//  - 대상 행: tblCompany의 유니버스 = 대회이고 대테마가 있는 행(테마·소속은 fnClassify가 수정표까지 반영한 값). 종목코드가 빈
//    경고 행·유니버스 밖 행은 빠진다. 세부테마가 빈 종목은 대테마 행에만 들어간다(세부테마 행을 만들지 않음).
//  - 행 순서: 대테마를 시가총액 합 내림차순(빈칸은 뒤, 같으면 이름순)으로, 각 대테마 바로 아래 그 세부테마 행들을 같은 규칙으로.
//    세부테마 행은 (대테마, 세부테마) 쌍으로 묶는다(같은 세부테마 이름이 다른 대테마에 있어도 따로).
//  - 열 (비율은 소수 — 페이지가 % 서식, 계산할 종목이 없으면 빈칸이고 0을 쓰지 않음):
//      단계            대테마 / 세부테마 (대테마 행은 세부테마 칸이 빈칸)
//      종목수          그 묶음의 종목 수
//      시가총액억      시가총액억 합(값이 있는 종목만, 모두 비면 빈칸)
//      1D…1Y           현재 시가총액 가중 평균 등락률 = Σ(시가총액 × 등락률) ÷ Σ시가총액 — 그 기간 등락률과 시가총액이 둘 다 있는
//                      종목만(가중치는 그 종목들로 다시 나눔)
//      상승비율1D·1W   그 기간 등락률이 있는 종목 중 등락률 > 0인 비율
//      외국인·기관 1D·1W·1M %  Σ순매수 억원 ÷ Σ시가총액억 — 그 순매수 값과 시가총액이 둘 다 있는 종목만(테마 시가총액 대비)
//      상태            OK / 데이터 없음(묶음의 어느 종목에도 시가총액·등락률이 없음 — 첫 [전체] 전 등, 정상 값)
//      조회시각        실제 계산 시각
//  - 실행 가드: 버튼 실행(quick·full·sector) 중 tblRunCtl started 이후에 이미 계산한 표(자기 표 조회시각 ≥ started)면 자기 표를
//    그대로 돌려준다(새로 고침 뒤 Excel의 백그라운드 재평가에서 다시 계산하지 않게 — T_PxMetrics·T_Company와 같은 가드).
//  - 성능: 행마다 값 목록을 버퍼로 고정해 묶음마다 위치로 읽는다(목록 식은 읽을 때마다 다시 계산되므로). 실측(2026-10-01, 대회 527종목
//    합성 입력, quick + started·동기 새로 고침) 0.3~0.5초.
//  - 실패: 직전 표 유지(fnFallback — 상태 '이전 데이터(갱신 실패: 사유)', 열 형식 다시 지정), 직전 표가 없으면 단계가 빈 1행에 "오류: 사유".
let
    Periods = {"1D", "1W", "1M", "3M", "6M", "YTD", "1Y"},
    FlowKeys = {"외국인1D", "외국인1W", "외국인1M", "기관1D", "기관1W", "기관1M"},
    Cols = {"단계", "대테마", "세부테마", "종목수", "시가총액억"} & Periods & {"상승비율1D", "상승비율1W"}
           & List.Transform(FlowKeys, each _ & "%") & {"상태", "조회시각"},
    NCols = List.Count(Cols),
    Types = List.Transform(Cols, each {_,
        if List.Contains({"단계", "대테마", "세부테마", "상태"}, _) then type text
        else if _ = "조회시각" then type datetime
        else type number}),
    Ctl = fnPxRunCtl(),
    OwnT = try Excel.CurrentWorkbook(){[Name = "tblThemeAgg"]}[Content] otherwise null,
    OwnOk = OwnT is table and List.ContainsAll(Table.ColumnNames(OwnT), Cols),
    LastStamp = if not OwnOk then null
                else List.Max(List.RemoveNulls(List.Transform(Table.Column(OwnT, "조회시각"), each try DateTime.From(_) otherwise null)), null),
    Already = List.Contains({"quick", "full", "sector"}, Ctl[Mode]) and Ctl[Started] <> null and LastStamp <> null and LastStamp >= Ctl[Started],
    KeptOwn = if OwnOk then Table.TransformColumnTypes(Table.SelectColumns(OwnT, Cols), Types) else null,

    Compute = () as table =>
        let
            Now = DateTime.LocalNow(),
            Txt = (v as any) as nullable text =>
                if v = null then null else let t = Text.Trim(try Text.From(v) otherwise "") in if t = "" then null else t,
            Num = (v as any) as nullable number =>
                if v = null then null
                else if v is number then (if Number.IsNaN(v) then null else v)
                else if v is text then fnNum(v)
                else null,
            Src = try Excel.CurrentWorkbook(){[Name = "tblCompany"]}[Content] otherwise null,
            Need = {"유니버스", "대테마", "세부테마", "시가총액억"} & Periods & List.Transform(FlowKeys, each _ & "억"),
            Recs = if not (Src is table) then {} else Table.ToRecords(Table.SelectColumns(Src, Need, MissingField.UseNull)),
            // 행마다 값 목록 v = {시가총액억, 1D…1Y(7), 외국인·기관 1D·1W·1M 억원(6)}을 버퍼로 고정(위치로 여러 번 읽어도 다시 계산하지 않음)
            Rows = List.Buffer(List.Select(List.Transform(Recs, (r) =>
                       [big = Txt(r[대테마]), small = Txt(r[세부테마]), uni = Txt(r[유니버스]),
                        v = List.Buffer({Num(r[시가총액억])} & List.Transform(Periods, (k) => Num(Record.Field(r, k)))
                                        & List.Transform(FlowKeys, (k) => Num(Record.Field(r, k & "억"))))]),
                   each [uni] = "대회" and [big] <> null)),
            // 묶음 하나 → 표 1행(값 순서 = Cols)
            Agg = (level as text, big as text, small as nullable text, g as list) as list =>
                let
                    caps = List.RemoveNulls(List.Transform(g, each [v]{0})),
                    capSum = if List.IsEmpty(caps) then null else List.Sum(caps),
                    Weighted = (i as number) as nullable number =>
                        let
                            pairs = List.Select(g, each [v]{0} <> null and [v]{1 + i} <> null),
                            w = List.Sum(List.Transform(pairs, each [v]{0}))
                        in
                            if List.IsEmpty(pairs) or w = null or w <= 0 then null
                            else List.Sum(List.Transform(pairs, each [v]{0} * [v]{1 + i})) / w,
                    UpRatio = (i as number) as nullable number =>
                        let x = List.RemoveNulls(List.Transform(g, each [v]{1 + i}))
                        in if List.IsEmpty(x) then null else List.Count(List.Select(x, each _ > 0)) / List.Count(x),
                    FlowPct = (j as number) as nullable number =>
                        let
                            pairs = List.Select(g, each [v]{0} <> null and [v]{8 + j} <> null),
                            w = List.Sum(List.Transform(pairs, each [v]{0}))
                        in
                            if List.IsEmpty(pairs) or w = null or w <= 0 then null
                            else List.Sum(List.Transform(pairs, each [v]{8 + j})) / w,
                    has = List.AnyTrue(List.Transform(g, each List.NonNullCount(List.FirstN([v], 8)) > 0))
                in
                    List.Buffer({level, big, small, List.Count(g), capSum}
                    & List.Transform(List.Positions(Periods), Weighted)
                    & {UpRatio(0), UpRatio(1)}
                    & List.Transform(List.Positions(FlowKeys), FlowPct)
                    & {if has then "OK" else "데이터 없음", Now}),
            // 시가총액 합 내림차순(빈칸은 뒤), 같으면 이름순
            ByCap = (xs as list) as list => List.Sort(xs, (a, b) =>
                let ca = a[row]{4}, cb = b[row]{4}
                in if ca = null and cb = null then Value.Compare(a[name], b[name])
                   else if ca = null then 1
                   else if cb = null then -1
                   else if ca <> cb then Value.Compare(cb, ca)
                   else Value.Compare(a[name], b[name])),
            Bigs = List.Distinct(List.Transform(Rows, each [big])),
            BigAggs = List.Transform(Bigs, (b) =>
                let g = List.Select(Rows, each [big] = b) in [name = b, g = g, row = Agg("대테마", b, null, g)]),
            Out = List.Buffer(List.Combine(List.Transform(ByCap(BigAggs), (b) =>
                let
                    smalls = List.Distinct(List.RemoveNulls(List.Transform(b[g], each [small]))),
                    subs = List.Transform(smalls, (s) => [name = s, row = Agg("세부테마", b[name], s, List.Select(b[g], each [small] = s))])
                in
                    {b[row]} & List.Transform(ByCap(subs), each [row])))),
            // 칸마다 한 번씩 읽어 계산 오류를 여기서 올린다(목록·표 버퍼는 칸 안 오류를 올리지 않고 빈칸으로 적재하므로)
            Counted = List.Accumulate(Out, 0, (n, x) => n + List.Count(List.Select(x, each _ = null or _ <> null)))
        in
            if Counted >= 0 then Table.TransformColumnTypes(#table(Cols, Out), Types) else error "평가 실패",
    Tried = try Table.Buffer(if Already then KeptOwn else Compute()),
    Reason = if not Tried[HasError] then null
             else let m = Tried[Error][Message] in if m = null or m = "" then "사유 불명" else m,
    KeptOnFail = if Tried[HasError] then fnFallback("tblThemeAgg", Cols, Reason) else null,
    Warn = #table(Cols, {List.Repeat({null}, NCols - 2) & {"오류: " & Text.Start(if Reason = null then "" else Reason, 120), DateTime.LocalNow()}}),
    // fnFallback·경고 행은 형식이 없는 표 → 열 형식을 다시 지정(조회시각이 일련번호로 보이지 않게)
    Final =
        if not Tried[HasError] then Tried[Value]
        else Table.TransformColumnTypes(if Table.RowCount(KeptOnFail) > 0 then KeptOnFail else Warn, Types)
in
    Final
