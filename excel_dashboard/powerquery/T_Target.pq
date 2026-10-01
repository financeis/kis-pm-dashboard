// T_Target : 목표주가 컨센서스 (spec R8·R20, 2026-10-01 T8) → tblTarget  [[전체] 버튼 전용 — mode = full일 때만 KIS 호출]
//  - 원천: 국내주식 종목투자의견[국내주식-188] FHKST663300C0 — fnTargetOpinions(날짜 분할로 기간 전체를 받음, 열 설명도 거기)
//  - 열: 종목코드, 목표가평균, 증권사수, 목표가평균_1M전, 목표가평균_3M전, 최근의견일, 최근증권사, 최근의견, 전체조사시각, 상태, 조회시각
//  계산 (기준일 D = 실행한 날짜, M = 설정 target_window_months, 기본 6)
//    - 목표가평균·증권사수 = fnTargetConsensus(의견, D, M): 증권사별로 (D − M개월, D] 안의 마지막 목표가(0·공란 제외)를 평균, 그 증권사 수.
//    - 목표가평균_1M전·_3M전 = 같은 규칙을 기준일 Date.AddMonths(D, -1)·Date.AddMonths(D, -3)에서 계산(기간 길이는 똑같이 M개월).
//      목표가 변화 1M·3M(= 목표가평균 ÷ _1M전·_3M전 − 1)과 괴리율(= 목표가평균 ÷ 현재가 − 1)은 대회종목 표(T16)가 계산한다.
//    - 조회 기간 = [Date.AddMonths(Date.AddMonths(D, -3), -M), D] — D−3개월 컨센서스의 기간 시작까지(M = 6이면 약 9개월).
//    - 최근의견일·최근증권사·최근의견 = 조회 기간에서 가장 최근 의견(같은 날이면 응답에서 먼저 온 행)의 일자·증권사·의견.
//      목표가가 없는 의견(NotRated 등)도 그대로 보여 준다.
//    - 기간 안에 목표가가 하나도 없으면 목표가 열은 빈칸이고 상태 '데이터 없음'(정상 값 — 이 종목군의 약 42%), 있으면 'OK'.
//  조사 범위 (spec R20, mode = full)
//    - 주간 = 페이지 행(fnPageRows) 전체 조회 후 모든 행의 전체조사시각 = 지금. 조건: 직전 표에 전체조사시각이 없음(첫 실행),
//      또는 그 날짜로부터 weekly_days(기본 7)일 이상 지남(날짜 차이로 셈 — 매일 누르는 시각이 달라도 같은 요일이면 돌게),
//      또는 설정 force_weekly = Y. (force_weekly를 N으로 되돌리는 일은 [전체] 매크로가 한다.)
//    - 매일 = 직전 행의 목표가평균이 있는 종목 + 직전 표에 없는 새 종목 + 직전 상태가 '오류'인 종목(실패한 조사를 다시 시도)만 조회.
//      나머지(직전 조사에서 목표가가 없던 종목)는 직전 행을 그대로 두되(조회시각 유지) 상태만 '데이터 없음'으로 맞춘다
//      — 직전 실패의 '이전 데이터' 표시가 남아 단계 판정이 실패로 보이지 않게. 전체조사시각은 직전 값 그대로(모든 행 같은 값).
//    - 행 = 지금의 페이지 행(순서도 fnPageRows). 페이지 행에서 빠진 종목의 직전 행은 버린다.
//    - 직전 표 = 자기 표 tblTarget → 없거나 비었으면 정적 표 tblTargetSeed(이관 때 빌더가 채움) → 없으면 빈 표.
//    - 같은 [전체] 실행 안에서 다시 평가되면(자기 표의 가장 늦은 조회시각이 tblRunCtl started 이후) 조회하지 않고 자기 표를 돌려준다
//      — 새로 고침 직후 Excel의 백그라운드 재평가가 매일 조사만큼 KIS를 또 부르는 것을 막음(아래 Main 주석). started가 비면 확인 안 함.
//    - 시험용 시계(tblRunCtl now_override)는 쓰지 않는다(plan §0: T_Sessions·T_PxStore 전용). 주간 규칙 시험은
//      tblTarget의 전체조사시각 셀을 바꾸거나 force_weekly·weekly_days 설정으로 한다.
//  mode가 full이 아니면(build·quick·sector·analysis·빈칸·알 수 없는 값) KIS를 부르지 않고 직전 표(자기 표 → Seed → 빈 표)를
//    그대로 돌려준다(상태도 그대로, 빈 표도 열은 모두 갖춤). 설정·페이지 행도 읽지 않는다.
//  실패: 종목별 조회 오류 → 그 종목 행만 '오류: 사유'(직전 값과 조회시각 유지; 직전 행이 없으면 값은 빈칸, 조회시각은 지금).
//    조회한 종목이 모두 실패하면 오류로 올려 직전 표를 유지(fnFallback — 상태 '이전 데이터(갱신 실패: 사유)', 전체조사시각도 그대로라
//    다음 [전체]에서 주간 조사를 다시 함). 직전 표를 다시 읽은 값은 열 형식이 없어 날짜가 일련번호로 보이므로 형식을 다시 지정한다.
//  호출 수: 종목당 1회(조회 기간의 의견이 100건을 넘는 대형주 — 삼성전자·SK하이닉스 등 — 는 2회). 주간 약 500~560회,
//    매일은 목표가가 있는 종목만이라 약 300회(이 종목군의 58%).
let
    Cols = {"종목코드", "목표가평균", "증권사수", "목표가평균_1M전", "목표가평균_3M전", "최근의견일", "최근증권사", "최근의견",
            "전체조사시각", "상태", "조회시각"},
    Types = {{"종목코드", type text}, {"목표가평균", type number}, {"증권사수", type number}, {"목표가평균_1M전", type number},
             {"목표가평균_3M전", type number}, {"최근의견일", type date}, {"최근증권사", type text}, {"최근의견", type text},
             {"전체조사시각", type datetime}, {"상태", type text}, {"조회시각", type datetime}},
    Schema = type table [종목코드 = text, 목표가평균 = nullable number, 증권사수 = nullable number, 목표가평균_1M전 = nullable number,
                         목표가평균_3M전 = nullable number, 최근의견일 = nullable date, 최근증권사 = nullable text, 최근의견 = nullable text,
                         전체조사시각 = nullable datetime, 상태 = nullable text, 조회시각 = nullable datetime],

    // 셀 값 → 형식 있는 값 (직전 표·Seed는 Excel에서 다시 읽어 열 형식이 없다; 변환할 수 없는 값은 null)
    Txt = (v as any) as nullable text =>
        if v = null then null else let t = try Text.Trim(Text.From(v)) otherwise null in if t = null or t = "" then null else t,
    Num = (v as any) as nullable number =>
        if v = null then null else if v is number then v else if v is text then fnNum(v) else null,
    Day = (v as any) as nullable date =>
        if v = null then null
        else if v is date then v
        else if v is datetime then Date.From(v)
        else if v is number then (try Date.From(v) otherwise null)
        else if v is text then
            (let t = Text.Trim(v) in
                if Text.Length(t) = 8 and Text.Select(t, {"0".."9"}) = t then fnYmd(t)
                else try Date.FromText(Text.Start(t, 10), "en-US") otherwise null)
        else null,
    Stamp = (v as any) as nullable datetime =>
        if v = null then null
        else if v is datetime then v
        else if v is date then DateTime.From(v)
        else if v is number then (try DateTime.From(v) otherwise null)
        else if v is text then (try DateTime.FromText(Text.Trim(v), "en-US") otherwise null)
        else null,
    // 종목코드 정규화(fnClassify와 같은 규칙: 대문자, 숫자만 6자리 미만이면 앞 0 채움, 6자리 [0-9A-Z]만 인정)
    Code = (v as any) as nullable text =>
        let
            t0 = Txt(v),
            t1 = if t0 = null then null else Text.Upper(t0),
            t2 = if t1 <> null and Text.Length(t1) < 6 and Text.Select(t1, {"0".."9"}) = t1 then Text.PadStart(t1, 6, "0") else t1
        in
            if t2 <> null and Text.Length(t2) = 6 and Text.Select(t2, {"0".."9", "A".."Z"}) = t2 then t2 else null,
    F = (r as record, k as text) as any => try Record.Field(r, k) otherwise null,
    ReadTable = (name as text) as nullable table => try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null,

    // 직전 표 → 행 목록(값 순서 = Cols). 종목코드가 없거나 잘못된 행은 버리고, 같은 코드는 처음 행만
    ToRows = (t as nullable table) as list =>
        if t = null then {}
        else
            let
                recs = Table.ToRecords(Table.SelectColumns(t, Cols, MissingField.UseNull)),
                rows = List.Transform(recs, (r) => {Code(F(r, "종목코드")), Num(F(r, "목표가평균")), Num(F(r, "증권사수")),
                    Num(F(r, "목표가평균_1M전")), Num(F(r, "목표가평균_3M전")), Day(F(r, "최근의견일")), Txt(F(r, "최근증권사")),
                    Txt(F(r, "최근의견")), Stamp(F(r, "전체조사시각")), Txt(F(r, "상태")), Stamp(F(r, "조회시각"))})
            in
                List.Distinct(List.Select(rows, each _{0} <> null), each _{0}),
    OwnRows = ToRows(ReadTable("tblTarget")),
    PrevRows = List.Buffer(if List.Count(OwnRows) > 0 then OwnRows else ToRows(ReadTable("tblTargetSeed"))),

    // 실행 제어 tblRunCtl(키, 값): mode — 없거나 빈칸·알 수 없는 값은 build와 같음(full일 때만 조회),
    // started — 버튼 실행 시작 시각([전체] 매크로가 씀)
    RunCtl = ReadTable("tblRunCtl"),
    RunRows = if RunCtl = null then {} else Table.ToRecords(Table.SelectColumns(RunCtl, {"키", "값"}, MissingField.UseNull)),
    RunValue = (key as text) as any =>
        let hit = List.Select(RunRows, (r) => let k = Txt(F(r, "키")) in k <> null and Text.Lower(k) = key)
        in if List.IsEmpty(hit) then null else F(hit{0}, "값"),
    ModeText = Txt(RunValue("mode")),
    Mode = if ModeText = null then "build" else Text.Lower(ModeText),
    // 같은 버튼 실행 안에서 다시 평가될 때(새로 고침 직후 Excel이 쿼리를 백그라운드에서 한 번 더 평가함 — 아래 Main 주석)는
    // 조회하지 않고 방금 적재한 자기 표를 돌려준다: 자기 표의 가장 늦은 조회시각이 이번 실행의 started 이후면 이미 조사한 것.
    // started가 비어 있으면(하네스 시험 등) 이 확인은 하지 않는다.
    Started = Stamp(RunValue("started")),
    LastStamp = List.Max(List.RemoveNulls(List.Transform(OwnRows, each _{10}))),
    Surveyed = Started <> null and LastStamp <> null and LastStamp >= Started,

    Survey = () as table =>
        let
            Now = DateTime.LocalNow(),
            D = Date.From(Now),
            // 설정(없거나 잘못된 값은 기본값): weekly_days 7, force_weekly N, target_window_months 6
            S = try fnSettings() otherwise [],
            Setting = (k as text) as any => try Record.FieldOrDefault(S, k, null) otherwise null,
            WeeklyDays = let v = Num(Setting("weekly_days")) in if v = null or v < 0 then 7 else v,
            Force = let v = Txt(Setting("force_weekly")) in v <> null and Text.Upper(v) = "Y",
            Months = let v = Num(Setting("target_window_months")) in if v = null or v < 1 then 6 else Number.Round(v),
            FromDate = Date.AddMonths(Date.AddMonths(D, -3), -Months),

            Codes = List.Buffer(List.Distinct(List.RemoveNulls(List.Transform(fnPageRows()[종목코드], Code)))),
            Prev = Record.FromList(PrevRows, List.Transform(PrevRows, each _{0})),
            LastFull = List.Max(List.RemoveNulls(List.Transform(PrevRows, each _{8}))),
            Weekly = Force or LastFull = null or Duration.Days(D - Date.From(LastFull)) >= WeeklyDays,
            FullAt = if Weekly then Now else LastFull,
            Needs = (c as text) as logical =>
                let p = Record.FieldOrDefault(Prev, c, null)
                in Weekly or p = null or p{1} <> null or (p{9} <> null and Text.StartsWith(p{9}, "오류")),
            Targets = List.Buffer(List.Select(Codes, Needs)),

            // 종목 하나: 의견 조회 → D·D−1M·D−3M 컨센서스, 증권사 수, 최근 의견
            Fetch = (c as text) as list =>
                let
                    R = fnTargetOpinions(c, FromDate, D),
                    C0 = fnTargetConsensus(R, D, Months),
                    C1 = fnTargetConsensus(R, Date.AddMonths(D, -1), Months),
                    C3 = fnTargetConsensus(R, Date.AddMonths(D, -3), Months),
                    Latest = Table.FirstN(Table.Sort(R, {{"일자", Order.Descending}, {"순번", Order.Ascending}}), 1),
                    L = if Table.RowCount(Latest) = 0 then null else Latest{0},
                    Vals = {C0[목표가평균], C0[증권사수], C1[목표가평균], C3[목표가평균],
                        if L = null then null else L[일자], if L = null then null else L[증권사], if L = null then null else L[의견]},
                    // M 지연 평가: 목록 항목·레코드 필드는 읽을 때 평가되고, 항목 오류는 List.Buffer로도 올라오지 않는다(항목 오류로
                    // 남아 셀만 빈칸이 됨 — 2026-10-01 모의 서버 시험에서 확인). 그래서 응답 표의 행 수를 먼저 세고(API 오류가 여기서 남)
                    // 값마다 한 번씩 비교해, 오류를 이 함수 안 — 아래 종목별 try 안 — 에서 올린다.
                    Forced = List.Accumulate(Vals, Table.RowCount(R), (n, v) => if v = null then n else n + 1)
                in
                    if Forced >= 0 then Vals else error "평가 실패",
            // 종목마다 한 번만 호출되도록 결과 목록을 버퍼로 고정(레코드 필드는 처음 읽을 때 한 번 평가)
            Results = List.Buffer(List.Transform(Targets, (c) => [code = c, r = try Fetch(c)])),
            Failed = List.Select(Results, each [r][HasError]),
            Reason = (e as record) as text =>
                let m = try Text.Trim(Text.From(e[Message])) otherwise "" in if m = "" then "조회 실패" else Text.Start(m, 200),
            // 조회한 종목이 모두 실패하면 오류로 올려 직전 표를 유지(fnFallback), 일부만 실패하면 그 종목만 오류 행
            Checked = if List.Count(Results) > 0 and List.Count(Failed) = List.Count(Results)
                      then error Error.Record("목표주가 조회 실패", Reason(Failed{0}[r][Error])) else Results,
            Done = Record.FromList(Checked, List.Transform(Checked, each [code])),
            Row = (c as text) as list =>
                let
                    p = Record.FieldOrDefault(Prev, c, null),
                    d = Record.FieldOrDefault(Done, c, null)
                in
                    if d = null then
                        // 매일 실행에서 조회하지 않은 종목(직전 조사에서 목표가 없음): 직전 행 유지, 상태만 정상 값으로
                        {c} & List.Range(p, 1, 7) & {FullAt, if p{1} = null then "데이터 없음" else "OK", p{10}}
                    else if d[r][HasError] then
                        (if p = null then {c, null, null, null, null, null, null, null} else {c} & List.Range(p, 1, 7))
                        & {FullAt, "오류: " & Reason(d[r][Error]), if p = null or p{10} = null then Now else p{10}}
                    else
                        let v = d[r][Value] in {c} & v & {FullAt, if v{0} = null then "데이터 없음" else "OK", Now},
            Rows = List.Buffer(List.Transform(Codes, Row)),
            // 행 목록의 항목 오류(모두 실패 오류 포함)는 표에 '오류 행'으로만 남고 Table.Buffer로도 올라오지 않으므로(빈 행이 적재됨 —
            // 2026-10-01 모의 서버 시험에서 확인) 행마다 한 번씩 읽어 여기서 오류를 올린다 → 아래 Tried가 잡아 직전 표를 유지
            Counted = List.Accumulate(Rows, List.Count(Checked), (n, x) => n + List.Count(x))
        in
            if Counted >= 0 then #table(Schema, Rows) else error "평가 실패",

    // Excel은 새로 고침이 끝난 직후 쿼리를 백그라운드에서 한 번 더 평가한다 — 최상위 단계들도 따로 평가한다(2026-10-01 하네스 +
    // 모의 서버 실측, 동기·백그라운드 새로 고침 모두). 이 평가는 Power Query 캐시를 써서 전에 백그라운드에서 받은 요청은 서버까지
    // 가지 않지만, 그날 처음 보는 요청(날짜가 바뀌면 전부)은 그대로 나간다.
    //  ① 조사를 값으로 된 단계로 두었을 때는 build 모드 새로 고침만으로도 페이지 행 전체 조회(11건)가 나갔다 → 조사는 함수로 두고
    //     full일 때만 부른다(함수 값인 단계는 평가해도 호출하지 않음).
    //  ② full 새로 고침 뒤에는 매일 조사만큼(목표가가 있는 종목) 중복 호출이 나갔다 → 위 Surveyed 확인으로 막는다.
    //     (조회가 모두 실패해 직전 표로 돌아간 경우엔 조회시각이 그대로라 막지 못하고 한 번 더 시도한다.)
    Main = if Mode = "full" and not Surveyed then Survey() else #table(Schema, PrevRows),
    // 실패(토큰·네트워크·설정·모두 실패 등) 시 오류 대화상자 대신 직전 결과를 유지하고, 열 형식을 다시 붙여 날짜 서식을 지킨다
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblTarget", Cols, Tried[Error][Message]), Types)
            else Tried[Value]
in
    Final
