// T_EstSnap : KIS 추정 Fwd EPS 스냅샷 — 영구 이력 (2026-10-01) → tblEstSnap  [KIS 호출 없음 — mode = full일 때만 추가·덮어쓰기]
//  - 열: 일자, 종목코드, FwdEPS, 종가, FwdPER, 추정일, 상태, 조회시각
//      일자     tblPxMetrics 가격기준일 = 그 종목의 가장 최근 봉 날짜(세션 일자; 거래정지 등이면 마지막 세션보다 이를 수 있음)
//      FwdEPS   fnFwdEps(tblEst FY1, FY1_EPS, FY2_EPS, 일자) — 가중치 r은 일자가 속한 달로 센다(fnFwdEps 식, 12월 결산·10월 → r = 3)
//      종가     tblPxMetrics 현재가(그 봉의 종가 — 장중에 [전체]를 누르면 그 시점 가격, 장 마감 뒤 다시 누르면 같은 일자로 덮음)
//      FwdPER   종가 ÷ FwdEPS (FwdEPS ≤ 0이거나 종가가 없으면 빈칸)
//      추정일   tblEst 추정일(output1.estdate — 그 Fwd EPS가 나온 KIS 보고서 일자)
//      상태     OK(성공한 실행은 모든 행을 OK로 — 저장소 쿼리 관례, docs/standards.md)
//      조회시각 그 행을 쓴(덮은) 시각
//  - 동작 (mode = full — [전체]에서 T_PxMetrics·T_Est 다음):
//    1) 직전 표 = 자기 표 tblEstSnap → 비었으면 정적 표 tblEstSnapSeed(이관 때 빌더가 채움; 이력 CSV에서 복원할 수도 있음) → 없으면 빈 표.
//       일자·종목코드가 없거나 잘못된 행은 버리고, 같은 (일자, 종목코드)가 여러 번 있으면 처음 행만 쓴다.
//    2) 이번 행 = tblEst에서 상태가 OK이고 FY1_EPS가 있는 종목 × tblPxMetrics의 그 종목 행(가격기준일이 있고 상태가 '이전 데이터'로
//       시작하지 않는 행 — 지표 표 전체가 직전 값이면 그 가격기준일은 이번 실행 것이 아니므로 쓰지 않음; 저장소 오류로 '오류:'인 행은
//       가격기준일과 현재가가 서로 맞는 한 쌍이라 그대로 씀). fnFwdEps가 null이면(가격기준일이 FY1 회계연도 시작 전 등) 행을 만들지 않는다.
//       tblEst 상태가 '오류:'(이번 조회 실패 — 직전 값 표시 중)이거나 '이전 데이터'(T_Est 전체 실패)인 종목도 행을 만들지 않는다:
//       스냅샷에는 그날 실제로 받은 추정만 남기고, 빈 날은 대회종목 표가 가까운 스냅샷(3세션 이내)이나 재구성으로 메운다(docs/business-rules.md의 'KIS 추정과 Fwd PER').
//    3) (일자, 종목코드)가 같은 직전 행은 이번 행으로 덮는다(같은 날 [전체]를 여러 번 누르면 마지막 값 — 행 수 그대로).
//       나머지 직전 행은 값과 조회시각을 그대로 두고 상태만 OK로 맞춘다(직전 실패의 '이전 데이터' 표시 정리).
//       행은 지우지 않는다 — 페이지 행에서 빠진 종목의 이력도 남는다.
//    4) 정렬: 일자 오름차순 → 종목코드 오름차순.
//  - 같은 [전체] 실행 안에서 다시 평가되면(자기 표의 가장 늦은 조회시각 ≥ tblRunCtl started) 계산하지 않고 자기 표를 돌려준다
//    (새로 고침 직후 Excel의 백그라운드 재평가에서 같은 행을 다시 덮어 조회시각만 바뀌는 일을 막음). started가 비면 확인하지 않는다.
//  - mode가 full이 아니면(build·quick·sector·analysis·빈칸·알 수 없는 값) 직전 표(자기 표 → Seed → 빈 표)를 그대로 돌려준다(상태 포함,
//    빈 표도 열은 모두 갖춤). tblEst·tblPxMetrics도 읽지 않는다. 이 쿼리는 어느 모드에서도 KIS를 부르지 않는다.
//  - 실패: tblEst나 tblPxMetrics 표가 통합문서에 없거나 계산 오류 → 직전 표 유지(fnFallback — 상태 '이전 데이터(갱신 실패: 사유)',
//    열 형식 다시 지정). 표는 있는데 비어 있으면(첫 [전체] 전 등) 오류가 아니라 이번 행이 없을 뿐이다.
//  - 이 표는 재생성 때 이관되고, [전체]가 끝날 때마다 VBA 매크로가 history\est_snap.csv로 복사한다(이 쿼리는 하지 않음).
//  - 크기: 추정 EPS가 있는 종목(약 110) × 세션 수 — 1년에 약 2.8만 행.
let
    Cols = {"일자", "종목코드", "FwdEPS", "종가", "FwdPER", "추정일", "상태", "조회시각"},
    Types = {{"일자", type date}, {"종목코드", type text}, {"FwdEPS", type number}, {"종가", type number},
             {"FwdPER", type number}, {"추정일", type date}, {"상태", type text}, {"조회시각", type datetime}},
    Schema = type table [일자 = date, 종목코드 = text, FwdEPS = nullable number, 종가 = nullable number,
                         FwdPER = nullable number, 추정일 = nullable date, 상태 = nullable text, 조회시각 = nullable datetime],

    // 셀 값 → 형식 있는 값 (자기 표·Seed·tblEst·tblPxMetrics는 Excel에서 다시 읽어 열 형식이 없다; 변환할 수 없는 값은 null)
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
    Records = (t as table, cols as list) as list => Table.ToRecords(Table.SelectColumns(t, cols, MissingField.UseNull)),
    // 스냅샷 키 (일자, 종목코드) → "yyyyMMdd|코드"
    Key = (d as date, c as text) as text => Date.ToText(d, [Format = "yyyyMMdd", Culture = "en-US"]) & "|" & c,

    // 직전 표 → 행 목록(값 순서 = Cols: 0 일자, 1 종목코드, 2 FwdEPS, 3 종가, 4 FwdPER, 5 추정일, 6 상태, 7 조회시각)
    ToRows = (t as nullable table) as list =>
        if t = null then {}
        else
            let
                rows = List.Transform(Records(t, Cols), (r) => {Day(F(r, "일자")), Code(F(r, "종목코드")), Num(F(r, "FwdEPS")),
                    Num(F(r, "종가")), Num(F(r, "FwdPER")), Day(F(r, "추정일")), Txt(F(r, "상태")), Stamp(F(r, "조회시각"))}),
                ok = List.Select(rows, each _{0} <> null and _{1} <> null)
            in
                List.Distinct(ok, each Key(_{0}, _{1})),
    OwnRows = ToRows(ReadTable("tblEstSnap")),
    PrevRows = List.Buffer(if List.Count(OwnRows) > 0 then OwnRows else ToRows(ReadTable("tblEstSnapSeed"))),

    // 실행 제어 tblRunCtl(키, 값): mode(없거나 빈칸·알 수 없는 값은 build와 같음 — full일 때만 추가·덮어쓰기), started
    RunCtl = ReadTable("tblRunCtl"),
    RunRows = if RunCtl = null then {} else Records(RunCtl, {"키", "값"}),
    RunValue = (key as text) as any =>
        let hit = List.Select(RunRows, (r) => let k = Txt(F(r, "키")) in k <> null and Text.Lower(k) = key)
        in if List.IsEmpty(hit) then null else F(hit{0}, "값"),
    ModeText = Txt(RunValue("mode")),
    Mode = if ModeText = null then "build" else Text.Lower(ModeText),
    Started = Stamp(RunValue("started")),
    LastStamp = List.Max(List.RemoveNulls(List.Transform(OwnRows, each _{7}))),
    Already = Started <> null and LastStamp <> null and LastStamp >= Started,

    // 계산은 함수 단계 안에서만(맨 위 단계를 따로 평가해도 계산하지 않음 — T_Target·T_PxMetrics와 같은 이유)
    Snapshot = () as table =>
        let
            Now = DateTime.LocalNow(),
            EstT = ReadTable("tblEst"),
            PxT = ReadTable("tblPxMetrics"),
            Missing = List.Select({{"tblEst", EstT}, {"tblPxMetrics", PxT}}, each _{1} = null)
        in
            if not List.IsEmpty(Missing) then
                error Error.Record("스냅샷 원천 없음", Text.Combine(List.Transform(Missing, each _{0}), ", ") & " 표가 통합문서에 없습니다.")
            else
                let
                    // 이번 추정: 상태 OK이고 FY1_EPS가 있는 종목 {코드, 추정일, FY1, FY1_EPS, FY2_EPS}, 같은 코드는 처음 행만
                    EstAll = List.Transform(Records(EstT, {"종목코드", "추정일", "FY1", "FY1_EPS", "FY2_EPS", "상태"}), (r) =>
                                {Code(F(r, "종목코드")), Day(F(r, "추정일")), Txt(F(r, "FY1")), Num(F(r, "FY1_EPS")),
                                 Num(F(r, "FY2_EPS")), Txt(F(r, "상태"))}),
                    Est = List.Distinct(List.Select(EstAll, each _{0} <> null and _{3} <> null and _{5} = "OK"), each _{0}),
                    // 가격: 종목별 {코드, 가격기준일, 현재가} — 가격기준일이 있고 상태가 '이전 데이터'로 시작하지 않는 행, 같은 코드는 처음 행만
                    PxAll = List.Transform(Records(PxT, {"종목코드", "가격기준일", "현재가", "상태"}), (r) =>
                                {Code(F(r, "종목코드")), Day(F(r, "가격기준일")), Num(F(r, "현재가")), Txt(F(r, "상태"))}),
                    PxOk = List.Distinct(List.Select(PxAll, each _{0} <> null and _{1} <> null
                                and not (_{3} <> null and Text.StartsWith(_{3}, "이전 데이터"))), each _{0}),
                    Px = Record.FromList(PxOk, List.Transform(PxOk, each _{0})),
                    Fresh = List.RemoveNulls(List.Transform(Est, (e) =>
                        let
                            p = Record.FieldOrDefault(Px, e{0}, null),
                            d = if p = null then null else p{1},
                            eps = if d = null then null else fnFwdEps(e{2}, e{3}, e{4}, d),
                            px = if p = null then null else p{2}
                        in
                            if eps = null then null
                            else {d, e{0}, eps, px, if px <> null and eps > 0 then px / eps else null, e{1}, "OK", Now})),
                    New = List.Buffer(List.Distinct(Fresh, each Key(_{0}, _{1}))),
                    NewKeys = Record.FromList(List.Repeat({true}, List.Count(New)), List.Transform(New, each Key(_{0}, _{1}))),
                    // 덮지 않는 직전 행: 값·조회시각 그대로, 상태만 OK
                    Kept = List.Transform(List.Select(PrevRows, each not Record.HasFields(NewKeys, Key(_{0}, _{1}))),
                                each List.FirstN(_, 6) & {"OK", _{7}}),
                    All = Kept & New,
                    // 칸마다 한 번씩 읽어 계산 오류를 여기서 올린다(목록·표 버퍼는 칸 안 오류를 올리지 않고 빈칸으로 적재하므로)
                    Counted = List.Accumulate(All, 0, (n, x) => n + List.Count(List.Select(x, each _ = null or _ <> null))),
                    Out = Table.Sort(#table(Schema, All), {{"일자", Order.Ascending}, {"종목코드", Order.Ascending}})
                in
                    if Counted >= 0 then Out else error "평가 실패",

    Main = if Mode = "full" and not Already then Snapshot() else #table(Schema, PrevRows),
    // 실패 시 오류 대화상자 대신 직전 결과를 유지하고(상태 '이전 데이터(갱신 실패: 사유)'), 열 형식을 다시 붙여 날짜 서식을 지킨다
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblEstSnap", Cols, Tried[Error][Message]), Types)
            else Tried[Value]
in
    Final
