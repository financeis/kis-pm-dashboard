// fnARun : 종목분석(tblA_*) 쿼리 공통 실행기 — 실행 모드·분석코드 검사·실패 시 직전 표 유지 (2026-10-01)
//  - 사용: T_A_X = fnARun("tblA_X", Types, (code as text, now as datetime) as table => <그 종목의 표>)
//    Types = {{"열", type …}, …} — 표의 열 이름·순서·형식. 모든 경로(성공·직전 표·Seed·빈 표·상태 1행)가 같은 열 구성.
//    main은 analysis 모드에서 분석코드가 올바를 때만 한 번 불린다(code = 정규화한 6자리, now = 이 쿼리의 조회시각).
//  - KIS 호출은 tblRunCtl의 mode = analysis([조회])일 때만. 그 밖의 모드(build·quick·full·sector·빈칸·알 수 없는 값)는
//    main을 부르지 않고(M 지연 평가 → 호출 0) 자기 표(직전 [조회] 결과) → 없으면 정적 표 tblA_XSeed → 없으면 빈 표를 그대로 돌려준다.
//  - analysis 모드에서 분석코드(이름 정의 — 종목분석 시트의 입력 칸)가 비었거나 형식이 틀리면 호출 없이 직전 표(→ Seed)의
//    상태를 모두 "오류: 종목코드 확인"으로 바꿔 돌려준다(데이터·조회시각은 그대로). 직전 표도 Seed도 없으면 그 상태의 1행.
//    형식 = fnClassify와 같은 정규화(앞뒤 공백 제거·대문자·숫자만 6자리 미만이면 앞 0 채움) 뒤 6자리 [0-9A-Z]. 상장 여부는 보지 않는다
//    (형식만 맞는 없는 코드는 각 표가 KIS 응답대로 '데이터 없음'이 되고 tblA_Info가 "오류: 종목코드 확인(…)"을 남긴다).
//  - main 실패(토큰·네트워크·호출 제한·응답 오류·셀 값 오류): 직전 표(→ Seed)를 유지하고 상태를 fnFallback과 같은 문구
//    "이전 데이터(갱신 실패: 사유 120자)"로. 직전 표도 Seed도 없으면 "오류: 사유" 1행 — 빈 표로는 실패가 보이지 않기 때문(T_Class와 같은 규칙).
//  - main 결과는 try 안에서 모든 행·모든 칸을 하나씩 try로 평가해 강제한다 — M 지연 평가 때문에 행·칸 안에서 난 API·변환 오류는
//    Table.Buffer·List.Buffer로는 올라오지 않고(빈 행으로 적재됨) 그대로 두면 표에 남으므로, 오류가 하나라도 있으면 위 실패 경로로 보낸다.
//  - 재호출 방지: Excel은 새로 고침 직후 쿼리를 백그라운드로 한 번 더 평가한다(동시 작업이 요청 수를 세는 목 서버로 실측).
//    tblRunCtl의 started(버튼 시작 시각 — 매크로가 기록)가 있고, 자기 표의 조회시각 최댓값 ≥ started이며 자기 표의 종목코드가
//    지금 분석코드뿐이면 이번 [조회]에서 이미 채운 것으로 보고 호출 없이 자기 표를 그대로 돌려준다(started가 비면 검사하지 않음).
//    KIS를 부르는 계산은 모두 main(함수 값 — 평가해도 호출 없음) 안에 있고 analysis 모드·올바른 코드·위 검사 통과 때만 불린다.
//  - 직전 표를 Excel에서 다시 읽으면 열 형식이 없어(any) 날짜·조회시각이 일련번호로 보이므로 모든 경로에서 Types를 다시 붙인다.
//    모든 칸이 빈 행(빌더가 빈 정적 표에 두는 자리 행)은 직전 표·Seed에서 버린다.
(tableName as text, types as list, main as function) as table =>
let
    Now = DateTime.LocalNow(),
    Cols = List.Transform(types, each _{0}),
    Txt = (v as any) as nullable text =>
        if v = null then null
        else let t = try Text.Trim(Text.From(v)) otherwise null in if t = null or t = "" then null else t,
    Typed = (t as table) as table => Table.TransformColumnTypes(Table.SelectColumns(t, Cols, MissingField.UseNull), types),

    // 통합문서의 표(자기 표 → Seed): 열이 모두 있을 때만, 모든 칸이 빈 행은 버림
    Stored = (name as text) as nullable table =>
        let
            t = try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null,
            ok = t <> null and List.ContainsAll(Table.ColumnNames(t), Cols),
            rows = if ok then Table.SelectRows(Table.SelectColumns(t, Cols), each List.NonNullCount(Record.FieldValues(_)) > 0) else null
        in
            if rows = null or Table.IsEmpty(rows) then null else rows,
    Own = Stored(tableName),
    Prev = if Own <> null then Own else Stored(tableName & "Seed"),

    // 상태만 바꾼 직전 표, 없으면 그 상태의 1행(종목코드는 비움)
    Marked = (status as text) as table =>
        if Prev <> null then Table.TransformColumns(Prev, {{"상태", each status, type text}})
        else fnAOneRow(types, null, status, Now),

    // 실행 모드 (tblRunCtl 키 mode — 없거나 빈칸이면 build), 버튼 시작 시각 started
    Ctl = try Excel.CurrentWorkbook(){[Name = "tblRunCtl"]}[Content] otherwise null,
    CtlVals = (key as text) as list =>
        if Ctl = null or not List.ContainsAll(Table.ColumnNames(Ctl), {"키", "값"}) then {}
        else Table.SelectRows(Ctl, each Txt([키]) = key)[값],
    ModeVals = List.Transform(CtlVals("mode"), Txt),
    Mode = if List.IsEmpty(ModeVals) or ModeVals{0} = null then "build" else Text.Lower(ModeVals{0}),
    AsDateTime = (v as any) as nullable datetime =>
        if v = null then null
        else if v is datetime then v
        else if v is date then DateTime.From(v)
        else if v is number then (try DateTime.From(v) otherwise null)
        else if v is text then (try DateTime.FromText(Text.Trim(v)) otherwise null)
        else null,
    Started = let s = CtlVals("started") in if List.IsEmpty(s) then null else AsDateTime(s{0}),

    // 분석코드 (이름 정의가 가리키는 셀 1개 — 머리글 없이 읽히므로 첫 열 첫 값)
    Raw = try
            let t = Excel.CurrentWorkbook(){[Name = "분석코드"]}[Content]
            in Table.Column(t, Table.ColumnNames(t){0}){0}
          otherwise null,
    Up = let t = Txt(Raw) in if t = null then null else Text.Upper(t),
    Padded = if Up <> null and Text.Length(Up) < 6 and Text.Select(Up, {"0".."9"}) = Up then Text.PadStart(Up, 6, "0") else Up,
    Code = if Padded <> null and Text.Length(Padded) = 6 and Text.Select(Padded, {"0".."9", "A".."Z"}) = Padded then Padded else null,

    // 이번 [조회](started 이후)에 이 종목으로 이미 채운 자기 표인지 — 백그라운드 재평가에서 KIS 재호출 방지
    LastStamp = if Own = null then null else List.Max(List.RemoveNulls(List.Transform(Table.Column(Own, "조회시각"), AsDateTime))),
    OwnCodes = if Own = null then {} else List.Distinct(List.RemoveNulls(List.Transform(Table.Column(Own, "종목코드"), Txt))),
    Already = Started <> null and LastStamp <> null and LastStamp >= Started and OwnCodes = {Code},

    // main 결과의 모든 행·칸을 try로 한 번씩 평가(행 자체가 오류인 경우 포함) → 오류가 있으면 첫 사유로 오류,
    // 없으면 평가된 값으로 표를 다시 만듦(같은 계산을 두 번 하지 않도록 행 우선 순서의 평가 결과를 열 수씩 잘라 씀)
    Forced = (t as table) as table =>
        let
            Names = Table.ColumnNames(t),
            Flat = List.Buffer(List.Combine(List.Transform(Table.ToRows(t), (row) =>
                        let r = try row in if r[HasError] then {r} else List.Transform(r[Value], (v) => try v)))),
            Errs = List.Select(Flat, each [HasError]),
            Msg = if List.IsEmpty(Errs) then null else Txt(Errs{0}[Error][Message])
        in
            if not List.IsEmpty(Errs) then error Error.Record("셀 값 오류", if Msg = null then "셀 값 오류" else Msg)
            else #table(Names, List.Transform(List.Split(Flat, List.Count(Names)), (chunk) => List.Transform(chunk, each [Value]))),

    Run = () as table =>
        let
            Tried = try Table.Buffer(Typed(Forced(Typed(main(Code, Now))))),
            Reason = if not Tried[HasError] then null
                     else let m = Txt(Tried[Error][Message]) in if m = null then "사유 불명" else Text.Start(m, 120)
        in
            if not Tried[HasError] then Tried[Value]
            else if Prev <> null then Marked("이전 데이터(갱신 실패: " & Reason & ")")
            else Marked("오류: " & Reason),

    Result =
        if Mode <> "analysis" then (if Prev <> null then Prev else #table(Cols, {}))
        else if Code = null then Marked("오류: 종목코드 확인")
        else if Already then Own
        else Run()
in
    Typed(Result)
