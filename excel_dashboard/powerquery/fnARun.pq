// fnARun : 종목분석(tblA_*) 쿼리 공통 실행기 — 실행 모드·분석코드 검사·실패 시 직전 표 유지 (spec R17, 2026-10-01 T15)
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
//  - main 결과는 Table.Buffer 뒤 오류 셀 검사(Table.SelectRowsWithErrors)로 모든 칸을 강제 평가한다 — M 지연 평가 때문에 셀 안에서 난
//    API·변환 오류가 표에 오류 칸으로 남지 않고 위 실패 경로로 가게 하려는 것.
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
    Prev = let p = Stored(tableName) in if p <> null then p else Stored(tableName & "Seed"),

    // 상태만 바꾼 직전 표, 없으면 그 상태의 1행(종목코드는 비움)
    Marked = (status as text) as table =>
        if Prev <> null then Table.TransformColumns(Prev, {{"상태", each status, type text}})
        else fnAOneRow(types, null, status, Now),

    // 실행 모드 (tblRunCtl 키 mode — 없거나 빈칸이면 build)
    Ctl = try Excel.CurrentWorkbook(){[Name = "tblRunCtl"]}[Content] otherwise null,
    ModeVals =
        if Ctl = null or not List.ContainsAll(Table.ColumnNames(Ctl), {"키", "값"}) then {}
        else List.Transform(Table.SelectRows(Ctl, each Txt([키]) = "mode")[값], Txt),
    Mode = if List.IsEmpty(ModeVals) or ModeVals{0} = null then "build" else Text.Lower(ModeVals{0}),

    // 분석코드 (이름 정의가 가리키는 셀 1개 — 머리글 없이 읽히므로 첫 열 첫 값)
    Raw = try
            let t = Excel.CurrentWorkbook(){[Name = "분석코드"]}[Content]
            in Table.Column(t, Table.ColumnNames(t){0}){0}
          otherwise null,
    Up = let t = Txt(Raw) in if t = null then null else Text.Upper(t),
    Padded = if Up <> null and Text.Length(Up) < 6 and Text.Select(Up, {"0".."9"}) = Up then Text.PadStart(Up, 6, "0") else Up,
    Code = if Padded <> null and Text.Length(Padded) = 6 and Text.Select(Padded, {"0".."9", "A".."Z"}) = Padded then Padded else null,

    // 첫 오류 칸의 메시지(대개 셀 안에서 평가된 API·설정 오류 그대로 — 예: kis_devlp.yaml 경로 오류)
    FirstError = (t as table) as text =>
        let
            r = Table.ToRecords(t){0},
            tries = List.Transform(Record.FieldNames(r), (n) => try Record.Field(r, n)),
            errs = List.Select(tries, each [HasError]),
            m = if List.IsEmpty(errs) then null else Txt(errs{0}[Error][Message])
        in
            if m = null then "셀 값 오류" else m,

    Run = () as table =>
        let
            Tried = try
                let
                    B = Table.Buffer(Typed(main(Code, Now))),
                    E = Table.SelectRowsWithErrors(B)
                in
                    if Table.IsEmpty(E) then B else error Error.Record("셀 값 오류", FirstError(E)),
            Reason = if not Tried[HasError] then null
                     else let m = Txt(Tried[Error][Message]) in if m = null then "사유 불명" else Text.Start(m, 120)
        in
            if not Tried[HasError] then Tried[Value]
            else if Prev <> null then Marked("이전 데이터(갱신 실패: " & Reason & ")")
            else Marked("오류: " & Reason),

    Result =
        if Mode <> "analysis" then (if Prev <> null then Prev else #table(Cols, {}))
        else if Code = null then Marked("오류: 종목코드 확인")
        else Run()
in
    Typed(Result)
