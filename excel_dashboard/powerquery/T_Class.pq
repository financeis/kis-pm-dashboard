// T_Class : 업종 분류표 — VALUESearch 내보내기 파일의 NICS 3단계·KSIC 세분류·주요상품·기업집단 (2026-10-01) → tblClass
//  - 원천: 설정 vs_path가 가리키는 수집기업_valuesearch.xlsx (시트 Sheet2, 1행 머리글). 파일은 읽기만 한다.
//  - KIS를 호출하지 않으므로 실행 모드(tblRunCtl mode)를 보지 않고 새로 고칠 때마다 파일을 읽는다
//    → 빌드(build 모드)와 [전체] 모두에서 갱신된다. '모두 새로 고침'에서 빼는 일은 빌더(연결 속성) 몫.
//  - 행 = 파일에서 종목코드가 있는 모든 행(ETF 등 포함). 대회 종목만이 아니라 종목DB(tblUniverse)가 시장 전체를 분류할 수 있게.
//  - 종목코드: 파일 형식 "A######" → 앞의 A를 뗀 6자리 텍스트(문자 포함 코드 예: A0126Z0 → 0126Z0). 같은 코드가 여럿이면 첫 행.
//  - NICS 3단계 (코드 N + 5자리, 파일 값 "N#####.이름"):
//      대분류 코드 = 앞 3자리 + "000", 업종 코드 = 앞 4자리 + "00", 세부 = 원래 코드.
//      대분류·업종 이름은 고정 코드표(fnNicsCodes = data/nics_codes.csv), 세부 이름은 파일의 이름.
//      업종 수준까지만 분류(N###00) → 세부 = 업종 이름 / 대분류 수준까지만(N##000) → 업종 = 세부 = "대분류이름 (기타)".
//      코드표에 없는 대분류·업종 코드(예: N00000.NISC — VALUESearch의 미분류 표시)는 이름을 정할 수 없어 세 단계를 비운다
//      (NICS코드 열에는 원래 코드를 남겨 이유를 볼 수 있게 함).
//  - 분류출처: NICS 세 단계를 만들었으면 "NICS", 아니면 공란. KRX 지수업종 대체(분류출처 KRX)는 fnClassify가 정한다
//    (KRX업종은 종목DB 쪽 값이고, 이 표는 파일만 반영하므로).
//  - KSIC 세분류·주요상품·기업집단: 파일 값 그대로(앞뒤 공백만 제거). KSIC는 "C26110.전자집적회로 제조업"처럼 코드를 포함한다.
//  - 실패(설정 없음·파일 없음·시트/열 이름 불일치): 직전 표를 그대로 두고 상태에 "이전 데이터(갱신 실패: 사유)"(fnFallback).
//    직전 표가 없으면(첫 실행) 종목코드가 빈 1행에 상태 "오류: 사유"를 남긴다 — 빈 표로는 실패가 보이지 않기 때문
//    (버튼 단계 판정이 이 행을 오류로 센다. fnClassify는 이 행을 무시하고, tblClass를 직접 읽는 쿼리도 종목코드가 빈 행은 건너뛴다).
//    두 경우 모두 열 형식을 다시 붙여 조회시각이 일련번호로 보이지 않게 한다.
let
    Columns = {"종목코드", "NICS코드", "NICS 대분류", "NICS 업종", "NICS 세부", "KSIC 세분류", "주요상품", "기업집단", "분류출처", "상태", "조회시각"},
    Main =
        let
            Now = DateTime.LocalNow(),
            Txt = (v as any) as nullable text =>
                if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,
            Digits = {"0".."9"},
            CodeChars = {"0".."9", "A".."Z"},

            // 설정 vs_path (탐색기 '경로로 복사'가 붙이는 큰따옴표도 허용)
            PathIn = Txt(Record.FieldOrDefault(fnSettings(), "vs_path", null)),
            Path = if PathIn = null then null else Text.Trim(PathIn, """"),
            // 파일을 먼저 버퍼로 읽어, 경로 오류를 알아보기 쉬운 사유로 바꿈. fnFallback이 사유를 120자로 자르므로
            // 긴 경로는 끝부분(파일 이름 쪽)만 보여 준다. 상태 문구에는 줄표 대신 쌍점을 쓴다(버튼 요약 대화상자가 CP949).
            Shown = if Path = null then "" else if Text.Length(Path) > 70 then "..." & Text.End(Path, 67) else Path,
            Bin =
                if Path = null or Path = "" then error "설정 vs_path가 비어 있습니다"
                else
                    let Read = try Binary.Buffer(File.Contents(Path)) in
                    if Read[HasError] then error ("VALUESearch 파일을 읽지 못했습니다(vs_path): " & Shown) else Read[Value],
            Nav = Excel.Workbook(Bin, true, true),
            Sheets = Table.SelectRows(Nav, each [Kind] = "Sheet"),
            Picked = Table.SelectRows(Sheets, each [Item] = "Sheet2"),
            Data =
                if Table.RowCount(Picked) = 0 then
                    error ("VALUESearch 파일에 시트 Sheet2가 없습니다(있는 시트: " & Text.Combine(Sheets[Item], ", ") & ")")
                else Picked{0}[Data],
            Required = {"종목코드", "691300.NICS 산업분류", "691440.KSIC-세분류(11차)", "691230.주요상품", "691280.소속기업집단"},
            Missing = List.Difference(Required, Table.ColumnNames(Data)),
            Src =
                if List.IsEmpty(Missing) then Table.SelectColumns(Data, Required)
                else error ("VALUESearch 열 이름 불일치(없는 열: " & Text.Combine(Missing, ", ") & ")"),

            // 고정 코드표 → 코드별 이름 (대분류 / 업종)
            Codes = fnNicsCodes(),
            SecT = Table.SelectRows(Codes, each [단계] = "대분류"),
            GrpT = Table.SelectRows(Codes, each [단계] = "업종"),
            SecName = Record.FromList(SecT[이름], SecT[코드]),
            GrpName = Record.FromList(GrpT[이름], GrpT[코드]),

            // 종목코드: "A######" → 앞의 A 제거. 숫자만 6자리 미만이면 앞 0 채움. 6자리 [0-9A-Z]가 아니면 버림
            StockCode = (v as any) as nullable text =>
                let
                    t0 = Txt(v),
                    t1 = if t0 = null then null else Text.Upper(t0),
                    t2 = if t1 <> null and Text.Length(t1) = 7 and Text.StartsWith(t1, "A") then Text.Middle(t1, 1) else t1,
                    t3 = if t2 <> null and Text.Length(t2) < 6 and Text.Select(t2, Digits) = t2 then Text.PadStart(t2, 6, "0") else t2
                in
                    if t3 <> null and Text.Length(t3) = 6 and Text.Select(t3, CodeChars) = t3 then t3 else null,

            // NICS "N#####.이름" → [코드, 대분류, 업종, 세부, 출처]
            Nics = (v as any) as record =>
                let
                    t = Txt(v),
                    d5 = if t <> null and Text.Length(t) >= 7 and Text.StartsWith(t, "N") and Text.Middle(t, 6, 1) = "."
                         then Text.Middle(t, 1, 5) else null,
                    ok = d5 <> null and Text.Select(d5, Digits) = d5,
                    name = if ok then Text.Trim(Text.Middle(t, 7)) else null,
                    sec = if ok then Record.FieldOrDefault(SecName, "N" & Text.Start(d5, 2) & "000", null) else null,
                    grp = if ok then Record.FieldOrDefault(GrpName, "N" & Text.Start(d5, 3) & "00", null) else null,
                    level = if not ok then null else if Text.End(d5, 3) = "000" then "대분류" else if Text.End(d5, 2) = "00" then "업종" else "세부",
                    l1 = sec,
                    l2 = if sec = null then null else if level = "대분류" then sec & " (기타)" else grp,
                    l3 = if sec = null then null
                         else if level = "대분류" then sec & " (기타)"
                         else if level = "업종" then grp
                         else if name = "" then grp else name,
                    usable = l1 <> null and l2 <> null and l3 <> null
                in
                    [코드 = if ok then "N" & d5 else null,
                     대분류 = if usable then l1 else null, 업종 = if usable then l2 else null, 세부 = if usable then l3 else null,
                     출처 = if usable then "NICS" else null],

            Recs = List.Transform(Table.ToRecords(Src), (r) =>
                let
                    code = StockCode(Record.Field(r, "종목코드")),
                    n = Nics(Record.Field(r, "691300.NICS 산업분류"))
                in
                    if code = null then null
                    else {code, n[코드], n[대분류], n[업종], n[세부], Txt(Record.Field(r, "691440.KSIC-세분류(11차)")),
                          Txt(Record.Field(r, "691230.주요상품")), Txt(Record.Field(r, "691280.소속기업집단")), n[출처], "OK", Now}),
            T = #table(type table [종목코드 = text, NICS코드 = text, #"NICS 대분류" = text, #"NICS 업종" = text, #"NICS 세부" = text,
                    #"KSIC 세분류" = text, 주요상품 = text, 기업집단 = text, 분류출처 = text, 상태 = text, 조회시각 = datetime],
                    List.RemoveNulls(Recs)),
            Dedup = Table.Distinct(T, {"종목코드"}),
            Result =
                if Table.RowCount(Dedup) = 0 then error "VALUESearch 파일에 종목코드가 있는 행이 없습니다"
                else Table.Sort(Dedup, {{"종목코드", Order.Ascending}})
        in
            Result,
    // 실패 시 오류 대화상자 대신 직전 결과를 유지(fnFallback). 직전 표가 없으면 경고 1행
    Tried = try Table.Buffer(Main),
    Reason = if not Tried[HasError] then null
             else let m = Tried[Error][Message] in if m = null or m = "" then "사유 불명" else m,
    Kept = if Tried[HasError] then fnFallback("tblClass", Columns, Reason) else null,
    Warn = #table(Columns, {{null, null, null, null, null, null, null, null, null, "오류: " & Text.Start(Reason, 120), DateTime.LocalNow()}}),
    // fnFallback·경고 행은 형식이 없는 표라 그대로 적재하면 Excel이 조회시각을 일반 서식(일련번호)으로 바꾼다 → 열 형식을 다시 붙임(값은 그대로)
    Types = {{"종목코드", type text}, {"NICS코드", type text}, {"NICS 대분류", type text}, {"NICS 업종", type text},
             {"NICS 세부", type text}, {"KSIC 세분류", type text}, {"주요상품", type text}, {"기업집단", type text},
             {"분류출처", type text}, {"상태", type text}, {"조회시각", type datetime}},
    Final =
        if not Tried[HasError] then Tried[Value]
        else Table.TransformColumnTypes(if Table.RowCount(Kept) > 0 then Kept else Warn, Types)
in
    Final
