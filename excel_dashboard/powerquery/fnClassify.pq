// fnClassify : 종목 분류·대회편입 판정의 단일 함수 (spec R2·R3·R4, 2026-10-01 T2)
//  - 종목DB(tblUniverse)·대회종목·순위·시세판·테마 집계·종목분석이 모두 이 함수로 분류를 붙여 같은 답을 쓴다.
//    수정표 규칙을 다른 쿼리에서 다시 구현하지 않는다(대상 종목 목록은 fnPageRows).
//  - 입력: 표 [종목코드, KRX업종(선택)]. KRX업종 열이 없거나 값이 비면 KRX 대체 없이 진행한다(MissingField.UseNull로 읽음 —
//    tblUniverse에 KRX업종 열이 생기기 전에도 동작). 종목코드 열이 없으면 오류.
//    종목코드는 정규화(앞뒤 공백 제거·대문자·숫자만 6자리 미만이면 앞 0 채움) 뒤 6자리 [0-9A-Z]인 것만 남긴다.
//    잘못된·빈 코드 행은 버리고, 같은 코드가 여러 번 오면 처음 행(그 행의 KRX업종)만 쓴다. 출력 순서 = 입력의 처음 등장 순서.
//  - 출력: 종목코드, 대회편입(Y/N), 대회비고, NICS 대분류, NICS 업종, NICS 세부, 대테마, 세부테마, 분류출처, 섹터 (모두 텍스트, 없으면 빈칸)
//  - 적용 순서(spec R4): 기본(명단 tblContest, NICS 분류 tblClass, 기본 테마표 tblThemeBase) → 그 위에 수정표 tblOverride가 항상 우선.
//    · 대회편입: tblContest에 있으면 Y, 없으면 N. 수정표 대회편입 = 추가 → Y, 제외 → N (모든 화면에서 같은 판정).
//    · 대회비고: tblContest 비고(외국기업·5일 미만 등)에, 수정표가 소속을 바꾼 경우 "수정표 추가"/"수정표 제외"를 "; "로 덧붙임
//      (이미 대회 종목인데 추가, 대회 밖인데 제외처럼 소속이 그대로면 덧붙이지 않음).
//    · NICS: tblClass에서 분류출처가 NICS인 행의 3단계 → 분류출처 NICS. 없으면(파일에 없음·NICS 빈 값·코드표 밖 코드)
//      입력 KRX업종을 세 단계 모두에 넣고 분류출처 KRX. KRX업종도 없으면 세 단계·분류출처 모두 빈칸.
//      tblClass가 '이전 데이터(갱신 실패)' 상태여도 그 분류를 쓴다(spec: 파일을 못 읽으면 직전 분류 유지).
//    · 대테마·세부테마: tblThemeBase(대회 종목이 아닌 종목은 보통 없음 → 빈칸, 사용자가 수정표로 지정 가능).
//    · 수정표의 NICS 대분류/업종/세부·대테마/세부테마는 비어 있지 않은 칸만 그 값으로 바꾼다. 분류출처는 NICS/KRX(또는 빈칸)
//      그대로 둔다(새 값을 만들지 않음).
//    · 섹터: 설정 sector_basis(대분류/업종/세부/대테마 — 비거나 다른 값이면 기본 세부)의 값. 그 값이 없으면
//      세부 → 업종 → 대분류 → 입력 KRX업종 순으로 대체(대테마 기준도 같음). 모두 없으면 빈칸.
//  - 수정표(tblOverride): 공란 = 수정 없음. 종목코드가 6자리 [0-9A-Z]가 아니거나(정규화 후) 대회편입이 공란·추가·제외가 아니거나
//    종목코드가 종목DB(tblUniverse)에 없는 행은 통째로 무시한다(오타 코드가 이름 없는 대회 종목 행이 되지 않게). 같은 종목코드가 여러
//    행이면 칸마다 아래쪽 행의 비어 있지 않은 값이 우선.
//    · 종목DB 확인은 tblUniverse가 있고 종목코드 칸이 하나라도 채워져 있을 때만 한다. 표가 없거나 비었으면(하네스 시험,
//      T_Universe를 처음 적재하기 전의 빌드) 확인하지 않고 형식만 본다.
//    · 셀 값 정리는 [설정]의 경고 수식(빌더 F_OVERRIDE_WARN — '수정표 n행 무시됨')과 같다: 줄바꿈 없는 공백(U+00A0)·전각 공백(U+3000)
//      → 일반 공백, 제어 문자(코드 0~31) 제거, 앞뒤 공백 제거, 종목코드는 대문자·숫자만 6자리 미만이면 앞 0 채움. 경고 수식도 종목DB가
//      비었으면 종목DB 확인을 하지 않으므로, 무시되는 행과 경고 건수가 같다.
//  - 읽는 표: tblContest·tblOverride·tblThemeBase·tblClass·tblSettings, 수정표 행이 있으면 tblUniverse(종목코드 열만)
//    (표가 없으면 빈 표로 취급). 빈 행(종목코드 없음)은 무시.
//  - 사용 예) fnClassify(Table.SelectColumns(U, {"종목코드", "KRX업종"}))   // 종목DB
//            fnClassify(#table({"종목코드"}, {{"005930"}}))                  // 종목 1개(종목분석)
(input as table) as table =>
let
    // 셀 값 → 앞뒤 공백을 뗀 텍스트(빈 값·오류 값은 null)
    Txt = (v as any) as nullable text =>
        if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,
    F = (r as record, k as text) as nullable text => Txt(try Record.FieldOrDefault(r, k, null) otherwise null),
    Digits = {"0".."9"},
    CodeChars = {"0".."9", "A".."Z"},
    // 종목코드 정규화 — 셀 서식이 숫자여서 앞 0이 빠진 코드(5930 → 005930)·소문자를 받아 주고, 6자리 [0-9A-Z]가 아니면 null
    Code = (v as any) as nullable text =>
        let
            t0 = Txt(v),
            t1 = if t0 = null then null else Text.Upper(t0),
            t2 = if t1 <> null and Text.Length(t1) < 6 and Text.Select(t1, Digits) = t1 then Text.PadStart(t1, 6, "0") else t1
        in
            if t2 <> null and Text.Length(t2) = 6 and Text.Select(t2, CodeChars) = t2 then t2 else null,
    Rows = (name as text) as list =>
        let t = try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null
        in if t = null then {} else List.Buffer(Table.ToRecords(t)),
    // {코드, 값} 목록 → 코드별 값 레코드 (같은 코드는 처음 것; Record.FromList는 이름이 겹치면 오류이므로 먼저 중복 제거)
    MapOf = (pairs as list) as record =>
        let d = List.Distinct(List.Select(pairs, each _{0} <> null), each _{0})
        in Record.FromList(List.Transform(d, each _{1}), List.Transform(d, each _{0})),

    // 기본 1: 대회 명단(값 = [비고]; 레코드로 감싸 '명단에 있고 비고 없음'과 '명단에 없음'을 구분)
    Contest = MapOf(List.Transform(Rows("tblContest"), (r) => {Code(F(r, "종목코드")), [비고 = F(r, "비고")]})),
    // 기본 2: NICS 분류(분류출처 NICS 행만; 종목코드가 빈 경고 행은 Code가 null이라 빠짐)
    Nics = MapOf(List.Transform(Rows("tblClass"), (r) =>
        if F(r, "분류출처") = "NICS" then {Code(F(r, "종목코드")), [대분류 = F(r, "NICS 대분류"), 업종 = F(r, "NICS 업종"), 세부 = F(r, "NICS 세부")]}
        else {null, null})),
    // 기본 3: 기본 테마표
    Theme = MapOf(List.Transform(Rows("tblThemeBase"), (r) => {Code(F(r, "종목코드")), [대테마 = F(r, "대테마"), 세부테마 = F(r, "세부테마")]})),

    // 수정표 셀 값 → 텍스트(빈 값은 null). 경고 수식의 정리 규칙(머리말)과 같게: U+00A0·U+3000 → 공백, 코드 0~31 문자 제거,
    // 앞뒤 공백(" ") 제거
    OvTxt = (v as any) as nullable text =>
        let
            s0 = if v = null then null else try Text.From(v) otherwise null,
            s1 = if s0 = null then null
                 else Text.Replace(Text.Replace(s0, Character.FromNumber(160), " "), Character.FromNumber(12288), " "),
            t = if s1 = null then null else Text.Trim(Text.Remove(s1, List.Transform({0..31}, Character.FromNumber)), " ")
        in
            if t = null or t = "" then null else t,
    OF = (r as record, k as text) as nullable text => OvTxt(try Record.FieldOrDefault(r, k, null) otherwise null),
    // 종목DB 종목코드 집합(정규화한 코드 → true). 표가 없거나 종목코드 칸이 모두 비었으면 null = 종목DB 확인 안 함(머리말 —
    // 경고 수식의 '종목DB에 종목코드가 하나라도 있을 때'와 같은 기준). 수정표에 형식이 맞는 행이 있을 때만 평가된다
    // (아래 'or'는 앞이 참이면 뒤를 평가하지 않음)
    UnivSet =
        let
            t = try Excel.CurrentWorkbook(){[Name = "tblUniverse"]}[Content] otherwise null,
            filled = if t = null or not List.Contains(Table.ColumnNames(t), "종목코드") then {}
                     else List.Select(List.Transform(Table.Column(t, "종목코드"), (v) => try Txt(v) otherwise null), each _ <> null),
            codes = List.Distinct(List.RemoveNulls(List.Transform(filled, Code)))
        in
            if List.IsEmpty(filled) then null else Record.FromList(List.Repeat({true}, List.Count(codes)), codes),
    // 수정표: 잘못된 행·종목DB에 없는 코드의 행은 통째로 무시, 같은 코드는 칸마다 아래쪽 행의 비어 있지 않은 값 우선
    OvFields = {"대회편입", "대테마", "세부테마", "NICS 대분류", "NICS 업종", "NICS 세부"},
    OvRows = List.RemoveNulls(List.Transform(Rows("tblOverride"), (r) =>
        let
            c = Code(OF(r, "종목코드")),
            m = OF(r, "대회편입")
        in
            if c = null or not (m = null or m = "추가" or m = "제외") or (UnivSet <> null and not Record.HasFields(UnivSet, c)) then null
            else {c, List.Transform(OvFields, (k) => OF(r, k))})),
    OvCodes = List.Distinct(List.Transform(OvRows, each _{0})),
    OvMerged = List.Transform(OvCodes, (c) =>
        let
            vals = List.Transform(List.Select(OvRows, each _{0} = c), each _{1}),
            merged = List.Accumulate(vals, List.Repeat({null}, List.Count(OvFields)),
                (acc, v) => List.Transform(List.Positions(OvFields), (i) => if v{i} <> null then v{i} else acc{i}))
        in
            Record.FromList(merged, OvFields)),
    Override = Record.FromList(OvMerged, OvCodes),

    // 섹터 기준 (설정이 없거나 목록 밖 값이면 기본 세부)
    S = try fnSettings() otherwise [],
    BasisIn = Txt(try Record.FieldOrDefault(S, "sector_basis", null) otherwise null),
    Basis = if List.Contains({"대분류", "업종", "세부", "대테마"}, BasisIn) then BasisIn else "세부",

    // 입력: 종목코드(필수) + KRX업종(선택)
    In0 =
        if List.Contains(Table.ColumnNames(input), "종목코드") then input
        else error "fnClassify: 입력 표에 종목코드 열이 없습니다",
    InRows = List.Buffer(Table.ToRecords(Table.SelectColumns(In0, {"종목코드", "KRX업종"}, MissingField.UseNull))),
    InPairs = List.Distinct(List.Select(List.Transform(InRows, (r) => {Code(F(r, "종목코드")), F(r, "KRX업종")}), each _{0} <> null), each _{0}),

    Out = List.Transform(InPairs, (p) =>
        let
            c = p{0},
            krx = p{1},
            con = Record.FieldOrDefault(Contest, c, null),
            ov = Record.FieldOrDefault(Override, c, null),
            // 수정표 칸이 비어 있지 않으면 그 값, 아니면 기본값
            Ov = (k as text, base as nullable text) as nullable text =>
                let v = if ov = null then null else Record.Field(ov, k) in if v <> null then v else base,
            inContest = con <> null,
            mem = Ov("대회편입", null),
            member = if mem = "추가" then "Y" else if mem = "제외" then "N" else if inContest then "Y" else "N",
            notes = List.RemoveNulls({
                if inContest then con[비고] else null,
                if mem = "추가" and not inContest then "수정표 추가" else null,
                if mem = "제외" and inContest then "수정표 제외" else null}),
            cls = Record.FieldOrDefault(Nics, c, null),
            source = if cls <> null then "NICS" else if krx <> null then "KRX" else null,
            l1 = Ov("NICS 대분류", if cls <> null then cls[대분류] else krx),
            l2 = Ov("NICS 업종", if cls <> null then cls[업종] else krx),
            l3 = Ov("NICS 세부", if cls <> null then cls[세부] else krx),
            th = Record.FieldOrDefault(Theme, c, null),
            t1 = Ov("대테마", if th = null then null else th[대테마]),
            t2 = Ov("세부테마", if th = null then null else th[세부테마]),
            chosen = if Basis = "대분류" then l1 else if Basis = "업종" then l2 else if Basis = "대테마" then t1 else l3,
            sector = List.First(List.RemoveNulls({chosen, l3, l2, l1, krx}), null)
        in
            {c, member, if List.IsEmpty(notes) then null else Text.Combine(notes, "; "), l1, l2, l3, t1, t2, source, sector}),
    Result = #table(type table [종목코드 = text, 대회편입 = text, 대회비고 = text, #"NICS 대분류" = text, #"NICS 업종" = text,
        #"NICS 세부" = text, 대테마 = text, 세부테마 = text, 분류출처 = text, 섹터 = text], Out)
in
    Result
