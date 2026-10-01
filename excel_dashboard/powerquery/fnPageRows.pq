// fnPageRows : 대회종목 페이지 행 = 대회 종목 ∪ (보유·관심 중 대회 종목이 아닌 종목) (2026-10-01)
//  - 가격 저장소·수급·목표주가·신용/공매도·이벤트·대회종목 표 등 다른 쿼리는 조회 대상 종목을 **이 함수로만** 얻는다.
//  - 반환: 표 [종목코드, 구분, 유니버스]
//      구분 = 보유 / 관심 / 보유·관심 (fnCodes: 보유 = 매매일지 기준 보유수량 > 0, 관심 = tblWatch), 그 밖은 빈칸
//      유니버스 = 대회(대회편입 Y) / 유니버스 밖(대회편입 N인 보유·관심 종목)
//  - 대회편입은 fnClassify가 정한다(tblContest ± tblOverride — 수정표 규칙을 여기서 다시 구현하지 않음):
//    후보 = tblContest 종목코드 ∪ tblOverride 종목코드(추가 행 포함; 잘못된 행·종목DB에 없는 코드의 행·제외는 fnClassify가 처리)
//    ∪ fnCodes() 종목코드 → fnClassify → 대회편입 Y이거나 보유·관심인 종목만 남김. 잘못된 형식의 코드(6자리 [0-9A-Z] 아님)는
//    fnClassify가 버리고, 종목DB에 없는 코드의 수정표 행은 무시되므로 오타로 '추가'한 코드는 페이지 행이 되지 않는다.
//  - KRX업종: tblUniverse에 그 열이 있으면 함께 넘기고 없으면 null. 이 함수가 돌려주는 열은 KRX업종과 무관하지만
//    fnClassify 입력 계약(종목코드, KRX업종)을 그대로 지킨다.
//  - 순서(결정적): 대회 종목(tblContest 순서 → 수정표로 추가한 종목) → 유니버스 밖(보유·관심 순서). 페이지는 따로 정렬한다.
//  - 사용 예) Codes = List.Buffer(fnPageRows()[종목코드])
() as table =>
let
    Txt = (v as any) as nullable text =>
        if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,
    // fnCodes·fnClassify와 같은 정규화(대문자, 숫자만 6자리 미만이면 앞 0 채움) — KRX업종 조회 키용. 형식 검사는 fnClassify가 한다
    Norm = (v as any) as nullable text =>
        let t0 = Txt(try v otherwise null), t = if t0 = null then null else Text.Upper(t0)
        in if t = null then null
           else if Text.Length(t) < 6 and Text.Select(t, {"0".."9"}) = t then Text.PadStart(t, 6, "0") else t,
    CodesOf = (name as text) as list =>
        let t = try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null
        in if t = null or not List.Contains(Table.ColumnNames(t), "종목코드") then {} else Table.Column(t, "종목코드"),

    // 보유·관심 [종목코드, 구분] — fnCodes가 코드를 정규화하고 중복을 없앤다
    Watch = Table.Buffer(fnCodes()),
    Gubun = Record.FromList(Watch[구분], Watch[종목코드]),

    // KRX업종 (tblUniverse에 열이 있을 때만)
    Univ = try Excel.CurrentWorkbook(){[Name = "tblUniverse"]}[Content] otherwise null,
    HasKrx = Univ <> null and List.ContainsAll(Table.ColumnNames(Univ), {"종목코드", "KRX업종"}),
    KrxPairs =
        if not HasKrx then {}
        else List.Distinct(List.Select(List.Transform(Table.ToRecords(Table.SelectColumns(Univ, {"종목코드", "KRX업종"})),
                (r) => {Norm(r[종목코드]), Txt(try r[KRX업종] otherwise null)}), each _{0} <> null), each _{0}),
    Krx = Record.FromList(List.Transform(KrxPairs, each _{1}), List.Transform(KrxPairs, each _{0})),

    Candidates = CodesOf("tblContest") & CodesOf("tblOverride") & Watch[종목코드],
    Input = #table({"종목코드", "KRX업종"}, List.Transform(Candidates, (v) =>
        let n = Norm(v) in {v, if n = null then null else Record.FieldOrDefault(Krx, n, null)})),
    Classified = fnClassify(Input),

    Kept = List.RemoveNulls(List.Transform(Table.ToRecords(Classified), (r) =>
        let
            g = Record.FieldOrDefault(Gubun, r[종목코드], null),
            inContest = r[대회편입] = "Y"
        in
            if inContest or g <> null then {r[종목코드], g, if inContest then "대회" else "유니버스 밖"} else null)),
    Ordered = List.Select(Kept, each _{2} = "대회") & List.Select(Kept, each _{2} = "유니버스 밖"),
    Result = #table(type table [종목코드 = text, 구분 = text, 유니버스 = text], Ordered)
in
    Result
