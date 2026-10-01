// fnCodes : 시세 조회 대상 종목 = 보유종목(매매일지 기준) ∪ 관심종목(tblWatch)
//  - 반환: 표 [종목코드, 구분(보유/관심/보유·관심)]
() as table =>
let
    Pos = try fnPositions(fnLedger(fnTrades())) otherwise #table({"종목코드", "보유수량"}, {}),
    Held = List.Transform(Table.SelectRows(Pos, each [보유수량] > 0)[종목코드], Text.From),
    W = try Excel.CurrentWorkbook(){[Name = "tblWatch"]}[Content] otherwise #table({"종목코드"}, {}),
    Norm = (v as any) as nullable text =>
        if v = null then null
        else let t = Text.Upper(Text.Trim(Text.From(v))) in
            if t = "" then null
            else if Text.Length(t) < 6 and (try Number.FromText(t) otherwise null) <> null then Text.PadStart(t, 6, "0") else t,
    Watch = List.RemoveNulls(List.Transform(W[종목코드], Norm)),
    All = List.Distinct(Held & Watch),
    T = #table(type table [종목코드 = text, 구분 = text], List.Transform(All, each {_,
            if List.Contains(Held, _) and List.Contains(Watch, _) then "보유·관심" else if List.Contains(Held, _) then "보유" else "관심"}))
in
    T
