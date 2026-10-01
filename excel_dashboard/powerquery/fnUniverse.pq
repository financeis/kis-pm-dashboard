// fnUniverse : 종목DB(tblUniverse, 직전 새로고침 결과)에서 종목명·시장·섹터 조회용 표 반환 (없으면 빈 표)
() as table =>
let
    Empty = #table(type table [종목코드 = text, 종목명 = text, 시장 = text, 섹터 = text, 규모 = text, 유의사항 = text], {}),
    U = try Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "tblUniverse"]}[Content],
            {"종목코드", "종목명", "시장", "섹터", "규모", "유의사항"}) otherwise Empty,
    Typed = Table.TransformColumnTypes(U, {{"종목코드", type text}, {"종목명", type text}, {"시장", type text},
            {"섹터", type text}, {"규모", type text}, {"유의사항", type text}})
in
    Table.Distinct(Typed, {"종목코드"})
