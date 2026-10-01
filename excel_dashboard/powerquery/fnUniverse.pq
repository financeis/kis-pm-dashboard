// fnUniverse : 종목DB(tblUniverse, 직전 새로고침 결과)에서 종목명·시장·섹터·대테마 조회용 표 반환 (없으면 빈 표)
//  - 섹터 = 설정 sector_basis 기준 값, 대테마 = 테마 분류(수정표 반영) — 둘 다 T_Universe가 fnClassify로 채운 값 (T22)
//  - tblUniverse에 아직 없는 열(예: 대테마 열이 생기기 전의 표)은 빈 값으로 채운다 — 한 열이 없다고 종목명·섹터 조회까지
//    통째로 빈 표가 되지 않게
() as table =>
let
    Cols = {"종목코드", "종목명", "시장", "섹터", "규모", "유의사항", "대테마"},
    Empty = #table(type table [종목코드 = text, 종목명 = text, 시장 = text, 섹터 = text, 규모 = text, 유의사항 = text, 대테마 = text], {}),
    U = try Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "tblUniverse"]}[Content], Cols, MissingField.UseNull) otherwise Empty,
    Typed = Table.TransformColumnTypes(U, List.Transform(Cols, each {_, type text}))
in
    Table.Distinct(Typed, {"종목코드"})
