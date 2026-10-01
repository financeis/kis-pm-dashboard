// fnAOneRow : 종목분석(tblA_*) 표의 '상태만 있는 1행' (2026-10-01 T15)
//  - 데이터 없음·추정 없음(정상 빈 결과)이나 오류를 빈 표 대신 1행으로 보여 주기 위해 쓴다(빈 표로는 결과·실패가 보이지 않음).
//  - 종목코드·상태·조회시각 열만 채우고 나머지 열은 빈칸. types = fnARun과 같은 {{"열", type …}, …} 목록.
(types as list, code as nullable text, status as text, now as datetime) as table =>
let
    Cols = List.Transform(types, each _{0}),
    Row = List.Transform(Cols, (c) => if c = "종목코드" then code else if c = "상태" then status else if c = "조회시각" then now else null)
in
    Table.TransformColumnTypes(#table(Cols, {Row}), types)
