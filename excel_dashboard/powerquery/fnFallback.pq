// fnFallback : 쿼리 실패 시 대체 결과 — 직전 새로 고침 결과(같은 표)를 그대로 유지, 없으면 빈 표
//  - 백그라운드 [모두 새로 고침] 중 오류 대화상자가 뜨지 않도록 모든 적재 쿼리의 마지막 단계에서 사용
//  - '상태' 열이 있는 표는 "이전 데이터(갱신 실패: 사유)"로 표시 → 대시보드 머리글 경고에 집계됨
//  - 사용 예) Tried = try Table.Buffer(Main), Final = if Tried[HasError] then fnFallback("tblFlow", {...}, Tried[Error][Message]) else Tried[Value]
(tableName as text, columns as list, optional reason as nullable text) as table =>
let
    Prev = try Excel.CurrentWorkbook(){[Name = tableName]}[Content] otherwise null,
    Usable = Prev <> null and List.ContainsAll(Table.ColumnNames(Prev), columns),
    Kept = if Usable then Table.SelectColumns(Prev, columns) else #table(columns, {}),
    Note = "이전 데이터(갱신 실패" & (if reason = null or reason = "" then "" else ": " & Text.Start(reason, 120)) & ")",
    Marked =
        if List.Contains(columns, "상태") and Table.RowCount(Kept) > 0 then
            Table.TransformColumns(Kept, {{"상태", each Note, type text}})
        else Kept
in
    Marked
