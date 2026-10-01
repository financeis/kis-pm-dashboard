// fnSettings : [설정] 시트의 tblSettings(키/값) 표를 레코드로 반환
// 사용 예) fnSettings()[init_capital]
() as record =>
let
    T = Excel.CurrentWorkbook(){[Name = "tblSettings"]}[Content],
    Rows = Table.SelectRows(T, each [키] <> null and Text.Trim(Text.From([키])) <> ""),
    Keys = List.Transform(Rows[키], each Text.Trim(Text.From(_))),
    Rec = Record.FromList(Rows[값], Keys)
in
    Rec
