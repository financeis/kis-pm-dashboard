// T_Global : 해외지수·환율·금리 일별 시계열 (최근 약 90일)  [해외 종목/지수/환율기간별시세 v1_해외주식-012] FHKST03030100
//  - 대상 목록은 [설정] 시트 tblMacro 표에서 추가/삭제 가능 (시장코드 N=해외지수, X=환율, I=금리)
let
    Main =
        let
            Today = Date.From(DateTime.LocalNow()),
            From = Date.AddDays(Today, -130),
            Default = #table({"구분", "이름", "시장코드", "심볼", "순서"}, {
                {"미국", "S&P500", "N", "SPX", 1}, {"미국", "NASDAQ", "N", "COMP", 2}, {"미국", "다우", "N", ".DJI", 3},
                {"미국", "필라델피아반도체", "N", "SOX", 4}, {"미국", "VIX", "N", "VIX", 5},
                {"아시아", "니케이225", "N", "JP#NI225", 6}, {"아시아", "항셍", "N", "HK#HS", 7}, {"아시아", "상해종합", "N", "SHANG", 8},
                {"환율", "원/달러", "X", "FX@KRW", 9},
                {"금리", "미국10년", "I", "Y0202", 10}, {"금리", "국고3년", "I", "Y0101", 11}, {"금리", "국고10년", "I", "Y0106", 12},
                {"금리", "CD91일", "I", "Y0112", 13}, {"금리", "회사채AA-", "I", "Y0102", 14}}),
            Src = try Excel.CurrentWorkbook(){[Name = "tblMacro"]}[Content] otherwise Default,
            Items = Table.SelectRows(Src, each [심볼] <> null and Text.Trim(Text.From([심볼])) <> ""),
            One = (r as record) as table =>
                let
                    res = try Table.Buffer(fnMacroChart(Text.Trim(Text.From(r[시장코드])), Text.Trim(Text.From(r[심볼])), From, Today)),
                    T = if res[HasError] then #table(type table [일자 = date, 종가 = number], {}) else res[Value],
                    N = Table.RowCount(T),
                    Closes = List.Buffer(T[종가]),
                    Idx = Table.AddIndexColumn(T, "__i", 0, 1),
                    AddChg = Table.AddColumn(Idx, "등락률", each if [__i] = 0 then null else [종가] / Closes{[__i] - 1} - 1, type number),
                    AddDiff = Table.AddColumn(AddChg, "전일대비", each if [__i] = 0 then null else [종가] - Closes{[__i] - 1}, type number),
                    Meta = Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.AddColumn(Table.RemoveColumns(AddDiff, {"__i"}),
                        "구분", each r[구분], type text), "이름", each r[이름], type text), "순서", each r[순서]),
                        "상태", each "OK", type text)
                in
                    if res[HasError] then #table({"구분", "이름", "순서", "일자", "종가", "등락률", "전일대비", "상태"},
                        {{r[구분], r[이름], r[순서], null, null, null, null, "오류: " & res[Error][Message]}})
                    else Meta,
            All0 = Table.Combine(List.Transform(Table.ToRecords(Items), One)),
            Errs = List.Select(All0[상태], each Text.StartsWith(Text.From(_), "오류")),
            // 전부 실패하면 오류로 올려 직전 데이터를 유지(fnFallback)
            All = if Table.RowCount(All0) > 0 and List.Count(Errs) = Table.RowCount(All0) then error Text.Middle(Errs{0}, 4) else All0,
            Final = Table.SelectColumns(All, {"구분", "이름", "순서", "일자", "종가", "등락률", "전일대비", "상태"}),
            Typed = Table.TransformColumnTypes(Final, {{"구분", type text}, {"이름", type text}, {"순서", type number}, {"일자", type date},
                {"종가", type number}, {"등락률", type number}, {"전일대비", type number}, {"상태", type text}}),
            Sorted = Table.Sort(Typed, {{"순서", Order.Ascending}, {"일자", Order.Descending}}),
            // 이름별 최신 행 표시 (대시보드 조회용)
            Latest = Table.Group(Sorted, {"이름"}, {{"maxd", each List.Max([일자]), type date}}),
            J = Table.NestedJoin(Sorted, {"이름"}, Latest, {"이름"}, "l", JoinKind.LeftOuter),
            E = Table.ExpandTableColumn(J, "l", {"maxd"}),
            Flag = Table.AddColumn(E, "최신여부", each if [일자] <> null and [일자] = [maxd] then "Y" else "N", type text),
            Out = Table.Sort(Table.RemoveColumns(Flag, {"maxd"}), {{"순서", Order.Ascending}, {"일자", Order.Descending}})
        in
            Out,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblGlobal", {"구분", "이름", "순서", "일자", "종가", "등락률", "전일대비", "상태", "최신여부"}, Tried[Error][Message]) else Tried[Value]
in
    Final
