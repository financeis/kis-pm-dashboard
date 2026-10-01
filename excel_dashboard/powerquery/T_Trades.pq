// T_Trades : 매매 원장 (거래별 수수료·세금·현금흐름·보유수량·평균단가·실현손익)  ← [매매분석] 시트
let
    Main =
        let
            L = fnLedger(fnTrades()),
            U = fnUniverse(),
            J = Table.NestedJoin(L, {"종목코드"}, U, {"종목코드"}, "u", JoinKind.LeftOuter),
            E = Table.ExpandTableColumn(J, "u", {"종목명", "섹터"}),
            Named = Table.AddColumn(E, "__n", each if [종목명] = null then [종목코드] else [종목명], type text),
            Fixed = Table.RenameColumns(Table.RemoveColumns(Named, {"종목명"}), {{"__n", "종목명"}}),
            Cols = {"순번", "일자", "종목코드", "종목명", "섹터", "구분", "수량", "단가", "금액", "수수료", "세금", "현금흐름", "보유수량",
                    "평균단가", "실현손익", "실현수익률", "종목누적실현", "보유일수", "결과", "전략", "매매근거", "목표가", "손절가", "메모", "확인"},
            Final = Table.SelectColumns(Fixed, Cols, MissingField.UseNull),
            Typed = Table.TransformColumnTypes(Final, {{"순번", Int64.Type}, {"일자", type date}, {"종목코드", type text}, {"종목명", type text},
                {"섹터", type text}, {"구분", type text}, {"수량", type number}, {"단가", type number}, {"금액", type number},
                {"수수료", type number}, {"세금", type number}, {"현금흐름", type number}, {"보유수량", type number}, {"평균단가", type number},
                {"실현손익", type number}, {"실현수익률", type number}, {"종목누적실현", type number}, {"보유일수", type number},
                {"결과", type text}, {"전략", type text}, {"매매근거", type text}, {"목표가", type number}, {"손절가", type number},
                {"메모", type text}, {"확인", type text}}),
            Sorted = Table.Sort(Typed, {{"일자", Order.Descending}, {"순번", Order.Descending}})
        in
            Sorted,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblTradeLog", {"순번", "일자", "종목코드", "종목명", "섹터", "구분", "수량", "단가", "금액", "수수료", "세금", "현금흐름", "보유수량", "평균단가", "실현손익", "실현수익률", "종목누적실현", "보유일수", "결과", "전략", "매매근거", "목표가", "손절가", "메모", "확인"}, Tried[Error][Message]) else Tried[Value]
in
    Final
