// T_Positions : 종목별 손익 귀속(Attribution) 기초표 — 보유·청산 종목 전체
//  - 실현손익은 여기서, 평가손익·총손익·기여도는 시트 계산열(현재가 연동)에서 계산
let
    Main =
        let
            P = fnPositions(fnLedger(fnTrades())),
            U = fnUniverse(),
            J = Table.NestedJoin(P, {"종목코드"}, U, {"종목코드"}, "u", JoinKind.LeftOuter),
            E = Table.ExpandTableColumn(J, "u", {"종목명", "시장", "섹터"}),
            Named = Table.AddColumn(E, "__n", each if [종목명] = null then [종목코드] else [종목명], type text),
            Fixed = Table.RenameColumns(Table.RemoveColumns(Named, {"종목명"}), {{"__n", "종목명"}}),
            Cols = {"종목코드", "종목명", "시장", "섹터", "상태", "보유수량", "평균단가", "매입금액", "실현손익", "매수금액", "매도금액",
                    "수수료합", "세금합", "매매횟수", "최초매수일", "최근매매일"},
            Final = Table.SelectColumns(Fixed, Cols, MissingField.UseNull),
            Typed = Table.TransformColumnTypes(Final, {{"종목코드", type text}, {"종목명", type text}, {"시장", type text}, {"섹터", type text},
                {"상태", type text}, {"보유수량", type number}, {"평균단가", type number}, {"매입금액", type number}, {"실현손익", type number},
                {"매수금액", type number}, {"매도금액", type number}, {"수수료합", type number}, {"세금합", type number},
                {"매매횟수", Int64.Type}, {"최초매수일", type date}, {"최근매매일", type date}}),
            Sorted = Table.Sort(Typed, {{"상태", Order.Ascending}, {"매수금액", Order.Descending}})
        in
            Sorted,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblPositions", {"종목코드", "종목명", "시장", "섹터", "상태", "보유수량", "평균단가", "매입금액", "실현손익", "매수금액", "매도금액", "수수료합", "세금합", "매매횟수", "최초매수일", "최근매매일"}, Tried[Error][Message]) else Tried[Value]
in
    Final
