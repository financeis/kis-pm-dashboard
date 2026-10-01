// T_Holdings : 현재 보유종목 (매매일지 기준)  ← [포트폴리오] 시트, 표 오른쪽 계산열은 시세판·지표와 연결된 수식
//  - 손절가/목표가를 매매일지에 적지 않았으면 [설정]의 기본 손절·목표 수익률을 평균단가에 적용
let
    Main =
        let
            S = fnSettings(),
            Today = Date.From(DateTime.LocalNow()),
            StopPct = try Number.From(S[stop_loss]) otherwise -0.08,
            TakePct = try Number.From(S[take_profit]) otherwise 0.2,
            P = fnPositions(fnLedger(fnTrades())),
            H = Table.SelectRows(P, each [보유수량] > 0),
            U = fnUniverse(),
            J = Table.NestedJoin(H, {"종목코드"}, U, {"종목코드"}, "u", JoinKind.LeftOuter),
            E = Table.ExpandTableColumn(J, "u", {"종목명", "시장", "섹터"}),
            Named = Table.AddColumn(E, "__n", each if [종목명] = null then [종목코드] else [종목명], type text),
            Fixed = Table.RenameColumns(Table.RemoveColumns(Named, {"종목명"}), {{"__n", "종목명"}}),
            Add1 = Table.AddColumn(Fixed, "보유일수", each if [진입일] = null then null else Duration.Days(Today - [진입일]), type number),
            Add2 = Table.AddColumn(Add1, "손절가적용", each if [손절가] <> null then [손절가] else Number.Round([평균단가] * (1 + StopPct), 0), type number),
            Add3 = Table.AddColumn(Add2, "목표가적용", each if [목표가] <> null then [목표가] else Number.Round([평균단가] * (1 + TakePct), 0), type number),
            // 시트에서 현재가·평가손익 등 계산열이 평균단가 뒤, 손절여유·신호 등이 목표가 뒤에 끼워짐
            Cols = {"종목코드", "종목명", "섹터", "보유수량", "평균단가", "손절가적용", "목표가적용", "매입금액", "실현손익", "진입일",
                    "보유일수", "시장", "전략", "매매근거", "확인"},
            Final = Table.SelectColumns(Add3, Cols, MissingField.UseNull),
            Renamed = Table.RenameColumns(Final, {{"손절가적용", "손절가"}, {"목표가적용", "목표가"}}),
            Typed = Table.TransformColumnTypes(Renamed, {{"종목코드", type text}, {"종목명", type text}, {"시장", type text}, {"섹터", type text},
                {"보유수량", type number}, {"평균단가", type number}, {"매입금액", type number}, {"실현손익", type number}, {"진입일", type date},
                {"보유일수", type number}, {"손절가", type number}, {"목표가", type number}, {"전략", type text}, {"매매근거", type text}, {"확인", type text}}),
            Sorted = Table.Sort(Typed, {{"매입금액", Order.Descending}})
        in
            Sorted,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblHoldings", {"종목코드", "종목명", "섹터", "보유수량", "평균단가", "손절가", "목표가", "매입금액", "실현손익", "진입일", "보유일수", "시장", "전략", "매매근거", "확인"}, Tried[Error][Message]) else Tried[Value]
in
    Final
