// fnPositions : 원장(fnLedger) → 종목별 현재 포지션 요약 (보유 + 청산 종목 모두)
(ledger as table) as table =>
let
    Grouped = Table.Group(ledger, {"종목코드"}, {
        {"보유수량", each List.Last([보유수량]), type number},
        {"평균단가", each List.Last([평균단가]), type number},
        {"실현손익", each List.Sum(List.RemoveNulls([실현손익])), type number},
        {"매수금액", each List.Sum(Table.SelectRows(_, each [구분] = "매수")[금액]), type number},
        {"매도금액", each List.Sum(Table.SelectRows(_, each [구분] = "매도")[금액]), type number},
        {"수수료합", each List.Sum([수수료]), type number},
        {"세금합", each List.Sum([세금]), type number},
        {"매매횟수", each Table.RowCount(_), Int64.Type},
        {"최초매수일", each List.Min(Table.SelectRows(_, each [구분] = "매수")[일자]), type date},
        {"최근매매일", each List.Max([일자]), type date},
        // 현재 보유분의 진입일 = 마지막으로 보유수량이 0이 된 이후 첫 매수일
        {"진입일", each
            let
                L = Table.ToRecords(_),
                LastZero = List.Last(List.Select(List.Positions(L), (i) => L{i}[보유수량] = 0), -1),
                After = List.Skip(L, LastZero + 1),
                Buys = List.Select(After, (x) => x[구분] = "매수")
            in
                if List.Count(Buys) = 0 then null else Buys{0}[일자], type date},
        {"목표가", each List.Last(List.RemoveNulls([목표가]), null), type number},
        {"손절가", each List.Last(List.RemoveNulls([손절가]), null), type number},
        {"전략", each List.Last(List.RemoveNulls(List.Transform([전략], (x) => if x = "" then null else x)), null)},
        {"매매근거", each List.Last(List.RemoveNulls(List.Transform([매매근거], (x) => if x = "" then null else x)), null)},
        {"확인", each Text.Combine(List.RemoveNulls([확인]), " / "), type text}
    }),
    WithCost = Table.AddColumn(Grouped, "매입금액", each [보유수량] * [평균단가], type number),
    WithState = Table.AddColumn(WithCost, "상태", each if [보유수량] > 0 then "보유" else "청산", type text)
in
    WithState
