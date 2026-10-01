// fnLedger : 매매내역 → 종목별 이동평균 원가 원장 (거래 1건당 1행)
//  - 평균단가는 매수 수수료를 포함한 원가 기준 → 실현손익 + 평가손익 = 순자산 - 초기자금 (정확히 일치)
//  - 매도: 실현손익 = (금액 - 수수료 - 세금) - 매도수량 × 평균단가
//  - 보유수량보다 많이 매도하면 보유수량까지만 반영하고 '확인필요' 표시
(trades as table) as table =>
let
    Rows = List.Buffer(Table.ToRecords(trades)),
    Codes = List.Distinct(List.Transform(Rows, each [종목코드])),
    PerCode = List.Transform(Codes, (c) =>
        let
            R = List.Select(Rows, each [종목코드] = c),
            Acc = List.Accumulate(R, [qty = 0, avg = 0, cum = 0, open = null, out = {}], (s, t) =>
                let
                    isBuy = t[구분] = "매수",
                    q = t[수량],
                    sellQ = if isBuy then 0 else List.Min({q, s[qty]}),
                    newQty = if isBuy then s[qty] + q else s[qty] - sellQ,
                    newAvg = if isBuy then (s[qty] * s[avg] + t[금액] + t[수수료]) / newQty
                             else if newQty = 0 then 0 else s[avg],
                    proceeds = if isBuy then 0 else (sellQ * t[단가]) - t[수수료] - t[세금],
                    realized = if isBuy then null else proceeds - sellQ * s[avg],
                    cost = if isBuy then null else sellQ * s[avg],
                    open = if isBuy and s[qty] = 0 then t[일자] else s[open],
                    held = if isBuy or s[open] = null then null else Duration.Days(t[일자] - s[open]),
                    row = t & [
                        보유수량 = newQty,
                        평균단가 = newAvg,
                        실현손익 = realized,
                        실현수익률 = if realized = null or cost = null or cost = 0 then null else realized / cost,
                        종목누적실현 = s[cum] + (if realized = null then 0 else realized),
                        보유일수 = held,
                        결과 = if realized = null then null else if realized > 0 then "승" else if realized < 0 then "패" else "무",
                        확인 = if not isBuy and q > s[qty] then "확인필요: 보유수량 초과 매도" else null]
                in
                    [qty = newQty, avg = newAvg, cum = row[종목누적실현], open = if newQty = 0 then null else open, out = s[out] & {row}])
        in
            Acc[out]),
    All = List.Combine(PerCode),
    Cols = Table.ColumnNames(trades) & {"보유수량", "평균단가", "실현손익", "실현수익률", "종목누적실현", "보유일수", "결과", "확인"},
    T = if List.Count(All) = 0 then #table(Cols, {}) else Table.FromRecords(All, Cols, MissingField.UseNull),
    Typed = Table.TransformColumnTypes(T, {{"보유수량", type number}, {"평균단가", type number}, {"실현손익", type number},
        {"실현수익률", type number}, {"종목누적실현", type number}, {"보유일수", type number}, {"결과", type text}, {"확인", type text}}),
    Sorted = Table.Sort(Typed, {{"일자", Order.Ascending}, {"순번", Order.Ascending}})
in
    Sorted
