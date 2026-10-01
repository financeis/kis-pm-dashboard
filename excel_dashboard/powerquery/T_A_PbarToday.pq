// T_A_PbarToday : 종목분석 3구역 (a) — KIS 당일 매물대(가격대별 거래량·비중) → tblA_PbarToday (2026-10-01)
//  - 원천: 국내주식 매물대/거래비중 FHPST01130000 /uapi/domestic-stock/v1/quotations/pbar-tratio (당일 자료만, 1회 호출)
//    + 기준일(당일 표시)은 fnAQuote(최근 일봉 일자, 호출 1회).
//  - 응답 output2 = 당일 체결된 가격대, 거래량 순위 순으로 최대 100개(2026-10-01 실측: 005930 31개·247540 83개는 거래량 합 =
//    output1.acml_vol로 전부, 005935는 100개에서 잘려 합 99.89%).
//      가격 stck_prpr, 거래량 cntg_vol, 순위 data_rank
//      비중 = 거래량 ÷ 당일 누적 거래량(output1.acml_vol; 없으면 가격대 합) — KIS acml_vol_rlim(%)과 반올림 차이만 있음.
//        가격대가 100개를 넘어 잘린 날은 합이 100%보다 약간 작다(빠진 가격대는 거래량 하위).
//      현재가구간 = output1.stck_prpr와 같은 가격대 "Y"(목록에 없으면 가장 가까운 가격대), 그 밖 빈칸
//  - 정렬: 가격 오름차순(가로 막대 차트에서 아래 = 낮은 가격). 가격대가 없으면 "데이터 없음" 1행.
let
    Types = {{"종목코드", type text}, {"기준일", type date}, {"가격", type number}, {"거래량", type number}, {"비중", type number},
             {"순위", type number}, {"현재가구간", type text}, {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Qt = fnAQuote(code, Date.From(now)),
            Base = Qt[기준일],
            J = fnKisGet("/uapi/domestic-stock/v1/quotations/pbar-tratio", "FHPST01130000",
                    [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code, FID_COND_SCR_DIV_CODE = "20113", FID_INPUT_HOUR_1 = ""]),
            O1 = let o = Record.FieldOrDefault(J, "output1", null) in if o is record then o else [],
            O2 = let o = Record.FieldOrDefault(J, "output2", null) in if o is list then o else {},
            N = (r as record, k as text) as nullable number => fnNum(Record.FieldOrDefault(r, k, null)),
            Levels = List.Buffer(List.Select(
                List.Transform(List.Select(O2, each _ is record), (r) => [가격 = N(r, "stck_prpr"), 거래량 = N(r, "cntg_vol"), 순위 = N(r, "data_rank")]),
                each [가격] <> null and [가격] > 0 and [거래량] <> null)),
            SumVol = List.Sum(List.Transform(Levels, each [거래량])),
            Acml = N(O1, "acml_vol"),
            Denom = if Acml <> null and Acml > 0 then Acml else SumVol,
            Px = let v = N(O1, "stck_prpr") in if v = null or v <= 0 then null else v,
            Prices = List.Transform(Levels, each [가격]),
            Diffs = if Px = null then {} else List.Transform(Prices, each Number.Abs(_ - Px)),
            Marked = if List.IsEmpty(Diffs) then null else Prices{List.PositionOf(Diffs, List.Min(Diffs))},
            Rows = List.Transform(Levels, (l) =>
                {code, Base, l[가격], l[거래량], if Denom = null or Denom = 0 then null else l[거래량] / Denom, l[순위],
                 if Marked <> null and l[가격] = Marked then "Y" else null, "OK", now}),
            T = #table(List.Transform(Types, each _{0}), Rows)
        in
            if List.IsEmpty(Levels) or SumVol = null or SumVol = 0 then
                Table.TransformColumns(fnAOneRow(Types, code, "데이터 없음", now), {{"기준일", each Base, type date}})
            else Table.Sort(T, {{"가격", Order.Ascending}})
in
    fnARun("tblA_PbarToday", Types, Main)
