// T_A_Estimate : 종목분석 4구역 — 외인·기관 추정가집계(당일) → tblA_Estimate (2026-10-01)
//  - 원천: 종목별 외인기관 추정가집계 HHPTJ04160200 /uapi/domestic-stock/v1/quotations/investor-trend-estimate (당일 자료만, 1회 호출)
//    + 기준일(당일 표시)·현재가는 fnAQuote(최근 일봉, 호출 1회).
//  - 응답 output2 = 이미 입력된 시점만, 최신 시점부터(2026-10-01 11:42 실측: 입력구분 3·2·1의 3행). 값은 장 시작부터의 누적 추정
//    순매수 수량(주, 0을 채운 부호 있는 문자열 "-00000000000871000").
//    입력구분(bsop_hour_gb) → 입력 시점: 1 = 09:30, 2 = 10:00, 3 = 11:20, 4 = 13:20, 5 = 14:30 (11:42에 1~3만 있음 확인)
//  - 행: 항상 5행(입력 시점 순). 아직 입력되지 않은 시점은 수량·금액 빈칸, 상태 "데이터 없음". 입력된 시점은 0도 값 그대로.
//      외국인수량 frgn_fake_ntby_qty, 기관수량 orgn_fake_ntby_qty, 합산수량 sum_fake_ntby_qty(= 외국인 + 기관)
//      외국인≈억·기관≈억·합산≈억 = 수량 × 현재가 ÷ 1억 (추정 수량에 지금 가격을 곱한 근사 금액)
let
    Types = {{"종목코드", type text}, {"기준일", type date}, {"순서", Int64.Type}, {"입력시점", type text},
             {"외국인수량", type number}, {"기관수량", type number}, {"합산수량", type number},
             {"외국인≈억", type number}, {"기관≈억", type number}, {"합산≈억", type number}, {"현재가", type number},
             {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Qt = fnAQuote(code, Date.From(now)),
            Base = Qt[기준일],
            Px = Qt[현재가],
            J = fnKisGet("/uapi/domestic-stock/v1/quotations/investor-trend-estimate", "HHPTJ04160200", [MKSC_SHRN_ISCD = code]),
            O = Record.FieldOrDefault(J, "output2", {}),
            L = List.Buffer(List.Select(if O is list then O else if O is record then {O} else {}, each _ is record)),
            Slots = {{"1", "09:30"}, {"2", "10:00"}, {"3", "11:20"}, {"4", "13:20"}, {"5", "14:30"}},
            Gb = (r as record) as text => let v = Record.FieldOrDefault(r, "bsop_hour_gb", null) in if v = null then "" else Text.Trim(Text.From(v)),
            Find = (g as text) as nullable record => let m = List.Select(L, each Gb(_) = g) in if List.IsEmpty(m) then null else m{0},
            N = (r as nullable record, k as text) as nullable number => if r = null then null else fnNum(Record.FieldOrDefault(r, k, null)),
            Eok = (q as nullable number) as nullable number => if q = null or Px = null then null else q * Px / 100000000,
            Rows = List.Transform(List.Positions(Slots), (i) =>
                let
                    r = Find(Slots{i}{0}),
                    f = N(r, "frgn_fake_ntby_qty"),
                    o = N(r, "orgn_fake_ntby_qty"),
                    s = N(r, "sum_fake_ntby_qty")
                in
                    {code, Base, i + 1, Slots{i}{1}, f, o, s, Eok(f), Eok(o), Eok(s), Px, if r = null then "데이터 없음" else "OK", now})
        in
            // 응답·가격을 먼저 평가해(행 수 확인) 호출 오류가 셀이 아닌 표 단위 오류가 되게 함
            if List.Count(L) >= 0 and (Px = null or Px > 0) then #table(List.Transform(Types, each _{0}), Rows) else error "응답 확인 실패"
in
    fnARun("tblA_Estimate", Types, Main)
