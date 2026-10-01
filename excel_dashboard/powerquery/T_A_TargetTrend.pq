// T_A_TargetTrend : 종목분석 6구역 (2) — 월말 목표주가 컨센서스 추이, 최근 12개월 + 현재 (spec R17-6, R8) → tblA_TargetTrend (2026-10-01 T15)
//  - 원천: fnAOpinions(종목투자의견 FHKST663300C0, 100건씩 페이지) — 기간 = 가장 이른 월말 − 기간 ~ 오늘.
//  - 기준일 D: 최근 12개 월말(오늘이 속한 달 이전 12개 달의 말일, 오래된 것부터, 구분 "월말") + 오늘(구분 "현재").
//  - R8 컨센서스(D) = 증권사별로 (D − 기간, D] 안의 마지막 0이 아닌 목표가를 모아 평균. 기간 = 설정 target_window_months(기본 6).
//      목표가평균(원), 증권사수(그 평균에 들어간 증권사 수), 최고목표가·최저목표가
//    증권사가 없는 기준일은 목표가평균 등 빈칸·증권사수 0·상태 "데이터 없음". 받은 자료에 목표가가 하나도 없으면 "데이터 없음" 1행.
//  - 괴리율(= 목표가평균 ÷ 현재가 − 1)은 페이지가 tblA_Info[현재가]로 계산한다.
let
    Types = {{"종목코드", type text}, {"기준일", type date}, {"구분", type text}, {"목표가평균", type number}, {"증권사수", Int64.Type},
             {"최고목표가", type number}, {"최저목표가", type number}, {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Today = Date.From(now),
            Win = fnASetting("target_window_months", 6, 1, 24),
            Ends = List.Transform({12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1}, (k) => Date.EndOfMonth(Date.AddMonths(Today, -k))),
            Points = List.Transform(Ends, each {_, "월말"}) & {{Today, "현재"}},
            Op = fnAOpinions(code, Date.AddMonths(Ends{0}, -Win), Today),
            // 일자 내림차순(같은 날은 응답 순서) → 증권사별 첫 행 = 그 기간의 마지막 목표가
            Tg = Table.Buffer(Table.SelectRows(Op, each [목표가] <> null and [증권사] <> null)),
            Cons = (d as date) as list =>
                let
                    w = Table.SelectRows(Tg, each [일자] > Date.AddMonths(d, -Win) and [일자] <= d),
                    // 묶음 안의 행 순서는 원래 순서 그대로 → 첫 값 = 증권사별 가장 최근 목표가
                    last = Table.Group(w, {"증권사"}, {{"목표가", each List.First([목표가]), type number}}),
                    v = List.Buffer(last[목표가])
                in
                    {if List.IsEmpty(v) then null else List.Average(v), List.Count(v), List.Max(v), List.Min(v)},
            Rows = List.Transform(Points, (p) =>
                let c = Cons(p{0}) in
                    {code, p{0}, p{1}, c{0}, c{1}, c{2}, c{3}, if c{1} > 0 then "OK" else "데이터 없음", now})
        in
            if Table.IsEmpty(Tg) then fnAOneRow(Types, code, "데이터 없음", now)
            else #table(List.Transform(Types, each _{0}), Rows)
in
    fnARun("tblA_TargetTrend", Types, Main)
