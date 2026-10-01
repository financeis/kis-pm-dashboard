// T_A_Targets : 종목분석 6구역 (1) — 증권사별 최근 의견·목표가 목록, 최근 6개월 (spec R17-6, R8) → tblA_Targets (2026-10-01 T15)
//  - 원천: fnAOpinions(종목투자의견 FHKST663300C0, 100건씩 페이지) — 기간 = 오늘 − (기간 + 6개월) ~ 오늘(직전 목표가를 찾으려고 6개월 더).
//    기간 = 설정 target_window_months(기본 6, 1~24). 목록 대상 = (오늘 − 기간, 오늘] 안에 의견을 낸 증권사.
//  - 행 = 증권사 1곳(가장 최근 의견일 내림차순, 같으면 증권사 이름순):
//      일자·투자의견·직전투자의견 = 기간 안의 가장 최근 의견(직전투자의견 = 응답의 rgbf_invt_opnn)
//      목표가·목표가일자 = 기간 안의 마지막 0이 아닌 목표가와 그 날짜(R8 컨센서스가 쓰는 값; 기간 안에 목표가가 없으면 빈칸)
//      직전목표가 = 그 목표가보다 앞선 날짜의 마지막 0이 아닌 목표가(받은 자료 안 — 기간 밖일 수 있음), 목표가변화율 = 목표가 ÷ 직전목표가 − 1
//    평균 목표가·괴리율은 페이지가 이 표(목표가 열)와 tblA_Info[현재가]로 계산하거나 tblA_TargetTrend의 '현재' 행을 쓴다.
//  - 기간 안에 의견이 없으면 "데이터 없음" 1행(이 종목군의 약 42%는 목표가가 없음 — 정상).
let
    Types = {{"종목코드", type text}, {"증권사", type text}, {"일자", type date}, {"투자의견", type text}, {"직전투자의견", type text},
             {"목표가", type number}, {"목표가일자", type date}, {"직전목표가", type number}, {"목표가변화율", type number},
             {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Today = Date.From(now),
            Win = fnASetting("target_window_months", 6, 1, 24),
            WinStart = Date.AddMonths(Today, -Win),
            Op = fnAOpinions(code, Date.AddMonths(Today, -(Win + 6)), Today),
            InWin = Table.Buffer(Table.SelectRows(Op, each [일자] > WinStart and [일자] <= Today and [증권사] <> null)),
            Brokers = List.Distinct(InWin[증권사]),
            Rows = List.Transform(Brokers, (b) =>
                let
                    mine = Table.SelectRows(InWin, each [증권사] = b),
                    latest = mine{0},
                    withTarget = Table.SelectRows(mine, each [목표가] <> null),
                    tgt = if Table.IsEmpty(withTarget) then null else withTarget{0},
                    older = if tgt = null then null
                            else Table.SelectRows(Op, each [증권사] = b and [목표가] <> null and [일자] < tgt[일자]),
                    prev = if older = null or Table.IsEmpty(older) then null else older{0}[목표가],
                    cur = if tgt = null then null else tgt[목표가]
                in
                    {code, b, latest[일자], latest[투자의견], latest[직전투자의견], cur, if tgt = null then null else tgt[일자],
                     prev, if cur = null or prev = null then null else cur / prev - 1, "OK", now}),
            T = #table(List.Transform(Types, each _{0}), Rows)
        in
            if Table.IsEmpty(InWin) then fnAOneRow(Types, code, "데이터 없음", now)
            else Table.Sort(T, {{"일자", Order.Descending}, {"증권사", Order.Ascending}})
in
    fnARun("tblA_Targets", Types, Main)
