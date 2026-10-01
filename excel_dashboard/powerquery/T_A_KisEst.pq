// T_A_KisEst : 종목분석 7구역 — KIS 추정실적 연도별 표 (spec R17-7, R9 단위 규칙) → tblA_KisEst (2026-10-01 T15)
//  - 원천: 국내주식 종목추정실적 [국내주식-187] HHKST668300C0 /uapi/domestic-stock/v1/quotations/estimate-perform (SHT_CD, 1회)
//    = 한국투자증권 리서치 자체 추정(컨센서스 아님). 저장소 예시의 열 매핑은 틀림 — 아래는 2026-09-30·10-01 실측 해석.
//    + 현재가는 fnAQuote(호출 1회 — 추정이 있는 종목만).
//  - 연도: output4[i].dt(예 2023.12 … 2027.12E)가 data(i+1) 열의 라벨 → 행 = 연도(라벨 순서 그대로). 끝이 E면 추정여부 = "E".
//  - output2 행(위치 순서): 0 매출액, 1 매출증가율, 2 영업이익, 3 영업이익증가율, 4 순이익(지배), 5 순이익증가율
//    output3 행(위치 순서): 0 EBITDA, 1 EPS, 2 EPS증가율, 3 PER, 4 EV/EBITDA, 5 ROE, 6 부채비율, 7 이자보상배율
//    단위: 금액(매출·영업이익·순이익·EBITDA)은 억원 그대로, EPS와 모든 비율·배수는 ×10으로 옴 → ÷10.
//          비율(증가율·ROE·부채비율)은 ÷10 한 % 값을 다시 ÷100 해서 소수로. (005930 2023: 매출 2,589,355억, 매출증가율 −143 → −14.3%,
//          EPS 21310 → 2,131원, PER 368 → 36.8배 — 손익계산서 API·공시값과 일치)
//    행이 '처음 N/A가 있는 행'에서 잘려 온다(적자 연도가 있으면 output3이 EBITDA·EPS·EPS증가율 3행뿐, 은행·증권은 output3 없음)
//    → 없는 행의 열은 빈칸.
//  - PER = 현재가 ÷ EPS(EPS ≤ 0 이면 빈칸) — KIS의 PER은 추정일(실적 연도는 결산일) 주가 기준이라 KIS_PER로 따로 둠.
//  - 추정이 없는 종목은 rt_cd 0에 빈 배열(실측) → "추정 없음" 1행(이 종목군의 약 21%만 제공 — 정상).
let
    Types = {{"종목코드", type text}, {"연도", type text}, {"결산년월", type text}, {"추정여부", type text},
             {"매출억", type number}, {"매출증가율", type number}, {"영업이익억", type number}, {"영업이익증가율", type number},
             {"순이익억", type number}, {"순이익증가율", type number}, {"EBITDA억", type number},
             {"EPS", type number}, {"EPS증가율", type number}, {"PER", type number}, {"KIS_PER", type number},
             {"EV_EBITDA", type number}, {"ROE", type number}, {"부채비율", type number}, {"이자보상배율", type number},
             {"추정일", type date}, {"KIS의견", type text}, {"현재가", type number}, {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            J = fnKisGet("/uapi/domestic-stock/v1/quotations/estimate-perform", "HHKST668300C0", [SHT_CD = code]),
            ListOf = (k as text) as list =>
                let o = Record.FieldOrDefault(J, k, null) in List.Buffer(if o is list then List.Select(o, each _ is record) else {}),
            O1 = let o = Record.FieldOrDefault(J, "output1", null) in if o is record then o else [],
            O2 = ListOf("output2"),
            O3 = ListOf("output3"),
            Labels = List.Buffer(List.Select(List.Transform(ListOf("output4"), (r) =>
                        let v = Record.FieldOrDefault(r, "dt", null) in if v = null then "" else Text.Trim(Text.From(v))), each _ <> "")),
            Px = if List.IsEmpty(Labels) then null else fnAQuote(code, Date.From(now))[현재가],
            // rows{row}의 data(col+1) 값
            V = (rows as list, row as number, col as number) as nullable number =>
                if row < List.Count(rows) then fnNum(Record.FieldOrDefault(rows{row}, "data" & Text.From(col + 1), null)) else null,
            Div = (v as nullable number, d as number) as nullable number => if v = null then null else v / d,
            Txt = (k as text) as nullable text =>
                let v = Record.FieldOrDefault(O1, k, null), t = if v = null then "" else Text.Trim(Text.From(v)) in if t = "" then null else t,
            EstDate = fnYmd(Txt("estdate")),
            Opinion = Txt("rcmd_name"),
            Rows = List.Transform(List.Positions(Labels), (i) =>
                let
                    lbl = Labels{i},
                    eps = Div(V(O3, 1, i), 10)
                in
                    {code, lbl, Text.Start(lbl, 7), if Text.EndsWith(lbl, "E") then "E" else null,
                     V(O2, 0, i), Div(V(O2, 1, i), 1000), V(O2, 2, i), Div(V(O2, 3, i), 1000), V(O2, 4, i), Div(V(O2, 5, i), 1000),
                     V(O3, 0, i), eps, Div(V(O3, 2, i), 1000),
                     if eps = null or eps <= 0 or Px = null then null else Px / eps,
                     Div(V(O3, 3, i), 10), Div(V(O3, 4, i), 10), Div(V(O3, 5, i), 1000), Div(V(O3, 6, i), 1000), Div(V(O3, 7, i), 10),
                     EstDate, Opinion, Px, "OK", now})
        in
            if List.IsEmpty(Labels) then fnAOneRow(Types, code, "추정 없음", now)
            else #table(List.Transform(Types, each _{0}), Rows)
in
    fnARun("tblA_KisEst", Types, Main)
