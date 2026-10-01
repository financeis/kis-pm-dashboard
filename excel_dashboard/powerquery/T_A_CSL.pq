// T_A_CSL : 종목분석 5구역 — 신용·공매도·대차 일별 추이, 최근 60세션 (spec R17-5) → tblA_CSL (2026-10-01 T15)
//  - 원천(호출 4회, 2026-10-01 005930 실측으로 필드·단위 확인):
//      공매도 일별추이 FHPST04830000 daily-short-sale  기간 1회(오늘−110일 ~ 어제, 100건 이내) — output2
//          공매도수량 ssts_cntg_qty(주), 공매도거래대금억 ssts_tr_pbmn(원 → 억원; 수량 × 평균가와 일치),
//          공매도비중 ssts_tr_pbmn_rlim(거래대금 대비 % → 소수; 공매도 대금 ÷ 그날 거래대금과 일치), 종가 stck_clpr
//      종목별 일별 대차거래추이 HHPST074500C0 daily-loan-trans  조회구분 3(종목), 기간 1회 — output1
//          대차잔고주수 rmnd_stcn, 대차잔고금액억 rmnd_amt(백만원 → 억원; 잔고주수 × 종가와 일치), 대차증감주수 prdy_rmnd_vrss
//      신용잔고 일별추이 FHPST04760000 daily-credit-balance  30건/회, 입력 = 결제일자(그 이하) → 가장 오래된 결제일 − 1일로 당겨 2회
//          일자 = 매매일(deal_date; 결제일 stlm_date는 매매일 + 2영업일이라 최근 2세션은 아직 없음),
//          신용잔고율 whol_loan_rmnd_rate(% → 소수), 신용잔고주수 whol_loan_rmnd_stcn
//          (융자 금액 whol_loan_rmnd_amt는 단위를 실측으로 확정하지 못해 쓰지 않음 — 금액 ÷ 주수가 주가의 약 0.8배라 만원으로 보이나 문서 없음)
//  - 일자 축: 세 원천의 일자 합집합 중 오늘 전(어제까지)의 최근 60개 — 오늘 행은 넣지 않는다(장중 공매도 행은 0으로 오고,
//    신용·대차의 당일분은 장 마감 뒤 늦게 공표). 정렬: 일자 오름차순. 원천에 없는 날의 칸은 빈칸.
//    종가 = 공매도 → 대차 → 신용 응답 순으로 먼저 있는 값. 종가가 빈 행(없는 코드·상장 전)은 원천에서 버린다.
//  - 실패: 세 원천이 모두 실패하면 오류 → 직전 표 유지(fnARun). 일부만 실패하면 받은 값은 보여 주고 모든 행의 상태를
//    "오류: <원천> 조회 실패(사유)"로(단계 판정상 일부 오류). 일자가 하나도 없으면 "데이터 없음" 1행.
let
    Types = {{"종목코드", type text}, {"일자", type date}, {"종가", type number}, {"신용잔고율", type number}, {"신용잔고주수", type number},
             {"공매도수량", type number}, {"공매도거래대금억", type number}, {"공매도비중", type number},
             {"대차잔고주수", type number}, {"대차잔고금액억", type number}, {"대차증감주수", type number},
             {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Keep = 60,
            Today = Date.From(now),
            From = Date.AddDays(Today, -110),
            Yesterday = Date.AddDays(Today, -1),
            Q = "/uapi/domestic-stock/v1/quotations/",
            N = (r as nullable record, k as text) as nullable number => if r = null then null else fnNum(Record.FieldOrDefault(r, k, null)),
            Pos = (v as nullable number) as nullable number => if v = null or v <= 0 then null else v,
            ListOf = (j as record, k as text) as list =>
                let o = Record.FieldOrDefault(j, k, null) in if o is list then List.Select(o, each _ is record) else if o is record then {o} else {},
            Valid = (rows as list, dateKey as text, priceKey as text) as list =>
                List.Select(rows, each fnYmd(Record.FieldOrDefault(_, dateKey, null)) <> null and Pos(N(_, priceKey)) <> null),

            Short = try List.Buffer(Valid(ListOf(fnKisGet(Q & "daily-short-sale", "FHPST04830000", [FID_COND_MRKT_DIV_CODE = "J",
                        FID_INPUT_ISCD = code, FID_INPUT_DATE_1 = fnYmdText(From), FID_INPUT_DATE_2 = fnYmdText(Yesterday)]), "output2"),
                        "stck_bsop_date", "stck_clpr")),
            Loan = try List.Buffer(Valid(ListOf(fnKisGet(Q & "daily-loan-trans", "HHPST074500C0", [MRKT_DIV_CLS_CODE = "3",
                        MKSC_SHRN_ISCD = code, START_DATE = fnYmdText(From), END_DATE = fnYmdText(Yesterday), CTS = ""]), "output1"),
                        "bsop_date", "stck_prpr")),
            CreditPage = (stlm as date) as list =>
                List.Buffer(Valid(ListOf(fnKisGet(Q & "daily-credit-balance", "FHPST04760000", [FID_COND_MRKT_DIV_CODE = "J",
                    FID_COND_SCR_DIV_CODE = "20476", FID_INPUT_ISCD = code, FID_INPUT_DATE_1 = fnYmdText(stlm)]), "output"),
                    "deal_date", "stck_prpr")),
            Credit = try
                let
                    p1 = CreditPage(Today),
                    stl = List.RemoveNulls(List.Transform(p1, each fnYmd(Record.FieldOrDefault(_, "stlm_date", null)))),
                    p2 = if List.Count(p1) < 30 or List.IsEmpty(stl) then {} else CreditPage(Date.AddDays(List.Min(stl), -1))
                in
                    List.Buffer(p1 & p2),

            Sources = {[이름 = "공매도", r = Short], [이름 = "대차", r = Loan], [이름 = "신용", r = Credit]},
            Failed = List.Select(Sources, each [r][HasError]),
            Reason = (e as record) as text => let m = try Text.Trim(Text.From(e[Message])) otherwise "" in if m = "" then "조회 실패" else m,
            Rows0 = (s as record) as list => if s[HasError] then {} else s[Value],
            // 일자(yyyyMMdd) → 응답 행 (같은 일자는 앞 행)
            MapBy = (rows as list, k as text) as record =>
                let pairs = List.Distinct(List.Transform(rows, (r) => {fnYmdText(fnYmd(Record.FieldOrDefault(r, k, null))), r}), each _{0})
                in Record.FromList(List.Transform(pairs, each _{1}), List.Transform(pairs, each _{0})),
            SMap = MapBy(Rows0(Short), "stck_bsop_date"),
            LMap = MapBy(Rows0(Loan), "bsop_date"),
            CMap = MapBy(Rows0(Credit), "deal_date"),
            Dates = List.Distinct(List.Transform(Rows0(Short), each fnYmd([stck_bsop_date]))
                    & List.Transform(Rows0(Loan), each fnYmd([bsop_date])) & List.Transform(Rows0(Credit), each fnYmd([deal_date]))),
            Axis = List.Sort(List.FirstN(List.Sort(List.Select(Dates, each _ <> null and _ < Today), Order.Descending), Keep), Order.Ascending),
            Status = if List.IsEmpty(Failed) then "OK"
                     else "오류: " & Text.Combine(List.Transform(Failed, each [이름]), "·") & " 조회 실패(" & Text.Start(Reason(Failed{0}[r][Error]), 80) & ")",
            Rows = List.Transform(Axis, (d) =>
                let
                    k = fnYmdText(d),
                    s = Record.FieldOrDefault(SMap, k, null),
                    l = Record.FieldOrDefault(LMap, k, null),
                    c = Record.FieldOrDefault(CMap, k, null),
                    px = List.First(List.RemoveNulls({Pos(N(s, "stck_clpr")), Pos(N(l, "stck_prpr")), Pos(N(c, "stck_prpr"))}), null),
                    rate = N(c, "whol_loan_rmnd_rate"),
                    sAmt = N(s, "ssts_tr_pbmn"),
                    sPct = N(s, "ssts_tr_pbmn_rlim"),
                    lAmt = N(l, "rmnd_amt")
                in
                    {code, d, px, if rate = null then null else rate / 100, N(c, "whol_loan_rmnd_stcn"),
                     N(s, "ssts_cntg_qty"), if sAmt = null then null else sAmt / 100000000, if sPct = null then null else sPct / 100,
                     N(l, "rmnd_stcn"), if lAmt = null then null else lAmt / 100, N(l, "prdy_rmnd_vrss"), Status, now})
        in
            if List.Count(Failed) = List.Count(Sources) then error Error.Record("조회 실패", Reason(Failed{0}[r][Error]))
            else if List.IsEmpty(Axis) then fnAOneRow(Types, code, if Status = "OK" then "데이터 없음" else Status, now)
            else #table(List.Transform(Types, each _{0}), Rows)
in
    fnARun("tblA_CSL", Types, Main)
