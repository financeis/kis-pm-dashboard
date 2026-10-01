// T_A_Investor : 종목분석 1구역 — 가격·투자자 일별 최근 120 완료 세션 → tblA_Investor (2026-10-01)
//  - 원천: 종목별 투자자매매동향(일별) FHPTJ04160001 /uapi/domestic-stock/v1/quotations/investor-trade-by-stock-daily
//    1회 = 입력일(FID_INPUT_DATE_1) 이하 최근 30세션(최신순). 가장 오래된 일자 − 1일로 입력일을 당겨 4회 → 120세션.
//    30행 미만이 오면(상장 직후) 거기서 멈춤 — 상장 전 날짜는 종가가 빈 행으로 오므로(없는 코드 999999도 같은 모양, 실측) 종가가 빈 행은 버린다.
//  - 완료 세션만: 오늘 날짜 입력은 00:00~15:40에 "OPSQ2001 TIME LIMIT"으로 거부되므로(2026-10-01 실측) 15:40 전에는 어제부터 받는다
//    → 장중(진행 중 세션)의 오늘 행은 이 표에 없다(당일 흐름은 2·3a·4구역). 15:40 이후에는 오늘 행이 들어온다
//      (그 뒤에도 시간외 거래가 18:00까지 거래대금에 더해져 오늘 행 값이 조금 바뀔 수 있음 — 15:42→15:47 실측 0.05억 차이).
//  - 열·단위(실측 검증):
//      종가 stck_clpr, 등락률 prdy_ctrt(% → 소수), 거래대금억 acml_tr_pbmn(원 → 억원; 일봉 API와 같은 값)
//      투자자 순매수(백만원 → 억원 ÷100): 외국인 frgn_ntby_tr_pbmn, 기관계 orgn_ntby_tr_pbmn, 연기금 fund_ntby_tr_pbmn(기금),
//        투신 ivtr_ntby_tr_pbmn, 사모 pe_fund_ntby_tr_pbmn, 금융투자 scrt_ntby_tr_pbmn(증권), 보험 insu_ntby_tr_pbmn,
//        은행 bank_ntby_tr_pbmn, 기타법인 etc_corp_ntby_tr_pbmn, 개인 prsn_ntby_tr_pbmn
//        (기관계 = 금융투자+투신+사모+은행+보험+종금+기금+기타단체, 외국인+기관계+개인+기타법인 ≈ 0 — 실측 확인)
//      누적(억원): 외국인·기관계·개인 일별 값을 120세션 중 가장 오래된 날부터 더한 값(빈 값은 0으로 건너뜀, 처음부터 모두 비면 빈칸)
//  - 정렬: 일자 오름차순(차트용, 마지막 행 = 최근 완료 세션). 행이 없으면 "데이터 없음" 1행.
let
    Types = {{"종목코드", type text}, {"일자", type date}, {"종가", type number}, {"등락률", type number}, {"거래대금억", type number},
             {"외국인억", type number}, {"기관계억", type number}, {"연기금억", type number}, {"투신억", type number}, {"사모억", type number},
             {"금융투자억", type number}, {"보험억", type number}, {"은행억", type number}, {"기타법인억", type number}, {"개인억", type number},
             {"외국인누적억", type number}, {"기관계누적억", type number}, {"개인누적억", type number},
             {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Keep = 120,
            PageSize = 30,
            Path = "/uapi/domestic-stock/v1/quotations/investor-trade-by-stock-daily",
            Today = Date.From(now),
            End0 = if Time.From(now) >= #time(15, 40, 0) then Today else Date.AddDays(Today, -1),
            N = (r as record, k as text) as nullable number => fnNum(Record.FieldOrDefault(r, k, null)),
            Page = (endDate as date) as list =>
                let
                    J = fnKisGet(Path, "FHPTJ04160001", [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code,
                            FID_INPUT_DATE_1 = fnYmdText(endDate), FID_ORG_ADJ_PRC = "", FID_ETC_CLS_CODE = ""]),
                    O = Record.FieldOrDefault(J, "output2", {}),
                    L = if O is list then O else {}
                in
                    List.Buffer(List.Select(L, each _ is record
                        and fnYmd(Record.FieldOrDefault(_, "stck_bsop_date", null)) <> null
                        and (let c = N(_, "stck_clpr") in c <> null and c > 0))),
            MinDate = (rows as list) as date => List.Min(List.Transform(rows, each fnYmd([stck_bsop_date]))),
            Pages = List.Generate(
                () => [rows = Page(End0), n = 1, total = 0],
                each List.Count([rows]) > 0,
                each
                    let got = [total] + List.Count([rows]) in
                        if got >= Keep or [n] >= 6 or List.Count([rows]) < PageSize then [rows = {}, n = [n] + 1, total = got]
                        else [rows = Page(Date.AddDays(MinDate([rows]), -1)), n = [n] + 1, total = got],
                each [rows]),
            Recs = List.Combine(Pages),
            Eok = (r as record, k as text) as nullable number => let v = N(r, k) in if v = null then null else v / 100,
            Daily = List.Transform(Recs, (r) => [
                일자 = fnYmd(r[stck_bsop_date]),
                종가 = N(r, "stck_clpr"),
                등락률 = let v = N(r, "prdy_ctrt") in if v = null then null else v / 100,
                거래대금억 = let v = N(r, "acml_tr_pbmn") in if v = null then null else v / 100000000,
                외국인억 = Eok(r, "frgn_ntby_tr_pbmn"), 기관계억 = Eok(r, "orgn_ntby_tr_pbmn"), 연기금억 = Eok(r, "fund_ntby_tr_pbmn"),
                투신억 = Eok(r, "ivtr_ntby_tr_pbmn"), 사모억 = Eok(r, "pe_fund_ntby_tr_pbmn"), 금융투자억 = Eok(r, "scrt_ntby_tr_pbmn"),
                보험억 = Eok(r, "insu_ntby_tr_pbmn"), 은행억 = Eok(r, "bank_ntby_tr_pbmn"), 기타법인억 = Eok(r, "etc_corp_ntby_tr_pbmn"),
                개인억 = Eok(r, "prsn_ntby_tr_pbmn")]),
            DailyCols = {"일자", "종가", "등락률", "거래대금억", "외국인억", "기관계억", "연기금억", "투신억", "사모억", "금융투자억",
                         "보험억", "은행억", "기타법인억", "개인억"},
            T0 = Table.FromRecords(Daily, DailyCols, MissingField.UseNull),
            Clean = Table.SelectRows(T0, each [일자] <> null and [일자] <= End0),
            Dedup = Table.Distinct(Clean, {"일자"}),
            Latest = Table.FirstN(Table.Sort(Dedup, {{"일자", Order.Descending}}), Keep),
            Asc = List.Buffer(Table.ToRecords(Table.Sort(Latest, {{"일자", Order.Ascending}}))),
            // 누적: 가장 오래된 날부터 더함(빈 값은 건너뜀, 처음부터 모두 비면 빈칸)
            Cum = (vals as list) as list =>
                List.Generate(
                    () => [i = 0, s = vals{0}?],
                    each [i] < List.Count(vals),
                    each [i = [i] + 1, s = let v = vals{[i] + 1}? in if v = null then [s] else if [s] = null then v else [s] + v],
                    each [s]),
            CumF = List.Buffer(Cum(List.Transform(Asc, each [외국인억]))),
            CumO = List.Buffer(Cum(List.Transform(Asc, each [기관계억]))),
            CumP = List.Buffer(Cum(List.Transform(Asc, each [개인억]))),
            Rows = List.Transform(List.Positions(Asc), (i) =>
                let r = Asc{i} in
                    {code, r[일자], r[종가], r[등락률], r[거래대금억], r[외국인억], r[기관계억], r[연기금억], r[투신억], r[사모억],
                     r[금융투자억], r[보험억], r[은행억], r[기타법인억], r[개인억], CumF{i}, CumO{i}, CumP{i}, "OK", now})
        in
            if List.IsEmpty(Asc) then fnAOneRow(Types, code, "데이터 없음", now)
            else #table(List.Transform(Types, each _{0}), Rows)
in
    fnARun("tblA_Investor", Types, Main)
