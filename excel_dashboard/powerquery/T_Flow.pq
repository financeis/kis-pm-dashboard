// T_Flow : 시장별 투자자 순매수(일별, 억원)  [시장별 투자자매매동향(일별) 국내주식-075] FHPTJ04040000
//  - 1회 호출로 약 300영업일 제공 → 최근 60영업일만 사용 (단위: 백만원 → 억원)
let
    Main =
        let
            Path = "/uapi/domestic-stock/v1/quotations/inquire-investor-daily-by-market",
            Today = fnYmdText(Date.From(DateTime.LocalNow())),
            Keep = 60,
            Get = (market as text, mktCode as text, idx as text) as table =>
                let
                    J = fnKisGet(Path, "FHPTJ04040000", [FID_COND_MRKT_DIV_CODE = "U", FID_INPUT_ISCD = idx, FID_INPUT_DATE_1 = Today,
                            FID_INPUT_ISCD_1 = mktCode, FID_INPUT_DATE_2 = Today, FID_INPUT_ISCD_2 = idx]),
                    O = Record.FieldOrDefault(J, "output", {}),
                    L = List.FirstN(List.Select(if O is list then O else {}, each Text.From(Record.FieldOrDefault(_, "stck_bsop_date", "")) <> ""), Keep),
                    Eok = (r as record, k as text) as nullable number => let v = fnNum(Record.FieldOrDefault(r, k, null)) in if v = null then null else v / 100,
                    Rows = List.Transform(L, (r) => [
                        시장 = market,
                        일자 = fnYmd(r[stck_bsop_date]),
                        지수 = fnNum(Record.FieldOrDefault(r, "bstp_nmix_prpr", null)),
                        등락률 = let c = fnNum(Record.FieldOrDefault(r, "bstp_nmix_prdy_ctrt", null)) in if c = null then null else c / 100,
                        외국인 = Eok(r, "frgn_ntby_tr_pbmn"),
                        기관계 = Eok(r, "orgn_ntby_tr_pbmn"),
                        개인 = Eok(r, "prsn_ntby_tr_pbmn"),
                        금융투자 = Eok(r, "scrt_ntby_tr_pbmn"),
                        투신 = Eok(r, "ivtr_ntby_tr_pbmn"),
                        사모 = Eok(r, "pe_fund_ntby_tr_pbmn"),
                        은행 = Eok(r, "bank_ntby_tr_pbmn"),
                        보험 = Eok(r, "insu_ntby_tr_pbmn"),
                        연기금 = Eok(r, "fund_ntby_tr_pbmn"),
                        기타법인 = Eok(r, "etc_corp_ntby_tr_pbmn")])
                in
                    Table.FromRecords(Rows, {"시장", "일자", "지수", "등락률", "외국인", "기관계", "개인", "금융투자", "투신", "사모", "은행", "보험", "연기금", "기타법인"}, MissingField.UseNull),
            All = Table.Combine({Get("KOSPI", "KSP", "0001"), Get("KOSDAQ", "KSQ", "1001")}),
            Typed = Table.TransformColumnTypes(All, {{"시장", type text}, {"일자", type date}, {"지수", type number}, {"등락률", type number},
                {"외국인", type number}, {"기관계", type number}, {"개인", type number}, {"금융투자", type number}, {"투신", type number},
                {"사모", type number}, {"은행", type number}, {"보험", type number}, {"연기금", type number}, {"기타법인", type number}}),
            Sorted = Table.Sort(Typed, {{"시장", Order.Ascending}, {"일자", Order.Descending}}),
            Keyed = Table.AddColumn(Sorted, "키", each [시장] & "|" & Date.ToText([일자], [Format = "yyyy-MM-dd", Culture = "en-US"]), type text)
        in
            Keyed,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblFlow", {"시장", "일자", "지수", "등락률", "외국인", "기관계", "개인", "금융투자", "투신", "사모", "은행", "보험", "연기금", "기타법인", "키"}, Tried[Error][Message]) else Tried[Value]
in
    Final
