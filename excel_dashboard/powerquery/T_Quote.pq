// T_Quote : 보유·관심 종목 시세판 (현재가·밸류에이션·52주·수급)
//  - 주식현재가 시세 [v1_국내주식-008] FHKST01010100 : 가격/PER/PBR/52주/외국인소진율/업종/시장경고
//  - 주식현재가 투자자 [v1_국내주식-012] FHKST01010900 : 최근 30영업일 외국인·기관·개인 순매수 → 5일/20일 합계(억원)
let
    Main =
        let
            S = fnSettings(),
            WithFlow = (try Text.Upper(Text.From(S[stock_flow])) otherwise "Y") = "Y",
            Now = DateTime.LocalNow(),
            Codes = fnCodes(),
            U = fnUniverse(),
            Q = "/uapi/domestic-stock/v1/quotations/",
            Get = (code as text) as record =>
                let
                    o = fnKisGet(Q & "inquire-price", "FHKST01010100", [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code])[output],
                    n = (k as text) => fnNum(Record.FieldOrDefault(o, k, null)),
                    p = (k as text) => let v = n(k) in if v = null then null else v / 100,
                    Inv = if not WithFlow then {}
                          else let r = try fnKisGet(Q & "inquire-investor", "FHKST01010900", [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code])[output] otherwise {}
                               in if r is list then List.Select(r, each Text.From(Record.FieldOrDefault(_, "stck_bsop_date", "")) <> "") else {},
                    SumN = (k as text, days as number) as nullable number =>
                        let v = List.RemoveNulls(List.Transform(List.FirstN(Inv, days), each fnNum(Record.FieldOrDefault(_, k, null))))
                        in if List.Count(v) = 0 then null else List.Sum(v) / 100,
                    warnCode = Text.From(Record.FieldOrDefault(o, "mrkt_warn_cls_code", "")),
                    flags = List.RemoveNulls({
                        if Record.FieldOrDefault(o, "temp_stop_yn", "N") = "Y" then "거래정지" else null,
                        if warnCode = "01" then "투자주의" else if warnCode = "02" then "투자경고" else if warnCode = "03" then "투자위험" else null,
                        if Record.FieldOrDefault(o, "invt_caful_yn", "N") = "Y" then "투자유의" else null,
                        if Record.FieldOrDefault(o, "short_over_yn", "N") = "Y" then "단기과열" else null,
                        if Record.FieldOrDefault(o, "mang_issu_cls_code", "N") = "Y" then "관리종목" else null,
                        if Text.From(Record.FieldOrDefault(o, "vi_cls_code", "N")) <> "N" and Text.From(Record.FieldOrDefault(o, "vi_cls_code", "N")) <> "" then "VI발동" else null})
                in
                    // 레코드 필드는 지연 평가되므로 응답 o를 먼저 평가해 try가 API 오류를 잡도록 함
                    if not (o is record) then error "빈 응답" else
                    [종목코드 = code,
                     현재가 = n("stck_prpr"), 전일대비 = n("prdy_vrss"), 등락률 = p("prdy_ctrt"),
                     시가 = n("stck_oprc"), 고가 = n("stck_hgpr"), 저가 = n("stck_lwpr"), 전일종가 = n("stck_sdpr"),
                     거래량 = n("acml_vol"), 거래대금억 = let v = n("acml_tr_pbmn") in if v = null then null else v / 100000000,
                     전일대비거래량 = p("prdy_vrss_vol_rate"),
                     시가총액억 = n("hts_avls"), PER = n("per"), PBR = n("pbr"), EPS = n("eps"), BPS = n("bps"),
                     외국인소진율 = p("hts_frgn_ehrt"),
                     고52주 = n("w52_hgpr"), 고52주일 = fnYmd(Record.FieldOrDefault(o, "w52_hgpr_date", null)),
                     저52주 = n("w52_lwpr"), 저52주일 = fnYmd(Record.FieldOrDefault(o, "w52_lwpr_date", null)),
                     외국인당일주 = n("frgn_ntby_qty"), 프로그램당일주 = n("pgtr_ntby_qty"),
                     외국인5일억 = SumN("frgn_ntby_tr_pbmn", 5), 기관5일억 = SumN("orgn_ntby_tr_pbmn", 5), 개인5일억 = SumN("prsn_ntby_tr_pbmn", 5),
                     외국인20일억 = SumN("frgn_ntby_tr_pbmn", 20), 기관20일억 = SumN("orgn_ntby_tr_pbmn", 20),
                     상한가 = n("stck_mxpr"), 하한가 = n("stck_llam"),
                     유의 = Text.Combine(flags, ","), 상태 = "OK", 조회시각 = Now],
            Rows = List.Transform(Table.ToRecords(Codes), (c) =>
                let r = try Get(c[종목코드]) in
                    (if r[HasError] then [종목코드 = c[종목코드], 상태 = "오류: " & r[Error][Message], 조회시각 = Now] else r[Value]) & [구분 = c[구분]]),
            Cols = {"구분", "종목코드", "현재가", "전일대비", "등락률", "시가", "고가", "저가", "전일종가", "거래량", "거래대금억",
                    "전일대비거래량", "시가총액억", "PER", "PBR", "EPS", "BPS", "외국인소진율", "고52주", "고52주일", "저52주", "저52주일",
                    "외국인당일주", "프로그램당일주", "외국인5일억", "기관5일억", "개인5일억", "외국인20일억", "기관20일억", "상한가", "하한가",
                    "유의", "상태", "조회시각"},
            Failed = List.Select(Rows, each Text.StartsWith(Text.From(Record.FieldOrDefault(_, "상태", "")), "오류")),
            // 전부 실패하면 오류로 올려 직전 데이터를 유지(fnFallback), 일부만 실패하면 그 행만 오류 표시
            Checked = if List.Count(Rows) > 0 and List.Count(Failed) = List.Count(Rows) then error Text.Middle(Failed{0}[상태], 4) else Rows,
            T0 = Table.FromRecords(Checked, Cols, MissingField.UseNull),
            T = Table.AddIndexColumn(T0, "__ord", 0, 1, Int64.Type),
            J = Table.NestedJoin(T, {"종목코드"}, U, {"종목코드"}, "u", JoinKind.LeftOuter),
            E = Table.ExpandTableColumn(J, "u", {"종목명", "시장", "섹터", "유의사항"}, {"종목명", "시장", "섹터", "DB유의"}),
            Named = Table.AddColumn(E, "__n", each if [종목명] = null then [종목코드] else [종목명], type text),
            Fixed = Table.RenameColumns(Table.RemoveColumns(Named, {"종목명"}), {{"__n", "종목명"}}),
            W52 = Table.AddColumn(Table.AddColumn(Fixed,
                    "고52주대비", each if [고52주] = null or [고52주] = 0 or [현재가] = null then null else [현재가] / [고52주] - 1, type number),
                    "저52주대비", each if [저52주] = null or [저52주] = 0 or [현재가] = null then null else [현재가] / [저52주] - 1, type number),
            // 핵심 열을 앞쪽에 배치 (시트에서 신호·추세 등 계산열이 등락률 뒤에 끼워짐)
            Order = {"구분", "종목코드", "종목명", "시장", "섹터", "현재가", "전일대비", "등락률", "거래대금억", "시가총액억", "PER", "PBR",
                     "외국인소진율", "고52주대비", "저52주대비", "외국인5일억", "기관5일억", "개인5일억", "외국인20일억", "기관20일억",
                     "시가", "고가", "저가", "전일종가", "거래량", "전일대비거래량", "EPS", "BPS", "고52주", "고52주일", "저52주", "저52주일",
                     "외국인당일주", "프로그램당일주", "상한가", "하한가", "유의", "DB유의", "상태", "조회시각"},
            // PER·PBR 등 0은 '데이터 없음'으로 처리
            NoZero = Table.TransformColumns(W52, List.Transform({"PER", "PBR", "EPS", "BPS"}, (c) => {c, each if _ = 0 then null else _})),
            Final = Table.SelectColumns(Table.Sort(NoZero, {{"__ord", Order.Ascending}}), Order),
            Typed = Table.TransformColumnTypes(Final, {
                {"구분", type text}, {"종목코드", type text}, {"종목명", type text}, {"시장", type text}, {"섹터", type text},
                {"현재가", type number}, {"전일대비", type number}, {"등락률", type number}, {"시가", type number}, {"고가", type number},
                {"저가", type number}, {"전일종가", type number}, {"거래량", type number}, {"거래대금억", type number}, {"전일대비거래량", type number},
                {"시가총액억", type number}, {"PER", type number}, {"PBR", type number}, {"EPS", type number}, {"BPS", type number},
                {"외국인소진율", type number}, {"고52주", type number}, {"고52주일", type date}, {"저52주", type number}, {"저52주일", type date},
                {"외국인당일주", type number}, {"프로그램당일주", type number}, {"외국인5일억", type number}, {"기관5일억", type number},
                {"개인5일억", type number}, {"외국인20일억", type number}, {"기관20일억", type number}, {"상한가", type number},
                {"하한가", type number}, {"유의", type text}, {"DB유의", type text}, {"상태", type text}, {"조회시각", type datetime}})
        in
            Typed,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblQuote", {"구분", "종목코드", "종목명", "시장", "섹터", "현재가", "전일대비", "등락률", "거래대금억", "시가총액억", "PER", "PBR", "외국인소진율", "고52주대비", "저52주대비", "외국인5일억", "기관5일억", "개인5일억", "외국인20일억", "기관20일억", "시가", "고가", "저가", "전일종가", "거래량", "전일대비거래량", "EPS", "BPS", "고52주", "고52주일", "저52주", "저52주일", "외국인당일주", "프로그램당일주", "상한가", "하한가", "유의", "DB유의", "상태", "조회시각"}, Tried[Error][Message]) else Tried[Value]
in
    Final
