// T_IndexNow : 국내 주요 지수 현재가 + 시장 폭(상승/하락 종목수)  [국내업종 현재지수 v1_국내주식-063] FHPUP02100000
let
    Main =
        let
            Path = "/uapi/domestic-stock/v1/quotations/inquire-index-price",
            List0 = {{"KOSPI", "0001"}, {"KOSDAQ", "1001"}, {"KOSPI200", "2001"}, {"VKOSPI", "0503"}},
            Now = DateTime.LocalNow(),
            Get = (name as text, code as text) as record =>
                let
                    o = fnKisGet(Path, "FHPUP02100000", [FID_COND_MRKT_DIV_CODE = "U", FID_INPUT_ISCD = code])[output],
                    n = (k as text) => fnNum(Record.FieldOrDefault(o, k, null)),
                    pb = n("acml_tr_pbmn"),
                    ppb = n("prdy_tr_pbmn")
                in
                    // 레코드 필드는 지연 평가되므로 응답 o를 먼저 평가해 try가 API 오류를 잡도록 함
                    if not (o is record) then error "빈 응답" else
                    [지수 = name, 코드 = code, 현재 = n("bstp_nmix_prpr"), 전일대비 = n("bstp_nmix_prdy_vrss"),
                     등락률 = let r = n("bstp_nmix_prdy_ctrt") in if r = null then null else r / 100,
                     시가 = n("bstp_nmix_oprc"), 고가 = n("bstp_nmix_hgpr"), 저가 = n("bstp_nmix_lwpr"),
                     거래대금억 = if pb = null then null else pb / 100, 전일거래대금억 = if ppb = null then null else ppb / 100,
                     상승 = n("ascn_issu_cnt"), 상한 = n("uplm_issu_cnt"), 보합 = n("stnr_issu_cnt"),
                     하락 = n("down_issu_cnt"), 하한 = n("lslm_issu_cnt"),
                     연중고가 = n("dryy_bstp_nmix_hgpr"), 연중고가일 = fnYmd(Record.FieldOrDefault(o, "dryy_bstp_nmix_hgpr_date", null)),
                     연중저가 = n("dryy_bstp_nmix_lwpr"), 연중저가일 = fnYmd(Record.FieldOrDefault(o, "dryy_bstp_nmix_lwpr_date", null)),
                     상태 = "OK", 조회시각 = Now],
            Rows = List.Transform(List0, each
                let r = try Get(_{0}, _{1}) in
                    if r[HasError] then [지수 = _{0}, 코드 = _{1}, 상태 = "오류: " & r[Error][Message], 조회시각 = Now] else r[Value]),
            Cols = {"지수", "코드", "현재", "전일대비", "등락률", "시가", "고가", "저가", "거래대금억", "전일거래대금억",
                    "상승", "상한", "보합", "하락", "하한", "연중고가", "연중고가일", "연중저가", "연중저가일", "상태", "조회시각"},
            Failed = List.Select(Rows, each Text.StartsWith(Text.From(Record.FieldOrDefault(_, "상태", "")), "오류")),
            // 전부 실패하면 오류로 올려 직전 데이터를 유지(fnFallback), 일부만 실패하면 그 행만 오류 표시
            Checked = if List.Count(Failed) = List.Count(Rows) then error Text.Middle(Failed{0}[상태], 4) else Rows,
            T = Table.FromRecords(Checked, Cols, MissingField.UseNull),
            Typed = Table.TransformColumnTypes(T, {
                {"지수", type text}, {"코드", type text}, {"현재", type number}, {"전일대비", type number}, {"등락률", type number},
                {"시가", type number}, {"고가", type number}, {"저가", type number}, {"거래대금억", type number}, {"전일거래대금억", type number},
                {"상승", type number}, {"상한", type number}, {"보합", type number}, {"하락", type number}, {"하한", type number},
                {"연중고가", type number}, {"연중고가일", type date}, {"연중저가", type number}, {"연중저가일", type date},
                {"상태", type text}, {"조회시각", type datetime}})
        in
            Typed,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblIndexNow", {"지수", "코드", "현재", "전일대비", "등락률", "시가", "고가", "저가", "거래대금억", "전일거래대금억", "상승", "상한", "보합", "하락", "하한", "연중고가", "연중고가일", "연중저가", "연중저가일", "상태", "조회시각"}, Tried[Error][Message]) else Tried[Value]
in
    Final
