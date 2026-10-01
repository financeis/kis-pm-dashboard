// T_Sector : KOSPI·KOSDAQ 업종지수 등락 (섹터 히트맵)  [국내업종 구분별전체시세 v1_국내주식-066] FHPUP02140000
//  - KOSPI: 소속구분 0(전업종) / KOSDAQ: 3(일반구분)+1(기타구분)+0 을 합쳐 업종코드로 필터
let
    Main =
        let
            Path = "/uapi/domestic-stock/v1/quotations/inquire-index-category-price",
            Now = DateTime.LocalNow(),
            Get = (mkt as text, isc as text, blng as text) as list =>
                let
                    J = try fnKisGet(Path, "FHPUP02140000", [FID_COND_MRKT_DIV_CODE = "U", FID_INPUT_ISCD = isc,
                            FID_COND_SCR_DIV_CODE = "20214", FID_MRKT_CLS_CODE = mkt, FID_BLNG_CLS_CODE = blng]) otherwise [],
                    O = Record.FieldOrDefault(J, "output2", {})
                in
                    if O is list then O else {},
            KospiCodes = List.Transform({5 .. 30}, each Text.PadStart(Text.From(_), 4, "0")),
            KosdaqCodes = List.Transform({1006 .. 1033}, Text.From),
            Parents = {"0021", "0027", "1009"},
            SizeCodes = {"0002", "0003", "0004", "1002", "1003", "1004"},
            ToRow = (r as record, market as text) as record =>
                let
                    code = Text.From(Record.FieldOrDefault(r, "bstp_cls_code", "")),
                    pb = fnNum(Record.FieldOrDefault(r, "acml_tr_pbmn", null)),
                    chg = fnNum(Record.FieldOrDefault(r, "bstp_nmix_prdy_ctrt", null))
                in
                    [시장 = market, 코드 = code, 업종명 = Text.Trim(Text.From(Record.FieldOrDefault(r, "hts_kor_isnm", ""))),
                     구분 = if List.Contains(SizeCodes, code) then "규모" else if List.Contains(Parents, code) then "대분류" else "업종",
                     지수 = fnNum(Record.FieldOrDefault(r, "bstp_nmix_prpr", null)),
                     전일대비 = fnNum(Record.FieldOrDefault(r, "bstp_nmix_prdy_vrss", null)),
                     등락률 = if chg = null then null else chg / 100,
                     거래대금억 = if pb = null then null else pb / 100,
                     조회시각 = Now],
            Kospi = List.Transform(List.Select(Get("K", "0001", "0"), each
                        List.Contains(KospiCodes & SizeCodes, Text.From(Record.FieldOrDefault(_, "bstp_cls_code", "")))), each ToRow(_, "KOSPI")),
            KosdaqRaw = Get("Q", "1001", "3") & Get("Q", "1001", "1") & Get("Q", "1001", "0"),
            Kosdaq = List.Transform(List.Select(KosdaqRaw, each
                        List.Contains(KosdaqCodes & SizeCodes, Text.From(Record.FieldOrDefault(_, "bstp_cls_code", "")))), each ToRow(_, "KOSDAQ")),
            T = Table.FromRecords(Kospi & Kosdaq, {"시장", "코드", "업종명", "구분", "지수", "전일대비", "등락률", "거래대금억", "조회시각"}, MissingField.UseNull),
            // 모든 호출이 실패해 빈 표가 되면 오류로 올려 직전 데이터를 유지(fnFallback)
            Dedup = if Table.RowCount(T) = 0 then error "업종 시세를 받지 못했습니다" else Table.Distinct(T, {"시장", "코드"}),
            // 시장별 거래대금 비중 (업종 구분만)
            Totals = Table.Group(Table.SelectRows(Dedup, each [구분] = "업종"), {"시장"}, {{"합계", each List.Sum([거래대금억]), type number}}),
            Joined = Table.NestedJoin(Dedup, {"시장"}, Totals, {"시장"}, "t", JoinKind.LeftOuter),
            WithShare = Table.AddColumn(Table.ExpandTableColumn(Joined, "t", {"합계"}), "거래비중", each
                if [구분] <> "업종" or [합계] = null or [합계] = 0 or [거래대금억] = null then null else [거래대금억] / [합계], type number),
            Final = Table.SelectColumns(WithShare, {"시장", "코드", "업종명", "구분", "지수", "전일대비", "등락률", "거래대금억", "거래비중", "조회시각"}),
            Typed = Table.TransformColumnTypes(Final, {{"시장", type text}, {"코드", type text}, {"업종명", type text}, {"구분", type text},
                {"지수", type number}, {"전일대비", type number}, {"등락률", type number}, {"거래대금억", type number}, {"조회시각", type datetime}}),
            Sorted = Table.Sort(Typed, {{"시장", Order.Ascending}, {"등락률", Order.Descending}})
        in
            Sorted,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblSector", {"시장", "코드", "업종명", "구분", "지수", "전일대비", "등락률", "거래대금억", "거래비중", "조회시각"}, Tried[Error][Message]) else Tried[Value]
in
    Final
