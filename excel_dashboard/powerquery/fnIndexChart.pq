// fnIndexChart : 국내주식업종기간별시세(일) [v1_국내주식-021] FHKUP03500100
//  - 1회 호출에 최대 50건 → 종료일을 앞으로 옮기며 반복 조회
//  - indexCode 예) 0001 KOSPI, 1001 KOSDAQ, 2001 KOSPI200
(indexCode as text, fromDate as date, toDate as date) as table =>
let
    Path = "/uapi/domestic-stock/v1/quotations/inquire-daily-indexchartprice",
    Page = (endDate as date) as list =>
        let
            J = fnKisGet(Path, "FHKUP03500100", [
                FID_COND_MRKT_DIV_CODE = "U", FID_INPUT_ISCD = indexCode,
                FID_INPUT_DATE_1 = fnYmdText(fromDate), FID_INPUT_DATE_2 = fnYmdText(endDate),
                FID_PERIOD_DIV_CODE = "D"]),
            Out = Record.FieldOrDefault(J, "output2", {}),
            L = if Out is list then Out else {}
        in
            List.Select(L, each Text.From(Record.FieldOrDefault(_, "stck_bsop_date", "")) <> ""),
    Pages = List.Generate(
        () => [rows = Page(toDate), n = 1],
        each List.Count([rows]) > 0,
        each
            let
                minD = List.Min(List.Transform([rows], each fnYmd([stck_bsop_date])))
            in
                if [n] >= 20 or List.Count([rows]) < 50 or minD = null or minD <= fromDate then [rows = {}, n = [n] + 1]
                else [rows = Page(Date.AddDays(minD, -1)), n = [n] + 1],
        each [rows]),
    Recs = List.Combine(Pages),
    T = Table.FromRecords(
        List.Transform(Recs, each [
            일자 = fnYmd([stck_bsop_date]),
            시가 = fnNum(Record.FieldOrDefault(_, "bstp_nmix_oprc", null)),
            고가 = fnNum(Record.FieldOrDefault(_, "bstp_nmix_hgpr", null)),
            저가 = fnNum(Record.FieldOrDefault(_, "bstp_nmix_lwpr", null)),
            종가 = fnNum(Record.FieldOrDefault(_, "bstp_nmix_prpr", null)),
            거래량 = fnNum(Record.FieldOrDefault(_, "acml_vol", null)),
            거래대금 = fnNum(Record.FieldOrDefault(_, "acml_tr_pbmn", null))]),
        type table [일자 = date, 시가 = number, 고가 = number, 저가 = number, 종가 = number, 거래량 = number, 거래대금 = number]),
    Clean = Table.SelectRows(T, each [일자] <> null and [일자] >= fromDate and [종가] <> null and [종가] > 0),
    Dedup = Table.Distinct(Clean, {"일자"}),
    Sorted = Table.Sort(Dedup, {{"일자", Order.Ascending}})
in
    Sorted
