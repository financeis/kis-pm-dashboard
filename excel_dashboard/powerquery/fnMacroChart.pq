// fnMacroChart : 해외주식 종목/지수/환율기간별시세(일) [v1_해외주식-012] FHKST03030100
//  - mkt: N 해외지수, X 환율, I 국채/금리  / symbol 예) SPX, COMP, .DJI, FX@KRW, Y0202(미국10년), Y0101(국고3년)
//  - 반환: 일자, 종가 (일자 오름차순)
(mkt as text, symbol as text, fromDate as date, toDate as date) as table =>
let
    Path = "/uapi/overseas-price/v1/quotations/inquire-daily-chartprice",
    Page = (endDate as date) as list =>
        let
            J = fnKisGet(Path, "FHKST03030100", [
                FID_COND_MRKT_DIV_CODE = mkt, FID_INPUT_ISCD = symbol,
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
                if [n] >= 10 or List.Count([rows]) < 100 or minD = null or minD <= fromDate then [rows = {}, n = [n] + 1]
                else [rows = Page(Date.AddDays(minD, -1)), n = [n] + 1],
        each [rows]),
    Recs = List.Combine(Pages),
    T = Table.FromRecords(
        List.Transform(Recs, each [
            일자 = fnYmd([stck_bsop_date]),
            종가 = fnNum(Record.FieldOrDefault(_, "ovrs_nmix_prpr", null))]),
        type table [일자 = date, 종가 = number]),
    Clean = Table.SelectRows(T, each [일자] <> null and [일자] >= fromDate and [종가] <> null and [종가] <> 0),
    Dedup = Table.Distinct(Clean, {"일자"}),
    Sorted = Table.Sort(Dedup, {{"일자", Order.Ascending}})
in
    Sorted
