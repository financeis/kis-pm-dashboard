// fnDailyChart : 국내주식기간별시세(일봉, 수정주가) [v1_국내주식-016] FHKST03010100
//  - 1회 호출에 최대 100건 → 종료일을 앞으로 옮기며 fromDate까지 반복 조회
//  - 반환: 일자, 시가, 고가, 저가, 종가, 거래량, 거래대금 (일자 오름차순)
(code as text, fromDate as date, toDate as date) as table =>
let
    Path = "/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice",
    Page = (endDate as date) as list =>
        let
            J = fnKisGet(Path, "FHKST03010100", [
                FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code,
                FID_INPUT_DATE_1 = fnYmdText(fromDate), FID_INPUT_DATE_2 = fnYmdText(endDate),
                FID_PERIOD_DIV_CODE = "D", FID_ORG_ADJ_PRC = "0"]),
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
                if [n] >= 15 or List.Count([rows]) < 100 or minD = null or minD <= fromDate then [rows = {}, n = [n] + 1]
                else [rows = Page(Date.AddDays(minD, -1)), n = [n] + 1],
        each [rows]),
    Recs = List.Combine(Pages),
    T = Table.FromRecords(
        List.Transform(Recs, each [
            일자 = fnYmd([stck_bsop_date]),
            시가 = fnNum(Record.FieldOrDefault(_, "stck_oprc", null)),
            고가 = fnNum(Record.FieldOrDefault(_, "stck_hgpr", null)),
            저가 = fnNum(Record.FieldOrDefault(_, "stck_lwpr", null)),
            종가 = fnNum(Record.FieldOrDefault(_, "stck_clpr", null)),
            거래량 = fnNum(Record.FieldOrDefault(_, "acml_vol", null)),
            거래대금 = fnNum(Record.FieldOrDefault(_, "acml_tr_pbmn", null))]),
        type table [일자 = date, 시가 = number, 고가 = number, 저가 = number, 종가 = number, 거래량 = number, 거래대금 = number]),
    Clean = Table.SelectRows(T, each [일자] <> null and [일자] >= fromDate and [종가] <> null and [종가] > 0),
    Dedup = Table.Distinct(Clean, {"일자"}),
    Sorted = Table.Sort(Dedup, {{"일자", Order.Ascending}})
in
    Sorted
