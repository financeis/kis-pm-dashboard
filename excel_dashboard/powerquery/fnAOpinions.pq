// fnAOpinions : 종목투자의견 [국내주식-188] FHKST663300C0 /uapi/domestic-stock/v1/quotations/invest-opinion 기간 전체 (2026-10-01)
//  - 1회 호출은 기간 안의 최신 100건만 준다(일자 내림차순, 2026-10-01 실측: 005930 2025-04-01~2026-10-01 → 2026-09-23~2026-01-30 100건).
//    100건째가 같은 날짜의 일부만 담고 잘리므로(1월 30일 16건 중 3건만 옴) 다음 호출은 종료일 = 직전 결과의 가장 오래된 날짜(포함)로
//    다시 받아 겹친 행을 중복 제거한다. 100건 미만이거나 시작일에 닿거나 날짜가 더 내려가지 않으면 멈춤(최대 8회).
//  - 날짜는 8자리(YYYYMMDD) — 저장소 예시의 "00"+8자리와 같은 결과(실측).
//  - 반환: 표 [일자, 증권사, 투자의견, 직전투자의견, 목표가, 전일종가] — 일자 내림차순(같은 날은 응답 순서), 목표가·전일종가 0·빈 값 → null.
//    증권사 = mbcr_name(저장소 열 매핑에 없는 필드, 실측), 목표가 = hts_goal_prc, 투자의견 = invt_opnn(BUY·매수·Buy 등 증권사 표기 그대로).
(code as text, fromDate as date, toDate as date) as table =>
let
    Path = "/uapi/domestic-stock/v1/quotations/invest-opinion",
    Txt = (r as record, k as text) as nullable text =>
        let v = Record.FieldOrDefault(r, k, null), t = if v = null then "" else Text.Trim(Text.From(v)) in if t = "" then null else t,
    Pos = (v as nullable number) as nullable number => if v = null or v <= 0 then null else v,
    Page = (endDate as date) as list =>
        let
            J = fnKisGet(Path, "FHKST663300C0", [
                    FID_COND_MRKT_DIV_CODE = "J", FID_COND_SCR_DIV_CODE = "16633", FID_INPUT_ISCD = code,
                    FID_INPUT_DATE_1 = fnYmdText(fromDate), FID_INPUT_DATE_2 = fnYmdText(endDate)]),
            O = Record.FieldOrDefault(J, "output", {}),
            L = if O is list then O else if O is record then {O} else {}
        in
            List.Buffer(List.Select(L, each _ is record and fnYmd(Record.FieldOrDefault(_, "stck_bsop_date", null)) <> null)),
    MinDate = (rows as list) as date => List.Min(List.Transform(rows, each fnYmd(Record.FieldOrDefault(_, "stck_bsop_date", null)))),
    Pages = List.Generate(
        () => [rows = Page(toDate), n = 1, prevMin = null],
        each List.Count([rows]) > 0,
        each
            let m = MinDate([rows]) in
                if [n] >= 8 or List.Count([rows]) < 100 or m <= fromDate or m = [prevMin] then [rows = {}, n = [n] + 1, prevMin = m]
                else [rows = Page(m), n = [n] + 1, prevMin = m],
        each [rows]),
    All = List.Buffer(List.Combine(Pages)),
    Rows = List.Transform(List.Positions(All), (i) =>
        let r = All{i} in
            {fnYmd(Record.FieldOrDefault(r, "stck_bsop_date", null)), Txt(r, "mbcr_name"), Txt(r, "invt_opnn"), Txt(r, "rgbf_invt_opnn"),
             Pos(fnNum(Record.FieldOrDefault(r, "hts_goal_prc", null))), Pos(fnNum(Record.FieldOrDefault(r, "stck_prdy_clpr", null))), i}),
    T = #table(type table [일자 = date, 증권사 = text, 투자의견 = text, 직전투자의견 = text, 목표가 = number, 전일종가 = number, __i = number], Rows),
    InRange = Table.SelectRows(T, each [일자] >= fromDate and [일자] <= toDate),
    Dedup = Table.Distinct(InRange, {"일자", "증권사", "투자의견", "직전투자의견", "목표가", "전일종가"}),
    Sorted = Table.Sort(Dedup, {{"일자", Order.Descending}, {"__i", Order.Ascending}})
in
    Table.Buffer(Table.RemoveColumns(Sorted, {"__i"}))
