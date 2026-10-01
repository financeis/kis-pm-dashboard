// fnAQuote : 종목분석용 현재가·최근 세션 일자 (국내주식기간별시세 일봉 FHKST03010100, 최근 20일 범위 1회 호출) (2026-10-01)
//  - 반환 레코드:
//      기준일   최근 일봉 일자 = 당일 자료가 속한 세션. 장중·장 마감 뒤에는 오늘(일봉에 오늘 행이 장중에도 있음 — 2026-10-01 11:40 실측),
//               장 시작 전·휴장일에는 직전 세션. 체결금액별·당일 매물대·추정가집계 표의 '기준일'(당일 표시)로 쓴다.
//      현재가   output1.stck_prpr (0·빈 값 → null). 추정가집계 ≈금액·KIS 추정 PER(현재가 기준)에 쓴다.
//      종목명   output1.hts_kor_isnm
//  - 없는 종목코드는 rt_cd 0에 0·빈 값으로 오므로(실측) 현재가·기준일이 null이 된다.
(code as text, today as date) as record =>
let
    J = fnKisGet("/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice", "FHKST03010100", [
            FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code,
            FID_INPUT_DATE_1 = fnYmdText(Date.AddDays(today, -20)), FID_INPUT_DATE_2 = fnYmdText(today),
            FID_PERIOD_DIV_CODE = "D", FID_ORG_ADJ_PRC = "0"]),
    O1 = let o = Record.FieldOrDefault(J, "output1", null) in if o is record then o else [],
    O2 = let o = Record.FieldOrDefault(J, "output2", null) in if o is list then o else {},
    Bars = List.Select(O2, each _ is record
            and (let c = fnNum(Record.FieldOrDefault(_, "stck_clpr", null)) in c <> null and c > 0)),
    Dates = List.RemoveNulls(List.Transform(Bars, each fnYmd(Record.FieldOrDefault(_, "stck_bsop_date", null)))),
    Price = let v = fnNum(Record.FieldOrDefault(O1, "stck_prpr", null)) in if v = null or v <= 0 then null else v,
    Name = let v = Record.FieldOrDefault(O1, "hts_kor_isnm", null), t = if v = null then "" else Text.Trim(Text.From(v))
           in if t = "" then null else t
in
    [기준일 = if List.IsEmpty(Dates) then null else List.Max(Dates), 현재가 = Price, 종목명 = Name]
