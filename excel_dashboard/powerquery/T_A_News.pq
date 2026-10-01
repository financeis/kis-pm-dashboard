// T_A_News : 종목분석 9구역 — 이 종목의 뉴스·공시 제목 최근 40건 (spec R17-9, R12) → tblA_News (2026-10-01 T15)
//  - 원천: 종합 시황/공시(제목) [국내주식-141] FHKST01011800 /uapi/domestic-stock/v1/quotations/news-title — FID_INPUT_ISCD만 지정(1회 = 40건).
//    T_News(보유·관심 종목 묶음)와 같은 해석: data_dt(YYYYMMDD) + data_tm(HHMMSS) → 일시, hts_pbnt_titl_cntt → 제목, dorg(자료원) → 출처.
//  - 같은 일시·제목·출처는 1건. 정렬: 일시 내림차순(같으면 기사 일련번호 cntt_usiq_srno 내림차순). 0건이면 "데이터 없음" 1행
//    (없는 종목코드도 rt_cd 0에 빈 목록으로 옴 — 999999 실측).
let
    Types = {{"종목코드", type text}, {"일시", type datetime}, {"제목", type text}, {"출처", type text},
             {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Keep = 40,
            J = fnKisGet("/uapi/domestic-stock/v1/quotations/news-title", "FHKST01011800", [FID_NEWS_OFER_ENTP_CODE = "",
                    FID_COND_MRKT_CLS_CODE = "", FID_INPUT_ISCD = code, FID_TITL_CNTT = "", FID_INPUT_DATE_1 = "",
                    FID_INPUT_HOUR_1 = "", FID_RANK_SORT_CLS_CODE = "", FID_INPUT_SRNO = ""]),
            O = Record.FieldOrDefault(J, "output", {}),
            L = List.Buffer(List.Select(if O is list then O else if O is record then {O} else {}, each _ is record)),
            Txt = (r as record, k as text) as text =>
                let v = Record.FieldOrDefault(r, k, null) in if v = null then "" else try Text.Trim(Text.From(v)) otherwise "",
            Blank = (t as text) as nullable text => if t = "" then null else t,
            // 작성일자 + 작성시간(HHMMSS, 4자리면 HHMM) → datetime. 날짜가 잘못되면 null, 시간만 잘못되면 00:00:00
            ToDt = (d as text, t as text) as nullable datetime =>
                let
                    day = fnYmd(d),
                    s = Text.Select(t, {"0".."9"}),
                    hms = if Text.Length(s) = 6 then s else if Text.Length(s) = 4 then s & "00" else "000000",
                    tm = try #time(Number.From(Text.Start(hms, 2)), Number.From(Text.Middle(hms, 2, 2)), Number.From(Text.End(hms, 2)))
                         otherwise #time(0, 0, 0)
                in
                    if day = null then null else day & tm,
            Rows = List.Transform(L, (r) => {code, ToDt(Txt(r, "data_dt"), Txt(r, "data_tm")), Blank(Txt(r, "hts_pbnt_titl_cntt")),
                        Blank(Txt(r, "dorg")), "OK", now, Txt(r, "cntt_usiq_srno")}),
            T = #table(List.Transform(Types, each _{0}) & {"__srno"}, Rows),
            Dedup = Table.Distinct(T, {"일시", "제목", "출처"}),
            Latest = Table.FirstN(Table.Sort(Dedup, {{"일시", Order.Descending}, {"__srno", Order.Descending}}), Keep)
        in
            if List.IsEmpty(L) then fnAOneRow(Types, code, "데이터 없음", now)
            else Table.RemoveColumns(Latest, {"__srno"})
in
    fnARun("tblA_News", Types, Main)
