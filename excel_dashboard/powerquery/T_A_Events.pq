// T_A_Events : 종목분석 10구역 — 이 종목의 기업 이벤트, 오늘 − 7일 ~ 오늘 + 60일 (spec R17-10, R11) → tblA_Events (2026-10-01 T15)
//  - 기간: 설정 event_days_back(기본 7)·event_days_ahead(기본 60). 원천 = 예탁원정보 6종(SHT_CD = 종목코드, 각 1회 — 호출 6회):
//      상장정보일정 HHKDB669107C0 list-info     일자 = 상장일 list_dt, 사유 issue_type으로 종류 분류:
//          사유에 CB·BW가 있으면(국내CB행사·BW행사 등) "CB·BW 전환 상장", '신규'가 있으면 "신규상장".
//          그 밖의 사유(유상증자 신주 상장·STOCKOPTION행사·주식전환·액면병합·상호변경 등)는 spec R11 종류에 없어 넣지 않는다.
//          종목을 지정해도 해당 없으면 빈 칸만 든 100행이 오므로(2026-10-01 실측) 종목코드·일자가 있는 행만 쓴다.
//      유상증자일정 HHKDB669100C0 paidin-capin (GB1 2 = 기준일별)  일자 = 신주배정기준일 record_date → "유상증자"
//      무상증자일정 HHKDB669101C0 bonus-issue                     일자 = record_date → "무상증자"
//      합병/분할일정 HHKDB669104C0 merger-split                   일자 = record_date → "합병·분할"
//      배당일정 HHKDB669102C0 dividend (GB1 0 = 전체)             일자 = record_date → "배당 기준일"
//      주주총회일정 HHKDB669111C0 sharehld-meet                   일자 = 주총일자 gen_meet_dt → "주주총회"
//          (조회 기간은 기준일 record_date에 걸리므로(실측) 기간 시작 90일 전부터 받아 주총일자로 다시 거름)
//  - 상세: 원천 값을 사람이 읽는 한 줄로(주식 수·배정률·발행가·권리락·청약·상장·지급일·주총 의안 등, 값이 있는 것만).
//  - D-day = 일자 − 오늘(일, 앞으로 남은 날이 양수 — 페이지에서 "D-"0 형식; 지난 이벤트는 음수).
//  - 같은 일자·종류·상세는 1건. 정렬: 일자 오름차순 → 종류. 이벤트가 없으면 "데이터 없음" 1행.
//  - 실패: 6종이 모두 실패하면 오류 → 직전 표 유지(fnARun). 일부만 실패하면 받은 이벤트 뒤에 실패한 원천마다 1행
//    (일자 빈칸, 종류 = 그 원천의 종류, 상태 "오류: <원천> 조회 실패(사유)").
let
    Types = {{"종목코드", type text}, {"일자", type date}, {"D-day", Int64.Type}, {"종류", type text}, {"상세", type text},
             {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Today = Date.From(now),
            Back = fnASetting("event_days_back", 7, 0, 366),
            Ahead = fnASetting("event_days_ahead", 60, 0, 366),
            F = Date.AddDays(Today, -Back),
            T = Date.AddDays(Today, Ahead),
            K = "/uapi/domestic-stock/v1/ksdinfo/",
            Get = (path as text, tr as text, params as record) as list =>
                let
                    J = fnKisGet(K & path, tr, params),
                    O = Record.FieldOrDefault(J, "output1", {}),
                    L = List.Select(if O is list then O else if O is record then {O} else {}, each _ is record)
                in
                    // 이 종목 행만(빈 자리 행·다른 종목 제외)
                    List.Buffer(List.Select(L, each let s = Text.Upper(Tx(_, "sht_cd")) in s = code or s = "A" & code)),
            Tx = (r as record, k as text) as text =>
                let v = Record.FieldOrDefault(r, k, null) in if v = null then "" else try Text.Trim(Text.From(v)) otherwise "",
            Num = (r as record, k as text) as nullable number => fnNum(Record.FieldOrDefault(r, k, null)),
            D8 = (s as text) as nullable date => fnYmd(Text.Start(Text.Select(s, {"0".."9"}), 8)),
            Ds = (s as text) as text => let d = D8(s) in if d = null then s else Date.ToText(d, [Format = "yyyy-MM-dd", Culture = "en-US"]),
            N0 = (v as nullable number) as text => if v = null then "" else Number.ToText(v, "N0", "en-US"),
            P2 = (v as nullable number) as text => if v = null then "" else Number.ToText(v, "0.##", "en-US"),
            Join = (parts as list) as text => Text.Combine(List.Select(parts, each _ <> null and _ <> ""), ", "),
            Opt = (cond as logical, t as text) as nullable text => if cond then t else null,
            Has = (r as record, k as text) as logical => Tx(r, k) <> "",
            Shares = (r as record, k as text, rateKey as nullable text) as nullable text =>
                let q = Num(r, k), rt = if rateKey = null then null else Num(r, rateKey)
                in if q = null or q <= 0 then null
                   else "신주 " & N0(q) & "주" & (if rt <> null and rt > 0 then "(배정률 " & P2(rt) & "%)" else ""),
            Ymd = fnYmdText,

            ListInfo = () as list => List.Transform(
                Get("list-info", "HHKDB669107C0", [SHT_CD = code, T_DT = Ymd(T), F_DT = Ymd(F), CTS = ""]), (r) =>
                    let
                        why = Tx(r, "issue_type"),
                        u = Text.Upper(why),
                        kind = if Text.Contains(u, "CB") or Text.Contains(u, "BW") then "CB·BW 전환 상장"
                               else if Text.Contains(why, "신규") then "신규상장" else null,
                        q = Num(r, "issue_stk_qty"),
                        p = Num(r, "issue_price")
                    in
                        {D8(Tx(r, "list_dt")), kind,
                         Join({why, Opt(q <> null and q > 0, N0(q) & "주"), Opt(p <> null and p > 0, "발행가 " & N0(p) & "원")})}),
            PaidIn = () as list => List.Transform(
                Get("paidin-capin", "HHKDB669100C0", [CTS = "", GB1 = "2", F_DT = Ymd(F), T_DT = Ymd(T), SHT_CD = code]), (r) =>
                    let p = Num(r, "fix_price") in
                        {D8(Tx(r, "record_date")), "유상증자",
                         Join({Shares(r, "issue_stk_qty", "fix_rate"), Opt(p <> null and p > 0, "발행예정가 " & N0(p) & "원"),
                               Opt(Has(r, "right_dt"), "권리락 " & Ds(Tx(r, "right_dt"))), Opt(Has(r, "sub_term"), "청약 " & Tx(r, "sub_term")),
                               Opt(Has(r, "list_date"), "상장 " & Ds(Tx(r, "list_date")))})}),
            Bonus = () as list => List.Transform(
                Get("bonus-issue", "HHKDB669101C0", [CTS = "", F_DT = Ymd(F), T_DT = Ymd(T), SHT_CD = code]), (r) =>
                    {D8(Tx(r, "record_date")), "무상증자",
                     Join({Shares(r, "issue_stk_qty", "fix_rate"), Opt(Has(r, "right_dt"), "권리락 " & Ds(Tx(r, "right_dt"))),
                           Opt(Has(r, "list_date"), "상장 " & Ds(Tx(r, "list_date")))})}),
            Merger = () as list => List.Transform(
                Get("merger-split", "HHKDB669104C0", [CTS = "", F_DT = Ymd(F), T_DT = Ymd(T), SHT_CD = code]), (r) =>
                    let rt = Num(r, "merge_rate"), who = Join({Tx(r, "opp_cust_nm"), Tx(r, "cust_nm")}) in
                        {D8(Tx(r, "record_date")), "합병·분할",
                         Join({Tx(r, "merge_type") & (if Has(r, "opp_cust_nm") and Has(r, "cust_nm")
                                                      then ": " & Tx(r, "opp_cust_nm") & " → " & Tx(r, "cust_nm")
                                                      else if who <> "" then ": " & who else ""),
                               Opt(rt <> null and rt > 0, "비율 " & P2(rt)), Opt(Has(r, "td_stop_dt"), "거래정지 " & Tx(r, "td_stop_dt")),
                               Opt(Has(r, "list_dt"), "상장 " & Ds(Tx(r, "list_dt")))})}),
            Dividend = () as list => List.Transform(
                Get("dividend", "HHKDB669102C0", [CTS = "", GB1 = "0", F_DT = Ymd(F), T_DT = Ymd(T), SHT_CD = code, HIGH_GB = ""]), (r) =>
                    let a = Num(r, "per_sto_divi_amt"), cr = Num(r, "divi_rate"), sr = Num(r, "stk_divi_rate"), kind = Tx(r, "divi_kind") in
                        {D8(Tx(r, "record_date")), "배당 기준일",
                         Join({(if kind = "" then "" else kind & " ") & "배당" & (if a <> null and a > 0 then " 주당 " & N0(a) & "원" else "(금액 미정)"),
                               Opt(cr <> null and cr > 0, "현금배당률 " & P2(cr) & "%"), Opt(sr <> null and sr > 0, "주식배당률 " & P2(sr) & "%"),
                               Opt(Has(r, "divi_pay_dt"), "지급 " & Ds(Tx(r, "divi_pay_dt")))})}),
            Meeting = () as list => List.Transform(
                Get("sharehld-meet", "HHKDB669111C0", [CTS = "", F_DT = Ymd(Date.AddDays(F, -90)), T_DT = Ymd(T), SHT_CD = code]), (r) =>
                    {D8(Tx(r, "gen_meet_dt")), "주주총회",
                     Join({Tx(r, "gen_meet_type") & (if Has(r, "agenda") then ": " & Tx(r, "agenda") else ""),
                           Opt(Has(r, "record_date"), "기준일 " & Ds(Tx(r, "record_date")))})}),

            Sources = {[이름 = "CB·BW 전환 상장·신규상장", f = ListInfo], [이름 = "유상증자", f = PaidIn], [이름 = "무상증자", f = Bonus],
                       [이름 = "합병·분할", f = Merger], [이름 = "배당 기준일", f = Dividend], [이름 = "주주총회", f = Meeting]},
            // 원천마다 한 번만 호출되도록 결과를 버퍼로 고정(try 안에서 강제 평가 → 호출 오류가 그 원천의 실패로 잡힘)
            Results = List.Buffer(List.Transform(Sources, (s) => [이름 = s[이름], r = try List.Buffer(List.Transform(s[f](), List.Buffer))])),
            Failed = List.Select(Results, each [r][HasError]),
            Reason = (e as record) as text => let m = try Text.Trim(Text.From(e[Message])) otherwise "" in if m = "" then "조회 실패" else m,
            Events = List.Select(List.Combine(List.Transform(List.Select(Results, each not [r][HasError]), each [r][Value])),
                        each _{0} <> null and _{1} <> null and _{0} >= F and _{0} <= T),
            Rows = List.Transform(Events, (e) => {code, e{0}, Duration.Days(e{0} - Today), e{1}, if e{2} = "" then null else e{2}, "OK", now}),
            Cols = List.Transform(Types, each _{0}),
            Ev = Table.Sort(Table.Distinct(#table(Cols, Rows), {"일자", "종류", "상세"}), {{"일자", Order.Ascending}, {"종류", Order.Ascending}}),
            ErrRows = #table(Cols, List.Transform(Failed, (x) =>
                        {code, null, null, x[이름], null, "오류: " & x[이름] & " 조회 실패(" & Text.Start(Reason(x[r][Error]), 80) & ")", now}))
        in
            if List.Count(Failed) = List.Count(Sources) then error Error.Record("조회 실패", Reason(Failed{0}[r][Error]))
            else if List.IsEmpty(Rows) and List.IsEmpty(Failed) then fnAOneRow(Types, code, "데이터 없음", now)
            else Table.Combine({Ev, ErrRows})
in
    fnARun("tblA_Events", Types, Main)
