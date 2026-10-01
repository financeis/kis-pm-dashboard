// T_A_Events : 종목분석 10구역 — 이 종목의 기업 이벤트, 오늘 − 7일 ~ 오늘 + 60일 (spec R17-10, R11) → tblA_Events (2026-10-01 T15)
//  - 기간: 설정 event_days_back(기본 7)·event_days_ahead(기본 60). 종류·일자·상세는 [전체]의 T_Events(→ tblEvents, T11)와 같은 규칙을
//    한 종목에 적용한다(두 페이지가 같은 종목에 같은 이벤트를 보이도록 — T11의 2026-10-01 실측 해석을 그대로 따름).
//  - 원천 = 예탁원정보 6종, 각 1회(호출 6회). 종목을 지정한 조회(SHT_CD)는 연속 조회 없이 한 쪽(실측, T11)이다.
//      종류                일자                                                    API (tr_id) · 조회 기간이 걸리는 필드
//      CB·BW 전환 상장     상장일 list_dt                                          상장정보일정 list-info (HHKDB669107C0) · list_dt. 사유 issue_type에 CB·BW·'전환'
//      신규상장            상장일 list_dt                                          상장정보일정 · 사유 '통일교체'(신규상장 표지). 같은 날 '유상증자' 행 = 공모(상세에 넣고 따로 안 냄)
//      유상증자            신주 상장일 list_date, 없으면 청약 시작일 sub_term_ft     유상증자일정 paidin-capin (HHKDB669100C0, GB1 1 = 청약일별) · 청약 시작일(시작 − 45일부터)
//                          (+ 상장정보일정의 사유 '유상증자' 신주 상장일 — 같은 일자는 유상증자일정 행만 남김)
//      무상증자            신주 상장일 list_date, 없으면 배정 기준일 record_date     무상증자일정 bonus-issue (HHKDB669101C0) · 기준일(시작 − 35일부터)
//                          (+ 상장정보일정의 사유 '무상증자' 신주 상장일 — 같은 일자는 무상증자일정 행만 남김)
//      합병·분할           합병·분할 기일 record_date                               합병/분할일정 merger-split (HHKDB669104C0) · 기준일. 이 종목이 피합병(분할)회사(sht_cd)이거나
//                          (+ 상장정보일정의 사유 합병·회사분할 신주 상장일 = 별도)   합병(분할)회사(cust_cd 5자리 + '0' = 보통주)인 행 → 종목 지정 조회로는 합병회사 쪽이
//                                                                                   안 잡히므로 이 API만 시장 전체로 1회(한 쪽, 100행까지) 받아 거른다(기간 행이 적음 — 2026-10-01 3행)
//      배당 기준일         배당 기준일 record_date                                  배당일정 dividend (HHKDB669102C0, GB1 0 = 전체) · 기준일
//      주주총회            총회일 gen_meet_dt(9999년 = 미정은 버림)                  주주총회일정 sharehld-meet (HHKDB669111C0) · 기준일 record_date — 총회일이 기준일보다
//                                                                                   최대 약 90일 뒤라 기준일은 시작 − 100일부터 받아 총회일로 거름. 안건 행은 (총회일, 구분)마다 1건
//    상장정보일정에서 버리는 사유: STOCKOPTION행사, 액면분할·액면병합, 상호변경, 자본감소, '-', 배당(주식배당 상장) 등 7종에 없는 것.
//    빈 자리 행(종목코드 빈칸 — 해당 없을 때 100행이 빈칸으로 옴)은 버린다.
//  - 상세: T11과 같은 형식(" · "로 이음, 수량 #,##0, 날짜 M/d, 150자에서 자름).
//  - D-day = 일자 − 오늘(일, 앞으로 남은 날이 양수 — 페이지에서 "D-"0 형식; 지난 이벤트는 음수).
//  - 같은 일자·종류·상세는 1건. 정렬: 일자 → 종류 → 상세 오름차순. 이벤트가 없으면 "데이터 없음" 1행.
//  - 실패: 6종이 모두 실패하면 오류 → 직전 표 유지(fnARun). 일부만 실패하면 받은 이벤트 뒤에 실패한 원천마다 1행
//    (일자 빈칸, 종류 = 그 원천, 상태 "오류: <원천> 조회 실패(사유)").
let
    Types = {{"종목코드", type text}, {"일자", type date}, {"D-day", Int64.Type}, {"종류", type text}, {"상세", type text},
             {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Today = Date.From(now),
            Back = fnASetting("event_days_back", 7, 0, 366),
            Ahead = fnASetting("event_days_ahead", 60, 0, 366),
            WStart = Date.AddDays(Today, -Back),
            WEnd = Date.AddDays(Today, Ahead),
            K = "/uapi/domestic-stock/v1/ksdinfo/",
            F = (d as date) as text => fnYmdText(d),

            // ---- 도우미 (T11 T_Events와 같은 표기)
            Txt = (r as record, k as text) as text =>
                let v = Record.FieldOrDefault(r, k, null) in if v = null then "" else try Text.Trim(Text.From(v)) otherwise "",
            Blank = (t as nullable text) as nullable text => if t = null or Text.Trim(t) = "" then null else Text.Trim(t),
            CodeOf = (r as record) as nullable text => Blank(Text.Upper(Txt(r, "sht_cd"))),
            // "20261125", "2026/11/25", "2026/11/25 ~ 2026/11/26"의 첫 날짜 / 두 번째 날짜. 잘못된 값·9999년(미정)은 null
            DateOf = (r as record, k as text) as nullable date =>
                let d = fnYmd(Text.Start(Text.Select(Txt(r, k), {"0".."9"}), 8)) in if d = null or Date.Year(d) >= 9999 then null else d,
            Date2Of = (r as record, k as text) as nullable date =>
                let t = Text.Select(Txt(r, k), {"0".."9"}), d = if Text.Length(t) >= 16 then fnYmd(Text.Middle(t, 8, 8)) else null
                in if d = null or Date.Year(d) >= 9999 then null else d,
            Pos = (r as record, k as text) as nullable number =>
                let v = fnNum(Record.FieldOrDefault(r, k, null)) in if v = null or v <= 0 then null else v,
            N0 = (v as nullable number) as nullable text => if v = null then null else Number.ToText(Number.Round(v, 0), "#,##0", "en-US"),
            Dec = (v as nullable number) as nullable text => if v = null then null else Number.ToText(Number.Round(v, 2), "0.##", "en-US"),
            MD = (d as nullable date) as nullable text => if d = null then null else Date.ToText(d, [Format = "M/d", Culture = "en-US"]),
            Join = (parts as list) as text => Text.Combine(List.RemoveNulls(parts), " · "),
            Cut = (t as text) as text => if Text.Length(t) > 150 then Text.Start(t, 149) & "…" else t,

            // 응답 output1의 레코드(종목코드가 있는 행만)
            Get = (path as text, tr as text, params as record) as list =>
                let
                    J = fnKisGet(K & path, tr, params),
                    O = Record.FieldOrDefault(J, "output1", {}),
                    L = if O is list then O else if O is record then {O} else {}
                in
                    List.Buffer(List.Select(L, each _ is record and CodeOf(_) <> null)),
            Mine = (rows as list) as list => List.Select(rows, each CodeOf(_) = code),

            // 이벤트 행 = {일자, 종류, 상세, 우선} — 우선 1 = 전용 API, 2 = 상장정보일정(같은 일자의 증자는 전용 API 쪽만 남김)
            ListInfo = () as list =>
                let
                    Rs = List.Buffer(Mine(Get("list-info", "HHKDB669107C0", [SHT_CD = code, T_DT = F(WEnd), F_DT = F(WStart), CTS = ""]))),
                    Typ = (r as record) as text => Txt(r, "issue_type"),
                    IpoDays = List.Buffer(List.Distinct(List.Transform(List.Select(Rs, each Typ(_) = "통일교체"), each Txt(_, "list_dt")))),
                    Ev = (r as record) as nullable list =>
                        let
                            t = Typ(r), up = Text.Upper(t), d = DateOf(r, "list_dt"),
                            qty = Pos(r, "issue_stk_qty"), px = Pos(r, "issue_price"), tot = Pos(r, "tot_issue_stk_qty"),
                            qtyTxt = if qty = null then null else N0(qty) & "주" & (if px = null then "" else " @ " & N0(px) & "원"),
                            totTxt = if tot = null then null else "상장 후 총 " & N0(tot) & "주",
                            kind =
                                if t = "통일교체" then "신규상장"
                                else if Text.Contains(up, "CB") or Text.Contains(up, "BW") or Text.Contains(t, "전환") then "CB·BW 전환 상장"
                                else if t = "유상증자" then (if List.Contains(IpoDays, Txt(r, "list_dt")) then null else "유상증자")
                                else if t = "무상증자" then "무상증자"
                                else if Text.Contains(t, "합병") or (Text.EndsWith(t, "분할") and not Text.Contains(t, "액면")) then "합병·분할"
                                else null,
                            Pub = List.Select(Rs, each Txt(_, "list_dt") = Txt(r, "list_dt") and Typ(_) = "유상증자"),
                            pubQty = if List.IsEmpty(Pub) then null else Pos(Pub{0}, "issue_stk_qty"),
                            pubPx = if List.IsEmpty(Pub) then null else Pos(Pub{0}, "issue_price"),
                            pubTxt = if pubQty = null then null else "공모 " & N0(pubQty) & "주" & (if pubPx = null then "" else " @ " & N0(pubPx) & "원"),
                            detail =
                                if kind = "신규상장" then Join({"신규상장", pubTxt, if tot = null then null else "상장주식 " & N0(tot) & "주"})
                                else if kind = "CB·BW 전환 상장" then Join({t, qtyTxt, totTxt})
                                else if kind = "유상증자" then Join({"신주 상장", qtyTxt, totTxt})
                                else if kind = "무상증자" then Join({"신주 상장", if qty = null then null else N0(qty) & "주", totTxt})
                                else Join({t & " 신주 상장", if qty = null then null else N0(qty) & "주", totTxt})
                        in
                            if kind = null or d = null then null else {d, kind, Cut(detail), 2}
                in
                    List.RemoveNulls(List.Transform(Rs, Ev)),
            PaidIn = () as list => List.RemoveNulls(List.Transform(
                Mine(Get("paidin-capin", "HHKDB669100C0", [CTS = "", GB1 = "1", F_DT = F(Date.AddDays(WStart, -45)), T_DT = F(WEnd), SHT_CD = code])),
                (r) =>
                    let
                        lst = DateOf(r, "list_date"), s1 = DateOf(r, "sub_term_ft"), s2 = Date2Of(r, "sub_term"),
                        s1b = if s1 = null then DateOf(r, "sub_term") else s1,
                        sub = if s1b = null then null else "청약 " & MD(s1b) & (if s2 <> null and s2 <> s1b then "~" & MD(s2) else ""),
                        d = if lst <> null then lst else s1b,
                        qty = Pos(r, "issue_stk_qty"), rate = Pos(r, "fix_rate"), px = Pos(r, "fix_price"), disc = Pos(r, "disc_rate"),
                        rl = DateOf(r, "right_dt"), rec = DateOf(r, "record_date"),
                        common = {if qty = null then null else "신주 " & N0(qty) & "주", if rate = null then null else "배정 " & Dec(rate) & "%",
                                  if px = null then null else "발행가 " & N0(px) & "원", if disc = null then null else "할인율 " & Dec(disc) & "%"},
                        detail = if lst <> null then Join({"신주 상장"} & common & {sub})
                                 else Join({sub} & common & {if rl = null then null else "권리락 " & MD(rl), if rec = null then null else "기준일 " & MD(rec)})
                    in
                        if d = null then null else {d, "유상증자", Cut(detail), 1})),
            Bonus = () as list => List.RemoveNulls(List.Transform(
                Mine(Get("bonus-issue", "HHKDB669101C0", [CTS = "", F_DT = F(Date.AddDays(WStart, -35)), T_DT = F(WEnd), SHT_CD = code])),
                (r) =>
                    let
                        lst = DateOf(r, "list_date"), rec = DateOf(r, "record_date"), rl = DateOf(r, "right_dt"),
                        d = if lst <> null then lst else rec,
                        qty = Pos(r, "issue_stk_qty"), rate = Pos(r, "fix_rate"),
                        rateTxt = if rate = null then null else "배정 " & Dec(rate) & "% (1주당 " & Dec(rate / 100) & "주)",
                        qtyTxt = if qty = null then null else "신주 " & N0(qty) & "주",
                        detail = if lst <> null then Join({"신주 상장", qtyTxt, rateTxt, if rec = null then null else "기준일 " & MD(rec)})
                                 else Join({"배정 기준일", rateTxt, qtyTxt, if rl = null then null else "권리락 " & MD(rl)})
                    in
                        if d = null then null else {d, "무상증자", Cut(detail), 1})),
            Merger = () as list => List.RemoveNulls(List.Transform(
                Get("merger-split", "HHKDB669104C0", [CTS = "", F_DT = F(WStart), T_DT = F(WEnd), SHT_CD = ""]),
                (r) =>
                    let
                        d = DateOf(r, "record_date"),
                        typ = Blank(Txt(r, "merge_type")), opp = Txt(r, "opp_cust_nm"), cust = Txt(r, "cust_nm"),
                        ratio = Pos(r, "merge_rate"),
                        stop1 = DateOf(r, "td_stop_dt"), stop2 = Date2Of(r, "td_stop_dt"),
                        stopTxt = if stop1 = null then null else "거래정지 " & MD(stop1) & "~" & (if stop2 = null then "" else MD(stop2)),
                        lst = DateOf(r, "list_dt"),
                        who = if opp <> "" and cust <> "" and opp <> cust then opp & " → " & cust else if opp <> "" then opp else cust,
                        detail = Cut(Join({(if typ = null then "합병·분할" else typ) & " 기일", if who = "" then null else who,
                                           if ratio = null then null else "비율 " & Dec(ratio), stopTxt,
                                           if lst = null then null else "신주 상장 " & MD(lst)})),
                        c2raw = Txt(r, "cust_cd"),
                        c2 = if Text.Length(c2raw) = 5 then Text.Upper(c2raw) & "0" else if Text.Length(c2raw) = 6 then Text.Upper(c2raw) else null,
                        hit = CodeOf(r) = code or c2 = code
                    in
                        if d = null or not hit then null else {d, "합병·분할", detail, 1})),
            Dividend = () as list => List.RemoveNulls(List.Transform(
                Mine(Get("dividend", "HHKDB669102C0", [CTS = "", GB1 = "0", F_DT = F(WStart), T_DT = F(WEnd), SHT_CD = code, HIGH_GB = ""])),
                (r) =>
                    let
                        d = DateOf(r, "record_date"), kind = Txt(r, "divi_kind"),
                        amt = Pos(r, "per_sto_divi_amt"), rate = Pos(r, "divi_rate"), stk = Pos(r, "stk_divi_rate"), face = Pos(r, "face_val"),
                        pay = DateOf(r, "divi_pay_dt"), stkPay = DateOf(r, "stk_div_pay_dt"),
                        label = if kind = "" then "분배금" else kind & " 현금배당",
                        cash = if amt = null then label & " 금액 미정" else label & " 주당 " & N0(amt) & "원",
                        detail = Join({if amt = null and stk <> null then null else cash,
                                       if amt = null or rate = null or face = null then null else "액면 " & Dec(rate) & "%",
                                       if stk = null then null else "주식배당 " & Dec(stk) & "%",
                                       if pay = null then null else "지급 " & MD(pay),
                                       if stkPay = null then null else "주식 지급 " & MD(stkPay),
                                       if Txt(r, "high_divi_gb") = "Y" then "고배당" else null})
                    in
                        if d = null then null else {d, "배당 기준일", Cut(detail), 1})),
            Meeting = () as list =>
                let
                    Rs = Mine(Get("sharehld-meet", "HHKDB669111C0", [CTS = "", F_DT = F(Date.AddDays(WStart, -100)), T_DT = F(WEnd), SHT_CD = code])),
                    Items = List.RemoveNulls(List.Transform(Rs, (r) =>
                        let d = DateOf(r, "gen_meet_dt") in if d = null then null else {d, Txt(r, "gen_meet_type"), Txt(r, "agenda")})),
                    T0 = #table(type table [d = date, t = text, a = text], Items),
                    G = Table.Group(T0, {"d", "t"}, {{"agendas", each List.Select(List.Distinct([a]), each _ <> ""), type list}}),
                    AgendaTxt = (l as list) as nullable text =>
                        if List.Count(l) = 0 then null
                        else "안건: " & Text.Combine(List.FirstN(l, 2), ", ") & (if List.Count(l) > 2 then " 외 " & Text.From(List.Count(l) - 2) & "건" else "")
                in
                    List.Transform(Table.ToRecords(G), (g) => {g[d], "주주총회", Cut(Join({if g[t] = "" then "총회" else g[t], AgendaTxt(g[agendas])})), 1}),

            Sources = {[이름 = "상장정보일정", f = ListInfo], [이름 = "유상증자", f = PaidIn], [이름 = "무상증자", f = Bonus],
                       [이름 = "합병·분할", f = Merger], [이름 = "배당 기준일", f = Dividend], [이름 = "주주총회", f = Meeting]},
            // 원천마다 한 번만 호출되도록 결과 목록을 버퍼로 고정(목록을 만드는 동안 난 호출 오류는 그 원천의 실패로 잡힘.
            // 행 안의 값 오류는 fnARun이 표 전체를 평가하며 잡는다)
            Results = List.Buffer(List.Transform(Sources, (s) => [이름 = s[이름], r = try List.Buffer(s[f]())])),
            Failed = List.Select(Results, each [r][HasError]),
            Reason = (e as record) as text => let m = try Text.Trim(Text.From(e[Message])) otherwise "" in if m = "" then "조회 실패" else m,
            Raw = List.Select(List.Combine(List.Transform(List.Select(Results, each not [r][HasError]), each [r][Value])),
                        each _{0} <> null and _{0} >= WStart and _{0} <= WEnd),
            T1 = #table(type table [일자 = date, 종류 = text, 상세 = text, 우선 = number], Raw),
            // 같은 일자의 유상증자·무상증자는 전용 API 행(우선 1)이 있으면 상장정보일정 행(우선 2)을 뺌
            Keyed = Table.AddColumn(T1, "키", each [종류] & "|" & fnYmdText([일자]), type text),
            BestT = Table.Group(Keyed, {"키"}, {{"최소", each List.Min([우선]), type number}}),
            BestRec = Record.FromList(BestT[최소], BestT[키]),
            Kept = Table.SelectRows(Keyed, each not List.Contains({"유상증자", "무상증자"}, [종류]) or [우선] = Record.Field(BestRec, [키])),
            Ev = Table.Sort(Table.Distinct(Table.SelectColumns(Kept, {"일자", "종류", "상세"})),
                        {{"일자", Order.Ascending}, {"종류", Order.Ascending}, {"상세", Order.Ascending}}),
            Cols = List.Transform(Types, each _{0}),
            Rows = List.Transform(Table.ToRecords(Ev), (e) =>
                        {code, e[일자], Duration.Days(e[일자] - Today), e[종류], if e[상세] = "" then null else e[상세], "OK", now}),
            ErrRows = List.Transform(Failed, (x) =>
                        {code, null, null, x[이름], null, "오류: " & x[이름] & " 조회 실패(" & Text.Start(Reason(x[r][Error]), 80) & ")", now})
        in
            if List.Count(Failed) = List.Count(Sources) then error Error.Record("조회 실패", Reason(Failed{0}[r][Error]))
            else if List.IsEmpty(Rows) and List.IsEmpty(ErrRows) then fnAOneRow(Types, code, "데이터 없음", now)
            else #table(Cols, Rows & ErrRows)
in
    fnARun("tblA_Events", Types, Main)
