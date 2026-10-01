// T_Events : 기업 이벤트 캘린더 — 예탁원정보 6종을 한 표로 (spec R11, 2026-10-01 T11, [전체] 버튼에서만 갱신)
//  - 대상 종목 = fnPageRows()(대회 종목 ∪ 보유·관심, 수정표 반영). 기간 = 오늘 − event_days_back(7) ~ 오늘 + event_days_ahead(60).
//  - 실행 모드(tblRunCtl mode): full([전체])만 KIS를 호출한다. 그 밖(quick·sector·analysis·build·빈칸·알 수 없는 값)은 호출 없이
//    자기 표(tblEvents) → 없거나 비었으면 정적 tblEventsSeed → 없으면 열만 있는 빈 표를 돌려준다(D-day만 오늘 기준으로 다시 계산).
//  - 열: 일자, D-day(일자 − 오늘, 지난 이벤트는 음수), 종목코드, 종목명, 종류, 상세, 보유(Y/빈칸), 상태, 조회시각.
//    정렬 = 일자 ↑, 종목코드 ↑. 같은 종목코드·일자·종류·상세는 1건. 종목명 = 대회 명단(tblContest) → 종목DB(tblUniverse) → 응답 이름.
//    보유 = fnPageRows 구분이 보유 또는 보유·관심일 때 Y. 값이 없으면 빈칸(0 아님).
//
//  원천(예탁원정보 ksdinfo) — 6개 모두 시장 전체 목록을 주므로 받은 뒤 페이지 행 종목만 남긴다. 응답 필드 이름은 2026-10-01 실측.
//    종류                일자(원천 필드)                                          API (tr_id) · 기간 조회 기준
//    CB·BW 전환 상장     상장일 list_dt                                           상장정보일정 list-info (HHKDB669107C0) · 상장일. 사유 issue_type에 CB·BW·'전환'
//    신규상장            상장일 list_dt                                           상장정보일정 · 사유 '통일교체'(상장 전 주식을 통일규격 증권으로 바꾸는 것 = 신규상장 표지,
//                                                                                 같은 종목·일자의 '유상증자' 행은 공모 주식수·공모가로 상세에 넣고 따로 안 냄)
//    유상증자            신주 상장일 list_date, 없으면 청약 시작일 sub_term_ft      유상증자일정 paidin-capin (HHKDB669100C0, GB1=1 청약일별) · 청약 시작일(시작일 − 45일부터)
//                        (+ 상장정보일정의 사유 '유상증자' 신주 상장일 — 같은 종목·일자는 유상증자일정 행으로 합침)
//    무상증자            신주 상장일 list_date, 없으면 배정 기준일 record_date      무상증자일정 bonus-issue (HHKDB669101C0) · 기준일(시작일 − 35일부터)
//                        (+ 상장정보일정의 사유 '무상증자' 신주 상장일 — 같은 종목·일자는 무상증자일정 행으로 합침)
//    합병·분할           합병·분할 기일 record_date(KSD '기준일' = 기일: 거래정지는 기일 전 영업일에   합병·분할일정 merger-split (HHKDB669104C0) · 기준일. 피합병(분할)회사 sht_cd와
//                        시작하고 신주 상장은 2~3주 뒤)                             합병(분할)회사 cust_cd+'0'(보통주) 두 종목 모두. 주식교환 행도 포함
//                        (+ 상장정보일정의 사유 합병·회사분할 신주 상장일 = 별도 이벤트)
//    배당 기준일         배당 기준일 record_date                                   배당일정 dividend (HHKDB669102C0, GB1=0 전체) · 기준일. ETF 분배금(배당종류 빈칸) 포함
//    주주총회            총회일 gen_meet_dt                                       주주총회일정 sharehld-meet (HHKDB669111C0) · 기준일(record_date)
//                        총회일이 기준일보다 중앙값 29일·90% 52일·최대 약 90일 뒤라 기준일은 시작일 − 45일부터 받아 총회일로 걸러 냄. 안건 행(1행 = 안건 1개)은 총회 1건으로 묶음.
//  상장정보일정에서 버리는 사유: STOCKOPTION행사(임직원 주식매수선택권), 액면분할·액면병합, 상호변경, 자본감소, '-'(사유 불명 — 주식교환 상장 등),
//    배당(주식배당 상장), 그 밖 7종에 해당하지 않는 사유. 주주총회 총회일이 9999/12/31(안건 미정)이면 일자가 없어 버림.
//    '주식전환'(전환주식·전환우선주의 보통주 전환 상장)은 'CB·BW 전환 상장'에 넣음(CB·BW와 같은 전환 상장).
//
//  연속 조회(실측): 6개 API 모두 응답이 한 쪽에 100행(주주총회는 20행)에서 잘리고(상장정보일정은 모자라는 행을 빈 행으로 채워 항상 100행),
//    tr_cont=N만 보내면 같은 쪽을 되풀이한다(그래서 '기간을 하루씩 쪼갰다'는 보고가 있었음). 요청 헤더 tr_cont=N + CTS = 앞 쪽 마지막 행의
//    (정렬 기준 일자 + 종목코드)를 보내면 이어서 준다 — 정렬 기준 일자: 유상증자일정 sub_term_ft, 상장정보일정 list_dt, 나머지 record_date, 모두 최신순(내림차순).
//    CTS가 같은 (일자, 종목)의 나머지 행을 건너뛰므로 쪽 끝의 같은 (일자, 종목) 행 묶음은 다음 쪽으로 미뤄 다시 받는다. 주주총회 API는 CTS의
//    종목코드에 영문이 있으면 CTS를 무시하고 첫 쪽을 되풀이해서(실측) 영문 앞까지만 쓰고 9로 채운 숫자 코드로 바꿔 보낸다(NumCode).
//    같은 쪽이 되풀이되면(응답이 같은 쪽) 연속 조회가 안 되는 것이므로 중단한다. 종목코드(SHT_CD)를 지정한 조회는 CTS를 무시하는 한 쪽짜리 조회.
//    응답 헤더 tr_cont는 파워 쿼리가 못 읽으므로 '쪽이 쪽 크기의 75% 이상 찼는지'로 다음 쪽을 판단한다(꽉 찬 마지막 쪽 뒤에는 빈 쪽 1회).
//    쪽 수 상한(상장정보·증자·합병분할 15, 배당·총회 60)에 닿거나, 총회의 한 기준일이 6쪽(약 120행) 넘게 이어지면(12/31 정기총회 등)
//    그 날짜 이전으로 새로 조회를 시작하고, 이런 경우와 연속 조회 중단은 '오류: 주주총회 … 일부 누락 가능' 행으로 남긴다.
//  주주총회는 보유·관심 종목(보유 우선, 최대 60)만 종목코드로 따로 한 번 더 조회한다(기준일 100일 앞부터, 최근 20행까지 한 쪽) — 시장 전체 조회의 하한(45일)과
//    쪽 상한을 넘는 정기총회도 내 종목은 빠지지 않게. 같은 종목·일자는 이 개별 조회 결과를 쓴다.
//
//  실패 처리: 6개 API마다 try로 끝까지 평가. 모두 실패하면 오류로 올려 직전 표 유지(fnFallback, 상태 '이전 데이터(갱신 실패: 사유)').
//    일부만 실패하면 나머지 행은 OK로 남기고 실패한 API마다 1행(상태 "오류: <API> <사유>", 종목코드·일자 빈칸)을 덧붙인다.
//    대상 종목이 하나도 없으면 호출 없이 빈 표(0행).
//  호출 수(2026-10-01 실측, 대회 527종목 + 샘플 보유·관심 18종목): 상장정보 1 + 유상증자 1 + 무상증자 1 + 합병·분할 1 + 배당 3 + 주주총회 50 + 보유·관심 개별 18 = 75회,
//    소요 20~40초(쪽당 0.3~0.5초). 주주총회가 대부분 — 상장정보는 기간 행이 75개를 넘으면 2쪽이 됨.
//  점검: Online의 Diag 단계가 API별 쪽 수·원천 행·이벤트 수·비고를 표로 만든다(출력에는 안 나옴).
let
    Cols = {"일자", "D-day", "종목코드", "종목명", "종류", "상세", "보유", "상태", "조회시각"},
    Types = {{"일자", type date}, {"D-day", Int64.Type}, {"종목코드", type text}, {"종목명", type text}, {"종류", type text},
             {"상세", type text}, {"보유", type text}, {"상태", type text}, {"조회시각", type datetime}},
    EvType = type table [일자 = nullable date, #"D-day" = nullable Int64.Type, 종목코드 = nullable text, 종목명 = nullable text,
                         종류 = nullable text, 상세 = nullable text, 보유 = nullable text, 상태 = text, 조회시각 = datetime],
    Empty = #table(EvType, {}),
    Now = DateTime.LocalNow(),
    Today = Date.From(Now),

    // ---- 공통 도우미
    Txt = (r as record, k as text) as text =>
        let v = Record.FieldOrDefault(r, k, null) in if v = null then "" else try Text.Trim(Text.From(v)) otherwise "",
    Blank = (t as nullable text) as nullable text => if t = null or Text.Trim(t) = "" then null else Text.Trim(t),
    CodeOf = (v as any) as nullable text =>
        let t = if v = null then "" else try Text.Upper(Text.Trim(Text.From(v))) otherwise "" in if t = "" then null else t,
    // 날짜 필드: "20261125", "2026/11/25", "2026/11/25 ~ 2026/11/26"의 첫 날짜. 잘못된 값·9999년(미정)은 null
    DateOf = (r as record, k as text) as nullable date =>
        let d = fnYmd(Text.Start(Text.Select(Txt(r, k), {"0".."9"}), 8))
        in if d = null or Date.Year(d) >= 9999 then null else d,
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
    // {코드, 값} 쌍 → 코드별 값 레코드(같은 코드는 앞의 것)
    MapOf = (pairs as list) as record =>
        let d = List.Distinct(pairs, each _{0})
        in Record.FromList(List.Transform(d, each _{1}), List.Transform(d, each _{0})),

    // ---- 설정·실행 모드
    S = try fnSettings() otherwise [],
    IntSetting = (k as text, dflt as number) as number =>
        let v = Record.FieldOrDefault(S, k, null), n = if v = null then null else try Number.From(v) otherwise null
        in if n = null or n < 0 then dflt else Number.RoundDown(n),
    Back = IntSetting("event_days_back", 7),
    Ahead = IntSetting("event_days_ahead", 60),
    RunCtl = try Excel.CurrentWorkbook(){[Name = "tblRunCtl"]}[Content] otherwise null,
    ModeRaw =
        if RunCtl = null or not List.ContainsAll(Table.ColumnNames(RunCtl), {"키", "값"}) then null
        else
            let m = Table.SelectRows(RunCtl, each [키] <> null and (try Text.Lower(Text.Trim(Text.From([키]))) = "mode" otherwise false))
            in if Table.RowCount(m) = 0 then null else m{0}[값],
    Mode =
        let t = if ModeRaw = null then "" else try Text.Lower(Text.Trim(Text.From(ModeRaw))) otherwise ""
        in if List.Contains({"quick", "full", "sector", "analysis"}, t) then t else "build",

    // ---- 호출 없는 경로(full 외): 직전 표 → 시드 → 빈 표. D-day만 오늘 기준으로 다시 계산
    Offline =
        let
            Pick = (t as nullable table) as nullable table =>
                if t = null then null
                else
                    let
                        u = Table.SelectColumns(t, Cols, MissingField.UseNull),
                        v = Table.SelectRows(u, each (try [일자] <> null or [종목코드] <> null or [상태] <> null otherwise true))
                    in
                        if Table.RowCount(v) = 0 then null else v,
            Own = Pick(try Excel.CurrentWorkbook(){[Name = "tblEvents"]}[Content] otherwise null),
            Seed = Pick(try Excel.CurrentWorkbook(){[Name = "tblEventsSeed"]}[Content] otherwise null),
            Base = if Own <> null then Own else if Seed <> null then Seed else Empty,
            Typed = Table.TransformColumnTypes(Base, Types),
            Re = Table.ReorderColumns(
                    Table.AddColumn(Table.RemoveColumns(Typed, {"D-day"}), "D-day",
                        each if [일자] = null then null else Duration.Days([일자] - Today), Int64.Type), Cols)
        in
            Re,

    // ---- 호출 경로(full)
    Online =
        let
            WStart = Date.AddDays(Today, -Back),
            WEnd = Date.AddDays(Today, Ahead),
            PR = Table.Buffer(fnPageRows()),
            PageCodes = List.Buffer(List.Distinct(List.Select(List.Transform(PR[종목코드], CodeOf), each _ <> null))),
            PageRec = Record.FromList(List.Repeat({true}, List.Count(PageCodes)), PageCodes),
            GubunRec = MapOf(List.Select(List.Transform(Table.ToRecords(PR), (r) => {CodeOf(r[종목코드]), Blank(r[구분])}), each _{0} <> null)),
            IsPage = (c as nullable text) as logical => c <> null and Record.HasFields(PageRec, c),
            Held = (c as text) as nullable text =>
                let g = Record.FieldOrDefault(GubunRec, c, null) in if g <> null and Text.StartsWith(g, "보유") then "Y" else null,
            // 종목명: 대회 명단 → 종목DB → 응답 이름
            NamesFrom = (tname as text) as list =>
                let t = try Excel.CurrentWorkbook(){[Name = tname]}[Content] otherwise null
                in if t = null or not List.ContainsAll(Table.ColumnNames(t), {"종목코드", "종목명"}) then {}
                   else List.RemoveNulls(List.Transform(Table.ToRecords(Table.SelectColumns(t, {"종목코드", "종목명"})), (r) =>
                        let c = CodeOf(r[종목코드]), n = Txt(r, "종목명") in if c = null or n = "" then null else {c, n})),
            NameRec = MapOf(NamesFrom("tblContest") & NamesFrom("tblUniverse")),

            // ---- 원천 행 → 이벤트 행(목록). 열 순서 = EvSchema: 일자, 종목코드, 응답이름, 종류, 상세, 우선
            //      우선: 0 = 보유·관심 종목 개별 조회(주주총회), 1 = 전용 API, 2 = 상장정보일정 — 같은 종목·일자의 증자·총회는 우선이 낮은 쪽만 남김
            EvSchema = type table [일자 = date, 종목코드 = text, 응답이름 = nullable text, 종류 = text, 상세 = nullable text, 우선 = number],
            Code = (r as record) as nullable text => CodeOf(Record.FieldOrDefault(r, "sht_cd", null)),
            Build = (rows as list) as table => Table.Buffer(#table(EvSchema, List.RemoveNulls(rows))),

            // 상장정보일정: 상장일 = list_dt
            LiBuild = (rows as list) as table =>
                let
                    Rs = List.Buffer(List.Select(rows, each IsPage(Code(_)))),
                    Typ = (r as record) as text => Txt(r, "issue_type"),
                    KeyOf = (r as record) as text => Code(r) & "|" & Txt(r, "list_dt"),
                    IpoKeys = List.Buffer(List.Distinct(List.Transform(List.Select(Rs, each Typ(_) = "통일교체"), KeyOf))),
                    Ev = (r as record) as nullable list =>
                        let
                            t = Typ(r), up = Text.Upper(t),
                            d = DateOf(r, "list_dt"), c = Code(r),
                            qty = Pos(r, "issue_stk_qty"), px = Pos(r, "issue_price"), tot = Pos(r, "tot_issue_stk_qty"),
                            qtyTxt = if qty = null then null else N0(qty) & "주" & (if px = null then "" else " @ " & N0(px) & "원"),
                            totTxt = if tot = null then null else "상장 후 총 " & N0(tot) & "주",
                            kind =
                                if t = "통일교체" then "신규상장"
                                else if Text.Contains(up, "CB") or Text.Contains(up, "BW") or Text.Contains(t, "전환") then "CB·BW 전환 상장"
                                else if t = "유상증자" then (if List.Contains(IpoKeys, KeyOf(r)) then null else "유상증자")
                                else if t = "무상증자" then "무상증자"
                                else if Text.Contains(t, "합병") or (Text.EndsWith(t, "분할") and not Text.Contains(t, "액면")) then "합병·분할"
                                else null,
                            // 신규상장: 같은 종목·일자의 '유상증자' 행 = 공모
                            Pub = List.Select(Rs, each KeyOf(_) = KeyOf(r) and Typ(_) = "유상증자"),
                            pubQty = if List.Count(Pub) = 0 then null else Pos(Pub{0}, "issue_stk_qty"),
                            pubPx = if List.Count(Pub) = 0 then null else Pos(Pub{0}, "issue_price"),
                            pubTxt = if pubQty = null then null else "공모 " & N0(pubQty) & "주" & (if pubPx = null then "" else " @ " & N0(pubPx) & "원"),
                            detail =
                                if kind = "신규상장" then Join({"신규상장", pubTxt, if tot = null then null else "상장주식 " & N0(tot) & "주"})
                                else if kind = "CB·BW 전환 상장" then Join({t, qtyTxt, totTxt})
                                else if kind = "유상증자" then Join({"신주 상장", qtyTxt, totTxt})
                                else if kind = "무상증자" then Join({"신주 상장", if qty = null then null else N0(qty) & "주", totTxt})
                                else Join({t & " 신주 상장", if qty = null then null else N0(qty) & "주", totTxt})
                        in
                            if kind = null or d = null or c = null then null else {d, c, Blank(Txt(r, "isin_name")), kind, Cut(detail), 2}
                in
                    Build(List.Transform(Rs, Ev)),

            // 유상증자일정(청약일별): 일자 = 신주 상장일 list_date, 없으면 청약 시작일 sub_term_ft
            PaBuild = (rows as list) as table =>
                let
                    Ev = (r as record) as nullable list =>
                        let
                            c = Code(r),
                            lst = DateOf(r, "list_date"), s1 = DateOf(r, "sub_term_ft"), s2 = Date2Of(r, "sub_term"),
                            s1b = if s1 = null then DateOf(r, "sub_term") else s1,
                            sub = if s1b = null then null else "청약 " & MD(s1b) & (if s2 <> null and s2 <> s1b then "~" & MD(s2) else ""),
                            d = if lst <> null then lst else s1b,
                            qty = Pos(r, "issue_stk_qty"), rate = Pos(r, "fix_rate"), px = Pos(r, "fix_price"), disc = Pos(r, "disc_rate"),
                            rl = DateOf(r, "right_dt"), rec = DateOf(r, "record_date"),
                            common = {if qty = null then null else "신주 " & N0(qty) & "주", if rate = null then null else "배정 " & Dec(rate) & "%",
                                      if px = null then null else "발행가 " & N0(px) & "원", if disc = null then null else "할인율 " & Dec(disc) & "%"},
                            detail =
                                if lst <> null then Join({"신주 상장"} & common & {sub})
                                else Join({sub} & common & {if rl = null then null else "권리락 " & MD(rl), if rec = null then null else "기준일 " & MD(rec)})
                        in
                            if d = null or c = null or not IsPage(c) then null else {d, c, Blank(Txt(r, "isin_name")), "유상증자", Cut(detail), 1}
                in
                    Build(List.Transform(rows, Ev)),

            // 무상증자일정: 일자 = 신주 상장일 list_date, 없으면 배정 기준일 record_date
            BoBuild = (rows as list) as table =>
                let
                    Ev = (r as record) as nullable list =>
                        let
                            c = Code(r),
                            lst = DateOf(r, "list_date"), rec = DateOf(r, "record_date"), rl = DateOf(r, "right_dt"),
                            d = if lst <> null then lst else rec,
                            qty = Pos(r, "issue_stk_qty"), rate = Pos(r, "fix_rate"),
                            rateTxt = if rate = null then null else "배정 " & Dec(rate) & "% (1주당 " & Dec(rate / 100) & "주)",
                            qtyTxt = if qty = null then null else "신주 " & N0(qty) & "주",
                            detail =
                                if lst <> null then Join({"신주 상장", qtyTxt, rateTxt, if rec = null then null else "기준일 " & MD(rec)})
                                else Join({"배정 기준일", rateTxt, qtyTxt, if rl = null then null else "권리락 " & MD(rl)})
                        in
                            if d = null or c = null or not IsPage(c) then null else {d, c, Blank(Txt(r, "isin_name")), "무상증자", Cut(detail), 1}
                in
                    Build(List.Transform(rows, Ev)),

            // 합병·분할일정: 일자 = 기일 record_date. 피합병(분할)회사 sht_cd와 합병(분할)회사 cust_cd + '0'(보통주) 두 종목 모두
            MeBuild = (rows as list) as table =>
                let
                    Ev = (r as record) as list =>
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
                            c1 = Code(r),
                            c2raw = Txt(r, "cust_cd"),
                            c2 = if Text.Length(c2raw) = 5 then Text.Upper(c2raw) & "0" else if Text.Length(c2raw) = 6 then Text.Upper(c2raw) else null,
                            row = (c as nullable text, nm as text) as nullable list =>
                                if d = null or c = null or not IsPage(c) then null else {d, c, Blank(nm), "합병·분할", detail, 1}
                        in
                            {row(c1, opp), if c2 = c1 then null else row(c2, cust)}
                in
                    Build(List.Combine(List.Transform(rows, Ev))),

            // 배당일정: 일자 = 배당 기준일 record_date (ETF는 분배금 — 배당종류가 빈칸)
            DiBuild = (rows as list) as table =>
                let
                    Ev = (r as record) as nullable list =>
                        let
                            c = Code(r), d = DateOf(r, "record_date"),
                            kind = Txt(r, "divi_kind"),
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
                            if d = null or c = null or not IsPage(c) then null else {d, c, Blank(Txt(r, "isin_name")), "배당 기준일", Cut(detail), 1}
                in
                    Build(List.Transform(rows, Ev)),

            // 주주총회일정: 일자 = 총회일 gen_meet_dt. 안건 행은 (종목, 총회일, 구분)마다 1건으로 묶음. pri = 우선(0 = 종목별 조회)
            MtBuild = (pri as number) as function => (rows as list) as table =>
                let
                    Rs = List.Select(rows, each IsPage(Code(_))),
                    Items = List.RemoveNulls(List.Transform(Rs, (r) =>
                        let d = DateOf(r, "gen_meet_dt") in
                            if d = null then null else {Code(r), d, Txt(r, "gen_meet_type"), Txt(r, "agenda"), Txt(r, "isin_name")})),
                    T0 = #table(type table [c = text, d = date, t = text, a = text, n = text], Items),
                    G = Table.Group(T0, {"c", "d", "t"}, {{"agendas", each List.Select(List.Distinct([a]), each _ <> ""), type list},
                                                         {"n", each List.First([n]), type text}}),
                    AgendaTxt = (l as list) as nullable text =>
                        if List.Count(l) = 0 then null
                        else "안건: " & Text.Combine(List.FirstN(l, 2), ", ") & (if List.Count(l) > 2 then " 외 " & Text.From(List.Count(l) - 2) & "건" else ""),
                    Ev = (g as record) as list =>
                        {g[d], g[c], Blank(g[n]), "주주총회", Cut(Join({if g[t] = "" then "총회" else g[t], AgendaTxt(g[agendas])})), pri}
                in
                    Build(List.Transform(Table.ToRecords(G), Ev)),

            // ---- 연속 조회: 한 API의 모든 쪽을 모아 [rows, pages, notes]로
            //      쪽을 순서대로 받으므로 호출은 초당 3건 안팎(쪽당 0.3~0.5초, 실측) — KIS 한도(앱키당 초당 약 20건)에 못 미쳐 따로 늦추지 않음
            Page = (d as record, q as record, cts as text) as list =>
                let
                    J = fnKisGetPage(d[path], d[trId], Record.AddField(q, "CTS", cts), if cts = "" then "" else "N"),
                    O = Record.FieldOrDefault(J, "output1", {}),
                    L = if O is list then O else if O is record then {O} else {}
                in
                    List.Buffer(List.Select(L, each _ is record and Txt(_, "sht_cd") <> "")),
            // CTS의 종목코드 자리에 영문이 든 코드(0190K0 등)를 보내면 주주총회 API가 CTS를 무시하고 첫 쪽을 되풀이한다(실측).
            // 첫 영문 앞까지만 쓰고 나머지를 9로 채운 숫자 코드로 바꿈 — 코드는 문자열 순서(숫자 < 영문)로 비교되므로 원래 코드보다 작아
            // 그 사이 행이 몇 개 다시 오더라도 빠지는 행은 없다(중복 행은 이벤트를 만들 때 합쳐짐).
            NumCode = (c as text) as text =>
                let p = List.First(List.Select(List.Positions(Text.ToList(c)), each not List.Contains({"0".."9"}, Text.At(c, _))), null)
                in if p = null then c else Text.Start(c, p) & Text.Repeat("9", Text.Length(c) - p),
            Chain = (d as record, q as record) as record =>
                let
                    KeyOf = (r as record) as text => Txt(r, d[dateField]) & Txt(r, "sht_cd"),
                    CursorOf = (r as record) as text => Txt(r, d[dateField]) & NumCode(Txt(r, "sht_cd")),
                    Loop = (q2 as record, cts as text, n as number, acc as list, prevDay as text, run as number, notes as list, seen as list) as record =>
                        let
                            rows = Page(d, q2, cts),
                            cnt = List.Count(rows),
                            more = cnt >= d[moreAt],
                            keys = List.Transform(rows, KeyOf),
                            sig = if cnt = 0 then "" else List.First(keys) & "|" & List.Last(keys) & "|" & Text.From(cnt),
                            repeated = cnt > 0 and List.Contains(seen, sig),
                            lastKey = if cnt = 0 then "" else List.Last(keys),
                            diff = List.Select(List.Positions(keys), each keys{_} <> lastKey),
                            keepN = if List.Count(diff) = 0 then cnt else List.Max(diff) + 1,
                            day0 = if cnt = 0 then "" else Txt(rows{0}, d[dateField]),
                            day1 = if cnt = 0 then "" else Txt(List.Last(rows), d[dateField]),
                            sameDay = day0 <> "" and day0 = day1,
                            run2 = if sameDay and day0 = prevDay then run + 1 else if sameDay then 1 else 0,
                            capped = n >= d[maxPages],
                            flooded = d[flood] > 0 and sameDay and run2 >= d[flood],
                            // 이어서 받을 때만 쪽 끝의 같은 (일자, 종목) 행 묶음을 다음 쪽으로 미룸(CTS가 그 종목의 나머지 행을 건너뛰므로)
                            Continue = more and not repeated and not capped and not flooded,
                            acc2 = acc & (if Continue then List.FirstN(rows, keepN) else rows),
                            nextCts = if cnt = 0 then "" else CursorOf(if List.Count(diff) = 0 then List.Last(rows) else rows{keepN - 1})
                        in
                            if not more then [rows = acc2, pages = n, notes = notes]
                            else if repeated then [rows = acc2, pages = n, notes = notes & {"연속 조회가 진행되지 않아 중단(일부 누락 가능)"}]
                            else if capped then [rows = acc2, pages = n, notes = notes & {"응답 " & Text.From(n) & "쪽에서 중단(일부 누락 가능)"}]
                            else if flooded then
                                @Loop(Record.TransformFields(q2, {{"T_DT", each fnYmdText(Date.AddDays(fnYmd(day0), -1))}}), "", n + 1, acc2, "", 0,
                                      notes & {"기준일 " & day0 & "은 행이 너무 많아 일부만 조회"}, {})
                            else @Loop(q2, nextCts, n + 1, acc2, if sameDay then day0 else "", run2, notes, seen & {sig})
                in
                    Loop(q, "", 1, {}, "", 0, {}, {}),

            // ---- API 정의
            F = (d as date) as text => fnYmdText(d),
            BasePath = "/uapi/domestic-stock/v1/ksdinfo/",
            Specs = {
                [key = "상장정보", path = BasePath & "list-info", trId = "HHKDB669107C0", dateField = "list_dt", moreAt = 75, maxPages = 15, flood = 0,
                 q = [SHT_CD = "", F_DT = F(WStart), T_DT = F(WEnd)], build = LiBuild],
                [key = "유상증자", path = BasePath & "paidin-capin", trId = "HHKDB669100C0", dateField = "sub_term_ft", moreAt = 75, maxPages = 15, flood = 0,
                 q = [GB1 = "1", F_DT = F(Date.AddDays(WStart, -45)), T_DT = F(WEnd), SHT_CD = ""], build = PaBuild],
                [key = "무상증자", path = BasePath & "bonus-issue", trId = "HHKDB669101C0", dateField = "record_date", moreAt = 75, maxPages = 15, flood = 0,
                 q = [F_DT = F(Date.AddDays(WStart, -35)), T_DT = F(WEnd), SHT_CD = ""], build = BoBuild],
                [key = "합병·분할", path = BasePath & "merger-split", trId = "HHKDB669104C0", dateField = "record_date", moreAt = 75, maxPages = 15, flood = 0,
                 q = [F_DT = F(WStart), T_DT = F(WEnd), SHT_CD = ""], build = MeBuild],
                [key = "배당", path = BasePath & "dividend", trId = "HHKDB669102C0", dateField = "record_date", moreAt = 75, maxPages = 60, flood = 0,
                 q = [GB1 = "0", F_DT = F(WStart), T_DT = F(WEnd), SHT_CD = "", HIGH_GB = ""], build = DiBuild],
                [key = "주주총회", path = BasePath & "sharehld-meet", trId = "HHKDB669111C0", dateField = "record_date", moreAt = 15, maxPages = 60, flood = 6,
                 q = [F_DT = F(Date.AddDays(WStart, -45)), T_DT = F(WEnd), SHT_CD = ""], build = MtBuild(1)]
            },
            Reason = (e as record) as text =>
                let m = try Text.Trim(Text.From(e[Message])) otherwise "" in if m = "" then "조회 실패" else Text.Start(m, 100),
            // try는 레코드 필드를 평가하지 않으므로 표 행 수·쪽 수를 읽어 이 안에서 끝까지 평가
            Run = (d as record, q as record, build as function) as nullable record =>
                let
                    c = Chain(d, q),
                    t = build(c[rows]),
                    ok = Table.RowCount(t) >= 0 and c[pages] >= 0
                in
                    if ok then [chain = c, events = t] else null,
            Results = List.Buffer(List.Transform(Specs, (d) => [spec = d, r = try Run(d, d[q], d[build])])),

            // 보유·관심 종목의 주주총회 개별 조회(보유 우선, 최대 60종목)
            HeldCodes =
                let
                    T = Table.AddColumn(Table.SelectRows(PR, each Blank([구분]) <> null), "p", each if Text.StartsWith([구분], "보유") then 0 else 1, Int64.Type),
                    Sorted = Table.Sort(T, {{"p", Order.Ascending}})
                in
                    List.FirstN(List.Distinct(List.Select(List.Transform(Sorted[종목코드], CodeOf), each _ <> null)), 60),
            MeetSpec = List.First(List.Select(Specs, each [key] = "주주총회")),
            Empty0 = #table(EvSchema, {}),
            PerStock = List.Buffer(List.Transform(HeldCodes, (c) =>
                [code = c, r = try Run(Record.TransformFields(MeetSpec, {{"maxPages", each 1}, {"flood", each 0}}),
                                       [F_DT = F(Date.AddDays(WStart, -100)), T_DT = F(WEnd), SHT_CD = c], MtBuild(0))])),

            // ---- 결과 합치기
            Failed = List.Select(Results, each [r][HasError]),
            Passed = List.Select(Results, each not [r][HasError]),
            PerFail = List.Select(PerStock, each [r][HasError]),
            AllFailed = List.Count(Failed) = List.Count(Results),
            Checked = if AllFailed then error Error.Record("이벤트 조회 실패", Reason(Failed{0}[r][Error]) & " (예탁원정보 " & Text.From(List.Count(Failed)) & "개 모두)") else Results,
            Tables = List.Transform(List.Select(Checked, each not [r][HasError]), each [r][Value][events])
                     & List.Transform(List.Select(PerStock, each not [r][HasError]), each [r][Value][events]),
            All = Table.Combine(Tables & {Empty0}),
            InWin = Table.SelectRows(All, each [일자] >= WStart and [일자] <= WEnd),
            // 같은 종목·일자의 유상증자·무상증자(전용 API ↔ 상장정보일정)·주주총회(개별 조회 ↔ 전체 조회)는 우선이 낮은 쪽만
            MergeKinds = {"유상증자", "무상증자", "주주총회"},
            Keyed = Table.AddColumn(InWin, "키", each [종류] & "|" & [종목코드] & "|" & fnYmdText([일자]), type text),
            Best = Table.Group(Keyed, {"키"}, {{"최소", each List.Min([우선]), type number}}),
            BestRec = Record.FromList(Best[최소], Best[키]),
            Pruned = Table.SelectRows(Keyed, each not List.Contains(MergeKinds, [종류]) or [우선] = Record.Field(BestRec, [키])),
            Dedup = Table.Distinct(Pruned, {"종목코드", "일자", "종류", "상세"}),
            Sorted = Table.Sort(Dedup, {{"일자", Order.Ascending}, {"종목코드", Order.Ascending}, {"종류", Order.Ascending}, {"상세", Order.Ascending}}),
            EventRows = List.Transform(Table.ToRecords(Sorted), (r) =>
                {r[일자], Duration.Days(r[일자] - Today), r[종목코드], Record.FieldOrDefault(NameRec, r[종목코드], r[응답이름]), r[종류], r[상세],
                 Held(r[종목코드]), "OK", Now}),
            // 실패·불완전한 API마다 1행(종목코드·일자 빈칸)
            ErrRow = (txt as text) as list => {null, null, null, null, null, null, null, "오류: " & txt, Now},
            ErrRows =
                List.Transform(Failed, each ErrRow([spec][key] & " " & Reason([r][Error])))
                & List.Transform(List.Select(Passed, each List.Count([r][Value][chain][notes]) > 0),
                                 each ErrRow([spec][key] & " " & Text.Combine([r][Value][chain][notes], "; ")))
                & (if List.Count(PerFail) = 0 or List.Contains(List.Transform(Failed, each [spec][key]), "주주총회") then {}
                   else {ErrRow("주주총회(보유·관심) " & Text.From(List.Count(PerFail)) & "종목 조회 실패: " & Reason(PerFail{0}[r][Error]))}),
            Events = #table(EvType, EventRows & ErrRows),
            // 점검용(표에는 나오지 않음): API마다 쪽 수·행 수·비고
            Diag = #table(type table [체인 = text, 쪽수 = number, 원천행 = number, 이벤트 = number, 비고 = text],
                List.Transform(Passed, each {[spec][key], [r][Value][chain][pages], List.Count([r][Value][chain][rows]), Table.RowCount([r][Value][events]),
                                             Text.Combine([r][Value][chain][notes], "; ")})
                & List.Transform(Failed, each {[spec][key], 0, 0, 0, "실패: " & Reason([r][Error])})
                & {{"주주총회(보유·관심)", List.Sum({0} & List.Transform(List.Select(PerStock, each not [r][HasError]), each [r][Value][chain][pages])),
                    List.Sum({0} & List.Transform(List.Select(PerStock, each not [r][HasError]), each List.Count([r][Value][chain][rows]))),
                    List.Sum({0} & List.Transform(List.Select(PerStock, each not [r][HasError]), each Table.RowCount([r][Value][events]))),
                    Text.From(List.Count(HeldCodes)) & "종목, 실패 " & Text.From(List.Count(PerFail))}})
        in
            if List.Count(PageCodes) = 0 then [Events = Empty, Diag = null] else [Events = Events, Diag = Diag],

    Main = if Mode <> "full" then Offline else Online[Events],
    // 실패(토큰·네트워크·호출제한·6개 API 모두 실패 등) 시 오류 대화상자 대신 직전 결과를 유지.
    // 직전 표를 Excel에서 다시 읽으면 열 형식이 없어(any) 날짜 서식이 일반으로 바뀌므로 같은 형식을 다시 지정함 — 값은 그대로
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblEvents", Cols, Tried[Error][Message]), Types) else Tried[Value]
in
    Final
