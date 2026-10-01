// T_PxStore : 가격 이력 저장소 — 페이지 행 종목의 일봉(수정주가), 변경분만 이어 붙임 (2026-10-01) → tblPxStore(숨김)
//  - 열: 종목코드(text), 일자(date), 시가·고가·저가·종가(원), 거래량(주), 거래대금억, 상태, 조회시각. (종목코드, 일자)는 1행.
//    조회시각 = 그 행을 마지막으로 받아 쓴 실제 시각(이번 실행에서 손대지 않은 행은 이전 값 유지). 성공한 실행은 모든 행 상태 OK.
//  - 대상 종목 = fnPageRows()(대회 종목 ± 수정표 ∪ 보유·관심). 페이지 행에서 빠진 종목의 행은 지우지 않고 그대로 둔다(다시 넣으면 재사용).
//  - 자기참조: 자기 표(tblPxStore)를 읽어 이어 쓴다. 자기 표가 없거나 비었으면 tblPxStoreSeed(이관용 초기값)를 읽는다.
//  - 세션 달력: fnPxCalendar(tblSessions — [시세]·[전체]에서 바로 앞 단계 T_Sessions가 새로 만든 표). 마지막 세션 L = 달력 최댓값.
//  - 실행 모드(fnPxRunCtl): quick·full만 KIS 호출. build·빈 값·그 밖 모드(빌드·표 우클릭 새로 고침·파일 열기)는 호출 없이 저장값을 그대로.
//    KIS 호출은 함수 단계 Fetch 안에서만 한다(Excel은 새로 고침 뒤 맨 위 단계들을 백그라운드에서 따로 다시 평가하는데, 값이 함수인
//    단계는 호출을 일으키지 않음). tblRunCtl started(버튼 시작 시각) 이후에 이미 쓴 저장소(자기 표 최근 조회시각 ≥ started)면
//    다시 부르지 않고 자기 표를 돌려준다 — 그 재평가가 시세·일봉 호출을 반복하지 않게(started가 비면 이 확인 없음).
//  - 안전 규칙(docs/business-rules.md의 '가격 이력과 지표'): 오늘이 개장일(fnPxCalendar: 개장확인일 = 오늘이면 오늘개장, 아니면 평일 & tblHolidays에 없음)이고 09:00 이후인데
//    달력에 오늘이 없으면 저장소를 바꾸지 않고 직전 표를 그대로 돌려준다(상태도 그대로 — 오류 표시는 tblSessions 1행). 개장일 09:00
//    이후에는 오늘 이전 날짜 봉에 시세를 쓰지 않는다. quick·full 모두 같다(full의 첫 채우기·보정도 하지 않음).
//  - quick([시세]): 이력이 있는(달력 범위 안 행이 있는) 페이지 행 종목만 멀티 시세(fnMultiQuote, 30종목/회)를 받아 마지막 세션 봉으로 기록.
//      · 개장일 09:00 이후(L = 오늘): 오늘 봉이 없으면 추가, 있으면 갱신(장중 진행 값 / 장 마감 후 확정 값).
//        거래량 0(거래정지·아직 체결 없음)이면 KIS 일봉과 같게 시가=고가=저가=종가=현재가, 거래량·거래대금 0.
//      · 장 시작 전(개장일 09:00 전)·휴장일: 새 봉을 만들지 않는다 — L 봉이 이미 있는 종목만, 거래량이 0보다 큰 시세로 갱신.
//        (장 시작 전 시세가 이미 오늘로 넘어가 거래량 0·시가 0인 경우 L 확정 봉을 덮어쓰지 않으려는 것 — 이때는 값은 그대로 두고
//        그 L 봉의 조회시각만 새로 써서 이번 실행을 기록한다)
//      · 현재가가 없는 시세(시세 없음·호출 실패)는 봉을 만들지 않고 그 종목의 모든 행 상태를 "오류: 사유"로.
//  - full([전체]): quick의 시세 기록 + 아래. 일봉은 fnDailyChart(국내주식기간별시세 FHKST03010100, 수정주가, 100행/회).
//      · 첫 채우기: 이력이 없는 페이지 행 종목 → 최근 260세션(달력 기준, 3회 호출). 실패하면 일자가 빈 상태 행 "오류: 사유" 1개.
//      · 수정주가 변동: 저장소의 '시세 세션 직전 세션' 종가 ≠ 멀티 시세 기준가(inter2_sdpr)이면서 기준가 ≠ 전일종가(inter2_prdy_clpr —
//        그날 권리·병합·분할로 기준가가 조정됨)이면 그 종목 이력 전체(260세션)를 다시 받는다.
//        시세 세션 = 개장일 09:00 이후엔 오늘, 그 밖엔 L. 장 시작 전에는 시세가 이미 오늘로 넘어갔을 수 있어 L 종가와 같아도 일치로 본다.
//      · 보정: 위 종가 ≠ 기준가인데 기준가 = 전일종가(그날 조정 없음 — 대개 장중 [시세]로 쓴 직전 봉이 확정 전 값)이거나,
//        이 종목의 첫 날짜 ~ L 사이에 달력 날짜가 빠졌으면 최근 100세션(빠진 날이 더 오래됐으면 그 전 세션부터)을 다시 받아 덮어쓴다.
//        받은 구간의 가장 오래된 겹치는 날짜 종가가 저장값과 다르면 과거 수정주가 변동으로 보고 이력 전체를 다시 받는다.
//      · 세션 달력 밖 날짜 행 삭제와 정리: 달력에 없는 날짜, 그리고 최근 260세션(달력 기준)보다 오래된 행을 지운다
//        (종목당 약 260행 유지 — [시세]가 하루 1행씩 더한 만큼은 다음 [전체]에서 정리. 저장소 크기가 [시세] 시간을 정하므로).
//        결과는 종목코드·일자 순으로 정렬한다(quick은 정렬하지 않고 새 봉을 뒤에 붙임 — 종목별 날짜 순서는 유지됨).
//  - 실측(2026-10-01): inter2_sdpr = 기준가(수정주가 기준 직전 종가), inter2_prdy_clpr = 실제 직전 종가 — 병합 재상장 007110은 3,015 vs 603.
//    KIS 일봉은 거래정지일에도 행이 있다(시가=고가=저가=종가=직전 종가, 거래량 0) → 거래정지 종목은 빠진 세션으로 잡히지 않는다.
//    수정주가 일봉은 가격·거래량을 조정하고 거래대금은 그대로 둔다(007110: 9/1 이전 거래량 ÷5, 거래대금 동일).
//  - 한계: 기준가 비교는 '시세 세션' 하루의 조정만 알아챈다. 권리락·병합일에 [전체]를 돌리지 못하고 그 뒤 [시세]가 새 봉을 쓰면
//    이후 [전체]에서는 직전 봉이 기준가와 같아져 과거 이력이 수정되지 않을 수 있다(보정 경로의 겹침 검사가 받은 구간 안의 변동만 잡음).
//  - 실패: 시세·일봉이 전부 실패하면 오류로 올려 직전 표 유지(fnFallback, 상태 "이전 데이터(갱신 실패: 사유)", 열 형식 다시 지정).
//    직전 표가 없으면 종목코드·일자가 빈 1행에 "오류: 사유". 일부 종목만 실패하면 그 종목 행만 "오류: 사유".
let
    Cols = {"종목코드", "일자", "시가", "고가", "저가", "종가", "거래량", "거래대금억", "상태", "조회시각"},
    Types = {{"종목코드", type text}, {"일자", type date}, {"시가", type number}, {"고가", type number}, {"저가", type number},
             {"종가", type number}, {"거래량", type number}, {"거래대금억", type number}, {"상태", type text}, {"조회시각", type datetime}},
    Schema = type table [종목코드 = text, 일자 = date, 시가 = nullable number, 고가 = nullable number, 저가 = nullable number,
                         종가 = number, 거래량 = nullable number, 거래대금억 = nullable number, 상태 = text, 조회시각 = nullable datetime],
    FillSessions = 260,      // 첫 채우기·전체 재조회 구간(세션 수, 최소 260세션)
    HealSessions = 100,      // 보정 재조회 구간(일봉 1회 호출 = 100행)
    Ctl = fnPxRunCtl(),
    Live = Ctl[Mode] = "quick" or Ctl[Mode] = "full",
    FetchedAt = DateTime.LocalNow(),
    Digits = {"0".."9"},
    CodeChars = {"0".."9", "A".."Z"},
    Txt = (v as any) as nullable text =>
        if v = null then null else let t = Text.Trim(try Text.From(v) otherwise "") in if t = "" then null else t,
    Code = (v as any) as nullable text =>
        let
            t0 = Txt(v),
            t1 = if t0 = null then null else Text.Upper(t0),
            t2 = if t1 <> null and Text.Length(t1) < 6 and Text.Select(t1, Digits) = t1 then Text.PadStart(t1, 6, "0") else t1
        in
            if t2 <> null and Text.Length(t2) = 6 and Text.Select(t2, CodeChars) = t2 then t2 else null,
    SetOf = (xs as list) as record => let d = List.Distinct(List.RemoveNulls(xs)) in Record.FromList(List.Repeat({true}, List.Count(d)), d),
    KeyOf = (d as date) as text => Date.ToText(d, [Format = "yyyyMMdd", Culture = "en-US"]),
    Same = (a as nullable number, b as nullable number) as logical =>
        a <> null and b <> null and Number.Round(a, 2) = Number.Round(b, 2),
    Read = (name as text) as any => try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null,

    // 통합문서 표 → 형식 있는 저장소 행. 자기 표는 이 쿼리가 형식을 붙여 쓴 값이라(불변식: 키 유일·형식 일치) 형식만 붙이고
    // 상태 행(일자 빈칸)만 버린다(527종목 × 260세션 = 13만 행에서 오류 값 정리·중복 제거가 약 2초). 시드는 코드 정규화·정리·중복 제거.
    FromSheet = (t as any, seed as logical) as table =>
        if not (t is table) then #table(Schema, {})
        else
            let
                s = Table.SelectColumns(t, Cols, MissingField.UseNull),
                n = if seed then Table.TransformColumns(s, {{"종목코드", Code}}) else s,
                typed = Table.TransformColumnTypes(n, Types),
                fixed = if seed then Table.ReplaceErrorValues(typed, List.Transform(Cols, each {_, null})) else typed,
                ok = Table.SelectRows(fixed, each [종목코드] <> null and [일자] <> null and [종가] <> null)
            in
                if seed then Table.Distinct(ok, {"종목코드", "일자"}) else ok,
    OwnRows = Table.Buffer(FromSheet(Read("tblPxStore"), false)),
    PrevStore = if Table.RowCount(OwnRows) > 0 then OwnRows else Table.Buffer(FromSheet(Read("tblPxStoreSeed"), true)),
    Kept = Table.TransformColumns(PrevStore, {{"상태", each if _ = null then "OK" else _, type text}}),
    PrevAllOk = List.MatchesAll(PrevStore[상태], each _ = "OK"),
    // 이번 버튼 실행(started 이후)에 이미 쓴 저장소면 다시 부르지 않음 — 새로 고침 뒤 Excel의 백그라운드 재평가가 호출을 반복하지 않게
    LastStamp = List.Max(List.RemoveNulls(OwnRows[조회시각]), null),
    Already = Ctl[Started] <> null and LastStamp <> null and LastStamp >= Ctl[Started],

    // 상태 열을 종목별 오류 사유(없으면 OK)로 다시 씀. 오류가 없고 원래 상태가 모두 OK면 그대로(13만 행 다시 쓰기 생략)
    WithStatus = (t as table, errMap as record, allOk as logical) as table =>
        if Record.FieldCount(errMap) = 0 and allOk then Table.SelectColumns(t, Cols)
        else Table.SelectColumns(
            Table.AddColumn(Table.RemoveColumns(t, {"상태"}), "상태", each Record.FieldOrDefault(errMap, [종목코드], "OK"), type text),
            Cols),
    Msg = (e as record) as text =>
        let m = try Text.Trim(Text.From(e[Message])) otherwise "" in if m = "" then "조회 실패" else Text.Start(m, 120),

    // KIS 호출은 모두 이 함수 안에서만 — Excel이 맨 위 단계들을 따로 평가해도 함수 값 단계는 호출을 일으키지 않는다
    Fetch = () as table =>
        let
            T = Ctl[Today],
            Tm = Ctl[Time],
            AfterOpen = Tm >= #time(9, 0, 0),
            CalInfo = fnPxCalendar(Ctl[Now]),
            Cal0 = CalInfo[Dates],
            Cal = if List.IsEmpty(Cal0) then error Error.Record("세션 달력 없음", "세션 달력(tblSessions)이 비어 있습니다(T_Sessions 실패 또는 첫 실행 전).") else Cal0,
            N = List.Count(Cal),
            OpenToday = CalInfo[OpenToday],
            Unconfirmed = OpenToday and AfterOpen and not List.Contains(Cal, T),
            L = List.Last(Cal),
            // 오늘 봉 추가·갱신: 달력(KOSPI 일봉)에 오늘이 있으면(09:00 이후에만 들어감) 오늘은 세션이 확실하다
            // — 개장 여부 판단(휴장일조회·대체 판단)보다 거래소 데이터를 우선
            TodayMode = L = T,
            PreOpen = OpenToday and not AfterOpen,                // 개장일 장 시작 전 — 마지막 세션 봉만 갱신
            DateAt = (p as number) as nullable date => if p < 0 or p >= N then null else Cal{p},
            PosRec = Record.FromList(List.Positions(Cal), List.Transform(Cal, KeyOf)),
            PosOf = (d as date) as nullable number => Record.FieldOrDefault(PosRec, KeyOf(d), null),
            PrevL = DateAt(N - 2),
            CalStart = Cal{0},
            FillFrom = DateAt(List.Max({0, N - FillSessions})),
            HealFrom = DateAt(List.Max({0, N - HealSessions})),
            KeepT = #table(type table [__d = date], List.Transform(List.Select(Cal, each _ >= FillFrom), each {_})),
            PageCodes = List.Buffer(List.Distinct(List.RemoveNulls(List.Transform(fnPageRows()[종목코드], Code)))),
            PageSet = SetOf(PageCodes),
            InWin = Table.SelectRows(PrevStore, each [일자] >= CalStart and [일자] <= L),
            HistSet = SetOf(List.Distinct(InWin[종목코드])),
            HistCodes = List.Buffer(List.Select(PageCodes, each Record.HasFields(HistSet, _))),

            // 멀티 시세 → 마지막 세션 봉 (봉을 쓰지 않으면 null). lrow = 저장소의 기존 L 봉(없으면 null)
            //  장 시작 전·휴장일에 거래 없는 시세(이미 오늘로 넘어간 시세·거래정지)는 값을 바꾸지 않고 기존 L 봉의 조회시각만 새로 씀
            //  — 이번 실행 기록이 남아야 started 확인이 백그라운드 재평가의 중복 호출을 막는다
            BarOf = (q as record, lrow as nullable record) as nullable list =>
                let
                    c = q[현재가],
                    v = q[거래량],
                    traded = v <> null and v > 0,
                    bar = if c = null then null
                          else if traded then {q[종목코드], L, q[시가], q[고가], q[저가], c, v, q[거래대금억], "OK", FetchedAt}
                          else {q[종목코드], L, c, c, c, c, 0, (if q[거래대금억] = null then 0 else q[거래대금억]), "OK", FetchedAt},
                    touch = if lrow = null then null
                            else {q[종목코드], L, lrow[시가], lrow[고가], lrow[저가], lrow[종가], lrow[거래량], lrow[거래대금억], "OK", FetchedAt}
                in
                    if q[상태] <> "OK" then null
                    else if TodayMode then bar
                    else if traded and lrow <> null then bar
                    else touch,
            Q = Table.Buffer(fnMultiQuote(HistCodes)),
            QRows = List.Buffer(Table.ToRecords(Q)),
            QRec = Record.FromList(QRows, Q[종목코드]),
            QErr = Record.FromList(List.Transform(List.Select(QRows, each [상태] <> "OK"), each [상태]),
                                   List.Transform(List.Select(QRows, each [상태] <> "OK"), each [종목코드])),
            QAnyOk = List.AnyTrue(List.Transform(QRows, each [상태] = "OK")),

            // ---------------- quick
            Quick =
                let
                    AtLRows = Table.Buffer(Table.SelectRows(PrevStore, each [일자] = L)),
                    AtL = Record.FromList(Table.ToRecords(AtLRows), AtLRows[종목코드]),
                    NewBars = List.RemoveNulls(List.Transform(QRows, (q) => BarOf(q, Record.FieldOrDefault(AtL, q[종목코드], null)))),
                    NewSet = SetOf(List.Transform(NewBars, each _{0})),
                    Base = Table.SelectRows(PrevStore, each [일자] <> L or not Record.HasFields(NewSet, [종목코드])),
                    Out = WithStatus(Table.Combine({Base, #table(Schema, NewBars)}), QErr, PrevAllOk)
                in
                    if List.Count(QRows) > 0 and not QAnyOk then error ("멀티 시세 전부 실패: " & Text.Middle(QRows{0}[상태], 4)) else Out,

            // ---------------- full
            Full =
                let
                    Groups = Table.Buffer(Table.Group(PrevStore, {"종목코드"}, {{"g", each _, Schema}})),   // 두 번 참조 → 버퍼(재계산 방지)
                    GroupRec = Record.FromList(Groups[g], Groups[종목코드]),
                    Chart = (c as text, from as date) as table => Table.Buffer(fnDailyChart(c, from, T)),
                    // 일봉 → 저장소 행(달력 날짜만 — 장 시작 전 오늘 행·달력 밖 날짜 제외)
                    ToStore = (c as text, t as table) as table =>
                        let
                            sel = Table.SelectRows(t, each Record.HasFields(PosRec, KeyOf([일자]))),
                            Px = (v as nullable number) as nullable number => if v = null or v = 0 then null else v,
                            rows = List.Transform(Table.ToRecords(sel), (r) =>
                                {c, r[일자], Px(r[시가]), Px(r[고가]), Px(r[저가]), r[종가], r[거래량],
                                 if r[거래대금] = null then null else r[거래대금] / 100000000, "OK", FetchedAt})
                        in
                            #table(Schema, rows),
                    PlanOf = (c as text) as record =>
                        let
                            g0 = Record.FieldOrDefault(GroupRec, c, null),
                            g = if g0 = null then null else Table.Buffer(Table.SelectRows(g0, each [일자] >= CalStart and [일자] <= L)),
                            has = g <> null and Table.RowCount(g) > 0,
                            closeRec = if has then Record.FromList(g[종가], List.Transform(g[일자], KeyOf)) else [],
                            CloseAt = (d as nullable date) as nullable number => if d = null then null else Record.FieldOrDefault(closeRec, KeyOf(d), null),
                            first = if has then List.Min(g[일자]) else null,
                            // 빠진 세션: 첫 날짜 ~ L 중 저장소에 없는 달력 날짜(장중 오늘은 시세가 쓰므로 제외)
                            missing = if not has then {}
                                      else List.Select(Cal, (d) => d >= first and not Record.HasFields(closeRec, KeyOf(d)) and not (TodayMode and d = T)),
                            q = Record.FieldOrDefault(QRec, c, null),
                            qok = q <> null and q[상태] = "OK",
                            sdpr = if qok then q[기준가] else null,
                            pcl = if qok then q[전일종가] else null,
                            cands = List.RemoveNulls(if PreOpen then {CloseAt(PrevL), CloseAt(L)} else {CloseAt(PrevL)}),
                            mismatch = sdpr <> null and not List.IsEmpty(cands) and not List.AnyTrue(List.Transform(cands, each Same(_, sdpr))),
                            event = sdpr <> null and pcl <> null and not Same(sdpr, pcl),
                            action = if not has then "first"
                                     else if mismatch and event then "full"
                                     else if mismatch or not List.IsEmpty(missing) then "heal"
                                     else "none",
                            earliest = if List.IsEmpty(missing) then null else List.Min(missing),
                            healFrom = if earliest = null then HealFrom
                                       else List.Min({HealFrom, DateAt(List.Max({0, PosOf(earliest) - 1}))})
                        in
                            [code = c, g = g, action = action, healFrom = healFrom, closeRec = closeRec],
                    // 보정·재조회 실행(오류는 바깥 try가 잡아 이전 행 유지)
                    Compute = (p as record) as table =>
                        let
                            c = p[code],
                            heal = p[action] = "heal",
                            h = if heal then ToStore(c, Chart(c, p[healFrom])) else null,
                            common = if heal then List.Select(h[일자], each Record.HasFields(p[closeRec], KeyOf(_))) else {},
                            overlap = if List.IsEmpty(common) then null else List.Min(common),
                            okHeal = heal and overlap <> null
                                     and Same(Record.Field(p[closeRec], KeyOf(overlap)), Table.SelectRows(h, each [일자] = overlap){0}[종가]),
                            hSet = if okHeal then SetOf(List.Transform(h[일자], KeyOf)) else [],
                            merged = if okHeal then Table.Combine({h, Table.SelectRows(p[g], each not Record.HasFields(hSet, KeyOf([일자])))})
                                     else ToStore(c, Chart(c, FillFrom))
                        in
                            Table.Sort(merged, {{"일자", Order.Ascending}}),
                    // 첫 채우기 실패 → rows = null(달력 범위 안 행이 원래 없음), 보정·재조회 실패 → 이전 행 유지.
                    // 버퍼는 칸 안의 오류를 올리지 않으므로(빈칸으로 적재됨) 오류 칸이 있으면 직접 오류로 올려 이 try가 잡게 함
                    RunOne = (p as record) as record =>
                        if p[action] = "none" then [code = p[code], action = p[action], rows = p[g], fetched = false, err = null]
                        else
                            let r = try (let t = Table.Buffer(Compute(p))
                                         in if Table.RowCount(Table.SelectRowsWithErrors(t)) > 0 then error "일봉 값 오류" else t)
                            in if r[HasError] then [code = p[code], action = p[action], rows = if p[action] = "first" then null else p[g],
                                                    fetched = false, err = "일봉 " & Msg(r[Error])]
                               else [code = p[code], action = p[action], rows = r[Value], fetched = true, err = null],
                    Runs = List.Buffer(List.Transform(PageCodes, (c) => RunOne(PlanOf(c)))),
                    // 다시 받지 않은 종목은 시세로 마지막 세션 봉 기록(quick과 같은 규칙)
                    Written = List.Transform(Runs, (x) =>
                        if x[rows] = null or x[fetched] then x
                        else
                            let
                                q = Record.FieldOrDefault(QRec, x[code], null),
                                lrows = Table.SelectRows(x[rows], each [일자] = L),
                                bar = if q = null then null else BarOf(q, if Table.IsEmpty(lrows) then null else lrows{0})
                            in
                                if bar = null then x
                                else x & [rows = Table.Combine({Table.SelectRows(x[rows], each [일자] <> L), #table(Schema, {bar})})]),
                    Others = Table.SelectRows(PrevStore, each not Record.HasFields(PageSet, [종목코드])),
                    All0 = Table.Combine(List.Transform(List.Select(Written, each [rows] <> null), each [rows]) & {Others}),
                    // 세션 달력 밖 날짜 삭제 + 최근 260세션(FillFrom 이후)만 남김 — 저장소 크기를 종목당 약 260행으로 유지
                    // ([시세]가 하루 1행씩 더한 행은 다음 [전체]에서 정리). 종목·일자 순으로 정렬해 둔다
                    InCal = Table.Sort(Table.RemoveColumns(Table.Join(All0, "일자", KeepT, "__d", JoinKind.Inner), {"__d"}),
                                       {{"종목코드", Order.Ascending}, {"일자", Order.Ascending}}),
                    FetchErr = List.Select(Runs, each [err] <> null),
                    ErrPairs = List.Transform(FetchErr, each {[code], [err]})
                               & List.Transform(List.Select(QRows, (q) => q[상태] <> "OK"
                                     and List.IsEmpty(List.Select(FetchErr, each [code] = q[종목코드]))), each {[종목코드], [상태]}),
                    ErrMap = Record.FromList(List.Transform(ErrPairs, each if Text.StartsWith(_{1}, "오류") then _{1} else "오류: " & _{1}),
                                             List.Transform(ErrPairs, each _{0})),
                    // 첫 채우기에 실패한 종목은 행이 없으므로 일자가 빈 상태 행 1개로 실패를 드러냄(다음 실행에서 읽을 때는 버려짐)
                    StatusOnly = #table(Schema, List.Transform(List.Select(FetchErr, each [rows] = null), (x) =>
                                     {x[code], null, null, null, null, null, null, null, "오류: " & x[err], FetchedAt})),
                    Out = Table.Combine({WithStatus(InCal, ErrMap, false), StatusOnly}),
                    Attempted = List.Count(QRows) + List.Count(List.Select(Runs, each [action] <> "none")),
                    AnyOk = QAnyOk or List.AnyTrue(List.Transform(Runs, each [fetched]))
                in
                    if Attempted > 0 and not AnyOk then
                        error ("시세·일봉 전부 실패: " & (if List.IsEmpty(ErrPairs) then "사유 불명" else ErrPairs{0}{1}))
                    else Out
        in
            if Unconfirmed then Kept else if Ctl[Mode] = "full" then Full else Quick,
    Tried = try Table.Buffer(if Live and not Already then Fetch() else Kept),
    Reason = if not Tried[HasError] then null
             else let m = Tried[Error][Message] in if m = null or m = "" then "사유 불명" else m,
    KeptOnFail = if Tried[HasError] then fnFallback("tblPxStore", Cols, Reason) else null,
    Warn = #table(Cols, {{null, null, null, null, null, null, null, null, "오류: " & Text.Start(if Reason = null then "" else Reason, 120), DateTime.LocalNow()}}),
    // fnFallback·경고 행은 형식이 없는 표 → 열 형식을 다시 지정(일자·조회시각이 일련번호로 보이지 않게)
    Final =
        if not Tried[HasError] then Tried[Value]
        else Table.TransformColumnTypes(if Table.RowCount(KeptOnFail) > 0 then KeptOnFail else Warn, Types)
in
    Final
