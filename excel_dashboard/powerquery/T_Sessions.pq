// T_Sessions : 세션 달력 — KOSPI 지수(0001) 일봉에 나타난 날짜 (세션·마지막 세션·안전 규칙 — docs/business-rules.md의 '용어'·'가격 이력과 지표', 2026-10-01) → tblSessions
//  - 열: 일자(date), 오늘개장(Y/N/공란), 개장확인일(date), 상태, 조회시각. 오늘개장·개장확인일은 모든 행에 같은 값. 정렬 = 일자 내림차순.
//  - 실행 모드(fnPxRunCtl = tblRunCtl mode): quick([시세])·full([전체])만 KIS를 부른다. 그 밖(build·빈 값·sector·analysis 등 —
//    빌드, 표 우클릭 새로 고침, 파일 열기)은 호출 없이 자기 표 → 없거나 비면 tblSessionsSeed → 없으면 열만 있는 빈 표.
//    KIS 호출은 함수 단계 Fetch 안에서만 한다(Excel은 새로 고침 뒤 맨 위 단계들을 백그라운드에서 따로 다시 평가하는데, 값이 함수인
//    단계는 호출을 일으키지 않음). 또 tblRunCtl started(버튼 시작 시각) 이후에 이미 받은 표(자기 표 조회시각 ≥ started)면 다시
//    부르지 않고 자기 표를 돌려준다 — 그 재평가가 같은 호출을 반복하지 않게(started가 비면 이 확인 없음).
//  - 달력 원천: 시장별 투자자매매동향(일별) [국내주식-075] FHPTJ04040000, U/0001(KOSPI)/KSP. 입력일 이하 최근 300영업일을 1회에 줌
//    (2026-10-01 실측 300행 2025-07-10 ~ 2026-10-01). 같은 KOSPI 일봉을 주는 업종기간별시세 FHKUP03500100은 50행,
//    업종 일자별지수 FHPUP02120000은 100행이라 이 API를 쓴다. 날짜는 stck_bsop_date만 쓴다.
//    1Y(250세션 전) + 줄무늬 계산에 261세션 이상이 필요하므로 261개 미만이면 가장 오래된 날짜 전날로 다시 조회(최대 3회).
//  - 장중 오늘 행(2026-10-01 실측): 11:37 조회에서 세 API(FHPTJ04040000·FHPUP02120000·FHKUP03500100) 모두 오늘(20261001) 행이 있었고,
//    같은 FHPTJ04040000은 10:50:50에도 KRX 업종 코드(0013·1028)로 오늘 행을 돌려줬다(병렬 작업 기록). 장중에는 진행 중 값으로 오늘 행이
//    있다. 09:00 직후 몇 분 만에 생기는지는 이 작업 시각상 관측하지 못했다 — 생기기 전에 [시세]를 누르면 아래 안전 규칙이
//    저장소 쓰기를 막고 이 표에 '오류: 오늘 세션 미확인'을 남긴다(다시 누르면 해소).
//  - 달력 규칙: '지금'(fnPxRunCtl — 시험용 now_override 반영) 뒤의 날짜는 빼고, 09:00 전이면 오늘 날짜도 뺀다
//    (장 시작 전 '마지막 세션' = 직전 영업일. 실제 시각에도 같은 규칙을 적용해 장 시작 전 응답에 오늘 행이 있어도 쓰지 않는다).
//    API 입력일은 min(모사 날짜, 실제 오늘)이라 미래 날짜를 모사해도 오늘까지의 달력을 받는다.
//  - 오늘 개장 여부: 국내휴장일조회 [국내주식-040] chk-holiday CTCA0903R의 opnd_yn(KIS가 하루 1회 호출을 요청). 날짜가 바뀐 뒤 첫 실행에서만
//    부르고(직전 표의 개장확인일 = 오늘이면 그 오늘개장 값을 재사용), 응답 output[]에서 bass_dt = 오늘인 행의 opnd_yn을 쓴다
//    (응답은 기준일부터 24일치 목록 — 2026-10-01 실측: 10/1·10/2 Y, 10/3·10/4·10/5 N, 10/9 N).
//    API가 실패하면 평일이고 tblHolidays에 없으면 Y, 아니면 N으로 정하고 개장확인일은 비운다(다음 실행에서 다시 확인).
//  - 안전 규칙: 오늘개장 = Y이고 09:00 이후인데 달력에 오늘이 없으면 가장 최근 날짜 행 1개에만
//    상태 "오류: 오늘 세션 미확인"(나머지 행은 OK) → 버튼 단계 판정 '일부 오류 1건'. 이때 T_PxStore는 저장소를 바꾸지 않는다
//    (T_PxStore가 같은 조건을 스스로 다시 판정 — 이 표가 직전 데이터로 남은 경우에도 막히도록). 휴장일에는 작동하지 않는다.
//  - 실패(토큰·네트워크·빈 응답): 직전 표 유지 + 상태 "이전 데이터(갱신 실패: 사유)"(fnFallback, 열 형식 다시 지정).
//    직전 표가 없으면 일자가 빈 1행에 상태 "오류: 사유"(빈 표로는 실패가 보이지 않으므로). 조회시각은 실제 시각.
let
    Cols = {"일자", "오늘개장", "개장확인일", "상태", "조회시각"},
    Types = {{"일자", type date}, {"오늘개장", type text}, {"개장확인일", type date}, {"상태", type text}, {"조회시각", type datetime}},
    Ctl = fnPxRunCtl(),
    Live = Ctl[Mode] = "quick" or Ctl[Mode] = "full",
    Read = (name as text) as any => try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null,
    ToDate = (v as any) as nullable date => if v = null then null else try Date.From(v) otherwise null,
    YN = (v as any) as nullable text =>
        let t = if v = null then "" else Text.Upper(Text.Trim(try Text.From(v) otherwise ""))
        in if t = "Y" or t = "N" then t else null,
    Tidy = (t as any) as table =>
        if not (t is table) then #table(Cols, {})
        else
            let
                s = Table.SelectColumns(t, Cols, MissingField.UseNull),
                x = Table.TransformColumns(s, {
                        {"일자", ToDate}, {"오늘개장", YN}, {"개장확인일", ToDate},
                        {"상태", each if _ = null then "OK" else Text.From(_)},
                        {"조회시각", each if _ = null then null else try DateTime.From(_) otherwise null}})
            in
                Table.Distinct(Table.SelectRows(x, each [일자] <> null), {"일자"}),
    OwnT = Table.Buffer(Tidy(Read("tblSessions"))),
    Prev = Table.Buffer(if Table.RowCount(OwnT) > 0 then OwnT else Tidy(Read("tblSessionsSeed"))),
    Kept = Table.TransformColumnTypes(Table.Sort(Prev, {{"일자", Order.Descending}}), Types),
    // 이번 버튼 실행(started 이후)에 이미 받은 표면 다시 부르지 않음 — 새로 고침 뒤 Excel의 백그라운드 재평가가 호출을 반복하지 않게
    LastStamp = List.Max(List.RemoveNulls(OwnT[조회시각]), null),
    Already = Ctl[Started] <> null and LastStamp <> null and LastStamp >= Ctl[Started],
    // KIS 호출은 모두 이 함수 안에서만 — Excel이 맨 위 단계들을 따로 평가해도 함수 값 단계는 호출을 일으키지 않는다
    Fetch = () as table =>
        let
            T = Ctl[Today],
            Tm = Ctl[Time],
            Open9 = #time(9, 0, 0),
            Now = DateTime.LocalNow(),
            RealToday = Date.From(Now),
            Path = "/uapi/domestic-stock/v1/quotations/inquire-investor-daily-by-market",
            Need = 261,
            Page = (endDate as date) as list =>
                let
                    d = fnYmdText(endDate),
                    J = fnKisGet(Path, "FHPTJ04040000", [FID_COND_MRKT_DIV_CODE = "U", FID_INPUT_ISCD = "0001", FID_INPUT_DATE_1 = d,
                            FID_INPUT_ISCD_1 = "KSP", FID_INPUT_DATE_2 = d, FID_INPUT_ISCD_2 = "0001"]),
                    O = Record.FieldOrDefault(J, "output", {}),
                    L = if O is list then O else {}
                in
                    List.Buffer(List.Distinct(List.RemoveNulls(List.Transform(L, each
                        if _ is record then fnYmd(Record.FieldOrDefault(_, "stck_bsop_date", null)) else null)))),
            Collect = (acc as list, endDate as date, n as number) as list =>
                let
                    pg = Page(endDate),
                    all = List.Distinct(acc & pg)
                in
                    if List.Count(all) >= Need or List.IsEmpty(pg) or n >= 3 then all
                    else @Collect(all, Date.AddDays(List.Min(pg), -1), n + 1),
            AllDates = Collect({}, List.Min({T, RealToday}), 1),
            Dates0 = List.Sort(List.Select(AllDates, each _ < T or (_ = T and Tm >= Open9)), Order.Descending),
            Dates = if List.IsEmpty(Dates0) then error Error.Record("데이터 없음", "KOSPI 일봉 응답에 날짜가 없습니다.") else Dates0,

            // 오늘 개장 여부 — 같은 날은 직전 값 재사용, 날짜가 바뀐 뒤 첫 실행에서만 API 호출
            PrevMatch = Table.SelectRows(Prev, each [개장확인일] = T and [오늘개장] <> null),
            Reuse = Table.RowCount(PrevMatch) > 0,
            ApiFetch = () as text =>
                let
                    J = fnKisGet("/uapi/domestic-stock/v1/quotations/chk-holiday", "CTCA0903R",
                            [BASS_DT = fnYmdText(T), CTX_AREA_NK = "", CTX_AREA_FK = ""]),
                    O = Record.FieldOrDefault(J, "output", {}),
                    L = if O is list then O else if O is record then {O} else {},
                    M = List.Select(L, each _ is record and fnYmd(Record.FieldOrDefault(_, "bass_dt", null)) = T),
                    v = if List.IsEmpty(M) then null else YN(Record.FieldOrDefault(M{0}, "opnd_yn", null))
                in
                    if v = null then error "국내휴장일조회 응답에 오늘 개장 여부가 없습니다" else v,
            Api = if Reuse then null else try ApiFetch(),
            HolT = Read("tblHolidays"),
            Holidays = if HolT is table and List.Contains(Table.ColumnNames(HolT), "휴장일")
                       then List.RemoveNulls(List.Transform(Table.Column(HolT, "휴장일"), ToDate)) else {},
            Fallback = if Date.DayOfWeek(T, Day.Monday) < 5 and not List.Contains(Holidays, T) then "Y" else "N",
            OpenYN = if Reuse then PrevMatch{0}[오늘개장] else if not Api[HasError] then Api[Value] else Fallback,
            ChkDate = if Reuse or not Api[HasError] then T else null,

            // 안전 규칙: 개장일 09:00 이후인데 달력에 오늘이 없으면 가장 최근 행 하나에만 오류 상태
            Unconfirmed = OpenYN = "Y" and Tm >= Open9 and not List.Contains(Dates, T),
            Top = Dates{0},
            Rows = List.Transform(Dates, (d) =>
                {d, OpenYN, ChkDate, if Unconfirmed and d = Top then "오류: 오늘 세션 미확인" else "OK", Now}),
            Out = #table(type table [일자 = date, 오늘개장 = nullable text, 개장확인일 = nullable date, 상태 = text, 조회시각 = datetime], Rows)
        in
            Out,
    Tried = try Table.Buffer(if Live and not Already then Fetch() else Kept),
    Reason = if not Tried[HasError] then null
             else let m = Tried[Error][Message] in if m = null or m = "" then "사유 불명" else m,
    KeptOnFail = if Tried[HasError] then fnFallback("tblSessions", Cols, Reason) else null,
    Warn = #table(Cols, {{null, null, null, "오류: " & Text.Start(if Reason = null then "" else Reason, 120), DateTime.LocalNow()}}),
    // fnFallback·경고 행은 형식이 없는 표 → 열 형식을 다시 지정(일자·조회시각이 일련번호로 보이지 않게)
    Final =
        if not Tried[HasError] then Tried[Value]
        else Table.TransformColumnTypes(if Table.RowCount(KeptOnFail) > 0 then KeptOnFail else Warn, Types)
in
    Final
