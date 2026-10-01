// fnPxCalendar : 세션 달력(tblSessions) 읽기 — T_PxStore·T_PxMetrics 공용 (spec §3 세션, 2026-10-01 T5)
//  - 일자 원천: tblSessions, 없거나 비면 tblSessionsSeed(이관용 초기값). 일자가 빈 행(실패 경고 행)은 건너뛴다.
//  - 반환 Dates = 오름차순 날짜 목록. '지금'(now 인자 — fnPxRunCtl의 시험용 시계) 기준으로 미래 날짜를 빼고, 09:00 전이면
//    오늘 날짜도 뺀다. T_Sessions가 이미 같은 규칙으로 만들지만, 오래된 달력·시드에도 같은 안전장치가 걸리도록 여기서 다시 적용.
//  - OpenToday(오늘 개장 여부): 표의 개장확인일 = 오늘이고 오늘개장이 Y/N이면 그 값(국내휴장일조회 결과),
//    아니면 평일이고 tblHolidays(휴장일 열)에 없으면 개장으로 본다(spec: API 실패 시 대체 판단). Confirmed = 앞의 경우.
//  - 반환: [Dates = list(date), OpenToday = logical, Confirmed = logical]
(now as datetime) as record =>
let
    T = Date.From(now),
    Tm = Time.From(now),
    Read = (name as text) as any => try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null,
    ToDate = (v as any) as nullable date => if v = null then null else try Date.From(v) otherwise null,
    YN = (v as any) as nullable text =>
        let t = if v = null then "" else Text.Upper(Text.Trim(try Text.From(v) otherwise ""))
        in if t = "Y" or t = "N" then t else null,
    RowsOf = (t as any) as list =>
        if not (t is table) or not List.Contains(Table.ColumnNames(t), "일자") then {}
        else List.Buffer(List.Select(
            List.Transform(Table.ToRecords(Table.SelectColumns(t, {"일자", "오늘개장", "개장확인일"}, MissingField.UseNull)),
                (r) => [d = ToDate(r[일자]), open = YN(r[오늘개장]), chk = ToDate(r[개장확인일])]),
            each [d] <> null)),
    Own = RowsOf(Read("tblSessions")),
    Src = if List.IsEmpty(Own) then RowsOf(Read("tblSessionsSeed")) else Own,
    All = List.Distinct(List.Transform(Src, each [d])),
    Dates = List.Sort(List.Select(All, each _ < T or (_ = T and Tm >= #time(9, 0, 0)))),
    Checked = List.Select(Src, each [chk] = T and [open] <> null),
    Confirmed = not List.IsEmpty(Checked),
    HolT = Read("tblHolidays"),
    Holidays = if HolT is table and List.Contains(Table.ColumnNames(HolT), "휴장일")
               then List.RemoveNulls(List.Transform(Table.Column(HolT, "휴장일"), ToDate)) else {},
    Fallback = Date.DayOfWeek(T, Day.Monday) < 5 and not List.Contains(Holidays, T),
    OpenToday = if Confirmed then Checked{0}[open] = "Y" else Fallback
in
    [Dates = List.Buffer(Dates), OpenToday = OpenToday, Confirmed = Confirmed]
