// T_FlowU : 종목별 투자자 순매수 — 외국인·기관계 1D·1W·1M (억원), 행 = 페이지 행 fnPageRows() (spec R6, 2026-10-01 T6) → 표 tblFlowU
//  [종목별 투자자매매동향(일별)] FHPTJ04160001  /uapi/domestic-stock/v1/quotations/investor-trade-by-stock-daily
//  - 호출: 종목당 1회. FID_COND_MRKT_DIV_CODE=J, FID_INPUT_ISCD=종목코드, FID_INPUT_DATE_1=조회일, FID_ORG_ADJ_PRC·FID_ETC_CLS_CODE=공백.
//    output2 = 조회일 '이하' 30세션(최신순, 연속 조회 없음). output1(현재가 요약)은 쓰지 않음. 조회일은 주말·휴일이어도 되고
//    그러면 직전 세션부터 옴(2026-10-01 실측: 일요일 9/27 → 9/23까지, 추석 연휴 반영).
//  - 응답 필드 → 열 (금액은 백만원 → ÷100 = 억원. 실측: 삼성전자 2026-09-30 외국인 -862,043백만원 = -8,620억원
//    ≈ 순매수 -3,191,284주 × 27만원)
//      stck_bsop_date    → 세션 날짜 (YYYYMMDD)
//      frgn_ntby_tr_pbmn → 외국인 = 외국인 전체: 등록(frgn_reg_ntby_pbmn) + 비등록 = 기타외국인(frgn_nreg_ntby_pbmn)의 합과 일치(±1 반올림)
//      orgn_ntby_tr_pbmn → 기관계 = API가 주는 기관 합계: 증권·투신·사모·은행·보험·종금·기금 순매수의 합과 일치(±2) — 여기서 따로 합산하지 않음
//      stck_clpr         → 종가: 세션 판정용(아래). prsn_ntby_tr_pbmn(개인) → 오늘 행이 채워졌는지 판정용
//      (외국인+기관계+개인+기타법인+기타단체 ≈ 0 항등식이 표본 전체 행에서 ±1 안으로 성립 → 필드 의미·단위 확인)
//  - 세션 = 종가(stck_clpr)가 있는 행. 응답은 시세가 없는 종목·날짜도 순매수 0으로 채워 주므로(실측) 0을 값으로 믿지 않는다.
//      · 없는 종목코드(예: 999999): rt_cd 0인데 30행이 모두 시세 빈칸·순매수 0 → 시세 있는 행이 하나도 없으면 그 종목만 "오류: 종목코드 확인 …"
//      · 신규 상장: 상장 전 날짜는 시세 빈칸·순매수 0, 상장일 행은 시세·거래량은 있는데 순매수가 0(투자자별 자료 없음; 2026-09 상장 4종목 실측)
//        → 더 과거에 시세 없는 행이 있는 가장 오래된 시세 행은 상장일로 보고(순매수 세 값이 모두 0일 때만) 합계에서 뺀다.
//        세션이 모자라면 해당 열은 빈칸: 1W는 5세션, 1M은 21세션이 있어야 함(합계 칸 안에 빈 값이 있어도 빈칸 — 0으로 채우지 않음)
//  - 완료 세션 (spec R6): 오늘 행은 15:30(장 마감) 전에는 쓰지 않음. 그리고 이 API는 오늘 날짜를 15:40 전에는 거절한다
//    (rt_cd 2 OPSQ2001 "TIME LIMIT 00:00 ~ 15:40"; 미래 날짜는 "must be before today" — 2026-10-01 11:36 실측) → 요청일 = 15:40 이후 오늘, 그 전 어제.
//    15:40 이후에도 오늘 날짜가 OPSQ2001로 거절되면(시계 차이 등) 첫 종목으로 한 번 확인해 모든 종목을 어제 기준으로 호출한다.
//    오늘 행이 와도 외국인·기관계·개인 순매수가 모두 0·빈칸이면(집계 전) 쓰지 않음. 따라서 장중과 15:30~15:40에는 직전 세션까지이고
//    [전체]는 15:40 이후에 돌려야 오늘이 들어간다. ※ 15:40 이후 오늘 행이 실제로 어떻게 오는지(채워진 시각·확정 여부)는 장중 개발이라
//    관측하지 못했고 위 방어 규칙만 적용했다. 시계 모사·합성 응답으로 이 분기를 시험함(아래 검증).
//  - 합계: 최근 1·5·21 완료 세션의 순매수 합(억원) — 1D·1W·1M. 기준일 = 합계에 쓴 가장 최근 완료 세션의 날짜(종목별).
//    시가총액 대비 %는 대회종목 표가 이 억원 합과 시가총액으로 계산한다.
//  - 실행 모드(tblRunCtl 키 mode): full([전체])에서만 호출. build(빈칸·알 수 없는 값 포함)·quick·sector·analysis는 호출 없이
//    자기 표 → 정적 표 tblFlowUSeed → 빈 표(열·형식 모두) 순으로 돌려준다.
//  - 상태: 성공 행 OK, 종목 하나가 실패하면 그 행만 "오류: 사유"(값은 빈칸, 조회시각만 채움), 전부 실패하면 오류로 올려 직전 표 유지
//    (fnFallback, 열 형식은 다시 입힘 — 날짜가 일련번호로 보이지 않게). 페이지 행이 없으면 열만 있는 빈 표.
//  - 행 순서: fnPageRows() 순서(대회 종목 → 유니버스 밖). 종목코드는 텍스트.
//  검증(2026-10-01 하네스 + 별도 KisClient 독립 재계산, 요청일 공란/지정):
//    표본 40종목(문자 코드 10·신규상장 4·외국기업 2 포함) 여섯 합계·기준일 40/40 일치, 표본 11종목은 항등식까지 확인.
//    모사 시계 6가지(마감 후·15:35·장중·월요일 아침 등)가 독립 기대값과 일치, 합성 응답 17가지 경우 × 5가지 시계 불일치 0.
//    소요: 40종목 9.4초(종목당 약 0.2초) → 대회 종목 약 530개면 2분 안팎.
let
    Columns = {"종목코드", "기준일", "외국인1D억", "외국인1W억", "외국인1M억", "기관1D억", "기관1W억", "기관1M억", "상태", "조회시각"},
    NumCols = {"외국인1D억", "외국인1W억", "외국인1M억", "기관1D억", "기관1W억", "기관1M억"},
    Types = {{"종목코드", type text}, {"기준일", type date}} & List.Transform(NumCols, each {_, type number})
            & {{"상태", type text}, {"조회시각", type datetime}},

    // 실행 모드: 버튼 매크로가 tblRunCtl의 mode에 full을 쓴 실행에서만 KIS를 호출
    Mode =
        let
            T = try Excel.CurrentWorkbook(){[Name = "tblRunCtl"]}[Content] otherwise null,
            Hit = if T = null or not Table.HasColumns(T, {"키", "값"}) then {}
                  else List.Select(Table.ToRecords(T), each (try Text.Lower(Text.Trim(Text.From([키]))) otherwise null) = "mode"),
            V = if List.IsEmpty(Hit) then null else Hit{0}[값]
        in
            if V = null then "" else (try Text.Lower(Text.Trim(Text.From(V))) otherwise ""),
    Live = Mode = "full",

    // build 등(호출 없음): 자기 표 → tblFlowUSeed → 빈 표. 값 형식을 계약대로 맞추고, 칸별 변환 실패는 빈칸
    Conform = (t as table) as table =>
        let
            Sel = Table.SelectColumns(t, Columns, MissingField.UseNull),
            AsText = (v) => try (if v = null then null else Text.From(v)) otherwise null,
            AsCode = (v) => try (if v = null then null else
                        let s = Text.Upper(Text.Trim(Text.From(v))) in
                            if s = "" then null
                            else if Text.Length(s) < 6 and Text.Select(s, {"0".."9"}) = s then Text.PadStart(s, 6, "0")
                            else s) otherwise null,
            AsNum = (v) => try (if v = null or v is number then v else Number.From(v)) otherwise null,
            AsDate = (v) => try (if v = null then null else Date.From(v)) otherwise null,
            AsTime = (v) => try (if v = null then null else DateTime.From(v)) otherwise null,
            Ops = {{"종목코드", AsCode, type text}, {"기준일", AsDate, type date}, {"상태", AsText, type text}, {"조회시각", AsTime, type datetime}}
                  & List.Transform(NumCols, each {_, AsNum, type number})
        in
            Table.TransformColumns(Sel, Ops),
    HasRows = (t as any) as logical =>
        t is table and Table.HasColumns(t, Columns)
        and not Table.IsEmpty(Table.SelectRows(t, each
            let c = try Text.Trim(Text.From([종목코드])) otherwise null in c <> null and c <> "")),
    Own = try Excel.CurrentWorkbook(){[Name = "tblFlowU"]}[Content] otherwise null,
    Seed = try Excel.CurrentWorkbook(){[Name = "tblFlowUSeed"]}[Content] otherwise null,
    Kept = if HasRows(Own) then Conform(Own)
           else if HasRows(Seed) then Conform(Seed)
           else Table.TransformColumnTypes(#table(Columns, {}), Types),

    // full: 페이지 행 종목마다 종목별 투자자매매동향(일별) 1회
    Main =
        let
            Now = DateTime.LocalNow(),
            Today = Date.From(Now),
            Yesterday = Date.AddDays(Today, -1),
            Completed = Time.From(Now) >= #time(15, 30, 0),    // 오늘 세션 완료(장 마감 15:30, spec R6)
            AskToday = Time.From(Now) >= #time(15, 40, 0),     // API가 오늘 날짜를 받는 시각(그 전은 OPSQ2001)
            Path = "/uapi/domestic-stock/v1/quotations/investor-trade-by-stock-daily",
            Codes = List.Buffer(List.Distinct(fnPageRows()[종목코드])),

            Call = (code as text, day as date) as record =>
                fnKisGet(Path, "FHPTJ04160001", [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code,
                    FID_INPUT_DATE_1 = fnYmdText(day), FID_ORG_ADJ_PRC = "", FID_ETC_CLS_CODE = ""]),
            // 오늘 날짜가 받아들여지는지 첫 종목으로 한 번 확인 (OPSQ2001이면 모든 종목을 어제 기준으로)
            Probe = if AskToday and not List.IsEmpty(Codes) then try Call(Codes{0}, Today) else null,
            ProbeMsg = if Probe <> null and Probe[HasError] then Text.From(Probe[Error][Message]) else null,
            TodayRejected = ProbeMsg <> null and Text.Contains(ProbeMsg, "OPSQ2001"),
            ReqDate = if AskToday and not TodayRejected then Today else Yesterday,

            Calc = (code as text) as record =>
                let
                    J = Call(code, ReqDate),
                    O = Record.FieldOrDefault(J, "output2", null),
                    Raw = if O is list then O else if O is record then {O} else {},
                    Parsed = List.Select(List.Transform(List.Select(Raw, each _ is record), (r) => [
                            d = fnYmd(Record.FieldOrDefault(r, "stck_bsop_date", null)),
                            px = fnNum(Record.FieldOrDefault(r, "stck_clpr", null)),
                            f = fnNum(Record.FieldOrDefault(r, "frgn_ntby_tr_pbmn", null)),
                            o = fnNum(Record.FieldOrDefault(r, "orgn_ntby_tr_pbmn", null)),
                            p = fnNum(Record.FieldOrDefault(r, "prsn_ntby_tr_pbmn", null))]), each [d] <> null),
                    // 최근 날짜부터, 같은 날짜는 첫 행만
                    Desc = List.Buffer(List.Distinct(List.Sort(Parsed, {each [d], Order.Descending}), each [d])),
                    Priced = (x as record) as logical => x[px] <> null and x[px] > 0,
                    NoFlow = (x as record) as logical =>
                        (x[f] = null or x[f] = 0) and (x[o] = null or x[o] = 0) and (x[p] = null or x[p] = 0),
                    PricedAt = List.Select(List.Positions(Desc), each Priced(Desc{_})),
                    LastAt = if List.IsEmpty(PricedAt) then -1 else List.Max(PricedAt),
                    // 더 과거에 시세 없는 행이 있는 가장 오래된 시세 행 = 상장일 (순매수가 모두 0이면 투자자별 자료가 없는 날)
                    ListingDay = if LastAt >= 0 and LastAt < List.Count(Desc) - 1 and NoFlow(Desc{LastAt}) then Desc{LastAt}[d] else null,
                    // 완료 세션: 시세가 있고, 상장일이 아니고, 오늘 행은 장 마감 뒤 순매수가 채워졌을 때만
                    Sessions = List.Buffer(List.Select(Desc, (x) =>
                        Priced(x) and x[d] <= Today and x[d] <> ListingDay
                        and (x[d] < Today or (Completed and not NoFlow(x))))),
                    Sum = (k as text, n as number) as nullable number =>
                        let v = List.Transform(List.FirstN(Sessions, n), each Record.Field(_, k))
                        in if List.Count(v) < n or List.Contains(v, null) then null else List.Sum(v) / 100,
                    Rec = [기준일 = Sessions{0}[d],
                           외국인1D억 = Sum("f", 1), 외국인1W억 = Sum("f", 5), 외국인1M억 = Sum("f", 21),
                           기관1D억 = Sum("o", 1), 기관1W억 = Sum("o", 5), 기관1M억 = Sum("o", 21)],
                    Out =
                        if List.IsEmpty(PricedAt) then error "종목코드 확인: 응답에 시세와 수급이 없습니다(없는 종목이거나 상장 전)"
                        else if List.IsEmpty(Sessions) then error "완료된 세션의 수급 자료가 없습니다(상장 당일 등)"
                        else Rec,
                    // 레코드 필드는 지연 평가되므로 여기서 모든 값을 평가해 오류가 이 종목의 try 안에서 잡히게 함
                    Values = Record.FieldValues(Out),
                    Forced = List.Count(List.Select(Values, each _ = null or _ is number or _ is date)) = List.Count(Values),
                    AllBlank = List.Count(List.RemoveNulls(List.Skip(Values, 1))) = 0
                in
                    if not Forced then error "계산 값 형식 오류"
                    else if AllBlank then error "수급 값이 비어 있습니다"
                    else Out,
            One = (code as text) as record =>
                let
                    R = try Calc(code),
                    Why = if not R[HasError] then null
                          else let m = R[Error][Message], w = R[Error][Reason] in
                              if m <> null and m <> "" then Text.From(m) else if w <> null and w <> "" then Text.From(w) else "알 수 없는 오류"
                in
                    if R[HasError] then [종목코드 = code, 상태 = "오류: " & Text.Start(Why, 200), 조회시각 = Now]
                    else [종목코드 = code] & R[Value] & [상태 = "OK", 조회시각 = Now],
            Rows = List.Buffer(List.Transform(Codes, One)),
            Failed = List.Select(Rows, each Text.StartsWith([상태], "오류")),
            // 전부 실패하면 오류로 올려 직전 데이터를 유지(fnFallback), 일부만 실패하면 그 행만 오류 표시
            Checked = if not List.IsEmpty(Rows) and List.Count(Failed) = List.Count(Rows)
                      then error Error.Record("종목 수급", "모든 종목 조회 실패: " & Text.Middle(Failed{0}[상태], 4))
                      else Rows
        in
            if List.IsEmpty(Checked) then Table.TransformColumnTypes(#table(Columns, {}), Types)
            else Table.TransformColumnTypes(Table.FromRecords(Checked, Columns, MissingField.UseNull), Types),

    Result = if Live then Main else Kept,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지.
    // fnFallback 결과는 형식이 없는 표라 날짜 열이 일련번호(일반 서식)로 적재되므로 Conform으로 열 형식을 다시 입힘(값은 그대로)
    Tried = try Table.Buffer(Result),
    Final = if Tried[HasError] then Conform(fnFallback("tblFlowU", Columns, Tried[Error][Message])) else Tried[Value]
in
    Final
