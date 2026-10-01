// T_CSL : 신용잔고·공매도·대차 주간 지표 (spec R10·R20, 2026-10-01 T10) → tblCSL
//  - 버튼 전용(주간) 쿼리: 실행 모드(tblRunCtl mode)가 full일 때만 KIS를 호출한다. 그 밖(quick·sector·analysis·build·빈칸·모르는 값)은
//    호출 없이 자기 표(tblCSL) → 없거나 비었으면 정적 표 tblCSLSeed → 그것도 없으면 열 형식을 갖춘 빈 표를 그대로 돌려준다.
//  - 대상 = fnPageRows() 행 전체(대회 종목 ∪ 보유·관심). full에서는 표를 페이지 행과 같은 종목·순서로 새로 쓴다(빠진 종목 행은 버림).
//  - 주간 규칙(spec R20): 직전 표(자기 표 → tblCSLSeed)에 그 종목 행이 있고 아래를 모두 만족하면 그 행을 그대로 둔다(호출 없음).
//      ① 상태가 `오류`·`이전 데이터`로 시작하지 않음(= OK 또는 데이터 없음)  ② 조회시각이 있음
//      ③ 오늘 날짜 − 조회시각 날짜 < weekly_days일(설정, 기본 7. 날짜 차이로 세어 매일 같은 시각에 돌려도 7일째에 조회됨)
//      ④ 설정 force_weekly가 Y가 아님(기본 N. 실행 뒤 N으로 되돌리는 일은 버튼 매크로 몫)
//    새 종목(직전 표에 없음)·오류였던 종목은 항상 조회한다. 조회시각 = 마지막으로 성공한 조회 시각(한 번의 실행은 모두 같은 시각).
//  - 종목당 호출 3회(연속 조회 없음; 실측 호출당 약 0.15초 → 22종목 66회 9.9초, 540종목이면 약 4분), 앞 호출이 실패하면 나머지는 부르지 않는다:
//      ① 신용잔고 일별추이 [국내주식-110] FHPST04760000 /uapi/domestic-stock/v1/quotations/daily-credit-balance
//         FID_COND_MRKT_DIV_CODE=J, FID_COND_SCR_DIV_CODE=20476, FID_INPUT_ISCD=종목코드, FID_INPUT_DATE_1=오늘(결제일자 이하 의미)
//         → output[] 최대 30건, 매매일자(deal_date) 내림차순. 최신 행은 약 2영업일 늦음(T+2 공표: 2026-10-01 낮 조회 시 최신 = 9/28)
//      ② 공매도 일별추이 [국내주식-134] FHPST04830000 /uapi/domestic-stock/v1/quotations/daily-short-sale
//         FID_COND_MRKT_DIV_CODE=J, FID_INPUT_ISCD=종목코드, FID_INPUT_DATE_1=오늘−30일, FID_INPUT_DATE_2=오늘
//         → output2[] 일자(stck_bsop_date) 내림차순. 이 구간에서 약 20행(최대 100행 제한 안). output1(현재가 요약)은 쓰지 않음
//      ③ 종목별 일별 대차거래추이 [국내주식-135] HHPST074500C0 /uapi/domestic-stock/v1/quotations/daily-loan-trans
//         MRKT_DIV_CLS_CODE=3(종목), MKSC_SHRN_ISCD=종목코드, START_DATE=오늘−60일, END_DATE=오늘, CTS=빈값
//         → output1[] 일자(bsop_date) 내림차순, 이 구간에서 약 40행(최대 100행 제한 안)
//  - 응답 필드 → 지표 (2026-10-01 실측. 비율은 모두 소수: 1.23% → 0.0123, %p 변화는 소수 차이)
//      신용잔고율        = whol_loan_rmnd_rate(전체 융자 잔고 비율, % 소수 둘째 자리) ÷ 100 — 최신 행(deal_date 최대)
//      신용잔고율1M변화  = (최신 행 − 21세션 전 행(내림차순 22번째 행)의 whol_loan_rmnd_rate) ÷ 100. 행이 22개 미만이면 빈칸
//                          (잔고율은 응답이 소수 둘째 자리 %라 변화의 최소 단위는 0.0001 = 0.01%p. 응답 비율의 분모(상장주식수)는 날짜마다
//                          그때의 값이라 증자로 주식 수가 바뀐 종목은 비율 자체가 그 영향을 포함한다 — 응답 값을 그대로 쓴다)
//      공매도비중5일     = 완료 세션 최근 5개 각각의 ssts_tr_pbmn(공매도 거래대금, 원) ÷ acml_tr_pbmn(그날 총 거래대금, 원)의 단순 평균.
//                          응답의 직접 비율 필드 ssts_tr_pbmn_rlim(%, 소수 둘째 자리)은 이 몫의 반올림값과 같다(2026-10-01 실측: 10종목 362행
//                          완료 세션 전부 일치, 차이 < 0.005%p) → 반올림 오차 없이 같은 몫을 원래 정밀도로 계산해 쓴다.
//                          완료 세션 = 일자 < 오늘(아래 참고). 총 거래대금이 0·빈 행(거래정지 등)은 세션으로 세지 않고 건너뜀. 5개 미만이면 빈칸
//      대차잔고1M변화율  = 최신 행 rmnd_stcn(당일 잔고 주수) ÷ 21세션 전 행(내림차순 22번째 행)의 rmnd_stcn − 1. 행이 22개 미만이거나
//                          그때 잔고가 0·빈 값이면 빈칸. 주수 기준으로 일관되게 계산한다(rmnd_amt 잔고 금액은 백만원 단위이고
//                          주수 × 그날 종가라서 주가 변동이 섞이므로 쓰지 않음. 분할·증자로 주식 수가 바뀐 달은 변화율이 왜곡될 수 있음)
//      기준일            = 지표를 낸 계열(신용·공매도·대차)의 최신 사용 행 날짜 중 가장 이른 날짜 = 모든 값이 이 날짜 이후까지 반영됨.
//                          (보통 신용잔고가 가장 늦어 신용 최신일 = 기준일, 공매도·대차는 1~2세션 더 최신). 지표가 하나도 없으면 빈칸
//  - '완료 세션'만 쓴다: 공매도 응답은 오늘 행이 장중·공표 전에 값 0(거래대금만 장중 누적)으로 들어 있어(2026-10-01 실측, 전 종목 0.00) 그대로 평균하면
//    비중이 낮게 나온다. 대차도 같은 기준을 적용한다. 장 마감 뒤에 돌려도 오늘 행은 쓰지 않는다(1세션 늦을 뿐 결과가 일정함).
//  - 응답 정리: 일자가 없는 행은 버린다(공매도·대차 응답은 없는 종목·상장 직후 종목에서 빈 레코드 {}로 채워 줌). 같은 일자가 겹치면 응답 순서상 앞 행.
//    없는 종목코드·상장 직후 종목은 rt_cd 0에 빈 목록이 와서 오류가 아니다 → 그 계열의 지표만 빈칸. 지표가 모두 비면 상태 `데이터 없음`.
//  - 상태: OK(하나 이상 지표 있음) / 데이터 없음(조회는 성공했으나 지표 없음) / 오류: 사유(그 종목 조회 실패 — 직전 값·직전 조회시각을 유지) /
//    이전 데이터(갱신 실패: 사유)(fnFallback — 조회한 종목이 전부 실패해 표 전체를 직전 값으로 유지).
//  - 종목별 실패(API 오류·네트워크)는 그 종목 행만 `오류: <사유>`로 하고 직전 값은 그대로 둔다. 조회를 시도한 종목이 하나 이상인데 사용할 수 있는 행
//    (OK·데이터 없음)이 하나도 없으면 오류로 올려 fnFallback이 직전 표를 유지하게 한다. 빈값은 빈칸(0으로 채우지 않음).
//  - 같은 [전체] 실행 안에서 다시 평가되면(자기 표 tblCSL의 가장 늦은 조회시각이 tblRunCtl started 이후) 조회하지 않고 자기 표를 돌려준다
//    — 새로 고침 직후 Excel의 백그라운드 재평가가 force_weekly·오류 재시도 종목을 또 조회하는 것을 막음(아래 Out 주석). started가 비면 확인 안 함.
let
    Cols = {"종목코드", "기준일", "신용잔고율", "신용잔고율1M변화", "공매도비중5일", "대차잔고1M변화율", "상태", "조회시각"},
    Types = {{"종목코드", type text}, {"기준일", type date}, {"신용잔고율", type number}, {"신용잔고율1M변화", type number},
             {"공매도비중5일", type number}, {"대차잔고1M변화율", type number}, {"상태", type text}, {"조회시각", type datetime}},
    Schema = type table [종목코드 = text, 기준일 = nullable date, 신용잔고율 = nullable number, 신용잔고율1M변화 = nullable number,
                         공매도비중5일 = nullable number, 대차잔고1M변화율 = nullable number, 상태 = nullable text, 조회시각 = nullable datetime],
    Main =
        let
            Now = DateTime.LocalNow(),
            Today = Date.From(Now),
            TodayTxt = fnYmdText(Today),
            ShortFrom = fnYmdText(Date.AddDays(Today, -30)),
            LoanFrom = fnYmdText(Date.AddDays(Today, -60)),
            PathCr = "/uapi/domestic-stock/v1/quotations/daily-credit-balance",
            PathSh = "/uapi/domestic-stock/v1/quotations/daily-short-sale",
            PathLn = "/uapi/domestic-stock/v1/quotations/daily-loan-trans",

            // 직전 표의 셀 값은 형식이 일정하지 않으므로(날짜가 일시·일련번호·텍스트로 올 수 있음) 오류 없이 읽는다
            SafeTxt = (v as any) as nullable text =>
                let t = try Text.Trim(Text.From(v)) otherwise null in if t = null or t = "" then null else t,
            SafeNum = (v as any) as nullable number =>
                let n = try (if v is number then v else fnNum(v)) otherwise null in if n = null or Number.IsNaN(n) then null else n,
            SafeDate = (v as any) as nullable date => try Date.From(v) otherwise null,
            SafeDt = (v as any) as nullable datetime => try DateTime.From(v) otherwise null,

            // 실행 제어 tblRunCtl(키, 값): mode — 값이 없거나 알 수 없으면 build, 호출은 full에서만. started — 버튼 실행 시작 시각(매크로가 씀)
            RunCtl = try Excel.CurrentWorkbook(){[Name = "tblRunCtl"]}[Content] otherwise null,
            CtlValue = (key as text) as any =>
                if RunCtl = null or not Table.HasColumns(RunCtl, {"키", "값"}) then null
                else
                    let hit = Table.SelectRows(RunCtl, each (let k = SafeTxt([키]) in k <> null and Text.Lower(k) = key))
                    in if Table.RowCount(hit) = 0 then null else hit{0}[값],
            ModeTxt = SafeTxt(CtlValue("mode")),
            Mode = if ModeTxt = null then "build" else Text.Lower(ModeTxt),

            // 설정(없으면 기본값): weekly_days 7, force_weekly N
            S = try fnSettings() otherwise [],
            WeeklyDays = let n = SafeNum(Record.FieldOrDefault(S, "weekly_days", null)) in if n = null or n < 0 then 7 else n,
            ForceWeekly = let t = SafeTxt(Record.FieldOrDefault(S, "force_weekly", null)) in t <> null and Text.Upper(t) = "Y",

            // 직전 표: 자기 표(tblCSL) → 비었거나 없으면 tblCSLSeed. 종목코드가 빈 행(빈 표의 자리 행 등)은 없는 행으로 본다
            Own = try Excel.CurrentWorkbook(){[Name = "tblCSL"]}[Content] otherwise null,
            Seed = try Excel.CurrentWorkbook(){[Name = "tblCSLSeed"]}[Content] otherwise null,
            AsPrev = (t as nullable table) as list =>
                if t = null or not Table.HasColumns(t, "종목코드") then {}
                else List.Select(
                    List.Transform(Table.ToRecords(t), (r) => [
                        code = let c = SafeTxt(Record.FieldOrDefault(r, "종목코드", null)) in if c = null then null else Text.Upper(c),
                        asof = SafeDate(Record.FieldOrDefault(r, "기준일", null)),
                        credit = SafeNum(Record.FieldOrDefault(r, "신용잔고율", null)),
                        chg = SafeNum(Record.FieldOrDefault(r, "신용잔고율1M변화", null)),
                        short = SafeNum(Record.FieldOrDefault(r, "공매도비중5일", null)),
                        loan = SafeNum(Record.FieldOrDefault(r, "대차잔고1M변화율", null)),
                        status = SafeTxt(Record.FieldOrDefault(r, "상태", null)),
                        ts = SafeDt(Record.FieldOrDefault(r, "조회시각", null))]),
                    each _[code] <> null),
            OwnRows = AsPrev(Own),
            PrevRows = if List.Count(OwnRows) > 0 then OwnRows else AsPrev(Seed),
            // 같은 버튼 실행 안의 재평가 확인: 자기 표(Seed는 보지 않음)의 가장 늦은 조회시각이 이번 실행의 started 이후면 이미 조회한 것
            Started = SafeDt(CtlValue("started")),
            OwnLast = List.Max(List.RemoveNulls(List.Transform(OwnRows, each _[ts]))),
            Surveyed = Started <> null and OwnLast <> null and OwnLast >= Started,
            // 같은 종목코드가 여러 행이면 처음 행(Record.FromList는 이름이 겹치면 오류라서 하나씩 더함)
            PrevByCode = List.Accumulate(PrevRows, [], (st, r) => if Record.HasFields(st, r[code]) then st else Record.AddField(st, r[code], r)),
            PrevList = List.Transform(Record.FieldNames(PrevByCode), each Record.Field(PrevByCode, _)),

            // 주간 규칙
            Healthy = (st as nullable text) as logical =>
                st <> null and not Text.StartsWith(st, "오류") and not Text.StartsWith(st, "이전 데이터"),
            Stale = (ts as nullable datetime) as logical =>
                ts = null or (let days = Duration.Days(Today - Date.From(ts)) in days < 0 or days >= WeeklyDays),
            NeedsFetch = (p as nullable record) as logical =>
                ForceWeekly or p = null or not Healthy(p[status]) or Stale(p[ts]),

            // 응답 목록 정리·오류 사유
            Items = (o as any) as list =>
                if o is list then List.Select(o, each _ is record) else if o is record then {o} else {},
            // 오류 메시지는 한 줄로(셀 안 줄바꿈 방지)
            Reason = (e as record) as text =>
                let m = try Text.Trim(Text.Replace(Text.Replace(Text.Replace(Text.From(e[Message]), "#(cr)", " "), "#(lf)", " "), "#(tab)", " ")) otherwise ""
                in if m = "" then "조회 실패" else m,

            // 세 응답 → 지표 한 줄. 계열마다 일자 내림차순·완료 세션만 쓰고, 지표를 낸 계열의 최신 날짜 중 가장 이른 날짜가 기준일
            Compute = (cr as list, sh as list, ln as list) as record =>
                let
                    // 신용잔고: 잔고율(%) → 소수. 행 위치 = 세션 위치(응답이 세션마다 한 행: 5종목 30행 모두 거래일 연속 확인)
                    CrT = #table(type table [d = nullable date, rate = nullable number],
                        List.Transform(cr, (r) => {fnYmd(Record.FieldOrDefault(r, "deal_date", null)),
                            fnNum(Record.FieldOrDefault(r, "whol_loan_rmnd_rate", null))})),
                    CrS = Table.Buffer(Table.Sort(Table.Distinct(Table.SelectRows(CrT, each [d] <> null), {"d"}), {{"d", Order.Descending}})),
                    Rate0 = if Table.RowCount(CrS) >= 1 then CrS[rate]{0} else null,
                    Rate21 = if Table.RowCount(CrS) >= 22 then CrS[rate]{21} else null,
                    Credit = if Rate0 = null then null else Number.Round(Rate0 / 100, 8),
                    CreditChg = if Rate0 = null or Rate21 = null then null else Number.Round((Rate0 - Rate21) / 100, 8),
                    CrAsOf = if Credit = null then null else CrS[d]{0},

                    // 공매도: 완료 세션(일자 < 오늘) 중 총 거래대금이 있는 최근 5개의 (공매도 거래대금 ÷ 총 거래대금) 평균
                    ShT = #table(type table [d = nullable date, s = nullable number, t = nullable number],
                        List.Transform(sh, (r) => {fnYmd(Record.FieldOrDefault(r, "stck_bsop_date", null)),
                            fnNum(Record.FieldOrDefault(r, "ssts_tr_pbmn", null)), fnNum(Record.FieldOrDefault(r, "acml_tr_pbmn", null))})),
                    ShS = Table.Buffer(Table.Sort(Table.Distinct(Table.SelectRows(ShT, each [d] <> null and [d] < Today), {"d"}), {{"d", Order.Descending}})),
                    ShOk = Table.SelectRows(ShS, each [s] <> null and [t] <> null and [t] > 0 and [s] >= 0),
                    Sh5 = Table.FirstN(ShOk, 5),
                    Short5 = if Table.RowCount(Sh5) < 5 then null else List.Average(List.Transform(Table.ToRecords(Sh5), each [s] / [t])),
                    ShAsOf = if Short5 = null then null else Sh5[d]{0},

                    // 대차: 완료 세션의 잔고 주수 — 최신 ÷ 21세션 전 − 1
                    LnT = #table(type table [d = nullable date, b = nullable number],
                        List.Transform(ln, (r) => {fnYmd(Record.FieldOrDefault(r, "bsop_date", null)),
                            fnNum(Record.FieldOrDefault(r, "rmnd_stcn", null))})),
                    LnS = Table.Buffer(Table.Sort(Table.Distinct(Table.SelectRows(LnT, each [d] <> null and [d] < Today), {"d"}), {{"d", Order.Descending}})),
                    Bal0 = if Table.RowCount(LnS) >= 1 then LnS[b]{0} else null,
                    Bal21 = if Table.RowCount(LnS) >= 22 then LnS[b]{21} else null,
                    LoanChg = if Bal0 = null or Bal21 = null or Bal21 <= 0 then null else Bal0 / Bal21 - 1,
                    LnAsOf = if LoanChg = null then null else LnS[d]{0},

                    AsOfs = List.RemoveNulls({CrAsOf, ShAsOf, LnAsOf}),
                    AsOf = if List.IsEmpty(AsOfs) then null else List.Min(AsOfs),
                    AnyData = Credit <> null or CreditChg <> null or Short5 <> null or LoanChg <> null
                in
                    [asof = AsOf, credit = Credit, chg = CreditChg, short = Short5, loan = LoanChg, status = if AnyData then "OK" else "데이터 없음"],

            // 종목 하나: 세 호출(앞이 실패하면 나머지는 건너뜀). 레코드·목록은 지연 평가되므로 응답을 List.Buffer로 여기서 강제 평가해
            // 호출하는 쪽의 행 단위 try가 API 오류를 잡게 하고, 계산 결과도 Record.ToList로 강제 평가한다
            Fetch = (code as text) as record =>
                let
                    Cr = try List.Buffer(Items(Record.FieldOrDefault(
                        fnKisGet(PathCr, "FHPST04760000", [FID_COND_MRKT_DIV_CODE = "J", FID_COND_SCR_DIV_CODE = "20476",
                            FID_INPUT_ISCD = code, FID_INPUT_DATE_1 = TodayTxt]), "output", {}))),
                    Sh = if Cr[HasError] then Cr else try List.Buffer(Items(Record.FieldOrDefault(
                        fnKisGet(PathSh, "FHPST04830000", [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code,
                            FID_INPUT_DATE_1 = ShortFrom, FID_INPUT_DATE_2 = TodayTxt]), "output2", {}))),
                    Ln = if Sh[HasError] then Sh else try List.Buffer(Items(Record.FieldOrDefault(
                        fnKisGet(PathLn, "HHPST074500C0", [MRKT_DIV_CLS_CODE = "3", MKSC_SHRN_ISCD = code,
                            START_DATE = LoanFrom, END_DATE = TodayTxt, CTS = ""]), "output1", {}))),
                    Failed =
                        if Cr[HasError] then "신용잔고 조회 실패: " & Reason(Cr[Error])
                        else if Sh[HasError] then "공매도 조회 실패: " & Reason(Sh[Error])
                        else if Ln[HasError] then "대차 조회 실패: " & Reason(Ln[Error])
                        else null,
                    Row = if Failed <> null then null else Compute(Cr[Value], Sh[Value], Ln[Value]),
                    Forced = if Failed <> null then null else Record.ToList(Row)
                in
                    if Failed <> null then error Failed
                    else if List.Count(Forced) < 0 then null
                    else Row,

            // full: 대상 종목과 그 종목의 조회 계획·결과 → 새 표. 함수로 두어 full이고 이번 실행에서 아직 조회하지 않았을 때만 부른다(함수 값은
            // 평가해도 호출하지 않음). 조회 결과(r)는 필요한 종목만, 종목마다 한 번만 평가된다(버퍼로 고정된 레코드의 필드는 처음 읽을 때 한 번 평가)
            Survey = () as table =>
                let
                    Codes = List.Buffer(List.Distinct(List.RemoveNulls(List.Transform(fnPageRows()[종목코드],
                                each let t = SafeTxt(_) in if t = null then null else Text.Upper(t))))),
                    Plan = List.Buffer(List.Transform(Codes, (c) =>
                        let p = Record.FieldOrDefault(PrevByCode, c, null), nd = NeedsFetch(p)
                        in [code = c, prev = p, need = nd, r = if nd then try Fetch(c) else null])),
                    RowOf = (x as record) as list =>
                        let
                            c = x[code],
                            p = x[prev],
                            PV = (k as text) as any => if p = null then null else Record.Field(p, k),
                            f = x[r]
                        in
                            if not x[need] then {c, PV("asof"), PV("credit"), PV("chg"), PV("short"), PV("loan"), PV("status"), PV("ts")}
                            else if f[HasError] then
                                {c, PV("asof"), PV("credit"), PV("chg"), PV("short"), PV("loan"), "오류: " & Text.Start(Reason(f[Error]), 120), PV("ts")}
                            else
                                let v = f[Value] in {c, v[asof], v[credit], v[chg], v[short], v[loan], v[status], Now},
                    Rows = List.Buffer(List.Transform(Plan, RowOf)),
                    Attempted = List.Count(List.Select(Plan, each _[need])),
                    FailedPlan = List.Select(Plan, each _[need] and _[r][HasError]),
                    Usable = List.Count(List.Select(Rows, each Healthy(_{6}))),
                    // 조회한 종목이 하나 이상인데 쓸 수 있는 행이 없으면(모두 실패) 오류로 올려 직전 표를 유지(fnFallback)
                    Checked = if Attempted > 0 and Usable = 0 and List.Count(FailedPlan) > 0 then error Reason(FailedPlan{0}[r][Error]) else Rows
                in
                    #table(Schema, Checked),

            PrevOut = #table(Schema, List.Transform(PrevList, (p) => {p[code], p[asof], p[credit], p[chg], p[short], p[loan], p[status], p[ts]})),
            // Excel은 새로 고침이 끝난 직후 쿼리를 백그라운드에서 한 번 더 평가한다(2026-10-01 하네스 + 모의 서버 실측, 동기·백그라운드 새로 고침
            // 모두). 이 평가는 방금 적재한 자기 표를 다시 읽으므로 주간 규칙상 조회할 종목이 없으면 호출이 없지만, force_weekly = Y(모든 종목)나
            // 오류 재시도 종목은 Power Query 캐시에 없는 요청을 한 번 더 보냈다 → 위 Surveyed 확인으로 막는다.
            // (조회한 종목이 모두 실패해 직전 표로 돌아간 경우엔 조회시각이 그대로라 막지 못하고 한 번 더 시도한다.)
            // 표를 처음 적재한 뒤·통합문서를 연 뒤 처음 하는 새로 고침에서는 추가 평가가 자기 표가 갱신되기 전에 함께 돌아 이 확인을 통과하므로
            // 그 한 번은 호출이 최대 2배가 될 수 있다(실측) — 버튼 매크로가 그 표를 build 모드로 먼저 한 번 새로 고침해 두면 생기지 않는다.
            Out = if Mode = "full" and not Surveyed then Survey() else PrevOut
        in
            Table.TransformColumnTypes(Out, Types),
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지.
    // 직전 표를 Excel에서 다시 읽으면 열 형식이 없어(any) Excel이 기준일·조회시각 셀 서식을 '일반'으로 바꾸므로 같은 형식을 다시 지정(값은 그대로)
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblCSL", Cols, Tried[Error][Message]), Types) else Tried[Value]
in
    Final
