// T_Fin : 분기 실적 저장소(분기 단독값) — 후행 PER·PBR·실적 YoY·실적 상태의 원천 (2026-10-01) → tblFin
//
//  열(정확히 이 순서): 종목코드, 결산년월, 공개기준일, 매출, 영업이익, 순이익, EPS, BPS, ROE, 상태, 조회시각
//    종목코드  text      6자리(앞 0 유지, 영문 포함 가능 예: 0126Z0)
//    결산년월  text      "yyyymm" = 그 분기의 결산월(분기말 월). 행 = (종목코드, 결산년월)마다 1행, 종목마다 최근 12개 분기
//    공개기준일 date     그 분기 실적이 '알려진' 첫날. 1~3분기 = 분기말(그 달 말일) + 45일, 4분기(회계연도 결산 분기) = 결산일 + 90일
//                        (설정 fin_lag_q·fin_lag_y로 조정 가능, 없으면 45·90). 기준일 d에 알려진 행 = 공개기준일 <= d.
//                        (공개기준일 − 결산월 말일 = 90일(fin_lag_y)인 행이 4분기 — 표에 회계 분기 번호 열이 따로 없음)
//    매출·영업이익·순이익  number  억원, 분기 단독값
//    EPS       number    원, 분기 단독값(회계연도 누적 EPS의 차이) → TTM EPS = 연속된 최근 4개 분기 EPS 합
//    BPS       number    원, 그 분기말 시점 값(차감하지 않음)
//    ROE       number    소수(API의 % 값 ÷ 100, 예: 31.39% → 0.3139), 그 행의 API 값 그대로(4분기 행 = 연간 ROE와 같음,
//                        1~3분기 행의 계산 방식은 KIS 내부 — 차감·환산하지 않음)
//    상태      text      OK / 데이터 없음 / 오류: <사유> / 이전 데이터(갱신 실패: <사유>)
//    조회시각  datetime  그 종목 행을 KIS에서 받은 시각(같은 종목의 행은 모두 같은 값)
//  정렬: fnPageRows 순서(대회 종목 → 유니버스 밖) → 종목 안에서는 결산년월 내림차순(첫 행 = 최근 분기)
//
//  원천(종목당 2회 호출, 조회 전용, FID_DIV_CLS_CODE=1 분기, fid_cond_mrkt_div_code=J):
//    국내주식 손익계산서 [v1_국내주식-079] FHKST66430200 /uapi/domestic-stock/v1/finance/income-statement
//      sale_account → 매출, bsop_prti → 영업이익, thtr_ntin → 순이익(연결 당기순이익 — 비지배지분 포함)
//    국내주식 재무비율   [v1_국내주식-080] FHKST66430300 /uapi/domestic-stock/v1/finance/financial-ratio
//      eps → EPS(지배주주 기준으로 보임), bps → BPS, roe_val → ROE
//    응답은 최신 분기부터 최대 30개 분기(약 7.5년)를 한 번에 줌(연속 조회 tr_cont 없음).
//
//  실측 사실(2026-10-01, 005930·000660·105560·000810·035420·950250·004310·0126Z0·950160·195940·087010·064760·323410·058470·0220W0):
//    · 분기 값은 회계연도 누적: 매출·영업이익·순이익과 EPS 모두. 예) 삼성전자 매출 202503 791,405 → 202506 1,537,068 →
//      202509 2,397,686 → 202512 3,336,059(= 연간), EPS 1,186 → 1,920 → 3,701 → 6,564(= 연간 EPS). BPS는 시점 값.
//    · 단위: 금액 억원(삼성전자 2024년 매출 3,008,709 = 300.9조), EPS·BPS 원, roe_val %.
//    · 빈 값 표시: "99.99"(쓰지 않는 항목 자리) 또는 0.00. 자료가 없는 분기는 모든 항목이 0.00(테라뷰 202501), 재무비율의
//      bps·roe_val 0.00도 빈 값(0126Z0 roe_val 전부 0.00, 테라뷰 202510 bps 0.00)
//      → 이 쿼리는 "", 99.99, 0을 모두 빈칸(null)으로 본다(금액이 정확히 0억인 진짜 값도 빈칸 — 0을 '값'으로 남기지 않음).
//    · KIS 재무 자료가 없는 종목이 꽤 있다: 모든 분기가 0.00(현대약품 004310·카카오뱅크 323410·리노공업 058470),
//      자료가 아직 없는 신규 상장(0220W0)과 없는 종목코드는 rt_cd 0에 0행 → 모두 '데이터 없음'.
//    · 분기 자료가 끊기고 연간 행만 이어지는 종목(HK이노엔 195940: 202209 다음에 202212·202312·202412),
//      연간(12월) 행만 있는 종목(코오롱티슈진 950160·펩트론 087010·티씨케이 064760, 최신 202412)이 있다
//      → 직전 분기가 없어 분기 단독값·TTM EPS는 빈칸(BPS·ROE만 남음).
//    · 12월 결산이 아닌 종목: 테라뷰(950250) 결산월 1·4·7·10월(4월 결산, 202507에서 누적이 다시 시작),
//      현대약품 2·5·8·11월(11월 결산, 값은 전부 0.00).
//    · 호출 1회 약 0.1~0.2초(시험: 10종목 20회 2.6초) → 대회 종목 527개 주간 전체 조회 ≈ 1,054회, 약 2~3분.
//
//  분기 단독값(매출·영업이익·순이익·EPS): 1분기 = 누적값 그대로, 2~4분기 = 누적값 − 같은 회계연도 직전 분기(정확히 3개월 전
//    결산월) 누적값. 직전 분기 행이 없거나 그 값이 빈칸이면 빈칸(추정하지 않음).
//  회계연도 결산월(4분기) 판정 — 응답에 결산월 정보가 없어 누적값이 '다시 작아지는' 지점으로 정한다(추가 호출 없음):
//    ① 3개월 간격으로 이어진 두 행에서 누적 매출이 줄면 뒤 행의 월을 '1분기 월' 후보로 한 표씩 센다(누적 매출은 회계연도 안에서는
//       늘기만 하므로 줄어든 곳 = 새 회계연도 시작). 매출이 없으면(빈칸) 영업이익·순이익의 절댓값이 같은 부호에서 줄어든 곳으로 센다.
//    ② 표가 가장 많은 월이 1분기 월(같으면 가장 최근에 나온 월), 결산월 = 1분기 월 − 3개월.
//    ③ 표가 하나도 없으면: 자료 행이 2개 이상이고 모두 같은 월이면(연간 행만 있는 종목) 그 월, 결산월 체계가 3·6·9·12월이면 12월,
//       그 밖은 판정 불가 → 분기 단독값 전부 빈칸, 공개기준일은 보수적으로 + 90일.
//  최근 12개 분기: 값이 하나라도 있는 가장 최근 분기부터 응답에 있는 분기 12개(중간에 값이 전부 빈 분기도 행으로 남김 — T_Company는
//    결산년월이 3개월씩 이어지는지로 연속성을 확인). 응답에 아예 없는 분기는 만들지 않는다.
//  데이터 없음: API가 두 응답 모두 0행이거나, 준 행이 전부 빈 값(0.00·99.99)이면 그 종목은 결산년월이 빈 1행, 상태 "데이터 없음".
//
//  실행 모드(tblRunCtl 키 mode)와 주간 규칙:
//    · mode = full 일 때만 호출한다. 그 밖(quick·sector·analysis·build·빈칸·알 수 없는 값)은 호출 없이 직전 tblFin을 그대로
//      돌려주고(값 형식만 정리), 없으면 정적 표 tblFinSeed, 그것도 없으면 빈 표.
//    · full: 대상 종목 = fnPageRows()만. 종목마다 아래 중 하나면 조회, 아니면 직전 행을 그대로 둔다.
//        설정 force_weekly = Y / 직전 표에 그 종목이 없음(새 종목) / 그 종목 상태가 "오류"로 시작 / 자료 행 없이 상태가
//        "데이터 없음"이 아님 / 오늘 날짜 − 그 종목 조회시각 날짜 >= 설정 weekly_days(기본 7, 달력 일수)
//      → 두 번째 [전체]가 weekly_days 안이면 호출 0회. force_weekly를 N으로 되돌리는 것은 매크로 몫.
//    · 대상에서 빠진 종목의 행은 지운다(표 = 현재 페이지 종목만). 다시 들어오면 새 종목으로 조회.
//    · 조회 실패(종목 단위, try로 강제 평가): 직전 행이 있으면 그 행들을 그대로 두고 상태만 "오류: 사유"(조회시각은 원래 값 →
//      다음 [전체]에서 다시 조회), 없으면 결산년월이 빈 1행 "오류: 사유". 조회한 종목이 전부 실패하면 오류로 올려
//      fnFallback이 직전 표 전체를 "이전 데이터(갱신 실패: 사유)"로 유지. 그 다음 [전체]에서 조회하지 않은 종목의
//      "이전 데이터" 상태는 OK로 되돌린다(자료 자체는 유효 — 단계 판정이 계속 실패로 보지 않도록). 결산년월이 빈 행은
//      원래 '데이터 없음'이었는지 '오류'였는지 알 수 없으므로 그 종목은 다시 조회한다(종목당 2회).
//    · 지금 시각 = tblRunCtl now_override(시험 전용)가 있으면 그 값, 없으면 DateTime.LocalNow(). 주간 판정과 조회시각에 쓴다.
//    · 같은 [전체] 실행 안에서 다시 평가되면(자기 표 tblFin의 가장 늦은 조회시각이 tblRunCtl started 이후) 조회하지 않고 자기 표를
//      돌려준다 — 새로 고침 직후 Excel의 백그라운드 재평가가 force_weekly·오류 재시도 종목을 또 조회하는 것을 막음(아래 Main 주석).
//      started가 비었거나 now_override를 쓰는 시험(조회시각이 시험 시각이라 비교가 맞지 않음)에서는 확인하지 않는다.
//  설정(tblSettings, 없으면 기본값): weekly_days 7, force_weekly N, fin_lag_q 45, fin_lag_y 90
//    (공개기준일 설정을 바꾸면 그 종목을 다음에 조회할 때 반영 — 바로 모두 반영하려면 force_weekly = Y로 [전체]).
let
    Cols = {"종목코드", "결산년월", "공개기준일", "매출", "영업이익", "순이익", "EPS", "BPS", "ROE", "상태", "조회시각"},
    Types = {{"종목코드", type text}, {"결산년월", type text}, {"공개기준일", type date}, {"매출", type number},
             {"영업이익", type number}, {"순이익", type number}, {"EPS", type number}, {"BPS", type number},
             {"ROE", type number}, {"상태", type text}, {"조회시각", type datetime}},
    Main =
        let
            // ---------------------------------------------------------------- 공통 도우미
            Txt = (v as any) as nullable text =>
                if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,
            // 종목코드 정규화(fnPageRows·fnClassify와 같은 규칙: 대문자, 숫자만 6자리 미만이면 앞 0 채움)
            Code = (v as any) as nullable text =>
                let t0 = Txt(try v otherwise null), t = if t0 = null then null else Text.Upper(t0)
                in if t = null then null
                   else if Text.Length(t) < 6 and Text.Select(t, {"0".."9"}) = t then Text.PadStart(t, 6, "0") else t,
            // 결산년월 → "yyyymm" (숫자 202606·날짜·"2026-06" 형식도 받아 줌 — 정적 Seed 표가 숫자로 들어올 수 있음)
            YmText = (v as any) as nullable text =>
                let
                    x = try v otherwise null,
                    t = if x = null then null
                        else if x is number then Number.ToText(Number.Round(x), "0", "en-US")
                        else if x is date or x is datetime then Date.ToText(Date.From(x), [Format = "yyyyMM", Culture = "en-US"])
                        else Text.Select(Text.From(x), {"0".."9"}),
                    m = if t <> null and Text.Length(t) = 6 then Number.From(Text.End(t, 2)) else null
                in
                    if m <> null and m >= 1 and m <= 12 then t else null,
            NumOrNull = (v as any) as nullable number => if v = null then null else try Number.From(v) otherwise null,
            DateOrNull = (v as any) as nullable date => if v = null then null else try Date.From(v) otherwise null,
            DateTimeOrNull = (v as any) as nullable datetime => if v = null then null else try DateTime.From(v) otherwise null,

            // ---------------------------------------------------------------- 실행 제어·설정
            RunCtl = try Excel.CurrentWorkbook(){[Name = "tblRunCtl"]}[Content] otherwise null,
            CtlValue = (key as text) as any =>
                if RunCtl = null or not Table.HasColumns(RunCtl, {"키", "값"}) then null
                else
                    let hit = Table.SelectRows(RunCtl, (r) => Txt(r[키]) = key)
                    in if Table.RowCount(hit) = 0 then null else hit{0}[값],
            Mode = let m = Txt(CtlValue("mode")) in if m = null then "build" else Text.Lower(m),
            NowOverride = DateTimeOrNull(CtlValue("now_override")),
            Now = if NowOverride <> null then NowOverride else DateTime.LocalNow(),
            Today = Date.From(Now),

            Settings = try fnSettings() otherwise [],
            SetNum = (key as text, default as number) as number =>
                let n = NumOrNull(Record.FieldOrDefault(Settings, key, null))
                in if n = null or n < 0 then default else n,
            WeeklyDays = SetNum("weekly_days", 7),
            ForceWeekly = let v = Txt(Record.FieldOrDefault(Settings, "force_weekly", null)) in v <> null and Text.Upper(v) = "Y",
            LagQ = Number.Round(SetNum("fin_lag_q", 45)),
            LagY = Number.Round(SetNum("fin_lag_y", 90)),

            // ---------------------------------------------------------------- 직전 표(자기참조) → 없으면 Seed → 없으면 빈 표
            ReadStore = (name as text) as nullable table =>
                let t = try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null
                in if t = null or not Table.HasColumns(t, Cols) then null else Table.SelectColumns(t, Cols),
            PrevTable = ReadStore("tblFin"),
            StoreRaw = if PrevTable <> null then PrevTable
                       else let s = ReadStore("tblFinSeed") in if s <> null then s else #table(Cols, {}),
            // 행 = 레코드(열 순서 = Cols), 값 형식 정리. 종목코드가 빈 행(빈 표의 빈 행 등)은 버림
            StoreRows = List.Buffer(List.Select(List.Transform(Table.ToRecords(StoreRaw), (r) => [
                    종목코드 = Code(r[종목코드]), 결산년월 = YmText(r[결산년월]), 공개기준일 = DateOrNull(r[공개기준일]),
                    매출 = NumOrNull(r[매출]), 영업이익 = NumOrNull(r[영업이익]), 순이익 = NumOrNull(r[순이익]),
                    EPS = NumOrNull(r[EPS]), BPS = NumOrNull(r[BPS]), ROE = NumOrNull(r[ROE]),
                    상태 = Txt(r[상태]), 조회시각 = DateTimeOrNull(r[조회시각])]),
                (r) => r[종목코드] <> null)),
            ToTable = (rows as list) as table =>
                Table.TransformColumnTypes(Table.FromRecords(rows, Cols, MissingField.UseNull), Types),

            // ---------------------------------------------------------------- 종목 1개 조회 → 그 종목의 행 목록(레코드)
            FetchStock = (code as text) as list =>
                let
                    Q = [FID_DIV_CLS_CODE = "1", fid_cond_mrkt_div_code = "J", fid_input_iscd = code],
                    JI = fnKisGet("/uapi/domestic-stock/v1/finance/income-statement", "FHKST66430200", Q),
                    JR = fnKisGet("/uapi/domestic-stock/v1/finance/financial-ratio", "FHKST66430300", Q),
                    Out = (j as record) as list =>
                        let o = Record.FieldOrDefault(j, "output", {})
                        in List.Select(if o is list then o else if o is record then {o} else {}, (x) => x is record),
                    // 빈 값 표시 "", "99.99", 0 → null
                    Val = (r as record, k as text) as nullable number =>
                        let n = fnNum(Record.FieldOrDefault(r, k, null))
                        in if n = null or n = 0 or n = 99.99 then null else n,
                    Ym = (r as record) as nullable text =>
                        let t = Txt(Record.FieldOrDefault(r, "stac_yymm", null))
                        in if t <> null and Text.Length(t) = 6 and Text.Select(t, {"0".."9"}) = t then YmText(t) else null,
                    IsRows = List.Distinct(List.Select(List.Transform(Out(JI), (r) =>
                        [ym = Ym(r), sales = Val(r, "sale_account"), op = Val(r, "bsop_prti"), ni = Val(r, "thtr_ntin")]),
                        (x) => x[ym] <> null), (x) => x[ym]),
                    FrRows = List.Distinct(List.Select(List.Transform(Out(JR), (r) =>
                        [ym = Ym(r), eps = Val(r, "eps"), bps = Val(r, "bps"),
                         roe = let p = Val(r, "roe_val") in if p = null then null else p / 100]),
                        (x) => x[ym] <> null), (x) => x[ym]),
                    IsBy = Record.FromList(IsRows, List.Transform(IsRows, (x) => x[ym])),
                    FrBy = Record.FromList(FrRows, List.Transform(FrRows, (x) => x[ym])),
                    Yms = List.Sort(List.Distinct(List.Transform(IsRows, (x) => x[ym]) & List.Transform(FrRows, (x) => x[ym]))),
                    // 결산년월별 누적값(오래된 분기 → 최근 분기)
                    Cum = List.Buffer(List.Transform(Yms, (y) =>
                        let a = Record.FieldOrDefault(IsBy, y, []), b = Record.FieldOrDefault(FrBy, y, [])
                        in [ym = y, sales = Record.FieldOrDefault(a, "sales", null), op = Record.FieldOrDefault(a, "op", null),
                            ni = Record.FieldOrDefault(a, "ni", null), eps = Record.FieldOrDefault(b, "eps", null),
                            bps = Record.FieldOrDefault(b, "bps", null), roe = Record.FieldOrDefault(b, "roe", null)])),
                    CumBy = Record.FromList(Cum, Yms),
                    HasAny = (c as record) as logical => List.NonNullCount({c[sales], c[op], c[ni], c[eps], c[bps], c[roe]}) > 0,
                    WithData = List.Select(Cum, HasAny),
                    MonthOf = (y as text) as number => Number.From(Text.End(y, 2)),
                    Back3 = (y as text) as text =>
                        let i = Number.From(Text.Start(y, 4)) * 12 + MonthOf(y) - 1 - 3
                        in Text.From(Number.IntegerDivide(i, 12)) & Text.PadStart(Text.From(Number.Mod(i, 12) + 1), 2, "0"),
                    Prev3 = (y as text) as nullable record => Record.FieldOrDefault(CumBy, Back3(y), null),

                    // 회계연도 결산월 판정 (머리말 ①~③)
                    Votes = (k as text) as list => List.RemoveNulls(List.Transform(Cum, (c) =>
                        let
                            p = Prev3(c[ym]),
                            a = if p = null then null else Record.Field(p, k),
                            b = Record.Field(c, k)
                        in
                            if a = null or b = null then null
                            else if k = "sales" then (if b < a then MonthOf(c[ym]) else null)
                            else if (a > 0) = (b > 0) and Number.Abs(b) < Number.Abs(a) then MonthOf(c[ym])
                            else null)),
                    SalesVotes = Votes("sales"),
                    UseVotes = if List.IsEmpty(SalesVotes) then Votes("op") & Votes("ni") else SalesVotes,
                    Q1Month =
                        if List.IsEmpty(UseVotes) then null
                        else
                            let
                                Tally = List.Transform(List.Distinct(UseVotes), (m) => [
                                    m = m,
                                    key = List.Count(List.Select(UseVotes, (v) => v = m)) * 1000
                                          + List.PositionOf(UseVotes, m, Occurrence.Last)])
                            in
                                List.Last(List.Sort(Tally, {(t) => t[key], Order.Ascending}))[m],
                    LatestYm = if List.IsEmpty(WithData) then null else List.Last(WithData)[ym],
                    DataMonths = List.Distinct(List.Transform(WithData, (c) => MonthOf(c[ym]))),
                    FyEnd =
                        if LatestYm = null then null
                        else if Q1Month <> null then Number.Mod(Q1Month + 8, 12) + 1
                        else if List.Count(WithData) >= 2 and List.Count(DataMonths) = 1 then DataMonths{0}
                        else if Number.Mod(MonthOf(LatestYm), 3) = 0 then 12
                        else null,
                    // 회계 분기 번호 1~4 (결산월 체계에 안 맞거나 판정 불가면 null)
                    QNum = (y as text) as nullable number =>
                        if FyEnd = null then null
                        else
                            let d = Number.Mod(MonthOf(y) - FyEnd + 12, 12)
                            in if Number.Mod(d, 3) <> 0 then null else if d = 0 then 4 else d / 3,
                    Disc = (c as record, k as text) as nullable number =>
                        let q = QNum(c[ym]), cur = Record.Field(c, k)
                        in if cur = null or q = null then null
                           else if q = 1 then cur
                           else
                               let p = Prev3(c[ym]), pv = if p = null then null else Record.Field(p, k)
                               in if pv = null then null else cur - pv,

                    // 최근 12개 분기: 값이 있는 가장 최근 분기부터 응답에 있는 분기 12개(최근 → 과거)
                    Window = if LatestYm = null then {}
                             else List.FirstN(List.Reverse(List.Select(Cum, (c) => c[ym] <= LatestYm)), 12),
                    Rows = List.Transform(Window, (c) =>
                        let
                            q = QNum(c[ym]),
                            MonthEnd = Date.EndOfMonth(#date(Number.From(Text.Start(c[ym], 4)), MonthOf(c[ym]), 1)),
                            Lag = if q = null or q = 4 then LagY else LagQ
                        in
                            [종목코드 = code, 결산년월 = c[ym], 공개기준일 = Date.AddDays(MonthEnd, Lag),
                             매출 = Disc(c, "sales"), 영업이익 = Disc(c, "op"), 순이익 = Disc(c, "ni"), EPS = Disc(c, "eps"),
                             BPS = c[bps], ROE = c[roe], 상태 = "OK", 조회시각 = Now]),
                    NoData = {[종목코드 = code, 결산년월 = null, 공개기준일 = null, 매출 = null, 영업이익 = null, 순이익 = null,
                               EPS = null, BPS = null, ROE = null, 상태 = "데이터 없음", 조회시각 = Now]},
                    Result = if List.IsEmpty(Rows) then NoData else Rows,
                    // 모든 칸을 여기서 평가 → API·변환 오류가 종목 단위 try 안에서 잡힘(M 지연 평가)
                    Touched = List.Sum({0} & List.Transform(Result, (r) => List.NonNullCount(Record.FieldValues(r))))
                in
                    if Touched >= 0 then List.Buffer(Result) else {},

            // ---------------------------------------------------------------- 같은 버튼 실행 안의 재평가 확인
            // 새로 고침 직후 Excel이 쿼리를 백그라운드에서 한 번 더 평가할 때(아래 Main 주석) 이번 실행에서 이미 조회했으면 조회하지 않고
            // 자기 표를 돌려준다: 자기 표(tblFin — Seed는 보지 않음)의 가장 늦은 조회시각이 이번 실행의 started(버튼 매크로가 씀) 이후면
            // 이미 조회한 것. started가 비었거나 now_override가 있으면 확인하지 않는다.
            Started = try DateTimeOrNull(CtlValue("started")) otherwise null,
            OwnLast = if PrevTable = null then null
                      else List.Max(List.RemoveNulls(List.Transform(Table.Column(PrevTable, "조회시각"),
                               (v) => try DateTimeOrNull(v) otherwise null))),
            Surveyed = NowOverride = null and Started <> null and OwnLast <> null and OwnLast >= Started,

            // ---------------------------------------------------------------- 저장소 갱신(full)
            // 함수로 두어 full이고 이번 실행에서 아직 조회하지 않았을 때만 부른다(함수 값은 평가해도 호출하지 않음)
            Survey = () as table =>
                let
                    Targets = List.Buffer(List.Distinct(List.RemoveNulls(List.Transform(fnPageRows()[종목코드], Code)))),
                    Grouped = Table.Group(Table.FromRecords(StoreRows, Cols, MissingField.UseNull), {"종목코드"},
                                          {{"R", (t) => Table.ToRecords(t), type list}}),
                    PrevBy = Record.FromList(Grouped[R], Grouped[종목코드]),
                    StatusOf = (r as record) as text => if r[상태] = null then "" else r[상태],
                    Latest12 = (rows as list) as list => List.FirstN(List.Sort(rows, {(r) => r[결산년월], Order.Descending}), 12),
                    Steps = List.Transform(Targets, (code) =>
                        let
                            prev = Record.FieldOrDefault(PrevBy, code, {}),
                            dataRows = List.Select(prev, (r) => r[결산년월] <> null),
                            statuses = List.Transform(prev, StatusOf),
                            hasError = List.AnyTrue(List.Transform(statuses, (s) => Text.StartsWith(s, "오류"))),
                            noDataOk = List.IsEmpty(dataRows) and not List.IsEmpty(prev)
                                       and List.AllTrue(List.Transform(statuses, (s) => s = "데이터 없음")),
                            last = List.Max(List.RemoveNulls(List.Transform(prev, (r) => r[조회시각]))),
                            age = if last = null then null else Duration.Days(Today - Date.From(last)),
                            need = ForceWeekly or List.IsEmpty(prev) or hasError or (List.IsEmpty(dataRows) and not noDataOk)
                                   or age = null or age >= WeeklyDays
                        in
                            if not need then
                                // 조회하지 않은 종목: 직전 행 유지. 전체 실패 대체(fnFallback)가 붙인 "이전 데이터" 상태는 OK로 되돌림
                                [attempted = false, ok = true, reason = null,
                                 rows = Latest12(List.Transform(prev, (r) =>
                                     if Text.StartsWith(StatusOf(r), "이전 데이터") then r & [상태 = "OK"] else r))]
                            else
                                let
                                    got = try FetchStock(code),
                                    reason = if got[HasError] then
                                                 let m = Txt(got[Error][Message]) in
                                                 Text.Start(Text.Clean(if m = null then "사유 불명" else m), 100)
                                             else null
                                in
                                    if not got[HasError] then [attempted = true, ok = true, reason = null, rows = got[Value]]
                                    else if List.IsEmpty(dataRows) then
                                        [attempted = true, ok = false, reason = reason,
                                         rows = {[종목코드 = code, 결산년월 = null, 공개기준일 = null, 매출 = null, 영업이익 = null,
                                                  순이익 = null, EPS = null, BPS = null, ROE = null, 상태 = "오류: " & reason, 조회시각 = Now]}]
                                    else
                                        // 직전 행을 그대로 두고 상태만 오류(조회시각은 원래 값 → 다음 [전체]에서 다시 조회)
                                        [attempted = true, ok = false, reason = reason,
                                         rows = Latest12(List.Transform(dataRows, (r) => r & [상태 = "오류: " & reason]))]),
                    // 여기서 모든 종목을 순서대로 평가(호출)
                    Done = List.Buffer(Steps),
                    Attempted = List.Select(Done, (s) => s[attempted]),
                    Failed = List.Select(Attempted, (s) => not s[ok]),
                    Checked =
                        if not List.IsEmpty(Attempted) and List.Count(Failed) = List.Count(Attempted) then
                            error Error.Record("KIS 재무 조회 실패",
                                "조회한 " & Text.From(List.Count(Attempted)) & "종목 모두 실패: " & Failed{0}[reason])
                        else Done
                in
                    ToTable(List.Combine(List.Transform(Checked, (s) => s[rows]))),

            // ---------------------------------------------------------------- full이 아닌 모드: 호출 없이 직전 표(→ Seed → 빈 표)
            Passive = ToTable(StoreRows)
        in
            // Excel은 새로 고침이 끝난 직후 쿼리를 백그라운드에서 한 번 더 평가한다(2026-10-01 하네스 + 모의 서버 실측, 동기·백그라운드
            // 새로 고침 모두). 이 평가는 방금 적재한 자기 표를 다시 읽으므로 주간 규칙상 조회할 종목이 없으면 호출이 없지만,
            // force_weekly = Y(모든 종목)나 오류 재시도 종목은 Power Query 캐시에 없는 요청을 한 번 더 보냈다 → Surveyed 확인으로 막는다.
            // (조회한 종목이 모두 실패해 직전 표로 돌아간 경우엔 조회시각이 그대로라 막지 못하고 한 번 더 시도한다.)
            // 표를 처음 적재한 뒤·통합문서를 연 뒤 처음 하는 새로 고침에서는 추가 평가가 자기 표가 갱신되기 전에 함께 돌아 이 확인을 통과하므로
            // 그 한 번은 호출이 최대 2배가 될 수 있다(실측) — 버튼 매크로가 그 표를 build 모드로 먼저 한 번 새로 고침해 두면 생기지 않는다.
            if Mode = "full" and not Surveyed then Survey() else Passive,
    // 실패(전 종목 조회 실패·토큰·네트워크·대상 종목 오류 등) 시 직전 표를 유지. 직전 표를 다시 읽으면 열 형식이 없어
    // 날짜가 일련번호로 보이므로 같은 형식을 다시 지정
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblFin", Cols, Tried[Error][Message]), Types) else Tried[Value]
in
    Final
