// T_Quote : 보유·관심 종목 시세판 (현재가·밸류에이션·52주·수급·추정가집계·신용/공매도·다음 이벤트)
//  - 주식현재가 시세 [v1_국내주식-008] FHKST01010100 : 가격/PER/PBR/52주/외국인소진율/업종/시장경고
//  - 주식현재가 투자자 [v1_국내주식-012] FHKST01010900 : 최근 30영업일 외국인·기관·개인 순매수 → 5일/20일 합계(억원)
//  - 종목별 외인기관 추정가집계 [v1_국내주식-046] HHPTJ04160200 /uapi/domestic-stock/v1/quotations/investor-trend-estimate
//    (spec R19·R20, 2026-10-01 T22) : MKSC_SHRN_ISCD = 종목코드 → output2[] = 이미 입력된 시점만, 최근 시점이 앞.
//      bsop_hour_gb 1~5 = 09:30·10:00·11:20·13:20·14:30, 각 행 = 그 시점까지의 누적 추정 순매수 수량(주):
//      frgn_fake_ntby_qty 외국인, orgn_fake_ntby_qty 기관, sum_fake_ntby_qty 합산.
//      실측(2026-10-01 12:44, 3종목): 행 3개(gb 3·2·1), gb 1(09:30)의 기관 수량은 0 — 기관 첫 입력은 10:00(gb 2).
//      당일 전용 자료이고 응답에 날짜가 없으므로 개장일(평일이고 tblHolidays에 없음) 09:00~15:30에만 호출한다.
//      그 밖·첫 입력 전·호출 실패는 빈칸(0으로 채우지 않음). 시세판 열:
//        추정시점 = 가장 최근 입력 시점("HH:MM"), 추정외국인주·추정기관주 = 그 시점의 누적 수량(기관은 10:00 입력부터 — 그 전은 빈칸),
//        추정합산억 ≈ 합산 수량 × 현재가 ÷ 1억(원 → 억원).
//  - 대테마 = 종목DB(tblUniverse)의 대테마(fnUniverse — 섹터와 같은 경로, 수정표 반영)
//  - 신용잔고율·공매도비중5일 = tblCSL([전체]의 주간 항목) 값 그대로(소수: 1.23% = 0.0123). 표나 종목 행이 없으면 빈칸
//  - 다음이벤트 = tblEvents에서 그 종목의 일자 ≥ 오늘인 가장 가까운 이벤트 → "D-n 종류"(n = 일자 − 오늘, 같은 날 여러 종류면
//    "·"로 묶음). 일자가 빈 행(조회 실패 표시 행)은 무시. 없으면 빈칸
let
    Cols = {"구분", "종목코드", "종목명", "시장", "섹터", "대테마", "현재가", "전일대비", "등락률", "다음이벤트", "거래대금억", "시가총액억",
            "PER", "PBR", "외국인소진율", "고52주대비", "저52주대비", "신용잔고율", "공매도비중5일", "추정시점", "추정외국인주", "추정기관주",
            "추정합산억", "외국인5일억", "기관5일억", "개인5일억", "외국인20일억", "기관20일억", "시가", "고가", "저가", "전일종가", "거래량",
            "전일대비거래량", "EPS", "BPS", "고52주", "고52주일", "저52주", "저52주일", "외국인당일주", "프로그램당일주", "상한가", "하한가",
            "유의", "DB유의", "상태", "조회시각"},
    Types = {{"구분", type text}, {"종목코드", type text}, {"종목명", type text}, {"시장", type text}, {"섹터", type text},
             {"대테마", type text}, {"현재가", type number}, {"전일대비", type number}, {"등락률", type number}, {"다음이벤트", type text},
             {"거래대금억", type number}, {"시가총액억", type number}, {"PER", type number}, {"PBR", type number},
             {"외국인소진율", type number}, {"고52주대비", type number}, {"저52주대비", type number}, {"신용잔고율", type number},
             {"공매도비중5일", type number}, {"추정시점", type text}, {"추정외국인주", type number}, {"추정기관주", type number},
             {"추정합산억", type number}, {"외국인5일억", type number}, {"기관5일억", type number}, {"개인5일억", type number},
             {"외국인20일억", type number}, {"기관20일억", type number}, {"시가", type number}, {"고가", type number}, {"저가", type number},
             {"전일종가", type number}, {"거래량", type number}, {"전일대비거래량", type number}, {"EPS", type number}, {"BPS", type number},
             {"고52주", type number}, {"고52주일", type date}, {"저52주", type number}, {"저52주일", type date}, {"외국인당일주", type number},
             {"프로그램당일주", type number}, {"상한가", type number}, {"하한가", type number}, {"유의", type text}, {"DB유의", type text},
             {"상태", type text}, {"조회시각", type datetime}},
    Main =
        let
            S = fnSettings(),
            WithFlow = (try Text.Upper(Text.From(S[stock_flow])) otherwise "Y") = "Y",
            Now = DateTime.LocalNow(),
            Today = Date.From(Now),
            Codes = fnCodes(),
            U = fnUniverse(),
            Q = "/uapi/domestic-stock/v1/quotations/",
            // 다른 표의 셀 값은 형식이 일정하지 않으므로(날짜가 일시·일련번호로 올 수 있음) 오류 없이 읽는다
            Key = (v as any) as nullable text =>
                let t = try Text.Upper(Text.Trim(Text.From(v))) otherwise null in if t = null or t = "" then null else t,
            AsNum = (v as any) as nullable number =>
                let x = try (if v is number then v else fnNum(v)) otherwise null in if x = null or Number.IsNaN(x) then null else x,
            AsDate = (v as any) as nullable date => try Date.From(v) otherwise null,
            Book = (name as text) as nullable table => try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null,

            // 추정가집계 호출 조건: 개장일 09:00~15:30 (당일 전용 자료 — 그 밖에는 호출하지 않고 빈칸)
            HolT = Book("tblHolidays"),
            Holidays = if HolT = null or not Table.HasColumns(HolT, "휴장일") then {}
                       else List.Buffer(List.RemoveNulls(List.Transform(HolT[휴장일], AsDate))),
            Clock = Time.From(Now),
            InSession = Date.DayOfWeek(Today, Day.Monday) < 5 and not List.Contains(Holidays, Today)
                        and Clock >= #time(9, 0, 0) and Clock <= #time(15, 30, 0),
            SlotTimes = {"09:30", "10:00", "11:20", "13:20", "14:30"},
            // 가장 최근 입력 시점의 [g 시점번호, f 외국인, i 기관, s 합산] — 입력 전·장 밖·실패는 null
            Estimate = (code as text) as nullable record =>
                if not InSession then null
                else
                    let
                        r = try fnKisGet(Q & "investor-trend-estimate", "HHPTJ04160200", [MKSC_SHRN_ISCD = code]) otherwise null,
                        o = if r = null then {} else Record.FieldOrDefault(r, "output2", {}),
                        L = if o is list then List.Select(o, each _ is record) else if o is record then {o} else {},
                        Slots = List.Select(List.Transform(L, (x) => [
                                    g = fnNum(Record.FieldOrDefault(x, "bsop_hour_gb", null)),
                                    f = fnNum(Record.FieldOrDefault(x, "frgn_fake_ntby_qty", null)),
                                    i = fnNum(Record.FieldOrDefault(x, "orgn_fake_ntby_qty", null)),
                                    s = fnNum(Record.FieldOrDefault(x, "sum_fake_ntby_qty", null))]),
                                each [g] <> null and List.Contains({1, 2, 3, 4, 5}, [g]))
                    in
                        if List.IsEmpty(Slots) then null else List.Last(List.Sort(Slots, each [g])),

            Get = (code as text) as record =>
                let
                    o = fnKisGet(Q & "inquire-price", "FHKST01010100", [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code])[output],
                    n = (k as text) => fnNum(Record.FieldOrDefault(o, k, null)),
                    p = (k as text) => let v = n(k) in if v = null then null else v / 100,
                    Inv = if not WithFlow then {}
                          else let r = try fnKisGet(Q & "inquire-investor", "FHKST01010900", [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code])[output] otherwise {}
                               in if r is list then List.Select(r, each Text.From(Record.FieldOrDefault(_, "stck_bsop_date", "")) <> "") else {},
                    SumN = (k as text, days as number) as nullable number =>
                        let v = List.RemoveNulls(List.Transform(List.FirstN(Inv, days), each fnNum(Record.FieldOrDefault(_, k, null))))
                        in if List.Count(v) = 0 then null else List.Sum(v) / 100,
                    Est = Estimate(code),
                    EstNo = if Est = null then null else Number.RoundDown(Est[g]),
                    Price = n("stck_prpr"),
                    warnCode = Text.From(Record.FieldOrDefault(o, "mrkt_warn_cls_code", "")),
                    flags = List.RemoveNulls({
                        if Record.FieldOrDefault(o, "temp_stop_yn", "N") = "Y" then "거래정지" else null,
                        if warnCode = "01" then "투자주의" else if warnCode = "02" then "투자경고" else if warnCode = "03" then "투자위험" else null,
                        if Record.FieldOrDefault(o, "invt_caful_yn", "N") = "Y" then "투자유의" else null,
                        if Record.FieldOrDefault(o, "short_over_yn", "N") = "Y" then "단기과열" else null,
                        if Record.FieldOrDefault(o, "mang_issu_cls_code", "N") = "Y" then "관리종목" else null,
                        if Text.From(Record.FieldOrDefault(o, "vi_cls_code", "N")) <> "N" and Text.From(Record.FieldOrDefault(o, "vi_cls_code", "N")) <> "" then "VI발동" else null})
                in
                    // 레코드 필드는 지연 평가되므로 응답 o를 먼저 평가해 try가 API 오류를 잡도록 함
                    if not (o is record) then error "빈 응답" else
                    [종목코드 = code,
                     현재가 = Price, 전일대비 = n("prdy_vrss"), 등락률 = p("prdy_ctrt"),
                     시가 = n("stck_oprc"), 고가 = n("stck_hgpr"), 저가 = n("stck_lwpr"), 전일종가 = n("stck_sdpr"),
                     거래량 = n("acml_vol"), 거래대금억 = let v = n("acml_tr_pbmn") in if v = null then null else v / 100000000,
                     전일대비거래량 = p("prdy_vrss_vol_rate"),
                     시가총액억 = n("hts_avls"), PER = n("per"), PBR = n("pbr"), EPS = n("eps"), BPS = n("bps"),
                     외국인소진율 = p("hts_frgn_ehrt"),
                     고52주 = n("w52_hgpr"), 고52주일 = fnYmd(Record.FieldOrDefault(o, "w52_hgpr_date", null)),
                     저52주 = n("w52_lwpr"), 저52주일 = fnYmd(Record.FieldOrDefault(o, "w52_lwpr_date", null)),
                     외국인당일주 = n("frgn_ntby_qty"), 프로그램당일주 = n("pgtr_ntby_qty"),
                     외국인5일억 = SumN("frgn_ntby_tr_pbmn", 5), 기관5일억 = SumN("orgn_ntby_tr_pbmn", 5), 개인5일억 = SumN("prsn_ntby_tr_pbmn", 5),
                     외국인20일억 = SumN("frgn_ntby_tr_pbmn", 20), 기관20일억 = SumN("orgn_ntby_tr_pbmn", 20),
                     추정시점 = if Est = null then null else SlotTimes{EstNo - 1},
                     추정외국인주 = if Est = null then null else Est[f],
                     추정기관주 = if Est = null or EstNo < 2 then null else Est[i],
                     추정합산억 = if Est = null or Est[s] = null or Price = null then null else Est[s] * Price / 100000000,
                     상한가 = n("stck_mxpr"), 하한가 = n("stck_llam"),
                     유의 = Text.Combine(flags, ","), 상태 = "OK", 조회시각 = Now],
            Rows = List.Transform(Table.ToRecords(Codes), (c) =>
                let r = try Get(c[종목코드]) in
                    (if r[HasError] then [종목코드 = c[종목코드], 상태 = "오류: " & r[Error][Message], 조회시각 = Now] else r[Value]) & [구분 = c[구분]]),
            RawCols = {"구분", "종목코드", "현재가", "전일대비", "등락률", "시가", "고가", "저가", "전일종가", "거래량", "거래대금억",
                       "전일대비거래량", "시가총액억", "PER", "PBR", "EPS", "BPS", "외국인소진율", "고52주", "고52주일", "저52주", "저52주일",
                       "외국인당일주", "프로그램당일주", "외국인5일억", "기관5일억", "개인5일억", "외국인20일억", "기관20일억",
                       "추정시점", "추정외국인주", "추정기관주", "추정합산억", "상한가", "하한가", "유의", "상태", "조회시각"},
            Failed = List.Select(Rows, each Text.StartsWith(Text.From(Record.FieldOrDefault(_, "상태", "")), "오류")),
            // 전부 실패하면 오류로 올려 직전 데이터를 유지(fnFallback), 일부만 실패하면 그 행만 오류 표시
            Checked = if List.Count(Rows) > 0 and List.Count(Failed) = List.Count(Rows) then error Text.Middle(Failed{0}[상태], 4) else Rows,
            T0 = Table.FromRecords(Checked, RawCols, MissingField.UseNull),
            T = Table.AddIndexColumn(T0, "__ord", 0, 1, Int64.Type),
            J = Table.NestedJoin(T, {"종목코드"}, U, {"종목코드"}, "u", JoinKind.LeftOuter),
            E = Table.ExpandTableColumn(J, "u", {"종목명", "시장", "섹터", "대테마", "유의사항"}, {"종목명", "시장", "섹터", "대테마", "DB유의"}),
            Named = Table.AddColumn(E, "__n", each if [종목명] = null then [종목코드] else [종목명], type text),
            Fixed = Table.RenameColumns(Table.RemoveColumns(Named, {"종목명"}), {{"__n", "종목명"}}),
            W52 = Table.AddColumn(Table.AddColumn(Fixed,
                    "고52주대비", each if [고52주] = null or [고52주] = 0 or [현재가] = null then null else [현재가] / [고52주] - 1, type number),
                    "저52주대비", each if [저52주] = null or [저52주] = 0 or [현재가] = null then null else [현재가] / [저52주] - 1, type number),

            // 신용잔고율·공매도비중5일 (tblCSL — 종목코드당 첫 행)
            CslT = Book("tblCSL"),
            CslPairs = if CslT = null or not Table.HasColumns(CslT, "종목코드") then {}
                       else List.Transform(Table.ToRecords(Table.SelectColumns(CslT, {"종목코드", "신용잔고율", "공매도비중5일"}, MissingField.UseNull)),
                                (r) => {Key(r[종목코드]), [cr = AsNum(r[신용잔고율]), sh = AsNum(r[공매도비중5일])]}),
            CslDist = List.Distinct(List.Select(CslPairs, each _{0} <> null), each _{0}),
            CslMap = Record.FromList(List.Transform(CslDist, each _{1}), List.Transform(CslDist, each _{0})),
            Csl = (code as any, field as text) as nullable number =>
                let k = Key(code), x = if k = null then null else Record.FieldOrDefault(CslMap, k, null)
                in if x = null then null else Record.Field(x, field),

            // 다음 이벤트 (tblEvents — 일자 ≥ 오늘, 일자·종목코드가 빈 행은 무시)
            EvT = Book("tblEvents"),
            EvRows = if EvT = null or not Table.HasColumns(EvT, {"일자", "종목코드", "종류"}) then {}
                     else List.Buffer(List.Select(List.Transform(Table.ToRecords(Table.SelectColumns(EvT, {"일자", "종목코드", "종류"})),
                                (r) => [d = AsDate(r[일자]), c = Key(r[종목코드]),
                                        k = let t = try Text.Trim(Text.From(r[종류])) otherwise null in if t = null or t = "" then null else t]),
                            each [d] <> null and [c] <> null and [d] >= Today)),
            NextEvent = (code as any) as nullable text =>
                let
                    kc = Key(code),
                    ev = if kc = null then {} else List.Select(EvRows, each [c] = kc),
                    d0 = if List.IsEmpty(ev) then null else List.Min(List.Transform(ev, each [d])),
                    kinds = if d0 = null then {} else List.Distinct(List.RemoveNulls(List.Transform(List.Select(ev, each [d] = d0), each [k])))
                in
                    if d0 = null then null
                    else "D-" & Text.From(Duration.Days(d0 - Today)) & (if List.IsEmpty(kinds) then "" else " " & Text.Combine(kinds, "·")),
            Looked = Table.AddColumn(Table.AddColumn(Table.AddColumn(W52,
                        "신용잔고율", each Csl([종목코드], "cr"), type number),
                        "공매도비중5일", each Csl([종목코드], "sh"), type number),
                        "다음이벤트", each NextEvent([종목코드]), type text),
            // PER·PBR 등 0은 '데이터 없음'으로 처리
            NoZero = Table.TransformColumns(Looked, List.Transform({"PER", "PBR", "EPS", "BPS"}, (c) => {c, each if _ = 0 then null else _})),
            // 핵심 열을 앞쪽에 배치 (시트에서 신호·추세 등 계산열이 등락률 뒤에, 관심가·MA 등이 기관20일억 뒤에 끼워짐)
            Final = Table.SelectColumns(Table.Sort(NoZero, {{"__ord", Order.Ascending}}), Cols),
            Typed = Table.TransformColumnTypes(Final, Types)
        in
            Typed,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지. 직전 표를 다시 읽으면 열 형식이 없어(any) Excel이
    // 날짜 칸 서식을 '일반'으로 바꾸므로 같은 형식을 다시 지정(값은 그대로)
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblQuote", Cols, Tried[Error][Message]), Types) else Tried[Value]
in
    Final
