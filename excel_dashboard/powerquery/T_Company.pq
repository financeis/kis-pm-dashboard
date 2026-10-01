// T_Company : 대회종목 페이지 표 — 페이지 행마다 분류·가격·수급·밸류·KIS 추정·목표주가·실적·신용/공매도/대차·이벤트를 한 행에 결합
//   (spec R16 데이터·R5~R11, 2026-10-01 T16) → tblCompany   [KIS 호출 없음 — 어느 모드에서나 통합문서에 있는 표로 계산]
//  - 행 = fnPageRows()(대회 종목 ± 수정표 ∪ 보유·관심). 정렬: 유니버스 = 대회 행을 시가총액억 내림차순(빈칸은 뒤) → 그 아래
//    유니버스 밖 행을 시가총액억 내림차순(빈칸은 뒤). 같은 시가총액은 fnPageRows 순서(명단 순서).
//  - 실행: [시세](quick)·[전체](full)의 T_PxMetrics 다음(전체는 T_CSL 다음), 그리고 빌드·수동 새로 고침(build)에서도 같은 계산.
//    단 버튼 실행(quick·full) 중 tblRunCtl started 이후에 이미 계산한 표(자기 표 조회시각 ≥ started)면 자기 표를 그대로 돌려준다
//    — 새로 고침 뒤 Excel이 백그라운드에서 쿼리를 다시 평가할 때 같은 계산을 반복하지 않게(T_PxMetrics와 같은 가드).
//    계산은 함수 단계(Compute) 안에서만 한다(맨 위 단계를 따로 평가해도 계산하지 않음).
//  - 원천 표(모두 없거나 비어 있어도 됨 → 그 열은 빈칸, 오류 아님. 종목코드가 빈 행(실패 경고 행)은 건너뜀. 같은 종목은 첫 행):
//      fnClassify(페이지 종목, KRX업종 = tblUniverse) — NICS 대분류·업종·세부, 대테마·세부테마, 분류출처(수정표 반영)
//      tblClass(KSIC 세분류·주요상품·기업집단), tblPxMetrics(가격 지표), tblFlowU, tblFin, tblTarget, tblEst, tblEstSnap, tblCSL,
//      tblEvents, tblUniverse(종목명·시장·유의사항·KRX업종), tblContest(종목명·시장 대체), 세션 달력 fnPxCalendar(tblSessions;
//      비면 T_PxMetrics처럼 tblPxStore → tblPxStoreSeed의 날짜), tblSettings(fin_lag_q 45·fin_lag_y 90 — 4분기 판정용).
//  - '오늘' = fnPxRunCtl()[Today] (평소 실제 날짜, 시험용 now_override가 있으면 그 날짜). 밸류 변화의 '오늘 알려진' TTM EPS·BPS,
//    Fwd EPS·다음 이벤트에 쓴다(후행PER·PBR은 오늘과 무관 — 아래). 조회시각 = 실제 계산 시각.
//  - 열 (금액 억원·원, 비율은 소수 — 페이지가 % 서식, 없으면 빈칸이고 0을 쓰지 않음):
//      기본       종목코드, 종목명·시장(tblUniverse → 없으면 tblContest, 칸마다), 구분·유니버스(fnPageRows), NICS 업종·대테마(fnClassify)
//      규모/가격  시가총액억, 거래대금60일억, 거래량60일천주, 현재가, 고52주대비, 가격기준일, 1D…1Y, D-19…D0 — tblPxMetrics 값 그대로
//      수급       외국인·기관 1D·1W·1M % = 순매수 억원 ÷ 시가총액억(소수; 시가총액이 없거나 ≤ 0이면 빈칸), 억원 열은 tblFlowU 그대로,
//                 수급기준일 = tblFlowU 기준일(장중에는 직전 완료 세션)
//      후행PER    현재가 ÷ 최근 TTM EPS. 최근 TTM EPS = 그 종목 tblFin의 최근 분기(결산년월이 있는 가장 최근 행)와 그 앞 3개 분기
//                 (결산년월이 정확히 3개월씩 이어짐)의 분기 단독 EPS 합 — 공개기준일과 무관(받은 자료는 이미 공시된 실적이므로 공시 지연을
//                 적용하지 않음). 넷 중 하나라도 행이 없거나 EPS가 비면 빈칸, 합 ≤ 0이면 빈칸.
//      PBR        현재가 ÷ 최근 분기의 BPS(공개기준일과 무관, 그 칸이 비면 빈칸). ≤ 0이면 빈칸.
//                 공시 지연(공개기준일)은 아래 '밸류 변화'의 분자·분모에만 쓴다.
//      FwdEPS     fnFwdEps(FY1, FY1_EPS, FY2_EPS, 오늘) — 가중 평균(FY2가 없으면 FY1 EPS). FwdPER = 현재가 ÷ FwdEPS(≤ 0이면 빈칸).
//                 tblEst 상태와 관계없이 표에 있는 값으로 계산한다('오류:'·'이전 데이터' 행은 직전 값 — 마지막으로 받은 추정).
//      FwdPER변화1W·1M  현재 FwdPER ÷ 기준 세션의 FwdPER − 1. 기준 세션 B = 가격기준일에서 5·21세션 전(세션 달력).
//                 ① tblEstSnap에 그 종목의 일자가 [B − 3세션, B]인 행이 있으면 가장 늦은 행의 FwdPER(그 칸이 비면 결과도 빈칸).
//                 ② 없고 tblEst 추정일 < B(같은 날은 아님)이면 재구성: 기준 세션 종가 ÷ fnFwdEps(FY1, FY1_EPS, FY2_EPS, B).
//                    기준 세션 종가 = 현재가 ÷ (1 + 1W 또는 1M) — tblPxMetrics 등락률의 기준 세션과 같은 봉(정확히 복원됨).
//                 ③ 그 밖은 빈칸. 어느 쪽 FwdPER이든 빈칸·≤ 0이면 빈칸.
//      추정일·FY1·FY1_EPS·FY2_EPS  tblEst 그대로. FY2 = tblEst FY2, 단 FY1 EPS는 있는데 FY2 EPS가 없으면 글자 "FY1만"(spec R9 '표시';
//                 이때 FwdEPS = FY1 EPS, FY2_EPS 빈칸).
//      목표주가   목표가평균·증권사수·최근의견일·최근증권사·최근의견 = tblTarget 그대로(증권사수 ≤ 0은 빈칸),
//                 괴리율 = 목표가평균 ÷ 현재가 − 1, 목표가변화1M·3M = 목표가평균 ÷ 목표가평균_1M전·_3M전 − 1(분모 ≤ 0·빈칸이면 빈칸)
//      실적       최근 분기 = 그 종목 tblFin에서 결산년월이 있는 가장 최근 행(공개기준일과 무관 — 받은 자료의 최신 분기).
//                 최근분기 = "yyyy.Qn": yyyy = 그 결산월의 연도(달력 연도), n = 회계 분기 번호. 4분기 = 회계연도 결산 분기 =
//                 공개기준일 − 결산월 말일 = fin_lag_y(기본 90)인 행(T_Fin 규칙). 최근 행이 그런 행이면 Q4, 아니면 그보다 앞선 가장 가까운
//                 4분기 행의 월 f로 n = ((월 − f) mod 12) ÷ 3 (12월 결산이면 월 ÷ 3; 테라뷰 4월 결산 202607 → 2026.Q1).
//                 4분기 행이 없으면 월이 3·6·9·12일 때 12월 결산으로 보고 n = 월 ÷ 3. 그래도 정할 수 없으면 "yyyy.mm"(결산년월 그대로).
//                 fin_lag_q = fin_lag_y로 설정하면 4분기를 가릴 수 없어 3·6·9·12월만 n = 월 ÷ 3, 그 밖은 "yyyy.mm".
//                 매출·영업이익·순이익 YoY = 최근 분기 단독값 ÷ 전년 동분기(결산년월 − 100) − 1. 둘 중 하나라도 없거나 전년 ≤ 0이면 빈칸.
//                 실적상태(영업이익): 흑자전환(전년 ≤ 0, 올해 > 0) / 적자전환(전년 > 0, 올해 ≤ 0) / 적자지속 / 흑자지속, 없으면 빈칸.
//                 ROE = 최근 분기 행의 ROE(T_Fin이 소수로 바꾼 API 값 그대로).
//      밸류 변화  PER변화k = (현재가 ÷ TTM EPS(오늘)) ÷ (기준 세션 종가 ÷ TTM EPS(기준일)) − 1, PBR변화k도 BPS(d)로 같은 방식.
//                 TTM EPS(d) = 날 d에 '알려진'(공개기준일 ≤ d) 분기 중 가장 최근 분기와 그 앞 3개 분기(결산년월이 정확히 3개월씩
//                 이어지고 모두 공개기준일 ≤ d, EPS가 있음)의 분기 단독 EPS 합 — 하나라도 없으면 빈칸. BPS(d) = 공개기준일 ≤ d인 가장
//                 최근 분기의 BPS. 분자(오늘)·분모(기준일) 모두 이 공시 지연 규칙(T_Fin이 1~3분기 + 45일·4분기 + 90일로 정함)을 쓰므로
//                 최근 분기가 아직 공개기준일 전이면 분자의 EPS·BPS는 후행PER·PBR의 값과 다르다(같은 기준끼리 비교).
//                 k = 3M·6M·1Y(가격기준일에서 63·126·250세션 전) / YTD(가격기준일 연도의 직전 연도 마지막 세션) — tblPxMetrics와 같은
//                 기준 세션, 기준 세션 종가 = 현재가 ÷ (1 + 그 기간 등락률). 어느 쪽 EPS·BPS라도 빈칸·≤ 0이면 빈칸.
//      신용 등    신용잔고율·신용잔고율1M변화·공매도비중5일·대차잔고1M변화율 = tblCSL 그대로, 신용기준일 = tblCSL 기준일
//      이벤트     다음이벤트 = tblEvents에서 그 종목의 일자 ≥ 오늘인 가장 가까운 일자 → "D-n 종류"(n = 일자 − 오늘, 오늘이면 D-0,
//                 같은 날 여러 종류는 표 순서대로 "·"로 묶음 — T_Quote 시세판과 같은 글자). 다음이벤트일 = 그 일자(페이지가 그날 기준
//                 D-day를 다시 계산할 수 있게; tblEvents의 D-day 열은 자정이 지나면 낡으므로 쓰지 않음). 일자·종목코드가 빈 행은 무시.
//      분류 세부  NICS 대분류·NICS 세부·세부테마·분류출처(fnClassify), KSIC 세분류 = tblClass 값에서 앞의 코드("C26420.")를 뗀 이름,
//                 주요상품·기업집단 = tblClass 그대로
//      상태       유의 = tblUniverse 유의사항(거래정지·관리종목 등). 상태 = tblPxMetrics 그 종목 상태(OK / 데이터 없음 / 오류: … /
//                 이전 데이터(…) — 가격 지표가 정상이면 OK); 지표 행이 없으면 데이터 없음. 다른 원천이 비어 있어도 상태를 오류로 하지 않는다.
//                 종목 행 계산이 실패하면 그 행만 "오류: 계산 실패(사유)"(코드·구분·유니버스만 채움).
//  - 성능: 원천 표마다 필요한 열만 읽어 페이지 종목 행만 남기고(해시 집합) 종목코드별 레코드로 만든 뒤 행마다 바로 꺼낸다. 금액·날짜
//    형식 변환은 표 단위(Table.TransformColumnTypes)로, 스냅샷은 기준 세션 창(가장 이른 가격기준일 − 24세션 이후)의 행만 묶는다.
//    행을 레코드 목록으로 바꿔 행마다 함수를 돌리는 방식은 tblClass(약 3,800행)·스냅샷(1년 약 2.8만 행)에서 수 초가 걸려 쓰지 않는다.
//    실측(2026-10-01, 하네스, 대회 527 + 관심 1종목·실적 6천 행·스냅샷 2.5만 행 합성 입력, [시세]처럼 quick + started·동기 새로 고침):
//    1.7~2.4초(중앙값 약 2초; 같은 PC에서 다른 Excel 작업이 돌면 더 걸림). 그중 fnPageRows + fnClassify 약 0.5초.
//  - 실패(통합문서 표를 읽다 생긴 예외 등): 직전 표 유지(fnFallback — 상태 '이전 데이터(갱신 실패: 사유)', 열 형식 다시 지정).
//    직전 표가 없으면 종목코드가 빈 1행에 "오류: 사유"(T_ThemeAgg는 유니버스가 빈 이 행을 건너뛴다).
let
    Periods = {"1D", "1W", "1M", "3M", "6M", "YTD", "1Y"},
    StripCols = List.Transform(List.Reverse(List.Numbers(0, 20)), each if _ = 0 then "D0" else "D-" & Text.From(_)),
    Cols = {"종목코드", "종목명", "시장", "구분", "유니버스", "NICS 업종", "대테마",
            "시가총액억", "거래대금60일억", "거래량60일천주",
            "현재가", "고52주대비", "가격기준일"}
           & Periods & StripCols
           & {"외국인1D%", "외국인1W%", "외국인1M%", "외국인1D억", "외국인1W억", "외국인1M억",
              "기관1D%", "기관1W%", "기관1M%", "기관1D억", "기관1W억", "기관1M억", "수급기준일",
              "후행PER", "PBR",
              "FwdPER", "FwdPER변화1W", "FwdPER변화1M", "FwdEPS", "추정일", "FY1", "FY1_EPS", "FY2", "FY2_EPS",
              "목표가평균", "괴리율", "증권사수", "목표가변화1M", "목표가변화3M", "최근의견일", "최근증권사", "최근의견",
              "최근분기", "매출YoY", "영업이익YoY", "순이익YoY", "실적상태", "ROE",
              "PER변화3M", "PER변화6M", "PER변화YTD", "PER변화1Y", "PBR변화3M", "PBR변화6M", "PBR변화YTD", "PBR변화1Y",
              "신용잔고율", "신용잔고율1M변화", "공매도비중5일", "대차잔고1M변화율", "신용기준일",
              "다음이벤트", "다음이벤트일",
              "NICS 대분류", "NICS 세부", "세부테마", "KSIC 세분류", "주요상품", "기업집단", "분류출처",
              "유의", "상태", "조회시각"},
    NCols = List.Count(Cols),
    TextCols = {"종목코드", "종목명", "시장", "구분", "유니버스", "NICS 업종", "대테마", "FY1", "FY2", "최근증권사", "최근의견",
                "최근분기", "실적상태", "다음이벤트", "NICS 대분류", "NICS 세부", "세부테마", "KSIC 세분류", "주요상품", "기업집단",
                "분류출처", "유의", "상태"},
    DateCols = {"가격기준일", "수급기준일", "추정일", "최근의견일", "신용기준일", "다음이벤트일"},
    Types = List.Transform(Cols, each {_,
        if List.Contains(TextCols, _) then type text
        else if List.Contains(DateCols, _) then type date
        else if _ = "조회시각" then type datetime
        else type number}),
    Ctl = fnPxRunCtl(),
    // 버튼 실행(quick·full) 중 이미 이번 실행(started 이후)에 계산한 표면 자기 표를 그대로(백그라운드 재평가의 반복 계산 방지).
    // build(빌드·수동 새로 고침)에서는 늘 다시 계산한다. 'and'는 앞이 거짓이면 뒤를 평가하지 않으므로 build에서는 자기 표도 읽지 않음
    OwnT = try Excel.CurrentWorkbook(){[Name = "tblCompany"]}[Content] otherwise null,
    OwnOk = OwnT is table and List.ContainsAll(Table.ColumnNames(OwnT), Cols),
    LastStamp = if not OwnOk then null
                else List.Max(List.RemoveNulls(List.Transform(Table.Column(OwnT, "조회시각"), each try DateTime.From(_) otherwise null)), null),
    Already = (Ctl[Mode] = "quick" or Ctl[Mode] = "full") and Ctl[Started] <> null and LastStamp <> null and LastStamp >= Ctl[Started],
    KeptOwn = if OwnOk then Table.TransformColumnTypes(Table.SelectColumns(OwnT, Cols), Types) else null,

    Compute = () as table =>
        let
            Today = Ctl[Today],
            Now = DateTime.LocalNow(),
            Digits = {"0".."9"},
            CodeChars = {"0".."9", "A".."Z"},
            // ---------------------------------------------------------------- 값 도우미 (통합문서 표 값은 형식이 없을 수 있음)
            Txt = (v as any) as nullable text =>
                if v = null then null else let t = Text.Trim(try Text.From(v) otherwise "") in if t = "" then null else t,
            // 종목코드 정규화(fnClassify와 같은 규칙: 대문자, 숫자만 6자리 미만이면 앞 0 채움, 6자리 [0-9A-Z]만 인정).
            // 이미 6자리 대문자·숫자 텍스트면 그대로(큰 표에서 행마다 부르므로 빠른 길)
            Code = (v as any) as nullable text =>
                if v is text and Text.Length(v) = 6 and Text.Select(v, CodeChars) = v then v
                else
                    let
                        t0 = Txt(v),
                        t1 = if t0 = null then null else Text.Upper(t0),
                        t2 = if t1 <> null and Text.Length(t1) < 6 and Text.Select(t1, Digits) = t1 then Text.PadStart(t1, 6, "0") else t1
                    in
                        if t2 <> null and Text.Length(t2) = 6 and Text.Select(t2, CodeChars) = t2 then t2 else null,
            Num = (v as any) as nullable number =>
                if v = null then null
                else if v is number then (if Number.IsNaN(v) then null else v)
                else if v is text then fnNum(v)
                else null,
            Day = (v as any) as nullable date =>
                if v = null then null
                else if v is date then v
                else if v is datetime then Date.From(v)
                else if v is datetimezone then Date.From(DateTimeZone.RemoveZone(v))
                else if v is number then (try Date.From(v) otherwise null)
                else if v is text then
                    (let t = Text.Trim(v) in
                        if Text.Length(t) = 8 and Text.Select(t, Digits) = t then fnYmd(t)
                        else try Date.FromText(Text.Start(t, 10), "en-US") otherwise null)
                else null,
            // 결산년월 → "yyyymm" (텍스트·숫자·날짜 모두 받음 — T_Fin과 같은 규칙)
            YmText = (v as any) as nullable text =>
                let
                    t = if v = null then null
                        else if v is number then Number.ToText(Number.Round(v), "0", "en-US")
                        else if v is date or v is datetime then Date.ToText(Date.From(v), [Format = "yyyyMM", Culture = "en-US"])
                        else Text.Select(try Text.From(v) otherwise "", Digits),
                    m = if t <> null and Text.Length(t) = 6 then Number.From(Text.End(t, 2)) else null
                in
                    if m <> null and m >= 1 and m <= 12 then t else null,
            F = (r as nullable record, k as text) as any => if r = null then null else Record.FieldOrDefault(r, k, null),
            Read = (name as text) as any => try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null,
            Div = (a as nullable number, b as nullable number) as nullable number =>
                if a = null or b = null or b <= 0 then null else a / b,
            Chg = (a as nullable number, b as nullable number) as nullable number =>
                if a = null or b = null or b <= 0 then null else a / b - 1,

            // ---------------------------------------------------------------- 페이지 행
            PageRows = List.Buffer(List.Select(List.Transform(Table.ToRecords(fnPageRows()), (r) =>
                           [c = Code(r[종목코드]), g = Txt(r[구분]), u = Txt(r[유니버스])]), each [c] <> null)),
            PageSet = Record.FromList(List.Repeat({true}, List.Count(PageRows)), List.Transform(PageRows, each [c])),
            // 원천 표 읽기: 필요한 열만(없는 열은 null) + 정규화한 종목코드 열 __c, 페이지 종목 행만 남김(해시 집합 — 행 순서 유지).
            // 표가 없으면 같은 열의 빈 표. 행을 레코드 목록으로 바꿔 함수를 돌리는 방식보다 훨씬 빠르다(tblClass 약 3,800행 등, 실측)
            PageTable = (name as text, cols as list) as table =>
                let
                    t0 = Read(name),
                    t1 = if t0 is table then Table.SelectColumns(t0, cols, MissingField.UseNull) else #table(cols, {}),
                    t2 = Table.AddColumn(t1, "__c", each Code([종목코드]), type text)
                in
                    Table.SelectRows(t2, each [__c] <> null and Record.HasFields(PageSet, [__c])),
            // 종목코드별 첫 행 레코드(같은 종목이 여러 행이면 표에서 먼저 나온 행)
            FirstBy = (t as table) as record =>
                let g = Table.Group(t, {"__c"}, {{"r", each Table.First(_), type record}})
                in Record.FromList(g[r], g[__c]),
            ByCode = (name as text, cols as list) as record => FirstBy(PageTable(name, cols)),

            // ---------------------------------------------------------------- 분류
            Univ = ByCode("tblUniverse", {"종목코드", "종목명", "시장", "유의사항", "KRX업종"}),
            Contest = ByCode("tblContest", {"종목코드", "종목명", "시장"}),
            ClsIn = #table({"종목코드", "KRX업종"}, List.Transform(PageRows, (p) => {p[c], Txt(F(Record.FieldOrDefault(Univ, p[c], null), "KRX업종"))})),
            Cls = FirstBy(Table.AddColumn(fnClassify(ClsIn), "__c", each [종목코드], type text)),
            Klass = ByCode("tblClass", {"종목코드", "KSIC 세분류", "주요상품", "기업집단"}),
            // KSIC "C26420.이름" → "이름" (앞부분이 2~8자 영숫자 코드일 때만 뗌)
            KsicName = (v as any) as nullable text =>
                let
                    t = Txt(v),
                    p = if t = null then -1 else Text.PositionOf(t, "."),
                    head = if p > 0 then Text.Start(t, p) else null,
                    tail = if p > 0 then Text.Trim(Text.Middle(t, p + 1)) else null,
                    isCode = head <> null and Text.Length(head) >= 2 and Text.Length(head) <= 8
                             and Text.Select(head, {"0".."9", "A".."Z", "a".."z"}) = head
                in
                    if t = null then null else if isCode and tail <> "" then tail else t,

            // 형식 붙이기(열 이름 → 형식 목록): Excel이 돌려준 형식 없는 값 → 형식 값, 바꿀 수 없는 값(오류)은 null
            Typed = (t as table, types as list) as table =>
                Table.ReplaceErrorValues(Table.TransformColumnTypes(t, types), List.Transform(types, each {_{0}, null})),

            // ---------------------------------------------------------------- 원천 표 (종목코드별 첫 행 레코드 — 값 형식은 행에서 변환)
            Px = ByCode("tblPxMetrics", {"종목코드", "가격기준일", "현재가"} & Periods & StripCols
                            & {"거래대금60일억", "거래량60일천주", "고52주대비", "시가총액억", "상태"}),
            Flow = ByCode("tblFlowU", {"종목코드", "기준일", "외국인1D억", "외국인1W억", "외국인1M억", "기관1D억", "기관1W억", "기관1M억"}),
            Tgt = ByCode("tblTarget", {"종목코드", "목표가평균", "증권사수", "목표가평균_1M전", "목표가평균_3M전",
                                       "최근의견일", "최근증권사", "최근의견"}),
            Est = ByCode("tblEst", {"종목코드", "추정일", "FY1", "FY1_EPS", "FY2", "FY2_EPS"}),
            Csl = ByCode("tblCSL", {"종목코드", "기준일", "신용잔고율", "신용잔고율1M변화", "공매도비중5일", "대차잔고1M변화율"}),
            // 실적: 종목별 분기 행 목록(결산년월 내림차순, 같은 결산년월은 한 행). 형식 변환·중복 제거·정렬은 표 전체에 한 번만
            // (정렬 뒤 버퍼로 순서를 고정 — 종목별 묶음 안의 행 순서 = 결산년월 내림차순)
            FinT = Table.Buffer(Table.Sort(Table.Distinct(Table.RenameColumns(Table.SelectRows(Table.TransformColumns(
                       Typed(PageTable("tblFin", {"종목코드", "결산년월", "공개기준일", "매출", "영업이익", "순이익", "EPS", "BPS", "ROE"}),
                             {{"공개기준일", type date}, {"매출", type number}, {"영업이익", type number}, {"순이익", type number},
                              {"EPS", type number}, {"BPS", type number}, {"ROE", type number}}),
                       {{"결산년월", YmText, type text}}), each [결산년월] <> null),
                       {{"결산년월", "ym"}, {"공개기준일", "pub"}, {"매출", "sales"}, {"영업이익", "op"}, {"순이익", "ni"},
                        {"EPS", "eps"}, {"BPS", "bps"}, {"ROE", "roe"}}), {"__c", "ym"}),
                       {{"__c", Order.Ascending}, {"ym", Order.Descending}})),
            FinG = Table.Group(FinT, {"__c"}, {{"rs", each List.Buffer(Table.ToRecords(Table.SelectColumns(_,
                       {"ym", "pub", "sales", "op", "ni", "eps", "bps", "roe"}))), type list}}),
            FinBy = Record.FromList(FinG[rs], FinG[__c]),
            FinOf = (c as text) as list => Record.FieldOrDefault(FinBy, c, {}),
            // 이벤트: 오늘 이후(오늘 포함) 행만, 종목별 {일자, 종류} 목록(표 순서). 일자·종목코드가 빈 행(실패 표시)은 빠진다
            EvT = Table.SelectRows(Typed(PageTable("tblEvents", {"일자", "종목코드", "종류"}), {{"일자", type date}}),
                      each [일자] <> null and [일자] >= Today),
            EvG = Table.Group(EvT, {"__c"}, {{"rs", each List.Buffer(List.Transform(Table.ToRecords(_), (r) => [d = r[일자], k = Txt(r[종류])])), type list}}),
            EvBy = Record.FromList(EvG[rs], EvG[__c]),
            NextEvent = (c as text) as record =>
                let
                    ev = Record.FieldOrDefault(EvBy, c, {}),
                    d0 = if List.IsEmpty(ev) then null else List.Min(List.Transform(ev, each [d])),
                    kinds = if d0 = null then {} else List.Distinct(List.RemoveNulls(List.Transform(List.Select(ev, each [d] = d0), each [k])))
                in
                    [t = if d0 = null then null
                         else "D-" & Text.From(Duration.Days(d0 - Today)) & (if List.IsEmpty(kinds) then "" else " " & Text.Combine(kinds, "·")),
                     d = d0],

            // ---------------------------------------------------------------- 세션 달력 (T_PxMetrics와 같은 달력 → 같은 기준 세션)
            StoreDates = () as list =>
                let
                    DatesOf = (name as text) as list =>
                        let t = Read(name)
                        in if not (t is table) or not List.ContainsAll(Table.ColumnNames(t), {"일자", "종가"}) then {}
                           else List.Distinct(Table.SelectRows(Typed(Table.SelectColumns(t, {"일자", "종가"}), {{"일자", type date}, {"종가", type number}}),
                                    each [일자] <> null and [종가] <> null)[일자]),
                    Own = DatesOf("tblPxStore"),
                    Src = if List.IsEmpty(Own) then DatesOf("tblPxStoreSeed") else Own
                in
                    List.Sort(List.Select(Src, each _ < Today or (_ = Today and Ctl[Time] >= #time(9, 0, 0)))),
            CalSes = fnPxCalendar(Ctl[Now])[Dates],
            Cal = List.Buffer(if not List.IsEmpty(CalSes) then CalSes else StoreDates()),
            KeyOf = (d as date) as text => Date.ToText(d, [Format = "yyyyMMdd", Culture = "en-US"]),
            CalPos = Record.FromList(List.Positions(Cal), List.Transform(Cal, KeyOf)),
            PosOf = (d as nullable date) as nullable number => if d = null then null else Record.FieldOrDefault(CalPos, KeyOf(d), null),
            Years = List.Distinct(List.Transform(Cal, Date.Year)),
            YearLast = Record.FromList(List.Transform(Years, (y) => List.Last(List.Select(List.Positions(Cal), each Date.Year(Cal{_}) = y))),
                                       List.Transform(Years, Text.From)),
            Back = [#"1W" = 5, #"1M" = 21, #"3M" = 63, #"6M" = 126, #"1Y" = 250],
            // 스냅샷: 기준 세션 창에 들 수 있는 행만 — 가장 이른 가격기준일에서 24세션(1M 기준 21세션 + 가까운 스냅샷 3세션) 전 이후.
            // 날짜를 먼저 걸러 이력(1년에 약 2.8만 행)이 늘어도 종목별 묶기·정규화는 창 안의 행에만 한다. 종목별 표(일자·FwdPER)
            MinP0 = List.Min(List.RemoveNulls(List.Transform(Record.FieldValues(Px), each PosOf(Day(F(_, "가격기준일"))))), null),
            SnapLo = if MinP0 = null then null else Cal{List.Max({0, MinP0 - Back[#"1M"] - 3})},
            SnapT =
                if SnapLo = null then #table({"일자", "종목코드", "FwdPER", "__c"}, {})
                else
                    let
                        t0 = Read("tblEstSnap"),
                        t1 = if t0 is table then Table.SelectColumns(t0, {"일자", "종목코드", "FwdPER"}, MissingField.UseNull)
                             else #table({"일자", "종목코드", "FwdPER"}, {}),
                        t2 = Table.SelectRows(Typed(t1, {{"일자", type date}, {"FwdPER", type number}}), each [일자] <> null and [일자] >= SnapLo),
                        t3 = Table.AddColumn(t2, "__c", each Code([종목코드]), type text)
                    in
                        Table.SelectRows(t3, each [__c] <> null and Record.HasFields(PageSet, [__c])),
            SnapG = Table.Group(SnapT, {"__c"}, {{"t", each Table.Buffer(_), type table}}),
            SnapBy = Record.FromList(SnapG[t], SnapG[__c]),

            // ---------------------------------------------------------------- 실적 도우미
            S = try fnSettings() otherwise [],
            SetNum = (k as text, dflt as number) as number =>
                let n = Num(try Record.FieldOrDefault(S, k, null) otherwise null) in if n = null or n < 0 then dflt else Number.Round(n),
            LagQ = SetNum("fin_lag_q", 45),
            LagY = SetNum("fin_lag_y", 90),
            MonthIdx = (ym as text) as number => Number.From(Text.Start(ym, 4)) * 12 + Number.From(Text.End(ym, 2)) - 1,
            YmOfIdx = (i as number) as text =>
                Text.PadStart(Text.From(Number.IntegerDivide(i, 12)), 4, "0") & Text.PadStart(Text.From(Number.Mod(i, 12) + 1), 2, "0"),
            MonthEnd = (ym as text) as date => Date.EndOfMonth(#date(Number.From(Text.Start(ym, 4)), Number.From(Text.End(ym, 2)), 1)),
            LagOf = (x as record) as nullable number => if x[pub] = null then null else Duration.Days(x[pub] - MonthEnd(x[ym])),
            // 날짜 d에 알려진 분기(공개기준일 ≤ d) — fin이 결산년월 내림차순이므로 결과도 내림차순(첫 행 = 가장 최근 분기).
            // 여러 번 읽으므로 버퍼(목록 식은 읽을 때마다 다시 계산됨)
            Known = (fin as list, d as date) as list => List.Buffer(List.Select(fin, each [pub] <> null and [pub] <= d)),
            // 분기 행 목록(결산년월 내림차순)의 첫 행과 그 앞 3개 분기(결산년월이 정확히 3개월씩 이어짐)의 분기 단독 EPS 합.
            // 넷 중 하나라도 행이 없거나 EPS가 비면 null
            TtmOf = (rows as list) as nullable number =>
                let
                    byYm = Record.FromList(rows, List.Transform(rows, each [ym])),
                    i0 = if List.IsEmpty(rows) then null else MonthIdx(rows{0}[ym]),
                    four = if i0 = null then {} else List.Transform({0, 1, 2, 3}, (j) => Record.FieldOrDefault(byYm, YmOfIdx(i0 - 3 * j), null)),
                    ok = List.Count(four) = 4 and List.AllTrue(List.Transform(four, each _ <> null and _[eps] <> null))
                in
                    if ok then List.Sum(List.Transform(four, each [eps])) else null,
            // 분기 행 목록의 첫 행(가장 최근 분기)의 BPS
            BpsOf = (rows as list) as nullable number => if List.IsEmpty(rows) then null else rows{0}[bps],
            // 밸류 변화용(공시 지연 규칙): 날짜 d에 알려진 분기만으로 계산
            TtmEps = (fin as list, d as date) as nullable number => TtmOf(Known(fin, d)),
            BpsAt = (fin as list, d as date) as nullable number => BpsOf(Known(fin, d)),
            QLabel = (fin as list) as nullable text =>
                let
                    x = if List.IsEmpty(fin) then null else fin{0},
                    y = if x = null then null else Text.Start(x[ym], 4),
                    m = if x = null then null else Number.From(Text.End(x[ym], 2)),
                    lagsDiffer = LagY <> LagQ,
                    isQ4 = lagsDiffer and x <> null and LagOf(x) = LagY,
                    earlierQ4 = if not lagsDiffer or x = null then {} else List.Select(List.Skip(fin, 1), each LagOf(_) = LagY),
                    f = if List.IsEmpty(earlierQ4) then null else Number.From(Text.End(earlierQ4{0}[ym], 2)),
                    d = if f = null then null else Number.Mod(m - f + 12, 12),
                    q = if x = null then null
                        else if isQ4 then 4
                        else if f <> null then (if Number.Mod(d, 3) = 0 and d <> 0 then d / 3 else null)
                        else if List.Contains({3, 6, 9, 12}, m) then m / 3
                        else null
                in
                    if x = null then null else if q = null then y & "." & Text.End(x[ym], 2) else y & ".Q" & Text.From(q),
            Yoy = (cur as nullable number, prv as nullable number) as nullable number =>
                if cur = null or prv = null or prv <= 0 then null else cur / prv - 1,
            OpState = (cur as nullable number, prv as nullable number) as nullable text =>
                if cur = null or prv = null then null
                else if prv <= 0 and cur > 0 then "흑자전환"
                else if prv > 0 and cur <= 0 then "적자전환"
                else if prv <= 0 then "적자지속"
                else "흑자지속",

            // ---------------------------------------------------------------- 종목 1행 (값 순서 = Cols)
            Row = (p as record) as list =>
                let
                    c = p[c],
                    u = Record.FieldOrDefault(Univ, c, null),
                    k = Record.FieldOrDefault(Contest, c, null),
                    cl = Record.FieldOrDefault(Cls, c, null),
                    kl = Record.FieldOrDefault(Klass, c, null),
                    name = let a = Txt(F(u, "종목명")) in if a <> null then a else Txt(F(k, "종목명")),
                    market = let a = Txt(F(u, "시장")) in if a <> null then a else Txt(F(k, "시장")),

                    // 가격 지표
                    x = Record.FieldOrDefault(Px, c, null),
                    price = Num(F(x, "현재가")),
                    mcap = Num(F(x, "시가총액억")),
                    d0 = Day(F(x, "가격기준일")),
                    Ret = (kk as text) as nullable number => Num(F(x, kk)),
                    pxState = Txt(F(x, "상태")),
                    status = if x = null then "데이터 없음" else if pxState <> null then pxState else if price = null then "데이터 없음" else "OK",
                    p0 = PosOf(d0),
                    BaseDate = (n as number) as nullable date => if p0 = null or p0 - n < 0 then null else Cal{p0 - n},
                    YtdDate = if d0 = null then null
                              else let yp = Record.FieldOrDefault(YearLast, Text.From(Date.Year(d0) - 1), null) in if yp = null then null else Cal{yp},
                    BaseClose = (kk as text) as nullable number =>
                        let r = Ret(kk) in if price = null or r = null or 1 + r = 0 then null else price / (1 + r),

                    // 수급
                    fl = Record.FieldOrDefault(Flow, c, null),
                    Fv = (kk as text) as nullable number => Num(F(fl, kk)),
                    Pct = (kk as text) as nullable number => let v = Fv(kk) in if v = null or mcap = null or mcap <= 0 then null else v / mcap,

                    // 실적·후행 밸류
                    fin = FinOf(c),
                    lat = if List.IsEmpty(fin) then null else fin{0},
                    prevYm = if lat = null then null else Text.From(Number.From(lat[ym]) - 100),
                    prv = if lat = null then null else List.First(List.Select(fin, each [ym] = prevYm), null),
                    // 후행PER·PBR: 받은 자료의 최근 분기 기준(공개기준일과 무관)
                    epsLatest = TtmOf(fin),
                    bpsLatest = BpsOf(fin),
                    // 밸류 변화의 분자: 오늘 알려진 값(공시 지연 규칙 — 분모의 기준일 값과 같은 기준)
                    epsKnown = TtmEps(fin, Today),
                    bpsKnown = BpsAt(fin, Today),
                    ValChg = (kk as text, isPer as logical) as nullable number =>
                        let
                            bd = if kk = "YTD" then YtdDate else BaseDate(Record.Field(Back, kk)),
                            cb = BaseClose(kk),
                            vNow = if isPer then epsKnown else bpsKnown,
                            vBase = if bd = null then null else if isPer then TtmEps(fin, bd) else BpsAt(fin, bd)
                        in
                            if bd = null or cb = null or price = null or vNow = null or vNow <= 0 or vBase = null or vBase <= 0 then null
                            else (price / vNow) / (cb / vBase) - 1,

                    // 목표주가
                    tg = Record.FieldOrDefault(Tgt, c, null),
                    m0 = Num(F(tg, "목표가평균")),
                    brokers = let n = Num(F(tg, "증권사수")) in if n = null or n <= 0 then null else n,

                    // KIS 추정·Fwd PER
                    es = Record.FieldOrDefault(Est, c, null),
                    fy1 = Txt(F(es, "FY1")),
                    e1 = Num(F(es, "FY1_EPS")),
                    fy2 = Txt(F(es, "FY2")),
                    e2 = Num(F(es, "FY2_EPS")),
                    estDay = Day(F(es, "추정일")),
                    fy1Only = e1 <> null and e2 = null,
                    fwd = if es = null then null else fnFwdEps(fy1, e1, e2, Today),
                    fper = if price = null or fwd = null or fwd <= 0 then null else price / fwd,
                    FwdChg = (n as number, kk as text) as nullable number =>
                        let
                            bd = BaseDate(n),
                            lo = if bd = null then null else Cal{List.Max({0, p0 - n - 3})},
                            snapT = if bd = null then null else Record.FieldOrDefault(SnapBy, c, null),
                            inWin = if snapT = null then null else Table.SelectRows(snapT, each [일자] >= lo and [일자] <= bd),
                            best = if inWin = null or Table.IsEmpty(inWin) then null else Table.Max(inWin, "일자"),
                            basePer =
                                if best <> null then best[FwdPER]
                                else if estDay <> null and bd <> null and estDay < bd then
                                    (let cb = BaseClose(kk), fb = fnFwdEps(fy1, e1, e2, bd) in if cb = null or fb = null or fb <= 0 then null else cb / fb)
                                else null
                        in
                            if fper = null or bd = null or basePer = null or basePer <= 0 then null else fper / basePer - 1,

                    // 신용·공매도·대차, 이벤트
                    cs = Record.FieldOrDefault(Csl, c, null),
                    ne = NextEvent(c)
                in
                    {c, name, market, p[g], p[u], F(cl, "NICS 업종"), F(cl, "대테마"),
                     mcap, Num(F(x, "거래대금60일억")), Num(F(x, "거래량60일천주")),
                     price, Num(F(x, "고52주대비")), d0}
                    & List.Transform(Periods, Ret)
                    & List.Transform(StripCols, (s) => Num(F(x, s)))
                    & {Pct("외국인1D억"), Pct("외국인1W억"), Pct("외국인1M억"), Fv("외국인1D억"), Fv("외국인1W억"), Fv("외국인1M억"),
                       Pct("기관1D억"), Pct("기관1W억"), Pct("기관1M억"), Fv("기관1D억"), Fv("기관1W억"), Fv("기관1M억"), Day(F(fl, "기준일")),
                       Div(price, epsLatest), Div(price, bpsLatest),
                       fper, FwdChg(5, "1W"), FwdChg(21, "1M"), fwd, estDay, fy1, e1,
                       if fy1Only then "FY1만" else fy2, if fy1Only then null else e2,
                       m0, Chg(m0, price), brokers, Chg(m0, Num(F(tg, "목표가평균_1M전"))), Chg(m0, Num(F(tg, "목표가평균_3M전"))),
                       Day(F(tg, "최근의견일")), Txt(F(tg, "최근증권사")), Txt(F(tg, "최근의견")),
                       QLabel(fin),
                       Yoy(F(lat, "sales"), F(prv, "sales")), Yoy(F(lat, "op"), F(prv, "op")), Yoy(F(lat, "ni"), F(prv, "ni")),
                       OpState(F(lat, "op"), F(prv, "op")), F(lat, "roe"),
                       ValChg("3M", true), ValChg("6M", true), ValChg("YTD", true), ValChg("1Y", true),
                       ValChg("3M", false), ValChg("6M", false), ValChg("YTD", false), ValChg("1Y", false),
                       Num(F(cs, "신용잔고율")), Num(F(cs, "신용잔고율1M변화")), Num(F(cs, "공매도비중5일")), Num(F(cs, "대차잔고1M변화율")),
                       Day(F(cs, "기준일")),
                       ne[t], ne[d],
                       F(cl, "NICS 대분류"), F(cl, "NICS 세부"), F(cl, "세부테마"), KsicName(F(kl, "KSIC 세분류")),
                       Txt(F(kl, "주요상품")), Txt(F(kl, "기업집단")), F(cl, "분류출처"),
                       Txt(F(u, "유의사항")), status, Now},
            // 종목 행마다 모든 칸을 강제 평가해 계산 오류를 그 행의 상태로(목록·표 버퍼는 칸 안 오류를 올리지 않고 빈칸으로 적재하므로)
            SafeRow = (p as record) as list =>
                let
                    r = try (let v = List.Buffer(Row(p)), n = List.Count(List.Select(v, each _ = null or _ <> null)) in if n = NCols then v else error "열 수 불일치"),
                    why = if r[HasError] then (try Text.Start(Text.From(r[Error][Message]), 100) otherwise "사유 불명") else null
                in
                    if r[HasError] then {p[c], null, null, p[g], p[u]} & List.Repeat({null}, NCols - 7) & {"오류: 계산 실패(" & why & ")", Now}
                    else r[Value],
            Rows = List.Transform(PageRows, SafeRow),
            // 정렬: 대회 → 유니버스 밖, 각 묶음 안에서 시가총액억 내림차순(빈칸은 뒤), 같으면 페이지 행 순서
            Indexed = Table.AddIndexColumn(#table(Cols, Rows), "__ord", 0, 1, Int64.Type),
            Keyed = Table.AddColumn(Table.AddColumn(Indexed,
                        "__grp", each if [유니버스] = "대회" then 0 else 1, Int64.Type),
                        "__nomc", each if [시가총액억] = null then 1 else 0, Int64.Type),
            Sorted = Table.Sort(Keyed, {{"__grp", Order.Ascending}, {"__nomc", Order.Ascending}, {"시가총액억", Order.Descending},
                                        {"__ord", Order.Ascending}}),
            Out = Table.TransformColumnTypes(Table.RemoveColumns(Sorted, {"__ord", "__grp", "__nomc"}), Types)
        in
            Out,
    Tried = try Table.Buffer(if Already then KeptOwn else Compute()),
    Reason = if not Tried[HasError] then null
             else let m = Tried[Error][Message] in if m = null or m = "" then "사유 불명" else m,
    KeptOnFail = if Tried[HasError] then fnFallback("tblCompany", Cols, Reason) else null,
    Warn = #table(Cols, {List.Repeat({null}, NCols - 2) & {"오류: " & Text.Start(if Reason = null then "" else Reason, 120), DateTime.LocalNow()}}),
    // fnFallback·경고 행은 형식이 없는 표 → 열 형식을 다시 지정(날짜·조회시각이 일련번호로 보이지 않게)
    Final =
        if not Tried[HasError] then Tried[Value]
        else Table.TransformColumnTypes(if Table.RowCount(KeptOnFail) > 0 then KeptOnFail else Warn, Types)
in
    Final
