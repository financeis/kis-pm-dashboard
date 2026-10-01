// T_Est : KIS 추정실적 — FY1·FY2 EPS (spec R9·R20, 2026-10-01 T9) → tblEst  [[전체] 버튼 전용 — mode = full일 때만 KIS 호출]
//  - 원천: 국내주식 종목추정실적[국내주식-187] HHKST668300C0, GET /uapi/domestic-stock/v1/quotations/estimate-perform, 입력 SHT_CD = 종목코드.
//    한국투자증권 리서치의 자체 추정이다(컨센서스 아님 — 이 종목군의 약 21%만 제공). 저장소 예제(chk_estimate_perform.py)의
//    COLUMN_MAPPING은 ELW 표를 베낀 것이라 틀렸다. 응답 구조(2026-09-30·10-01 실측, 표본 15종목):
//      output1(레코드)  estdate = KIS 보고서 일자(YYYYMMDD), sht_cd = "A" + 코드, item_kor_nm, name1(애널리스트), rcmd_name(KIS 의견)
//      output4(목록)    dt = data1…data5 열의 결산년월 라벨(보통 2023.12, 2024.12, 2025.12, 2026.12E, 2027.12E — 오래된 추정은
//                       2025.12E에서 끝나기도 함) → 연도는 항상 이 라벨로 대응한다(열 위치를 연도로 가정하지 않음)
//      output2(목록)    행 순서 매출액, 매출증가율, 영업이익, 영업이익증가율, 순이익(지배), 순이익증가율 — 이 쿼리는 쓰지 않음
//      output3(목록)    행 순서 EBITDA, EPS, EPS증가율, PER, EV/EBITDA, ROE, 부채비율, 이자보상배율 → EPS = 두 번째 행(행 이름 필드 없음)
//    단위: 금액은 억원 그대로, EPS와 모든 비율·배수는 ×10으로 오므로 ÷10 한다(예: 삼성전자 2023.12 EPS "21310.0" → 2,131원).
//    행 잘림: 값이 N/A인 행에서 그 행부터 끝까지 빠진다 — 적자 연도가 있으면 output3이 3행(EBITDA·EPS·EPS증가율)만 오고
//    (SK하이닉스·롯데케미칼·에코프로비엠), 은행·증권·카드는 output3이 빈 목록이다(KB금융) → EPS 없음.
//    추정이 없는 종목은 rt_cd 0에 빈 목록·estdate 빈칸으로 온다(오류 아님). 연속 조회 없음(응답 tr_cont 빈칸) — 종목당 1회.
//  - 열: 종목코드, 추정일, FY1, FY1_EPS, FY2, FY2_EPS, 전체조사시각, 상태, 조회시각
//      추정일   output1.estdate(date). 보고서가 있으면 EPS가 없어도(은행 등) 채운다. 없으면 빈칸.
//      FY1      dt 라벨 가운데 결산년월(라벨의 숫자 앞 6자리 = yyyymm)이 기준일 D(조회한 날)의 연월 이상인 것 중 가장 이른 것 —
//               아직 끝나지 않은 회계연도(실적 라벨이든 추정 라벨이든 같은 규칙). 12월 결산이 아니어도 같다: 2026년 10월에 3월 결산이면
//               "2027.03E", 9월 결산이면 "2027.09E"(2026.09는 이미 끝남), 10월 결산이면 "2026.10E"(이번 달 결산 포함).
//               값은 KIS 라벨 글자 그대로(예: "2026.12E") — 결산월은 쓰는 쪽(fnFwdEps)이 라벨의 숫자에서 다시 읽는다.
//      FY1_EPS  FY1 열의 EPS ÷ 10(원)
//      FY2      dt에서 FY1 다음(결산년월이 더 늦은 것 중 가장 이른) 라벨, FY2_EPS = 그 열의 EPS ÷ 10.
//               FY2 라벨이나 그 EPS가 없으면 FY2·FY2_EPS 둘 다 빈칸 → 쓰는 쪽이 'FY1만'으로 표시하고 Fwd EPS = FY1 EPS(fnFwdEps).
//      FY1이 없거나(미제공·오래된 추정) FY1 EPS가 없으면(은행 등 output3이 비었거나 1행) FY1·FY1_EPS·FY2·FY2_EPS 모두 빈칸이고
//      상태 '추정 없음'(정상 값). Fwd EPS·Fwd PER는 이 표에 두지 않는다 — 날짜마다 fnFwdEps(FY1, FY1_EPS, FY2_EPS, 날짜)로 계산한다
//      (스냅샷 T_EstSnap = 가격기준일, 대회종목 표 = 오늘·기준 세션).
//  조사 범위 (spec R20, mode = full) — T_Target과 같은 규칙
//    - 주간 = 페이지 행(fnPageRows) 전체 조회 후 모든 행의 전체조사시각 = 지금. 조건: 직전 표에 전체조사시각이 없음(첫 실행),
//      또는 그 날짜로부터 weekly_days(기본 7)일 이상 지남(날짜 차이로 셈 — 누르는 시각이 달라도 같은 요일이면 돌게),
//      또는 설정 force_weekly = Y. (force_weekly를 N으로 되돌리는 일은 [전체] 매크로가 한다.)
//    - 매일 = 직전 행의 FY1_EPS가 있는 종목 + 직전 표에 없는 새 종목 + 직전 상태가 '오류'인 종목(실패한 조사를 다시 시도)만 조회.
//      나머지(직전 조사에서 추정이 없던 종목)는 직전 행을 그대로 두되(조회시각 유지) 상태만 '추정 없음'으로 맞춘다 — 직전 실패의
//      '이전 데이터' 표시가 남아 단계 판정이 실패로 보이지 않게. 전체조사시각은 직전 값 그대로(모든 행 같은 값).
//    - 행 = 지금의 페이지 행(순서도 fnPageRows). 페이지 행에서 빠진 종목의 직전 행은 버린다(스냅샷 tblEstSnap에는 남는다).
//    - 직전 표 = 자기 표 tblEst → 없거나 비었으면 정적 표 tblEstSeed(이관 때 빌더가 채움) → 없으면 빈 표.
//    - 같은 [전체] 실행 안에서 다시 평가되면(자기 표의 가장 늦은 조회시각이 tblRunCtl started 이후) 조회하지 않고 자기 표를 돌려준다
//      — 새로 고침 직후 Excel의 백그라운드 재평가가 매일 조사만큼 KIS를 또 부르는 것을 막음(아래 Main 주석). started가 비면 확인 안 함.
//    - 기준일 D는 실제 오늘(DateTime.LocalNow)이다. 시험용 시계(tblRunCtl now_override)는 쓰지 않는다(plan §0: T_Sessions·T_PxStore
//      전용). 주간 규칙 시험은 tblEst의 전체조사시각 셀을 바꾸거나 force_weekly·weekly_days 설정으로 한다.
//  mode가 full이 아니면(build·quick·sector·analysis·빈칸·알 수 없는 값) KIS를 부르지 않고 직전 표(자기 표 → Seed → 빈 표)를
//    그대로 돌려준다(상태도 그대로, 빈 표도 열은 모두 갖춤). 설정·페이지 행도 읽지 않는다.
//  실패: 종목별 조회 오류 → 그 종목 행만 '오류: 사유'(직전 값과 조회시각 유지; 직전 행이 없으면 값은 빈칸, 조회시각은 지금).
//    조회한 종목이 모두 실패하면 오류로 올려 직전 표를 유지(fnFallback — 상태 '이전 데이터(갱신 실패: 사유)', 전체조사시각도 그대로라
//    주간 조사였다면 다음 [전체]에서 다시 함). 직전 표를 다시 읽은 값은 열 형식이 없어 날짜가 일련번호로 보이므로 형식을 다시 지정한다.
//  호출 수: 종목당 1회. 주간 = 페이지 행 수(약 527), 매일 = 추정 EPS가 있는 종목 수(추정 제공 약 21% ≈ 110 이하 — 은행 등 EPS 없는
//    종목은 주간에만) + 새 종목 + 직전 오류 종목. 2026-10-01 실측 표본 15종목: 주간 15회, 매일 6회.
let
    Cols = {"종목코드", "추정일", "FY1", "FY1_EPS", "FY2", "FY2_EPS", "전체조사시각", "상태", "조회시각"},
    Types = {{"종목코드", type text}, {"추정일", type date}, {"FY1", type text}, {"FY1_EPS", type number},
             {"FY2", type text}, {"FY2_EPS", type number}, {"전체조사시각", type datetime}, {"상태", type text},
             {"조회시각", type datetime}},
    Schema = type table [종목코드 = text, 추정일 = nullable date, FY1 = nullable text, FY1_EPS = nullable number,
                         FY2 = nullable text, FY2_EPS = nullable number, 전체조사시각 = nullable datetime,
                         상태 = nullable text, 조회시각 = nullable datetime],
    // 행 목록의 값 위치(Cols 순서: 0 종목코드, 1 추정일, 2 FY1, 3 FY1_EPS, 4 FY2, 5 FY2_EPS, 6 전체조사시각, 7 상태, 8 조회시각)
    iEps1 = 3, iFull = 6, iState = 7, iStamp = 8,

    // 셀 값 → 형식 있는 값 (직전 표·Seed는 Excel에서 다시 읽어 열 형식이 없다; 변환할 수 없는 값은 null)
    Txt = (v as any) as nullable text =>
        if v = null then null else let t = try Text.Trim(Text.From(v)) otherwise null in if t = null or t = "" then null else t,
    Num = (v as any) as nullable number =>
        if v = null then null else if v is number then v else if v is text then fnNum(v) else null,
    Day = (v as any) as nullable date =>
        if v = null then null
        else if v is date then v
        else if v is datetime then Date.From(v)
        else if v is number then (try Date.From(v) otherwise null)
        else if v is text then
            (let t = Text.Trim(v) in
                if Text.Length(t) = 8 and Text.Select(t, {"0".."9"}) = t then fnYmd(t)
                else try Date.FromText(Text.Start(t, 10), "en-US") otherwise null)
        else null,
    Stamp = (v as any) as nullable datetime =>
        if v = null then null
        else if v is datetime then v
        else if v is date then DateTime.From(v)
        else if v is number then (try DateTime.From(v) otherwise null)
        else if v is text then (try DateTime.FromText(Text.Trim(v), "en-US") otherwise null)
        else null,
    // 종목코드 정규화(fnClassify와 같은 규칙: 대문자, 숫자만 6자리 미만이면 앞 0 채움, 6자리 [0-9A-Z]만 인정)
    Code = (v as any) as nullable text =>
        let
            t0 = Txt(v),
            t1 = if t0 = null then null else Text.Upper(t0),
            t2 = if t1 <> null and Text.Length(t1) < 6 and Text.Select(t1, {"0".."9"}) = t1 then Text.PadStart(t1, 6, "0") else t1
        in
            if t2 <> null and Text.Length(t2) = 6 and Text.Select(t2, {"0".."9", "A".."Z"}) = t2 then t2 else null,
    F = (r as record, k as text) as any => try Record.Field(r, k) otherwise null,
    ReadTable = (name as text) as nullable table => try Excel.CurrentWorkbook(){[Name = name]}[Content] otherwise null,

    // 직전 표 → 행 목록(값 순서 = Cols). 종목코드가 없거나 잘못된 행은 버리고, 같은 코드는 처음 행만
    ToRows = (t as nullable table) as list =>
        if t = null then {}
        else
            let
                recs = Table.ToRecords(Table.SelectColumns(t, Cols, MissingField.UseNull)),
                rows = List.Transform(recs, (r) => {Code(F(r, "종목코드")), Day(F(r, "추정일")), Txt(F(r, "FY1")),
                    Num(F(r, "FY1_EPS")), Txt(F(r, "FY2")), Num(F(r, "FY2_EPS")), Stamp(F(r, "전체조사시각")),
                    Txt(F(r, "상태")), Stamp(F(r, "조회시각"))})
            in
                List.Distinct(List.Select(rows, each _{0} <> null), each _{0}),
    OwnRows = ToRows(ReadTable("tblEst")),
    PrevRows = List.Buffer(if List.Count(OwnRows) > 0 then OwnRows else ToRows(ReadTable("tblEstSeed"))),

    // 실행 제어 tblRunCtl(키, 값): mode — 없거나 빈칸·알 수 없는 값은 build와 같음(full일 때만 조회),
    // started — 버튼 실행 시작 시각([전체] 매크로가 씀)
    RunCtl = ReadTable("tblRunCtl"),
    RunRows = if RunCtl = null then {} else Table.ToRecords(Table.SelectColumns(RunCtl, {"키", "값"}, MissingField.UseNull)),
    RunValue = (key as text) as any =>
        let hit = List.Select(RunRows, (r) => let k = Txt(F(r, "키")) in k <> null and Text.Lower(k) = key)
        in if List.IsEmpty(hit) then null else F(hit{0}, "값"),
    ModeText = Txt(RunValue("mode")),
    Mode = if ModeText = null then "build" else Text.Lower(ModeText),
    // 같은 버튼 실행 안에서 다시 평가될 때(새로 고침 직후 Excel이 쿼리를 백그라운드에서 한 번 더 평가함 — 아래 Main 주석)는
    // 조회하지 않고 방금 적재한 자기 표를 돌려준다: 자기 표의 가장 늦은 조회시각이 이번 실행의 started 이후면 이미 조사한 것.
    // started가 비어 있으면(하네스 시험 등) 이 확인은 하지 않는다.
    Started = Stamp(RunValue("started")),
    LastStamp = List.Max(List.RemoveNulls(List.Transform(OwnRows, each _{iStamp}))),
    Surveyed = Started <> null and LastStamp <> null and LastStamp >= Started,

    // dt 라벨 → 결산년월 키(연 × 12 + 월). 라벨의 숫자만 모아 앞 4자리 = 연, 다음 2자리 = 월("2026.12E" → 2026·12). 해석 못 하면 null
    YmKey = (label as nullable text) as nullable number =>
        let
            d = if label = null then "" else Text.Select(label, {"0".."9"}),
            y = if Text.Length(d) >= 6 then Number.From(Text.Start(d, 4)) else null,
            m = if y = null then null else Number.From(Text.Middle(d, 4, 2))
        in
            if y = null or m < 1 or m > 12 then null else y * 12 + m,

    Survey = () as table =>
        let
            Now = DateTime.LocalNow(),
            D = Date.From(Now),
            Cur = Date.Year(D) * 12 + Date.Month(D),
            // 설정(없거나 잘못된 값은 기본값): weekly_days 7, force_weekly N
            S = try fnSettings() otherwise [],
            Setting = (k as text) as any => try Record.FieldOrDefault(S, k, null) otherwise null,
            WeeklyDays = let v = Num(Setting("weekly_days")) in if v = null or v < 0 then 7 else v,
            Force = let v = Txt(Setting("force_weekly")) in v <> null and Text.Upper(v) = "Y",

            Codes = List.Buffer(List.Distinct(List.RemoveNulls(List.Transform(fnPageRows()[종목코드], Code)))),
            Prev = Record.FromList(PrevRows, List.Transform(PrevRows, each _{0})),
            LastFull = List.Max(List.RemoveNulls(List.Transform(PrevRows, each _{iFull}))),
            Weekly = Force or LastFull = null or Duration.Days(D - Date.From(LastFull)) >= WeeklyDays,
            FullAt = if Weekly then Now else LastFull,
            Needs = (c as text) as logical =>
                let p = Record.FieldOrDefault(Prev, c, null)
                in Weekly or p = null or p{iEps1} <> null or (p{iState} <> null and Text.StartsWith(p{iState}, "오류")),
            Targets = List.Buffer(List.Select(Codes, Needs)),

            // 종목 하나: 추정실적 조회 → {추정일, FY1, FY1_EPS, FY2, FY2_EPS}
            Fetch = (c as text) as list =>
                let
                    J = fnKisGet("/uapi/domestic-stock/v1/quotations/estimate-perform", "HHKST668300C0", [SHT_CD = c]),
                    Recs = (k as text) as list =>
                        let v = Record.FieldOrDefault(J, k, null)
                        in if v is list then List.Select(v, each _ is record) else if v is record then {v} else {},
                    O1 = let l = Recs("output1") in if List.IsEmpty(l) then [] else l{0},
                    O3 = Recs("output3"),
                    O4 = Recs("output4"),
                    EstDate = fnYmd(Record.FieldOrDefault(O1, "estdate", null)),
                    EpsRow = if List.Count(O3) >= 2 then O3{1} else null,
                    // 라벨 위치 i(0부터) ↔ 값 필드 data(i+1). 해석할 수 없는 라벨은 버린다
                    Cands = List.Select(List.Transform(List.Positions(O4), (i) =>
                                let
                                    label = Txt(Record.FieldOrDefault(O4{i}, "dt", null)),
                                    raw = if EpsRow = null then null else fnNum(Record.FieldOrDefault(EpsRow, "data" & Text.From(i + 1), null))
                                in
                                    [label = label, key = YmKey(label), eps = if raw = null then null else raw / 10]),
                            each [key] <> null),
                    Ahead = List.Sort(List.Select(Cands, each [key] >= Cur), {each [key], Order.Ascending}),
                    P1 = if List.IsEmpty(Ahead) then null else Ahead{0},
                    Later = if P1 = null then {} else List.Select(Ahead, each [key] > P1[key]),
                    P2 = if List.IsEmpty(Later) then null else Later{0},
                    Has1 = P1 <> null and P1[eps] <> null,
                    Has2 = Has1 and P2 <> null and P2[eps] <> null,
                    Vals = {EstDate, if Has1 then P1[label] else null, if Has1 then P1[eps] else null,
                            if Has2 then P2[label] else null, if Has2 then P2[eps] else null},
                    // M 지연 평가: 목록 항목은 읽을 때 평가되고, 항목 오류는 List.Buffer로도 올라오지 않는다(T_Target 실측).
                    // 그래서 응답 목록의 길이를 먼저 세고(API 오류가 여기서 남) 값마다 한 번씩 비교해, 오류를 이 함수 안 —
                    // 아래 종목별 try 안 — 에서 올린다.
                    Forced = List.Accumulate(Vals, List.Count(O3) + List.Count(O4), (n, v) => if v = null then n else n + 1)
                in
                    if Forced >= 0 then Vals else error "평가 실패",
            // 종목마다 한 번만 호출되도록 결과 목록을 버퍼로 고정(레코드 필드는 처음 읽을 때 한 번 평가)
            Results = List.Buffer(List.Transform(Targets, (c) => [code = c, r = try Fetch(c)])),
            Failed = List.Select(Results, each [r][HasError]),
            Reason = (e as record) as text =>
                let m = try Text.Trim(Text.From(e[Message])) otherwise "" in if m = "" then "조회 실패" else Text.Start(m, 200),
            // 조회한 종목이 모두 실패하면 오류로 올려 직전 표를 유지(fnFallback), 일부만 실패하면 그 종목만 오류 행
            Checked = if List.Count(Results) > 0 and List.Count(Failed) = List.Count(Results)
                      then error Error.Record("KIS 추정 조회 실패", Reason(Failed{0}[r][Error])) else Results,
            Done = Record.FromList(Checked, List.Transform(Checked, each [code])),
            Row = (c as text) as list =>
                let
                    p = Record.FieldOrDefault(Prev, c, null),
                    d = Record.FieldOrDefault(Done, c, null)
                in
                    if d = null then
                        // 매일 실행에서 조회하지 않은 종목(직전 조사에서 추정 EPS 없음): 직전 행 유지, 상태만 정상 값으로
                        {c} & List.Range(p, 1, 5) & {FullAt, if p{iEps1} = null then "추정 없음" else "OK", p{iStamp}}
                    else if d[r][HasError] then
                        (if p = null then {c, null, null, null, null, null} else {c} & List.Range(p, 1, 5))
                        & {FullAt, "오류: " & Reason(d[r][Error]), if p = null or p{iStamp} = null then Now else p{iStamp}}
                    else
                        let v = d[r][Value] in {c} & v & {FullAt, if v{2} = null then "추정 없음" else "OK", Now},
            Rows = List.Buffer(List.Transform(Codes, Row)),
            // 행 목록의 항목 오류(모두 실패 오류 포함)는 표에 '오류 행'으로만 남고 Table.Buffer로도 올라오지 않으므로(빈 행이 적재됨 —
            // T_Target 실측) 행마다 한 번씩 읽어 여기서 오류를 올린다 → 아래 Tried가 잡아 직전 표를 유지
            Counted = List.Accumulate(Rows, List.Count(Checked), (n, x) => n + List.Count(x))
        in
            if Counted >= 0 then #table(Schema, Rows) else error "평가 실패",

    // Excel은 새로 고침이 끝난 직후 쿼리를 백그라운드에서 한 번 더 평가한다 — 최상위 단계들도 따로 평가한다(T_Target 하네스 +
    // 모의 서버 실측, 동기·백그라운드 새로 고침 모두). 이 평가에서 Power Query 캐시에 없는 요청은 KIS까지 다시 나간다.
    //  ① 조사를 값으로 된 단계로 두면 build 모드 새로 고침만으로도 페이지 행 전체 조회가 나간다 → 조사는 함수로 두고
    //     full일 때만 부른다(함수 값인 단계는 평가해도 호출하지 않음).
    //  ② full 새로 고침 뒤에는 매일 조사만큼 중복 호출이 나간다 → 위 Surveyed 확인으로 막는다.
    //     (조회가 모두 실패해 직전 표로 돌아간 경우엔 조회시각이 그대로라 막지 못하고 한 번 더 시도한다.)
    //  ③ 이 재평가는 새로 고침 몇 초 뒤에 돈다. 그래서 build 새로 고침 직후 몇 초 안에 mode를 full로 바꿔 새로 고치면 build 쪽
    //     재평가가 full을 보고 함께 조사해 요청이 두 번씩 나가고, 새로 추가한 쿼리의 첫 새로 고침이 full이어도 두 번 평가된다
    //     (2026-10-01 T9 모의 서버 실측: build 1.1초 뒤 full → 15종목 30건, 10초 뒤 → 15건). [전체]는 빌드(build) 뒤 한참 지나
    //     누르므로 해당 없지만, 시험·매크로에서 T_Est를 build로 새로 고친 바로 뒤에 full로 돌리지 말 것.
    Main = if Mode = "full" and not Surveyed then Survey() else #table(Schema, PrevRows),
    // 실패(토큰·네트워크·설정·모두 실패 등) 시 오류 대화상자 대신 직전 결과를 유지하고, 열 형식을 다시 붙여 날짜 서식을 지킨다
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblEst", Cols, Tried[Error][Message]), Types)
            else Tried[Value]
in
    Final
