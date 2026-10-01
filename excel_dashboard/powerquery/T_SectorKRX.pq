// T_SectorKRX : KRX 업종지수 기간 등락 + 업종별 투자자 순매수 (업종 페이지 A) → 표 tblSectorKRX
//  [시장별 투자자매매동향(일별) 국내주식-075] FHPTJ04040000 — 업종마다 1회 호출로 약 300세션의 업종지수 종가 + 투자자별 순매수(백만원)
//  - 호출 인자(2026-10-01 실측): FID_INPUT_ISCD = 업종코드 → 지수 계열, FID_INPUT_ISCD_2 = 업종코드 → 순매수 계열,
//    FID_INPUT_ISCD_1 = KSP(KOSPI)/KSQ(KOSDAQ), FID_INPUT_DATE_1·2 = 오늘. 둘 중 한 칸에 0001을 넣으면 그쪽만 시장 전체 값이 온다.
//  - 업종 목록·업종명: KIS 업종코드 마스터 idxcode.mst.zip(T_Universe와 같은 파일, 인증 불필요)의 KOSPI(시장구분 0) 0005~0030
//    + KOSDAQ(시장구분 1) 1006~1033 코드. 규모 지수(0002~0004·1002~1004)는 제외하고, 구분은 T_Sector와 같은 규칙
//    (0021 금융·0027 제조·1009 제조 = 대분류, 나머지 = 업종).
//    T_Sector의 원천(국내업종 구분별전체시세)은 KOSDAQ 목록이 100행에서 잘리고 호출마다 구성이 달라 1009 제조·1010 건설·
//    1032 통신·1033 IT 서비스가 빠지곤 해서(실측) 목록으로 쓰지 않는다(업종명은 두 원천이 같다). 범위 안의 없는 코드(예: 1007·1012)도
//    이 API는 다른 지수 값을 순매수 0으로 돌려주므로, '응답이 오는지'로 업종 존재를 판정하지 않는다.
//  - 지수·기간 등락: 가장 최근 지수 값이 '현재'(장중이면 오늘 진행 중 값, 09:00 전에는 오늘 행을 쓰지 않음), 기준일 = 그 값의 날짜.
//    기간 등락 = 현재 ÷ 기준 세션 종가 − 1 (세션 = 이 지수 계열의 날짜): 1D=직전 세션, 1W=5, 1M=21, 3M=63, 6M=126,
//    1Y=250세션 전, YTD=기준일 전년도의 마지막 세션. 이력이 모자라면 빈칸.
//  - 순매수(억원 = 백만원 ÷ 100): 외국인 frgn_ntby_tr_pbmn, 기관계 orgn_ntby_tr_pbmn, 개인 prsn_ntby_tr_pbmn.
//    1D·1W·1M = 최근 1·5·21개 '완료' 세션의 합(종목 수급과 같은 규칙). 15:30 전에는 오늘 행(진행 중 세션)을 빼므로 장중에는
//    직전 세션까지의 합이다 — 이때 순매수는 기준일(오늘, 지수 기준)보다 하루 앞선 세션까지. 창 안에 빈 값이 있거나 세션이 모자라면 빈칸.
//  - 실행 모드(tblRunCtl 키 mode): sector·full에서만 호출(업종 수만큼 — 2026-10-01 기준 46회 — + 마스터 파일 1회).
//    build(빈칸·알 수 없는 값 포함)·quick·analysis는 호출 없이(마스터 파일도 받지 않음) 자기 표 → 정적 표 tblSectorKRXSeed → 빈 표(열·형식 모두)
//    순으로 돌려준다.
//  - 같은 버튼 실행([업종]·[전체]) 안에서 다시 평가되면(자기 표의 가장 늦은 조회시각이 tblRunCtl started 이후) 조회하지 않고 자기 표를
//    돌려준다 — 새로 고침 직후 Excel의 백그라운드 재평가가 47회를 또 부르는 것을 막음(아래 Main 주석). started가 비면 확인 안 함.
//  - 상태: 성공 행 OK, 업종 하나가 실패하면 그 행만 "오류: 사유"(값은 빈칸), 전부 실패·목록 실패는 직전 표 유지(fnFallback,
//    열 형식은 다시 입힘).
//  - 행 순서: KOSPI → KOSDAQ, 각 시장 안에서 코드 순.
let
    Columns = {"시장", "코드", "업종명", "구분", "기준일", "지수", "1D", "1W", "1M", "3M", "6M", "YTD", "1Y",
               "외국인1D억", "외국인1W억", "외국인1M억", "기관1D억", "기관1W억", "기관1M억", "개인1D억", "개인1W억", "개인1M억",
               "상태", "조회시각"},
    TextCols = {"시장", "코드", "업종명", "구분", "상태"},
    NumCols = {"지수", "1D", "1W", "1M", "3M", "6M", "YTD", "1Y", "외국인1D억", "외국인1W억", "외국인1M억",
               "기관1D억", "기관1W억", "기관1M억", "개인1D억", "개인1W억", "개인1M억"},
    Types = List.Transform(TextCols, each {_, type text}) & List.Transform(NumCols, each {_, type number})
            & {{"기준일", type date}, {"조회시각", type datetime}},

    // 실행 제어 tblRunCtl(키, 값): mode — 버튼 매크로가 sector/full을 쓴 실행에서만 KIS를 호출, started — 버튼 실행 시작 시각(매크로가 씀)
    RunCtl = try Excel.CurrentWorkbook(){[Name = "tblRunCtl"]}[Content] otherwise null,
    RunValue = (key as text) as any =>
        let
            Hit = if RunCtl = null or not Table.HasColumns(RunCtl, {"키", "값"}) then {}
                  else List.Select(Table.ToRecords(RunCtl), each (try Text.Lower(Text.Trim(Text.From([키]))) otherwise null) = key)
        in
            if List.IsEmpty(Hit) then null else Hit{0}[값],
    Mode = let V = RunValue("mode") in if V = null then "" else (try Text.Lower(Text.Trim(Text.From(V))) otherwise ""),
    Live = List.Contains({"sector", "full"}, Mode),

    // build 등(호출 없음): 자기 표 → tblSectorKRXSeed → 빈 표. 값 형식을 계약대로 맞추고, 칸별 변환 실패는 빈칸
    Conform = (t as table) as table =>
        let
            Sel = Table.SelectColumns(t, Columns, MissingField.UseNull),
            AsText = (v) => try (if v = null then null else Text.From(v)) otherwise null,
            AsCode = (v) => try (if v = null then null else
                        let s = Text.Trim(Text.From(v)) in if s = "" then null else Text.PadStart(s, 4, "0")) otherwise null,
            AsNum = (v) => try (if v = null or v is number then v else Number.From(v)) otherwise null,
            AsDate = (v) => try (if v = null then null else Date.From(v)) otherwise null,
            AsTime = (v) => try (if v = null then null else DateTime.From(v)) otherwise null,
            Ops = List.Transform({"시장", "업종명", "구분", "상태"}, each {_, AsText, type text})
                  & {{"코드", AsCode, type text}, {"기준일", AsDate, type date}, {"조회시각", AsTime, type datetime}}
                  & List.Transform(NumCols, each {_, AsNum, type number})
        in
            Table.TransformColumns(Sel, Ops),
    HasRows = (t as any) as logical =>
        t is table and Table.HasColumns(t, Columns)
        and not Table.IsEmpty(Table.SelectRows(t, each
            let c = try Text.Trim(Text.From([코드])) otherwise null in c <> null and c <> "")),
    Own = try Excel.CurrentWorkbook(){[Name = "tblSectorKRX"]}[Content] otherwise null,
    Seed = try Excel.CurrentWorkbook(){[Name = "tblSectorKRXSeed"]}[Content] otherwise null,
    Kept = if HasRows(Own) then Conform(Own)
           else if HasRows(Seed) then Conform(Seed)
           else Table.TransformColumnTypes(#table(Columns, {}), Types),

    // 같은 버튼 실행 안에서 다시 평가될 때(새로 고침 직후 Excel이 쿼리를 백그라운드에서 한 번 더 평가함 — 아래 Main 주석)는
    // 조회하지 않고 방금 적재한 자기 표를 돌려준다: 자기 표의 가장 늦은 조회시각이 이번 실행의 started 이후면 이미 조회한 것.
    // started가 비어 있으면(하네스 시험 등) 이 확인은 하지 않는다. 셀 값은 일시·날짜·일련번호·텍스트로 올 수 있다(변환 실패는 빈칸).
    Stamp = (v as any) as nullable datetime =>
        if v = null then null
        else if v is datetime then v
        else if v is date then DateTime.From(v)
        else if v is number then (try DateTime.From(v) otherwise null)
        else if v is text then (try DateTime.FromText(Text.Trim(v), "en-US") otherwise null)
        else null,
    Started = try Stamp(RunValue("started")) otherwise null,
    LastStamp = if not HasRows(Own) then null
                else List.Max(List.RemoveNulls(List.Transform(Table.Column(Own, "조회시각"), (v) => try Stamp(v) otherwise null))),
    Surveyed = Started <> null and LastStamp <> null and LastStamp >= Started,

    // sector·full: 업종 목록(마스터) → 업종마다 시장별 투자자매매동향(일별) 1회. 함수로 두어 sector·full이고 이번 실행에서 아직
    // 조회하지 않았을 때만 부른다(마스터 파일 내려받기도 이 안)
    Survey = () as table =>
        let
            Now = DateTime.LocalNow(),
            Today = Date.From(Now),
            Clock = Time.From(Now),
            TodayText = fnYmdText(Today),
            SkipTodayIndex = Clock < #time(9, 0, 0),      // 장 시작 전: 오늘 행은 아직 세션이 아님
            SkipTodayFlow = Clock < #time(15, 30, 0),     // 장 마감 전: 오늘은 완료 세션이 아님
            Path = "/uapi/domestic-stock/v1/quotations/inquire-investor-daily-by-market",

            KospiCodes = List.Transform({5 .. 30}, each Text.PadStart(Text.From(_), 4, "0")),
            KosdaqCodes = List.Transform({1006 .. 1033}, Text.From),
            SizeCodes = {"0002", "0003", "0004", "1002", "1003", "1004"},
            Parents = {"0021", "0027", "1009"},
            // idxcode.mst: 1자리 시장구분(0 KOSPI, 1 KOSDAQ) + 4자리 업종코드 + 업종명(고정폭, cp949)
            MasterZip = Web.Contents("https://new.real.download.dws.co.kr",
                            [RelativePath = "common/master/idxcode.mst.zip", Timeout = #duration(0, 0, 2, 0)]),
            MasterLines = List.Select(Lines.FromText(Text.FromBinary(fnUnzip(MasterZip), 949)), each Text.Length(_) > 5),
            MasterRecs = List.Transform(MasterLines, each [mk = Text.Start(_, 1), code = Text.Middle(_, 1, 4), name = Text.Trim(Text.Middle(_, 5))]),
            Pick = (prefix as text, codes as list, market as text, iscd1 as text) as list =>
                let
                    Hits = List.Select(MasterRecs, each [mk] = prefix and List.Contains(codes, [code]) and not List.Contains(SizeCodes, [code])),
                    Uniq = List.Sort(List.Distinct(Hits, each [code]), each [code])
                in
                    List.Transform(Uniq, each [시장 = market, 코드 = [code], 업종명 = [name],
                        구분 = if List.Contains(Parents, [code]) then "대분류" else "업종", iscd1 = iscd1]),
            Kospi = Pick("0", KospiCodes, "KOSPI", "KSP"),
            Kosdaq = Pick("1", KosdaqCodes, "KOSDAQ", "KSQ"),
            Industries =
                if List.IsEmpty(Kospi) or List.IsEmpty(Kosdaq) then
                    error Error.Record("업종 목록", "업종코드 마스터(idxcode.mst)에서 " & (if List.IsEmpty(Kospi) then "KOSPI" else "KOSDAQ")
                        & " 업종을 찾지 못했습니다")
                else List.Buffer(Kospi & Kosdaq),

            Calc = (ind as record) as record =>
                let
                    J = fnKisGet(Path, "FHPTJ04040000", [FID_COND_MRKT_DIV_CODE = "U", FID_INPUT_ISCD = ind[코드], FID_INPUT_DATE_1 = TodayText,
                            FID_INPUT_ISCD_1 = ind[iscd1], FID_INPUT_DATE_2 = TodayText, FID_INPUT_ISCD_2 = ind[코드]]),
                    O = Record.FieldOrDefault(J, "output", null),
                    Raw = if O is list then O else error "응답에 일별 목록(output)이 없습니다",
                    Parsed = List.Select(List.Transform(Raw, (r) => [
                            d = fnYmd(Record.FieldOrDefault(r, "stck_bsop_date", null)),
                            px = fnNum(Record.FieldOrDefault(r, "bstp_nmix_prpr", null)),
                            f = fnNum(Record.FieldOrDefault(r, "frgn_ntby_tr_pbmn", null)),
                            o = fnNum(Record.FieldOrDefault(r, "orgn_ntby_tr_pbmn", null)),
                            p = fnNum(Record.FieldOrDefault(r, "prsn_ntby_tr_pbmn", null))]), each [d] <> null),
                    // 최근 날짜부터, 같은 날짜는 첫 행만
                    Desc = List.Buffer(List.Distinct(List.Sort(Parsed, {each [d], Order.Descending}), each [d])),
                    Px = List.Buffer(List.Select(Desc, each [px] <> null and [px] > 0 and not (SkipTodayIndex and [d] = Today))),
                    Fl = List.Buffer(List.Select(Desc, each not (SkipTodayFlow and [d] = Today))),
                    Cur = if List.IsEmpty(Desc) then error "일별 데이터가 없습니다"
                          else if List.IsEmpty(Px) then error "지수 값이 없습니다"
                          else Px{0},
                    Ret = (k as number) as nullable number => let b = Px{k}? in if b = null then null else Cur[px] / b[px] - 1,
                    PrevYear = List.Select(Px, each Date.Year([d]) < Date.Year(Cur[d])),
                    Ytd = if List.IsEmpty(PrevYear) then null else Cur[px] / PrevYear{0}[px] - 1,
                    Flow = (k as text, n as number) as nullable number =>
                        let v = List.Transform(List.FirstN(Fl, n), each Record.Field(_, k))
                        in if List.Count(v) < n or List.Contains(v, null) then null else List.Sum(v) / 100,
                    S = [기준일 = Cur[d], 지수 = Cur[px],
                         #"1D" = Ret(1), #"1W" = Ret(5), #"1M" = Ret(21), #"3M" = Ret(63), #"6M" = Ret(126), YTD = Ytd, #"1Y" = Ret(250),
                         외국인1D억 = Flow("f", 1), 외국인1W억 = Flow("f", 5), 외국인1M억 = Flow("f", 21),
                         기관1D억 = Flow("o", 1), 기관1W억 = Flow("o", 5), 기관1M억 = Flow("o", 21),
                         개인1D억 = Flow("p", 1), 개인1W억 = Flow("p", 5), 개인1M억 = Flow("p", 21)],
                    // 레코드 필드는 지연 평가되므로 여기서 모든 값을 평가해 오류가 이 업종의 try 안에서 잡히게 함
                    // (List.Buffer는 항목 오류를 일으키지 않음 — 실측)
                    Values = Record.FieldValues(S),
                    Forced = List.Count(List.Select(Values, each _ = null or _ is number or _ is date)) = List.Count(Values)
                in
                    if Forced then S else error "계산 값 형식 오류",
            One = (ind as record) as record =>
                let
                    Head = [시장 = ind[시장], 코드 = ind[코드], 업종명 = ind[업종명], 구분 = ind[구분]],
                    R = try Calc(ind),
                    Why = if not R[HasError] then null
                          else let m = R[Error][Message], w = R[Error][Reason] in
                              if m <> null and m <> "" then Text.From(m) else if w <> null and w <> "" then Text.From(w) else "알 수 없는 오류"
                in
                    if R[HasError] then Head & [상태 = "오류: " & Text.Start(Why, 200), 조회시각 = Now]
                    else Head & R[Value] & [상태 = "OK", 조회시각 = Now],
            Rows = List.Buffer(List.Transform(Industries, One)),
            Failed = List.Select(Rows, each Text.StartsWith([상태], "오류")),
            // 전부 실패하면 오류로 올려 직전 데이터를 유지(fnFallback), 일부만 실패하면 그 행만 오류 표시
            Checked = if List.Count(Failed) = List.Count(Rows)
                      then error Error.Record("업종 수급", "모든 업종 조회 실패: " & Text.Middle(Failed{0}[상태], 4))
                      else Rows,
            T = Table.FromRecords(Checked, Columns, MissingField.UseNull)
        in
            Table.TransformColumnTypes(T, Types),

    // Excel은 새로 고침이 끝난 직후 쿼리를 백그라운드에서 한 번 더 평가하고, 최상위 단계도 따로 평가한다(2026-10-01 하네스 +
    // 모의 서버 실측, 동기·백그라운드 새로 고침 모두). 이 평가는 Power Query 캐시에 없는 요청을 서버로 그대로 보낸다.
    //  ① 조회를 값으로 된 단계로 두었을 때는 표를 처음 적재하는 새로 고침이 build 모드여도 그 단계를 평가해 마스터 파일 + 업종 수만큼
    //     호출했다 → 조회는 함수(Survey)로 두고 sector·full일 때만 부른다(함수 값인 단계는 평가해도 호출하지 않음).
    //  ② sector·full 새로 고침 뒤에는 같은 요청이 한 번 더 나갔다(호출 2배) → 위 Surveyed 확인으로 막는다.
    //     (모든 업종이 실패해 직전 표로 돌아간 경우엔 조회시각이 그대로라 막지 못하고 한 번 더 시도한다.)
    //  ③ 표를 처음 적재한 뒤·통합문서를 연 뒤 처음 하는 새로 고침에서는 추가 평가가 자기 표가 갱신되기 전에 함께 돌아 ②의 확인을
    //     통과하므로 그 한 번은 호출이 최대 2배가 될 수 있다(실측). 버튼 매크로가 그 표를 build 모드로 먼저 한 번 새로 고침해 두면 생기지 않는다.
    Main = if Live and not Surveyed then Survey() else Kept,
    // 실패(토큰·네트워크·호출제한·마스터 파일 등) 시 오류 대화상자 대신 직전 결과를 유지.
    // fnFallback 결과는 형식이 없는 표라 날짜 열이 일련번호(일반 서식)로 적재되므로 Conform으로 열 형식을 다시 입힘(값은 그대로)
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Conform(fnFallback("tblSectorKRX", Columns, Tried[Error][Message])) else Tried[Value]
in
    Final
