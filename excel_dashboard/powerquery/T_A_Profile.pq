// T_A_Profile : 종목분석 3구역 (b) — 최근 N세션 일봉으로 계산한 매물대 → tblA_Profile (2026-10-01)
//  - N = 설정 profile_days(기본 60, 5~500). 일봉(수정주가) = fnDailyChart(국내주식기간별시세 FHKST03010100, 100건/회 —
//    기본 60세션은 조회 기간 N×1.5+20일이라 1회 호출)의 최근 N개 봉.
//    장중이면 오늘의 진행 중 봉도 N개에 들어간다 → 현재가(= 마지막 봉 종가)가 항상 [최저가, 최고가] 안에 있어 현재가 구간이 정해진다.
//  - 구간: 최근 N봉의 최저 저가 L ~ 최고 고가 H를 같은 폭 20개로 나눔(구간 1 = 가장 낮은 가격). 구간 i = [하단, 상단), 마지막 구간만 상단 포함.
//  - 배분(그날 고가~저가에 걸친 구간에 균등 배분): 각 세션 거래량을 그날 [저가, 고가] 위에 가격당 균등하게 펼쳐
//    구간과 겹친 길이에 비례해 나눔 — 완전히 덮인 구간은 같은 몫, 끝에 살짝 걸친 구간은 걸친 만큼만(걸치기만 해도 같은 몫을 주면
//    양 끝 구간이 부풀려지므로). 고가 = 저가인 세션(상·하한가 고정 등)은 그 가격이 속한 구간에 전부. 거래량 0인 세션은 영향 없음.
//    세션 범위가 모두 [L, H] 안이므로 배분 합 = 전체 거래량 → 비중 합 = 100%.
//  - 열: 구간, 구간하단, 구간상단, 가격대(차트 축 이름 "52,300~53,100"), 거래량(배분된 주식 수, 소수 가능), 비중(소수),
//    현재가구간(현재가가 속한 구간 "Y", 그 밖 빈칸), 현재가, 시작일·종료일(사용한 봉의 첫·마지막 일자), 세션수.
//  - 모든 봉이 같은 가격이면(H = L) 20구간 폭이 0 — 거래량은 구간 1에. 봉이 없으면 "데이터 없음" 1행.
let
    Types = {{"종목코드", type text}, {"구간", Int64.Type}, {"구간하단", type number}, {"구간상단", type number}, {"가격대", type text},
             {"거래량", type number}, {"비중", type number}, {"현재가구간", type text}, {"현재가", type number},
             {"시작일", type date}, {"종료일", type date}, {"세션수", Int64.Type}, {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Bins = 20,
            Today = Date.From(now),
            Days = fnASetting("profile_days", 60, 5, 500),
            Raw = fnDailyChart(code, Date.AddDays(Today, -Number.RoundUp(Days * 1.5 + 20)), Today),
            Bars = Table.Buffer(Table.LastN(Table.SelectRows(Raw, each [고가] <> null and [저가] <> null and [고가] > 0 and [저가] > 0), Days)),
            Recs = List.Buffer(Table.ToRecords(Bars)),
            L = List.Min(Bars[저가]),
            H = List.Max(Bars[고가]),
            W = (H - L) / Bins,
            Px = List.Last(Bars[종가]),
            BinOf = (x as number) as number =>
                if W = 0 then 0 else List.Max({0, List.Min({Bins - 1, Number.RoundDown((x - L) / W)})}),
            Low = (i as number) as number => L + i * W,
            High = (i as number) as number => if i = Bins - 1 then H else L + (i + 1) * W,
            Alloc = (i as number) as number =>
                List.Sum({0} & List.Transform(Recs, (b) =>
                    let v = b[거래량], hb = b[고가], lb = b[저가] in
                        if v = null or v <= 0 then 0
                        else if hb <= lb then (if BinOf(hb) = i then v else 0)
                        else v * List.Max({0, List.Min({High(i), hb}) - List.Max({Low(i), lb})}) / (hb - lb))),
            Vols = List.Buffer(List.Transform({0 .. Bins - 1}, Alloc)),
            Total = List.Sum(Vols),
            PxBin = if Px = null then null else BinOf(Px),
            Fmt = (x as number) as text => Number.ToText(Number.Round(x, 0), "N0", "en-US"),
            Start = List.Min(Bars[일자]),
            Stop = List.Max(Bars[일자]),
            Count = Table.RowCount(Bars),
            Rows = List.Transform({0 .. Bins - 1}, (i) =>
                {code, i + 1, Low(i), High(i), Fmt(Low(i)) & "~" & Fmt(High(i)), Vols{i},
                 if Total > 0 then Vols{i} / Total else null, if PxBin = i then "Y" else null, Px, Start, Stop, Count, "OK", now})
        in
            if Table.IsEmpty(Bars) or Total = null or Total <= 0 then fnAOneRow(Types, code, "데이터 없음", now)
            else #table(List.Transform(Types, each _{0}), Rows)
in
    fnARun("tblA_Profile", Types, Main)
