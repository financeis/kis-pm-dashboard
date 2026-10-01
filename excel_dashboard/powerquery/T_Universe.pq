// T_Universe : KOSPI·KOSDAQ 종목 유니버스 (KIS 종목정보 마스터파일, 인증 불필요) + 분류·대회편입 (2026-10-01)
//  - 저장소 stocks_info/kis_kospi_code_mst.py, kis_kosdaq_code_mst.py, sector_code.py 의 정제 로직을 M으로 옮긴 것
//  - 주식(ST)·외국주식(FS)·DR만 남기고 ETF/ETN/리츠/펀드 등은 제외
//  - KRX업종 = 기존 섹터 계산 규칙의 KRX 지수업종 이름(중분류 → 없으면 대분류 → 없으면 "기타"). 분류 대체의 입력이자
//    다른 쿼리(fnPageRows 등)가 읽는 열이다.
//  - 분류·대회편입은 fnClassify(종목코드, KRX업종)가 정한다(명단·수정표·NICS·테마 규칙을 여기서 다시 구현하지 않음):
//      대회편입 = 대회 종목 여부 Y/N(명단 tblContest ± 수정표 tblOverride), 대회비고, 섹터 = 설정 sector_basis 기준 값
//      (없으면 세부 → 업종 → 대분류 → KRX업종), 대분류 = NICS 대분류, NICS 업종, NICS 세부, 대테마, 세부테마, 분류출처(NICS/KRX).
//    기존 열 이름(섹터·대분류·대회편입)은 그대로 두고 값의 기준만 바뀐다. fnClassify가 버리는 형식의 종목코드(6자리 [0-9A-Z]가
//    아님)는 대회편입 N, 섹터 = KRX업종, 나머지 분류 칸은 빈칸.
//  - [모두 새로 고침]으로 갱신하는 일반 쿼리. tblClass(T_Class)·tblContest·tblOverride·tblThemeBase·tblSettings를 읽으므로
//    빌드 때는 T_Class 다음에 새로 고친다(같은 새로 고침 안에서 다른 쿼리가 적재한 표는 직전 결과로 읽힘).
//  - 앞 14개 열은 [종목DB] 검색 결과와 같은 순서(열 너비 공유). 새 분류 열은 대분류 뒤에 모아 둔다.
let
    Cols = {"종목코드", "종목명", "시장", "섹터", "규모", "지수편입", "유의사항", "대회편입", "기준가", "시가총액억", "매출액억",
            "영업이익억", "영업이익률", "ROE", "재무기준", "상장일", "상장주식수", "대분류", "NICS 업종", "NICS 세부", "대테마",
            "세부테마", "분류출처", "대회비고", "KRX업종", "우선주", "SPAC", "거래정지", "관리종목", "시장경고", "종류"},
    Types = {{"종목코드", type text}, {"종목명", type text}, {"시장", type text}, {"섹터", type text}, {"규모", type text},
             {"지수편입", type text}, {"유의사항", type text}, {"대회편입", type text}, {"기준가", type number},
             {"시가총액억", type number}, {"매출액억", type number}, {"영업이익억", type number}, {"영업이익률", type number},
             {"ROE", type number}, {"재무기준", type text}, {"상장일", type date}, {"상장주식수", type number}, {"대분류", type text},
             {"NICS 업종", type text}, {"NICS 세부", type text}, {"대테마", type text}, {"세부테마", type text},
             {"분류출처", type text}, {"대회비고", type text}, {"KRX업종", type text}, {"우선주", type text}, {"SPAC", type text},
             {"거래정지", type text}, {"관리종목", type text}, {"시장경고", type text}, {"종류", type text}},
    Main =
        let
            BaseUrl = "https://new.real.download.dws.co.kr",
            GetText = (file as text) as text =>
                Text.FromBinary(fnUnzip(Web.Contents(BaseUrl, [RelativePath = "common/master/" & file, Timeout = #duration(0, 0, 2, 0)])), 949),

            // 업종코드(idxcode.mst): 1자리 시장구분 + 4자리 업종코드 + 업종명
            IdxLines = List.Select(Lines.FromText(GetText("idxcode.mst.zip")), each Text.Length(_) > 5),
            IdxPairs = List.Distinct(List.Transform(IdxLines, each {Text.Start(_, 5), Text.Trim(Text.Middle(_, 5))}), each _{0}),
            IdxMap = Record.FromList(List.Transform(IdxPairs, each _{1}), List.Transform(IdxPairs, each _{0})),
            IdxName = (key as text) as nullable text => Record.FieldOrDefault(IdxMap, key, null),

            Sub = (s as text, spec as list) as text => Text.Trim(Text.Middle(s, spec{0}, spec{1})),
            Num = (s as text) as nullable number => if s = "" then null else try Number.FromText(s, "en-US") otherwise null,

            // 고정폭 필드 위치 (offset, length) — 마스터 파일 part2 기준
            KospiSpec = [tail = 227, prefix = "0", grp = {0, 2}, size = {2, 1}, big = {3, 4}, mid = {7, 4}, idxflag = {18, 1},
                         spac = {29, 1}, base = {41, 9}, halt = {60, 1}, liq = {61, 1}, admin = {62, 1}, warn = {63, 2},
                         warnpre = {65, 1}, unfair = {66, 1}, overheat = {32, 1}, listdt = {105, 8}, shares = {113, 15},
                         pref = {158, 1}, shortheat = {159, 1}, surge = {160, 1}, sales = {163, 9}, op = {172, 9},
                         roe = {195, 9}, fym = {204, 8}, mcap = {212, 9}],
            KosdaqSpec = [tail = 221, prefix = "1", grp = {0, 2}, size = {2, 1}, big = {3, 4}, mid = {7, 4}, idxflag = {35, 1},
                          spac = {24, 1}, base = {36, 9}, halt = {55, 1}, liq = {56, 1}, admin = {57, 1}, warn = {58, 2},
                          warnpre = {60, 1}, unfair = {61, 1}, overheat = {27, 1}, listdt = {100, 8}, shares = {108, 15},
                          pref = {153, 1}, shortheat = {154, 1}, surge = {155, 1}, sales = {157, 9}, op = {166, 9},
                          roe = {189, 9}, fym = {198, 8}, mcap = {206, 9}],

            ParseFile = (file as text, market as text, S as record) as list =>
                let
                    Tail = S[tail],
                    Lines0 = List.Select(Lines.FromText(GetText(file)), each Text.Length(_) > Tail + 21),
                    Rows = List.Transform(Lines0, (line) =>
                        let
                            p1 = Text.Start(line, Text.Length(line) - Tail),
                            p2 = Text.End(line, Tail),
                            F = (k as text) as text => Sub(p2, Record.Field(S, k)),
                            big = F("big"),
                            mid = F("mid"),
                            // 기존 섹터 규칙(지수업종 중분류 → 대분류) = KRX업종
                            sector =
                                if mid <> "" and mid <> "0000" and IdxName(S[prefix] & mid) <> null then IdxName(S[prefix] & mid)
                                else if big <> "" and big <> "0000" then IdxName(S[prefix] & big)
                                else null,
                            idxflag = F("idxflag"),
                            warn = F("warn"),
                            pref = F("pref"),
                            sales = Num(F("sales")),
                            op = Num(F("op")),
                            roe = Num(F("roe")),
                            fym = F("fym")
                        in
                            [
                                종목코드 = Text.Trim(Text.Start(p1, 9)),
                                종목명 = Text.Trim(Text.Middle(p1, 21)),
                                시장 = market,
                                종류 = F("grp"),
                                KRX업종 = if sector = null then "기타" else sector,
                                규모 = let z = F("size") in if z = "1" then "대형" else if z = "2" then "중형" else if z = "3" then "소형" else "미분류",
                                지수편입 = if market = "KOSPI" then (if idxflag <> "" and idxflag <> "0" then "KOSPI200" else null)
                                          else (if idxflag = "Y" then "KOSDAQ150" else null),
                                기준가 = Num(F("base")),
                                시가총액억 = Num(F("mcap")),
                                상장주식수 = let s = Num(F("shares")) in if s = null then null else s * 1000,
                                매출액억 = sales,
                                영업이익억 = op,
                                영업이익률 = if sales = null or sales = 0 or op = null then null else op / sales,
                                ROE = if roe = null then null else roe / 100,
                                재무기준 = if Text.Length(fym) = 8 then Text.Start(fym, 4) & "-" & Text.Middle(fym, 4, 2) else null,
                                상장일 = fnYmd(F("listdt")),
                                우선주 = if pref = "" or pref = "0" then "N" else "Y",
                                SPAC = if F("spac") = "Y" then "Y" else "N",
                                거래정지 = F("halt"),
                                정리매매 = F("liq"),
                                관리종목 = F("admin"),
                                시장경고 = if warn = "01" then "투자주의" else if warn = "02" then "투자경고" else if warn = "03" then "투자위험" else "",
                                불성실공시 = F("unfair"),
                                단기과열 = if F("overheat") = "" or F("overheat") = "0" or F("overheat") = "N" then "N" else "Y"
                            ])
                in
                    Rows,

            All = ParseFile("kospi_code.mst.zip", "KOSPI", KospiSpec) & ParseFile("kosdaq_code.mst.zip", "KOSDAQ", KosdaqSpec),
            Tbl = Table.FromRecords(All),
            Stocks = Table.SelectRows(Tbl, each List.Contains({"ST", "FS", "DR"}, [종류])),
            WithFlag = Table.AddColumn(Stocks, "유의사항", each
                Text.Combine(List.RemoveNulls({
                    if [거래정지] = "Y" then "거래정지" else null,
                    if [정리매매] = "Y" then "정리매매" else null,
                    if [관리종목] = "Y" then "관리종목" else null,
                    if [시장경고] <> "" then [시장경고] else null,
                    if [불성실공시] = "Y" then "불성실공시" else null,
                    if [단기과열] = "Y" then "단기과열" else null,
                    if [SPAC] = "Y" then "SPAC" else null,
                    if [우선주] = "Y" then "우선주" else null}), ","), type text),
            // 마스터 파일 순서 — 조인은 행 순서를 보장하지 않으므로 마지막 정렬의 동점 처리에 씀
            Indexed = Table.AddIndexColumn(WithFlag, "__ord", 0, 1, Int64.Type),

            // 분류·대회편입 (종목코드당 1행, 형식 밖 코드는 빠짐 → 종목코드로 다시 붙임)
            Cls = Table.Buffer(fnClassify(Table.SelectColumns(Indexed, {"종목코드", "KRX업종"}))),
            Joined = Table.NestedJoin(Indexed, {"종목코드"}, Cls, {"종목코드"}, "c", JoinKind.LeftOuter),
            Expanded = Table.ExpandTableColumn(Joined, "c",
                {"대회편입", "대회비고", "NICS 대분류", "NICS 업종", "NICS 세부", "대테마", "세부테마", "분류출처", "섹터"},
                {"__mem", "대회비고", "대분류", "NICS 업종", "NICS 세부", "대테마", "세부테마", "분류출처", "__sec"}),
            Filled = Table.AddColumn(Table.AddColumn(Expanded,
                        "대회편입", each if [__mem] = "Y" then "Y" else "N", type text),
                        "섹터", each if [__sec] <> null then [__sec] else [KRX업종], type text),
            Final = Table.SelectColumns(Filled, Cols & {"__ord"}),
            Typed = Table.TransformColumnTypes(Final, Types),
            Sorted = Table.RemoveColumns(Table.Sort(Typed, {{"시가총액억", Order.Descending}, {"__ord", Order.Ascending}}), {"__ord"})
        in
            Sorted,
    // 실패(네트워크 등) 시 오류 대화상자 대신 직전 결과를 유지. 직전 표를 다시 읽으면 열 형식이 없어(any) Excel이 날짜 칸
    // 서식을 '일반'으로 바꾸므로 같은 형식을 다시 지정(값은 그대로)
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblUniverse", Cols, Tried[Error][Message]), Types) else Tried[Value]
in
    Final
