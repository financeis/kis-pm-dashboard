// T_News : 보유·관심 종목 뉴스·공시 제목 — 시간순으로 합친 최근 100건 ([모두 새로 고침]에서 갱신)
//  - 종합 시황/공시(제목) [국내주식-141] FHKST01011800 /uapi/domestic-stock/v1/quotations/news-title
//    종목당 1회 호출 = 그 종목코드가 등록된 최근 40건(연속 조회 없음). FID_INPUT_ISCD만 지정하고 나머지 입력은 공백(현재 기준).
//  - 응답 output[] → 열: data_dt(작성일자 YYYYMMDD) + data_tm(작성시간 HHMMSS) → 일시, hts_pbnt_titl_cntt → 제목,
//    dorg(자료원 = 연합뉴스·서울경제·인포스탁 같은 제공처 이름) → 출처. news_ofer_entp_code(제공업체 한 글자 코드)는 쓰지 않음.
//  - 종목명: 응답의 관련 종목 쌍(iscd1~10 ↔ kor_isnm1~10)에서 그 종목코드의 이름 → 없으면 종목DB(tblUniverse) → 없으면 빈칸.
//    관련 종목 칸에는 다른 종목이 오는 기사가 많아(시황 기사 등) 행마다가 아니라 종목 단위로 한 번 정함.
//  - 같은 종목 안에서 일시·제목·출처가 같은 행은 1건으로 줄임. 여러 종목에 걸린 기사는 종목마다 1행.
//  - 일시 내림차순 최근 100건 + 조회 실패 종목마다 1행(일시·제목·출처 빈칸, 상태 "오류: 사유")을 뒤에 덧붙임.
//    실패 = API 오류 / 종목코드 형식(6자리 영숫자, ETN은 Q+6자리)이 아님(호출 안 함) / 응답 0건.
//    없는 종목코드(예: 999999)도 rt_cd 0에 빈 목록으로 오고, 상장 종목은 소형주도 최근 1~3주에 40건이 차므로(실측) 0건은 코드 확인 대상.
//    모든 종목이 실패하면 오류로 올려 직전 표를 유지(fnFallback). 보유·관심 종목이 없으면 열만 있는 빈 표.
let
    Cols = {"일시", "종목코드", "종목명", "제목", "출처", "상태", "조회시각"},
    Types = {{"일시", type datetime}, {"종목코드", type text}, {"종목명", type text}, {"제목", type text}, {"출처", type text},
             {"상태", type text}, {"조회시각", type datetime}},
    Main =
        let
            Now = DateTime.LocalNow(),
            Keep = 100,
            Path = "/uapi/domestic-stock/v1/quotations/news-title",
            Schema = type table [일시 = nullable datetime, 종목코드 = text, 종목명 = nullable text, 제목 = nullable text,
                                 출처 = nullable text, 상태 = text, 조회시각 = datetime, __srno = text],
            Codes = List.Distinct(List.Select(List.Transform(fnCodes()[종목코드],
                        each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_)))), each _ <> "")),
            U = Table.Buffer(Table.SelectRows(fnUniverse(), each List.Contains(Codes, [종목코드]))),
            Txt = (r as record, k as text) as text =>
                let v = Record.FieldOrDefault(r, k, null) in if v = null then "" else try Text.Trim(Text.From(v)) otherwise "",
            Blank = (t as text) as nullable text => if t = "" then null else t,
            UName = (code as text) as nullable text =>
                let m = Table.SelectRows(U, each [종목코드] = code)
                in if Table.RowCount(m) = 0 then null else Blank(try Text.Trim(Text.From(m{0}[종목명])) otherwise ""),
            ValidCode = (code as text) as logical =>
                List.Contains({6, 7}, Text.Length(code)) and Text.Select(code, {"0".."9", "A".."Z"}) = code
                and (Text.Length(code) = 6 or (Text.StartsWith(code, "Q") and Text.Select(code, {"0".."9"}) = Text.Middle(code, 1))),
            // 작성일자(YYYYMMDD) + 작성시간(HHMMSS, 4자리면 HHMM) → datetime. 날짜가 잘못되면 null, 시간만 잘못되면 00:00:00
            ToDt = (d as text, t as text) as nullable datetime =>
                let
                    day = fnYmd(d),
                    s = Text.Select(t, {"0".."9"}),
                    hms = if Text.Length(s) = 6 then s else if Text.Length(s) = 4 then s & "00" else "000000",
                    tm = try #time(Number.From(Text.Start(hms, 2)), Number.From(Text.Middle(hms, 2, 2)), Number.From(Text.End(hms, 2)))
                         otherwise #time(0, 0, 0)
                in
                    if day = null then null else day & tm,
            Fetch = (code as text) as table =>
                let
                    J = if ValidCode(code) then
                            fnKisGet(Path, "FHKST01011800", [FID_NEWS_OFER_ENTP_CODE = "", FID_COND_MRKT_CLS_CODE = "", FID_INPUT_ISCD = code,
                                FID_TITL_CNTT = "", FID_INPUT_DATE_1 = "", FID_INPUT_HOUR_1 = "", FID_RANK_SORT_CLS_CODE = "", FID_INPUT_SRNO = ""])
                        else error "종목코드 형식 확인(6자리 영숫자, ETN은 Q+6자리)",
                    O = Record.FieldOrDefault(J, "output", {}),
                    L = List.Buffer(List.Select(if O is list then O else if O is record then {O} else {}, each _ is record)),
                    // 응답의 관련 종목 쌍에서 이 종목코드의 이름
                    Pairs = List.Combine(List.Transform(L, (r) =>
                                List.Transform({1..10}, (i) => {Text.Upper(Txt(r, "iscd" & Text.From(i))), Txt(r, "kor_isnm" & Text.From(i))}))),
                    Named = List.Select(Pairs, each _{0} = code and _{1} <> ""),
                    Name = if List.Count(Named) > 0 then Named{0}{1} else UName(code),
                    Rows = List.Transform(L, (r) => {ToDt(Txt(r, "data_dt"), Txt(r, "data_tm")), code, Name,
                                Blank(Txt(r, "hts_pbnt_titl_cntt")), Blank(Txt(r, "dorg")), "OK", Now, Txt(r, "cntt_usiq_srno")}),
                    T = Table.Distinct(#table(Schema, Rows), {"종목코드", "일시", "제목", "출처"})
                in
                    // 레코드·목록은 지연 평가되므로 응답 J를 먼저 평가해 행 단위 try가 API 오류를 잡도록 하고, 표는 버퍼로 확정
                    if not (J is record) then error "빈 응답"
                    else if List.Count(L) = 0 then error "뉴스 0건(종목코드 확인)"
                    else Table.Buffer(T),
            // 종목마다 한 번만 호출되도록 결과 목록을 버퍼로 고정 (레코드 필드는 처음 읽을 때 한 번 평가)
            Results = List.Buffer(List.Transform(Codes, (c) => [code = c, r = try Fetch(c)])),
            Failed = List.Select(Results, each [r][HasError]),
            Reason = (e as record) as text =>
                let m = try Text.Trim(Text.From(e[Message])) otherwise "" in if m = "" then "조회 실패" else m,
            // 전부 실패하면 오류로 올려 직전 표를 유지(fnFallback), 일부만 실패하면 그 종목만 오류 행
            Checked = if List.Count(Results) > 0 and List.Count(Failed) = List.Count(Results)
                      then error Reason(Failed{0}[r][Error]) else Results,
            News = Table.Combine(List.Transform(List.Select(Checked, each not [r][HasError]), each [r][Value]) & {#table(Schema, {})}),
            Latest = Table.FirstN(Table.Sort(News, {{"일시", Order.Descending}, {"__srno", Order.Descending}, {"종목코드", Order.Ascending}}), Keep),
            ErrRows = #table(Schema, List.Transform(Failed, each {null, [code], UName([code]), null, null, "오류: " & Reason([r][Error]), Now, ""})),
            Out = Table.RemoveColumns(Table.Combine({Latest, ErrRows}), {"__srno"})
        in
            Out,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지. 직전 표를 다시 읽으면 열 형식이 없어(any)
    // 일시·조회시각이 일련번호로 보이므로 같은 형식을 다시 지정한다
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblNews", Cols, Tried[Error][Message]), Types) else Tried[Value]
in
    Final
