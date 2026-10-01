// fnKisGetPage : KIS REST GET 한 쪽 — 연속 조회(tr_cont 요청 헤더)가 필요한 API용 (2026-10-01)
//  - fnKisGet과 같은 헤더·재시도 규칙(설정·토큰 문제는 재시도 없이 즉시 오류, EGW00201만 짧은 백오프로 최대 5회,
//    네트워크 오류는 1회)에 요청 헤더 tr_cont만 더한다. fnKisGet은 이 헤더를 보내지 못한다.
//  - 반환: 응답 본문 레코드(rt_cd = 0 확인됨). rt_cd <> "0"이면 KIS 메시지를 담은 오류.
//  - trCont: "" = 첫 쪽(헤더 생략), "N" = 다음 쪽 요청. 어디서부터 이어 받을지는 query의 CTS가 정한다.
//    예탁원정보(ksdinfo) API의 CTS = 앞 쪽 마지막 행의 정렬 기준 일자(YYYYMMDD) + 종목코드(sht_cd)를 이어 붙인 값
//    (2026-10-01 실측: 'CTS 없이 tr_cont=N만' 또는 '형식이 틀린 CTS'는 같은/첫 쪽을 되풀이해 연속 조회가 진행되지 않고,
//     올바른 CTS + tr_cont=N이어야 다음 행부터 줌. CTS만 보내고 tr_cont를 빼면 CTS를 무시함)
//  - 응답 헤더 tr_cont(F/M = 다음 쪽 있음)는 파워 쿼리가 돌려주지 않아(Value.Metadata의 Headers에는 Content-Type·
//    Content-Length·Date만 있음, 2026-10-01 실측) 이 함수는 읽지 않는다. 호출하는 쪽이 '쪽이 가득 찼는지'로 판단한다.
//  - 사용 예) fnKisGetPage("/uapi/domestic-stock/v1/ksdinfo/dividend", "HHKDB669102C0",
//                [CTS = "20260930" & "0174J0", GB1 = "0", F_DT = "20260924", T_DT = "20261130", SHT_CD = "", HIGH_GB = ""], "N")
(path as text, trId as text, query as record, trCont as text) as record =>
let
    Cfg = fnCfg(),
    Base = Text.Trim(Cfg[prod]),
    Token = fnToken("prod"),
    Headers0 = [
        #"content-type" = "application/json; charset=utf-8",
        authorization = "Bearer " & Token,
        appkey = Text.Trim(Cfg[my_app]),
        appsecret = Text.Trim(Cfg[my_sec]),
        tr_id = trId,
        custtype = "P"
    ],
    Headers = if trCont = "" then Headers0 else Record.AddField(Headers0, "tr_cont", trCont),
    // 여기서 설정·토큰을 먼저 평가 → 문제가 있으면 재시도 없이 바로 오류
    Ready = Text.Length(Base) > 0 and Text.Length(Headers0[authorization]) > 7 and Text.Length(Headers0[appkey]) > 0,
    RelPath = if Text.StartsWith(path, "/") then Text.Middle(path, 1) else path,
    Fetch = (attempt as number) as record =>
        let
            Raw = Web.Contents(Base, [
                RelativePath = RelPath,
                Query = query,
                Headers = Headers,
                ManualStatusHandling = {400, 401, 403, 404, 429, 500, 502, 503, 504},
                IsRetry = attempt > 0,
                Timeout = #duration(0, 0, 0, 20)]),
            Parsed = try Json.Document(Raw)
        in
            if Parsed[HasError] then [rt_cd = "-1", msg_cd = "HTTP", msg1 = "네트워크/서버 응답 오류: " & Parsed[Error][Message]]
            else if Parsed[Value] is record then Parsed[Value]
            else [rt_cd = "-1", msg_cd = "FORMAT", msg1 = "예상하지 못한 응답 형식"],
    Code = (b as record) as text => Text.From(Record.FieldOrDefault(b, "msg_cd", "")),
    Attempt = (n as number) as record =>
        let
            r = if n = 0 then Fetch(0)
                else Function.InvokeAfter(() => Fetch(n), #duration(0, 0, 0, 0.3 * Number.Power(2, n - 1)))
        in
            if (Code(r) = "EGW00201" and n < 5) or (Code(r) = "HTTP" and n < 1) then @Attempt(n + 1) else r,
    Body = if Ready then Attempt(0) else error Error.Record("KIS 설정", "kis_devlp.yaml의 prod 도메인/앱키 또는 접근토큰이 비어 있습니다."),
    Result =
        if Text.From(Record.FieldOrDefault(Body, "rt_cd", "")) = "0" then Body
        else error Error.Record(
            "KIS API 오류",
            Code(Body) & " " & Text.From(Record.FieldOrDefault(Body, "msg1", "")),
            [path = path, tr_id = trId])
in
    Result
