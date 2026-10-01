// fnKisGet : KIS REST GET 공통 호출 (kis_auth.py의 _url_fetch와 같은 헤더 구성)
//  - 설정 파일·토큰 문제는 재시도하지 않고 즉시 오류 (호출마다 기다리지 않도록 먼저 평가)
//  - 초당 호출 제한(EGW00201)만 짧은 지수 백오프로 최대 5회 재시도 (IsRetry=true로 캐시 무시)
//  - 네트워크/서버 응답 오류는 1회만 재시도
//  - rt_cd <> "0" 이면 KIS 메시지를 담은 오류를 발생시킵니다.
// 사용 예) fnKisGet("/uapi/domestic-stock/v1/quotations/inquire-price", "FHKST01010100", [FID_COND_MRKT_DIV_CODE="J", FID_INPUT_ISCD="005930"])
(path as text, trId as text, query as record) as record =>
let
    Cfg = fnCfg(),
    Base = Text.Trim(Cfg[prod]),
    Token = fnToken("prod"),
    Headers = [
        #"content-type" = "application/json; charset=utf-8",
        authorization = "Bearer " & Token,
        appkey = Text.Trim(Cfg[my_app]),
        appsecret = Text.Trim(Cfg[my_sec]),
        tr_id = trId,
        custtype = "P"
    ],
    // 여기서 설정·토큰을 먼저 평가 → 문제가 있으면 재시도 없이 바로 오류
    Ready = Text.Length(Base) > 0 and Text.Length(Headers[authorization]) > 7 and Text.Length(Headers[appkey]) > 0,
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
