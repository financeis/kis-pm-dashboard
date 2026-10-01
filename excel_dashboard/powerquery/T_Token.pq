// T_Token : KIS REST 접근토큰 캐시 (자기 자신의 출력표 tblToken을 다시 읽는 '자기참조' 쿼리)
//  - 토큰 유효기간 24시간 / 6시간 이내 재발급 시 같은 토큰 반환 / 발급할 때마다 알림톡 발송 (kis_auth.py 주석 참고)
//  - 남은 유효시간이 3시간 이상이면 재사용하고, 그 미만이거나 앱키가 바뀌면 새로 발급합니다.
//  - 발급 실패(1분당 1회 제한 등) 시 기존 토큰이 아직 유효하면 계속 사용합니다.
let
    CfgTry = try fnCfg(),
    Cfg = if CfgTry[HasError] then [my_app = "", my_sec = "", prod = ""] else CfgTry[Value],
    CfgError = if CfgTry[HasError] then "설정 오류: " & CfgTry[Error][Message] else null,
    Now = DateTime.LocalNow(),
    AppKey = Text.Trim(Cfg[my_app]),
    AppSec = Text.Trim(Cfg[my_sec]),
    Base = Text.Trim(Cfg[prod]),
    // 앱키 자체는 저장하지 않고 변경 여부만 확인하기 위한 간단한 체크섬
    KeySig = Text.From(List.Sum({0} & List.Transform(Text.ToList(AppKey), Character.ToNumber)) * 7 + Text.Length(AppKey)),
    Schema = type table [env = text, token = text, expires = datetime, issued = datetime, key_sig = text, status = text, checked = datetime],

    Cached = try Excel.CurrentWorkbook(){[Name = "tblToken"]}[Content] otherwise null,
    CachedRow =
        if Cached = null or not Table.HasColumns(Cached, {"env", "token", "expires", "issued", "key_sig"}) then null
        else
            let r = Table.SelectRows(Cached, each [env] = "prod" and [token] <> null and [token] <> "")
            in if Table.RowCount(r) = 0 then null else r{0},
    CachedExp = if CachedRow = null then null else try DateTime.From(CachedRow[expires]) otherwise null,
    CachedIssued = if CachedRow = null then null else try DateTime.From(CachedRow[issued]) otherwise null,
    CachedValid = CachedExp <> null and CachedExp > Now + #duration(0, 0, 10, 0),
    SameKey = CachedRow <> null and Text.From(CachedRow[key_sig]) = KeySig,
    Reuse = CachedValid and SameKey and CachedExp - Now > #duration(0, 3, 0, 0),

    ParseDT = (s as text) as datetime =>
        #datetime(
            Number.From(Text.Middle(s, 0, 4)), Number.From(Text.Middle(s, 5, 2)), Number.From(Text.Middle(s, 8, 2)),
            Number.From(Text.Middle(s, 11, 2)), Number.From(Text.Middle(s, 14, 2)), Number.From(Text.Middle(s, 17, 2))),

    Issue = () as record =>
        let
            Body = Json.FromValue([grant_type = "client_credentials", appkey = AppKey, appsecret = AppSec]),
            Resp = Web.Contents(Base, [
                RelativePath = "oauth2/tokenP",
                Headers = [#"Content-Type" = "application/json; charset=utf-8"],
                Content = Body,
                ManualStatusHandling = {400, 401, 403, 404, 429, 500, 502, 503},
                Timeout = #duration(0, 0, 0, 30)]),
            J = try Json.Document(Resp) otherwise [error_description = "응답 파싱 실패"],
            Out =
                if Record.HasFields(J, "access_token") then
                    [ok = true, token = J[access_token], expires = ParseDT(J[access_token_token_expired]), msg = "신규 발급"]
                else
                    [ok = false, token = null, expires = null,
                     msg = "발급 실패: " & Text.From(Record.FieldOrDefault(J, "error_code", "")) & " "
                         & Text.From(Record.FieldOrDefault(J, "error_description", Record.FieldOrDefault(J, "msg1", "")))]
        in
            Out,

    KeepCached = (msg as text) as list =>
        if CachedValid then {"prod", CachedRow[token], CachedExp, CachedIssued, Text.From(CachedRow[key_sig]), msg & " → 기존 토큰 사용", Now}
        else {"prod", null, null, null, null, msg, Now},

    Main =
        if CfgError <> null then
            #table(Schema, {KeepCached(CfgError)})
        else if Reuse then
            #table(Schema, {{"prod", CachedRow[token], CachedExp, CachedIssued, KeySig, "재사용", Now}})
        else
            let
                N = try Issue() otherwise [ok = false, token = null, expires = null, msg = "발급 실패: 네트워크/설정 오류"],
                Row =
                    if N[ok] then {"prod", N[token], N[expires], Now, KeySig, N[msg], Now}
                    else if CachedValid and SameKey then KeepCached(N[msg])
                    else {"prod", null, null, null, KeySig, N[msg], Now}
            in
                #table(Schema, {Row}),
    // 예상치 못한 오류도 대화상자 없이 상태 문자열로 남김
    Result = try Table.Buffer(Main) otherwise #table(Schema, {KeepCached("토큰 쿼리 오류")})
in
    Result
