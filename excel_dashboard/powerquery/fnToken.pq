// fnToken : 통합문서에 캐시된 접근토큰(tblToken)을 읽기만 합니다. (발급은 T_Token 쿼리만 수행)
//  - 여러 쿼리가 동시에 토큰을 발급하면 KIS 제한(1분당 1회)에 걸리고 알림톡이 여러 번 발송되므로
//    토큰 발급은 한 곳(T_Token)에서만 하고, 나머지 쿼리는 이 함수를 통해 캐시된 토큰만 사용합니다.
(optional env as nullable text) as text =>
let
    e = if env = null then "prod" else env,
    T = try Excel.CurrentWorkbook(){[Name = "tblToken"]}[Content] otherwise null,
    Rows = if T = null or not Table.HasColumns(T, {"env", "token", "expires"}) then null
           else Table.SelectRows(T, each [env] = e and [token] <> null and [token] <> ""),
    Row = if Rows = null or Table.RowCount(Rows) = 0 then null else Rows{0},
    Expires = if Row = null then null else try DateTime.From(Row[expires]) otherwise null,
    Valid = Expires <> null and Expires > DateTime.LocalNow() + #duration(0, 0, 5, 0),
    Result =
        if Valid then Row[token]
        else error Error.Record(
            "KIS 토큰",
            "유효한 접근토큰이 없습니다. 방금 T_Token 쿼리가 새 토큰을 발급했다면 [데이터] > [모두 새로 고침]을 한 번 더 실행하세요.",
            [env = e])
in
    Result
