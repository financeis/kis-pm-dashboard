// fnNum : KIS 응답의 숫자 문자열("-3000", "1.10", "", null) → number / null
(v as any) as nullable number =>
let
    t = if v = null then null else Text.Trim(Text.From(v)),
    r = if t = null or t = "" then null else try Number.FromText(t, "en-US") otherwise null
in
    r
