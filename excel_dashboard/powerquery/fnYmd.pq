// fnYmd : "YYYYMMDD" → date (잘못된 값이면 null)
(v as any) as nullable date =>
let
    t = if v = null then "" else Text.Trim(Text.From(v)),
    r = if Text.Length(t) <> 8 then null
        else try #date(Number.From(Text.Start(t, 4)), Number.From(Text.Middle(t, 4, 2)), Number.From(Text.End(t, 2))) otherwise null
in
    r
