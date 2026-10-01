// fnYmdText : date → "YYYYMMDD"
(d as date) as text =>
    Date.ToText(d, [Format = "yyyyMMdd", Culture = "en-US"])
