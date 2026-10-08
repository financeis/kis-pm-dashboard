// fnCfg : KIS 설정 파일(kis_devlp.yaml)을 읽어 레코드로 반환
//  - 앱키/시크릿은 통합문서에 저장하지 않고, KIS 공식 샘플코드(kis_auth.py)와 같은 파일을 그대로 읽습니다.
//  - "key: value", "key: ""value""", "key: 'value'  # 주석" 형식을 모두 지원
() as record =>
let
    Path = Text.Trim(Text.From(fnSettings()[cfg_path])),
    Lines = Lines.FromBinary(File.Contents(Path), null, null, 65001),
    Parse = (raw as text) as nullable list =>
        let
            line = Text.Trim(raw),
            isKV = Text.Length(line) > 0 and not Text.StartsWith(line, "#") and Text.Contains(line, ":"),
            key = if isKV then Text.Trim(Text.BeforeDelimiter(line, ":")) else null,
            rest = if isKV then Text.Trim(Text.AfterDelimiter(line, ":")) else null,
            val =
                if rest = null then null
                else if Text.StartsWith(rest, """") then Text.BetweenDelimiters(rest, """", """")
                else if Text.StartsWith(rest, "'") then Text.BetweenDelimiters(rest, "'", "'")
                else Text.Trim(Text.BeforeDelimiter(rest & " #", " #"))
        in
            if key = null or Text.Contains(key, " ") then null else {key, val},
    Pairs = List.RemoveNulls(List.Transform(Lines, Parse)),
    // 같은 키가 여러 번 나오면 첫 번째 값을 사용
    Dedup = List.Accumulate(Pairs, [],
        (state, p) => if Record.HasFields(state, p{0}) then state else Record.AddField(state, p{0}, p{1})),
    Required = {"my_app", "my_sec", "prod"},
    Missing = List.Select(Required, each not Record.HasFields(Dedup, _) or Record.Field(Dedup, _) = null or Record.Field(Dedup, _) = ""),
    Result =
        if List.Count(Missing) > 0 then
            error Error.Record("KIS 설정", "kis_devlp.yaml에 필수 항목이 없습니다: " & Text.Combine(Missing, ", "), Path)
        else Dedup
in
    Result
