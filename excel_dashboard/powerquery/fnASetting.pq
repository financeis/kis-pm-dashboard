// fnASetting : 종목분석 쿼리의 숫자 설정값 (tblSettings 키 → 정수, 없거나 숫자가 아니면 기본값, 범위 밖이면 경계값) (2026-10-01 T15)
//  - 사용 키: profile_days(60, N일 매물대 세션 수), target_window_months(6, 목표주가 기간), event_days_ahead(60)·event_days_back(7, 이벤트 기간)
//  - 설정표가 없어도(빌더가 새 키를 넣기 전) 기본값으로 동작한다.
(key as text, default as number, minValue as number, maxValue as number) as number =>
let
    S = try fnSettings() otherwise [],
    V = try Number.From(Record.FieldOrDefault(S, key, null)) otherwise null
in
    if V = null then default else List.Max({minValue, List.Min({maxValue, Number.RoundDown(V)})})
