// fnTargetConsensus : 기준일 asOf의 목표주가 컨센서스 (2026-10-01, 규칙은 docs/business-rules.md의 '목표주가 컨센서스')
//  - 규칙: 기간 (asOf − months개월, asOf] 안에서 — asOf 포함, 시작일 제외 — 목표가가 있는 행만 보고 증권사별 마지막 목표가
//    (가장 늦은 일자, 같은 날이면 순번이 앞선 행 = 응답에서 먼저 온 행)를 모아 단순 평균한다.
//    목표가 0·공란 행은 목표가가 아니므로 건너뛴다(규칙 문구 '마지막 목표가(0·공란 제외)'를 그대로 따름 — 그 증권사의
//    더 이른 목표가가 기간 안에 있으면 그 값을 쓴다. 2026-10-01 표본 56종목에서는 0인 행이 모두 같은 증권사의 목표가보다
//    이전 날짜라, '증권사의 마지막 행이 0이면 그 증권사를 뺀다'로 읽어도 결과가 같았다).
//    증권사명이 빈 행은 어느 증권사인지 알 수 없어 뺀다. 증권사는 이름(앞뒤 공백 제거)으로 구분한다.
//  - 기간 시작일 = Date.AddMonths(asOf, -months) — 없는 날은 그 달 마지막 날로 맞춘다(예: 3/31의 1개월 전 = 2/28).
//  - 입력: rows = fnTargetOpinions 형식의 표(일자 date, 증권사 text, 목표가 number — 없으면 null, 순번 number; 다른 열은 무시),
//    asOf = 기준일, months = 기간 개월 수(정수, 생략하면 6 — 설정 target_window_months).
//  - 반환: [목표가평균 = number 또는 null, 증권사수 = number 또는 null] — 목표가가 하나도 없으면 둘 다 null(0을 쓰지 않음).
//    레코드 필드는 읽을 때 계산되므로(지연 평가) rows의 오류(예: API 실패)를 try로 잡으려면 try 안에서 필드를 먼저 읽어야 한다.
//  - 사용 예) R = fnTargetOpinions("005930", #date(2026, 1, 1), #date(2026, 10, 1)),
//            fnTargetConsensus(R, #date(2026, 10, 1), 6)[목표가평균]
//            fnTargetConsensus(R, Date.AddMonths(#date(2026, 10, 1), -1), 6)   // 1개월 전 컨센서스
(rows as table, asOf as date, optional months as nullable number) as record =>
let
    Months = if months = null then 6 else months,
    Start = Date.AddMonths(asOf, -Months),
    InWindow = Table.SelectRows(rows, each [일자] <> null and [일자] > Start and [일자] <= asOf
        and [목표가] <> null and [목표가] > 0 and [증권사] <> null and Text.Trim([증권사]) <> ""),
    Brokers = List.Distinct(List.Transform(InWindow[증권사], Text.Trim)),
    Last = List.Transform(Brokers, (b) =>
        Table.First(Table.Sort(Table.SelectRows(InWindow, each Text.Trim([증권사]) = b),
            {{"일자", Order.Descending}, {"순번", Order.Ascending}}))[목표가]),
    N = List.Count(Last)
in
    [목표가평균 = if N = 0 then null else List.Average(Last), 증권사수 = if N = 0 then null else N]
