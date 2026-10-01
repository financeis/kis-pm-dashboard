// fnFwdEps : 날짜 asOf의 12개월 선행 EPS(Fwd EPS) — KIS 추정 FY1·FY2 EPS의 가중 평균 (2026-10-01)
//  - 식: Fwd EPS = FY1 EPS × r/12 + FY2 EPS × (12 − r)/12,
//        r = asOf가 속한 달부터 FY1 결산월까지의 달 수(양 끝 포함) = (FY1 연 × 12 + FY1 월) − (asOf 연 × 12 + asOf 월) + 1.
//    예) 12월 결산·10월 → r = 3(10·11·12월), 3월 결산 FY1 2027.03E·2026년 10월 → r = 6, 결산월과 같은 달 → r = 1,
//        결산 다음 달부터 시작하는 회계연도의 첫 달(12월 결산이면 1월) → r = 12(= FY1 EPS 그대로).
//  - FY2 EPS가 없으면(null) FY1 EPS를 그대로 돌려준다(FY2가 없으면 FY1 EPS만 쓰고 그렇다고 표시 — 표시는 쓰는 쪽이
//    tblEst의 FY2 빈칸으로 판단한다).
//  - null을 돌려주는 경우(계산할 수 없음):
//      FY1 라벨을 해석할 수 없음 · FY1 EPS 없음 · asOf 없음 ·
//      r < 1 — asOf가 FY1 결산월보다 뒤(FY1 회계연도가 이미 끝남) ·
//      r > 12 — asOf가 FY1 회계연도 시작 전. 그 날의 12개월 선행 EPS에는 지금의 FY1·FY2에 없는 직전 회계연도 EPS가
//               필요하므로 지어내지 않는다(예: 1월 초에 [전체]를 눌러 FY1이 다음 해로 넘어간 뒤 전년 12월 날짜의 값).
//    → 스냅샷(T_EstSnap)은 그 종목 행을 만들지 않고, 대회종목 표의 Fwd PER 1W·1M 재구성은 공란이 된다.
//  - 입력(형식 없는 셀 값도 받음 — Excel 표를 Excel.CurrentWorkbook으로 읽은 값을 그대로 넘겨도 됨):
//      fy1    FY1 결산년월 라벨(tblEst FY1, 예 "2026.12E"). 숫자만 모아 앞 4자리 = 연, 다음 2자리 = 월
//             ("2026.12", "2026.12(E)"도 같음). 숫자로 와도 글자로 바꿔 같은 규칙.
//      fy1Eps, fy2Eps  원 단위 EPS(tblEst FY1_EPS·FY2_EPS — ÷10 보정 후). 숫자 또는 숫자 글자, 없으면 null
//      asOf   기준 날짜(date, datetime — 날짜 부분만, 또는 Excel 일련번호)
//  - 반환: number(원, 반올림 안 함) 또는 null.
//  - 사용 예) fnFwdEps("2026.12E", 46209, 67239.8, #date(2026, 10, 1))   // = 46209 × 3/12 + 67239.8 × 9/12 = 61982.1
//            fnFwdEps(e[FY1], e[FY1_EPS], e[FY2_EPS], 기준세션일)          // 대회종목 표의 재구성(기준일 가중치)
(fy1 as any, fy1Eps as any, fy2Eps as any, asOf as any) as nullable number =>
let
    Num = (v as any) as nullable number =>
        if v = null then null else if v is number then v else if v is text then fnNum(v) else null,
    Label = if fy1 = null then "" else (try Text.From(fy1) otherwise ""),
    Digits = Text.Select(Label, {"0".."9"}),
    Year = if Text.Length(Digits) >= 6 then Number.From(Text.Start(Digits, 4)) else null,
    Month = if Year = null then null else Number.From(Text.Middle(Digits, 4, 2)),
    AsOfDate =
        if asOf = null then null
        else if asOf is date then asOf
        else if asOf is datetime then Date.From(asOf)
        else if asOf is datetimezone then Date.From(DateTimeZone.RemoveZone(asOf))
        else try Date.From(asOf) otherwise null,
    E1 = Num(fy1Eps),
    E2 = Num(fy2Eps),
    R = if Year = null or Month < 1 or Month > 12 or AsOfDate = null then null
        else (Year * 12 + Month) - (Date.Year(AsOfDate) * 12 + Date.Month(AsOfDate)) + 1,
    Result =
        if R = null or E1 = null or R < 1 or R > 12 then null
        else if E2 = null then E1
        else E1 * R / 12 + E2 * (12 - R) / 12
in
    Result
