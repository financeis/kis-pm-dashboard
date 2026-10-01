// fnTargetOpinions : 한 종목의 종목투자의견 원천 행 (2026-10-01)
//  - API: 국내주식 종목투자의견[국내주식-188] FHKST663300C0 /uapi/domestic-stock/v1/quotations/invest-opinion
//    입력 FID_COND_MRKT_DIV_CODE = J, FID_COND_SCR_DIV_CODE = 16633, FID_INPUT_DATE_1 = 시작(이후), FID_INPUT_DATE_2 = 끝(이전),
//    날짜는 YYYYMMDD, 기간은 양 끝 포함.
//  - 한 번의 응답은 그 기간의 최신 100건까지(최신순)이고 연속 조회가 되지 않는다(2026-10-01 실측: 삼성전자 2026-01-01~10-01
//    요청 → 100건, 2026-01-30~09-23, 응답 헤더 tr_cont ""; 요청 헤더 tr_cont = N으로 다시 보내도 같은 100건). 그래서 날짜를 나눠 이어 받는다:
//    100건이 꽉 찬 페이지는 그 페이지의 가장 오래된 날짜(oldest)보다 새 행만 쓰고 [시작, oldest]를 다시 요청한다.
//    잘린 날짜(oldest)의 행은 다음 요청이 그 날짜째로 다시 받으므로 빠지거나 겹치는 행이 없다. 겹치게 받은 뒤 중복을 지우는
//    방식을 쓰지 않는 이유: 괴리도·괴리율 필드(stft_esdg·dprt)가 요청 끝 날짜에 따라 값이 달라 같은 의견도 행이 같지 않다.
//    멈춤: 100건 미만이거나, oldest가 시작일 이하이거나, oldest가 요청 끝 날짜와 같거나(한 날짜에 100건 이상), 10회에 이름.
//  - 반환 열(응답 필드):
//      일자     stck_bsop_date (YYYYMMDD → date, 형식이 틀린 행과 기간 밖 행은 버림)
//      증권사   mbcr_name — 저장소 예제의 COLUMN_MAPPING에는 없지만 응답에 있음(미래에셋·KB·신한투자증권 등). 빈 값은 null
//      의견     invt_opnn — 증권사 표기 그대로(BUY·매수·Buy·Strong BUY·Hold·중립·Neutral·NotRated 등)
//      직전의견 rgbf_invt_opnn
//      목표가   hts_goal_prc(원) — 0·공란·음수는 '목표가 없음'이라 null(NotRated와 일부 중립 보고서가 0으로 옴)
//      순번     응답 순서(0부터, 최신순으로 이어 붙인 순서) — 같은 날짜 안의 선후를 정하는 기준(앞이 최근)
//    쓰지 않는 필드: invt_opnn_cls_code·rgbf_invt_opnn_cls_code(의견 구분 코드), stck_prdy_clpr(전일 종가),
//    stck_nday_esdg·nday_dprt·stft_esdg·dprt(괴리도·괴리율 — 요청 기간에 따라 바뀜).
//  - 오류: fnKisGet 오류(설정·토큰·rt_cd <> 0)를 그대로 올린다 — 반환 표를 처음 평가할 때(행 수를 셀 때 등) 표 전체의 오류로 난다.
//    M은 지연 평가라 호출한 쪽의 종목별 try가 오류를 잡으려면 try 안에서 결과를 먼저 평가해야 한다(예: Table.RowCount(결과)).
//    응답은 모두 받아 버퍼 표로 돌려주므로 결과를 여러 번 읽어도 다시 호출하지 않는다.
//    의견이 없는 종목은 rt_cd 0에 빈 목록으로 오므로 빈 표(오류 아님).
//  - 사용 예) fnTargetOpinions("005930", #date(2026, 1, 1), #date(2026, 10, 1))
(code as text, fromDate as date, toDate as date) as table =>
let
    Path = "/uapi/domestic-stock/v1/quotations/invest-opinion",
    PageSize = 100,
    MaxPages = 10,
    FromText = fnYmdText(fromDate),
    Txt = (r as record, k as text) as text =>
        let v = Record.FieldOrDefault(r, k, null) in if v = null then "" else try Text.Trim(Text.From(v)) otherwise "",
    Blank = (t as text) as nullable text => if t = "" then null else t,
    // 기간 [FromText, d2]의 한 페이지(최신 100건까지)
    Page = (d2 as text) as list =>
        let
            J = fnKisGet(Path, "FHKST663300C0", [FID_COND_MRKT_DIV_CODE = "J", FID_COND_SCR_DIV_CODE = "16633",
                    FID_INPUT_ISCD = code, FID_INPUT_DATE_1 = FromText, FID_INPUT_DATE_2 = d2]),
            O = Record.FieldOrDefault(J, "output", {}),
            L = if O is list then O else if O is record then {O} else {}
        in
            List.Buffer(List.Select(L, each _ is record)),
    // 날짜 분할 반복: 꽉 찬 페이지는 oldest보다 새 행만 남기고 [시작, oldest]를 다시 요청
    Loop = (d2 as text, n as number, acc as list) as list =>
        let
            L = Page(d2),
            Dates = List.Select(List.Transform(L, each Txt(_, "stck_bsop_date")), each Text.Length(_) = 8),
            Oldest = if List.IsEmpty(Dates) then "" else List.Min(Dates),
            More = List.Count(L) >= PageSize and Oldest > FromText and Oldest < d2 and n + 1 < MaxPages,
            Kept = if More then List.Select(L, each Txt(_, "stck_bsop_date") > Oldest) else L
        in
            if More then @Loop(Oldest, n + 1, acc & Kept) else acc & Kept,
    Raw = List.Buffer(Loop(fnYmdText(toDate), 0, {})),
    Rows = List.Transform(List.Positions(Raw), (i) =>
        let
            r = Raw{i},
            t = fnNum(Txt(r, "hts_goal_prc"))
        in
            {fnYmd(Txt(r, "stck_bsop_date")), Blank(Txt(r, "mbcr_name")), Blank(Txt(r, "invt_opnn")),
             Blank(Txt(r, "rgbf_invt_opnn")), if t = null or t <= 0 then null else t, i}),
    InRange = List.Select(Rows, each _{0} <> null and _{0} >= fromDate and _{0} <= toDate),
    Result = #table(type table [일자 = date, 증권사 = nullable text, 의견 = nullable text, 직전의견 = nullable text,
                                목표가 = nullable number, 순번 = Int64.Type], InRange)
in
    Table.Buffer(Result)
