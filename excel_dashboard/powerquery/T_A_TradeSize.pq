// T_A_TradeSize : 종목분석 2구역 — 체결금액별 매매비중(당일) (spec R17-2) → tblA_TradeSize (2026-10-01 T15)
//  - 원천: 국내주식 체결금액별 매매비중 FHKST111900C0 /uapi/domestic-stock/v1/quotations/tradprt-byamt (당일 자료만, 8구간 1회 호출)
//    + 기준일(당일 표시)은 fnAQuote(최근 일봉 일자, 호출 1회) — 장중·장 마감 뒤 오늘, 장 시작 전·휴장일은 직전 세션.
//  - 구간: 응답 prpr_name("3백 이하"…"5억 초과", 만원 단위 표기)을 "3백만 이하"…"5억 초과"로 바꿔 씀. 순서 = 응답 순서(작은 체결금액부터).
//  - 열(2026-10-01 실측으로 의미 확인 — 005930):
//      매수거래량 shnu_cnqn_smtn / 매도거래량 seln_cnqn_smtn (그 구간 체결의 매수·매도 주체 거래량, 구간 합계 = 종목 당일 거래량)
//      매수비율 whol_shun_vol_rate / 매도비율 whol_seln_vol_rate / 순매수비율 whol_ntby_qty_rate : 종목 당일 전체 거래량 대비(% → 소수)
//        (8구간 매수비율 + 매도비율 합 = 100%, 순매수비율 = 매수비율 − 매도비율)
//      매수건수 shnu_cntg_csnu / 매도건수 seln_cntg_csnu / 순매수건수 ntby_cntg_csnu(= 매수건수 − 매도건수), 순매수거래량 = 매수 − 매도
//      평균가 smtn_avrg_prpr (0 → 빈칸)
//  - 거래량 합계가 0이면(장 시작 전·거래 없음·없는 코드 — 없는 코드도 rt_cd 0에 0으로 옴) "데이터 없음" 1행.
let
    Types = {{"종목코드", type text}, {"기준일", type date}, {"순서", Int64.Type}, {"구간", type text},
             {"매수거래량", type number}, {"매도거래량", type number}, {"순매수거래량", type number},
             {"매수비율", type number}, {"매도비율", type number}, {"순매수비율", type number},
             {"매수건수", type number}, {"매도건수", type number}, {"순매수건수", type number}, {"평균가", type number},
             {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Qt = fnAQuote(code, Date.From(now)),
            Base = Qt[기준일],
            J = fnKisGet("/uapi/domestic-stock/v1/quotations/tradprt-byamt", "FHKST111900C0",
                    [FID_COND_MRKT_DIV_CODE = "J", FID_COND_SCR_DIV_CODE = "11119", FID_INPUT_ISCD = code]),
            O = Record.FieldOrDefault(J, "output", {}),
            L = List.Buffer(List.Select(if O is list then O else if O is record then {O} else {}, each _ is record)),
            N = (r as record, k as text) as nullable number => fnNum(Record.FieldOrDefault(r, k, null)),
            Pct = (r as record, k as text) as nullable number => let v = N(r, k) in if v = null then null else v / 100,
            Label = (r as record) as nullable text =>
                let v = Record.FieldOrDefault(r, "prpr_name", null), t = if v = null then "" else Text.Trim(Text.From(v))
                in if t = "" then null else Text.Replace(Text.Replace(t, "백 ", "백만 "), "천 ", "천만 "),
            Volumes = List.Transform(L, (r) => List.Sum(List.RemoveNulls({N(r, "shnu_cnqn_smtn"), N(r, "seln_cnqn_smtn")}))),
            Total = List.Sum(List.RemoveNulls(Volumes)),
            Rows = List.Transform(List.Positions(L), (i) =>
                let
                    r = L{i},
                    buy = N(r, "shnu_cnqn_smtn"),
                    sell = N(r, "seln_cnqn_smtn"),
                    avg = N(r, "smtn_avrg_prpr")
                in
                    {code, Base, i + 1, Label(r), buy, sell, if buy = null or sell = null then null else buy - sell,
                     Pct(r, "whol_shun_vol_rate"), Pct(r, "whol_seln_vol_rate"), Pct(r, "whol_ntby_qty_rate"),
                     N(r, "shnu_cntg_csnu"), N(r, "seln_cntg_csnu"), N(r, "ntby_cntg_csnu"),
                     if avg = null or avg <= 0 then null else avg, "OK", now})
        in
            if List.IsEmpty(L) or Total = null or Total = 0 then
                Table.TransformColumns(fnAOneRow(Types, code, "데이터 없음", now), {{"기준일", each Base, type date}})
            else #table(List.Transform(Types, each _{0}), Rows)
in
    fnARun("tblA_TradeSize", Types, Main)
