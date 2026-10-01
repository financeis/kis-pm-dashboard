// T_MktFunds : 국내 증시자금 일별 추이(억원) — 고객예탁금·신용융자·미수금·주식형 펀드·MMF  [국내 증시자금 종합 국내주식-193] FHKST649100C0
//  - 원천은 금융투자협회 집계(HTS [0470] 증시자금 종합). 1회 호출로 입력일 이하 최근 100영업일을 최신순으로 줌 → 전부 사용
//  - [모두 새로 고침]으로 갱신하는 일반 쿼리(호출 1회). 버튼·실행 모드(tblRunCtl)와 무관
//  - FID_INPUT_DATE_1 = 오늘: '그 날짜 이하' 의미라 휴일·주말이어도 직전 영업일까지 옴(빈 값도 같은 결과, 2026-10-01 실측)
//  - 공표가 늦음: 2026-10-01 오전 조회 시 최신 일자 2026-09-29(전 영업일 9/30 자료는 아직 없음) → 화면은 '일자'를 기준일로 표시
//
//  응답 필드 → 열 (다섯 금액 필드 모두 응답이 이미 억원 → 단위 변환 없음, ×1)
//    bsop_date      → 일자          "YYYYMMDD" → date
//    cust_dpmn_amt  → 고객예탁금억  고객(투자자)예탁금 잔고 (2026-09-29: 1,077,257억 ≈ 107.7조)
//    crdt_loan_rmnd → 신용융자억    신용융자 잔고 (329,200억 ≈ 32.9조, 하루 변화 중앙값 0.7%의 완만한 잔고)
//    uncl_amt       → 미수금억      (위탁매매) 미수금 (11,480억 ≈ 1.1조, 하루 변화 중앙값 7%로 크게 출렁임)
//    sttp_amt       → 주식형억      주식형 펀드 잔고 (3,427,598억 ≈ 342.8조). 하루 변화가 KOSPI 등락에 강하게 연동(일간 변화율의
//                                   KOSPI 베타 0.54, R² 0.93 — 평가액 기준으로 보임) → 1D·1W·1M 변화는 자금 유출입보다 시장 등락이 대부분
//    mmf_amt        → MMF억         MMF 잔고 (2,384,496억 ≈ 238.4조). 분기 말 급감(6/30 -5.7% 뒤 7/1~2 반등, 9월 말에도 감소) — MMF 특유의 계절성
//  단위·의미 검증(2026-10-01 실측, 응답 100행 2026-05-04 ~ 2026-09-29):
//    ① amt_tnrt(금액회전율 %) × cust_dpmn_amt = 같은 날 KOSPI+KOSDAQ 거래대금(지수 일봉 acml_tr_pbmn 백만원 ÷ 100)
//       — 지수 일봉과 겹치는 39영업일(8/3~9/29) 모두 일치(비율 0.9998~1.0002) → 고객예탁금은 억원.
//       나머지 금액 필드도 같은 원천·같은 문서 단위(억원)이고 크기도 각 항목에 맞는 수준(신용융자 수십 조, 미수금 약 1조)
//    ② cust_dpmn_amt_prdy_vrss = 전일 행과의 차이(±1억 이내 — 반올림 차이) → 행 = 영업일별 잔고
//    ③ bstp_nmix_prpr = 같은 날 KOSPI 종가(겹치는 39영업일 모두 일치, 날짜 목록도 KOSPI 영업일과 같음) → bsop_date와 값의 날짜가 맞음
//  쓰지 않는 필드: bstp_nmix_prpr(KOSPI 종가), prdy_ctrt(전일 대비 비율 — 등락률이 아님: 99.73 = -0.27%), amt_tnrt(금액회전율 %),
//    hts_avls·futs_tfam_amt·mxtp_amt·bntp_amt·secu_lend_amt(저장소 표기: 시가총액·선물예수금·혼합형·채권형·담보대출잔고)는 미검증
//  빈칸 규칙: 빈 문자열·0·음수는 '데이터 없음' → 빈칸(잔고는 0이 될 수 없으므로 0은 미공표로 봄). 다섯 값이 모두 빈 날짜 행은 버림
//  정렬: 일자 오름차순(오래된 날 → 최신) — 차트가 그대로 쓰고, 최신 기준일은 마지막 행
let
    Cols = {"일자", "고객예탁금억", "신용융자억", "미수금억", "주식형억", "MMF억", "상태", "조회시각"},
    Types = {{"일자", type date}, {"고객예탁금억", type number}, {"신용융자억", type number}, {"미수금억", type number},
             {"주식형억", type number}, {"MMF억", type number}, {"상태", type text}, {"조회시각", type datetime}},
    Main =
        let
            Path = "/uapi/domestic-stock/v1/quotations/mktfunds",
            Now = DateTime.LocalNow(),
            Today = fnYmdText(Date.From(Now)),
            J = fnKisGet(Path, "FHKST649100C0", [FID_INPUT_DATE_1 = Today]),
            O = Record.FieldOrDefault(J, "output", {}),
            L = List.Select(if O is list then O else if O is record then {O} else {},
                each _ is record and Text.Trim(Text.From(Record.FieldOrDefault(_, "bsop_date", ""))) <> ""),
            // 잔고 값: 빈 문자열·0·음수 → null(빈칸). 응답이 이미 억원이라 나누지 않음
            Amt = (r as record, k as text) as nullable number =>
                let v = fnNum(Record.FieldOrDefault(r, k, null)) in if v = null or v <= 0 then null else v,
            Rows = List.Transform(L, (r) => [
                일자 = fnYmd(r[bsop_date]),
                고객예탁금억 = Amt(r, "cust_dpmn_amt"),
                신용융자억 = Amt(r, "crdt_loan_rmnd"),
                미수금억 = Amt(r, "uncl_amt"),
                주식형억 = Amt(r, "sttp_amt"),
                MMF억 = Amt(r, "mmf_amt")]),
            T0 = Table.FromRecords(Rows, {"일자", "고객예탁금억", "신용융자억", "미수금억", "주식형억", "MMF억"}, MissingField.UseNull),
            HasData = Table.SelectRows(T0, each [일자] <> null
                and List.NonNullCount({[고객예탁금억], [신용융자억], [미수금억], [주식형억], [MMF억]}) > 0),
            // 행 수를 세어 응답(fnKisGet)을 여기서 강제 평가 → API 오류가 아래 try 안에서 잡힘.
            // 유효한 행이 하나도 없으면 오류로 올려 직전 표를 유지(fnFallback) — 빈 표로 덮어쓰지 않음
            Checked = if Table.RowCount(HasData) = 0 then error Error.Record("데이터 없음", "증시자금 응답에 유효한 행이 없습니다.") else HasData,
            // 같은 일자가 겹치면 응답 순서상 앞 행을 남김(최신순 응답)
            Dedup = Table.Distinct(Checked, {"일자"}),
            Sorted = Table.Sort(Dedup, {{"일자", Order.Ascending}}),
            WithStatus = Table.AddColumn(Table.AddColumn(Sorted, "상태", each "OK", type text), "조회시각", each Now, type datetime),
            Typed = Table.TransformColumnTypes(WithStatus, Types)
        in
            Typed,
    // 실패(토큰·네트워크·호출제한·빈 응답 등) 시 오류 대화상자 대신 직전 결과를 유지.
    // 직전 표를 Excel에서 다시 읽으면 열 형식이 없어(any) Excel이 일자·조회시각 셀 서식을 '일반'으로 바꾸므로(일련번호로 보임,
    // 2026-10-01 실측) 같은 형식을 다시 지정해 날짜 서식을 유지함 — 값은 그대로
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then Table.TransformColumnTypes(fnFallback("tblMktFunds", Cols, Tried[Error][Message]), Types) else Tried[Value]
in
    Final
