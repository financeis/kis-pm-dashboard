// T_A_Info : 종목분석 머리 정보 1행 → tblA_Info (2026-10-01)
//  - 실행 규칙은 fnARun: KIS 호출은 [조회](tblRunCtl mode = analysis)에서만, 다른 모드는 직전 표 그대로, 분석코드가 틀리면 "오류: 종목코드 확인".
//  - 원천(호출 2회):
//      주식기본조회 CTPF1002R search-stock-info      → 종목명(prdt_abrv_name), 시장(mket_id_cd: STK→KOSPI, KSQ→KOSDAQ, KNX→KONEX)
//      주식현재가 시세 FHKST01010100 inquire-price     → 현재가·전일대비·등락률(%→소수)·거래대금(원→억원)·시가총액(hts_avls, 억원)·PER·PBR
//                                                       ·외국인소진율(%→소수)·52주 고저·유의(T_Quote와 같은 규칙)·KRX 업종명(bstp_kor_isnm)
//  - 분류: fnClassify(종목코드, KRX업종) → 대회편입·대회비고·NICS 3단계·대테마·세부테마·분류출처 (수정표 규칙을 여기서 다시 구현하지 않음).
//    KRX업종 = tblUniverse에 KRX업종 열이 있으면 그 값, 없으면 inquire-price의 업종명. NICS가 없을 때의 대체에만 쓰인다.
//    KSIC 세분류·주요상품·기업집단 = tblClass(종목코드가 빈 경고 행은 건너뜀).
//  - 상태: OK / "오류: 종목코드 확인(KIS 시세 없음)" — 형식은 맞지만 상장 시세가 없는 코드. 시장(mket_id_cd)도 현재가도 비어 오면 이 상태
//    (999999 실측: 주식기본조회는 비상장 법인 이름을 주지만 시장은 빈칸, 시세는 모두 0). 이때 시세 열은 모두 빈칸.
//  - 0·빈 값은 빈칸(PER·PBR 0 포함).
let
    Types = {{"종목코드", type text}, {"종목명", type text}, {"시장", type text}, {"대회편입", type text}, {"대회비고", type text},
             {"NICS 대분류", type text}, {"NICS 업종", type text}, {"NICS 세부", type text}, {"대테마", type text}, {"세부테마", type text},
             {"분류출처", type text}, {"KRX업종", type text}, {"KSIC 세분류", type text}, {"주요상품", type text}, {"기업집단", type text},
             {"현재가", type number}, {"전일대비", type number}, {"등락률", type number}, {"거래대금억", type number}, {"시가총액억", type number},
             {"PER", type number}, {"PBR", type number}, {"외국인소진율", type number}, {"고52주", type number}, {"저52주", type number},
             {"유의", type text}, {"상태", type text}, {"조회시각", type datetime}},
    Main = (code as text, now as datetime) as table =>
        let
            Q = "/uapi/domestic-stock/v1/quotations/",
            Txt = (v as any) as nullable text =>
                if v = null then null else let t = try Text.Trim(Text.From(v)) otherwise null in if t = null or t = "" then null else t,
            Norm = (v as any) as nullable text =>
                let t0 = Txt(v), t = if t0 = null then null else Text.Upper(t0)
                in if t <> null and Text.Length(t) < 6 and Text.Select(t, {"0".."9"}) = t then Text.PadStart(t, 6, "0") else t,
            Rec = (j as record, k as text) as record => let o = Record.FieldOrDefault(j, k, null) in if o is record then o else [],
            SI = Rec(fnKisGet(Q & "search-stock-info", "CTPF1002R", [PRDT_TYPE_CD = "300", PDNO = code]), "output"),
            P = Rec(fnKisGet(Q & "inquire-price", "FHKST01010100", [FID_COND_MRKT_DIV_CODE = "J", FID_INPUT_ISCD = code]), "output"),
            F = (r as record, k as text) as nullable text => Txt(Record.FieldOrDefault(r, k, null)),
            N = (k as text) as nullable number => fnNum(Record.FieldOrDefault(P, k, null)),
            NZ = (k as text) as nullable number => let v = N(k) in if v = null or v = 0 then null else v,
            Pos = (k as text) as nullable number => let v = N(k) in if v = null or v <= 0 then null else v,
            Pct = (k as text) as nullable number => let v = N(k) in if v = null then null else v / 100,

            Name = F(SI, "prdt_abrv_name"),
            MarketCode = F(SI, "mket_id_cd"),
            Market = if MarketCode = "STK" then "KOSPI" else if MarketCode = "KSQ" then "KOSDAQ" else if MarketCode = "KNX" then "KONEX" else MarketCode,
            Price = Pos("stck_prpr"),

            // KRX 업종: tblUniverse의 KRX업종 열 → 없으면 현재가 시세의 업종명
            Univ = try Excel.CurrentWorkbook(){[Name = "tblUniverse"]}[Content] otherwise null,
            UnivKrx =
                if Univ = null or not List.ContainsAll(Table.ColumnNames(Univ), {"종목코드", "KRX업종"}) then null
                else let m = Table.SelectRows(Univ, each Norm([종목코드]) = code) in if Table.IsEmpty(m) then null else Txt(m{0}[KRX업종]),
            Krx = if UnivKrx <> null then UnivKrx else F(P, "bstp_kor_isnm"),
            Cls = fnClassify(#table({"종목코드", "KRX업종"}, {{code, Krx}})),
            C = if Table.IsEmpty(Cls) then [] else Cls{0},
            Cf = (k as text) as nullable text => Txt(Record.FieldOrDefault(C, k, null)),

            // tblClass(KSIC·주요상품·기업집단) — 종목코드가 빈 경고 행은 건너뜀
            ClassT = try Excel.CurrentWorkbook(){[Name = "tblClass"]}[Content] otherwise null,
            ClassRow =
                if ClassT = null or not List.Contains(Table.ColumnNames(ClassT), "종목코드") then []
                else let m = Table.SelectRows(ClassT, each Norm([종목코드]) = code) in if Table.IsEmpty(m) then [] else m{0},
            Kf = (k as text) as nullable text => Txt(Record.FieldOrDefault(ClassRow, k, null)),

            // 유의 (T_Quote와 같은 규칙)
            Warn = F(P, "mrkt_warn_cls_code"),
            Vi = F(P, "vi_cls_code"),
            Flags = List.RemoveNulls({
                if F(P, "temp_stop_yn") = "Y" then "거래정지" else null,
                if Warn = "01" then "투자주의" else if Warn = "02" then "투자경고" else if Warn = "03" then "투자위험" else null,
                if F(P, "invt_caful_yn") = "Y" then "투자유의" else null,
                if F(P, "short_over_yn") = "Y" then "단기과열" else null,
                if F(P, "mang_issu_cls_code") = "Y" then "관리종목" else null,
                if Vi <> null and Vi <> "N" then "VI발동" else null}),
            Found = Market <> null or Price <> null,
            // 시세가 없으면(상장 종목 아님) 0으로 오는 시세 열을 모두 빈칸으로
            Q0 = (v as nullable number) as nullable number => if Price = null then null else v,
            Row = {code, Name, Market, Cf("대회편입"), Cf("대회비고"), Cf("NICS 대분류"), Cf("NICS 업종"), Cf("NICS 세부"),
                   Cf("대테마"), Cf("세부테마"), Cf("분류출처"), Krx, Kf("KSIC 세분류"), Kf("주요상품"), Kf("기업집단"),
                   Price, Q0(N("prdy_vrss")), Q0(Pct("prdy_ctrt")), Q0(let v = N("acml_tr_pbmn") in if v = null then null else v / 100000000),
                   Q0(Pos("hts_avls")), Q0(NZ("per")), Q0(NZ("pbr")), Q0(Pct("hts_frgn_ehrt")), Q0(Pos("w52_hgpr")), Q0(Pos("w52_lwpr")),
                   if List.IsEmpty(Flags) then null else Text.Combine(Flags, ","),
                   if Found then "OK" else "오류: 종목코드 확인(KIS 시세 없음)", now}
        in
            #table(List.Transform(Types, each _{0}), {Row})
in
    fnARun("tblA_Info", Types, Main)
