// T_Rank : 시장 주도주/특징주 순위 (거래대금·거래량급증·등락률·외국인/기관 순매수·신고가근접)
//  - 순위 API: 거래량순위 FHPST01710000, 등락률 FHPST01700000, 기관·외국인 가집계 FHPTJ04400000, 신고/신저근접 FHPST01870000
//  - 시장 전체 개별종목(2026-10-01): API의 ETF/ETN/SPAC 제외 + 종목DB(tblUniverse)에 있는 종목만(ETF·ETN·리츠는
//    종목DB에 없음) + 종목DB의 SPAC = Y 제외. 대회편입으로는 거르지 않는다(9/30 뒤 주도주까지 보이게).
//  - 열 `대회` = ★(대회 종목) / 공란 — fnClassify의 대회편입(명단 ± 수정표; 규칙을 여기서 다시 구현하지 않음).
//    종목명은 바꾸지 않는다(화면의 보유·관심 강조가 종목명으로 찾기 때문). 화면은 종목명 옆 칸에 ★, 대회 밖은 회색 글꼴.
//  - 거른 뒤 목록별 순위를 다시 매긴다(원래 순서 기준). 섹터·시장은 종목DB 값(섹터 = 설정 sector_basis 기준).
let
    Main =
        let
            S = fnSettings(),
            Mk = let m = try Text.From(S[rank_market]) otherwise "전체" in
                 if m = "코스피" or m = "KOSPI" then "0001" else if m = "코스닥" or m = "KOSDAQ" then "1001" else "0000",
            Now = DateTime.LocalNow(),
            Q = "/uapi/domestic-stock/v1/quotations/",
            R = "/uapi/domestic-stock/v1/ranking/",
            Exls = "0000001101",   // 투자위험/경고/주의,관리,정리매매,불성실,우선주,거래정지,ETF,ETN,신용불가,SPAC (1=제외)
            Out = (j as record) as list => let o = Record.FieldOrDefault(j, "output", {}) in if o is list then o else {},
            N = (r as record, k as text) as nullable number => fnNum(Record.FieldOrDefault(r, k, null)),
            Pct = (r as record, k as text) as nullable number => let v = N(r, k) in if v = null then null else v / 100,

            VolRank = (blng as text, volMin as text, p1 as text, p2 as text) as list => Out(fnKisGet(Q & "volume-rank", "FHPST01710000", [
                FID_COND_MRKT_DIV_CODE = "J", FID_COND_SCR_DIV_CODE = "20171", FID_INPUT_ISCD = Mk, FID_DIV_CLS_CODE = "0",
                FID_BLNG_CLS_CODE = blng, FID_TRGT_CLS_CODE = "111111111", FID_TRGT_EXLS_CLS_CODE = Exls,
                FID_INPUT_PRICE_1 = p1, FID_INPUT_PRICE_2 = p2, FID_VOL_CNT = volMin, FID_INPUT_DATE_1 = ""])),
            Fluct = (sort as text) as list => Out(fnKisGet(R & "fluctuation", "FHPST01700000", [
                fid_cond_mrkt_div_code = "J", fid_cond_scr_div_code = "20170", fid_input_iscd = Mk, fid_rank_sort_cls_code = sort,
                fid_input_cnt_1 = "0", fid_prc_cls_code = "1", fid_input_price_1 = "", fid_input_price_2 = "", fid_vol_cnt = "",
                fid_trgt_cls_code = "0", fid_trgt_exls_cls_code = Exls, fid_div_cls_code = "0", fid_rsfl_rate1 = "", fid_rsfl_rate2 = ""])),
            FrgnInst = (sort as text, who as text) as list => Out(fnKisGet(Q & "foreign-institution-total", "FHPTJ04400000", [
                FID_COND_MRKT_DIV_CODE = "V", FID_COND_SCR_DIV_CODE = "16449", FID_INPUT_ISCD = Mk, FID_DIV_CLS_CODE = "1",
                FID_RANK_SORT_CLS_CODE = sort, FID_ETC_CLS_CODE = who])),
            NearHigh = () as list => Out(fnKisGet(R & "near-new-highlow", "FHPST01870000", [
                fid_aply_rang_vol = "100", fid_cond_mrkt_div_code = "J", fid_cond_scr_div_code = "20187", fid_div_cls_code = "0",
                fid_input_cnt_1 = "0", fid_input_cnt_2 = "3", fid_prc_cls_code = "0", fid_input_iscd = Mk, fid_trgt_cls_code = "0",
                fid_trgt_exls_cls_code = "0", fid_aply_rang_prc_1 = "", fid_aply_rang_prc_2 = ""])),

            Code = (r as record) as text => Text.From(Record.FieldOrDefault(r, "mksc_shrn_iscd", Record.FieldOrDefault(r, "stck_shrn_iscd", ""))),
            Base = (r as record, listName as text, rank as number) as record => [
                목록 = listName, 순위 = rank, 종목코드 = Code(r),
                종목명 = Text.From(Record.FieldOrDefault(r, "hts_kor_isnm", "")),
                현재가 = N(r, "stck_prpr"), 등락률 = Pct(r, "prdy_ctrt"), 거래량 = N(r, "acml_vol")],
            Build = (listName as text, rows as list, extra as function) as list =>
                List.Transform(List.Positions(rows), (i) => Base(rows{i}, listName, i + 1) & extra(rows{i})),

            Lists = {
                {"거래대금상위", () => Build("거래대금상위", VolRank("3", "", "", ""), (r) =>
                    [거래대금억 = let v = N(r, "acml_tr_pbmn") in if v = null then null else v / 100000000, 지표 = Pct(r, "vol_inrt"), 지표명 = "거래증가율"])},
                {"거래량급증", () => Build("거래량급증", VolRank("1", "100000", "1000", "99999999"), (r) =>
                    [거래대금억 = let v = N(r, "acml_tr_pbmn") in if v = null then null else v / 100000000, 지표 = Pct(r, "vol_inrt"), 지표명 = "거래증가율"])},
                {"상승률상위", () => Build("상승률상위", Fluct("0"), (r) =>
                    [거래대금억 = let p = N(r, "stck_prpr"), v = N(r, "acml_vol") in if p = null or v = null then null else p * v / 100000000,
                     지표 = N(r, "cnnt_ascn_dynu"), 지표명 = "연속상승일"])},
                {"하락률상위", () => Build("하락률상위", Fluct("1"), (r) =>
                    [거래대금억 = let p = N(r, "stck_prpr"), v = N(r, "acml_vol") in if p = null or v = null then null else p * v / 100000000,
                     지표 = N(r, "cnnt_down_dynu"), 지표명 = "연속하락일"])},
                {"외국인순매수", () => Build("외국인순매수", FrgnInst("0", "1"), (r) =>
                    [거래대금억 = null, 지표 = let v = N(r, "frgn_ntby_tr_pbmn") in if v = null then null else v / 100, 지표명 = "외국인순매수(억)"])},
                {"외국인순매도", () => Build("외국인순매도", FrgnInst("1", "1"), (r) =>
                    [거래대금억 = null, 지표 = let v = N(r, "frgn_ntby_tr_pbmn") in if v = null then null else v / 100, 지표명 = "외국인순매수(억)"])},
                {"기관순매수", () => Build("기관순매수", FrgnInst("0", "2"), (r) =>
                    [거래대금억 = null, 지표 = let v = N(r, "orgn_ntby_tr_pbmn") in if v = null then null else v / 100, 지표명 = "기관순매수(억)"])},
                {"기관순매도", () => Build("기관순매도", FrgnInst("1", "2"), (r) =>
                    [거래대금억 = null, 지표 = let v = N(r, "orgn_ntby_tr_pbmn") in if v = null then null else v / 100, 지표명 = "기관순매수(억)"])},
                {"신고가근접", () => Build("신고가근접", NearHigh(), (r) =>
                    [거래대금억 = let p = N(r, "stck_prpr"), v = N(r, "acml_vol") in if p = null or v = null then null else p * v / 100000000,
                     지표 = let v = N(r, "hprc_near_rate") in if v = null then null else v / 100, 지표명 = "신고가괴리"])}
            },
            Rows = List.Combine(List.Transform(Lists, (x) =>
                let r = try List.Buffer(x{1}()) in
                    if r[HasError] then {[목록 = x{0}, 순위 = 0, 종목명 = "오류: " & r[Error][Message]]} else r[Value])),
            Cols = {"목록", "순위", "종목코드", "종목명", "현재가", "등락률", "거래량", "거래대금억", "지표", "지표명"},
            Errs = List.Select(Rows, each Text.StartsWith(Text.From(Record.FieldOrDefault(_, "종목명", "")), "오류")),
            // 모든 목록이 실패하면 오류로 올려 직전 데이터를 유지(fnFallback)
            Checked = if List.Count(Rows) > 0 and List.Count(Errs) = List.Count(Rows) then error Text.Middle(Errs{0}[종목명], 4) else Rows,
            T0 = Table.FromRecords(Checked, Cols, MissingField.UseNull),
            // 원래 순서(목록 순서 + API 순위)를 조인 전에 기록 — 조인은 행 순서를 보장하지 않음
            T = Table.AddIndexColumn(T0, "__ord", 0, 1, Int64.Type),

            // 종목DB(이전 새로고침 결과)가 있으면 시장 전체 개별종목(종목DB에 있는 종목, SPAC 제외)만 남기고 섹터·시장을 붙임.
            // 종목DB가 없거나 비었으면 거르지 않는다(API의 ETF·ETN·SPAC 제외만 적용)
            U0 = try Excel.CurrentWorkbook(){[Name = "tblUniverse"]}[Content] otherwise null,
            U = if U0 = null or not Table.HasColumns(U0, "종목코드") then null
                else Table.Buffer(Table.Distinct(Table.SelectColumns(U0, {"종목코드", "섹터", "시장", "SPAC", "KRX업종"}, MissingField.UseNull), {"종목코드"})),
            Filtered =
                if U = null or Table.RowCount(U) = 0 then
                    Table.AddColumn(Table.AddColumn(Table.AddColumn(T, "섹터", each null), "시장", each null), "KRX업종", each null)
                else
                    let
                        Uin = Table.AddColumn(U, "__in", each true, type logical),
                        J = Table.NestedJoin(T, {"종목코드"}, Uin, {"종목코드"}, "u", JoinKind.LeftOuter),
                        E = Table.ExpandTableColumn(J, "u", {"섹터", "시장", "SPAC", "KRX업종", "__in"}),
                        K = Table.SelectRows(E, each [순위] = 0 or ([__in] = true and [SPAC] <> "Y"))
                    in
                        Table.RemoveColumns(K, {"SPAC", "__in"}),
            // 대회 ★ = fnClassify의 대회편입 Y (오류 행처럼 종목코드가 없는 행은 빼고 넘김 → 공란)
            ClsIn = Table.SelectColumns(Table.SelectRows(Filtered, each [종목코드] <> null and [종목코드] <> ""), {"종목코드", "KRX업종"}),
            Members = List.Buffer(Table.SelectRows(fnClassify(ClsIn), each [대회편입] = "Y")[종목코드]),
            Starred = Table.AddColumn(Table.RemoveColumns(Filtered, {"KRX업종"}), "대회", each
                let c = if [종목코드] = null then null else Text.Upper(Text.Trim(Text.From([종목코드])))
                in if c <> null and List.Contains(Members, c) then "★" else null, type text),
            // 필터 후 목록별 순위 다시 매기기 (원래 순서 기준)
            Ordered = Table.Sort(Starred, {{"__ord", Order.Ascending}}),
            Regroup = Table.Group(Ordered, {"목록"}, {{"t", each
                Table.AddIndexColumn(Table.Sort(Table.RemoveColumns(_, {"순위"}), {{"__ord", Order.Ascending}}), "순위", 1, 1, Int64.Type), type table}}),
            Expanded = Table.ExpandTableColumn(Regroup, "t", {"순위", "종목코드", "종목명", "현재가", "등락률", "거래량", "거래대금억", "지표", "지표명", "섹터", "시장", "대회", "__ord"}),
            Sorted = Table.Sort(Expanded, {{"__ord", Order.Ascending}}),
            Final = Table.AddColumn(Table.RemoveColumns(Sorted, {"__ord"}), "조회시각", each Now, type datetime),
            Typed = Table.TransformColumnTypes(Final, {{"목록", type text}, {"순위", Int64.Type}, {"종목코드", type text}, {"종목명", type text},
                {"현재가", type number}, {"등락률", type number}, {"거래량", type number}, {"거래대금억", type number}, {"지표", type number},
                {"지표명", type text}, {"섹터", type text}, {"시장", type text}, {"대회", type text}}),
            Keyed = Table.AddColumn(Typed, "키", each [목록] & "|" & Text.From([순위]), type text)
        in
            Keyed,
    // 실패(토큰·네트워크·호출제한 등) 시 오류 대화상자 대신 직전 결과를 유지
    Tried = try Table.Buffer(Main),
    Final = if Tried[HasError] then fnFallback("tblRank", {"목록", "순위", "종목코드", "종목명", "현재가", "등락률", "거래량", "거래대금억", "지표", "지표명", "섹터", "시장", "대회", "조회시각", "키"}, Tried[Error][Message]) else Tried[Value]
in
    Final
