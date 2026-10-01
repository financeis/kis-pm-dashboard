// fnPxRunCtl : 버튼 실행 모드와 '지금'(시험용 시계) — T_Sessions·T_PxStore·T_PxMetrics 공용 (2026-10-01)
//  - tblRunCtl(키, 값)의 mode: quick([시세]) / full([전체]) / sector / analysis / build. 비었거나 모르는 값이면 build로 본다.
//    KIS를 부르는 것은 quick·full뿐이고, 그 밖의 모드(빌드·표 우클릭 새로 고침·파일 열기)는 호출 없이 저장값을 돌려준다.
//  - now_override(일시, 시험 전용 — 평소 비어 있음): 있으면 시스템 시각 대신 이 값을 '지금'으로 쓴다(장 시작 전·휴장일·
//    안전 규칙 모사). 이 시계는 판단에만 쓰고, 각 표의 조회시각 열에는 실제 시각(DateTime.LocalNow)을 적는다.
//  - started(일시): 버튼 매크로가 실행 시작 때 적는 시각. Excel은 새로 고침이 끝난 뒤 쿼리를 백그라운드에서 한 번 더 평가하므로
//    (병렬 작업이 요청 수 계수 모의 서버로 실측), KIS를 부르는 쿼리는 자기 표의 최근 조회시각 ≥ started이면 이번 버튼 실행에서
//    이미 받은 것으로 보고 다시 부르지 않는다. 비어 있으면(하네스 시험 등) 이 확인을 하지 않는다.
//  - 반환: [Mode = text, Now = datetime, Today = date, Time = time, Override = logical, Started = nullable datetime]
//  - 사용 예) Ctl = fnPxRunCtl(), Live = Ctl[Mode] = "quick" or Ctl[Mode] = "full"
() as record =>
let
    Ctl = try Excel.CurrentWorkbook(){[Name = "tblRunCtl"]}[Content] otherwise null,
    Rows = if Ctl = null or not List.ContainsAll(Table.ColumnNames(Ctl), {"키", "값"}) then {}
           else List.Buffer(Table.ToRecords(Table.SelectColumns(Ctl, {"키", "값"}))),
    Get = (k as text) as any =>
        let m = List.Select(Rows, each (try Text.Trim(Text.From([키])) otherwise "") = k)
        in if List.IsEmpty(m) then null else m{0}[값],
    ModeRaw = Get("mode"),
    ModeIn = if ModeRaw = null then "" else Text.Lower(Text.Trim(Text.From(ModeRaw))),
    Mode = if List.Contains({"quick", "full", "sector", "analysis", "build"}, ModeIn) then ModeIn else "build",
    OvRaw = Get("now_override"),
    // 셀 값이 일시면 그대로, 일련번호(서식이 풀린 경우)·글자면 변환, 빈칸·변환 불가는 '없음'
    Ov = if OvRaw = null or (OvRaw is text and Text.Trim(OvRaw) = "") then null
         else try DateTime.From(OvRaw) otherwise null,
    Now = if Ov = null then DateTime.LocalNow() else Ov,
    StRaw = Get("started"),
    Started = if StRaw = null or (StRaw is text and Text.Trim(StRaw) = "") then null
              else try DateTime.From(StRaw) otherwise null
in
    [Mode = Mode, Now = Now, Today = Date.From(Now), Time = Time.From(Now), Override = Ov <> null, Started = Started]
