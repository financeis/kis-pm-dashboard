'==================================================================================================
' mod_refresh : 대회종목·업종·종목분석 페이지의 버튼 매크로 (KIS PM 대시보드, excel_dashboard)
'
' 원본: excel_dashboard/vba/mod_refresh.bas (UTF-8, Attribute 줄 없음). 빌더가 이 텍스트를 새 표준 모듈에
'   CodeModule.AddFromString으로 넣는다(권장 모듈 이름 mod_refresh). .bas 파일 가져오기는 이 PC에서
'   CP949로 읽혀 한글이 깨지므로 쓰지 않는다. 이 텍스트에 Option Explicit가 있으므로, 새 모듈에 VBE가
'   자동으로 넣은 줄(Option Explicit 등)은 먼저 지운 뒤 넣어야 중복 오류가 나지 않는다.
'   모듈 텍스트는 코드 페이지 949로 표현되는 문자만 쓴다(VBE가 949로 저장해 다른 문자는 "?"가 됨).
'   949에 없는 줄표(U+2014)는 실행 중에 ChrW로 만든다(말줄임표도 ChrW로 만든다).
'
' 매크로 (버튼에 연결하는 이름, 고정)
'   RefreshQuick  [시세]  mode=quick    : T_Token > T_Sessions > T_PxStore > T_PxMetrics > T_Company > T_ThemeAgg
'   RefreshFull   [전체]  mode=full     : T_Token > T_Class > T_Sessions > T_PxStore > T_PxMetrics > T_FlowU >
'                                         T_Target > T_Est > T_EstSnap > T_Events > T_Fin > T_CSL > T_Company >
'                                         T_ThemeAgg > T_SectorKRX > 이력 CSV 저장
'   RefreshSector [업종]  mode=sector   : T_Token > T_SectorKRX > T_ThemeAgg
'   RunAnalysis   [조회]  mode=analysis : T_Token > T_A_* (이름이 tblA_로 시작하는 표 전부, 이름이 Seed로 끝나는 정적 표 제외)
'   쿼리 T_X는 표 tblX에 적재된다(T_A_X는 tblA_X). 매크로는 표의 QueryTable을 새로 고친다.
'
' 실행 순서
'   0) 예열: 통합문서를 연 뒤 표마다 첫 새로 고침에서는 Excel이 쿼리를 한 번 더 평가해 KIS 호출이 겹칠 수 있다(실측).
'      그래서 이 Excel 세션에서 아직 한 번도 새로 고치지 않은 이 버튼의 표(T_Token과 CSV 단계 제외)를 먼저 mode=build로 두고
'      한 번씩 동기 새로 고침한다 - build 모드라 KIS 호출이 없다(2026-10-01). 새로 고친 표는 모듈 변수 m_refreshed(표 이름
'      사전)에 남겨 표마다 한 번만 예열한다 - 예: [시세]나 [업종]을 먼저 누른 뒤 [전체]를 누르면 [전체]에만 있는 표들을 예열한다
'      (세션 첫 버튼 하나만 예열하던 때는 그 뒤 다른 버튼의 표가 첫 새로 고침에서 호출을 겹쳐 냈다 - 표별 기록으로 고침).
'      예열은 [전체](mode=full)에서만 한다(WARMUP_MODES). [시세]·[업종]·[조회]는 예열 비용이 막는 호출보다 크다(2026-10-01 실측,
'      복사본·새 Excel 세션의 첫 실행: [시세] 예열 있음 79초 / 없음 54초(두 번째 실행 37~40초), [업종] 40초 / 30초(20초);
'      예열 없이 겹칠 수 있는 호출은 [시세] 약 20회(멀티 시세 18 + 세션 달력), [업종] 약 47회이고 걸리는 시간 차이는 2초 안쪽).
'      그 버튼들이 새로 고친 표도 m_refreshed에 남으므로 뒤이은 [전체]는 그 표를 다시 예열하지 않는다.
'      예열 중 실패는 무시한다(본 단계에서 다시 판정). Esc로 멈추면 다음 실행에서 남은 표를 다시 예열한다.
'   1) tblRunCtl에 mode(위 값)와 started(시작 시각)를 쓴다. 예열을 했으면 그 전에 두 가지를 기다린다.
'      ① 마지막 예열 새로 고침 뒤 WARMUP_SETTLE_SECONDS(15)초 + 예열에 걸린 시간(최대 WARMUP_SETTLE_MAX_SECONDS 120초)
'         (SettleAfterWarmUp, 상태 표시줄 "준비 중... n초"): Excel은 새로 고친 쿼리들을 몇 초 뒤 백그라운드에서 한 번 더
'         평가하는데(Excel이 쉴 때 시작해 예열한 표를 하나씩 다시 평가하는 묶음), 그때 이미 mode=full과 새 started가 적혀
'         있으면 그 평가가 조회를 다시 해 KIS 호출이 겹친다(예: 주간 조사일의 신용·공매도·대차 약 1,500회). 묶음이 다시
'         평가하는 쿼리는 예열에서 방금 평가한 것들이라 걸리는 시간도 예열과 비슷하므로 기다림을 예열 시간에 맞춰 늘린다.
'         2026-10-02 요청 수 계수 모의 서버 실측(새 Excel 세션의 첫 [전체], 7종목): 기다림 없음 - 본 단계가 길면 묶음이
'         본 단계 중에 돌아 중복(표 14개: 47·128건), 15초 고정 - 묶음이 15초보다 길어 전환 뒤까지 이어져 중복(표 9개:
'         81건, 6/6회), 30·60초 고정 0건(7/7회), 이 규칙(15초 + 예열 시간, 실제 34·50초) 0건(4/4회).
'         예열에서 새로 고친 쿼리 표가 없으면(이 세션에서 이미 예열했거나 쿼리에 연결되지 않은 표뿐) 기다리지 않는다.
'      ② 시계가 다음 초로 넘어갈 때까지(WaitNextSecond, 1초 이내): 쿼리의 재호출 방지 확인은 '자기 표 조회시각 >= started'인데
'         Now는 초 단위로 잘리므로, 예열이 같은 초 안에 남긴 조회시각이 started 이상이 되어 그 표의 본 단계가 계산 없이
'         예열 결과를 돌려줄 수 있었다(2026-10-01 실측: 새 세션 첫 [시세]에서 테마 집계가 대회종목 표보다 먼저 계산된 값으로 남음).
'      둘 중 하나라도 쓰지 못하면 단계를 하나도 실행하지 않고
'      멈춘다(요약에 사유). mode를 못 쓰면 쿼리가 build로 돌아 조회가 없고, started를 못 쓰면 지난 실행의 started가 남아
'      쿼리의 재호출 방지 확인(자기 표 조회시각 >= started)이 이번 조회까지 막을 수 있기 때문.
'   2) 단계마다 표의 QueryTable을 동기 새로 고침한다(BackgroundQuery를 잠시 끄고 Refresh False, 끝나면 원래
'      값으로 되돌림. 그 표가 백그라운드 새로 고침 중이면 끝날 때까지 최대 120초 기다림). 상태 표시줄에
'      "[전체] 4/16 T_PxStore 새로 고침 중..." 형태로 진행을 표시한다. 한 단계가 실패해도 멈추지 않고 다음
'      단계로 간다. 표가 없거나 쿼리에 연결되지 않은 표여도 그 단계만 실패로 판정한다.
'   3) 단계 판정 (docs/business-rules.md의 '버튼 결과 판정')
'      실패      : 새로 고침 오류, 표 없음, 상태 열 없음, 또는 상태 중 하나라도 "이전 데이터"로 시작
'      일부 오류 : 상태가 "오류:"로 시작하는 행이 있음. 종목코드 열이 있으면 서로 다른 종목 수(한 종목이 여러
'                  행이어도 1건), 없으면 행 수. 단 tblSessions(세션 달력)는 달력 하나가 단위라서 오류 행이
'                  몇 개든 1건(안전 규칙의 "오류: 오늘 세션 미확인" = 일부 오류 1건)
'      정상      : 그 밖(OK, 데이터 없음, 추정 없음 등)
'      T_Token   : 새로 고침 뒤 tblToken에 env=prod(다른 쿼리가 읽는 토큰)이고 token이 비어 있지 않으며
'                  expires까지 5분 이상 남은 행이 없으면 실패. 이 표의 상태 열은 status이며 판정에는 쓰지
'                  않는다. 토큰 값 자체는 어디에도 쓰거나 보여 주지 않는다.
'   4) [전체]만: 마지막 단계로 tblEstSnap을 통합문서 폴더의 history\est_snap.csv로 저장한다
'      (UTF-8 BOM, 줄 끝 CRLF, 머리글 = 표 열 이름, 날짜 yyyy-mm-dd, 시각이 있는 열은 yyyy-mm-dd hh:nn:ss,
'      숫자는 로캘과 무관하게 소수점 ".", 쉼표·큰따옴표·줄바꿈이 든 글자는 큰따옴표로 감쌈). 이 단계도 따로
'      판정한다. 표가 비었는데 CSV가 이미 있으면 이력을 지우지 않도록 덮어쓰지 않고 실패로 판정한다.
'      스냅샷 표는 행을 지우지 않으므로(쌓기만 함) 표의 행 수가 기존 CSV의 데이터 행 수보다 적거나 기존 CSV의 행 수를 읽지
'      못하면 기존 파일을 그대로 두고 새 내용을 history\est_snap_yyyymmdd_hhnnss.csv로 따로 저장한 뒤 그 단계를 일부 오류
'      1건("오류: ..." 사유)으로 판정한다(실행은 계속). 기존 CSV 행 수는 큰따옴표 안의 줄바꿈을 빼고 센다.
'      이어서 tblSettings의 force_weekly를 N으로 되돌린다(사용자가 Esc로 중단한 실행은 되돌리지 않음).
'   5) 이름 칸 ..._최근조회에 종료 시각, ..._상태에 요약을 쓴다. 실패한 단계가 있으면
'      "실패: <단계들> (줄표) 이전 데이터 표시 중", 없고 일부 오류가 있으면 "일부 오류 n건"(건수 합), 그 밖은 "정상".
'   6) 정리(오류가 나도 항상 실행): mode를 build로 되돌리고 상태 표시줄·화면 갱신·이벤트·경고·커서·취소 키
'      설정을 원래대로 돌린다. 단계별 요약 문장을 tblRunCtl의 last_summary에 쓰고(행이 없으면 추가),
'      Application.Visible = True이고 tblRunCtl의 quiet가 Y가 아닐 때만 요약 MsgBox를 띄운다
'      (COM으로 숨김 Excel에서 호출할 때 모달 창이 자동화를 멈추지 않도록). VBA MsgBox는 949로 바꿔
'      보여 주므로 대화상자 안에서만 줄표를 가로줄(U+2015)로 바꾼다. 셀에는 원래 줄표를 쓴다.
'      상태 표시줄은 StatusBar = False로 Excel에 돌려준다. 이 Excel은 그 뒤에도 StatusBar를 읽으면 글자
'      "FALSE"를 돌려주므로(화면은 기본 "준비" 표시), 저장해 둔 값이 "FALSE"면 그 글자를 다시 넣지 않는다.
'   실행 중 Esc를 누르면 진행 중인 단계 뒤로 멈추고, 남은 단계는 "실행 안 함"(실패)으로 기록한다.
'
' tblRunCtl (정적 표, 열: 키, 값) - 빌더가 만들 키
'   mode         : build(평소) / quick / full / sector / analysis. 매크로가 쓰고 끝나면 build로 되돌린다.
'   started      : 버튼 실행 시작 시각(일시 값). 매크로가 쓴다.
'   now_override : 시험 전용(쿼리가 읽음). 매크로는 읽지도 쓰지도 않는다.
'   quiet        : Y면 요약 MsgBox를 띄우지 않는다(자동화·시험용, 평소 비워 둠). 매크로는 읽기만 한다.
'   last_summary : 마지막 버튼 실행의 단계별 요약 문장. 매크로가 쓴다.
'   값 열은 일반 서식으로 둔다(텍스트 서식이면 started가 글자로 저장된다). 매크로는 없는 키 행을 추가하지만
'   (ListRows.Add), 빌더가 다섯 키를 모두 만들어 두면 행 추가가 일어나지 않는다. last_summary는 여러 줄 글자다.
'
' 이름 정의: 시세_최근조회·시세_상태, 전체_최근조회·전체_상태, 업종_최근조회·업종_상태, 분석_최근조회·분석_상태
'   (쓰기), 분석코드(읽기: 요약 제목에 종목코드 표시). 통합문서 범위 이름을 우선 찾고, 없으면 시트 범위
'   이름도 찾는다. 이름이 없으면 오류 없이 요약에 경고 줄만 남긴다.
' 그 밖의 표: tblSettings(키, 값: force_weekly, cfg_path, vs_path), tblToken(env, token, expires, status), tblEstSnap.
'
' 다른 PC로 옮긴 통합문서: Auto_Open(사용자가 통합문서를 열 때 Excel이 부름 - COM 자동화로 열 때는 불리지 않음)과
'   모든 버튼의 시작에서 FixLocalPaths가 [설정] cfg_path·vs_path를 이 PC에 맞춘다. 지금 경로에 파일이 없고 표준 위치
'   (cfg_path: %USERPROFILE%\KIS\config\kis_devlp.yaml, vs_path: 통합문서 폴더 또는 그 상위 폴더의 수집기업_valuesearch.xlsx)에
'   파일이 있을 때만 바꾼다. 버튼에서 바꾸면 요약에 [주의] 줄로 알린다.
'
' 불변식: 비동기 쿼리가 모두 끝날 때까지 기다리는 Application 메서드(CalculateUntil...)는 쓰지 않는다
'   (통합문서 표를 읽는 쿼리와 교착됨). 표별 동기 새로 고침과 QueryTable.Refreshing 확인만 쓴다.
' 참조 추가 없음: ADODB.Stream, Scripting.FileSystemObject, Scripting.Dictionary는 CreateObject로 쓴다.
' 시험용 공개 함수: JudgeTable(표 이름), SaveEstSnapCsv() - "판정|건수|사유" 문자열을 돌려준다.
'   RefreshedTables() - 이 Excel 세션에서 새로 고친 표 이름(소문자, 쉼표로 이음) - 예열 동작 확인용.
'   FixLocalPaths() - 바꾼 설정 경로("키 -> 경로" 줄들, 없으면 빈 문자열).
'==================================================================================================
Option Explicit

Private Const V_OK As Long = 0
Private Const V_PARTIAL As Long = 1
Private Const V_FAIL As Long = 2

Private Const STEP_CSV As String = "이력 CSV 저장"
Private Const STEP_ANALYSIS As String = "T_A_*"
Private Const NOT_RUN As String = "실행 안 함"
Private Const PREFIX_PREV As String = "이전 데이터"
Private Const PREFIX_ERR As String = "오류:"
Private Const HISTORY_DIR As String = "history"
Private Const SNAP_FILE As String = "est_snap.csv"
' 기존 CSV를 덮지 않을 때 새 내용을 따로 저장하는 파일 이름 앞부분(뒤에 yyyymmdd_hhnnss.csv) - 머리 주석 4번
Private Const SNAP_ALT_PREFIX As String = "est_snap_"
' 다른 PC 경로 맞추기(FixLocalPaths)의 표준 위치
Private Const CFG_REL_PATH As String = "\KIS\config\kis_devlp.yaml"
Private Const VS_FILE As String = "수집기업_valuesearch.xlsx"
Private Const TOKEN_MIN_MINUTES As Long = 5
Private Const WAIT_BUSY_SECONDS As Long = 120
Private Const REASON_MAX As Long = 200
Private Const REASON_MAX_BOX As Long = 60

' 예열하는 실행 모드(머리 주석 0번 - 실측으로 [전체]만). 형식: "|모드|모드|"
Private Const WARMUP_MODES As String = "|full|"
' 마지막 예열 새로 고침 뒤 mode·started를 쓰기 전에 기다리는 시간(머리 주석 1번 ①): 이 초 + 예열에 걸린 초, 최대 MAX 초
Private Const WARMUP_SETTLE_SECONDS As Long = 15
Private Const WARMUP_SETTLE_MAX_SECONDS As Long = 120

' 이 Excel 세션에서 한 번이라도 새로 고친 표(예열과 본 단계 모두): 표 이름(소문자) -> True (Scripting.Dictionary).
' 아직 새로 고치지 않은 표만 예열한다(머리 주석 0번). VBA 프로젝트가 초기화되면(통합문서를 다시 열면) 비워진다.
Private m_refreshed As Object

' 한 번의 버튼 실행 상태 (재진입은 m_running으로 막는다)
Private m_running As Boolean
Private m_cancelled As Boolean
Private m_n As Long
Private m_names() As String
Private m_tables() As String
Private m_verdict() As Long
Private m_count() As Long
Private m_reason() As String
Private m_notes As String

'--------------------------------------------------------------------------------------------------
' 버튼 매크로
'--------------------------------------------------------------------------------------------------
Public Sub RefreshQuick()
    RunButton "quick", "[시세]", "시세_최근조회", "시세_상태", _
        Array("T_Token", "T_Sessions", "T_PxStore", "T_PxMetrics", "T_Company", "T_ThemeAgg")
End Sub

Public Sub RefreshFull()
    RunButton "full", "[전체]", "전체_최근조회", "전체_상태", _
        Array("T_Token", "T_Class", "T_Sessions", "T_PxStore", "T_PxMetrics", "T_FlowU", "T_Target", _
              "T_Est", "T_EstSnap", "T_Events", "T_Fin", "T_CSL", "T_Company", "T_ThemeAgg", "T_SectorKRX", _
              STEP_CSV)
End Sub

Public Sub RefreshSector()
    RunButton "sector", "[업종]", "업종_최근조회", "업종_상태", Array("T_Token", "T_SectorKRX", "T_ThemeAgg")
End Sub

Public Sub RunAnalysis()
    Dim code As String
    code = Clip(CleanText(CellText(NamedValue("분석코드"))), 20)
    If Len(code) > 0 Then
        RunButton "analysis", "[조회] " & code, "분석_최근조회", "분석_상태", Array("T_Token", STEP_ANALYSIS)
    Else
        RunButton "analysis", "[조회]", "분석_최근조회", "분석_상태", Array("T_Token", STEP_ANALYSIS)
    End If
End Sub

' 사용자가 통합문서를 열 때(매크로를 허용한 뒤) Excel이 부른다. COM 자동화로 열 때는 불리지 않는다.
Public Sub Auto_Open()
    On Error Resume Next
    FixLocalPaths
End Sub

'--------------------------------------------------------------------------------------------------
' 시험·점검용 공개 함수 ("판정|건수|사유")
'--------------------------------------------------------------------------------------------------
Public Function JudgeTable(ByVal tableName As String) As String
    Dim lo As ListObject
    Dim verdict As Long
    Dim cnt As Long
    Dim reason As String
    On Error GoTo Fail
    Set lo = FindTable(tableName)
    If lo Is Nothing Then
        verdict = V_FAIL
        reason = "표 없음: " & tableName
    Else
        JudgeListObject lo, verdict, cnt, reason
    End If
    JudgeTable = VerdictName(verdict) & "|" & cnt & "|" & reason
    Exit Function
Fail:
    JudgeTable = VerdictName(V_FAIL) & "|0|" & CleanText(Err.Description)
End Function

Public Function SaveEstSnapCsv() As String
    Dim verdict As Long
    Dim cnt As Long
    Dim reason As String
    SaveEstSnapCsvCore verdict, cnt, reason
    SaveEstSnapCsv = VerdictName(verdict) & "|" & cnt & "|" & reason
End Function

' [설정]의 cfg_path·vs_path가 이 PC에 없는 파일을 가리키면, 표준 위치에 파일이 있을 때만 그 경로로 바꾼다
' (다른 PC에서 만들거나 옮겨 온 통합문서). 지금 경로에 파일이 있거나 후보에도 없으면 그대로 둔다.
' 반환: 바꾼 키와 새 경로(줄마다 "키 -> 경로"), 바꾼 것이 없으면 빈 문자열. 오류는 삼킨다(버튼 실행을 막지 않음).
Public Function FixLocalPaths() As String
    Dim fso As Object
    Dim msg As String
    Dim home As String
    Dim wbDir As String
    Dim upDir As String
    Dim cfgCands As Variant
    Dim vsCands As Variant
    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    home = Environ$("USERPROFILE")
    wbDir = ThisWorkbook.Path
    If Len(wbDir) > 0 Then upDir = fso.GetParentFolderName(wbDir)
    If Len(home) > 0 Then cfgCands = Array(home & CFG_REL_PATH) Else cfgCands = Array()
    If Len(wbDir) = 0 Then
        vsCands = Array()
    ElseIf Len(upDir) = 0 Then
        vsCands = Array(fso.BuildPath(wbDir, VS_FILE))
    Else
        vsCands = Array(fso.BuildPath(wbDir, VS_FILE), fso.BuildPath(upDir, VS_FILE))
    End If
    msg = FixOnePath(fso, "cfg_path", cfgCands)
    msg = msg & FixOnePath(fso, "vs_path", vsCands)
Done:
    FixLocalPaths = msg
End Function

' FixLocalPaths 도우미: 키 하나. 바꿨으면 "키 -> 경로" & 줄바꿈, 아니면 빈 문자열.
Private Function FixOnePath(ByVal fso As Object, ByVal key As String, ByVal candidates As Variant) As String
    Dim cur As String
    Dim c As Variant
    Dim why As String
    On Error GoTo Done
    cur = Trim$(CellText(GetKeyValue("tblSettings", key)))
    If Len(cur) > 0 Then
        If fso.FileExists(cur) Then Exit Function
    End If
    For Each c In candidates
        If Len(CStr(c)) > 0 Then
            If fso.FileExists(CStr(c)) Then
                If SetKeyValue("tblSettings", key, CStr(c), False, why) Then
                    FixOnePath = key & " -> " & CStr(c) & vbLf
                End If
                Exit Function
            End If
        End If
    Next c
Done:
End Function

Public Function RefreshedTables() As String
    If m_refreshed Is Nothing Then Exit Function
    If m_refreshed.Count = 0 Then Exit Function
    ' keys: 이 모듈의 변수 keys(JudgeStatus)와 같은 대소문자로 써야 VBE가 바꾸지 않는다(빌더의 넣은 코드 대조).
    RefreshedTables = Join(m_refreshed.keys, ",")
End Function

'--------------------------------------------------------------------------------------------------
' 실행기
'--------------------------------------------------------------------------------------------------
Private Sub RunButton(ByVal modeName As String, ByVal title As String, ByVal timeName As String, _
                      ByVal statusName As String, ByVal steps As Variant)
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean
    Dim prevAlerts As Boolean
    Dim prevCursor As Long
    Dim prevCancel As Long
    Dim prevStatus As Variant
    Dim stateSaved As Boolean
    Dim t0 As Double
    Dim warmT0 As Double
    Dim secs As Double
    Dim endTime As Date
    Dim i As Long
    Dim nSteps As Long
    Dim why As String
    Dim overall As String
    Dim fullText As String
    Dim fixed As String

    If m_running Then
        Beep
        Exit Sub
    End If
    m_running = True
    m_cancelled = False
    m_notes = ""
    m_n = 0
    Erase m_names, m_tables, m_verdict, m_count, m_reason
    t0 = Timer
    On Error GoTo Crash

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevCursor = Application.Cursor
    prevCancel = Application.EnableCancelKey
    prevStatus = Application.StatusBar
    stateSaved = True
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    ' 다른 PC로 옮긴 통합문서면 설정 경로를 이 PC에 맞춘다(머리 주석 '다른 PC로 옮긴 통합문서')
    fixed = FixLocalPaths()
    If Right$(fixed, 1) = vbLf Then fixed = Left$(fixed, Len(fixed) - 1)
    If Len(fixed) > 0 Then AddNote "이 PC에 맞게 설정 경로를 바꿈: " & Replace(fixed, vbLf, " / ")
    Application.Cursor = xlWait
    Application.StatusBar = title & " 준비 중" & Ellipsis()

    BuildSteps steps
    nSteps = m_n

    If NeedsWarmUp(modeName, nSteps) Then
        If Not SetKeyValue("tblRunCtl", "mode", "build", True, why) Then
            AbortRun "mode를 build로 쓰지 못함: " & why
            GoTo Finalize
        End If
        warmT0 = Timer
        If WarmUp(title, modeName, nSteps) > 0 Then
            If m_cancelled Then GoTo Finalize
            SettleAfterWarmUp title, ElapsedSince(warmT0)
        End If
        If m_cancelled Then GoTo Finalize
        WaitNextSecond
    End If

    If Not SetKeyValue("tblRunCtl", "mode", modeName, True, why) Then
        AbortRun "mode를 쓰지 못함: " & why
        GoTo Finalize
    End If
    If Not SetKeyValue("tblRunCtl", "started", Now, True, why) Then
        AbortRun "started를 쓰지 못함: " & why
        GoTo Finalize
    End If

    For i = 1 To nSteps
        If m_cancelled Then Exit For
        If m_names(i) = STEP_CSV Then
            Application.StatusBar = title & " " & i & "/" & nSteps & " " & m_names(i) & " 중" & Ellipsis()
        Else
            Application.StatusBar = title & " " & i & "/" & nSteps & " " & m_names(i) & " 새로 고침 중" & Ellipsis()
        End If
        DoEvents
        RunStepSafe i
    Next i

Finalize:
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    If Not SetKeyValue("tblRunCtl", "mode", "build", True, why) Then AddNote "mode를 build로 되돌리지 못함: " & why
    If m_n = 0 Then AddResult "실행", "", V_FAIL, 0, "단계를 시작하지 못함"
    If modeName = "full" Then
        If m_cancelled Then
            AddNote "사용자 중단: force_weekly는 그대로 둠"
        ElseIf Not SetKeyValue("tblSettings", "force_weekly", "N", False, why) Then
            AddNote "force_weekly를 N으로 되돌리지 못함: " & why
        End If
    End If
    endTime = Now
    secs = ElapsedSince(t0)
    overall = OverallStatus()
    If Not WriteNamed(timeName, endTime, why) Then AddNote why
    If Not WriteNamed(statusName, overall, why) Then AddNote why
    fullText = SummaryText(title, endTime, secs, overall, REASON_MAX)
    If Not SetKeyValue("tblRunCtl", "last_summary", fullText, True, why) Then AddNote "last_summary 기록 실패: " & why
    If stateSaved Then
        ' 이 Excel은 StatusBar = False 뒤에도 읽으면 글자 "FALSE"를 돌려준다(화면은 기본 표시로 돌아감).
        ' 그 값을 다시 넣으면 "FALSE" 글자가 표시되므로, 다른 매크로가 남긴 글자일 때만 되돌리고 그 밖은 False.
        If VarType(prevStatus) = vbString Then
            If StrComp(CStr(prevStatus), "FALSE", vbTextCompare) <> 0 And Len(CStr(prevStatus)) > 0 Then
                Application.StatusBar = CStr(prevStatus)
            Else
                Application.StatusBar = False
            End If
        Else
            Application.StatusBar = False
        End If
        Application.ScreenUpdating = prevScreen
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.Cursor = prevCursor
        Application.EnableCancelKey = prevCancel
    Else
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Cursor = xlDefault
        Application.EnableCancelKey = xlInterrupt
    End If
    m_running = False
    ShowSummary SummaryText(title, endTime, secs, overall, REASON_MAX_BOX), overall, title
    Exit Sub

Crash:
    If Err.Number = 18 Then
        m_cancelled = True
    Else
        AddResult "매크로 오류", "", V_FAIL, 0, CleanText(Err.Description)
    End If
    Resume Finalize
End Sub

' 실행 제어 값을 쓰지 못해 버튼 실행을 멈춘다. 단계들은 BuildSteps가 채운 '실행 안 함'(실패)으로 남고,
' 정리 단계(Finalize)가 mode를 build로 되돌리고 상태 칸·요약에 이 사유를 남긴다.
Private Sub AbortRun(ByVal reason As String)
    AddResult "실행 제어(tblRunCtl)", "", V_FAIL, 0, _
        Clip(CleanText(reason), REASON_MAX) & " - 실행을 멈춤(KIS 호출 없음)"
End Sub

' 예열: 이 버튼의 표 중 이 세션에서 아직 새로 고치지 않은 표를 build 모드(KIS 호출 없음)로 한 번씩 새로 고친다(머리 주석 0번).
' tblToken(T_Token은 모드와 상관없이 토큰을 확인)과 CSV 단계(표 없음)는 건너뛴다. 결과는 판정에 쓰지 않는다.
' 새로 고친 표는 RefreshTable이 m_refreshed에 남긴다. 새로 고침을 시도한 쿼리 표의 수를 돌려준다(쿼리에 연결되지 않은 표는
' 평가되는 쿼리가 없으므로 세지 않음. 0이면 뒤의 기다림을 건너뜀).
Private Function WarmUp(ByVal title As String, ByVal modeName As String, ByVal nSteps As Long) As Long
    Dim i As Long
    Dim lo As ListObject
    Dim why As String
    Dim tried As Long
    For i = 1 To nSteps
        If m_cancelled Then Exit For
        If WantsWarmUp(modeName, m_tables(i)) Then
            Set lo = FindTable(m_tables(i))
            If Not lo Is Nothing Then
                Application.StatusBar = title & " 첫 실행 준비(호출 없음) " & i & "/" & nSteps & " " & m_names(i) & Ellipsis()
                DoEvents
                RefreshTable lo, why
                If lo.SourceType <> xlSrcRange Then tried = tried + 1
            End If
        End If
    Next i
    WarmUp = tried
End Function

' 마지막 예열 새로 고침 뒤 기다린다(머리 주석 1번 ①): WARMUP_SETTLE_SECONDS + 예열에 걸린 초(warmSecs), 최대
' WARMUP_SETTLE_MAX_SECONDS. 그동안 mode는 build이고, DoEvents로 Excel이 예열한 쿼리들의 백그라운드 재평가를 처리하게 한다
' (그 평가는 build 모드라 KIS를 부르지 않음). 상태 표시줄에 남은 초를 보여 준다. Esc는 다른 단계처럼 실행을 멈춘다(오류 18 ->
' RunButton의 정리 단계).
Private Sub SettleAfterWarmUp(ByVal title As String, ByVal warmSecs As Double)
    Dim t0 As Double
    Dim settleSecs As Long
    Dim secsLeft As Long
    Dim lastShown As Long
    settleSecs = WARMUP_SETTLE_SECONDS + CLng(warmSecs)
    If settleSecs > WARMUP_SETTLE_MAX_SECONDS Then settleSecs = WARMUP_SETTLE_MAX_SECONDS
    t0 = Timer
    lastShown = -1
    Do
        secsLeft = settleSecs - Int(ElapsedSince(t0))
        If secsLeft <= 0 Then Exit Do
        If secsLeft <> lastShown Then
            Application.StatusBar = title & " 준비 중" & Ellipsis() & " " & secsLeft & "초"
            lastShown = secsLeft
        End If
        DoEvents
        If m_cancelled Then Exit Do
    Loop
End Sub

' 예열할 표인지: 예열하는 모드(WARMUP_MODES)이고, 표 단계이며 tblToken이 아니고, 이 세션에서 아직 새로 고치지 않은 표.
Private Function WantsWarmUp(ByVal modeName As String, ByVal tableName As String) As Boolean
    If InStr(1, WARMUP_MODES, "|" & modeName & "|", vbTextCompare) = 0 Then Exit Function
    If Len(tableName) = 0 Then Exit Function
    If StrComp(tableName, "tblToken", vbTextCompare) = 0 Then Exit Function
    WantsWarmUp = Not WasRefreshed(tableName)
End Function

' 이 버튼 실행에 예열할 표가 하나라도 있는지(없으면 mode를 build로 쓰는 일도 건너뜀).
Private Function NeedsWarmUp(ByVal modeName As String, ByVal nSteps As Long) As Boolean
    Dim i As Long
    For i = 1 To nSteps
        If WantsWarmUp(modeName, m_tables(i)) Then
            If Not FindTable(m_tables(i)) Is Nothing Then
                NeedsWarmUp = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function WasRefreshed(ByVal tableName As String) As Boolean
    If m_refreshed Is Nothing Then Exit Function
    WasRefreshed = m_refreshed.Exists(LCase$(tableName))
End Function

Private Sub MarkRefreshed(ByVal tableName As String)
    If m_refreshed Is Nothing Then Set m_refreshed = CreateObject("Scripting.Dictionary")
    m_refreshed(LCase$(tableName)) = True
End Sub

' 단계 목록을 결과 배열에 '실행 안 함'(실패)으로 먼저 채운다. 실행된 단계만 결과를 덮어쓴다.
Private Sub BuildSteps(ByVal steps As Variant)
    Dim i As Long
    Dim s As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim found As Long
    For i = LBound(steps) To UBound(steps)
        s = CStr(steps(i))
        If s = STEP_ANALYSIS Then
            found = 0
            For Each ws In ThisWorkbook.Worksheets
                For Each lo In ws.ListObjects
                    If IsAnalysisTable(lo.Name) Then
                        AddResult "T_" & Mid$(lo.Name, 4), lo.Name, V_FAIL, 0, NOT_RUN
                        found = found + 1
                    End If
                Next lo
            Next ws
            If found = 0 Then AddResult STEP_ANALYSIS, "", V_FAIL, 0, "표 없음: 이름이 tblA_로 시작하는 표가 없음"
        ElseIf s = STEP_CSV Then
            AddResult s, "", V_FAIL, 0, NOT_RUN
        Else
            AddResult s, "tbl" & Mid$(s, 3), V_FAIL, 0, NOT_RUN
        End If
    Next i
End Sub

Private Function IsAnalysisTable(ByVal tableName As String) As Boolean
    If StrComp(Left$(tableName, 5), "tblA_", vbTextCompare) <> 0 Then Exit Function
    If StrComp(Right$(tableName, 4), "Seed", vbTextCompare) = 0 Then Exit Function
    IsAnalysisTable = True
End Function

' 단계 하나를 실행한다. 어떤 오류도 이 단계의 실패로 바꾸고 호출자에게 넘기지 않는다.
Private Sub RunStepSafe(ByVal i As Long)
    On Error GoTo Fail
    If m_names(i) = STEP_CSV Then
        SaveEstSnapCsvCore m_verdict(i), m_count(i), m_reason(i)
    ElseIf Len(m_tables(i)) > 0 Then
        RunQueryStep m_tables(i), m_verdict(i), m_count(i), m_reason(i)
    End If
    Exit Sub
Fail:
    If Err.Number = 18 Then m_cancelled = True
    m_verdict(i) = V_FAIL
    m_count(i) = 0
    m_reason(i) = "예기치 않은 오류: " & CleanText(Err.Description)
End Sub

Private Sub RunQueryStep(ByVal tableName As String, ByRef verdict As Long, ByRef cnt As Long, ByRef reason As String)
    Dim lo As ListObject
    Dim why As String
    verdict = V_FAIL
    cnt = 0
    reason = ""
    Set lo = FindTable(tableName)
    If lo Is Nothing Then
        reason = "표 없음: " & tableName
        Exit Sub
    End If
    If Not RefreshTable(lo, why) Then
        reason = "새로 고침 오류: " & why
        Exit Sub
    End If
    JudgeListObject lo, verdict, cnt, reason
End Sub

' 표의 QueryTable을 동기로 새로 고친다. 성공하면 True, 실패하면 False와 사유.
Private Function RefreshTable(ByVal lo As ListObject, ByRef why As String) As Boolean
    Dim qt As QueryTable
    Dim prevBg As Boolean
    Dim haveBg As Boolean
    On Error GoTo Fail
    If lo.SourceType = xlSrcRange Then
        why = "쿼리에 연결되지 않은 표(" & lo.Name & ")"
        GoTo Done
    End If
    Set qt = lo.QueryTable
    If Not WaitNotRefreshing(qt) Then
        why = "백그라운드 새로 고침이 " & WAIT_BUSY_SECONDS & "초 안에 끝나지 않음"
        GoTo Done
    End If
    prevBg = qt.BackgroundQuery
    haveBg = True
    qt.BackgroundQuery = False
    qt.Refresh False
    RefreshTable = True
    MarkRefreshed lo.Name
Done:
    On Error Resume Next
    If haveBg Then qt.BackgroundQuery = prevBg
    Exit Function
Fail:
    If Err.Number = 18 Then m_cancelled = True
    why = CleanText(Err.Description)
    If Len(why) = 0 Then why = "오류 번호 " & Err.Number
    RefreshTable = False
    Resume Done
End Function

Private Function WaitNotRefreshing(ByVal qt As QueryTable) As Boolean
    Dim t0 As Double
    t0 = Timer
    Do While qt.Refreshing
        If ElapsedSince(t0) > WAIT_BUSY_SECONDS Then Exit Function
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    WaitNotRefreshing = True
End Function

'--------------------------------------------------------------------------------------------------
' 판정
'--------------------------------------------------------------------------------------------------
Private Sub JudgeListObject(ByVal lo As ListObject, ByRef verdict As Long, ByRef cnt As Long, ByRef reason As String)
    If StrComp(lo.Name, "tblToken", vbTextCompare) = 0 Then
        JudgeToken lo, verdict, cnt, reason
    Else
        JudgeStatus lo, verdict, cnt, reason
    End If
End Sub

Private Sub JudgeStatus(ByVal lo As ListObject, ByRef verdict As Long, ByRef cnt As Long, ByRef reason As String)
    Dim sc As Long
    Dim cc As Long
    Dim n As Long
    Dim r As Long
    Dim st As Variant
    Dim codes As Variant
    Dim s As String
    Dim k As String
    Dim prevRows As Long
    Dim errRows As Long
    Dim firstPrev As String
    Dim firstErr As String
    Dim keys As Object

    verdict = V_FAIL
    cnt = 0
    reason = ""
    sc = ColIndex(lo, "상태")
    If sc = 0 Then
        reason = "상태 열 없음(" & lo.Name & ")"
        Exit Sub
    End If
    n = lo.ListRows.Count
    If n = 0 Then
        verdict = V_OK
        reason = "0행"
        Exit Sub
    End If
    st = ColumnValues(lo, sc)
    cc = ColIndex(lo, "종목코드")
    If cc > 0 Then codes = ColumnValues(lo, cc)
    Set keys = CreateObject("Scripting.Dictionary")
    For r = 1 To n
        s = LTrim$(CellText(st(r, 1)))
        If Left$(s, Len(PREFIX_PREV)) = PREFIX_PREV Then
            prevRows = prevRows + 1
            If Len(firstPrev) = 0 Then firstPrev = s
        ElseIf Left$(s, Len(PREFIX_ERR)) = PREFIX_ERR Then
            errRows = errRows + 1
            If Len(firstErr) = 0 Then firstErr = s
            If cc > 0 Then
                k = NormCode(codes(r, 1))
                If Not keys.Exists(k) Then keys.Add k, True
            End If
        End If
    Next r
    If prevRows > 0 Then
        verdict = V_FAIL
        reason = Clip(CleanText(firstPrev), REASON_MAX)
    ElseIf errRows > 0 Then
        verdict = V_PARTIAL
        If StrComp(lo.Name, "tblSessions", vbTextCompare) = 0 Then
            cnt = 1
        ElseIf cc > 0 Then
            cnt = keys.Count
        Else
            cnt = errRows
        End If
        reason = Clip(CleanText(firstErr), REASON_MAX)
    Else
        verdict = V_OK
        reason = n & "행"
    End If
End Sub

' tblToken: env=prod, token 있음, expires까지 TOKEN_MIN_MINUTES분 이상 남은 행이 있으면 정상. 토큰 값은 길이만 본다.
Private Sub JudgeToken(ByVal lo As ListObject, ByRef verdict As Long, ByRef cnt As Long, ByRef reason As String)
    Dim ec As Long
    Dim tc As Long
    Dim xc As Long
    Dim sc As Long
    Dim n As Long
    Dim r As Long
    Dim envs As Variant
    Dim toks As Variant
    Dim exps As Variant
    Dim sts As Variant
    Dim d As Date
    Dim mins As Double
    Dim best As Double
    Dim found As Boolean
    Dim lastStatus As String

    verdict = V_FAIL
    cnt = 0
    reason = ""
    tc = ColIndex(lo, "token")
    xc = ColIndex(lo, "expires")
    ec = ColIndex(lo, "env")
    sc = ColIndex(lo, "status")
    If tc = 0 Or xc = 0 Then
        reason = "토큰 표에 token/expires 열 없음"
        Exit Sub
    End If
    n = lo.ListRows.Count
    If n > 0 Then
        toks = ColumnValues(lo, tc)
        exps = ColumnValues(lo, xc)
        If ec > 0 Then envs = ColumnValues(lo, ec)
        If sc > 0 Then sts = ColumnValues(lo, sc)
    End If
    For r = 1 To n
        If ec > 0 Then
            If LCase$(Trim$(CellText(envs(r, 1)))) <> "prod" Then GoTo NextRow
        End If
        If sc > 0 Then lastStatus = CellText(sts(r, 1))
        If Len(Trim$(CellText(toks(r, 1)))) = 0 Then GoTo NextRow
        If Not TryDate(exps(r, 1), d) Then GoTo NextRow
        mins = (CDbl(d) - CDbl(Now)) * 1440#
        If mins >= TOKEN_MIN_MINUTES Then
            If Not found Or mins > best Then best = mins
            found = True
        End If
NextRow:
    Next r
    If found Then
        verdict = V_OK
        reason = "토큰 남은 시간 " & Int(best / 60) & "시간 " & (Int(best) Mod 60) & "분"
    Else
        reason = "유효한 토큰 없음(만료까지 " & TOKEN_MIN_MINUTES & "분 이상 남은 prod 토큰이 없음)"
        If Len(lastStatus) > 0 Then reason = reason & ": " & Clip(CleanText(lastStatus), 120)
    End If
End Sub

Private Function OverallStatus() As String
    Dim i As Long
    Dim failed As String
    Dim partialCount As Long
    Dim anyPartial As Boolean
    For i = 1 To m_n
        If m_verdict(i) = V_FAIL Then
            If Len(failed) > 0 Then failed = failed & ", "
            failed = failed & m_names(i)
        ElseIf m_verdict(i) = V_PARTIAL Then
            anyPartial = True
            partialCount = partialCount + m_count(i)
        End If
    Next i
    If Len(failed) > 0 Then
        OverallStatus = "실패: " & failed & " " & EmDash() & " 이전 데이터 표시 중"
    ElseIf anyPartial Then
        OverallStatus = "일부 오류 " & partialCount & "건"
    Else
        OverallStatus = "정상"
    End If
End Function

Private Function VerdictName(ByVal verdict As Long) As String
    If verdict = V_OK Then
        VerdictName = "정상"
    ElseIf verdict = V_PARTIAL Then
        VerdictName = "일부 오류"
    Else
        VerdictName = "실패"
    End If
End Function

Private Function SummaryText(ByVal title As String, ByVal endTime As Date, ByVal secs As Double, _
                             ByVal overall As String, ByVal maxReason As Long) As String
    Dim i As Long
    Dim s As String
    Dim lineText As String
    s = title & " " & IsoDateText(endTime, True) & " 종료 (소요 " & CLng(secs) & "초)"
    If m_cancelled Then s = s & " - 사용자 중단"
    s = s & vbLf & "결과: " & overall
    For i = 1 To m_n
        lineText = i & ". " & m_names(i) & ": " & VerdictName(m_verdict(i))
        If m_verdict(i) = V_PARTIAL Then lineText = lineText & " " & m_count(i) & "건"
        If Len(m_reason(i)) > 0 Then lineText = lineText & " (" & Clip(m_reason(i), maxReason) & ")"
        s = s & vbLf & lineText
    Next i
    If Len(m_notes) > 0 Then s = s & vbLf & m_notes
    SummaryText = s
End Function

' 요약 대화상자는 Excel 창이 보이고 tblRunCtl의 quiet가 Y가 아닐 때만 띄운다(숨김 자동화에서 멈추지 않도록).
' VBA MsgBox는 글자를 코드 페이지 949로 바꿔 보여 주므로 줄표(U+2014)는 949에 있는 가로줄(U+2015)로 바꾼다.
' 셀(..._상태, last_summary)에는 원래 줄표를 그대로 쓴다.
Private Sub ShowSummary(ByVal msgText As String, ByVal overall As String, ByVal title As String)
    Dim icon As VbMsgBoxStyle
    If Not Application.Visible Then Exit Sub
    If IsQuiet() Then Exit Sub
    If overall = "정상" Then
        icon = vbInformation
    Else
        icon = vbExclamation
    End If
    MsgBox Clip(Replace(msgText, EmDash(), ChrW(&H2015)), 1000), icon, "KIS 대시보드 " & title
End Sub

Private Function IsQuiet() As Boolean
    IsQuiet = (UCase$(Trim$(CellText(GetKeyValue("tblRunCtl", "quiet")))) = "Y")
End Function

Private Sub AddResult(ByVal stepName As String, ByVal tableName As String, ByVal verdict As Long, _
                      ByVal cnt As Long, ByVal reason As String)
    m_n = m_n + 1
    ReDim Preserve m_names(1 To m_n)
    ReDim Preserve m_tables(1 To m_n)
    ReDim Preserve m_verdict(1 To m_n)
    ReDim Preserve m_count(1 To m_n)
    ReDim Preserve m_reason(1 To m_n)
    m_names(m_n) = stepName
    m_tables(m_n) = tableName
    m_verdict(m_n) = verdict
    m_count(m_n) = cnt
    m_reason(m_n) = reason
End Sub

Private Sub AddNote(ByVal note As String)
    If Len(m_notes) > 0 Then m_notes = m_notes & vbLf
    m_notes = m_notes & "[주의] " & CleanText(note)
End Sub

'--------------------------------------------------------------------------------------------------
' 이력 CSV: tblEstSnap -> <통합문서 폴더>\history\est_snap.csv (UTF-8 BOM)
'--------------------------------------------------------------------------------------------------
Private Sub SaveEstSnapCsvCore(ByRef verdict As Long, ByRef cnt As Long, ByRef reason As String)
    Dim lo As ListObject
    Dim fso As Object
    Dim folder As String
    Dim filePath As String
    Dim nRows As Long
    Dim oldRows As Long
    Dim readWhy As String
    Dim altName As String
    verdict = V_FAIL
    cnt = 0
    reason = ""
    On Error GoTo Fail
    Set lo = FindTable("tblEstSnap")
    If lo Is Nothing Then
        reason = "표 없음: tblEstSnap"
        Exit Sub
    End If
    If Len(ThisWorkbook.Path) = 0 Then
        reason = "통합문서가 아직 저장되지 않아 폴더를 알 수 없음"
        Exit Sub
    End If
    If StrComp(Left$(ThisWorkbook.Path, 4), "http", vbTextCompare) = 0 Then
        reason = "통합문서 경로가 웹 주소(OneDrive 등)라 CSV를 쓸 수 없음"
        Exit Sub
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = ThisWorkbook.Path & "\" & HISTORY_DIR
    filePath = folder & "\" & SNAP_FILE
    nRows = lo.ListRows.Count
    If nRows = 0 And fso.FileExists(filePath) Then
        reason = "tblEstSnap이 비어 있어 기존 CSV를 덮어쓰지 않음"
        Exit Sub
    End If
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    ' 스냅샷 표는 쌓기만 하므로 기존 CSV보다 행이 적으면(또는 기존 행 수를 모르면) 덮지 않고 따로 저장한다(머리 주석 4번)
    If fso.FileExists(filePath) Then
        oldRows = CsvDataRows(filePath, readWhy)
        If oldRows < 0 Or nRows < oldRows Then
            altName = SNAP_ALT_PREFIX & FileStamp(Now) & ".csv"
            WriteUtf8Bom folder & "\" & altName, TableCsv(lo)
            verdict = V_PARTIAL
            cnt = 1
            If oldRows < 0 Then
                reason = PREFIX_ERR & " 기존 " & SNAP_FILE & "의 행 수를 읽지 못해(" & Clip(readWhy, 80) & ") 덮어쓰지 않고 " & _
                         nRows & "행을 " & HISTORY_DIR & "\" & altName & "로 따로 저장"
            Else
                reason = PREFIX_ERR & " tblEstSnap " & nRows & "행이 기존 " & SNAP_FILE & " " & oldRows & "행보다 적어 " & _
                         "기존 파일을 그대로 두고 " & HISTORY_DIR & "\" & altName & "로 따로 저장"
            End If
            Exit Sub
        End If
    End If
    WriteUtf8Bom filePath, TableCsv(lo)
    verdict = V_OK
    reason = nRows & "행 저장: " & HISTORY_DIR & "\" & SNAP_FILE
    Exit Sub
Fail:
    If Err.Number = 18 Then m_cancelled = True
    verdict = V_FAIL
    reason = "CSV 저장 실패: " & CleanText(Err.Description)
End Sub

Private Function TableCsv(ByVal lo As ListObject) As String
    Dim nCols As Long
    Dim nRows As Long
    Dim r As Long
    Dim c As Long
    Dim data As Variant
    Dim one As Variant
    Dim lines() As String
    Dim fields() As String
    Dim hasTime() As Boolean
    nCols = lo.ListColumns.Count
    nRows = lo.ListRows.Count
    ReDim lines(0 To nRows)
    ReDim fields(1 To nCols)
    ReDim hasTime(1 To nCols)
    For c = 1 To nCols
        fields(c) = CsvText(lo.ListColumns(c).Name)
    Next c
    lines(0) = Join(fields, ",")
    If nRows > 0 Then
        data = lo.DataBodyRange.Value
        If Not IsArray(data) Then
            one = data
            ReDim data(1 To 1, 1 To 1)
            data(1, 1) = one
        End If
        ' 시각이 들어 있는 날짜 열은 열 전체를 일시 형식으로 쓴다(자정 값만 날짜로 바뀌지 않도록)
        For c = 1 To nCols
            For r = 1 To nRows
                If VarType(data(r, c)) = vbDate Then
                    If HasTimePart(data(r, c)) Then
                        hasTime(c) = True
                        Exit For
                    End If
                End If
            Next r
        Next c
        For r = 1 To nRows
            For c = 1 To nCols
                fields(c) = CsvValue(data(r, c), hasTime(c))
            Next c
            lines(r) = Join(fields, ",")
        Next r
    End If
    TableCsv = Join(lines, vbCrLf) & vbCrLf
End Function

Private Function CsvValue(ByVal v As Variant, ByVal withTime As Boolean) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        CsvValue = ""
        Exit Function
    End If
    Select Case VarType(v)
        Case vbDate
            CsvValue = IsoDateText(v, withTime)
        Case vbBoolean
            If v Then CsvValue = "TRUE" Else CsvValue = "FALSE"
        Case vbString
            CsvValue = CsvText(CStr(v))
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger, vbByte
            CsvValue = NumText(v)
        Case Else
            CsvValue = CsvText(CStr(v))
    End Select
End Function

' Str$는 로캘과 무관하게 소수점을 "."으로 쓴다. 앞 공백을 지우고 ".5" 꼴은 "0.5"로 고친다.
Private Function NumText(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(Str$(v))
    If Left$(s, 1) = "." Then
        s = "0" & s
    ElseIf Left$(s, 2) = "-." Then
        s = "-0" & Mid$(s, 2)
    End If
    NumText = s
End Function

Private Function CsvText(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        CsvText = """" & Replace(s, """", """""") & """"
    Else
        CsvText = s
    End If
End Function

Private Function HasTimePart(ByVal v As Variant) As Boolean
    Dim secs As Double
    secs = Round(CDbl(v) * 86400#, 0)
    HasTimePart = (secs - Int(secs / 86400#) * 86400#) <> 0
End Function

' 초 단위로 반올림한 뒤 yyyy-mm-dd[ hh:nn:ss]. Format$의 날짜·시간 구분자는 로캘을 따르므로 숫자만 서식한다.
Private Function IsoDateText(ByVal v As Variant, ByVal withTime As Boolean) As String
    Dim d As Date
    Dim s As String
    d = CDate(Round(CDbl(v) * 86400#, 0) / 86400#)
    s = Format$(Year(d), "0000") & "-" & Format$(Month(d), "00") & "-" & Format$(Day(d), "00")
    If withTime Then
        s = s & " " & Format$(Hour(d), "00") & ":" & Format$(Minute(d), "00") & ":" & Format$(Second(d), "00")
    End If
    IsoDateText = s
End Function

' ADODB.Stream의 utf-8 문자 집합은 BOM(EF BB BF)을 붙여 저장한다.
Private Sub WriteUtf8Bom(ByVal filePath As String, ByVal content As String)
    Dim st As Object
    Dim errNo As Long
    Dim errDesc As String
    Set st = CreateObject("ADODB.Stream")
    On Error GoTo Fail
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText content
    st.SaveToFile filePath, 2
    st.Close
    Exit Sub
Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Resume CloseStream
CloseStream:
    On Error Resume Next
    st.Close
    On Error GoTo 0
    Err.Raise errNo, "WriteUtf8Bom", errDesc
End Sub

' UTF-8 텍스트 파일 전체를 읽는다(ADODB.Stream의 utf-8 문자 집합은 앞의 BOM을 건너뜀).
Private Function ReadUtf8(ByVal filePath As String) As String
    Dim st As Object
    Dim errNo As Long
    Dim errDesc As String
    Set st = CreateObject("ADODB.Stream")
    On Error GoTo Fail
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile filePath
    ReadUtf8 = st.ReadText(-1)
    st.Close
    Exit Function
Fail:
    errNo = Err.Number
    errDesc = Err.Description
    Resume CloseStream
CloseStream:
    On Error Resume Next
    st.Close
    On Error GoTo 0
    Err.Raise errNo, "ReadUtf8", errDesc
End Function

' CSV 파일의 데이터 행 수(머리글 1행과 끝의 빈 줄 제외). 큰따옴표 안의 줄바꿈은 행으로 세지 않는다
' (큰따옴표로 나눈 조각 중 짝수 번째가 따옴표 밖 - 이스케이프된 "" 는 빈 조각이 되어 안팎이 그대로 유지됨).
' 읽지 못하면 -1을 돌려주고 사유를 why에 쓴다.
Private Function CsvDataRows(ByVal filePath As String, ByRef why As String) As Long
    Dim fileBody As String
    Dim pieces As Variant
    Dim i As Long
    Dim n As Long
    why = ""
    On Error GoTo Fail
    fileBody = Replace(Replace(ReadUtf8(filePath), vbCrLf, vbLf), vbCr, vbLf)
    Do While Right$(fileBody, 1) = vbLf
        fileBody = Left$(fileBody, Len(fileBody) - 1)
    Loop
    If Len(fileBody) = 0 Then Exit Function
    pieces = Split(fileBody, """")
    For i = LBound(pieces) To UBound(pieces) Step 2
        n = n + Len(pieces(i)) - Len(Replace(pieces(i), vbLf, ""))
    Next i
    CsvDataRows = n
    Exit Function
Fail:
    why = CleanText(Err.Description)
    If Len(why) = 0 Then why = "오류 번호 " & Err.Number
    CsvDataRows = -1
End Function

' 파일 이름용 시각 yyyymmdd_hhnnss. Format$의 날짜 서식은 로캘을 따르므로 숫자만 서식한다.
Private Function FileStamp(ByVal d As Date) As String
    FileStamp = Format$(Year(d), "0000") & Format$(Month(d), "00") & Format$(Day(d), "00") & "_" & _
                Format$(Hour(d), "00") & Format$(Minute(d), "00") & Format$(Second(d), "00")
End Function

'--------------------------------------------------------------------------------------------------
' 표·이름·키/값 도우미
'--------------------------------------------------------------------------------------------------
Private Function FindTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function ColIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).Name), colName, vbTextCompare) = 0 Then
            ColIndex = i
            Exit Function
        End If
    Next i
End Function

' 표 열의 값을 항상 (1 To n, 1 To 1) 2차원 배열로 돌려준다(1행이면 Range.Value가 배열이 아니므로).
Private Function ColumnValues(ByVal lo As ListObject, ByVal colIdx As Long) As Variant
    Dim rng As Range
    Dim v As Variant
    Dim a() As Variant
    Set rng = lo.ListColumns(colIdx).DataBodyRange
    If rng Is Nothing Then
        ReDim a(1 To 1, 1 To 1)
        ColumnValues = a
        Exit Function
    End If
    v = rng.Value
    If IsArray(v) Then
        ColumnValues = v
    Else
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = v
        ColumnValues = a
    End If
End Function

Private Function FindKeyRow(ByVal lo As ListObject, ByVal keyCol As Long, ByVal key As String) As Long
    Dim v As Variant
    Dim r As Long
    If lo.ListRows.Count = 0 Then Exit Function
    v = ColumnValues(lo, keyCol)
    For r = 1 To lo.ListRows.Count
        If StrComp(Trim$(CellText(v(r, 1))), key, vbTextCompare) = 0 Then
            FindKeyRow = r
            Exit Function
        End If
    Next r
End Function

' 키/값 표(tblRunCtl, tblSettings)의 값 쓰기. 키 행이 없으면 addIfMissing일 때만 추가.
Private Function SetKeyValue(ByVal tableName As String, ByVal key As String, ByVal v As Variant, _
                             ByVal addIfMissing As Boolean, ByRef why As String) As Boolean
    Dim lo As ListObject
    Dim kc As Long
    Dim vc As Long
    Dim r As Long
    Dim lr As ListRow
    why = ""
    On Error GoTo Fail
    Set lo = FindTable(tableName)
    If lo Is Nothing Then
        why = "표 없음: " & tableName
        Exit Function
    End If
    kc = ColIndex(lo, "키")
    vc = ColIndex(lo, "값")
    If kc = 0 Or vc = 0 Then
        why = tableName & "에 키/값 열 없음"
        Exit Function
    End If
    r = FindKeyRow(lo, kc, key)
    If r = 0 Then
        If Not addIfMissing Then
            why = tableName & "에 키 " & key & " 없음"
            Exit Function
        End If
        Set lr = lo.ListRows.Add
        lr.Range.Cells(1, kc).Value = key
        lr.Range.Cells(1, vc).Value = v
    Else
        lo.DataBodyRange.Cells(r, vc).Value = v
    End If
    SetKeyValue = True
    Exit Function
Fail:
    why = tableName & "의 " & key & " 기록 실패: " & CleanText(Err.Description)
    SetKeyValue = False
End Function

Private Function GetKeyValue(ByVal tableName As String, ByVal key As String) As Variant
    Dim lo As ListObject
    Dim kc As Long
    Dim vc As Long
    Dim r As Long
    GetKeyValue = Empty
    On Error GoTo Fail
    Set lo = FindTable(tableName)
    If lo Is Nothing Then Exit Function
    kc = ColIndex(lo, "키")
    vc = ColIndex(lo, "값")
    If kc = 0 Or vc = 0 Then Exit Function
    r = FindKeyRow(lo, kc, key)
    If r > 0 Then GetKeyValue = lo.DataBodyRange.Cells(r, vc).Value
    Exit Function
Fail:
    GetKeyValue = Empty
End Function

' 이름 정의가 가리키는 첫 셀. 통합문서 범위 이름을 먼저, 없으면 시트 범위 이름(시트!이름)을 찾는다.
Private Function NamedCell(ByVal nm As String) As Range
    Dim n As Name
    Dim hit As Name
    Dim p As Long
    On Error Resume Next
    Set hit = ThisWorkbook.Names(nm)
    If hit Is Nothing Then
        For Each n In ThisWorkbook.Names
            p = InStrRev(n.Name, "!")
            If p > 0 Then
                If StrComp(Mid$(n.Name, p + 1), nm, vbTextCompare) = 0 Then
                    Set hit = n
                    Exit For
                End If
            End If
        Next n
    End If
    If Not hit Is Nothing Then Set NamedCell = hit.RefersToRange.Cells(1, 1)
End Function

Private Function NamedValue(ByVal nm As String) As Variant
    Dim c As Range
    NamedValue = Empty
    On Error Resume Next
    Set c = NamedCell(nm)
    If Not c Is Nothing Then NamedValue = c.Value
End Function

Private Function WriteNamed(ByVal nm As String, ByVal v As Variant, ByRef why As String) As Boolean
    Dim c As Range
    why = ""
    On Error GoTo Fail
    Set c = NamedCell(nm)
    If c Is Nothing Then
        why = "이름 정의 없음: " & nm
        Exit Function
    End If
    c.Value = v
    WriteNamed = True
    Exit Function
Fail:
    why = "이름 " & nm & " 기록 실패: " & CleanText(Err.Description)
    WriteNamed = False
End Function

'--------------------------------------------------------------------------------------------------
' 값·문자열 도우미
'--------------------------------------------------------------------------------------------------
Private Function CellText(ByVal v As Variant) As String
    If IsArray(v) Then Exit Function
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    CellText = CStr(v)
End Function

' 종목코드 정규화: 숫자로 저장된 5자리 이하 코드는 앞에 0을 채워 6자리로 맞춘다(같은 종목을 두 번 세지 않도록).
Private Function NormCode(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CellText(v))
    If Len(s) > 0 And Len(s) < 6 Then
        If Not s Like "*[!0-9]*" Then s = Right$("000000" & s, 6)
    End If
    NormCode = UCase$(s)
End Function

Private Function TryDate(ByVal v As Variant, ByRef d As Date) As Boolean
    On Error GoTo Fail
    If IsArray(v) Then Exit Function
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            d = v
        Case vbDouble, vbSingle, vbCurrency, vbDecimal, vbLong, vbInteger
            d = CDate(CDbl(v))
        Case vbString
            If Not IsDate(v) Then Exit Function
            d = CDate(v)
        Case Else
            Exit Function
    End Select
    TryDate = True
    Exit Function
Fail:
    TryDate = False
End Function

Private Function CleanText(ByVal s As String) As String
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    CleanText = Trim$(s)
End Function

Private Function Clip(ByVal s As String, ByVal maxLen As Long) As String
    If Len(s) > maxLen And maxLen > 1 Then
        Clip = Left$(s, maxLen - 1) & Ellipsis()
    Else
        Clip = s
    End If
End Function

Private Function ElapsedSince(ByVal t0 As Double) As Double
    Dim t As Double
    t = Timer
    If t < t0 Then t = t + 86400#
    ElapsedSince = t - t0
End Function

Private Function EmDash() As String
    EmDash = ChrW(&H2014)
End Function

Private Function Ellipsis() As String
    Ellipsis = ChrW(&H2026)
End Function

' 시계(Now, 초 단위)가 다음 초로 넘어갈 때까지 기다린다(1초 이내). 예열 뒤 started가 예열의 조회시각보다 늦게 하려는 것(머리 주석 1번).
Private Sub WaitNextSecond()
    Dim t As Date
    t = Now
    Do While Now = t
        DoEvents
    Loop
End Sub
