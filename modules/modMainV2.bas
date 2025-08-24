Attribute VB_Name = "modMainV2"
' =====================================
' STRIX v2 - Main Module (한글 안정화 버전)
' 보고 업무 자동화 시스템
' =====================================
Option Explicit

' =====================================
' 메인 진입점 - 대시보드 생성
' =====================================
Sub CreateSTRIXDashboard()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 모듈 초기화
    Call modConfig.InitializeConfig
    Call modLabels.InitializeLabels
    
    ' 기존 시트 정리
    Call CleanupSheets
    
    ' 메인 대시보드 생성
    Call CreateMainDashboard
    
    ' Mock 데이터 로드
    Call LoadInitialMockData
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' 완료 메시지
    MsgBox GetLabel("MSG_WELCOME") & vbLf & vbLf & _
           GetLabel("MAIN_TITLE") & vbLf & _
           "Version: " & VERSION & vbLf & vbLf & _
           GetLabel("STATUS_READY"), _
           vbInformation, APP_NAME
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox GetLabel("ERR_GENERAL") & ": " & Err.Description, vbCritical, APP_NAME
End Sub

' =====================================
' 기존 시트 정리
' =====================================
Private Sub CleanupSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "STRIX*" Or ws.Name Like "Phase*" Then
            ws.Delete
        End If
    Next ws
End Sub

' =====================================
' 메인 대시보드 생성
' =====================================
Private Sub CreateMainDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = SHEET_MAIN
    
    With ws
        ' 전체 배경
        .Cells.Interior.Color = COLOR_BACKGROUND
        
        ' 열 너비 설정
        .Columns("A").ColumnWidth = 2
        .Columns("B:M").ColumnWidth = 12
        .Columns("N").ColumnWidth = 2
        
        ' 메인 헤더 - 라벨 사용
        With .Range("B2:M3")
            .Merge
            .Value = GetLabel("MAIN_TITLE")
            .Font.Name = "맑은 고딕"
            .Font.Size = 28
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = COLOR_PRIMARY
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' 서브 헤더 - 동적 시간 표시
        With .Range("B4:M4")
            .Merge
            .Value = GetLabel("MAIN_SUBTITLE") & " | " & _
                    GetLabel("LAST_UPDATE") & ": " & _
                    FormatDateTime(Now)
            .Font.Size = 14
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter
        End With
        
        ' 3단계 프로세스 카드
        Call CreateProcessCards(ws, 6)
        
        ' 실시간 지표
        Call CreateMetricsSection(ws, 16)
        
        ' 빠른 실행 버튼
        Call CreateQuickActions(ws, 26)
        
        ' Mock 데이터 영역
        Call CreateMockDataSection(ws, 34)
    End With
    
    ' 화면 설정
    ActiveWindow.DisplayGridlines = False
    ws.Range("B2").Select
End Sub

' =====================================
' 프로세스 카드 생성
' =====================================
Private Sub CreateProcessCards(ws As Worksheet, startRow As Integer)
    Dim i As Integer
    Dim phases As Variant
    Dim colors As Variant
    
    ' 3단계 정보
    phases = Array( _
        Array("PHASE1_TITLE", "PHASE1_DESC", "📥"), _
        Array("PHASE2_TITLE", "PHASE2_DESC", "📝"), _
        Array("PHASE3_TITLE", "PHASE3_DESC", "📤") _
    )
    
    colors = Array(COLOR_PHASE1, COLOR_PHASE2, COLOR_PHASE3)
    
    For i = 0 To 2
        With ws.Range(ws.Cells(startRow, 2 + i * 4), ws.Cells(startRow + 6, 5 + i * 4))
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
        
        ' 카드 헤더
        With ws.Range(ws.Cells(startRow, 2 + i * 4), ws.Cells(startRow + 1, 5 + i * 4))
            .Merge
            .Interior.Color = colors(i)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 14
            .Value = phases(i)(2) & " " & GetLabel(phases(i)(0))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' 카드 설명
        With ws.Range(ws.Cells(startRow + 2, 2 + i * 4), ws.Cells(startRow + 5, 5 + i * 4))
            .Merge
            .Value = GetLabel(phases(i)(1))
            .WrapText = True
            .Font.Size = 11
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlLeft
        End With
        
        ' 실행 버튼
        Dim btn As Object
        Set btn = ws.Buttons.Add( _
            ws.Cells(startRow + 7, 2 + i * 4).Left, _
            ws.Cells(startRow + 7, 2 + i * 4).Top, _
            150, 30)
        
        With btn
            .Caption = "Phase " & (i + 1) & " " & GetLabel("BTN_EXECUTE")
            .OnAction = "RunPhase" & (i + 1)
            .Font.Size = 11
        End With
    Next i
End Sub

' =====================================
' 실시간 지표 섹션
' =====================================
Private Sub CreateMetricsSection(ws As Worksheet, startRow As Integer)
    ' 섹션 헤더
    With ws.Range("B" & startRow & ":M" & startRow)
        .Merge
        .Value = "📊 " & GetLabel("STATUS_READY")
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = COLOR_INFO
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 지표 데이터
    Dim metrics As Variant
    metrics = Array( _
        Array("METRIC_DOCS", "182", "METRIC_COUNT"), _
        Array("METRIC_ANALYZED", "95", "METRIC_PERCENT"), _
        Array("METRIC_ISSUES", "7", "METRIC_COUNT"), _
        Array("METRIC_ACTIONS", "12", "METRIC_COUNT"), _
        Array("METRIC_FEEDBACK", "3", "METRIC_COUNT"), _
        Array("METRIC_ACCURACY", "94", "METRIC_PERCENT") _
    )
    
    Dim col As Integer
    For col = 0 To 5
        Dim metricCol As Integer
        metricCol = 2 + col * 2
        
        ' 지표명
        ws.Cells(startRow + 2, metricCol).Value = GetLabel(metrics(col)(0))
        ws.Cells(startRow + 2, metricCol).Font.Size = 10
        ws.Cells(startRow + 2, metricCol).Font.Color = RGB(100, 100, 100)
        
        ' 지표값
        With ws.Range(ws.Cells(startRow + 3, metricCol), ws.Cells(startRow + 4, metricCol + 1))
            .Merge
            .Value = metrics(col)(1) & GetLabel(metrics(col)(2))
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = COLOR_SUCCESS
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next col
End Sub

' =====================================
' 빠른 실행 버튼
' =====================================
Private Sub CreateQuickActions(ws As Worksheet, startRow As Integer)
    ' 섹션 헤더
    With ws.Range("B" & startRow & ":M" & startRow)
        .Merge
        .Value = "⚡ " & GetLabel("BTN_EXECUTE")
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = RGB(44, 62, 80)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 버튼 생성
    Dim actions As Variant
    actions = Array( _
        Array("전체 워크플로우", "RunFullWorkflow"), _
        Array("AI " & GetLabel("BTN_ANALYZE"), "RunAIAnalysis"), _
        Array(GetLabel("BTN_GENERATE"), "GenerateReport"), _
        Array(GetLabel("PHASE3_FEEDBACK"), "CollectFeedback"), _
        Array(GetLabel("BTN_REFRESH"), "RefreshData"), _
        Array(GetLabel("BTN_SETTINGS"), "ShowSettings") _
    )
    
    Dim i As Integer
    For i = 0 To 5
        Dim btn As Object
        Set btn = ws.Buttons.Add( _
            ws.Cells(startRow + 2, 2 + i * 2).Left, _
            ws.Cells(startRow + 2, 2 + i * 2).Top, _
            140, 35)
        
        With btn
            .Caption = actions(i)(0)
            .OnAction = actions(i)(1)
            .Font.Size = 11
        End With
    Next i
End Sub

' =====================================
' Mock 데이터 섹션
' =====================================
Private Sub CreateMockDataSection(ws As Worksheet, startRow As Integer)
    ' 섹션 헤더
    With ws.Range("B" & startRow & ":M" & startRow)
        .Merge
        .Value = "📁 " & GetLabel("PHASE1_COLLECT")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(149, 165, 166)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 테이블 헤더
    Dim headers As Variant
    headers = Array("COL_NO", "COL_TITLE", "COL_CATEGORY", _
                   "COL_SOURCE", "COL_DATE", "COL_TYPE")
    
    Dim col As Integer
    For col = 0 To 5
        ws.Cells(startRow + 1, 2 + col * 2).Value = GetLabel(headers(col))
        With ws.Range(ws.Cells(startRow + 1, 2 + col * 2), _
                     ws.Cells(startRow + 1, 3 + col * 2))
            .Merge
            .Font.Bold = True
            .Interior.Color = RGB(230, 230, 230)
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
    Next col
End Sub

' =====================================
' 초기 Mock 데이터 로드
' =====================================
Private Sub LoadInitialMockData()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_MAIN)
    
    If ws Is Nothing Then Exit Sub
    
    ' Mock 데이터를 시트에 로드
    Call modMockData.LoadMockDataToSheet(ws, "all")
End Sub

' =====================================
' Phase 실행 함수들
' =====================================
Sub RunPhase1()
    MsgBox GetLabel("PHASE1_TITLE") & vbLf & vbLf & _
           "• " & GetLabel("PHASE1_FEEDBACK") & vbLf & _
           "• " & GetLabel("PHASE1_COLLECT") & " (182" & GetLabel("METRIC_COUNT") & ")" & vbLf & _
           "• " & GetLabel("PHASE1_IDENTIFY"), _
           vbInformation, GetLabel("PHASE1_TITLE")
End Sub

Sub RunPhase2()
    MsgBox GetLabel("PHASE2_TITLE") & vbLf & vbLf & _
           "• " & GetLabel("PHASE2_CONSOLIDATE") & vbLf & _
           "• " & GetLabel("PHASE2_ANALYZE") & vbLf & _
           "• " & GetLabel("PHASE2_WRITE"), _
           vbInformation, GetLabel("PHASE2_TITLE")
End Sub

Sub RunPhase3()
    MsgBox GetLabel("PHASE3_TITLE") & vbLf & vbLf & _
           "• " & GetLabel("PHASE3_FEEDBACK") & vbLf & _
           "• " & GetLabel("PHASE3_UPDATE") & vbLf & _
           "• " & GetLabel("PHASE3_TRACK"), _
           vbInformation, GetLabel("PHASE3_TITLE")
End Sub

' =====================================
' 빠른 실행 함수들
' =====================================
Sub RunFullWorkflow()
    Application.StatusBar = GetLabel("STATUS_PROCESSING")
    ' 전체 워크플로우 실행 로직
    Application.Wait Now + TimeValue("00:00:02")
    Application.StatusBar = GetLabel("STATUS_COMPLETE")
    MsgBox GetLabel("MSG_SUCCESS"), vbInformation, APP_NAME
End Sub

Sub RunAIAnalysis()
    Application.StatusBar = GetLabel("STATUS_ANALYZING")
    ' AI 분석 로직
    Application.Wait Now + TimeValue("00:00:02")
    Application.StatusBar = GetLabel("STATUS_COMPLETE")
End Sub

Sub GenerateReport()
    Application.StatusBar = GetLabel("STATUS_GENERATING")
    ' 보고서 생성 로직
    Application.Wait Now + TimeValue("00:00:02")
    Application.StatusBar = GetLabel("STATUS_COMPLETE")
End Sub

Sub CollectFeedback()
    Dim feedback As String
    feedback = InputBox(GetLabel("PHASE3_FEEDBACK"), GetLabel("PHASE3_TITLE"))
    If feedback <> "" Then
        MsgBox GetLabel("MSG_SUCCESS"), vbInformation
    End If
End Sub

Sub RefreshData()
    Application.StatusBar = GetLabel("STATUS_LOADING")
    Call LoadInitialMockData
    Application.StatusBar = GetLabel("STATUS_COMPLETE")
End Sub

Sub ShowSettings()
    MsgBox GetLabel("BTN_SETTINGS") & " (준비 중)", vbInformation, APP_NAME
End Sub