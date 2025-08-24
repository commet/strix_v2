Attribute VB_Name = "modMain"
' =====================================
' STRIX v2 - Main Module
' Executive Intelligence System
' =====================================
Option Explicit

' 전역 상수
Public Const APP_NAME As String = "STRIX Executive v2"
Public Const API_URL As String = "http://localhost:5000/api"
Public Const VERSION As String = "2.0.0"

' 전역 변수
Public g_CurrentPhase As Integer
Public g_LastUpdate As Date
Public g_UserRole As String

' =====================================
' 메인 진입점
' =====================================
Sub CreateExecutiveDashboard()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 초기화
    Call InitializeSystem
    
    ' 메인 대시보드 생성
    Call CreateMainDashboard
    
    ' 3단계 워크플로우 탭 생성
    Call CreatePhase1Tab    ' 보고 준비 이전
    Call CreatePhase2Tab    ' 보고 준비
    Call CreatePhase3Tab    ' 보고 이후
    
    ' 부가 기능 탭
    Call CreateAnalyticsTab ' 분석 대시보드
    Call CreateAlertsTab    ' Smart Alerts
    
    ' 초기 데이터 로드
    Call LoadInitialData
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' 완료 메시지
    MsgBox "STRIX Executive v2가 성공적으로 생성되었습니다!" & vbLf & vbLf & _
           "🎯 주요 기능:" & vbLf & _
           "• 3단계 업무 워크플로우 자동화" & vbLf & _
           "• AI 기반 실시간 분석" & vbLf & _
           "• 피드백 자동 학습 시스템" & vbLf & _
           "• Critical Issue 즉시 알림" & vbLf & vbLf & _
           "경영진 시연 준비 완료!", _
           vbInformation, APP_NAME
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "오류 발생: " & Err.Description, vbCritical, APP_NAME
End Sub

' =====================================
' 시스템 초기화
' =====================================
Private Sub InitializeSystem()
    ' 기존 시트 정리
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "STRIX*" Or ws.Name Like "Phase*" Then
            ws.Delete
        End If
    Next ws
    
    ' 전역 변수 초기화
    g_CurrentPhase = 1
    g_LastUpdate = Now
    g_UserRole = "Executive"
    
    ' 설정 로드
    Call LoadSettings
End Sub

' =====================================
' 메인 대시보드 생성
' =====================================
Private Sub CreateMainDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "STRIX Dashboard"
    
    With ws
        ' 전체 배경
        .Cells.Interior.Color = RGB(248, 249, 250)
        
        ' 열 너비 설정
        .Columns("A").ColumnWidth = 2
        .Columns("B:M").ColumnWidth = 12
        .Columns("N").ColumnWidth = 2
        
        ' 메인 헤더
        With .Range("B2:M3")
            .Merge
            .Value = "STRIX Executive Intelligence System v2"
            .Font.Name = "맑은 고딕"
            .Font.Size = 32
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(25, 45, 95)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' 서브 헤더
        With .Range("B4:M4")
            .Merge
            .Value = "AI 기반 보고 업무 자동화 | 실시간 업데이트: " & Format(Now, "yyyy-mm-dd hh:mm")
            .Font.Size = 14
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter
        End With
        
        ' 3단계 프로세스 카드
        Call CreateProcessCards(ws, 6)
        
        ' 실시간 지표
        Call CreateMetricsSection(ws, 15)
        
        ' 빠른 실행 버튼
        Call CreateQuickActions(ws, 25)
        
        ' 최근 활동 로그
        Call CreateActivityLog(ws, 35)
    End With
    
    ' 화면 설정
    ActiveWindow.DisplayGridlines = False
    ws.Range("B2").Select
End Sub

' =====================================
' 프로세스 카드 생성
' =====================================
Private Sub CreateProcessCards(ws As Worksheet, startRow As Integer)
    Dim phases As Variant
    Dim colors As Variant
    Dim i As Integer
    
    phases = Array( _
        Array("Phase 1: 보고 준비 이전", "이전 피드백 확인 & 자료 수집", "📥"), _
        Array("Phase 2: 보고 준비", "자료 종합 & AI 분석 & 보고서 작성", "📝"), _
        Array("Phase 3: 보고 이후", "피드백 반영 & RAG 업데이트 & 추적", "📤") _
    )
    
    colors = Array(RGB(52, 152, 219), RGB(46, 204, 113), RGB(155, 89, 182))
    
    For i = 0 To 2
        With ws.Range(ws.Cells(startRow, 2 + i * 4), ws.Cells(startRow + 5, 5 + i * 4))
            .Merge
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
        
        ' 카드 헤더
        With ws.Range(ws.Cells(startRow, 2 + i * 4), ws.Cells(startRow + 1, 5 + i * 4))
            .Interior.Color = colors(i)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .Font.Size = 14
            .Value = phases(i)(2) & " " & phases(i)(0)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' 카드 내용
        With ws.Range(ws.Cells(startRow + 2, 2 + i * 4), ws.Cells(startRow + 5, 5 + i * 4))
            .Value = phases(i)(1)
            .WrapText = True
            .Font.Size = 11
            .VerticalAlignment = xlTop
            .HorizontalAlignment = xlLeft
        End With
        
        ' 실행 버튼
        Dim btn As Object
        Set btn = ws.Buttons.Add( _
            ws.Cells(startRow + 6, 2 + i * 4).Left, _
            ws.Cells(startRow + 6, 2 + i * 4).Top, _
            150, 30)
        
        With btn
            .Caption = "Phase " & (i + 1) & " 실행"
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
        .Value = "📊 실시간 핵심 지표"
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = RGB(41, 128, 185)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 지표 데이터
    Dim metrics As Variant
    metrics = Array( _
        Array("수집 문서", "182", "건", RGB(52, 152, 219)), _
        Array("분석 완료", "95", "%", RGB(46, 204, 113)), _
        Array("Critical Issues", "3", "건", RGB(231, 76, 60)), _
        Array("처리 시간", "2.3", "시간", RGB(241, 196, 15)), _
        Array("피드백 반영", "100", "%", RGB(155, 89, 182)), _
        Array("정확도", "94", "%", RGB(26, 188, 156)) _
    )
    
    Dim col As Integer
    For col = 0 To 5
        Dim metricCol As Integer
        metricCol = 2 + col * 2
        
        ' 지표명
        ws.Cells(startRow + 2, metricCol).Value = metrics(col)(0)
        ws.Cells(startRow + 2, metricCol).Font.Size = 10
        ws.Cells(startRow + 2, metricCol).Font.Color = RGB(100, 100, 100)
        
        ' 지표값
        With ws.Range(ws.Cells(startRow + 3, metricCol), ws.Cells(startRow + 4, metricCol + 1))
            .Merge
            .Value = metrics(col)(1) & metrics(col)(2)
            .Font.Size = 24
            .Font.Bold = True
            .Font.Color = metrics(col)(3)
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
        .Value = "⚡ 빠른 실행"
        .Font.Bold = True
        .Font.Size = 16
        .Interior.Color = RGB(44, 62, 80)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 버튼 생성
    Dim actions As Variant
    actions = Array( _
        "전체 워크플로우 실행", _
        "AI 통합 분석", _
        "보고서 생성", _
        "피드백 수집", _
        "Smart Alert 확인", _
        "설정" _
    )
    
    Dim i As Integer
    For i = 0 To 5
        Dim btn As Object
        Set btn = ws.Buttons.Add( _
            ws.Cells(startRow + 2, 2 + i * 2).Left, _
            ws.Cells(startRow + 2, 2 + i * 2).Top, _
            140, 35)
        
        With btn
            .Caption = actions(i)
            .OnAction = "QuickAction" & (i + 1)
            .Font.Size = 11
        End With
    Next i
End Sub

' =====================================
' 활동 로그
' =====================================
Private Sub CreateActivityLog(ws As Worksheet, startRow As Integer)
    ' 섹션 헤더
    With ws.Range("B" & startRow & ":M" & startRow)
        .Merge
        .Value = "📝 최근 활동"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(149, 165, 166)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 로그 영역
    With ws.Range("B" & (startRow + 1) & ":M" & (startRow + 5))
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(220, 220, 220)
        .Font.Size = 10
        .Font.Color = RGB(100, 100, 100)
    End With
    
    ' 샘플 로그
    ws.Cells(startRow + 1, 2).Value = Format(Now - 0.1, "hh:mm") & " - AI 분석 완료: SK온 합병 영향 분석"
    ws.Cells(startRow + 2, 2).Value = Format(Now - 0.05, "hh:mm") & " - Critical Alert: BYD 5분 충전 기술 발표"
    ws.Cells(startRow + 3, 2).Value = Format(Now - 0.02, "hh:mm") & " - 보고서 생성: 2025년 8월 월간 보고"
    ws.Cells(startRow + 4, 2).Value = Format(Now, "hh:mm") & " - 피드백 수집: CEO 검토 완료"
End Sub

' =====================================
' Phase 실행 함수들
' =====================================
Sub RunPhase1()
    MsgBox "Phase 1: 보고 준비 이전 단계를 실행합니다." & vbLf & vbLf & _
           "• 이전 피드백 확인" & vbLf & _
           "• 자료 수집 (182건)" & vbLf & _
           "• AI 이슈 식별", _
           vbInformation, "Phase 1 실행"
    Call Phase1_Execute
End Sub

Sub RunPhase2()
    MsgBox "Phase 2: 보고 준비 단계를 실행합니다." & vbLf & vbLf & _
           "• 자료 종합 분석" & vbLf & _
           "• AI 보고서 작성" & vbLf & _
           "• 핵심 인사이트 도출", _
           vbInformation, "Phase 2 실행"
    Call Phase2_Execute
End Sub

Sub RunPhase3()
    MsgBox "Phase 3: 보고 이후 단계를 실행합니다." & vbLf & vbLf & _
           "• 피드백 수집/분류" & vbLf & _
           "• RAG 시스템 업데이트" & vbLf & _
           "• Issue Tracking", _
           vbInformation, "Phase 3 실행"
    Call Phase3_Execute
End Sub

' =====================================
' 빠른 실행 함수들
' =====================================
Sub QuickAction1()
    ' 전체 워크플로우 실행
    Call RunFullWorkflow
End Sub

Sub QuickAction2()
    ' AI 통합 분석
    Call RunAIAnalysis
End Sub

Sub QuickAction3()
    ' 보고서 생성
    Call GenerateExecutiveReport
End Sub

Sub QuickAction4()
    ' 피드백 수집
    Call CollectFeedback
End Sub

Sub QuickAction5()
    ' Smart Alert 확인
    Call ShowSmartAlerts
End Sub

Sub QuickAction6()
    ' 설정
    Call ShowSettings
End Sub

' =====================================
' 설정 관련
' =====================================
Private Sub LoadSettings()
    ' 설정 파일에서 로드 (구현 예정)
End Sub

Private Sub ShowSettings()
    MsgBox "설정 화면 (구현 예정)", vbInformation, APP_NAME
End Sub

' =====================================
' 초기 데이터 로드
' =====================================
Private Sub LoadInitialData()
    ' API에서 초기 데이터 로드 (구현 예정)
End Sub