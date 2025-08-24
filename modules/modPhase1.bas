Attribute VB_Name = "modPhase1"
' =====================================
' STRIX v2 - Phase 1 Module
' 보고 준비 이전 단계
' =====================================
Option Explicit

' Phase 1 데이터 타입
Public Type FeedbackItem
    id As String
    reportDate As String
    department As String
    feedbackType As String  ' 개선요청/질문/지적사항
    content As String
    status As String  ' 반영완료/진행중/대기
    priority As String  ' High/Medium/Low
End Type

Public Type CollectedDocument
    id As String
    title As String
    source As String
    collectedDate As String
    docType As String
    relevance As Double
    summary As String
End Type

' =====================================
' Phase 1 시트 생성
' =====================================
Sub CreatePhase1Sheet()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_PHASE1)
    
    With ws
        .Cells.Clear
        .Cells.Interior.Color = COLOR_BACKGROUND
        
        ' 열 너비 설정
        .Columns("A").ColumnWidth = 2
        .Columns("B:E").ColumnWidth = 20
        .Columns("F").ColumnWidth = 3
        .Columns("G:J").ColumnWidth = 20
        .Columns("K").ColumnWidth = 3
        .Columns("L:N").ColumnWidth = 25
        .Columns("O").ColumnWidth = 2
        
        ' Phase 1 헤더
        With .Range("B2:N2")
            .Merge
            .Value = GetLabel("PHASE1_TITLE")
            .Font.Name = "맑은 고딕"
            .Font.Size = 24
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = COLOR_PHASE1
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 50
        End With
        
        ' 부제목
        With .Range("B3:N3")
            .Merge
            .Value = GetLabel("PHASE1_DESC") & " | " & FormatDateTime(Now)
            .Font.Size = 12
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter
            .RowHeight = 25
        End With
        
        ' 3개 섹션 생성
        Call CreateFeedbackSection(ws, 5)      ' 이전 피드백
        Call CreateCollectionSection(ws, 5)    ' 자료 수집
        Call CreateIssueSection(ws, 5)         ' 이슈 식별
    End With
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox GetLabel("ERR_GENERAL") & ": " & Err.Description, vbCritical
End Sub

' =====================================
' 이전 피드백 섹션
' =====================================
Private Sub CreateFeedbackSection(ws As Worksheet, startCol As Integer)
    Dim startRow As Integer
    startRow = 5
    
    ' 섹션 헤더
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Merge
        .Value = "📋 " & GetLabel("PHASE1_FEEDBACK")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(41, 128, 185)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 최근 피드백 요약
    startRow = startRow + 2
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 3, startCol + 3))
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    
    ' 피드백 로드
    ws.Cells(startRow, startCol).Value = GetPreviousFeedback()
    
    ' 버튼들
    startRow = startRow + 5
    Dim btn As Object
    
    ' 피드백 조회 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol).Left, _
                             ws.Cells(startRow, startCol).Top, 110, 30)
    With btn
        .Caption = GetLabel("BTN_LOAD")
        .OnAction = "LoadPreviousFeedback"
        .Font.Size = 11
    End With
    
    ' 피드백 분석 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol + 2).Left, _
                             ws.Cells(startRow, startCol + 2).Top, 110, 30)
    With btn
        .Caption = GetLabel("BTN_ANALYZE")
        .OnAction = "AnalyzeFeedback"
        .Font.Size = 11
    End With
    
    ' 피드백 테이블
    startRow = startRow + 3
    Call CreateFeedbackTable(ws, startRow, startCol)
End Sub

' =====================================
' 자료 수집 섹션
' =====================================
Private Sub CreateCollectionSection(ws As Worksheet, startCol As Integer)
    Dim colOffset As Integer
    colOffset = 5  ' G열부터 시작
    startCol = startCol + colOffset
    
    Dim startRow As Integer
    startRow = 5
    
    ' 섹션 헤더
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Merge
        .Value = "📂 " & GetLabel("PHASE1_COLLECT")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(46, 204, 113)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 수집 현황 요약
    startRow = startRow + 2
    Call CreateCollectionSummary(ws, startRow, startCol)
    
    ' 버튼들
    startRow = startRow + 6
    Dim btn As Object
    
    ' 자료 수집 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol).Left, _
                             ws.Cells(startRow, startCol).Top, 110, 30)
    With btn
        .Caption = GetLabel("PHASE1_COLLECT")
        .OnAction = "CollectDocuments"
        .Font.Size = 11
    End With
    
    ' AI 요약 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol + 2).Left, _
                             ws.Cells(startRow, startCol + 2).Top, 110, 30)
    With btn
        .Caption = "AI " & GetLabel("PHASE2_SUMMARY")
        .OnAction = "GenerateAISummary"
        .Font.Size = 11
    End With
    
    ' 수집 문서 테이블
    startRow = startRow + 3
    Call CreateDocumentTable(ws, startRow, startCol)
End Sub

' =====================================
' 이슈 식별 섹션
' =====================================
Private Sub CreateIssueSection(ws As Worksheet, startCol As Integer)
    Dim colOffset As Integer
    colOffset = 10  ' L열부터 시작
    startCol = startCol + colOffset
    
    Dim startRow As Integer
    startRow = 5
    
    ' 섹션 헤더
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 2))
        .Merge
        .Value = "🎯 " & GetLabel("PHASE1_IDENTIFY")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(155, 89, 182)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' AI 분석 결과
    startRow = startRow + 2
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 3, startCol + 2))
        .Merge
        .Interior.Color = RGB(255, 250, 205)
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    
    ws.Cells(startRow, startCol).Value = GetAIInsights()
    
    ' 버튼
    startRow = startRow + 5
    Dim btn As Object
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol).Left, _
                             ws.Cells(startRow, startCol).Top, 150, 30)
    With btn
        .Caption = "AI " & GetLabel("PHASE1_IDENTIFY")
        .OnAction = "IdentifyKeyIssues"
        .Font.Size = 11
    End With
    
    ' 핵심 이슈 리스트
    startRow = startRow + 3
    Call CreateIssueList(ws, startRow, startCol)
End Sub

' =====================================
' 피드백 테이블 생성
' =====================================
Private Sub CreateFeedbackTable(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 테이블 헤더
    ws.Cells(startRow, startCol).Value = GetLabel("COL_DATE")
    ws.Cells(startRow, startCol + 1).Value = GetLabel("COL_TYPE")
    ws.Cells(startRow, startCol + 2).Value = GetLabel("PHASE3_FEEDBACK")
    ws.Cells(startRow, startCol + 3).Value = GetLabel("COL_STATUS")
    
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    ' 샘플 피드백 데이터
    Dim feedbacks As Variant
    feedbacks = Array( _
        Array("2025-07", "개선요청", "SK온 재무구조 구체화", "진행중"), _
        Array("2025-07", "지적사항", "경쟁사 대비 부족", "반영완료"), _
        Array("2025-07", "질문", "IRA 시나리오 추가", "대기") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(feedbacks)
        Dim row As Integer
        row = startRow + i + 1
        
        ws.Cells(row, startCol).Value = feedbacks(i)(0)
        ws.Cells(row, startCol + 1).Value = feedbacks(i)(1)
        ws.Cells(row, startCol + 2).Value = feedbacks(i)(2)
        ws.Cells(row, startCol + 3).Value = feedbacks(i)(3)
        
        ' 상태별 색상
        If feedbacks(i)(3) = "반영완료" Then
            ws.Cells(row, startCol + 3).Font.Color = COLOR_SUCCESS
        ElseIf feedbacks(i)(3) = "진행중" Then
            ws.Cells(row, startCol + 3).Font.Color = RGB(241, 196, 15)
        Else
            ws.Cells(row, startCol + 3).Font.Color = RGB(100, 100, 100)
        End If
        
        With ws.Range(ws.Cells(row, startCol), ws.Cells(row, startCol + 3))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' 수집 현황 요약
' =====================================
Private Sub CreateCollectionSummary(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 요약 테이블
    Dim categories As Variant
    categories = Array( _
        Array(GetLabel("PHASE1_INTERNAL"), "25", "2025-08-04"), _
        Array(GetLabel("PHASE1_EXTERNAL"), "127", "2025-08-04"), _
        Array(GetLabel("PHASE1_COMPETITOR"), "18", "2025-08-03"), _
        Array(GetLabel("PHASE1_POLICY"), "12", "2025-08-02") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(categories)
        ws.Cells(startRow + i, startCol).Value = categories(i)(0)
        ws.Cells(startRow + i, startCol + 1).Value = categories(i)(1) & GetLabel("METRIC_COUNT")
        ws.Cells(startRow + i, startCol + 2).Value = categories(i)(2)
        
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + 2))
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' 문서 테이블 생성
' =====================================
Private Sub CreateDocumentTable(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' Mock 데이터에서 최근 5개 문서 표시
    Dim reports As Collection
    Set reports = modMockData.InitializeMockData()
    
    ' 헤더
    ws.Cells(startRow, startCol).Value = GetLabel("COL_TITLE")
    ws.Cells(startRow, startCol + 1).Value = GetLabel("COL_SOURCE")
    ws.Cells(startRow, startCol + 2).Value = GetLabel("COL_DATE")
    ws.Cells(startRow, startCol + 3).Value = GetLabel("COL_RELEVANCE")
    
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    ' 데이터 표시 (최대 5개)
    Dim i As Integer
    Dim report As MockReport
    i = 0
    
    Dim item As Variant
    For Each item In reports
        If i >= 5 Then Exit For
        report = item
        i = i + 1
        
        ws.Cells(startRow + i, startCol).Value = Left(report.title, 20) & "..."
        ws.Cells(startRow + i, startCol + 1).Value = report.organization
        ws.Cells(startRow + i, startCol + 2).Value = report.reportDate
        ws.Cells(startRow + i, startCol + 3).Value = "95%"
        
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + 3))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
            .Font.Size = 10
        End With
    Next item
End Sub

' =====================================
' 이슈 리스트 생성
' =====================================
Private Sub CreateIssueList(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 핵심 이슈들
    Dim issues As Variant
    issues = Array( _
        "SK온 합병 시너지 실현", _
        "BYD 기술 격차 대응", _
        "IRA 정책 리스크", _
        "원자재 가격 변동", _
        "전고체 배터리 개발" _
    )
    
    Dim i As Integer
    For i = 0 To UBound(issues)
        With ws.Cells(startRow + i, startCol)
            .Value = "• " & issues(i)
            .Font.Size = 11
            .WrapText = True
        End With
        
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + 2))
            .Merge
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' 데이터 로드 함수들
' =====================================
Private Function GetPreviousFeedback() As String
    GetPreviousFeedback = "📌 7월 보고 피드백:" & vbLf & _
        "• SK온 재무구조 개선 계획 구체화 필요" & vbLf & _
        "• 전고체 배터리 경쟁사 대비 우위 분석" & vbLf & _
        "• IRA 정책 변화 시나리오별 대응안"
End Function

Private Function GetAIInsights() As String
    GetAIInsights = "🔍 AI 분석 결과:" & vbLf & _
        "• SK온-SK엔무브 합병 진행 상황" & vbLf & _
        "• BYD 5분 충전 기술 대응 필요" & vbLf & _
        "• IRA 정책 변화 리스크 증가"
End Function

' =====================================
' 액션 함수들
' =====================================
Sub LoadPreviousFeedback()
    Application.StatusBar = GetLabel("STATUS_LOADING")
    Application.Wait Now + TimeValue("00:00:01")
    
    ' 실제로는 DB나 파일에서 로드
    MsgBox "이전 피드백 3건을 불러왔습니다.", vbInformation, GetLabel("PHASE1_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub AnalyzeFeedback()
    Application.StatusBar = GetLabel("STATUS_ANALYZING")
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox "피드백 분석 완료:" & vbLf & vbLf & _
           "• 반영 완료: 5건" & vbLf & _
           "• 진행 중: 3건" & vbLf & _
           "• 대기: 2건", _
           vbInformation, GetLabel("PHASE1_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub CollectDocuments()
    Application.StatusBar = GetLabel("STATUS_PROCESSING")
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox GetLabel("PHASE1_COLLECT") & " 완료:" & vbLf & vbLf & _
           "• " & GetLabel("PHASE1_INTERNAL") & ": 25건" & vbLf & _
           "• " & GetLabel("PHASE1_EXTERNAL") & ": 127건" & vbLf & _
           "• " & GetLabel("PHASE1_COMPETITOR") & ": 18건" & vbLf & _
           "• " & GetLabel("PHASE1_POLICY") & ": 12건" & vbLf & vbLf & _
           "총 182건", _
           vbInformation, GetLabel("PHASE1_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub GenerateAISummary()
    Application.StatusBar = GetLabel("STATUS_GENERATING")
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox "AI 요약 생성 완료:" & vbLf & vbLf & _
           "주요 내용:" & vbLf & _
           "• SK온 합병으로 시너지 20조원 예상" & vbLf & _
           "• BYD 초급속 충전 기술 위협" & vbLf & _
           "• IRA 정책 불확실성 증가" & vbLf & _
           "• K배터리 점유율 회복 조짐", _
           vbInformation, GetLabel("PHASE1_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub IdentifyKeyIssues()
    Application.StatusBar = GetLabel("STATUS_ANALYZING")
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox "핵심 이슈 식별 완료:" & vbLf & vbLf & _
           "Critical (즉시 대응):" & vbLf & _
           "• SK온 합병 준비" & vbLf & _
           "• BYD 기술 대응" & vbLf & vbLf & _
           "High (단기 대응):" & vbLf & _
           "• IRA 정책 시나리오" & vbLf & _
           "• 원자재 가격 헤징", _
           vbInformation, GetLabel("PHASE1_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

' =====================================
' Phase 1 실행 (메인에서 호출)
' =====================================
Sub Phase1_Execute()
    ' Phase 1 시트 생성
    Call CreatePhase1Sheet
    
    ' 자동 실행 순서
    Call LoadPreviousFeedback
    Call CollectDocuments
    Call IdentifyKeyIssues
    
    MsgBox GetLabel("PHASE1_TITLE") & " " & GetLabel("STATUS_COMPLETE"), _
           vbInformation, APP_NAME
End Sub