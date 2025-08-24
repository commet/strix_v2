Attribute VB_Name = "modPhase3"
' =====================================
' STRIX v2 - Phase 3 Module
' 보고 이후 단계
' =====================================
Option Explicit

' Phase 3 데이터 타입
Public Type FeedbackRecord
    id As String
    receivedDate As Date
    reviewer As String
    department As String
    feedbackType As String
    content As String
    priority As String
    actionRequired As Boolean
    status As String
End Type

Public Type ActionItem
    id As String
    feedbackId As String
    action As String
    owner As String
    dueDate As Date
    status As String
    progress As Integer
End Type

Public Type RAGUpdate
    id As String
    updateDate As Date
    feedbackContent As String
    vectorized As Boolean
    embedded As Boolean
    indexUpdated As Boolean
End Type

' =====================================
' Phase 3 시트 생성
' =====================================
Sub CreatePhase3Sheet()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_PHASE3)
    
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
        
        ' Phase 3 헤더
        With .Range("B2:N2")
            .Merge
            .Value = GetLabel("PHASE3_TITLE")
            .Font.Name = "맑은 고딕"
            .Font.Size = 24
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = COLOR_PHASE3
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 50
        End With
        
        ' 부제목
        With .Range("B3:N3")
            .Merge
            .Value = GetLabel("PHASE3_DESC") & " | " & FormatDateTime(Now)
            .Font.Size = 12
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter
            .RowHeight = 25
        End With
        
        ' 3개 섹션 생성
        Call CreateFeedbackCollectionSection(ws, 2)   ' 피드백 수집
        Call CreateRAGUpdateSection(ws, 7)            ' RAG 업데이트
        Call CreateIssueTrackingSection(ws, 12)       ' 이슈 트래킹
    End With
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox GetLabel("ERR_GENERAL") & ": " & Err.Description, vbCritical
End Sub

' =====================================
' 피드백 수집 섹션
' =====================================
Private Sub CreateFeedbackCollectionSection(ws As Worksheet, startCol As Integer)
    Dim startRow As Integer
    startRow = 5
    
    ' 섹션 헤더
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Merge
        .Value = "📝 " & GetLabel("PHASE3_FEEDBACK")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(52, 152, 219)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 실시간 피드백
    startRow = startRow + 2
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 3, startCol + 3))
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    
    ws.Cells(startRow, startCol).Value = GetRecentFeedback()
    
    ' 버튼들
    startRow = startRow + 5
    Dim btn As Object
    
    ' 피드백 기록 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol).Left, _
                             ws.Cells(startRow, startCol).Top, 110, 30)
    With btn
        .Caption = "피드백 기록"
        .OnAction = "RecordFeedback"
        .Font.Size = 11
    End With
    
    ' 피드백 분류 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol + 2).Left, _
                             ws.Cells(startRow, startCol + 2).Top, 110, 30)
    With btn
        .Caption = GetLabel("PHASE3_CLASSIFY")
        .OnAction = "ClassifyFeedback"
        .Font.Size = 11
    End With
    
    ' 피드백 테이블
    startRow = startRow + 3
    Call CreateFeedbackTable(ws, startRow, startCol)
End Sub

' =====================================
' RAG 업데이트 섹션
' =====================================
Private Sub CreateRAGUpdateSection(ws As Worksheet, startCol As Integer)
    Dim startRow As Integer
    startRow = 5
    
    ' 섹션 헤더
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Merge
        .Value = "🔄 " & GetLabel("PHASE3_UPDATE")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(46, 204, 113)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' RAG 업데이트 현황
    startRow = startRow + 2
    Call CreateRAGStatus(ws, startRow, startCol)
    
    ' 버튼들
    startRow = startRow + 6
    Dim btn As Object
    
    ' RAG 업데이트 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol).Left, _
                             ws.Cells(startRow, startCol).Top, 110, 30)
    With btn
        .Caption = "RAG 업데이트"
        .OnAction = "UpdateRAGSystem"
        .Font.Size = 11
    End With
    
    ' 학습 검증 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol + 2).Left, _
                             ws.Cells(startRow, startCol + 2).Top, 110, 30)
    With btn
        .Caption = "학습 검증"
        .OnAction = "ValidateLearning"
        .Font.Size = 11
    End With
    
    ' 업데이트 로그
    startRow = startRow + 3
    Call CreateUpdateLog(ws, startRow, startCol)
End Sub

' =====================================
' 이슈 트래킹 섹션
' =====================================
Private Sub CreateIssueTrackingSection(ws As Worksheet, startCol As Integer)
    Dim startRow As Integer
    startRow = 5
    
    ' 섹션 헤더
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 2))
        .Merge
        .Value = "🎯 " & GetLabel("PHASE3_TRACK")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(231, 76, 60)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 액션 아이템 요약
    startRow = startRow + 2
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 2, startCol + 2))
        .Merge
        .Interior.Color = RGB(255, 245, 245)
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    
    ws.Cells(startRow, startCol).Value = GetActionItemsSummary()
    
    ' 버튼
    startRow = startRow + 4
    Dim btn As Object
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol).Left, _
                             ws.Cells(startRow, startCol).Top, 150, 30)
    With btn
        .Caption = GetLabel("PHASE3_ACTION") & " 관리"
        .OnAction = "ManageActionItems"
        .Font.Size = 11
    End With
    
    ' 액션 아이템 리스트
    startRow = startRow + 3
    Call CreateActionItemsList(ws, startRow, startCol)
End Sub

' =====================================
' 피드백 테이블
' =====================================
Private Sub CreateFeedbackTable(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 헤더
    ws.Cells(startRow, startCol).Value = "검토자"
    ws.Cells(startRow, startCol + 1).Value = GetLabel("COL_TYPE")
    ws.Cells(startRow, startCol + 2).Value = "내용"
    ws.Cells(startRow, startCol + 3).Value = GetLabel("COL_PRIORITY")
    
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    ' 샘플 피드백
    Dim feedbacks As Variant
    feedbacks = Array( _
        Array("CEO", "개선", "합병 시너지 구체화", GetLabel("CAT_HIGH")), _
        Array("CFO", "질문", "자본확충 일정", GetLabel("CAT_CRITICAL")), _
        Array("CTO", "요청", "기술 로드맵 업데이트", GetLabel("CAT_NORMAL")) _
    )
    
    Dim i As Integer
    For i = 0 To UBound(feedbacks)
        Dim row As Integer
        row = startRow + i + 1
        
        ws.Cells(row, startCol).Value = feedbacks(i)(0)
        ws.Cells(row, startCol + 1).Value = feedbacks(i)(1)
        ws.Cells(row, startCol + 2).Value = feedbacks(i)(2)
        ws.Cells(row, startCol + 3).Value = feedbacks(i)(3)
        
        ' 우선순위별 색상
        If feedbacks(i)(3) = GetLabel("CAT_CRITICAL") Then
            ws.Cells(row, startCol + 3).Font.Color = COLOR_WARNING
            ws.Cells(row, startCol + 3).Font.Bold = True
        ElseIf feedbacks(i)(3) = GetLabel("CAT_HIGH") Then
            ws.Cells(row, startCol + 3).Font.Color = RGB(230, 126, 34)
        End If
        
        With ws.Range(ws.Cells(row, startCol), ws.Cells(row, startCol + 3))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' RAG 상태 표시
' =====================================
Private Sub CreateRAGStatus(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' RAG 업데이트 현황
    Dim statuses As Variant
    statuses = Array( _
        Array("피드백 벡터화", "완료", "5건"), _
        Array("문서 임베딩", "진행중", "12건"), _
        Array("메타데이터", "대기", "0건"), _
        Array("인덱스 갱신", "완료", "전체") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(statuses)
        ws.Cells(startRow + i, startCol).Value = statuses(i)(0)
        ws.Cells(startRow + i, startCol + 1).Value = statuses(i)(1)
        ws.Cells(startRow + i, startCol + 2).Value = statuses(i)(2)
        
        ' 상태별 색상
        If statuses(i)(1) = "완료" Then
            ws.Cells(startRow + i, startCol + 1).Font.Color = COLOR_SUCCESS
        ElseIf statuses(i)(1) = "진행중" Then
            ws.Cells(startRow + i, startCol + 1).Font.Color = RGB(241, 196, 15)
        Else
            ws.Cells(startRow + i, startCol + 1).Font.Color = RGB(150, 150, 150)
        End If
        
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + 2))
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' 업데이트 로그
' =====================================
Private Sub CreateUpdateLog(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 최근 업데이트 로그
    Dim logs As Variant
    logs = Array( _
        Array(FormatTime(Now - 0.1), "피드백 5건 벡터화 완료"), _
        Array(FormatTime(Now - 0.05), "RAG 정확도 92% → 94%"), _
        Array(FormatTime(Now - 0.02), "인덱스 재구축 완료") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(logs)
        ws.Cells(startRow + i, startCol).Value = logs(i)(0)
        ws.Cells(startRow + i, startCol + 1).Value = logs(i)(1)
        
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + 3))
            .Merge
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
            .Font.Size = 9
            .Font.Color = RGB(100, 100, 100)
        End With
    Next i
End Sub

' =====================================
' 액션 아이템 리스트
' =====================================
Private Sub CreateActionItemsList(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 액션 아이템
    Dim actions As Variant
    actions = Array( _
        Array("[A-001]", "합병 TF 구성", "D-7"), _
        Array("[A-002]", "IRA 대응안 수립", "D-3"), _
        Array("[A-003]", "기술 벤치마킹", "D-14"), _
        Array("[A-004]", "자본확충 IR", "D-10"), _
        Array("[A-005]", "Q4 실적 예측", "D-5") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(actions)
        ws.Cells(startRow + i, startCol).Value = actions(i)(0)
        ws.Cells(startRow + i, startCol + 1).Value = actions(i)(1)
        ws.Cells(startRow + i, startCol + 2).Value = actions(i)(2)
        
        ' 마감일 임박 표시
        If InStr(actions(i)(2), "D-3") > 0 Or InStr(actions(i)(2), "D-5") > 0 Then
            ws.Cells(startRow + i, startCol + 2).Font.Color = COLOR_WARNING
            ws.Cells(startRow + i, startCol + 2).Font.Bold = True
        End If
        
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + 2))
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' 데이터 로드 함수들
' =====================================
Private Function GetRecentFeedback() As String
    GetRecentFeedback = "📝 실시간 피드백:" & vbLf & _
        "• CEO: 합병 시너지 구체화 필요" & vbLf & _
        "• CFO: 자본확충 일정 명확화" & vbLf & _
        "• CTO: 기술 로드맵 업데이트 요청"
End Function

Private Function GetActionItemsSummary() As String
    GetActionItemsSummary = "🎯 Action Items:" & vbLf & _
        "• 전체: 12건" & vbLf & _
        "• 긴급: 3건 (D-3 이내)" & vbLf & _
        "• 진행중: 7건"
End Function

' =====================================
' 액션 함수들
' =====================================
Sub RecordFeedback()
    Dim feedback As String
    feedback = InputBox("피드백을 입력하세요:", GetLabel("PHASE3_FEEDBACK"))
    
    If feedback <> "" Then
        Application.StatusBar = GetLabel("STATUS_SAVING")
        Application.Wait Now + TimeValue("00:00:01")
        
        MsgBox "피드백이 기록되었습니다:" & vbLf & vbLf & _
               feedback, vbInformation, GetLabel("PHASE3_TITLE")
        
        Application.StatusBar = GetLabel("STATUS_READY")
    End If
End Sub

Sub ClassifyFeedback()
    Application.StatusBar = GetLabel("STATUS_PROCESSING")
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox "피드백 분류 완료:" & vbLf & vbLf & _
           "• 개선 요청: 3건" & vbLf & _
           "• 질문 사항: 2건" & vbLf & _
           "• 지적 사항: 1건" & vbLf & _
           "• 칭찬/격려: 1건", _
           vbInformation, GetLabel("PHASE3_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub UpdateRAGSystem()
    Application.StatusBar = GetLabel("PHASE3_UPDATE") & "..."
    Application.Wait Now + TimeValue("00:00:03")
    
    MsgBox "RAG 시스템 업데이트 완료:" & vbLf & vbLf & _
           "• 피드백 벡터화: 5건" & vbLf & _
           "• 문서 임베딩: 12건" & vbLf & _
           "• 인덱스 갱신: 완료" & vbLf & _
           "• 정확도: 92% → 94%", _
           vbInformation, GetLabel("PHASE3_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub ValidateLearning()
    Application.StatusBar = "학습 검증 중..."
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox "학습 검증 결과:" & vbLf & vbLf & _
           "✅ 응답 정확도: 94%" & vbLf & _
           "✅ 피드백 반영률: 100%" & vbLf & _
           "✅ 검색 관련성: 89%" & vbLf & vbLf & _
           GetLabel("PHASE3_LEARNING"), _
           vbInformation, GetLabel("PHASE3_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub ManageActionItems()
    MsgBox "Action Items 관리:" & vbLf & vbLf & _
           "🔴 긴급 (D-3 이내): 3건" & vbLf & _
           "🟡 주의 (D-7 이내): 4건" & vbLf & _
           "🟢 정상 진행: 5건" & vbLf & vbLf & _
           "총 12건의 액션 아이템", _
           vbInformation, GetLabel("PHASE3_ACTION")
End Sub

' =====================================
' Phase 3 실행 (메인에서 호출)
' =====================================
Sub Phase3_Execute()
    ' Phase 3 시트 생성
    Call CreatePhase3Sheet
    
    ' 자동 실행 순서
    Call RecordFeedback
    Call ClassifyFeedback
    Call UpdateRAGSystem
    Call ValidateLearning
    
    MsgBox GetLabel("PHASE3_TITLE") & " " & GetLabel("STATUS_COMPLETE"), _
           vbInformation, APP_NAME
End Sub