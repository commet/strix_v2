Attribute VB_Name = "modPhase2"
' =====================================
' STRIX v2 - Phase 2 Module
' 보고 준비 단계
' =====================================
Option Explicit

' Phase 2 데이터 타입
Public Type ConsolidatedData
    id As String
    category As String
    keyFindings As String
    dataPoints As Collection
    confidence As Double
    source As String
End Type

Public Type ReportSection
    sectionName As String
    content As String
    status As String  ' 작성중/검토중/완료
    lastUpdate As Date
    wordCount As Long
End Type

Public Type Insight
    id As String
    category As String
    insight As String
    impact As String  ' High/Medium/Low
    recommendation As String
    confidence As Double
End Type

' =====================================
' Phase 2 시트 생성
' =====================================
Sub CreatePhase2Sheet()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_PHASE2)
    
    With ws
        .Cells.Clear
        .Cells.Interior.Color = COLOR_BACKGROUND
        
        ' 열 너비 설정
        .Columns("A").ColumnWidth = 2
        .Columns("B:E").ColumnWidth = 22
        .Columns("F").ColumnWidth = 3
        .Columns("G:J").ColumnWidth = 22
        .Columns("K").ColumnWidth = 3
        .Columns("L:N").ColumnWidth = 25
        .Columns("O").ColumnWidth = 2
        
        ' Phase 2 헤더
        With .Range("B2:N2")
            .Merge
            .Value = GetLabel("PHASE2_TITLE")
            .Font.Name = "맑은 고딕"
            .Font.Size = 24
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = COLOR_PHASE2
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 50
        End With
        
        ' 부제목
        With .Range("B3:N3")
            .Merge
            .Value = GetLabel("PHASE2_DESC") & " | " & FormatDateTime(Now)
            .Font.Size = 12
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlCenter
            .RowHeight = 25
        End With
        
        ' 3개 섹션 생성
        Call CreateConsolidationSection(ws, 5)   ' 자료 종합
        Call CreateReportSection(ws, 5)          ' 보고서 작성
        Call CreateInsightSection(ws, 5)         ' 핵심 인사이트
    End With
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox GetLabel("ERR_GENERAL") & ": " & Err.Description, vbCritical
End Sub

' =====================================
' 자료 종합 섹션
' =====================================
Private Sub CreateConsolidationSection(ws As Worksheet, startCol As Integer)
    Dim startRow As Integer
    startRow = 5
    
    ' 섹션 헤더
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Merge
        .Value = "📊 " & GetLabel("PHASE2_CONSOLIDATE")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(52, 152, 219)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 종합 분석 결과
    startRow = startRow + 2
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 3, startCol + 3))
        .Merge
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    
    ws.Cells(startRow, startCol).Value = GetConsolidationSummary()
    
    ' 버튼들
    startRow = startRow + 5
    Dim btn As Object
    
    ' 자료 종합 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol).Left, _
                             ws.Cells(startRow, startCol).Top, 110, 30)
    With btn
        .Caption = GetLabel("PHASE2_CONSOLIDATE")
        .OnAction = "ConsolidateData"
        .Font.Size = 11
    End With
    
    ' AI 분석 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol + 2).Left, _
                             ws.Cells(startRow, startCol + 2).Top, 110, 30)
    With btn
        .Caption = GetLabel("PHASE2_ANALYZE")
        .OnAction = "RunAIAnalysis"
        .Font.Size = 11
    End With
    
    ' 종합 데이터 테이블
    startRow = startRow + 3
    Call CreateDataTable(ws, startRow, startCol)
End Sub

' =====================================
' 보고서 작성 섹션
' =====================================
Private Sub CreateReportSection(ws As Worksheet, startCol As Integer)
    Dim colOffset As Integer
    colOffset = 5  ' G열부터 시작
    startCol = startCol + colOffset
    
    Dim startRow As Integer
    startRow = 5
    
    ' 섹션 헤더
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Merge
        .Value = "📝 " & GetLabel("PHASE2_WRITE")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(241, 196, 15)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 보고서 작성 현황
    startRow = startRow + 2
    Call CreateReportProgress(ws, startRow, startCol)
    
    ' 버튼들
    startRow = startRow + 7
    Dim btn As Object
    
    ' 보고서 생성 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol).Left, _
                             ws.Cells(startRow, startCol).Top, 110, 30)
    With btn
        .Caption = GetLabel("BTN_GENERATE")
        .OnAction = "GenerateReport"
        .Font.Size = 11
    End With
    
    ' 미리보기 버튼
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol + 2).Left, _
                             ws.Cells(startRow, startCol + 2).Top, 110, 30)
    With btn
        .Caption = GetLabel("BTN_PREVIEW")
        .OnAction = "PreviewReport"
        .Font.Size = 11
    End With
    
    ' 보고서 섹션 체크리스트
    startRow = startRow + 3
    Call CreateReportChecklist(ws, startRow, startCol)
End Sub

' =====================================
' 핵심 인사이트 섹션
' =====================================
Private Sub CreateInsightSection(ws As Worksheet, startCol As Integer)
    Dim colOffset As Integer
    colOffset = 10  ' L열부터 시작
    startCol = startCol + colOffset
    
    Dim startRow As Integer
    startRow = 5
    
    ' 섹션 헤더
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 2))
        .Merge
        .Value = "💡 " & GetLabel("PHASE2_INSIGHT")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(155, 89, 182)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 핵심 인사이트
    startRow = startRow + 2
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 4, startCol + 2))
        .Merge
        .Interior.Color = RGB(255, 250, 205)
        .Borders.LineStyle = xlContinuous
        .WrapText = True
        .VerticalAlignment = xlTop
        .Font.Bold = True
    End With
    
    ws.Cells(startRow, startCol).Value = GetKeyInsights()
    
    ' 버튼
    startRow = startRow + 6
    Dim btn As Object
    Set btn = ws.Buttons.Add(ws.Cells(startRow, startCol).Left, _
                             ws.Cells(startRow, startCol).Top, 150, 30)
    With btn
        .Caption = GetLabel("PHASE2_INSIGHT") & " " & GetLabel("BTN_GENERATE")
        .OnAction = "GenerateInsights"
        .Font.Size = 11
    End With
    
    ' 인사이트 리스트
    startRow = startRow + 3
    Call CreateInsightsList(ws, startRow, startCol)
End Sub

' =====================================
' 데이터 테이블 생성
' =====================================
Private Sub CreateDataTable(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 테이블 헤더
    ws.Cells(startRow, startCol).Value = GetLabel("COL_CATEGORY")
    ws.Cells(startRow, startCol + 1).Value = "핵심 발견"
    ws.Cells(startRow, startCol + 2).Value = "데이터 수"
    ws.Cells(startRow, startCol + 3).Value = "신뢰도"
    
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 230)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
    End With
    
    ' 종합 데이터
    Dim consolidatedData As Variant
    consolidatedData = Array( _
        Array(GetLabel("PHASE2_TREND"), "CATL 점유율 40% 돌파", "45", "95%"), _
        Array(GetLabel("PHASE2_RISK"), "IRA 정책 변경 임박", "23", "88%"), _
        Array("기술", "전고체 2027년 양산", "18", "92%"), _
        Array("재무", "4Q 흑자 전환 예상", "12", "85%") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(consolidatedData)
        Dim row As Integer
        row = startRow + i + 1
        
        ws.Cells(row, startCol).Value = consolidatedData(i)(0)
        ws.Cells(row, startCol + 1).Value = consolidatedData(i)(1)
        ws.Cells(row, startCol + 2).Value = consolidatedData(i)(2)
        ws.Cells(row, startCol + 3).Value = consolidatedData(i)(3)
        
        With ws.Range(ws.Cells(row, startCol), ws.Cells(row, startCol + 3))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' 보고서 진행 현황
' =====================================
Private Sub CreateReportProgress(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 진행률 바
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 3))
        .Merge
        .Value = "전체 진행률: 75%"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    ' 진행률 시각화
    startRow = startRow + 1
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 2))
        .Merge
        .Interior.Color = COLOR_SUCCESS
    End With
    With ws.Cells(startRow, startCol + 3)
        .Interior.Color = RGB(230, 230, 230)
    End With
    
    ' 섹션별 상태
    startRow = startRow + 2
    Dim sections As Variant
    sections = Array( _
        Array("Executive Summary", "완료", COLOR_SUCCESS), _
        Array(GetLabel("PHASE2_TREND"), "완료", COLOR_SUCCESS), _
        Array("경쟁사 현황", "완료", COLOR_SUCCESS), _
        Array(GetLabel("PHASE2_RISK"), "작성중", RGB(241, 196, 15)), _
        Array(GetLabel("PHASE2_STRATEGY"), "대기", RGB(200, 200, 200)) _
    )
    
    For i = 0 To UBound(sections)
        ws.Cells(startRow + i, startCol).Value = sections(i)(0)
        ws.Cells(startRow + i, startCol + 1).Value = sections(i)(1)
        
        ' 상태별 색상
        ws.Cells(startRow + i, startCol + 1).Font.Color = sections(i)(2)
        ws.Cells(startRow + i, startCol + 1).Font.Bold = True
        
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + 1))
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' 보고서 체크리스트
' =====================================
Private Sub CreateReportChecklist(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 체크리스트 항목
    Dim checklist As Variant
    checklist = Array( _
        Array("✅", "데이터 검증 완료"), _
        Array("✅", "그래프/차트 생성"), _
        Array("✅", "핵심 메시지 정리"), _
        Array("⏳", "임원 검토 반영"), _
        Array("⏳", "최종 교정") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(checklist)
        ws.Cells(startRow + i, startCol).Value = checklist(i)(0)
        ws.Cells(startRow + i, startCol + 1).Value = checklist(i)(1)
        
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + 3))
            .Merge
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' 인사이트 리스트
' =====================================
Private Sub CreateInsightsList(ws As Worksheet, startRow As Integer, startCol As Integer)
    ' 인사이트 카테고리별 표시
    Dim insights As Variant
    insights = Array( _
        Array("전략", "SK온 합병 시너지 20조"), _
        Array("기술", "전고체 3년내 양산 가능"), _
        Array("시장", "중국 점유율 60% 돌파"), _
        Array("리스크", "IRA 폐지시 -2000억/년"), _
        Array("기회", "ESS 시장 300% 성장") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(insights)
        With ws.Cells(startRow + i, startCol)
            .Value = insights(i)(0)
            .Font.Bold = True
            .Font.Color = COLOR_INFO
        End With
        
        ws.Cells(startRow + i, startCol + 1).Value = insights(i)(1)
        
        With ws.Range(ws.Cells(startRow + i, startCol), ws.Cells(startRow + i, startCol + 2))
            .Merge
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
            .WrapText = True
        End With
    Next i
End Sub

' =====================================
' 데이터 로드 함수들
' =====================================
Private Function GetConsolidationSummary() As String
    GetConsolidationSummary = "📊 종합 분석 결과:" & vbLf & _
        "• 내외부 자료 182건 통합 완료" & vbLf & _
        "• 핵심 트렌드 7개 도출" & vbLf & _
        "• 리스크 요인 5개 식별" & vbLf & _
        "• 기회 요인 3개 발굴"
End Function

Private Function GetKeyInsights() As String
    GetKeyInsights = "💡 핵심 인사이트:" & vbLf & vbLf & _
        "1. SK온 합병 시너지 20조원" & vbLf & _
        "2. 전고체 2027년 양산 가능" & vbLf & _
        "3. 중국 대응 전략 시급"
End Function

' =====================================
' 액션 함수들
' =====================================
Sub ConsolidateData()
    Application.StatusBar = GetLabel("STATUS_PROCESSING")
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox GetLabel("PHASE2_CONSOLIDATE") & " 완료:" & vbLf & vbLf & _
           "• 총 182건 자료 분석" & vbLf & _
           "• 중복 제거: 23건" & vbLf & _
           "• 핵심 자료: 98건" & vbLf & _
           "• 보조 자료: 61건", _
           vbInformation, GetLabel("PHASE2_TITLE")
    
    ' Phase 2 시트 업데이트
    Call UpdateConsolidationStatus
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub RunAIAnalysis()
    Application.StatusBar = GetLabel("STATUS_ANALYZING")
    Application.Wait Now + TimeValue("00:00:03")
    
    MsgBox "AI 분석 완료:" & vbLf & vbLf & _
           "📈 시장 전망:" & vbLf & _
           "• 2025년 배터리 수요 500GWh" & vbLf & _
           "• 중국 점유율 65% 예상" & vbLf & vbLf & _
           "⚠️ 주요 리스크:" & vbLf & _
           "• IRA 정책 변경 (확률 70%)" & vbLf & _
           "• 원자재 가격 상승 지속", _
           vbInformation, GetLabel("PHASE2_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub GenerateReport()
    Application.StatusBar = GetLabel("STATUS_GENERATING")
    Application.Wait Now + TimeValue("00:00:03")
    
    ' 보고서 생성 시뮬레이션
    Dim reportName As String
    reportName = "STRIX_Report_" & Format(Date, "yyyymmdd") & ".xlsx"
    
    MsgBox GetLabel("PHASE2_WRITE") & " 완료:" & vbLf & vbLf & _
           "파일명: " & reportName & vbLf & _
           "페이지: 25p" & vbLf & _
           "차트: 12개" & vbLf & _
           "표: 8개" & vbLf & vbLf & _
           "저장 위치: Documents\Reports\", _
           vbInformation, GetLabel("PHASE2_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

Sub PreviewReport()
    MsgBox "보고서 미리보기:" & vbLf & vbLf & _
           "1. Executive Summary" & vbLf & _
           "2. 시장 동향 분석" & vbLf & _
           "3. 경쟁사 현황" & vbLf & _
           "4. 기술 개발 현황" & vbLf & _
           "5. 리스크 평가" & vbLf & _
           "6. 전략 제언" & vbLf & _
           "7. Appendix", _
           vbInformation, GetLabel("BTN_PREVIEW")
End Sub

Sub GenerateInsights()
    Application.StatusBar = GetLabel("STATUS_GENERATING")
    Application.Wait Now + TimeValue("00:00:02")
    
    MsgBox GetLabel("PHASE2_INSIGHT") & " 도출:" & vbLf & vbLf & _
           "🎯 전략적 시사점:" & vbLf & _
           "• SK온 합병은 생존 필수 전략" & vbLf & _
           "• 전고체 선점이 게임 체인저" & vbLf & _
           "• 중국 시장 재진입 검토 필요" & vbLf & vbLf & _
           "💰 재무적 영향:" & vbLf & _
           "• 합병 시너지: +20조원 (5년)" & vbLf & _
           "• IRA 리스크: -2,000억/년", _
           vbInformation, GetLabel("PHASE2_TITLE")
    
    Application.StatusBar = GetLabel("STATUS_READY")
End Sub

' =====================================
' 상태 업데이트
' =====================================
Private Sub UpdateConsolidationStatus()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_PHASE2)
    
    If Not ws Is Nothing Then
        ' 진행률 업데이트 등
        ws.Range("G7").Value = "📊 종합 분석 완료 ✅"
    End If
End Sub

' =====================================
' Phase 2 실행 (메인에서 호출)
' =====================================
Sub Phase2_Execute()
    ' Phase 2 시트 생성
    Call CreatePhase2Sheet
    
    ' 자동 실행 순서
    Call ConsolidateData
    Call RunAIAnalysis
    Call GenerateReport
    Call GenerateInsights
    
    MsgBox GetLabel("PHASE2_TITLE") & " " & GetLabel("STATUS_COMPLETE"), _
           vbInformation, APP_NAME
End Sub