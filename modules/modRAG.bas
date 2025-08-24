Attribute VB_Name = "modRAG"
' =====================================
' STRIX v2 - RAG Integration Module
' API 서버 연동 및 RAG 처리
' =====================================
Option Explicit

' API 응답 타입
Public Type RAGResponse
    success As Boolean
    answer As String
    confidence As Double
    sources As Collection
    internalCount As Integer
    externalCount As Integer
    errorMessage As String
    responseTime As Double
End Type

' 소스 문서 타입
Public Type SourceDoc
    id As String
    title As String
    content As String
    docType As String
    organization As String
    docDate As String
    relevance As Double
End Type

' =====================================
' RAG API 호출 (UTF-8 안전)
' =====================================
Public Function CallRAGAPI(question As String, Optional docType As String = "both") As RAGResponse
    Dim response As RAGResponse
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim startTime As Double
    
    On Error GoTo ErrorHandler
    
    response.success = False
    startTime = Timer
    
    ' HTTP 객체 생성
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' API URL
    url = GetAPIUrl("query")
    
    ' JSON 요청 생성 (UTF-8 안전)
    requestBody = CreateRAGRequest(question, docType)
    
    ' API 호출
    With http
        .Open "POST", url, False
        .setTimeouts API_TIMEOUT, API_TIMEOUT, API_TIMEOUT, API_TIMEOUT
        
        ' UTF-8 헤더 설정
        Call modUTF8.SetUTF8Headers(http)
        
        ' 요청 전송
        .send modUTF8.StringToUTF8Bytes(requestBody)
        
        ' 응답 처리
        If .Status = 200 Then
            response = ParseRAGResponse(.responseBody)
            response.success = True
        Else
            response.errorMessage = "API Error: " & .Status & " - " & .statusText
        End If
    End With
    
    response.responseTime = Timer - startTime
    CallRAGAPI = response
    Exit Function
    
ErrorHandler:
    response.errorMessage = "Error: " & Err.Description
    response.success = False
    CallRAGAPI = response
End Function

' =====================================
' RAG 요청 JSON 생성
' =====================================
Private Function CreateRAGRequest(question As String, docType As String) As String
    Dim json As String
    
    json = "{"
    json = json & """question"":""" & modUTF8.EscapeJSON(question) & ""","
    json = json & """doc_type"":""" & docType & ""","
    json = json & """max_results"":10,"
    json = json & """include_sources"":true"
    json = json & "}"
    
    CreateRAGRequest = json
End Function

' =====================================
' RAG 응답 파싱
' =====================================
Private Function ParseRAGResponse(responseBody As Variant) As RAGResponse
    Dim response As RAGResponse
    Dim jsonText As String
    
    On Error GoTo ParseError
    
    ' UTF-8 디코딩
    jsonText = modUTF8.DecodeUTF8Response(responseBody)
    
    ' 간단한 JSON 파싱 (실제로는 JSON 파서 사용 권장)
    response.answer = ExtractJSONValue(jsonText, "answer")
    response.confidence = CDbl(ExtractJSONValue(jsonText, "confidence", "0.0"))
    response.internalCount = CInt(ExtractJSONValue(jsonText, "internal_docs", "0"))
    response.externalCount = CInt(ExtractJSONValue(jsonText, "external_docs", "0"))
    
    ' 소스 문서 파싱
    Set response.sources = ParseSources(jsonText)
    
    ParseRAGResponse = response
    Exit Function
    
ParseError:
    response.errorMessage = "Parse error: " & Err.Description
    ParseRAGResponse = response
End Function

' =====================================
' JSON 값 추출 (간단한 파서)
' =====================================
Private Function ExtractJSONValue(json As String, key As String, Optional defaultValue As String = "") As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String
    
    startPos = InStr(json, """" & key & """:")
    If startPos = 0 Then
        ExtractJSONValue = defaultValue
        Exit Function
    End If
    
    startPos = InStr(startPos, json, ":") + 1
    
    ' 값이 문자열인 경우
    If Mid(json, startPos + 1, 1) = """" Then
        startPos = startPos + 2
        endPos = InStr(startPos, json, """")
        value = Mid(json, startPos, endPos - startPos)
    Else
        ' 값이 숫자인 경우
        endPos = InStr(startPos, json, ",")
        If endPos = 0 Then endPos = InStr(startPos, json, "}")
        value = Trim(Mid(json, startPos, endPos - startPos))
    End If
    
    ExtractJSONValue = value
End Function

' =====================================
' 소스 문서 파싱
' =====================================
Private Function ParseSources(json As String) As Collection
    Dim sources As New Collection
    Dim source As SourceDoc
    
    ' 간단한 구현 - 실제로는 JSON 파서 사용
    ' Mock 데이터로 대체
    Dim mockReports As Collection
    Set mockReports = modMockData.InitializeMockData()
    
    Dim i As Integer
    Dim report As MockReport
    For i = 1 To 3
        If i <= mockReports.Count Then
            report = mockReports(i)
            source.id = report.id
            source.title = report.title
            source.organization = report.organization
            source.docDate = report.reportDate
            source.docType = report.docType
            source.relevance = 0.95
            sources.Add source
        End If
    Next i
    
    Set ParseSources = sources
End Function

' =====================================
' Excel 시트에서 RAG 실행
' =====================================
Sub RunRAGSearch()
    Dim ws As Worksheet
    Dim question As String
    Dim response As RAGResponse
    
    On Error GoTo ErrorHandler
    
    ' 현재 시트 확인
    Set ws = ActiveSheet
    
    ' 질문 입력 받기
    question = InputBox("질문을 입력하세요:", "STRIX RAG Search")
    If question = "" Then Exit Sub
    
    ' 상태 표시
    Application.StatusBar = GetLabel("STATUS_PROCESSING") & " RAG API..."
    
    ' RAG API 호출
    response = CallRAGAPI(question)
    
    ' 결과 처리
    If response.success Then
        Call DisplayRAGResults(ws, response)
        MsgBox "검색 완료!" & vbLf & vbLf & _
               "응답 시간: " & Format(response.responseTime, "0.00") & "초" & vbLf & _
               "신뢰도: " & Format(response.confidence * 100, "0") & "%" & vbLf & _
               "참조 문서: " & (response.internalCount + response.externalCount) & "건", _
               vbInformation, "RAG Search"
    Else
        ' 오프라인 모드로 전환
        If InStr(response.errorMessage, "Error") > 0 Then
            MsgBox "API 서버 연결 실패" & vbLf & vbLf & _
                   "Mock 데이터 모드로 전환합니다.", vbExclamation
            Call RunMockRAGSearch(question)
        Else
            MsgBox "오류: " & response.errorMessage, vbCritical
        End If
    End If
    
    Application.StatusBar = GetLabel("STATUS_READY")
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = GetLabel("STATUS_READY")
    MsgBox GetLabel("ERR_GENERAL") & ": " & Err.Description, vbCritical
End Sub

' =====================================
' Mock RAG 검색 (오프라인 모드)
' =====================================
Private Sub RunMockRAGSearch(question As String)
    Dim results As Collection
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    
    ' Mock 데이터에서 검색
    Set results = modMockData.SearchMockData(question)
    
    ' 결과 표시
    If results.Count > 0 Then
        MsgBox "Mock 데이터에서 " & results.Count & "건을 찾았습니다.", _
               vbInformation, "Mock RAG"
    Else
        MsgBox "관련 문서를 찾을 수 없습니다.", vbInformation
    End If
End Sub

' =====================================
' RAG 결과 표시
' =====================================
Private Sub DisplayRAGResults(ws As Worksheet, response As RAGResponse)
    Dim startRow As Long
    startRow = 40
    
    ' 답변 표시
    With ws.Range("B" & startRow & ":M" & startRow)
        .Merge
        .Value = "💡 AI 답변"
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = COLOR_INFO
        .Font.Color = RGB(255, 255, 255)
    End With
    
    With ws.Range("B" & (startRow + 1) & ":M" & (startRow + 5))
        .Merge
        .Value = response.answer
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' 소스 문서 표시
    startRow = startRow + 7
    With ws.Range("B" & startRow & ":M" & startRow)
        .Merge
        .Value = "📚 참조 문서 (" & response.sources.Count & "건)"
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 230)
    End With
    
    Dim i As Integer
    Dim source As SourceDoc
    For i = 1 To response.sources.Count
        If i > 5 Then Exit For
        source = response.sources(i)
        
        ws.Cells(startRow + i, 2).Value = i
        ws.Cells(startRow + i, 3).Value = source.title
        ws.Cells(startRow + i, 5).Value = source.organization
        ws.Cells(startRow + i, 7).Value = source.docDate
        ws.Cells(startRow + i, 9).Value = Format(source.relevance * 100, "0") & "%"
        
        With ws.Range(ws.Cells(startRow + i, 2), ws.Cells(startRow + i, 13))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
        End With
    Next i
End Sub

' =====================================
' 피드백을 RAG에 반영
' =====================================
Sub UpdateRAGWithFeedback(feedback As String)
    Dim requestBody As String
    Dim http As Object
    Dim url As String
    
    On Error GoTo ErrorHandler
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = GetAPIUrl("feedback")
    
    ' JSON 생성
    requestBody = "{"
    requestBody = requestBody & """feedback"":""" & modUTF8.EscapeJSON(feedback) & ""","
    requestBody = requestBody & """timestamp"":""" & Format(Now, "yyyy-mm-dd hh:mm:ss") & ""","
    requestBody = requestBody & """user"":""" & Application.UserName & """"
    requestBody = requestBody & "}"
    
    ' API 호출
    With http
        .Open "POST", url, False
        Call modUTF8.SetUTF8Headers(http)
        .send modUTF8.StringToUTF8Bytes(requestBody)
        
        If .Status = 200 Then
            MsgBox "피드백이 RAG 시스템에 반영되었습니다.", vbInformation
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    ' 오류 시 로컬 저장
    Call SaveFeedbackLocally(feedback)
End Sub

' =====================================
' 피드백 로컬 저장
' =====================================
Private Sub SaveFeedbackLocally(feedback As String)
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\feedback_" & Format(Now, "yyyymmdd") & ".txt"
    
    Call modUTF8.WriteFileUTF8(filePath, feedback & vbCrLf & _
                                "Timestamp: " & Now & vbCrLf & _
                                "User: " & Application.UserName & vbCrLf & _
                                "---" & vbCrLf)
    
    MsgBox "피드백이 로컬에 저장되었습니다.", vbInformation
End Sub

' =====================================
' RAG 시스템 상태 확인
' =====================================
Function CheckRAGStatus() As Boolean
    Dim http As Object
    Dim url As String
    
    On Error GoTo Offline
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = GetAPIUrl("health")
    
    http.Open "GET", url, False
    http.setTimeouts 5000, 5000, 5000, 5000
    http.send
    
    CheckRAGStatus = (http.Status = 200)
    Exit Function
    
Offline:
    CheckRAGStatus = False
End Function