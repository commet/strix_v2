Attribute VB_Name = "modConfig"
' =====================================
' STRIX v2 - Configuration Module
' 시스템 설정 및 상수 관리
' =====================================
Option Explicit

' ===== 시스템 상수 =====
Public Const APP_NAME As String = "STRIX v2"
Public Const VERSION As String = "2.0.0"
Public Const BUILD_DATE As String = "2025-08-04"

' ===== API 설정 =====
Public Const API_BASE_URL As String = "http://localhost:5000"
Public Const API_TIMEOUT As Long = 30000  ' 30초
Public Const API_RETRY_COUNT As Integer = 3

' ===== 파일 경로 =====
Public Const DATA_FOLDER As String = "data\"
Public Const LOG_FOLDER As String = "logs\"
Public Const TEMP_FOLDER As String = "temp\"

' ===== 시트 이름 (영문) =====
Public Const SHEET_MAIN As String = "STRIX_Main"
Public Const SHEET_PHASE1 As String = "Phase1_PreReport"
Public Const SHEET_PHASE2 As String = "Phase2_Reporting"
Public Const SHEET_PHASE3 As String = "Phase3_PostReport"
Public Const SHEET_ANALYTICS As String = "Analytics"
Public Const SHEET_ALERTS As String = "Alerts"
Public Const SHEET_SETTINGS As String = "Settings"
Public Const SHEET_LOGS As String = "Logs"

' ===== 색상 테마 =====
Public Const COLOR_PRIMARY As Long = 1643520    ' RGB(25, 45, 95)
Public Const COLOR_PHASE1 As Long = 14395790    ' RGB(52, 152, 219)
Public Const COLOR_PHASE2 As Long = 7528448     ' RGB(46, 204, 113)
Public Const COLOR_PHASE3 As Long = 11957230    ' RGB(155, 89, 182)
Public Const COLOR_WARNING As Long = 3447039    ' RGB(231, 76, 60)
Public Const COLOR_SUCCESS As Long = 7528448    ' RGB(46, 204, 113)
Public Const COLOR_INFO As Long = 14395790      ' RGB(52, 152, 219)
Public Const COLOR_BACKGROUND As Long = 16316664 ' RGB(248, 249, 250)

' ===== 설정 변수 =====
Private g_Settings As Object  ' Dictionary 객체

' =====================================
' 설정 초기화
' =====================================
Public Sub InitializeConfig()
    Set g_Settings = CreateObject("Scripting.Dictionary")
    
    ' 기본 설정값
    g_Settings.Add "AutoSave", True
    g_Settings.Add "AutoRefresh", True
    g_Settings.Add "RefreshInterval", 300  ' 5분
    g_Settings.Add "MaxDocuments", 200
    g_Settings.Add "EnableLogging", True
    g_Settings.Add "Language", "ko-KR"
    g_Settings.Add "Theme", "Default"
    
    ' 설정 시트에서 로드
    Call LoadSettingsFromSheet
End Sub

' =====================================
' 설정 시트에서 로드
' =====================================
Private Sub LoadSettingsFromSheet()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_SETTINGS)
    
    If ws Is Nothing Then Exit Sub
    
    ' 설정값이 있으면 로드
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        ' 설정이 없으면 기본값 저장
        Call SaveSettingsToSheet
        Exit Sub
    End If
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            g_Settings(ws.Cells(i, 1).Value) = ws.Cells(i, 2).Value
        End If
    Next i
End Sub

' =====================================
' 설정 시트에 저장
' =====================================
Public Sub SaveSettingsToSheet()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_SETTINGS)
    
    If ws Is Nothing Then Exit Sub
    
    ' 시트 초기화
    ws.Cells.Clear
    
    ' 헤더
    ws.Range("A1").Value = "Setting"
    ws.Range("B1").Value = "Value"
    ws.Range("C1").Value = "Description"
    ws.Range("A1:C1").Font.Bold = True
    
    ' 설정값 저장
    Dim key As Variant
    Dim row As Long
    row = 2
    
    For Each key In g_Settings.Keys
        ws.Cells(row, 1).Value = key
        ws.Cells(row, 2).Value = g_Settings(key)
        row = row + 1
    Next key
    
    ' 시트 숨기기
    ws.Visible = xlSheetVeryHidden
End Sub

' =====================================
' 설정값 가져오기
' =====================================
Public Function GetSetting(key As String, Optional defaultValue As Variant = "") As Variant
    If g_Settings Is Nothing Then
        Call InitializeConfig
    End If
    
    If g_Settings.Exists(key) Then
        GetSetting = g_Settings(key)
    Else
        GetSetting = defaultValue
    End If
End Function

' =====================================
' 설정값 저장
' =====================================
Public Sub SetSetting(key As String, value As Variant)
    If g_Settings Is Nothing Then
        Call InitializeConfig
    End If
    
    If g_Settings.Exists(key) Then
        g_Settings(key) = value
    Else
        g_Settings.Add key, value
    End If
    
    ' 시트에 저장
    Call SaveSettingsToSheet
End Sub

' =====================================
' 시트 가져오기 또는 생성
' =====================================
Public Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    Set GetOrCreateSheet = ws
End Function

' =====================================
' API URL 생성
' =====================================
Public Function GetAPIUrl(endpoint As String) As String
    GetAPIUrl = API_BASE_URL & "/api/" & endpoint
End Function

' =====================================
' 로그 파일 경로
' =====================================
Public Function GetLogFilePath() As String
    GetLogFilePath = ThisWorkbook.Path & "\" & LOG_FOLDER & _
                     Format(Date, "yyyymmdd") & "_strix.log"
End Function

' =====================================
' 임시 파일 경로
' =====================================
Public Function GetTempFilePath(Optional extension As String = "tmp") As String
    GetTempFilePath = ThisWorkbook.Path & "\" & TEMP_FOLDER & _
                      "temp_" & Format(Now, "yyyymmdd_hhmmss") & "." & extension
End Function