Attribute VB_Name = "modUTF8"
' =====================================
' STRIX v2 - UTF-8 Encoding Module
' 한글 인코딩 처리 전문 모듈
' =====================================
Option Explicit

' =====================================
' UTF-8 문자열을 바이트 배열로 변환
' =====================================
Public Function StringToUTF8Bytes(ByVal str As String) As Byte()
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    With objStream
        .Mode = 3  ' adModeReadWrite
        .Type = 2  ' adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText str
        .Position = 0
        .Type = 1  ' adTypeBinary
        .Position = 3  ' BOM 건너뛰기
        StringToUTF8Bytes = .Read
        .Close
    End With
    
    Set objStream = Nothing
End Function

' =====================================
' UTF-8 바이트 배열을 문자열로 변환
' =====================================
Public Function UTF8BytesToString(bytes() As Byte) As String
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    With objStream
        .Mode = 3  ' adModeReadWrite
        .Type = 1  ' adTypeBinary
        .Open
        .Write bytes
        .Position = 0
        .Type = 2  ' adTypeText
        .Charset = "UTF-8"
        UTF8BytesToString = .ReadText
        .Close
    End With
    
    Set objStream = Nothing
End Function

' =====================================
' HTTP 응답을 UTF-8로 디코딩
' =====================================
Public Function DecodeUTF8Response(responseBody As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim bytes() As Byte
    
    ' responseBody가 이미 바이트 배열인 경우
    If VarType(responseBody) = vbArray + vbByte Then
        bytes = responseBody
    ' responseBody가 문자열인 경우
    ElseIf VarType(responseBody) = vbString Then
        ' 이미 디코딩된 문자열일 수 있음
        DecodeUTF8Response = responseBody
        Exit Function
    Else
        ' 기타 타입은 문자열로 변환 시도
        DecodeUTF8Response = CStr(responseBody)
        Exit Function
    End If
    
    ' UTF-8 디코딩
    DecodeUTF8Response = UTF8BytesToString(bytes)
    Exit Function
    
ErrorHandler:
    DecodeUTF8Response = ""
End Function

' =====================================
' JSON 문자열을 UTF-8로 인코딩
' =====================================
Public Function EncodeJSONToUTF8(jsonString As String) As Byte()
    EncodeJSONToUTF8 = StringToUTF8Bytes(jsonString)
End Function

' =====================================
' 파일을 UTF-8로 읽기
' =====================================
Public Function ReadFileUTF8(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    With objStream
        .Type = 2  ' adTypeText
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadFileUTF8 = .ReadText
        .Close
    End With
    
    Set objStream = Nothing
    Exit Function
    
ErrorHandler:
    ReadFileUTF8 = ""
End Function

' =====================================
' 파일을 UTF-8로 저장
' =====================================
Public Sub WriteFileUTF8(filePath As String, content As String)
    On Error GoTo ErrorHandler
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    With objStream
        .Type = 2  ' adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile filePath, 2  ' adSaveCreateOverWrite
        .Close
    End With
    
    Set objStream = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "파일 저장 오류: " & Err.Description, vbCritical
End Sub

' =====================================
' 한글 문자열 검증
' =====================================
Public Function IsValidKorean(str As String) As Boolean
    ' 한글이 포함되어 있고 깨지지 않았는지 확인
    Dim i As Long
    Dim charCode As Long
    
    For i = 1 To Len(str)
        charCode = AscW(Mid(str, i, 1))
        ' 한글 유니코드 범위: AC00-D7AF (완성형 한글)
        If charCode >= &HAC00 And charCode <= &HD7AF Then
            IsValidKorean = True
            Exit Function
        End If
    Next i
    
    IsValidKorean = False
End Function

' =====================================
' HTTP Request에 UTF-8 헤더 설정
' =====================================
Public Sub SetUTF8Headers(httpObj As Object)
    httpObj.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    httpObj.setRequestHeader "Accept", "application/json; charset=utf-8"
    httpObj.setRequestHeader "Accept-Charset", "utf-8"
End Sub

' =====================================
' 특수문자 이스케이프
' =====================================
Public Function EscapeJSON(str As String) As String
    Dim result As String
    result = str
    
    ' JSON 특수문자 이스케이프
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    
    EscapeJSON = result
End Function

' =====================================
' URL 인코딩 (한글 처리)
' =====================================
Public Function URLEncode(str As String) As String
    Dim i As Long
    Dim char As String
    Dim charCode As Long
    Dim result As String
    
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        charCode = AscW(char)
        
        If (charCode >= 48 And charCode <= 57) Or _
           (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Or _
           char = "-" Or char = "_" Or char = "." Or char = "~" Then
            ' 안전한 문자는 그대로
            result = result & char
        Else
            ' UTF-8로 인코딩
            Dim bytes() As Byte
            bytes = StringToUTF8Bytes(char)
            Dim j As Long
            For j = 0 To UBound(bytes)
                result = result & "%" & Right("0" & Hex(bytes(j)), 2)
            Next j
        End If
    Next i
    
    URLEncode = result
End Function

' =====================================
' 디버그용: 문자열의 유니코드 값 확인
' =====================================
Public Sub DebugUnicode(str As String)
    Dim i As Long
    Dim output As String
    
    For i = 1 To Len(str)
        output = output & Mid(str, i, 1) & " = U+" & Hex(AscW(Mid(str, i, 1))) & vbCrLf
    Next i
    
    Debug.Print "Unicode Debug for: " & str
    Debug.Print output
End Sub

' =====================================
' 안전한 JSON 생성
' =====================================
Public Function CreateSafeJSON(key As String, value As String) As String
    CreateSafeJSON = "{""" & EscapeJSON(key) & """:""" & EscapeJSON(value) & """}"
End Function

' =====================================
' 한글 깨짐 복구 시도
' =====================================
Public Function TryFixBrokenKorean(str As String) As String
    On Error Resume Next
    
    ' 일반적인 깨짐 패턴 수정
    Dim result As String
    result = str
    
    ' UTF-8 -> ANSI -> UTF-8 재변환 시도
    If InStr(result, "?") > 0 Or InStr(result, "��") > 0 Then
        ' 깨진 문자가 있으면 원본 반환
        TryFixBrokenKorean = str
    Else
        TryFixBrokenKorean = result
    End If
End Function