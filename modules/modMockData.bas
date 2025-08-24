Attribute VB_Name = "modMockData"
' =====================================
' STRIX v2 - Mock Data Module
' 한글 보고서 데이터 안전 처리
' =====================================
Option Explicit

' Mock 데이터 타입 정의
Public Type MockReport
    id As String
    title As String
    category As String
    organization As String
    reportDate As String
    content As String
    docType As String  ' internal/external
End Type

' =====================================
' Mock 데이터 초기화 (메모리에 저장)
' =====================================
Public Function InitializeMockData() As Collection
    Dim reports As New Collection
    Dim report As MockReport
    
    ' ===== 내부 보고서 Mock 데이터 =====
    
    ' 1. SK온 합병 관련
    report.id = "INT_001"
    report.title = "SK온-SK엔무브 합병 시너지 분석"
    report.category = "전략기획"
    report.organization = "전략기획팀"
    report.reportDate = "2025-07-30"
    report.content = GetSKMergerContent()
    report.docType = "internal"
    reports.Add report
    
    ' 2. 전고체 배터리 개발
    report.id = "INT_002"
    report.title = "전고체 배터리 개발 현황 및 로드맵"
    report.category = "R&D"
    report.organization = "배터리연구소"
    report.reportDate = "2025-08-01"
    report.content = GetSolidStateContent()
    report.docType = "internal"
    reports.Add report
    
    ' 3. IRA 대응 전략
    report.id = "INT_003"
    report.title = "IRA 정책 변화 대응 시나리오"
    report.category = "정책"
    report.organization = "정책대응팀"
    report.reportDate = "2025-08-02"
    report.content = GetIRAContent()
    report.docType = "internal"
    reports.Add report
    
    ' ===== 외부 뉴스 Mock 데이터 =====
    
    ' 4. BYD 기술 동향
    report.id = "EXT_001"
    report.title = "BYD 5분 충전 기술 공개 임팩트"
    report.category = "경쟁사"
    report.organization = "PR팀"
    report.reportDate = "2025-08-03"
    report.content = GetBYDContent()
    report.docType = "external"
    reports.Add report
    
    ' 5. 시장 동향
    report.id = "EXT_002"
    report.title = "글로벌 배터리 시장 동향 브리핑"
    report.category = "시장"
    report.organization = "마케팅팀"
    report.reportDate = "2025-08-04"
    report.content = GetMarketContent()
    report.docType = "external"
    reports.Add report
    
    Set InitializeMockData = reports
End Function

' =====================================
' SK온 합병 콘텐츠
' =====================================
Private Function GetSKMergerContent() As String
    Dim content As String
    content = "SK온-SK엔무브 합병 시너지 분석 보고서" & vbCrLf & vbCrLf
    content = content & "1. 합병 개요" & vbCrLf
    content = content & "- 합병 예정일: 2025년 11월 1일" & vbCrLf
    content = content & "- 통합법인명: SK온 (존속법인)" & vbCrLf
    content = content & "- 예상 시너지: 20조원 (5년간)" & vbCrLf & vbCrLf
    content = content & "2. 핵심 시너지 효과" & vbCrLf
    content = content & "가. 기술 시너지" & vbCrLf
    content = content & "- 배터리-분리막 통합 기술 개발" & vbCrLf
    content = content & "- 전고체 배터리 공동 개발 가속화" & vbCrLf
    content = content & "- 특허 포트폴리오 통합 (15,000건)" & vbCrLf & vbCrLf
    content = content & "나. 비용 절감" & vbCrLf
    content = content & "- 중복 투자 제거: 연 3조원" & vbCrLf
    content = content & "- 구매력 강화: 원재료 10% 절감" & vbCrLf
    content = content & "- 운영 효율화: 인력 20% 감축" & vbCrLf & vbCrLf
    content = content & "3. 리스크 요인" & vbCrLf
    content = content & "- 조직 문화 통합 과제" & vbCrLf
    content = content & "- 고객사 승인 필요" & vbCrLf
    content = content & "- 규제 당국 심사"
    GetSKMergerContent = content
End Function

' =====================================
' 전고체 배터리 콘텐츠
' =====================================
Private Function GetSolidStateContent() As String
    Dim content As String
    content = "전고체 배터리 개발 현황 보고" & vbCrLf & vbCrLf
    content = content & "1. 기술 개발 현황" & vbCrLf
    content = content & "- 에너지 밀도: 400Wh/kg 달성" & vbCrLf
    content = content & "- 충전 시간: 10분 (80% 충전)" & vbCrLf
    content = content & "- 수명: 3,000 사이클 확보" & vbCrLf & vbCrLf
    content = content & "2. 양산 로드맵" & vbCrLf
    content = content & "- 2025년: 파일럿 라인 구축" & vbCrLf
    content = content & "- 2026년: 소규모 양산 (10GWh)" & vbCrLf
    content = content & "- 2027년: 대량 양산 (50GWh)" & vbCrLf & vbCrLf
    content = content & "3. 투자 계획" & vbCrLf
    content = content & "- 총 투자: 10조원 (3년간)" & vbCrLf
    content = content & "- R&D: 3조원" & vbCrLf
    content = content & "- 생산설비: 7조원"
    GetSolidStateContent = content
End Function

' =====================================
' IRA 대응 콘텐츠
' =====================================
Private Function GetIRAContent() As String
    Dim content As String
    content = "IRA 정책 변화 대응 시나리오" & vbCrLf & vbCrLf
    content = content & "1. 정책 변화 시나리오" & vbCrLf
    content = content & "시나리오 A: 현행 유지 (확률 30%)" & vbCrLf
    content = content & "- AMPC 세액공제 유지" & vbCrLf
    content = content & "- 연간 혜택: 2,000억원" & vbCrLf & vbCrLf
    content = content & "시나리오 B: 부분 수정 (확률 50%)" & vbCrLf
    content = content & "- 세액공제 50% 축소" & vbCrLf
    content = content & "- 연간 혜택: 1,000억원" & vbCrLf & vbCrLf
    content = content & "시나리오 C: 전면 폐지 (확률 20%)" & vbCrLf
    content = content & "- AMPC 완전 폐지" & vbCrLf
    content = content & "- 대체 지원책 모색 필요" & vbCrLf & vbCrLf
    content = content & "2. 대응 전략" & vbCrLf
    content = content & "- 북미 현지 생산 확대" & vbCrLf
    content = content & "- 유럽 시장 다변화" & vbCrLf
    content = content & "- 원가 절감 가속화"
    GetIRAContent = content
End Function

' =====================================
' BYD 기술 콘텐츠
' =====================================
Private Function GetBYDContent() As String
    Dim content As String
    content = "BYD 5분 충전 기술 분석" & vbCrLf & vbCrLf
    content = content & "1. 기술 사양" & vbCrLf
    content = content & "- 충전 시간: 5분 (10-80%)" & vbCrLf
    content = content & "- 주행거리: 400km" & vbCrLf
    content = content & "- 배터리 용량: 80kWh" & vbCrLf
    content = content & "- 충전 출력: 800kW" & vbCrLf & vbCrLf
    content = content & "2. 핵심 기술" & vbCrLf
    content = content & "- 실리콘 음극재 적용" & vbCrLf
    content = content & "- 고전압 아키텍처 (1200V)" & vbCrLf
    content = content & "- 혁신적 냉각 시스템" & vbCrLf & vbCrLf
    content = content & "3. 시장 영향" & vbCrLf
    content = content & "- 충전 인프라 재편 예상" & vbCrLf
    content = content & "- 경쟁사 기술 개발 가속화" & vbCrLf
    content = content & "- 전기차 대중화 촉진"
    GetBYDContent = content
End Function

' =====================================
' 시장 동향 콘텐츠
' =====================================
Private Function GetMarketContent() As String
    Dim content As String
    content = "글로벌 배터리 시장 동향" & vbCrLf & vbCrLf
    content = content & "1. 시장 규모" & vbCrLf
    content = content & "- 2025년: 500GWh" & vbCrLf
    content = content & "- 2030년: 2,000GWh (전망)" & vbCrLf
    content = content & "- CAGR: 32%" & vbCrLf & vbCrLf
    content = content & "2. 점유율 현황" & vbCrLf
    content = content & "- CATL: 37.9%" & vbCrLf
    content = content & "- BYD: 15.7%" & vbCrLf
    content = content & "- LG에너지솔루션: 13.5%" & vbCrLf
    content = content & "- 파나소닉: 8.2%" & vbCrLf
    content = content & "- SK온: 5.8%" & vbCrLf
    content = content & "- 삼성SDI: 4.9%" & vbCrLf & vbCrLf
    content = content & "3. 주요 트렌드" & vbCrLf
    content = content & "- LFP 배터리 비중 증가 (60%)" & vbCrLf
    content = content & "- 전고체 배터리 개발 경쟁" & vbCrLf
    content = content & "- ESS 시장 급성장"
    GetMarketContent = content
End Function

' =====================================
' Mock 데이터를 시트에 로드
' =====================================
Public Sub LoadMockDataToSheet(ws As Worksheet, Optional dataType As String = "all")
    On Error GoTo ErrorHandler
    
    Dim reports As Collection
    Set reports = InitializeMockData()
    
    Dim startRow As Long
    startRow = 10  ' 데이터 시작 행
    
    Dim report As MockReport
    Dim i As Long
    i = 0
    
    Dim item As Variant
    For Each item In reports
        report = item
        
        ' 필터링
        If dataType <> "all" Then
            If dataType <> report.docType Then
                GoTo NextItem
            End If
        End If
        
        i = i + 1
        
        ' 데이터 입력
        ws.Cells(startRow + i, 2).Value = i
        ws.Cells(startRow + i, 3).Value = report.title
        ws.Cells(startRow + i, 4).Value = report.category
        ws.Cells(startRow + i, 5).Value = report.organization
        ws.Cells(startRow + i, 6).Value = report.reportDate
        ws.Cells(startRow + i, 7).Value = report.docType
        
        ' 서식 적용
        With ws.Range(ws.Cells(startRow + i, 2), ws.Cells(startRow + i, 7))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(220, 220, 220)
            If i Mod 2 = 0 Then
                .Interior.Color = RGB(248, 248, 248)
            End If
        End With
        
NextItem:
    Next item
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Mock 데이터 로드 오류: " & Err.Description, vbCritical
End Sub

' =====================================
' Mock 데이터 검색
' =====================================
Public Function SearchMockData(keyword As String) As Collection
    Dim results As New Collection
    Dim reports As Collection
    Set reports = InitializeMockData()
    
    Dim report As MockReport
    Dim item As Variant
    
    For Each item In reports
        report = item
        
        ' 제목 또는 내용에서 키워드 검색
        If InStr(1, report.title, keyword, vbTextCompare) > 0 Or _
           InStr(1, report.content, keyword, vbTextCompare) > 0 Then
            results.Add report
        End If
    Next item
    
    Set SearchMockData = results
End Function

' =====================================
' Mock 데이터를 JSON으로 변환
' =====================================
Public Function MockDataToJSON(report As MockReport) As String
    Dim json As String
    
    json = "{"
    json = json & """id"":""" & modUTF8.EscapeJSON(report.id) & ""","
    json = json & """title"":""" & modUTF8.EscapeJSON(report.title) & ""","
    json = json & """category"":""" & modUTF8.EscapeJSON(report.category) & ""","
    json = json & """organization"":""" & modUTF8.EscapeJSON(report.organization) & ""","
    json = json & """date"":""" & report.reportDate & ""","
    json = json & """content"":""" & modUTF8.EscapeJSON(report.content) & ""","
    json = json & """type"":""" & report.docType & """"
    json = json & "}"
    
    MockDataToJSON = json
End Function