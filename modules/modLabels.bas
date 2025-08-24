Attribute VB_Name = "modLabels"
' =====================================
' STRIX v2 - Labels Module
' 한글 텍스트 중앙 관리
' =====================================
Option Explicit

Private g_Labels As Object  ' Dictionary 객체

' =====================================
' 라벨 초기화
' =====================================
Public Sub InitializeLabels()
    Set g_Labels = CreateObject("Scripting.Dictionary")
    
    ' ===== 메인 UI 라벨 =====
    g_Labels.Add "MAIN_TITLE", "STRIX 보고 업무 자동화 시스템 v2"
    g_Labels.Add "MAIN_SUBTITLE", "AI 기반 3단계 워크플로우"
    g_Labels.Add "LAST_UPDATE", "마지막 업데이트"
    
    ' ===== Phase 1: 보고 준비 이전 =====
    g_Labels.Add "PHASE1_TITLE", "Phase 1: 보고 준비 이전"
    g_Labels.Add "PHASE1_DESC", "이전 피드백 확인 및 자료 수집"
    g_Labels.Add "PHASE1_FEEDBACK", "이전 피드백 확인"
    g_Labels.Add "PHASE1_COLLECT", "자료 수집"
    g_Labels.Add "PHASE1_IDENTIFY", "핵심 이슈 식별"
    g_Labels.Add "PHASE1_INTERNAL", "내부 보고서"
    g_Labels.Add "PHASE1_EXTERNAL", "외부 뉴스"
    g_Labels.Add "PHASE1_COMPETITOR", "경쟁사 동향"
    g_Labels.Add "PHASE1_POLICY", "정책/규제"
    
    ' ===== Phase 2: 보고 준비 =====
    g_Labels.Add "PHASE2_TITLE", "Phase 2: 보고 준비"
    g_Labels.Add "PHASE2_DESC", "자료 종합 및 보고서 작성"
    g_Labels.Add "PHASE2_CONSOLIDATE", "자료 종합"
    g_Labels.Add "PHASE2_ANALYZE", "AI 분석"
    g_Labels.Add "PHASE2_WRITE", "보고서 작성"
    g_Labels.Add "PHASE2_INSIGHT", "핵심 인사이트"
    g_Labels.Add "PHASE2_SUMMARY", "요약"
    g_Labels.Add "PHASE2_TREND", "시장 동향"
    g_Labels.Add "PHASE2_RISK", "리스크 평가"
    g_Labels.Add "PHASE2_STRATEGY", "전략 제언"
    
    ' ===== Phase 3: 보고 이후 =====
    g_Labels.Add "PHASE3_TITLE", "Phase 3: 보고 이후"
    g_Labels.Add "PHASE3_DESC", "피드백 반영 및 추적 관리"
    g_Labels.Add "PHASE3_FEEDBACK", "피드백 수집"
    g_Labels.Add "PHASE3_CLASSIFY", "피드백 분류"
    g_Labels.Add "PHASE3_UPDATE", "RAG 시스템 업데이트"
    g_Labels.Add "PHASE3_TRACK", "이슈 추적"
    g_Labels.Add "PHASE3_ACTION", "액션 아이템"
    g_Labels.Add "PHASE3_LEARNING", "학습 완료"
    
    ' ===== 버튼 라벨 =====
    g_Labels.Add "BTN_EXECUTE", "실행"
    g_Labels.Add "BTN_REFRESH", "새로고침"
    g_Labels.Add "BTN_SAVE", "저장"
    g_Labels.Add "BTN_LOAD", "불러오기"
    g_Labels.Add "BTN_EXPORT", "내보내기"
    g_Labels.Add "BTN_SETTINGS", "설정"
    g_Labels.Add "BTN_HELP", "도움말"
    g_Labels.Add "BTN_SEARCH", "검색"
    g_Labels.Add "BTN_ANALYZE", "분석"
    g_Labels.Add "BTN_GENERATE", "생성"
    g_Labels.Add "BTN_PREVIEW", "미리보기"
    g_Labels.Add "BTN_APPLY", "적용"
    g_Labels.Add "BTN_CANCEL", "취소"
    
    ' ===== 상태 메시지 =====
    g_Labels.Add "STATUS_READY", "준비 완료"
    g_Labels.Add "STATUS_PROCESSING", "처리 중..."
    g_Labels.Add "STATUS_COMPLETE", "완료"
    g_Labels.Add "STATUS_ERROR", "오류 발생"
    g_Labels.Add "STATUS_LOADING", "로딩 중..."
    g_Labels.Add "STATUS_SAVING", "저장 중..."
    g_Labels.Add "STATUS_CONNECTING", "연결 중..."
    g_Labels.Add "STATUS_ANALYZING", "분석 중..."
    g_Labels.Add "STATUS_GENERATING", "생성 중..."
    
    ' ===== 메시지 =====
    g_Labels.Add "MSG_WELCOME", "STRIX v2에 오신 것을 환영합니다"
    g_Labels.Add "MSG_COMPLETE", "작업이 완료되었습니다"
    g_Labels.Add "MSG_CONFIRM", "계속 진행하시겠습니까?"
    g_Labels.Add "MSG_NO_DATA", "데이터가 없습니다"
    g_Labels.Add "MSG_NO_CONNECTION", "서버에 연결할 수 없습니다"
    g_Labels.Add "MSG_INVALID_INPUT", "올바른 값을 입력해주세요"
    g_Labels.Add "MSG_SUCCESS", "성공적으로 처리되었습니다"
    g_Labels.Add "MSG_FAILED", "처리에 실패했습니다"
    
    ' ===== 지표 라벨 =====
    g_Labels.Add "METRIC_DOCS", "수집 문서"
    g_Labels.Add "METRIC_ANALYZED", "분석 완료"
    g_Labels.Add "METRIC_ISSUES", "미해결 이슈"
    g_Labels.Add "METRIC_ACTIONS", "진행중 액션"
    g_Labels.Add "METRIC_FEEDBACK", "피드백"
    g_Labels.Add "METRIC_ACCURACY", "정확도"
    g_Labels.Add "METRIC_TIME", "처리 시간"
    g_Labels.Add "METRIC_COUNT", "건"
    g_Labels.Add "METRIC_PERCENT", "%"
    g_Labels.Add "METRIC_HOURS", "시간"
    
    ' ===== 테이블 헤더 =====
    g_Labels.Add "COL_NO", "번호"
    g_Labels.Add "COL_TITLE", "제목"
    g_Labels.Add "COL_TYPE", "유형"
    g_Labels.Add "COL_DATE", "날짜"
    g_Labels.Add "COL_STATUS", "상태"
    g_Labels.Add "COL_PRIORITY", "우선순위"
    g_Labels.Add "COL_OWNER", "담당자"
    g_Labels.Add "COL_DUE_DATE", "마감일"
    g_Labels.Add "COL_PROGRESS", "진행률"
    g_Labels.Add "COL_CATEGORY", "카테고리"
    g_Labels.Add "COL_SOURCE", "출처"
    g_Labels.Add "COL_RELEVANCE", "관련도"
    
    ' ===== 카테고리 =====
    g_Labels.Add "CAT_INTERNAL", "내부"
    g_Labels.Add "CAT_EXTERNAL", "외부"
    g_Labels.Add "CAT_URGENT", "긴급"
    g_Labels.Add "CAT_NORMAL", "일반"
    g_Labels.Add "CAT_LOW", "낮음"
    g_Labels.Add "CAT_HIGH", "높음"
    g_Labels.Add "CAT_CRITICAL", "매우 중요"
    
    ' ===== 날짜/시간 =====
    g_Labels.Add "TIME_TODAY", "오늘"
    g_Labels.Add "TIME_YESTERDAY", "어제"
    g_Labels.Add "TIME_THIS_WEEK", "이번 주"
    g_Labels.Add "TIME_LAST_WEEK", "지난 주"
    g_Labels.Add "TIME_THIS_MONTH", "이번 달"
    g_Labels.Add "TIME_LAST_MONTH", "지난 달"
    
    ' ===== 에러 메시지 =====
    g_Labels.Add "ERR_CONNECTION", "연결 오류"
    g_Labels.Add "ERR_TIMEOUT", "시간 초과"
    g_Labels.Add "ERR_INVALID", "잘못된 요청"
    g_Labels.Add "ERR_NOTFOUND", "찾을 수 없음"
    g_Labels.Add "ERR_PERMISSION", "권한 없음"
    g_Labels.Add "ERR_GENERAL", "일반 오류"
End Sub

' =====================================
' 라벨 가져오기
' =====================================
Public Function GetLabel(key As String, Optional defaultValue As String = "") As String
    If g_Labels Is Nothing Then
        Call InitializeLabels
    End If
    
    If g_Labels.Exists(key) Then
        GetLabel = g_Labels(key)
    ElseIf defaultValue <> "" Then
        GetLabel = defaultValue
    Else
        GetLabel = "[" & key & "]"  ' 키가 없으면 키 자체를 반환
    End If
End Function

' =====================================
' 포맷된 라벨 가져오기
' =====================================
Public Function GetFormattedLabel(key As String, ParamArray args() As Variant) As String
    Dim label As String
    label = GetLabel(key)
    
    Dim i As Integer
    For i = 0 To UBound(args)
        label = Replace(label, "{" & i & "}", CStr(args(i)))
    Next i
    
    GetFormattedLabel = label
End Function

' =====================================
' 동적 라벨 추가
' =====================================
Public Sub AddLabel(key As String, value As String)
    If g_Labels Is Nothing Then
        Call InitializeLabels
    End If
    
    If g_Labels.Exists(key) Then
        g_Labels(key) = value
    Else
        g_Labels.Add key, value
    End If
End Sub

' =====================================
' 라벨 존재 확인
' =====================================
Public Function LabelExists(key As String) As Boolean
    If g_Labels Is Nothing Then
        Call InitializeLabels
    End If
    
    LabelExists = g_Labels.Exists(key)
End Function

' =====================================
' 날짜 포맷
' =====================================
Public Function FormatDateTime(dt As Date) As String
    FormatDateTime = Format(dt, "yyyy-mm-dd hh:mm")
End Function

Public Function FormatDate(dt As Date) As String
    FormatDate = Format(dt, "yyyy-mm-dd")
End Function

Public Function FormatTime(dt As Date) As String
    FormatTime = Format(dt, "hh:mm:ss")
End Function