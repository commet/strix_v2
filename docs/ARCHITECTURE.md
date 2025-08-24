# STRIX v2 아키텍처 설계

## 🎯 핵심 원칙
1. **안정성**: 한글 인코딩 문제 완벽 해결
2. **실용성**: 보고 준비 실무자 중심 UI/UX
3. **확장성**: 모듈간 낮은 결합도

## 📊 시스템 구조

### 1. 기술 스택 (기존 STRIX 계승)
```
Frontend: Excel VBA
Backend: Python Flask + LangChain
Database: Supabase + pgvector
AI Model: GPT-4 (OpenAI API)
Encoding: UTF-8 throughout
```

### 2. 한글 처리 전략
```vba
' ❌ 피해야 할 방식
ws.Range("A1").Value = "보고서 작성"  ' 코드에 한글 직접 입력

' ✅ 권장 방식
ws.Range("A1").Value = GetLabel("REPORT_TITLE")  ' 설정에서 읽기
```

### 3. 모듈 구조
```
strix_v2/
├── modules/
│   ├── modMain.bas           # 진입점
│   ├── modConfig.bas         # 설정 및 상수
│   ├── modLabels.bas         # 한글 텍스트 관리
│   ├── modUTF8.bas           # UTF-8 인코딩 처리
│   ├── modPhase1.bas         # Phase 1: 보고 준비 이전
│   ├── modPhase2.bas         # Phase 2: 보고 준비
│   ├── modPhase3.bas         # Phase 3: 보고 이후
│   ├── modRAG.bas            # RAG API 연동
│   └── modUtils.bas          # 공통 유틸리티
```

### 4. API 구조
```python
/api/
├── query           # RAG 질의
├── collect         # 자료 수집
├── analyze         # AI 분석
├── feedback        # 피드백 처리
└── report          # 보고서 생성
```

### 5. 데이터 흐름
```
Excel VBA → UTF-8 JSON → Flask API → LangChain → Supabase
                ↓                         ↓
            한글 보존                  Vector Search
                ↓                         ↓
            UTF-8 JSON ← Flask API ← GPT-4 Response
```

## 🔧 구현 순서

### Phase 0: 기반 구축
1. modConfig - 설정 관리
2. modLabels - 한글 텍스트 중앙화
3. modUTF8 - 인코딩 처리

### Phase 1: 핵심 기능
1. modPhase1 - 자료 수집
2. modPhase2 - 분석/작성
3. modPhase3 - 피드백

### Phase 2: 연동
1. modRAG - API 연동
2. API 서버 구축

## ⚠️ 주의사항

### 1. VBA 한글 처리
- 모든 한글 텍스트는 modLabels에 집중
- 동적 생성 텍스트는 UTF-8 처리 필수
- API 응답은 반드시 UTF-8 디코딩

### 2. 오류 처리
- 모든 API 호출에 타임아웃 설정
- 오프라인 모드 대비
- 상세 로깅

### 3. 성능
- 대용량 데이터 처리시 청크 단위
- 비동기 처리 고려
- 캐싱 전략

## 📈 개선 사항 (v1 대비)

| 영역 | v1 문제점 | v2 해결책 |
|------|-----------|-----------|
| 한글 | 코드 내 깨짐 | 중앙화 관리 |
| 모듈 | 높은 결합도 | 독립적 구조 |
| 오류 | 불명확한 메시지 | 상세 로깅 |
| UI | 개발자 중심 | 실무자 중심 |