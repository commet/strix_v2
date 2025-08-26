# STRIX v2 RAG API Server

## 개요
STRIX v2의 Python 기반 RAG (Retrieval Augmented Generation) API 서버입니다.
LangChain, Supabase, FastAPI를 활용하여 구현되었습니다.

## 주요 기능
- 📚 **문서 임베딩 및 벡터 검색**: Supabase Vector Store 활용
- 🤖 **LLM 기반 답변 생성**: OpenAI GPT-4 또는 Google Gemini
- 📄 **다양한 문서 형식 지원**: PDF, DOCX, XLSX, TXT 등
- 🔄 **실시간 문서 업로드 및 처리**
- 🎯 **컨텍스트 기반 정확한 답변**

## 설치 방법

### 1. 의존성 설치
```bash
pip install -r requirements.txt
```

### 2. 환경 변수 설정
`.env.example`을 `.env`로 복사하고 설정값 입력:
```bash
cp .env.example .env
```

필수 설정:
- `OPENAI_API_KEY`: OpenAI API 키
- `SUPABASE_URL`: Supabase 프로젝트 URL
- `SUPABASE_KEY`: Supabase anon 키

### 3. Supabase 설정
Supabase 프로젝트에서 벡터 확장 활성화:
```sql
-- Enable vector extension
create extension if not exists vector;

-- Create documents table
create table strix_documents (
  id uuid primary key default uuid_generate_v4(),
  content text,
  metadata jsonb,
  embedding vector(1536),
  created_at timestamp with time zone default now()
);

-- Create similarity search function
create or replace function match_strix_documents(
  query_embedding vector(1536),
  match_threshold float,
  match_count int
)
returns table (
  id uuid,
  content text,
  metadata jsonb,
  similarity float
)
language sql stable
as $$
  select
    id,
    content,
    metadata,
    1 - (strix_documents.embedding <=> query_embedding) as similarity
  from strix_documents
  where 1 - (strix_documents.embedding <=> query_embedding) > match_threshold
  order by similarity desc
  limit match_count;
$$;
```

## 서버 실행

### 개발 모드
```bash
python main.py
```

### 프로덕션 모드
```bash
uvicorn main:app --host 0.0.0.0 --port 5000 --workers 4
```

### Mock 모드 (테스트용)
`.env` 파일에서 `MOCK_MODE=true` 설정

## API 엔드포인트

### 1. RAG 질의
```http
POST /api/query
Content-Type: application/json

{
  "question": "SK온 합병 계획은?",
  "doc_type": "both",
  "max_results": 10,
  "include_sources": true
}
```

응답:
```json
{
  "answer": "SK온과 SK이노베이션의 배터리 사업 통합은...",
  "confidence": 0.92,
  "internal_docs": 3,
  "external_docs": 5,
  "sources": [...],
  "timestamp": "2025-08-26T10:00:00"
}
```

### 2. 문서 업로드
```http
POST /api/documents/upload
Content-Type: multipart/form-data

file: [binary]
doc_type: "internal"
organization: "전략기획팀"
```

### 3. 문서 검색
```http
GET /api/documents/search?query=전고체배터리&limit=5
```

### 4. 피드백 제출
```http
POST /api/feedback
Content-Type: application/json

{
  "feedback": "답변이 유용했습니다",
  "question": "원래 질문",
  "answer": "받은 답변",
  "rating": 5
}
```

## VBA 연동

Excel VBA에서 API 호출 예시:

```vba
Function CallRAGAPI(question As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim url As String
    url = "http://localhost:5000/api/query"
    
    Dim requestBody As String
    requestBody = "{""question"":""" & question & """,""doc_type"":""both""}"
    
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send requestBody
    
    If http.Status = 200 Then
        CallRAGAPI = http.responseText
    Else
        CallRAGAPI = "Error: " & http.Status
    End If
End Function
```

## 프로젝트 구조
```
api/
├── main.py           # FastAPI 메인 서버
├── config.py         # 환경 설정
├── requirements.txt  # 의존성 목록
├── .env.example     # 환경 변수 템플릿
└── rag/
    ├── __init__.py
    ├── chain.py          # LangGraph RAG 체인
    ├── vector_store.py   # Supabase 벡터 스토어
    └── document_loader.py # 문서 로더
```

## 문제 해결

### Mock 모드가 작동하지 않을 때
- `.env` 파일의 `MOCK_MODE=true` 확인
- 서버 재시작

### Supabase 연결 오류
- `SUPABASE_URL`과 `SUPABASE_KEY` 확인
- Supabase 프로젝트의 벡터 확장 활성화 확인

### 문서 업로드 실패
- 지원되는 파일 형식 확인 (.pdf, .docx, .xlsx, .txt)
- 파일 크기 제한 확인 (기본 10MB)

## 라이선스
내부 사용 전용