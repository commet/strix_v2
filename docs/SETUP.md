# STRIX v2 설치 가이드

## 📋 사전 요구사항

### 필수
- Microsoft Excel 2016 이상 (VBA 지원)
- Windows 10/11

### 선택 (API 서버용)
- Python 3.8+
- pip

## 🚀 빠른 설치 (3분)

### Step 1: Excel 설정
1. Excel 열기
2. 파일 → 옵션 → 리본 사용자 지정
3. "개발 도구" 체크 ✅
4. 확인

### Step 2: VBA 모듈 설치
1. 새 Excel 파일 생성
2. `Alt + F11` (VBA 편집기)
3. 프로젝트 → 가져오기
4. `strix_v2/modules` 폴더의 모든 `.bas` 파일 선택:
   - modConfig.bas
   - modLabels.bas  
   - modUTF8.bas
   - modMockData.bas
   - modMainV2.bas
   - modPhase1.bas
   - modPhase2.bas
   - modPhase3.bas
   - modRAG.bas

### Step 3: 참조 설정
VBA 편집기에서:
1. 도구 → 참조
2. 체크할 항목:
   - ✅ Microsoft XML, v6.0
   - ✅ Microsoft WinHTTP Services, version 5.1
   - ✅ Microsoft Scripting Runtime

### Step 4: 파일 저장
1. 파일 → 다른 이름으로 저장
2. 파일 형식: **Excel 매크로 사용 통합 문서 (*.xlsm)**
3. 파일명: `STRIX_v2.xlsm`

### Step 5: 실행
1. `Alt + F8`
2. `CreateSTRIXDashboard` 선택
3. 실행

## 🐍 API 서버 설치 (선택사항)

### 1. Python 환경 설정
```bash
cd strix_v2
python -m venv venv
venv\Scripts\activate  # Windows
```

### 2. 패키지 설치
```bash
pip install flask flask-cors
```

### 3. 서버 실행
```bash
cd api
python app.py
```

서버가 `http://localhost:5000`에서 실행됩니다.

## ✅ 설치 확인

### Excel에서:
1. STRIX Dashboard 시트가 생성되었는지 확인
2. 한글이 정상 표시되는지 확인
3. 버튼 클릭 시 동작 확인

### API 서버 (선택):
브라우저에서 http://localhost:5000/api/health 접속
```json
{
  "status": "healthy",
  "mode": "mock",
  "timestamp": "2025-08-04T10:00:00"
}
```

## 🔧 문제 해결

### 문제: 한글이 깨져서 표시됨
**해결:**
- Windows 제어판 → 지역 → 관리자 옵션
- "유니코드를 지원하지 않는 프로그램용 언어" → 한국어
- 재부팅

### 문제: 매크로 차단됨
**해결:**
- 파일 → 옵션 → 보안 센터 → 보안 센터 설정
- 매크로 설정 → "모든 매크로 포함" 또는 "디지털 서명된 매크로만"

### 문제: API 연결 실패
**해결:**
- Mock 모드로 자동 전환됨 (정상)
- API 서버 실행 확인: `python app.py`
- 방화벽 확인: 5000번 포트 허용

## 📁 폴더 구조

```
strix_v2/
├── modules/          # VBA 모듈 (필수)
├── api/             # Python API (선택)
├── docs/            # 문서
└── STRIX_v2.xlsm    # 완성된 Excel 파일
```

## 🎯 다음 단계

1. **Phase 1 실행**: 이전 피드백 확인 및 자료 수집
2. **Phase 2 실행**: 자료 분석 및 보고서 생성
3. **Phase 3 실행**: 피드백 반영 및 학습

## 📞 지원

문제 발생 시:
- GitHub Issues: github.com/your-org/strix_v2
- Email: strix-support@company.com