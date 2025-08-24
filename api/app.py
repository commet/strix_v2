"""
STRIX v2 API Server
Flask + LangChain + Supabase
"""
from flask import Flask, request, jsonify, Response
from flask_cors import CORS
import json
import os
import sys
from datetime import datetime

# Strix v1 경로 추가 (기존 모듈 재사용)
sys.path.append(os.path.join(os.path.dirname(__file__), '../../strix/src'))

# Mock 모드 설정 (Supabase 없이도 실행 가능)
MOCK_MODE = os.getenv('MOCK_MODE', 'true').lower() == 'true'

app = Flask(__name__)
CORS(app)

# Mock 데이터
MOCK_RESPONSES = {
    "default": {
        "answer": "SK온과 SK엔무브의 합병은 2025년 11월 1일 예정이며, 예상 시너지는 5년간 20조원입니다. 전고체 배터리는 2027년 양산을 목표로 개발 중이며, BYD의 5분 충전 기술에 대응하기 위한 전략을 수립 중입니다.",
        "confidence": 0.92,
        "internal_docs": 3,
        "external_docs": 5,
        "sources": [
            {
                "title": "SK온-SK엔무브 합병 시너지 분석",
                "organization": "전략기획팀",
                "date": "2025-07-30",
                "type": "internal",
                "relevance": 0.95
            },
            {
                "title": "전고체 배터리 개발 현황",
                "organization": "R&D센터",
                "date": "2025-08-01", 
                "type": "internal",
                "relevance": 0.88
            },
            {
                "title": "BYD 5분 충전 기술 분석",
                "organization": "기술전략팀",
                "date": "2025-08-03",
                "type": "external",
                "relevance": 0.85
            }
        ]
    }
}

@app.route('/api/query', methods=['POST'])
def query():
    """RAG 질의 처리"""
    try:
        data = request.get_json()
        question = data.get('question', '')
        doc_type = data.get('doc_type', 'both')
        
        if not question:
            return jsonify({"error": "No question provided"}), 400
        
        # Mock 모드
        if MOCK_MODE:
            response = MOCK_RESPONSES["default"].copy()
            response["question"] = question
            response["timestamp"] = datetime.now().isoformat()
            
            # UTF-8 응답
            return Response(
                json.dumps(response, ensure_ascii=False),
                mimetype='application/json; charset=utf-8'
            )
        
        # 실제 RAG 처리 (Supabase + LangChain)
        from rag.strix_chain import STRIXChain
        chain = STRIXChain()
        result = chain.invoke({"question": question})
        
        response = {
            "answer": result.get('answer', ''),
            "confidence": 0.9,
            "internal_docs": len(result.get('internal_docs', [])),
            "external_docs": len(result.get('external_docs', [])),
            "sources": [],
            "timestamp": datetime.now().isoformat()
        }
        
        return Response(
            json.dumps(response, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/collect', methods=['GET'])
def collect():
    """자료 수집"""
    try:
        # Mock 자료 수집 결과
        result = {
            "status": "success",
            "collected": {
                "internal": 25,
                "external": 127,
                "competitor": 18,
                "policy": 12,
                "total": 182
            },
            "timestamp": datetime.now().isoformat()
        }
        
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/analyze', methods=['POST'])
def analyze():
    """AI 분석"""
    try:
        data = request.get_json()
        
        # Mock 분석 결과
        result = {
            "status": "success",
            "insights": [
                {"category": "전략", "insight": "SK온 합병으로 20조원 시너지", "confidence": 0.92},
                {"category": "기술", "insight": "전고체 배터리 2027년 양산", "confidence": 0.88},
                {"category": "리스크", "insight": "IRA 정책 변경 가능성 70%", "confidence": 0.85}
            ],
            "timestamp": datetime.now().isoformat()
        }
        
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/feedback', methods=['POST'])
def feedback():
    """피드백 처리"""
    try:
        data = request.get_json()
        feedback_text = data.get('feedback', '')
        
        # 피드백 저장 (실제로는 DB에 저장)
        result = {
            "status": "success",
            "message": "피드백이 저장되었습니다",
            "feedback_id": f"FB_{datetime.now().strftime('%Y%m%d%H%M%S')}",
            "timestamp": datetime.now().isoformat()
        }
        
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/report', methods=['POST'])
def generate_report():
    """보고서 생성"""
    try:
        data = request.get_json()
        
        # Mock 보고서 생성
        result = {
            "status": "success",
            "report": {
                "title": "월간 배터리 산업 동향 보고서",
                "date": datetime.now().strftime('%Y-%m-%d'),
                "sections": [
                    "Executive Summary",
                    "시장 동향",
                    "경쟁사 분석",
                    "기술 개발 현황",
                    "리스크 평가",
                    "전략 제언"
                ],
                "pages": 25,
                "charts": 12,
                "tables": 8
            },
            "file_path": f"reports/STRIX_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
            "timestamp": datetime.now().isoformat()
        }
        
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/issues/predict', methods=['POST'])
def predict_issues():
    """이슈 예측"""
    try:
        # Mock 예측 결과
        result = {
            "status": "success",
            "predictions": [
                {
                    "issue": "SK온 합병 지연 가능성",
                    "probability": 0.15,
                    "impact": "high",
                    "recommendation": "대체 시나리오 준비"
                },
                {
                    "issue": "원자재 가격 20% 상승",
                    "probability": 0.65,
                    "impact": "medium",
                    "recommendation": "헤징 전략 수립"
                }
            ],
            "timestamp": datetime.now().isoformat()
        }
        
        return Response(
            json.dumps(result, ensure_ascii=False),
            mimetype='application/json; charset=utf-8'
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/health', methods=['GET'])
def health():
    """서버 상태 확인"""
    return jsonify({
        "status": "healthy",
        "mode": "mock" if MOCK_MODE else "production",
        "timestamp": datetime.now().isoformat()
    })

if __name__ == '__main__':
    print(f"""
    ========================================
    STRIX v2 API Server
    Mode: {'Mock' if MOCK_MODE else 'Production'}
    URL: http://localhost:5000
    ========================================
    """)
    app.run(host='0.0.0.0', port=5000, debug=True)