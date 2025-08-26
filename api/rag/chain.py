"""
RAG Chain module for STRIX v2
Implements the core RAG pipeline using LangGraph
"""
from typing import List, Dict, Any, Optional, TypedDict
from langchain_core.documents import Document
from langchain_core.prompts import PromptTemplate, ChatPromptTemplate
from langchain_openai import ChatOpenAI
from langchain_google_genai import ChatGoogleGenerativeAI
from langgraph.graph import StateGraph, START
import logging
from datetime import datetime
from .vector_store import STRIXVectorStore
from ..config import config

logger = logging.getLogger(__name__)

# State definition for LangGraph
class RAGState(TypedDict):
    """State for RAG processing"""
    question: str
    doc_type: str
    context: List[Document]
    internal_docs: List[Document]
    external_docs: List[Document]
    answer: str
    confidence: float
    sources: List[Dict[str, Any]]

class STRIXRAGChain:
    """Main RAG chain for STRIX system"""
    
    def __init__(self):
        """Initialize RAG chain components"""
        self.vector_store = STRIXVectorStore()
        self.llm = self._initialize_llm()
        self.graph = self._build_graph()
        
        # Prompts
        self.qa_prompt = self._create_qa_prompt()
        self.query_analysis_prompt = self._create_query_analysis_prompt()
    
    def _initialize_llm(self):
        """Initialize LLM based on configuration"""
        if config.MOCK_MODE:
            return None
        
        if config.LLM_PROVIDER == 'openai':
            return ChatOpenAI(
                model="gpt-4-turbo-preview",
                temperature=config.TEMPERATURE,
                max_tokens=config.MAX_TOKENS,
                api_key=config.OPENAI_API_KEY
            )
        elif config.LLM_PROVIDER == 'google':
            return ChatGoogleGenerativeAI(
                model="gemini-2.0-flash",
                temperature=config.TEMPERATURE,
                max_tokens=config.MAX_TOKENS,
                google_api_key=config.GOOGLE_API_KEY
            )
        else:
            raise ValueError(f"Unsupported LLM provider: {config.LLM_PROVIDER}")
    
    def _create_qa_prompt(self) -> ChatPromptTemplate:
        """Create Q&A prompt template"""
        template = """You are STRIX, an intelligent assistant for battery industry analysis.
        
Use the following context to answer the question. If you don't know the answer, 
say so clearly. Provide specific details and cite sources when possible.

Context from internal documents:
{internal_context}

Context from external sources:
{external_context}

Question: {question}

Please provide a comprehensive answer in Korean, including:
1. Direct answer to the question
2. Supporting evidence from the documents
3. Any relevant additional insights
4. Confidence level in your answer

Answer:"""
        
        return ChatPromptTemplate.from_template(template)
    
    def _create_query_analysis_prompt(self) -> ChatPromptTemplate:
        """Create query analysis prompt for better retrieval"""
        template = """Analyze the following question and generate an optimized search query.
Consider synonyms, related terms, and key concepts that might help find relevant documents.

Original Question: {question}

Generate a search query that will help find the most relevant documents.
Focus on key terms and concepts.

Optimized Query:"""
        
        return ChatPromptTemplate.from_template(template)
    
    def _build_graph(self) -> StateGraph:
        """Build LangGraph processing pipeline"""
        
        def analyze_query(state: RAGState) -> Dict:
            """Analyze and optimize query for retrieval"""
            if config.MOCK_MODE or not self.llm:
                return {"question": state["question"]}
            
            try:
                # Generate optimized query
                prompt = self.query_analysis_prompt.invoke({
                    "question": state["question"]
                })
                response = self.llm.invoke(prompt)
                optimized_query = response.content
                
                logger.info(f"Optimized query: {optimized_query}")
                return {"question": optimized_query}
            except Exception as e:
                logger.error(f"Query analysis failed: {e}")
                return {"question": state["question"]}
        
        def retrieve_documents(state: RAGState) -> Dict:
            """Retrieve relevant documents"""
            doc_type = state.get("doc_type", "both")
            
            # Prepare filters based on document type
            internal_filter = {"doc_type": "internal"} if doc_type in ["internal", "both"] else None
            external_filter = {"doc_type": "external"} if doc_type in ["external", "both"] else None
            
            # Search internal documents
            internal_docs = []
            if internal_filter:
                internal_results = self.vector_store.similarity_search_with_score(
                    state["question"],
                    k=config.MAX_SEARCH_RESULTS // 2,
                    filter=internal_filter
                )
                internal_docs = [doc for doc, score in internal_results]
            
            # Search external documents  
            external_docs = []
            if external_filter:
                external_results = self.vector_store.similarity_search_with_score(
                    state["question"],
                    k=config.MAX_SEARCH_RESULTS // 2,
                    filter=external_filter
                )
                external_docs = [doc for doc, score in external_results]
            
            # Combine all documents
            all_docs = internal_docs + external_docs
            
            logger.info(f"Retrieved {len(internal_docs)} internal, {len(external_docs)} external documents")
            
            return {
                "context": all_docs,
                "internal_docs": internal_docs,
                "external_docs": external_docs
            }
        
        def generate_answer(state: RAGState) -> Dict:
            """Generate answer using LLM"""
            if config.MOCK_MODE or not self.llm:
                return self._mock_generate_answer(state)
            
            try:
                # Prepare context
                internal_context = "\n\n".join([
                    f"[{doc.metadata.get('title', 'Document')}]\n{doc.page_content}"
                    for doc in state["internal_docs"][:3]
                ])
                
                external_context = "\n\n".join([
                    f"[{doc.metadata.get('title', 'Document')}]\n{doc.page_content}"
                    for doc in state["external_docs"][:3]
                ])
                
                # Generate answer
                prompt = self.qa_prompt.invoke({
                    "internal_context": internal_context or "No internal documents found",
                    "external_context": external_context or "No external documents found",
                    "question": state["question"]
                })
                
                response = self.llm.invoke(prompt)
                answer = response.content
                
                # Calculate confidence (simplified)
                confidence = min(0.95, 0.7 + (len(state["context"]) * 0.05))
                
                # Extract sources
                sources = self._extract_sources(state["context"])
                
                return {
                    "answer": answer,
                    "confidence": confidence,
                    "sources": sources
                }
                
            except Exception as e:
                logger.error(f"Answer generation failed: {e}")
                return {
                    "answer": "죄송합니다. 답변 생성 중 오류가 발생했습니다.",
                    "confidence": 0.0,
                    "sources": []
                }
        
        # Build graph
        graph_builder = StateGraph(RAGState)
        
        # Add nodes
        graph_builder.add_node("analyze_query", analyze_query)
        graph_builder.add_node("retrieve", retrieve_documents)
        graph_builder.add_node("generate", generate_answer)
        
        # Add edges
        graph_builder.add_edge(START, "analyze_query")
        graph_builder.add_edge("analyze_query", "retrieve")
        graph_builder.add_edge("retrieve", "generate")
        
        return graph_builder.compile()
    
    def _mock_generate_answer(self, state: RAGState) -> Dict:
        """Generate mock answer for testing"""
        answer = f"""SK온과 SK이노베이션의 배터리 사업 통합과 관련하여 다음과 같은 정보를 찾았습니다:

1. **합병 계획**: 2025년 11월 1일 예정으로 SK온과 SK이노베이션의 배터리 사업부가 통합됩니다.

2. **시너지 효과**: 5년간 약 20조원의 시너지 효과가 예상되며, 주로 R&D 효율화와 생산 규모 확대에서 발생할 것으로 분석됩니다.

3. **기술 개발**: 전고체 배터리는 2027년 양산을 목표로 개발 중이며, NCM 9.5.5 배터리는 2026년 상반기 출시 예정입니다.

4. **경쟁사 대응**: BYD의 블레이드 배터리와 5분 급속충전 기술에 대응하기 위한 자체 기술 개발이 진행 중입니다.

관련 질문: {state['question']}
문서 참조: 내부 {len(state['internal_docs'])}건, 외부 {len(state['external_docs'])}건"""
        
        sources = self._extract_sources(state["context"])
        confidence = 0.92
        
        return {
            "answer": answer,
            "confidence": confidence,
            "sources": sources
        }
    
    def _extract_sources(self, documents: List[Document]) -> List[Dict[str, Any]]:
        """Extract source information from documents"""
        sources = []
        
        for doc in documents[:5]:  # Top 5 sources
            source = {
                "title": doc.metadata.get("title", "Unknown"),
                "organization": doc.metadata.get("organization", "Unknown"),
                "date": doc.metadata.get("date", datetime.now().strftime("%Y-%m-%d")),
                "type": doc.metadata.get("doc_type", "unknown"),
                "relevance": doc.metadata.get("relevance", 0.8)
            }
            sources.append(source)
        
        return sources
    
    def invoke(self, question: str, doc_type: str = "both") -> Dict[str, Any]:
        """
        Process a question through the RAG pipeline
        
        Args:
            question: User's question
            doc_type: Type of documents to search ("internal", "external", "both")
            
        Returns:
            Dictionary with answer, confidence, sources, etc.
        """
        try:
            # Initialize state
            initial_state = {
                "question": question,
                "doc_type": doc_type,
                "context": [],
                "internal_docs": [],
                "external_docs": [],
                "answer": "",
                "confidence": 0.0,
                "sources": []
            }
            
            # Run graph
            result = self.graph.invoke(initial_state)
            
            # Format response
            response = {
                "answer": result["answer"],
                "confidence": result["confidence"],
                "internal_docs": len(result["internal_docs"]),
                "external_docs": len(result["external_docs"]),
                "sources": result["sources"],
                "timestamp": datetime.now().isoformat()
            }
            
            logger.info(f"RAG query processed successfully")
            return response
            
        except Exception as e:
            logger.error(f"RAG processing failed: {e}")
            return {
                "answer": "처리 중 오류가 발생했습니다.",
                "confidence": 0.0,
                "internal_docs": 0,
                "external_docs": 0,
                "sources": [],
                "error": str(e),
                "timestamp": datetime.now().isoformat()
            }