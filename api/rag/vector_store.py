"""
Vector Store module for STRIX v2
Handles Supabase vector database operations
"""
from typing import List, Dict, Any, Optional
from langchain_core.documents import Document
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import SupabaseVectorStore
from supabase import create_client, Client
import logging
from ..config import config

logger = logging.getLogger(__name__)

class STRIXVectorStore:
    """Manages vector store operations for STRIX RAG system"""
    
    def __init__(self):
        """Initialize vector store with Supabase"""
        self.client: Optional[Client] = None
        self.vector_store: Optional[SupabaseVectorStore] = None
        self.embeddings = None
        
        if not config.MOCK_MODE:
            self._initialize_store()
    
    def _initialize_store(self):
        """Initialize Supabase client and vector store"""
        try:
            # Initialize Supabase client
            self.client = create_client(
                config.SUPABASE_URL,
                config.SUPABASE_KEY
            )
            
            # Initialize embeddings
            self.embeddings = OpenAIEmbeddings(
                model=config.EMBEDDING_MODEL,
                openai_api_key=config.OPENAI_API_KEY
            )
            
            # Initialize vector store
            self.vector_store = SupabaseVectorStore(
                client=self.client,
                embedding=self.embeddings,
                table_name=config.VECTOR_COLLECTION_NAME,
                query_name=f"match_{config.VECTOR_COLLECTION_NAME}"
            )
            
            logger.info("Vector store initialized successfully")
            
        except Exception as e:
            logger.error(f"Failed to initialize vector store: {e}")
            raise
    
    def add_documents(self, documents: List[Document]) -> List[str]:
        """
        Add documents to vector store
        
        Args:
            documents: List of documents to add
            
        Returns:
            List of document IDs
        """
        if config.MOCK_MODE:
            return [f"mock_id_{i}" for i in range(len(documents))]
        
        try:
            ids = self.vector_store.add_documents(documents)
            logger.info(f"Added {len(ids)} documents to vector store")
            return ids
        except Exception as e:
            logger.error(f"Failed to add documents: {e}")
            raise
    
    def similarity_search(
        self, 
        query: str, 
        k: int = None,
        filter: Optional[Dict[str, Any]] = None
    ) -> List[Document]:
        """
        Search for similar documents
        
        Args:
            query: Search query
            k: Number of results to return
            filter: Optional metadata filter
            
        Returns:
            List of relevant documents
        """
        if config.MOCK_MODE:
            return self._mock_search(query, k)
        
        k = k or config.MAX_SEARCH_RESULTS
        
        try:
            results = self.vector_store.similarity_search(
                query,
                k=k,
                filter=filter
            )
            logger.info(f"Found {len(results)} similar documents")
            return results
        except Exception as e:
            logger.error(f"Search failed: {e}")
            return []
    
    def similarity_search_with_score(
        self, 
        query: str, 
        k: int = None,
        filter: Optional[Dict[str, Any]] = None
    ) -> List[tuple[Document, float]]:
        """
        Search with relevance scores
        
        Args:
            query: Search query
            k: Number of results
            filter: Optional metadata filter
            
        Returns:
            List of (document, score) tuples
        """
        if config.MOCK_MODE:
            docs = self._mock_search(query, k)
            return [(doc, 0.95 - i*0.05) for i, doc in enumerate(docs)]
        
        k = k or config.MAX_SEARCH_RESULTS
        
        try:
            results = self.vector_store.similarity_search_with_score(
                query,
                k=k,
                filter=filter
            )
            # Filter by minimum relevance score
            filtered_results = [
                (doc, score) for doc, score in results 
                if score >= config.MIN_RELEVANCE_SCORE
            ]
            logger.info(f"Found {len(filtered_results)} relevant documents")
            return filtered_results
        except Exception as e:
            logger.error(f"Search with score failed: {e}")
            return []
    
    def _mock_search(self, query: str, k: int = None) -> List[Document]:
        """Mock search for testing"""
        k = k or config.MAX_SEARCH_RESULTS
        
        mock_docs = [
            Document(
                page_content=f"SK온과 SK이노베이션의 배터리 사업 분사 및 합병 계획에 대한 분석. {query}와 관련된 내용.",
                metadata={
                    "title": "SK온 합병 시너지 분석 보고서",
                    "organization": "전략기획팀",
                    "doc_type": "internal",
                    "date": "2025-08-01",
                    "author": "김철수"
                }
            ),
            Document(
                page_content=f"전고체 배터리 기술 개발 현황 및 2027년 양산 계획. {query}와 관련된 기술 동향.",
                metadata={
                    "title": "차세대 배터리 기술 로드맵",
                    "organization": "R&D센터",
                    "doc_type": "internal",
                    "date": "2025-07-28",
                    "author": "이영희"
                }
            ),
            Document(
                page_content=f"BYD의 블레이드 배터리 및 5분 급속충전 기술 분석. {query}에 대한 경쟁사 동향.",
                metadata={
                    "title": "경쟁사 기술 분석 리포트",
                    "organization": "기술전략팀",
                    "doc_type": "external",
                    "date": "2025-08-03",
                    "source": "Bloomberg"
                }
            )
        ]
        
        return mock_docs[:k]
    
    def delete_documents(self, ids: List[str]) -> bool:
        """
        Delete documents from vector store
        
        Args:
            ids: List of document IDs to delete
            
        Returns:
            Success status
        """
        if config.MOCK_MODE:
            return True
        
        try:
            # Supabase delete implementation
            for doc_id in ids:
                self.client.table(config.VECTOR_COLLECTION_NAME).delete().eq('id', doc_id).execute()
            logger.info(f"Deleted {len(ids)} documents")
            return True
        except Exception as e:
            logger.error(f"Failed to delete documents: {e}")
            return False
    
    def clear(self) -> bool:
        """Clear all documents from vector store"""
        if config.MOCK_MODE:
            return True
        
        try:
            self.client.table(config.VECTOR_COLLECTION_NAME).delete().execute()
            logger.info("Cleared vector store")
            return True
        except Exception as e:
            logger.error(f"Failed to clear vector store: {e}")
            return False