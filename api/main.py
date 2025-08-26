"""
STRIX v2 FastAPI Server
Modern async API server with LangChain RAG integration
"""
from fastapi import FastAPI, HTTPException, UploadFile, File, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import uvicorn
import logging
from datetime import datetime
import os
import tempfile

from config import config
from rag import STRIXRAGChain, STRIXDocumentLoader, STRIXVectorStore

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="STRIX v2 RAG API",
    description="Battery Industry Intelligence System with RAG",
    version="2.0.0"
)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize RAG components
rag_chain = STRIXRAGChain()
document_loader = STRIXDocumentLoader()
vector_store = STRIXVectorStore()

# Pydantic models for request/response
class QueryRequest(BaseModel):
    question: str
    doc_type: Optional[str] = "both"
    max_results: Optional[int] = 10
    include_sources: Optional[bool] = True

class QueryResponse(BaseModel):
    answer: str
    confidence: float
    internal_docs: int
    external_docs: int
    sources: List[Dict[str, Any]]
    timestamp: str

class DocumentUploadResponse(BaseModel):
    status: str
    message: str
    document_count: int
    chunk_count: int

class FeedbackRequest(BaseModel):
    feedback: str
    question: Optional[str] = None
    answer: Optional[str] = None
    rating: Optional[int] = None

# API Endpoints
@app.get("/")
async def root():
    """Root endpoint"""
    return {
        "name": "STRIX v2 RAG API",
        "status": "operational",
        "mode": "mock" if config.MOCK_MODE else "production",
        "timestamp": datetime.now().isoformat()
    }

@app.get("/api/health")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "mode": "mock" if config.MOCK_MODE else "production",
        "components": {
            "rag_chain": "operational",
            "vector_store": "operational",
            "document_loader": "operational"
        },
        "timestamp": datetime.now().isoformat()
    }

@app.post("/api/query", response_model=QueryResponse)
async def query_rag(request: QueryRequest):
    """
    Main RAG query endpoint
    Process user questions using the RAG pipeline
    """
    try:
        logger.info(f"Processing query: {request.question}")
        
        # Process through RAG chain
        result = rag_chain.invoke(
            question=request.question,
            doc_type=request.doc_type
        )
        
        # Format response
        response = QueryResponse(
            answer=result.get("answer", ""),
            confidence=result.get("confidence", 0.0),
            internal_docs=result.get("internal_docs", 0),
            external_docs=result.get("external_docs", 0),
            sources=result.get("sources", []) if request.include_sources else [],
            timestamp=result.get("timestamp", datetime.now().isoformat())
        )
        
        return response
        
    except Exception as e:
        logger.error(f"Query processing failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/documents/upload")
async def upload_document(
    file: UploadFile = File(...),
    doc_type: str = "internal",
    organization: Optional[str] = None,
    background_tasks: BackgroundTasks = BackgroundTasks()
):
    """
    Upload and process a document
    Adds document to vector store for RAG
    """
    try:
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=file.filename) as tmp_file:
            content = await file.read()
            tmp_file.write(content)
            tmp_path = tmp_file.name
        
        # Process document
        metadata = {
            "doc_type": doc_type,
            "organization": organization or "Unknown",
            "uploaded_at": datetime.now().isoformat(),
            "file_name": file.filename
        }
        
        # Load and split document
        documents = document_loader.load_document(tmp_path, metadata)
        
        if not documents:
            raise HTTPException(status_code=400, detail="Failed to process document")
        
        # Add to vector store (can be done in background)
        background_tasks.add_task(vector_store.add_documents, documents)
        
        # Clean up temp file
        background_tasks.add_task(os.unlink, tmp_path)
        
        return DocumentUploadResponse(
            status="success",
            message=f"Document '{file.filename}' uploaded successfully",
            document_count=1,
            chunk_count=len(documents)
        )
        
    except Exception as e:
        logger.error(f"Document upload failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/documents/batch")
async def batch_upload_documents(
    files: List[UploadFile] = File(...),
    doc_type: str = "internal",
    background_tasks: BackgroundTasks = BackgroundTasks()
):
    """
    Upload multiple documents at once
    """
    try:
        total_chunks = 0
        
        for file in files:
            # Save file temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix=file.filename) as tmp_file:
                content = await file.read()
                tmp_file.write(content)
                tmp_path = tmp_file.name
            
            # Process document
            metadata = {
                "doc_type": doc_type,
                "uploaded_at": datetime.now().isoformat(),
                "file_name": file.filename
            }
            
            documents = document_loader.load_document(tmp_path, metadata)
            total_chunks += len(documents)
            
            # Add to vector store in background
            background_tasks.add_task(vector_store.add_documents, documents)
            background_tasks.add_task(os.unlink, tmp_path)
        
        return DocumentUploadResponse(
            status="success",
            message=f"Uploaded {len(files)} documents successfully",
            document_count=len(files),
            chunk_count=total_chunks
        )
        
    except Exception as e:
        logger.error(f"Batch upload failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/feedback")
async def submit_feedback(request: FeedbackRequest):
    """
    Submit feedback for continuous improvement
    """
    try:
        # In production, save to database
        feedback_data = {
            "feedback": request.feedback,
            "question": request.question,
            "answer": request.answer,
            "rating": request.rating,
            "timestamp": datetime.now().isoformat(),
            "feedback_id": f"FB_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        }
        
        logger.info(f"Feedback received: {feedback_data['feedback_id']}")
        
        return {
            "status": "success",
            "message": "피드백이 저장되었습니다",
            "feedback_id": feedback_data["feedback_id"]
        }
        
    except Exception as e:
        logger.error(f"Feedback submission failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/documents/search")
async def search_documents(
    query: str,
    doc_type: Optional[str] = None,
    limit: int = 10
):
    """
    Search documents directly without generating answer
    """
    try:
        filter = {"doc_type": doc_type} if doc_type else None
        
        results = vector_store.similarity_search_with_score(
            query=query,
            k=limit,
            filter=filter
        )
        
        documents = []
        for doc, score in results:
            documents.append({
                "content": doc.page_content[:500] + "..." if len(doc.page_content) > 500 else doc.page_content,
                "metadata": doc.metadata,
                "relevance_score": score
            })
        
        return {
            "query": query,
            "document_count": len(documents),
            "documents": documents,
            "timestamp": datetime.now().isoformat()
        }
        
    except Exception as e:
        logger.error(f"Document search failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/api/documents/clear")
async def clear_documents():
    """
    Clear all documents from vector store (admin only)
    """
    try:
        success = vector_store.clear()
        
        if success:
            return {
                "status": "success",
                "message": "All documents cleared from vector store"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to clear documents")
            
    except Exception as e:
        logger.error(f"Document clear failed: {e}")
        raise HTTPException(status_code=500, detail=str(e))

# Main entry point
if __name__ == "__main__":
    # Validate configuration
    try:
        config.validate()
    except ValueError as e:
        logger.error(f"Configuration error: {e}")
        logger.info("Running in MOCK_MODE due to configuration issues")
        config.MOCK_MODE = True
    
    # Print startup message
    print(f"""
    ╔══════════════════════════════════════╗
    ║       STRIX v2 RAG API Server        ║
    ║                                      ║
    ║  Mode: {'MOCK' if config.MOCK_MODE else 'PRODUCTION':15}       ║
    ║  Host: {config.HOST:15}       ║
    ║  Port: {config.PORT:15}       ║
    ║                                      ║
    ║  LangChain + Supabase + FastAPI     ║
    ╚══════════════════════════════════════╝
    """)
    
    # Run server
    uvicorn.run(
        app,
        host=config.HOST,
        port=config.PORT,
        reload=config.DEBUG_MODE
    )