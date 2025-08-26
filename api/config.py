"""
Configuration module for STRIX v2 RAG system
"""
import os
from dotenv import load_dotenv
from typing import Optional

# Load environment variables
load_dotenv()

class Config:
    """Application configuration"""
    
    # LLM Settings
    OPENAI_API_KEY: str = os.getenv('OPENAI_API_KEY', '')
    GOOGLE_API_KEY: str = os.getenv('GOOGLE_API_KEY', '')
    LLM_PROVIDER: str = os.getenv('LLM_PROVIDER', 'openai')  # 'openai' or 'google'
    
    # Supabase Settings
    SUPABASE_URL: str = os.getenv('SUPABASE_URL', '')
    SUPABASE_KEY: str = os.getenv('SUPABASE_KEY', '')
    SUPABASE_SERVICE_KEY: str = os.getenv('SUPABASE_SERVICE_KEY', '')
    
    # Vector Store Settings
    VECTOR_COLLECTION_NAME: str = os.getenv('VECTOR_COLLECTION_NAME', 'strix_documents')
    EMBEDDING_MODEL: str = os.getenv('EMBEDDING_MODEL', 'text-embedding-3-small')
    CHUNK_SIZE: int = int(os.getenv('CHUNK_SIZE', '1000'))
    CHUNK_OVERLAP: int = int(os.getenv('CHUNK_OVERLAP', '200'))
    
    # Application Settings
    MOCK_MODE: bool = os.getenv('MOCK_MODE', 'false').lower() == 'true'
    DEBUG_MODE: bool = os.getenv('DEBUG_MODE', 'true').lower() == 'true'
    PORT: int = int(os.getenv('PORT', '5000'))
    HOST: str = os.getenv('HOST', '0.0.0.0')
    
    # RAG Settings
    MAX_SEARCH_RESULTS: int = int(os.getenv('MAX_SEARCH_RESULTS', '10'))
    MIN_RELEVANCE_SCORE: float = float(os.getenv('MIN_RELEVANCE_SCORE', '0.7'))
    TEMPERATURE: float = float(os.getenv('TEMPERATURE', '0.7'))
    MAX_TOKENS: int = int(os.getenv('MAX_TOKENS', '2000'))
    
    # Document Types
    INTERNAL_DOC_TYPES = ['report', 'analysis', 'memo', 'presentation']
    EXTERNAL_DOC_TYPES = ['news', 'research', 'competitor', 'policy']
    
    @classmethod
    def validate(cls) -> bool:
        """Validate configuration"""
        if not cls.MOCK_MODE:
            # Check required API keys
            if cls.LLM_PROVIDER == 'openai' and not cls.OPENAI_API_KEY:
                raise ValueError("OPENAI_API_KEY is required when MOCK_MODE is false")
            if cls.LLM_PROVIDER == 'google' and not cls.GOOGLE_API_KEY:
                raise ValueError("GOOGLE_API_KEY is required when MOCK_MODE is false")
                
            # Check Supabase settings
            if not cls.SUPABASE_URL or not cls.SUPABASE_KEY:
                raise ValueError("Supabase credentials are required when MOCK_MODE is false")
        
        return True

# Create config instance
config = Config()