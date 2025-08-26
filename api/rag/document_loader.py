"""
Document Loader module for STRIX v2
Handles loading and processing various document types
"""
from typing import List, Dict, Any, Optional
from pathlib import Path
import logging
from langchain_core.documents import Document
from langchain_community.document_loaders import (
    PyPDFLoader,
    Docx2txtLoader,
    UnstructuredExcelLoader,
    TextLoader,
    WebBaseLoader
)
from langchain_text_splitters import RecursiveCharacterTextSplitter
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from ..config import config

logger = logging.getLogger(__name__)

class STRIXDocumentLoader:
    """Handles document loading and processing for STRIX RAG system"""
    
    def __init__(self):
        """Initialize document loader with text splitter"""
        self.text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=config.CHUNK_SIZE,
            chunk_overlap=config.CHUNK_OVERLAP,
            length_function=len,
            add_start_index=True,
        )
        
        # Supported file extensions
        self.supported_extensions = {
            '.pdf': self._load_pdf,
            '.docx': self._load_docx,
            '.doc': self._load_docx,
            '.xlsx': self._load_excel,
            '.xls': self._load_excel,
            '.txt': self._load_text,
            '.md': self._load_text
        }
    
    def load_document(self, file_path: str, metadata: Optional[Dict[str, Any]] = None) -> List[Document]:
        """
        Load a document from file path
        
        Args:
            file_path: Path to the document
            metadata: Optional metadata to attach
            
        Returns:
            List of document chunks
        """
        path = Path(file_path)
        
        if not path.exists():
            logger.error(f"File not found: {file_path}")
            return []
        
        # Get file extension
        ext = path.suffix.lower()
        
        if ext not in self.supported_extensions:
            logger.warning(f"Unsupported file type: {ext}")
            return []
        
        try:
            # Load document based on type
            documents = self.supported_extensions[ext](str(path))
            
            # Add metadata
            if metadata:
                for doc in documents:
                    doc.metadata.update(metadata)
            
            # Add common metadata
            for doc in documents:
                doc.metadata.update({
                    'source': str(path),
                    'file_name': path.name,
                    'file_type': ext[1:],
                    'loaded_at': datetime.now().isoformat()
                })
            
            # Split documents into chunks
            chunks = self.text_splitter.split_documents(documents)
            
            logger.info(f"Loaded {len(chunks)} chunks from {file_path}")
            return chunks
            
        except Exception as e:
            logger.error(f"Failed to load document {file_path}: {e}")
            return []
    
    def _load_pdf(self, file_path: str) -> List[Document]:
        """Load PDF document"""
        loader = PyPDFLoader(file_path)
        return loader.load()
    
    def _load_docx(self, file_path: str) -> List[Document]:
        """Load Word document"""
        loader = Docx2txtLoader(file_path)
        return loader.load()
    
    def _load_excel(self, file_path: str) -> List[Document]:
        """Load Excel document"""
        try:
            # Read Excel file
            df = pd.read_excel(file_path, sheet_name=None)
            documents = []
            
            # Process each sheet
            for sheet_name, sheet_df in df.items():
                # Convert dataframe to text
                text = f"Sheet: {sheet_name}\n"
                text += sheet_df.to_string()
                
                doc = Document(
                    page_content=text,
                    metadata={'sheet_name': sheet_name}
                )
                documents.append(doc)
            
            return documents
            
        except Exception as e:
            logger.error(f"Failed to load Excel file: {e}")
            # Fallback to unstructured loader
            loader = UnstructuredExcelLoader(file_path)
            return loader.load()
    
    def _load_text(self, file_path: str) -> List[Document]:
        """Load text file"""
        loader = TextLoader(file_path, encoding='utf-8')
        return loader.load()
    
    def load_web_content(self, url: str, metadata: Optional[Dict[str, Any]] = None) -> List[Document]:
        """
        Load content from web URL
        
        Args:
            url: Web URL to load
            metadata: Optional metadata
            
        Returns:
            List of document chunks
        """
        try:
            loader = WebBaseLoader(url)
            documents = loader.load()
            
            # Add metadata
            if metadata:
                for doc in documents:
                    doc.metadata.update(metadata)
            
            # Add URL metadata
            for doc in documents:
                doc.metadata.update({
                    'source': url,
                    'source_type': 'web',
                    'loaded_at': datetime.now().isoformat()
                })
            
            # Split into chunks
            chunks = self.text_splitter.split_documents(documents)
            
            logger.info(f"Loaded {len(chunks)} chunks from {url}")
            return chunks
            
        except Exception as e:
            logger.error(f"Failed to load web content from {url}: {e}")
            return []
    
    def load_directory(self, directory_path: str, recursive: bool = True) -> List[Document]:
        """
        Load all supported documents from a directory
        
        Args:
            directory_path: Path to directory
            recursive: Whether to search recursively
            
        Returns:
            List of all document chunks
        """
        path = Path(directory_path)
        
        if not path.is_dir():
            logger.error(f"Directory not found: {directory_path}")
            return []
        
        all_documents = []
        
        # Get all files
        if recursive:
            files = path.rglob('*')
        else:
            files = path.glob('*')
        
        # Process each file
        for file_path in files:
            if file_path.is_file() and file_path.suffix.lower() in self.supported_extensions:
                documents = self.load_document(str(file_path))
                all_documents.extend(documents)
        
        logger.info(f"Loaded {len(all_documents)} chunks from directory {directory_path}")
        return all_documents
    
    def create_document_from_text(
        self, 
        text: str, 
        metadata: Optional[Dict[str, Any]] = None
    ) -> List[Document]:
        """
        Create documents from raw text
        
        Args:
            text: Raw text content
            metadata: Optional metadata
            
        Returns:
            List of document chunks
        """
        doc = Document(
            page_content=text,
            metadata=metadata or {}
        )
        
        # Add timestamp
        doc.metadata['created_at'] = datetime.now().isoformat()
        
        # Split into chunks
        chunks = self.text_splitter.split_documents([doc])
        
        return chunks