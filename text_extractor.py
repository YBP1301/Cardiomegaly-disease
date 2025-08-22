import os
import logging
from pathlib import Path
from typing import Dict, Optional, Tuple
import magic

# Document processing imports
import PyPDF2
import docx2txt
from pptx import Presentation
import extract_msg

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class TextExtractor:
    """Handles text extraction from various file formats"""
    
    def __init__(self):
        self.supported_formats = {
            '.pdf': self._extract_from_pdf,
            '.docx': self._extract_from_docx,
            '.doc': self._extract_from_doc,
            '.pptx': self._extract_from_pptx,
            '.ppt': self._extract_from_ppt,
            '.msg': self._extract_from_msg,
            '.txt': self._extract_from_txt
        }
    
    def extract_text(self, file_path: str) -> Tuple[str, str]:
        """
        Extract text from a file
        
        Args:
            file_path: Path to the file
            
        Returns:
            Tuple of (extracted_text, detected_file_type)
        """
        try:
            file_path = Path(file_path)
            
            if not file_path.exists():
                raise FileNotFoundError(f"File not found: {file_path}")
            
            # Detect file type
            file_extension = file_path.suffix.lower()
            detected_type = self._detect_file_type(file_path)
            
            # Extract text based on file type
            if file_extension in self.supported_formats:
                text = self.supported_formats[file_extension](file_path)
                return text.strip(), detected_type
            else:
                logger.warning(f"Unsupported file format: {file_extension}")
                return "", detected_type
                
        except Exception as e:
            logger.error(f"Error extracting text from {file_path}: {str(e)}")
            return "", "unknown"
    
    def _detect_file_type(self, file_path: Path) -> str:
        """Detect file type using magic numbers and extension"""
        try:
            mime_type = magic.from_file(str(file_path), mime=True)
            extension = file_path.suffix.lower()
            
            # Map MIME types to readable format
            type_mapping = {
                'application/pdf': 'PDF',
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'Word Document (DOCX)',
                'application/msword': 'Word Document (DOC)',
                'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PowerPoint (PPTX)',
                'application/vnd.ms-powerpoint': 'PowerPoint (PPT)',
                'application/vnd.ms-outlook': 'Outlook Message (MSG)',
                'text/plain': 'Text File'
            }
            
            return type_mapping.get(mime_type, f"Unknown ({extension})")
            
        except Exception:
            return f"Unknown ({file_path.suffix})"
    
    def _extract_from_pdf(self, file_path: Path) -> str:
        """Extract text from PDF files"""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
                return text
        except Exception as e:
            logger.error(f"Error extracting PDF {file_path}: {str(e)}")
            return ""
    
    def _extract_from_docx(self, file_path: Path) -> str:
        """Extract text from DOCX files"""
        try:
            return docx2txt.process(str(file_path))
        except Exception as e:
            logger.error(f"Error extracting DOCX {file_path}: {str(e)}")
            return ""
    
    def _extract_from_doc(self, file_path: Path) -> str:
        """Extract text from DOC files (legacy format)"""
        try:
            # For .doc files, we'll try to use python-docx2txt
            # Note: This might not work perfectly for all .doc files
            return docx2txt.process(str(file_path))
        except Exception as e:
            logger.error(f"Error extracting DOC {file_path}: {str(e)}")
            return ""
    
    def _extract_from_pptx(self, file_path: Path) -> str:
        """Extract text from PPTX files"""
        try:
            presentation = Presentation(file_path)
            text = ""
            
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text += shape.text + "\n"
                        
            return text
        except Exception as e:
            logger.error(f"Error extracting PPTX {file_path}: {str(e)}")
            return ""
    
    def _extract_from_ppt(self, file_path: Path) -> str:
        """Extract text from PPT files (legacy format)"""
        try:
            # For legacy PPT files, try using python-pptx
            # Note: This might not work for all legacy PPT files
            presentation = Presentation(file_path)
            text = ""
            
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text += shape.text + "\n"
                        
            return text
        except Exception as e:
            logger.error(f"Error extracting PPT {file_path}: {str(e)}")
            return ""
    
    def _extract_from_msg(self, file_path: Path) -> str:
        """Extract text from MSG files"""
        try:
            msg = extract_msg.Message(str(file_path))
            text = ""
            
            # Extract subject, body, sender, recipients
            if msg.subject:
                text += f"Subject: {msg.subject}\n"
            if msg.sender:
                text += f"From: {msg.sender}\n"
            if msg.to:
                text += f"To: {msg.to}\n"
            if msg.body:
                text += f"Body: {msg.body}\n"
                
            return text
        except Exception as e:
            logger.error(f"Error extracting MSG {file_path}: {str(e)}")
            return ""
    
    def _extract_from_txt(self, file_path: Path) -> str:
        """Extract text from TXT files"""
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                return file.read()
        except Exception as e:
            logger.error(f"Error extracting TXT {file_path}: {str(e)}")
            return ""