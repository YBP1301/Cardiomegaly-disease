#!/usr/bin/env python3
"""
Document Processing Pipeline - All-in-One Solution
=================================================

A comprehensive solution for document analysis, text extraction, LLM-based summarization, 
and duplicate detection with Excel export - all integrated in a single file.

Features:
- Multi-format text extraction (PDF, Word, PowerPoint, MSG, TXT)
- Azure OpenAI integration for summarization and analysis
- Advanced duplicate detection using content similarity
- Professional Excel export with formatting

Usage:
    python document_processor_all_in_one.py --folder /path/to/documents
    python document_processor_all_in_one.py --file /path/to/document.pdf
    python document_processor_all_in_one.py --folder /path/to/documents --output results.xlsx --stats

Requirements:
    pip install python-docx2txt PyPDF2 python-pptx extract-msg openai azure-identity pandas openpyxl python-magic fuzzywuzzy python-Levenshtein nltk scikit-learn numpy tqdm

Environment Variables:
    AZURE_OPENAI_API_KEY="your-api-key"
    AZURE_OPENAI_ENDPOINT="https://your-resource.openai.azure.com/"
    AZURE_OPENAI_DEPLOYMENT="gpt-4"
"""

import os
import sys
import argparse
import logging
import hashlib
import json
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from datetime import datetime
from difflib import SequenceMatcher
import concurrent.futures
from threading import Lock

# Document processing imports
try:
    import PyPDF2
    import docx2txt
    from pptx import Presentation
    import extract_msg
    import magic
    from openai import AzureOpenAI
    import pandas as pd
    from fuzzywuzzy import fuzz
    import nltk
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity
    import numpy as np
    from tqdm import tqdm
except ImportError as e:
    print(f"❌ Missing required package: {e}")
    print("💡 Install with: pip install python-docx2txt PyPDF2 python-pptx extract-msg openai azure-identity pandas openpyxl python-magic fuzzywuzzy python-Levenshtein nltk scikit-learn numpy tqdm")
    sys.exit(1)

# ============================================================================
# CONFIGURATION
# ============================================================================

# Azure OpenAI Configuration
AZURE_OPENAI_CONFIG = {
    "api_key": os.getenv("AZURE_OPENAI_API_KEY", "your-api-key-here"),
    "azure_endpoint": os.getenv("AZURE_OPENAI_ENDPOINT", "https://your-resource.openai.azure.com/"),
    "api_version": "2024-02-15-preview",
    "deployment_name": os.getenv("AZURE_OPENAI_DEPLOYMENT", "gpt-4"),
    "model": "gpt-4"
}

# System prompts for different tasks
SYSTEM_PROMPTS = {
    "summarization": """You are an expert document analyzer. Your task is to:
1. Provide a comprehensive summarization covering all necessary contents
2. Create a brief description with 2-3 bullet points highlighting key aspects
3. Extract file metadata including file_type, version, and relevant tags/keywords

Format your response as JSON:
{
    "summary": "detailed summary here",
    "description": "brief description",
    "bullet_points": ["point 1", "point 2", "point 3"],
    "file_type": "detected file type",
    "version": "version if available or 'N/A'",
    "tags": ["keyword1", "keyword2", "keyword3"]
}""",
    
    "duplicate_analysis": """You are an expert at analyzing document similarity. Compare the provided texts and determine:
1. If they are duplicates (same content with minor differences)
2. Similarity percentage
3. Reason for duplication (exact copy, minor edits, formatting differences, etc.)
4. Which one appears to be the master/original version

Respond in JSON format:
{
    "is_duplicate": true/false,
    "similarity_percentage": 85.5,
    "duplicate_reason": "reason for duplication",
    "master_file": "filename of master version"
}"""
}

# File processing settings
MAX_TEXT_LENGTH = 50000
SUPPORTED_EXTENSIONS = {'.pdf', '.docx', '.doc', '.pptx', '.ppt', '.msg', '.txt'}
DUPLICATE_THRESHOLD = 0.85

# Excel output settings
EXCEL_COLUMNS = [
    'file_name', 'file_type', 'extracted_text', 'summary', 
    'description', 'bullet_points', 'version', 'tags', 
    'duplicates', 'duplicates_percentage', 'master_one'
]

# ============================================================================
# TEXT EXTRACTION CLASS
# ============================================================================

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
        """Extract text from a file"""
        try:
            file_path = Path(file_path)
            
            if not file_path.exists():
                raise FileNotFoundError(f"File not found: {file_path}")
            
            file_extension = file_path.suffix.lower()
            detected_type = self._detect_file_type(file_path)
            
            if file_extension in self.supported_formats:
                text = self.supported_formats[file_extension](file_path)
                return text.strip(), detected_type
            else:
                logging.warning(f"Unsupported file format: {file_extension}")
                return "", detected_type
                
        except Exception as e:
            logging.error(f"Error extracting text from {file_path}: {str(e)}")
            return "", "unknown"
    
    def _detect_file_type(self, file_path: Path) -> str:
        """Detect file type using magic numbers and extension"""
        try:
            mime_type = magic.from_file(str(file_path), mime=True)
            extension = file_path.suffix.lower()
            
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
            logging.error(f"Error extracting PDF {file_path}: {str(e)}")
            return ""
    
    def _extract_from_docx(self, file_path: Path) -> str:
        """Extract text from DOCX files"""
        try:
            return docx2txt.process(str(file_path))
        except Exception as e:
            logging.error(f"Error extracting DOCX {file_path}: {str(e)}")
            return ""
    
    def _extract_from_doc(self, file_path: Path) -> str:
        """Extract text from DOC files"""
        try:
            return docx2txt.process(str(file_path))
        except Exception as e:
            logging.error(f"Error extracting DOC {file_path}: {str(e)}")
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
            logging.error(f"Error extracting PPTX {file_path}: {str(e)}")
            return ""
    
    def _extract_from_ppt(self, file_path: Path) -> str:
        """Extract text from PPT files"""
        try:
            presentation = Presentation(file_path)
            text = ""
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        text += shape.text + "\n"
            return text
        except Exception as e:
            logging.error(f"Error extracting PPT {file_path}: {str(e)}")
            return ""
    
    def _extract_from_msg(self, file_path: Path) -> str:
        """Extract text from MSG files"""
        try:
            msg = extract_msg.Message(str(file_path))
            text = ""
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
            logging.error(f"Error extracting MSG {file_path}: {str(e)}")
            return ""
    
    def _extract_from_txt(self, file_path: Path) -> str:
        """Extract text from TXT files"""
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                return file.read()
        except Exception as e:
            logging.error(f"Error extracting TXT {file_path}: {str(e)}")
            return ""

# ============================================================================
# LLM ANALYZER CLASS
# ============================================================================

class LLMAnalyzer:
    """Handles Azure OpenAI integration for document analysis"""
    
    def __init__(self):
        self.client = AzureOpenAI(
            api_key=AZURE_OPENAI_CONFIG["api_key"],
            azure_endpoint=AZURE_OPENAI_CONFIG["azure_endpoint"],
            api_version=AZURE_OPENAI_CONFIG["api_version"]
        )
        self.deployment_name = AZURE_OPENAI_CONFIG["deployment_name"]
    
    def analyze_document(self, text: str, file_name: str) -> Dict[str, Any]:
        """Analyze document text and extract summary, metadata, and tags"""
        try:
            if len(text) > MAX_TEXT_LENGTH:
                text = text[:MAX_TEXT_LENGTH] + "... [truncated]"
                logging.warning(f"Text truncated for {file_name} due to length")
            
            user_prompt = f"""
            Please analyze the following document text from file: {file_name}
            
            Document content:
            {text}
            
            Provide a comprehensive analysis including summary, description, bullet points, file type, version, and relevant tags.
            """
            
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPTS["summarization"]},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
                max_tokens=1500
            )
            
            response_text = response.choices[0].message.content
            
            try:
                analysis_result = json.loads(response_text)
                return self._validate_analysis_result(analysis_result)
            except json.JSONDecodeError:
                logging.error(f"Failed to parse JSON response for {file_name}")
                return self._create_fallback_result(text, file_name)
                
        except Exception as e:
            logging.error(f"Error analyzing document {file_name}: {str(e)}")
            return self._create_fallback_result(text, file_name)
    
    def _validate_analysis_result(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """Validate and clean analysis result"""
        validated = {
            "summary": result.get("summary", "No summary available"),
            "description": result.get("description", "No description available"),
            "bullet_points": result.get("bullet_points", []),
            "file_type": result.get("file_type", "Unknown"),
            "version": result.get("version", "N/A"),
            "tags": result.get("tags", [])
        }
        
        if isinstance(validated["bullet_points"], str):
            validated["bullet_points"] = [validated["bullet_points"]]
        if isinstance(validated["tags"], str):
            validated["tags"] = [validated["tags"]]
            
        return validated
    
    def _create_fallback_result(self, text: str, file_name: str) -> Dict[str, Any]:
        """Create fallback result when LLM analysis fails"""
        return {
            "summary": f"Text extraction completed for {file_name}. LLM analysis failed.",
            "description": "Document processed but detailed analysis unavailable",
            "bullet_points": ["Text successfully extracted", "Analysis requires manual review"],
            "file_type": "Unknown",
            "version": "N/A",
            "tags": ["unprocessed"]
        }

# ============================================================================
# DUPLICATE DETECTOR CLASS
# ============================================================================

class DuplicateDetector:
    """Handles duplicate detection using multiple similarity algorithms"""
    
    def __init__(self):
        try:
            nltk.download('punkt', quiet=True)
            nltk.download('stopwords', quiet=True)
        except:
            pass
    
    def find_duplicates(self, documents: Dict[str, str]) -> Dict[str, Dict]:
        """Find duplicates among a collection of documents"""
        duplicate_info = {}
        file_paths = list(documents.keys())
        
        # Initialize all files as non-duplicates
        for file_path in file_paths:
            duplicate_info[file_path] = {
                "is_duplicate": False,
                "similarity_percentage": 0.0,
                "duplicate_reason": "No duplicates found",
                "master_file": file_path,
                "duplicates_of": []
            }
        
        # Compare each file with every other file
        for i, file1 in enumerate(file_paths):
            for j, file2 in enumerate(file_paths[i+1:], i+1):
                similarity_result = self._calculate_similarity(
                    documents[file1], documents[file2], file1, file2
                )
                
                if similarity_result["similarity_percentage"] > DUPLICATE_THRESHOLD * 100:
                    master_file = self._determine_master_file(file1, file2, documents)
                    duplicate_file = file2 if master_file == file1 else file1
                    
                    duplicate_info[duplicate_file].update({
                        "is_duplicate": True,
                        "similarity_percentage": similarity_result["similarity_percentage"],
                        "duplicate_reason": similarity_result["reason"],
                        "master_file": master_file,
                        "duplicates_of": [master_file]
                    })
                    
                    if duplicate_file not in duplicate_info[master_file]["duplicates_of"]:
                        duplicate_info[master_file]["duplicates_of"].append(duplicate_file)
        
        return duplicate_info
    
    def _calculate_similarity(self, text1: str, text2: str, file1: str, file2: str) -> Dict:
        """Calculate similarity between two texts using multiple methods"""
        
        # Method 1: Exact hash comparison
        hash1 = hashlib.md5(text1.encode()).hexdigest()
        hash2 = hashlib.md5(text2.encode()).hexdigest()
        
        if hash1 == hash2:
            return {
                "similarity_percentage": 100.0,
                "reason": "Exact duplicate (identical hash)"
            }
        
        # Method 2: Sequence matcher
        seq_similarity = SequenceMatcher(None, text1, text2).ratio() * 100
        
        # Method 3: Fuzzy string matching
        fuzzy_ratio = fuzz.ratio(text1, text2)
        fuzzy_partial = fuzz.partial_ratio(text1, text2)
        fuzzy_token_sort = fuzz.token_sort_ratio(text1, text2)
        fuzzy_token_set = fuzz.token_set_ratio(text1, text2)
        
        # Method 4: TF-IDF cosine similarity
        tfidf_similarity = self._calculate_tfidf_similarity(text1, text2)
        
        # Combine all similarity scores
        similarities = [seq_similarity, fuzzy_ratio, fuzzy_partial, 
                       fuzzy_token_sort, fuzzy_token_set, tfidf_similarity * 100]
        
        weights = [0.15, 0.15, 0.15, 0.15, 0.15, 0.25]
        weighted_similarity = sum(s * w for s, w in zip(similarities, weights))
        
        reason = self._determine_duplicate_reason(similarities)
        
        return {
            "similarity_percentage": round(weighted_similarity, 2),
            "reason": reason
        }
    
    def _calculate_tfidf_similarity(self, text1: str, text2: str) -> float:
        """Calculate TF-IDF cosine similarity"""
        try:
            vectorizer = TfidfVectorizer(stop_words='english', max_features=1000)
            tfidf_matrix = vectorizer.fit_transform([text1, text2])
            similarity = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0]
            return similarity
        except Exception:
            return 0.0
    
    def _determine_duplicate_reason(self, similarities: List[float]) -> str:
        """Determine the reason for duplication based on similarity patterns"""
        seq_sim, fuzzy_ratio, fuzzy_partial, fuzzy_token_sort, fuzzy_token_set, tfidf_sim = similarities
        
        if seq_sim > 95:
            return "Nearly identical content"
        elif fuzzy_ratio > 90:
            return "Minor formatting differences"
        elif fuzzy_token_sort > 85:
            return "Same content, different word order"
        elif fuzzy_token_set > 85:
            return "Similar content with additions/deletions"
        elif tfidf_sim > 85:
            return "Semantically similar content"
        else:
            return "Low similarity - likely not duplicates"
    
    def _determine_master_file(self, file1: str, file2: str, documents: Dict[str, str]) -> str:
        """Determine which file is the master based on various criteria"""
        path1, path2 = Path(file1), Path(file2)
        text1, text2 = documents[file1], documents[file2]
        
        # File with more content
        if len(text1) != len(text2):
            return file1 if len(text1) > len(text2) else file2
        
        # File with earlier modification time
        try:
            stat1, stat2 = path1.stat(), path2.stat()
            if stat1.st_mtime != stat2.st_mtime:
                return file1 if stat1.st_mtime < stat2.st_mtime else file2
        except:
            pass
        
        # File with simpler name
        if len(file1) != len(file2):
            return file1 if len(file1) < len(file2) else file2
        
        # Alphabetical order as fallback
        return file1 if file1 < file2 else file2

# ============================================================================
# EXCEL EXPORTER CLASS
# ============================================================================

class ExcelExporter:
    """Handles Excel export functionality"""
    
    def __init__(self):
        self.columns = EXCEL_COLUMNS
    
    def export_results(self, results: List[Dict[str, Any]], output_path: str = None) -> str:
        """Export analysis results to Excel file"""
        try:
            df = pd.DataFrame(results)
            
            # Ensure all required columns exist
            for col in self.columns:
                if col not in df.columns:
                    df[col] = ""
            
            df = df[self.columns]
            
            if output_path is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = f"document_analysis_results_{timestamp}.xlsx"
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Document Analysis', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Document Analysis']
                self._apply_excel_formatting(workbook, worksheet, df)
            
            logging.info(f"Results exported to: {output_path}")
            return output_path
            
        except Exception as e:
            logging.error(f"Error exporting to Excel: {str(e)}")
            raise
    
    def _apply_excel_formatting(self, workbook, worksheet, df):
        """Apply formatting to the Excel worksheet"""
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        # Header formatting
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
        
        # Column widths
        column_widths = {
            'A': 25, 'B': 15, 'C': 50, 'D': 50, 'E': 30, 'F': 30,
            'G': 12, 'H': 25, 'I': 30, 'J': 15, 'K': 25
        }
        
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
        
        # Borders and alignment
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        cell_alignment = Alignment(vertical="top", wrap_text=True)
        
        for row in worksheet.iter_rows(min_row=1, max_row=len(df) + 1):
            for cell in row:
                cell.border = thin_border
                if cell.row > 1:
                    cell.alignment = cell_alignment
        
        worksheet.freeze_panes = "A2"
        
        # Highlight duplicates
        duplicate_fill = PatternFill(start_color="FFE6E6", end_color="FFE6E6", fill_type="solid")
        
        for row_idx, row in enumerate(df.itertuples(), start=2):
            if hasattr(row, 'duplicates_percentage') and row.duplicates_percentage > 0:
                for col_idx in range(1, len(self.columns) + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    cell.fill = duplicate_fill
    
    def prepare_row_data(self, file_info: Dict[str, Any]) -> Dict[str, Any]:
        """Prepare row data for Excel export"""
        bullet_points = file_info.get('bullet_points', [])
        if isinstance(bullet_points, list):
            bullet_points = '\n'.join([f"• {point}" for point in bullet_points])
        
        tags = file_info.get('tags', [])
        if isinstance(tags, list):
            tags = ', '.join(tags)
        
        extracted_text = file_info.get('extracted_text', '')
        if len(extracted_text) > 1000:
            extracted_text = extracted_text[:1000] + "..."
        
        return {
            'file_name': file_info.get('file_name', ''),
            'file_type': file_info.get('file_type', ''),
            'extracted_text': extracted_text,
            'summary': file_info.get('summary', ''),
            'description': file_info.get('description', ''),
            'bullet_points': bullet_points,
            'version': file_info.get('version', 'N/A'),
            'tags': tags,
            'duplicates': file_info.get('duplicate_reason', 'No duplicates found'),
            'duplicates_percentage': file_info.get('similarity_percentage', 0.0),
            'master_one': file_info.get('master_file', file_info.get('file_name', ''))
        }

# ============================================================================
# MAIN DOCUMENT PROCESSOR CLASS
# ============================================================================

class DocumentProcessor:
    """Main orchestrator for document processing pipeline"""
    
    def __init__(self):
        self.text_extractor = TextExtractor()
        self.llm_analyzer = LLMAnalyzer()
        self.duplicate_detector = DuplicateDetector()
        self.excel_exporter = ExcelExporter()
        self.results_lock = Lock()
    
    def process_folder(self, folder_path: str, output_excel_path: Optional[str] = None) -> str:
        """Process all supported documents in a folder"""
        try:
            folder_path = Path(folder_path)
            if not folder_path.exists():
                raise FileNotFoundError(f"Folder not found: {folder_path}")
            
            logging.info(f"Starting document processing for folder: {folder_path}")
            
            # Step 1: Find all supported files
            supported_files = self._find_supported_files(folder_path)
            logging.info(f"Found {len(supported_files)} supported files")
            
            if not supported_files:
                logging.warning("No supported files found in the folder")
                return ""
            
            # Step 2: Extract text from all files
            logging.info("Extracting text from documents...")
            extracted_documents = self._extract_texts_parallel(supported_files)
            
            # Step 3: Analyze documents with LLM
            logging.info("Analyzing documents with Azure OpenAI...")
            analysis_results = self._analyze_documents_parallel(extracted_documents)
            
            # Step 4: Detect duplicates
            logging.info("Detecting duplicates...")
            duplicate_info = self.duplicate_detector.find_duplicates(
                {path: info['extracted_text'] for path, info in extracted_documents.items()}
            )
            
            # Step 5: Combine all results
            logging.info("Combining results...")
            final_results = self._combine_results(analysis_results, duplicate_info, extracted_documents)
            
            # Step 6: Export to Excel
            logging.info("Exporting to Excel...")
            excel_path = self.excel_exporter.export_results(final_results, output_excel_path)
            
            logging.info(f"Document processing completed. Results saved to: {excel_path}")
            return excel_path
            
        except Exception as e:
            logging.error(f"Error processing folder: {str(e)}")
            raise
    
    def process_single_file(self, file_path: str) -> Dict[str, Any]:
        """Process a single document file"""
        try:
            file_path = Path(file_path)
            
            # Extract text
            extracted_text, file_type = self.text_extractor.extract_text(str(file_path))
            
            if not extracted_text:
                logging.warning(f"No text extracted from {file_path}")
                return self._create_empty_result(str(file_path), file_type)
            
            # Analyze with LLM
            analysis = self.llm_analyzer.analyze_document(extracted_text, file_path.name)
            
            # Combine results
            result = {
                'file_name': file_path.name,
                'file_path': str(file_path),
                'file_type': file_type,
                'extracted_text': extracted_text,
                **analysis,
                'similarity_percentage': 0.0,
                'duplicate_reason': 'Single file - no comparison',
                'master_file': file_path.name
            }
            
            return result
            
        except Exception as e:
            logging.error(f"Error processing single file {file_path}: {str(e)}")
            return self._create_empty_result(str(file_path), "Unknown")
    
    def _find_supported_files(self, folder_path: Path) -> List[Path]:
        """Find all supported files in the folder recursively"""
        supported_files = []
        for file_path in folder_path.rglob("*"):
            if file_path.is_file() and file_path.suffix.lower() in SUPPORTED_EXTENSIONS:
                supported_files.append(file_path)
        return supported_files
    
    def _extract_texts_parallel(self, file_paths: List[Path]) -> Dict[str, Dict[str, Any]]:
        """Extract text from multiple files in parallel"""
        extracted_documents = {}
        
        def extract_single(file_path):
            text, file_type = self.text_extractor.extract_text(str(file_path))
            return str(file_path), {
                'extracted_text': text,
                'file_type': file_type,
                'file_name': file_path.name
            }
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            future_to_file = {executor.submit(extract_single, fp): fp for fp in file_paths}
            
            for future in tqdm(concurrent.futures.as_completed(future_to_file), 
                             total=len(file_paths), desc="Extracting text"):
                try:
                    file_path, result = future.result()
                    extracted_documents[file_path] = result
                except Exception as e:
                    file_path = future_to_file[future]
                    logging.error(f"Error extracting text from {file_path}: {str(e)}")
                    extracted_documents[str(file_path)] = {
                        'extracted_text': '',
                        'file_type': 'Unknown',
                        'file_name': file_path.name
                    }
        
        return extracted_documents
    
    def _analyze_documents_parallel(self, extracted_documents: Dict[str, Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
        """Analyze documents with LLM in parallel"""
        analysis_results = {}
        
        def analyze_single(file_path, doc_info):
            if doc_info['extracted_text']:
                analysis = self.llm_analyzer.analyze_document(
                    doc_info['extracted_text'], 
                    doc_info['file_name']
                )
                return file_path, analysis
            else:
                return file_path, self._create_empty_analysis(doc_info['file_name'])
        
        with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
            future_to_file = {
                executor.submit(analyze_single, fp, info): fp 
                for fp, info in extracted_documents.items()
            }
            
            for future in tqdm(concurrent.futures.as_completed(future_to_file), 
                             total=len(extracted_documents), desc="Analyzing with LLM"):
                try:
                    file_path, analysis = future.result()
                    analysis_results[file_path] = analysis
                except Exception as e:
                    file_path = future_to_file[future]
                    logging.error(f"Error analyzing {file_path}: {str(e)}")
                    analysis_results[file_path] = self._create_empty_analysis(
                        extracted_documents[file_path]['file_name']
                    )
        
        return analysis_results
    
    def _combine_results(self, analysis_results: Dict[str, Dict], 
                        duplicate_info: Dict[str, Dict], 
                        extracted_documents: Dict[str, Dict]) -> List[Dict[str, Any]]:
        """Combine all analysis results into final format"""
        final_results = []
        
        for file_path in extracted_documents:
            doc_info = extracted_documents[file_path]
            analysis = analysis_results.get(file_path, {})
            dup_info = duplicate_info.get(file_path, {})
            
            combined_info = {
                'file_name': doc_info['file_name'],
                'file_path': file_path,
                'file_type': doc_info['file_type'],
                'extracted_text': doc_info['extracted_text'],
                **analysis,
                'similarity_percentage': dup_info.get('similarity_percentage', 0.0),
                'duplicate_reason': dup_info.get('duplicate_reason', 'No duplicates found'),
                'master_file': Path(dup_info.get('master_file', file_path)).name
            }
            
            excel_row = self.excel_exporter.prepare_row_data(combined_info)
            final_results.append(excel_row)
        
        return final_results
    
    def _create_empty_result(self, file_path: str, file_type: str) -> Dict[str, Any]:
        """Create empty result for failed processing"""
        return {
            'file_name': Path(file_path).name,
            'file_path': file_path,
            'file_type': file_type,
            'extracted_text': '',
            'summary': 'Failed to process document',
            'description': 'Error occurred during processing',
            'bullet_points': ['Processing failed'],
            'version': 'N/A',
            'tags': ['error'],
            'similarity_percentage': 0.0,
            'duplicate_reason': 'Processing failed',
            'master_file': Path(file_path).name
        }
    
    def _create_empty_analysis(self, file_name: str) -> Dict[str, Any]:
        """Create empty analysis for failed LLM processing"""
        return {
            'summary': f'Failed to analyze {file_name}',
            'description': 'LLM analysis failed',
            'bullet_points': ['Analysis failed'],
            'file_type': 'Unknown',
            'version': 'N/A',
            'tags': ['analysis_failed']
        }
    
    def get_processing_stats(self, results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Generate processing statistics"""
        total_files = len(results)
        duplicates = sum(1 for r in results if r.get('duplicates_percentage', 0) > 0)
        file_types = {}
        
        for result in results:
            file_type = result.get('file_type', 'Unknown')
            file_types[file_type] = file_types.get(file_type, 0) + 1
        
        return {
            'total_files_processed': total_files,
            'duplicates_found': duplicates,
            'unique_files': total_files - duplicates,
            'file_types_distribution': file_types,
            'duplicate_percentage': round((duplicates / total_files) * 100, 2) if total_files > 0 else 0
        }

# ============================================================================
# COMMAND LINE INTERFACE
# ============================================================================

def print_stats(stats: dict):
    """Print processing statistics"""
    print("\n📈 Processing Statistics:")
    print("=" * 50)
    print(f"Total files processed: {stats['total_files_processed']}")
    print(f"Unique files: {stats['unique_files']}")
    print(f"Duplicates found: {stats['duplicates_found']}")
    print(f"Duplicate percentage: {stats['duplicate_percentage']}%")
    print("\nFile types distribution:")
    for file_type, count in stats['file_types_distribution'].items():
        print(f"  • {file_type}: {count}")

def check_environment():
    """Check if environment is properly configured"""
    print("🔧 Environment Check")
    print("=" * 30)
    
    required_env_vars = [
        "AZURE_OPENAI_API_KEY",
        "AZURE_OPENAI_ENDPOINT", 
        "AZURE_OPENAI_DEPLOYMENT"
    ]
    
    missing_vars = []
    for var in required_env_vars:
        if not os.getenv(var) or os.getenv(var) == "your-api-key-here":
            missing_vars.append(var)
    
    if missing_vars:
        print("❌ Missing environment variables:")
        for var in missing_vars:
            print(f"  • {var}")
        print("\n💡 Set these variables:")
        print("export AZURE_OPENAI_API_KEY='your-key'")
        print("export AZURE_OPENAI_ENDPOINT='https://your-resource.openai.azure.com/'")
        print("export AZURE_OPENAI_DEPLOYMENT='gpt-4'")
        return False
    else:
        print("✅ All required environment variables are set")
        return True

def setup_example():
    """Display setup and usage examples"""
    example_text = """
🚀 Document Processing Pipeline - All-in-One Solution
===================================================

📋 SETUP INSTRUCTIONS:

1. Install dependencies:
   pip install python-docx2txt PyPDF2 python-pptx extract-msg openai azure-identity pandas openpyxl python-magic fuzzywuzzy python-Levenshtein nltk scikit-learn numpy tqdm

2. Set environment variables:
   export AZURE_OPENAI_API_KEY="your-api-key"
   export AZURE_OPENAI_ENDPOINT="https://your-resource.openai.azure.com/"
   export AZURE_OPENAI_DEPLOYMENT="gpt-4"

📖 USAGE EXAMPLES:

Process a folder:
   python document_processor_all_in_one.py --folder /path/to/documents

Process single file:
   python document_processor_all_in_one.py --file /path/to/document.pdf

Custom output file:
   python document_processor_all_in_one.py --folder /path/to/documents --output my_results.xlsx

Include statistics:
   python document_processor_all_in_one.py --folder /path/to/documents --stats

📊 SUPPORTED FILE FORMATS:
- PDF (.pdf)
- Word Documents (.docx, .doc)
- PowerPoint (.pptx, .ppt)
- Outlook Messages (.msg)
- Text files (.txt)

📋 EXCEL OUTPUT COLUMNS:
file_name | file_type | extracted_text | summary | description | bullet_points | version | tags | duplicates | duplicates_percentage | master_one

🎯 FEATURES:
✅ Multi-format text extraction
✅ AI-powered summarization with Azure OpenAI
✅ Content-based duplicate detection
✅ Professional Excel export with formatting
✅ Parallel processing for performance
✅ Comprehensive error handling
"""
    print(example_text)

def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description="Process documents with text extraction, LLM analysis, and duplicate detection"
    )
    
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--folder", "-f", type=str, help="Path to folder containing documents")
    group.add_argument("--file", "-sf", type=str, help="Path to single file to process")
    
    parser.add_argument("--output", "-o", type=str, help="Output Excel file path (optional)")
    parser.add_argument("--stats", "-s", action="store_true", help="Display processing statistics")
    
    args = parser.parse_args()
    
    # Setup logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    # Check environment
    if not check_environment():
        print("\n⚠️ Please set up your Azure OpenAI credentials before proceeding.")
        return
    
    try:
        processor = DocumentProcessor()
        
        if args.folder:
            # Process entire folder
            logging.info(f"Processing folder: {args.folder}")
            excel_path = processor.process_folder(args.folder, args.output)
            
            if excel_path:
                print(f"\n✅ Processing completed successfully!")
                print(f"📊 Results exported to: {excel_path}")
                
                if args.stats:
                    df = pd.read_excel(excel_path)
                    stats = processor.get_processing_stats(df.to_dict('records'))
                    print_stats(stats)
            else:
                print("❌ No files were processed")
                
        elif args.file:
            # Process single file
            logging.info(f"Processing single file: {args.file}")
            result = processor.process_single_file(args.file)
            
            excel_path = processor.excel_exporter.export_results([
                processor.excel_exporter.prepare_row_data(result)
            ], args.output)
            
            print(f"\n✅ File processed successfully!")
            print(f"📊 Results exported to: {excel_path}")
            
            if args.stats:
                stats = processor.get_processing_stats([result])
                print_stats(stats)
    
    except KeyboardInterrupt:
        logging.info("Processing interrupted by user")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Error during processing: {str(e)}")
        print(f"❌ Error: {str(e)}")
        sys.exit(1)

# ============================================================================
# PROGRAMMATIC USAGE EXAMPLES
# ============================================================================

def example_programmatic_usage():
    """Example of how to use the processor programmatically"""
    
    # Initialize processor
    processor = DocumentProcessor()
    
    # Example 1: Process a folder
    try:
        excel_output = processor.process_folder(
            folder_path="/path/to/your/documents",
            output_excel_path="my_analysis.xlsx"
        )
        print(f"Folder processing completed: {excel_output}")
    except Exception as e:
        print(f"Error: {e}")
    
    # Example 2: Process a single file
    try:
        result = processor.process_single_file("/path/to/your/document.pdf")
        print(f"Single file processed: {result['file_name']}")
        
        # Export to Excel
        excel_output = processor.excel_exporter.export_results([
            processor.excel_exporter.prepare_row_data(result)
        ])
        print(f"Results exported to: {excel_output}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) == 1:
        setup_example()
    else:
        main()