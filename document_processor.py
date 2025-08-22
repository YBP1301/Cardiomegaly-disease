import os
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional
from tqdm import tqdm
import concurrent.futures
from threading import Lock

from text_extractor import TextExtractor
from llm_analyzer import LLMAnalyzer
from duplicate_detector import DuplicateDetector
from excel_exporter import ExcelExporter
from config import SUPPORTED_EXTENSIONS

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class DocumentProcessor:
    """Main orchestrator for document processing pipeline"""
    
    def __init__(self):
        self.text_extractor = TextExtractor()
        self.llm_analyzer = LLMAnalyzer()
        self.duplicate_detector = DuplicateDetector()
        self.excel_exporter = ExcelExporter()
        self.results_lock = Lock()
    
    def process_folder(self, folder_path: str, output_excel_path: Optional[str] = None) -> str:
        """
        Process all supported documents in a folder
        
        Args:
            folder_path: Path to the folder containing documents
            output_excel_path: Optional custom path for Excel output
            
        Returns:
            Path to the generated Excel file
        """
        try:
            folder_path = Path(folder_path)
            if not folder_path.exists():
                raise FileNotFoundError(f"Folder not found: {folder_path}")
            
            logger.info(f"Starting document processing for folder: {folder_path}")
            
            # Step 1: Find all supported files
            supported_files = self._find_supported_files(folder_path)
            logger.info(f"Found {len(supported_files)} supported files")
            
            if not supported_files:
                logger.warning("No supported files found in the folder")
                return ""
            
            # Step 2: Extract text from all files
            logger.info("Extracting text from documents...")
            extracted_documents = self._extract_texts_parallel(supported_files)
            
            # Step 3: Analyze documents with LLM
            logger.info("Analyzing documents with Azure OpenAI...")
            analysis_results = self._analyze_documents_parallel(extracted_documents)
            
            # Step 4: Detect duplicates
            logger.info("Detecting duplicates...")
            duplicate_info = self.duplicate_detector.find_duplicates(
                {path: info['extracted_text'] for path, info in extracted_documents.items()}
            )
            
            # Step 5: Combine all results
            logger.info("Combining results...")
            final_results = self._combine_results(analysis_results, duplicate_info, extracted_documents)
            
            # Step 6: Export to Excel
            logger.info("Exporting to Excel...")
            excel_path = self.excel_exporter.export_results(final_results, output_excel_path)
            
            logger.info(f"Document processing completed. Results saved to: {excel_path}")
            return excel_path
            
        except Exception as e:
            logger.error(f"Error processing folder: {str(e)}")
            raise
    
    def process_single_file(self, file_path: str) -> Dict[str, Any]:
        """
        Process a single document file
        
        Args:
            file_path: Path to the document file
            
        Returns:
            Dictionary containing analysis results
        """
        try:
            file_path = Path(file_path)
            
            # Extract text
            extracted_text, file_type = self.text_extractor.extract_text(str(file_path))
            
            if not extracted_text:
                logger.warning(f"No text extracted from {file_path}")
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
            logger.error(f"Error processing single file {file_path}: {str(e)}")
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
        
        # Use ThreadPoolExecutor for I/O bound operations
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            future_to_file = {executor.submit(extract_single, fp): fp for fp in file_paths}
            
            for future in tqdm(concurrent.futures.as_completed(future_to_file), 
                             total=len(file_paths), desc="Extracting text"):
                try:
                    file_path, result = future.result()
                    extracted_documents[file_path] = result
                except Exception as e:
                    file_path = future_to_file[future]
                    logger.error(f"Error extracting text from {file_path}: {str(e)}")
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
        
        # Use ThreadPoolExecutor for API calls
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
                    logger.error(f"Error analyzing {file_path}: {str(e)}")
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
            
            # Combine all information
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
            
            # Prepare for Excel export
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