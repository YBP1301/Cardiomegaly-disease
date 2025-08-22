import json
import logging
from typing import Dict, Any, Optional
from openai import AzureOpenAI
from config import AZURE_OPENAI_CONFIG, SYSTEM_PROMPTS, MAX_TEXT_LENGTH

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
        """
        Analyze document text and extract summary, metadata, and tags
        
        Args:
            text: Extracted text from document
            file_name: Name of the file being analyzed
            
        Returns:
            Dictionary containing analysis results
        """
        try:
            # Truncate text if too long
            if len(text) > MAX_TEXT_LENGTH:
                text = text[:MAX_TEXT_LENGTH] + "... [truncated]"
                logger.warning(f"Text truncated for {file_name} due to length")
            
            # Prepare the prompt
            user_prompt = f"""
            Please analyze the following document text from file: {file_name}
            
            Document content:
            {text}
            
            Provide a comprehensive analysis including summary, description, bullet points, file type, version, and relevant tags.
            """
            
            # Call Azure OpenAI
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPTS["summarization"]},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
                max_tokens=1500
            )
            
            # Parse the response
            response_text = response.choices[0].message.content
            
            try:
                analysis_result = json.loads(response_text)
                return self._validate_analysis_result(analysis_result)
            except json.JSONDecodeError:
                logger.error(f"Failed to parse JSON response for {file_name}")
                return self._create_fallback_result(text, file_name)
                
        except Exception as e:
            logger.error(f"Error analyzing document {file_name}: {str(e)}")
            return self._create_fallback_result(text, file_name)
    
    def compare_documents(self, text1: str, text2: str, file1: str, file2: str) -> Dict[str, Any]:
        """
        Compare two documents for similarity and duplicate detection
        
        Args:
            text1: Text from first document
            text2: Text from second document
            file1: Name of first file
            file2: Name of second file
            
        Returns:
            Dictionary containing comparison results
        """
        try:
            # Truncate texts if too long
            if len(text1) > MAX_TEXT_LENGTH // 2:
                text1 = text1[:MAX_TEXT_LENGTH // 2] + "... [truncated]"
            if len(text2) > MAX_TEXT_LENGTH // 2:
                text2 = text2[:MAX_TEXT_LENGTH // 2] + "... [truncated]"
            
            user_prompt = f"""
            Compare these two documents for similarity:
            
            Document 1 ({file1}):
            {text1}
            
            Document 2 ({file2}):
            {text2}
            
            Determine if they are duplicates and provide detailed analysis.
            """
            
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPTS["duplicate_analysis"]},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.1,
                max_tokens=800
            )
            
            response_text = response.choices[0].message.content
            
            try:
                comparison_result = json.loads(response_text)
                return self._validate_comparison_result(comparison_result, file1, file2)
            except json.JSONDecodeError:
                logger.error(f"Failed to parse JSON response for comparison {file1} vs {file2}")
                return self._create_fallback_comparison(file1, file2)
                
        except Exception as e:
            logger.error(f"Error comparing documents {file1} vs {file2}: {str(e)}")
            return self._create_fallback_comparison(file1, file2)
    
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
        
        # Ensure bullet_points is a list
        if isinstance(validated["bullet_points"], str):
            validated["bullet_points"] = [validated["bullet_points"]]
        
        # Ensure tags is a list
        if isinstance(validated["tags"], str):
            validated["tags"] = [validated["tags"]]
            
        return validated
    
    def _validate_comparison_result(self, result: Dict[str, Any], file1: str, file2: str) -> Dict[str, Any]:
        """Validate and clean comparison result"""
        return {
            "is_duplicate": result.get("is_duplicate", False),
            "similarity_percentage": float(result.get("similarity_percentage", 0.0)),
            "duplicate_reason": result.get("duplicate_reason", "No analysis available"),
            "master_file": result.get("master_file", file1)
        }
    
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
    
    def _create_fallback_comparison(self, file1: str, file2: str) -> Dict[str, Any]:
        """Create fallback comparison when LLM analysis fails"""
        return {
            "is_duplicate": False,
            "similarity_percentage": 0.0,
            "duplicate_reason": "Analysis failed",
            "master_file": file1
        }