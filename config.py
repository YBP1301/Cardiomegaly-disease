import os
from typing import Dict, Any

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
MAX_TEXT_LENGTH = 50000  # Maximum text length to send to LLM
SUPPORTED_EXTENSIONS = {'.pdf', '.docx', '.doc', '.pptx', '.ppt', '.msg', '.txt'}
DUPLICATE_THRESHOLD = 0.85  # Similarity threshold for duplicate detection

# Excel output settings
EXCEL_COLUMNS = [
    'file_name', 'file_type', 'extracted_text', 'summary', 
    'description', 'bullet_points', 'version', 'tags', 
    'duplicates', 'duplicates_percentage', 'master_one'
]