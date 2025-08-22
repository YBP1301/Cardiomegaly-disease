#!/usr/bin/env python3
"""
Installation and Test Script for Document Processing Pipeline
============================================================

This script will:
1. Install all required dependencies
2. Create sample test files
3. Run a quick test to verify everything works
"""

import os
import sys
import subprocess
from pathlib import Path

def install_dependencies():
    """Install all required packages"""
    print("📦 Installing dependencies...")
    
    packages = [
        "python-docx2txt==0.8",
        "PyPDF2==3.0.1", 
        "python-pptx==0.6.21",
        "extract-msg==0.45.0",
        "openai==1.12.0",
        "azure-identity==1.15.0",
        "pandas==2.2.0",
        "openpyxl==3.1.2",
        "python-magic==0.4.27",
        "fuzzywuzzy==0.18.0",
        "python-Levenshtein==0.20.9",
        "nltk==3.8.1",
        "scikit-learn==1.4.0",
        "numpy==1.26.3",
        "tqdm==4.66.1"
    ]
    
    try:
        for package in packages:
            print(f"Installing {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package], 
                                capture_output=True, text=True)
        print("✅ All dependencies installed successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install dependencies: {e}")
        return False

def create_test_files():
    """Create sample test files for testing"""
    print("📁 Creating test files...")
    
    test_folder = Path("test_documents")
    test_folder.mkdir(exist_ok=True)
    
    # Create sample documents with different content and duplicates
    test_files = {
        "document1.txt": """
        Project Proposal: AI-Powered Document Management System
        
        Executive Summary:
        This document outlines a comprehensive proposal for implementing an AI-powered document management system that will revolutionize how our organization handles, processes, and analyzes documents.
        
        Key Features:
        - Automated text extraction from multiple file formats
        - Intelligent document categorization and tagging
        - Duplicate detection and content analysis
        - Advanced search capabilities with semantic understanding
        
        Implementation Timeline:
        Phase 1: Infrastructure setup (2 months)
        Phase 2: Core functionality development (4 months)
        Phase 3: Testing and deployment (2 months)
        
        Budget Estimate: $150,000
        """,
        
        "document2.txt": """
        Marketing Strategy Report Q4 2024
        
        Overview:
        This report presents our comprehensive marketing strategy for the fourth quarter of 2024, focusing on digital transformation and customer engagement initiatives.
        
        Key Objectives:
        - Increase brand awareness by 25%
        - Boost online engagement by 40%
        - Generate 500 new qualified leads
        - Improve customer retention rate to 85%
        
        Tactics:
        1. Social media advertising campaigns
        2. Content marketing and SEO optimization
        3. Email marketing automation
        4. Influencer partnerships
        5. Trade show participation
        
        Budget Allocation: $75,000
        """,
        
        "duplicate_document1.txt": """
        Project Proposal: AI-Powered Document Management System
        
        Executive Summary:
        This document outlines a comprehensive proposal for implementing an AI-powered document management system that will revolutionize how our organization handles, processes, and analyzes documents.
        
        Key Features:
        - Automated text extraction from multiple file formats
        - Intelligent document categorization and tagging  
        - Duplicate detection and content analysis
        - Advanced search capabilities with semantic understanding
        
        Implementation Timeline:
        Phase 1: Infrastructure setup (2 months)
        Phase 2: Core functionality development (4 months)
        Phase 3: Testing and deployment (2 months)
        
        Budget Estimate: $150,000
        """,
        
        "similar_document.txt": """
        AI Document Management Proposal
        
        Summary:
        We propose implementing an artificial intelligence-powered system for managing documents that will transform our organization's document handling and analysis processes.
        
        Main Features:
        - Automatic text extraction from various file types
        - Smart document classification and keyword tagging
        - Detection of duplicate content and analysis
        - Enhanced search with semantic capabilities
        
        Project Timeline:
        Stage 1: Setup infrastructure (2 months)
        Stage 2: Build core features (4 months) 
        Stage 3: Test and deploy (2 months)
        
        Cost Estimate: $150,000
        """
    }
    
    for filename, content in test_files.items():
        file_path = test_folder / filename
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content.strip())
    
    print(f"✅ Test files created in: {test_folder.absolute()}")
    return test_folder

def run_test():
    """Run a quick test of the document processor"""
    print("🧪 Running test...")
    
    try:
        # Import the processor
        sys.path.append(str(Path.cwd()))
        from document_processor_all_in_one import DocumentProcessor
        
        # Test with sample files
        test_folder = Path("test_documents")
        if test_folder.exists():
            processor = DocumentProcessor()
            
            # Test single file processing
            test_file = test_folder / "document1.txt"
            if test_file.exists():
                result = processor.process_single_file(str(test_file))
                print(f"✅ Single file test passed: {result['file_name']}")
            
            print("✅ All tests passed!")
            print(f"💡 Run full test with: python document_processor_all_in_one.py --folder {test_folder}")
            return True
        else:
            print("❌ Test folder not found")
            return False
            
    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False

def main():
    """Main installation and test function"""
    print("🚀 Document Processing Pipeline - Installation & Test")
    print("=" * 60)
    
    print("Current working directory:", Path.cwd())
    
    # Install dependencies
    if not install_dependencies():
        print("❌ Installation failed. Please check the error messages above.")
        return
    
    # Create test files
    test_folder = create_test_files()
    
    # Run basic test
    if run_test():
        print("\n🎉 Installation and test completed successfully!")
        print("\n📚 Next steps:")
        print("1. Set your Azure OpenAI environment variables:")
        print("   export AZURE_OPENAI_API_KEY='your-api-key'")
        print("   export AZURE_OPENAI_ENDPOINT='https://your-resource.openai.azure.com/'")
        print("   export AZURE_OPENAI_DEPLOYMENT='gpt-4'")
        print(f"\n2. Test with sample files:")
        print(f"   python document_processor_all_in_one.py --folder {test_folder} --stats")
        print(f"\n3. Process your own documents:")
        print(f"   python document_processor_all_in_one.py --folder /path/to/your/documents")
    else:
        print("⚠️ Installation completed but test failed. Please check error messages.")

if __name__ == "__main__":
    main()