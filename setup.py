#!/usr/bin/env python3
"""
Setup script for Document Processing Pipeline
===========================================

This script helps set up the environment and install dependencies
"""

import os
import sys
import subprocess
from pathlib import Path

def install_dependencies():
    """Install required Python packages"""
    print("📦 Installing Python dependencies...")
    
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Dependencies installed successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to install dependencies: {e}")
        return False

def check_python_version():
    """Check if Python version is compatible"""
    print("🐍 Checking Python version...")
    
    version = sys.version_info
    if version.major >= 3 and version.minor >= 8:
        print(f"✅ Python {version.major}.{version.minor}.{version.micro} is compatible")
        return True
    else:
        print(f"❌ Python {version.major}.{version.minor}.{version.micro} is not compatible")
        print("💡 Please upgrade to Python 3.8 or higher")
        return False

def setup_environment_variables():
    """Guide user through setting up environment variables"""
    print("\n🔑 Setting up Azure OpenAI environment variables...")
    
    env_vars = {
        "AZURE_OPENAI_API_KEY": "Your Azure OpenAI API key",
        "AZURE_OPENAI_ENDPOINT": "Your Azure OpenAI endpoint (e.g., https://your-resource.openai.azure.com/)",
        "AZURE_OPENAI_DEPLOYMENT": "Your deployment name (e.g., gpt-4)"
    }
    
    print("Please set the following environment variables:")
    print("You can add these to your ~/.bashrc or ~/.zshrc file for persistence")
    print()
    
    for var, description in env_vars.items():
        current_value = os.getenv(var, "Not set")
        print(f"export {var}=\"your-value-here\"  # {description}")
        print(f"  Current value: {current_value}")
        print()

def create_sample_folder():
    """Create a sample folder structure for testing"""
    print("📁 Creating sample folder structure...")
    
    sample_folder = Path("sample_documents")
    sample_folder.mkdir(exist_ok=True)
    
    # Create sample text files for testing
    sample_files = {
        "sample1.txt": "This is a sample document for testing the document processing pipeline. It contains some basic text content.",
        "sample2.txt": "This is another sample document with different content. It talks about various topics and concepts.",
        "duplicate_sample.txt": "This is a sample document for testing the document processing pipeline. It contains some basic text content."  # Duplicate of sample1
    }
    
    for filename, content in sample_files.items():
        file_path = sample_folder / filename
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
    
    print(f"✅ Sample documents created in: {sample_folder.absolute()}")
    print("💡 You can test the pipeline with: python main.py --folder sample_documents")

def run_test():
    """Run a quick test to verify everything is working"""
    print("\n🧪 Running quick test...")
    
    try:
        # Test imports
        from document_processor import DocumentProcessor
        from text_extractor import TextExtractor
        from llm_analyzer import LLMAnalyzer
        from duplicate_detector import DuplicateDetector
        from excel_exporter import ExcelExporter
        
        print("✅ All modules imported successfully")
        
        # Test text extractor with sample file
        if Path("sample_documents/sample1.txt").exists():
            extractor = TextExtractor()
            text, file_type = extractor.extract_text("sample_documents/sample1.txt")
            if text:
                print("✅ Text extraction working")
            else:
                print("⚠️ Text extraction returned empty result")
        
        print("✅ Basic functionality test passed")
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False
    
    return True

def main():
    """Main setup function"""
    print("🚀 Document Processing Pipeline Setup")
    print("=" * 50)
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Install dependencies
    if not install_dependencies():
        sys.exit(1)
    
    # Setup environment variables
    setup_environment_variables()
    
    # Create sample folder
    create_sample_folder()
    
    # Run test
    if run_test():
        print("\n🎉 Setup completed successfully!")
        print("\n📚 Next steps:")
        print("1. Set your Azure OpenAI environment variables")
        print("2. Test with sample documents: python main.py --folder sample_documents")
        print("3. Process your own documents: python main.py --folder /path/to/your/documents")
        print("4. For help: python main.py --help")
    else:
        print("\n⚠️ Setup completed with warnings. Please check the error messages above.")

if __name__ == "__main__":
    main()