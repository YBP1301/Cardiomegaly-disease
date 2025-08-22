#!/usr/bin/env python3
"""
Example Usage Script for Document Processing Pipeline
====================================================

This script demonstrates how to use the DocumentProcessor class programmatically
"""

import os
from pathlib import Path
from document_processor import DocumentProcessor

def example_folder_processing():
    """Example: Process all documents in a folder"""
    print("🚀 Example: Processing folder of documents")
    print("=" * 50)
    
    # Initialize the processor
    processor = DocumentProcessor()
    
    # Example folder path (replace with your actual path)
    folder_path = "/path/to/your/documents"
    
    # Check if folder exists (for demo purposes, we'll create a sample structure)
    if not Path(folder_path).exists():
        print(f"📁 Folder {folder_path} not found. Please update the path in the script.")
        return
    
    try:
        # Process the folder
        excel_output = processor.process_folder(
            folder_path=folder_path,
            output_excel_path="my_document_analysis.xlsx"
        )
        
        print(f"✅ Processing completed!")
        print(f"📊 Results saved to: {excel_output}")
        
        # Get processing statistics
        import pandas as pd
        df = pd.read_excel(excel_output)
        stats = processor.get_processing_stats(df.to_dict('records'))
        
        print(f"\n📈 Statistics:")
        print(f"  • Total files: {stats['total_files_processed']}")
        print(f"  • Duplicates: {stats['duplicates_found']}")
        print(f"  • File types: {list(stats['file_types_distribution'].keys())}")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")

def example_single_file_processing():
    """Example: Process a single document"""
    print("\n🚀 Example: Processing single document")
    print("=" * 50)
    
    # Initialize the processor
    processor = DocumentProcessor()
    
    # Example file path (replace with your actual file)
    file_path = "/path/to/your/document.pdf"
    
    if not Path(file_path).exists():
        print(f"📄 File {file_path} not found. Please update the path in the script.")
        return
    
    try:
        # Process the single file
        result = processor.process_single_file(file_path)
        
        print(f"✅ File processed successfully!")
        print(f"📄 File: {result['file_name']}")
        print(f"📝 Summary: {result['summary'][:100]}...")
        print(f"🏷️ Tags: {', '.join(result['tags'])}")
        
        # Export to Excel
        excel_output = processor.excel_exporter.export_results([
            processor.excel_exporter.prepare_row_data(result)
        ], "single_file_analysis.xlsx")
        
        print(f"📊 Results saved to: {excel_output}")
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")

def example_custom_configuration():
    """Example: Using custom configuration"""
    print("\n🚀 Example: Custom configuration")
    print("=" * 50)
    
    # You can modify config.py or override settings programmatically
    from config import AZURE_OPENAI_CONFIG, DUPLICATE_THRESHOLD
    
    print(f"Current Azure OpenAI endpoint: {AZURE_OPENAI_CONFIG['azure_endpoint']}")
    print(f"Current duplicate threshold: {DUPLICATE_THRESHOLD}")
    print("💡 Modify config.py to change these settings")

def example_batch_processing():
    """Example: Process multiple specific files"""
    print("\n🚀 Example: Batch processing specific files")
    print("=" * 50)
    
    processor = DocumentProcessor()
    
    # List of specific files to process
    files_to_process = [
        "/path/to/document1.pdf",
        "/path/to/document2.docx",
        "/path/to/presentation.pptx"
    ]
    
    results = []
    
    for file_path in files_to_process:
        if Path(file_path).exists():
            try:
                result = processor.process_single_file(file_path)
                results.append(processor.excel_exporter.prepare_row_data(result))
                print(f"✅ Processed: {Path(file_path).name}")
            except Exception as e:
                print(f"❌ Failed to process {file_path}: {str(e)}")
        else:
            print(f"📄 File not found: {file_path}")
    
    if results:
        # Export all results to Excel
        excel_output = processor.excel_exporter.export_results(
            results, 
            "batch_processing_results.xlsx"
        )
        print(f"📊 Batch results saved to: {excel_output}")

def check_environment():
    """Check if environment is properly configured"""
    print("🔧 Environment Check")
    print("=" * 50)
    
    # Check Azure OpenAI configuration
    required_env_vars = [
        "AZURE_OPENAI_API_KEY",
        "AZURE_OPENAI_ENDPOINT", 
        "AZURE_OPENAI_DEPLOYMENT"
    ]
    
    missing_vars = []
    for var in required_env_vars:
        if not os.getenv(var):
            missing_vars.append(var)
    
    if missing_vars:
        print("❌ Missing environment variables:")
        for var in missing_vars:
            print(f"  • {var}")
        print("\n💡 Set these variables before running the processor:")
        print("export AZURE_OPENAI_API_KEY='your-key'")
        print("export AZURE_OPENAI_ENDPOINT='https://your-resource.openai.azure.com/'")
        print("export AZURE_OPENAI_DEPLOYMENT='gpt-4'")
    else:
        print("✅ All required environment variables are set")
    
    # Check if dependencies are installed
    try:
        import PyPDF2, docx2txt, extract_msg, openai, pandas, openpyxl
        print("✅ All required packages are installed")
    except ImportError as e:
        print(f"❌ Missing package: {e}")
        print("💡 Run: pip install -r requirements.txt")

if __name__ == "__main__":
    print("🔧 Document Processing Pipeline - Example Usage")
    print("=" * 60)
    
    # Check environment first
    check_environment()
    
    # Run examples (comment out the ones you don't want to run)
    # example_folder_processing()
    # example_single_file_processing()
    # example_custom_configuration()
    # example_batch_processing()
    
    print("\n💡 Uncomment the example functions you want to run and update file paths!")
    print("💡 Or use the command line interface: python main.py --help")