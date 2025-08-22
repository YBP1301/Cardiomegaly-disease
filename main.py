#!/usr/bin/env python3
"""
Document Processing Pipeline
============================

A comprehensive solution for:
1. Extracting text from various document formats (PDF, Word, PowerPoint, MSG, etc.)
2. LLM-based analysis and summarization using Azure OpenAI
3. Duplicate detection using content similarity
4. Excel export with detailed analysis results

Usage:
    python main.py --folder /path/to/documents
    python main.py --file /path/to/single/document.pdf
    python main.py --folder /path/to/documents --output results.xlsx
"""

import argparse
import logging
import sys
from pathlib import Path
from document_processor import DocumentProcessor

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description="Process documents with text extraction, LLM analysis, and duplicate detection"
    )
    
    # Mutually exclusive group for folder or single file
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--folder", "-f",
        type=str,
        help="Path to folder containing documents to process"
    )
    group.add_argument(
        "--file", "-sf",
        type=str,
        help="Path to single file to process"
    )
    
    parser.add_argument(
        "--output", "-o",
        type=str,
        help="Output Excel file path (optional)"
    )
    
    parser.add_argument(
        "--stats", "-s",
        action="store_true",
        help="Display processing statistics"
    )
    
    args = parser.parse_args()
    
    try:
        # Initialize processor
        processor = DocumentProcessor()
        
        if args.folder:
            # Process entire folder
            logger.info(f"Processing folder: {args.folder}")
            excel_path = processor.process_folder(args.folder, args.output)
            
            if excel_path:
                print(f"\n✅ Processing completed successfully!")
                print(f"📊 Results exported to: {excel_path}")
                
                if args.stats:
                    # Read back the results to show stats
                    import pandas as pd
                    df = pd.read_excel(excel_path)
                    stats = processor.get_processing_stats(df.to_dict('records'))
                    print_stats(stats)
            else:
                print("❌ No files were processed")
                
        elif args.file:
            # Process single file
            logger.info(f"Processing single file: {args.file}")
            result = processor.process_single_file(args.file)
            
            # Export single result to Excel
            excel_path = processor.excel_exporter.export_results([
                processor.excel_exporter.prepare_row_data(result)
            ], args.output)
            
            print(f"\n✅ File processed successfully!")
            print(f"📊 Results exported to: {excel_path}")
            
            if args.stats:
                stats = processor.get_processing_stats([result])
                print_stats(stats)
    
    except KeyboardInterrupt:
        logger.info("Processing interrupted by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error during processing: {str(e)}")
        print(f"❌ Error: {str(e)}")
        sys.exit(1)

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

def setup_example():
    """Create example setup instructions"""
    example_text = """
Example Usage:
=============

1. Set up environment variables:
   export AZURE_OPENAI_API_KEY="your-api-key"
   export AZURE_OPENAI_ENDPOINT="https://your-resource.openai.azure.com/"
   export AZURE_OPENAI_DEPLOYMENT="gpt-4"

2. Install dependencies:
   pip install -r requirements.txt

3. Process a folder of documents:
   python main.py --folder /path/to/documents --output analysis_results.xlsx

4. Process a single file:
   python main.py --file /path/to/document.pdf

5. Include statistics:
   python main.py --folder /path/to/documents --stats

Supported file formats:
- PDF (.pdf)
- Word Documents (.docx, .doc)
- PowerPoint (.pptx, .ppt)
- Outlook Messages (.msg)
- Text files (.txt)

Output Excel columns:
- file_name: Name of the processed file
- file_type: Detected file type
- extracted_text: Text extracted from the document
- summary: LLM-generated comprehensive summary
- description: Brief description with bullet points
- bullet_points: Key points in bullet format
- version: Document version if available
- tags: Relevant keywords/tags
- duplicates: Reason for duplication if found
- duplicates_percentage: Similarity percentage with duplicates
- master_one: Master file among duplicates
"""
    print(example_text)

if __name__ == "__main__":
    if len(sys.argv) == 1:
        setup_example()
    else:
        main()