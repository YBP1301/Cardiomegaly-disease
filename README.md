# Document Processing Pipeline

A comprehensive Python solution for document analysis, text extraction, LLM-based summarization, and duplicate detection with Excel export.

## Features

### 📄 Multi-Format Text Extraction
- **PDF** documents (.pdf)
- **Word** documents (.docx, .doc)
- **PowerPoint** presentations (.pptx, .ppt)
- **Outlook** messages (.msg)
- **Text** files (.txt)

### 🤖 AI-Powered Analysis
- Comprehensive summarization using Azure OpenAI
- Brief descriptions with bullet points
- Automatic metadata extraction (file type, version, tags)
- Content-based keyword generation

### 🔍 Advanced Duplicate Detection
- Multi-algorithm similarity detection
- Content hash comparison
- Fuzzy string matching
- TF-IDF semantic similarity
- Master file identification

### 📊 Excel Export
- Formatted Excel output with all analysis results
- Conditional formatting for duplicates
- Professional styling and layout

## Installation

1. **Clone or download the project files**

2. **Install dependencies:**
```bash
pip install -r requirements.txt
```

3. **Set up Azure OpenAI credentials:**
```bash
export AZURE_OPENAI_API_KEY="your-api-key-here"
export AZURE_OPENAI_ENDPOINT="https://your-resource.openai.azure.com/"
export AZURE_OPENAI_DEPLOYMENT="gpt-4"
```

## Usage

### Process a folder of documents:
```bash
python main.py --folder /path/to/documents
```

### Process a single file:
```bash
python main.py --file /path/to/document.pdf
```

### Specify output file:
```bash
python main.py --folder /path/to/documents --output my_analysis.xlsx
```

### Include processing statistics:
```bash
python main.py --folder /path/to/documents --stats
```

## Excel Output Columns

| Column | Description |
|--------|-------------|
| `file_name` | Name of the processed file |
| `file_type` | Detected file type |
| `extracted_text` | Text extracted from document |
| `summary` | LLM-generated comprehensive summary |
| `description` | Brief description |
| `bullet_points` | Key points in bullet format |
| `version` | Document version if available |
| `tags` | Relevant keywords/tags |
| `duplicates` | Reason for duplication if found |
| `duplicates_percentage` | Similarity percentage |
| `master_one` | Master file among duplicates |

## Configuration

Edit `config.py` to customize:
- Azure OpenAI settings
- System prompts for LLM analysis
- Duplicate detection thresholds
- File processing limits

## Architecture

- **`text_extractor.py`**: Handles multi-format text extraction
- **`llm_analyzer.py`**: Azure OpenAI integration for analysis
- **`duplicate_detector.py`**: Advanced duplicate detection algorithms
- **`excel_exporter.py`**: Professional Excel export with formatting
- **`document_processor.py`**: Main orchestration pipeline
- **`main.py`**: Command-line interface

## Error Handling

The system includes comprehensive error handling:
- Graceful fallbacks for unsupported files
- Retry mechanisms for API failures
- Detailed logging for troubleshooting
- Partial processing continuation on individual file failures

## Performance

- Parallel text extraction for faster processing
- Concurrent LLM API calls (rate-limited)
- Efficient duplicate detection algorithms
- Memory-optimized for large document sets