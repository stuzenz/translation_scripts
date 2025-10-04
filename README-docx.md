# DOCX Translator - Usage Guide

A command-line tool for translating Microsoft Word documents while preserving formatting, styles, tables, and document structure. Features automatic Table of Contents (TOC) handling and enhanced translation quality through context-aware processing.

## Quick Start

```bash
# Translate a single file to Japanese
python docx_translator8.py document.docx --target-lang ja

# Translate all DOCX files in a directory to multiple languages
python docx_translator8.py --input-files ./documents/ --output-files ./translated/ --target-langs ja,es,fr

# Use business style with smart context
python docx_translator8.py document.docx --target-lang ja --style-prompt business --smart-context
```

## Command Line Options

### Input Options

| Option | Description | Example |
|--------|-------------|---------|
| `input_file` | Single DOCX file to translate (positional) | `document.docx` |
| `--input-files` | Directory containing DOCX files to process | `--input-files ./documents/` |

**Note**: Use either `input_file` OR `--input-files`, not both.

### Language Options

| Option | Description | Example |
|--------|-------------|---------|
| `--source-lang` | Source language code (default: auto-detect) | `--source-lang en` |
| `--target-lang` | Single target language code | `--target-lang ja` |
| `--target-langs` | Multiple target languages (comma-separated) | `--target-langs ja,es,fr,de` |

**Note**: Use either `--target-lang` OR `--target-langs`, not both.

#### Common Language Codes
- `en` - English
- `ja` - Japanese
- `es` - Spanish
- `fr` - French
- `de` - German
- `it` - Italian
- `pt` - Portuguese
- `zh` - Chinese (Simplified)
- `ko` - Korean

### Translation Quality Options

| Option | Description | Default | Example |
|--------|-------------|---------|---------|
| `--style-prompt` | Translation style | `business` | `--style-prompt technical` |
| `--smart-context` | Use enhanced context extraction | `false` | `--smart-context` |
| `--context-file` | Path to glossary JSON file | None | `--context-file terms.json` |

#### Available Style Prompts

- **`business`** (default) - Formal business language with appropriate terminology
- **`casual`** - Natural, friendly language for general communication
- **`technical`** - Precise technical terminology for specialized content
- **`academic`** - Scholarly language for academic contexts
- **`marketing`** - Persuasive, engaging language for marketing materials

### Output Options

| Option | Description | Default | Example |
|--------|-------------|---------|---------|
| `--output-dir` | Output directory | `.` (current) | `--output-dir ./translated/` |
| `--output-files` | Same as `--output-dir` (for compatibility) | `.` (current) | `--output-files ./translated/` |

### Performance Options

| Option | Description | Default | Example |
|--------|-------------|---------|---------|
| `--model` | Gemini model to use | `gemini-flash-latest` | `--model gemini-2.0-flash` |
| `--batch-size` | Texts per API call | `10` | `--batch-size 20` |
| `--concurrency` | Concurrent API calls | `4` | `--concurrency 8` |

### Information Options

| Option | Description |
|--------|-------------|
| `--debug` | Enable debug logging |
| `--list-styles` | Show available style prompts |
| `--version` | Show version information |

## Glossary File Format

The `--context-file` option accepts a JSON file containing terminology mappings for consistent translation.

### Structure

```json
{
    "source_term": "target_translation",
    "another_term": "another_translation"
}
```

### Example Glossary (`business-terms.json`)

```json
{
    "Revenue": "収益",
    "Operating Expenses": "営業費用",
    "Net Profit": "純利益",
    "Q1": "第1四半期",
    "Q2": "第2四半期",
    "Q3": "第3四半期",
    "Q4": "第4四半期",
    "FY2024": "2024年度",
    "Dashboard": "ダッシュボード",
    "API": "API",
    "SDK": "SDK",
    "User Interface": "ユーザーインターフェース",
    "Database": "データベース"
}
```

### Technical Terms Example (`tech-glossary.json`)

```json
{
    "Authentication": "認証",
    "Authorization": "承認",
    "Load Balancer": "ロードバランサー",
    "Microservices": "マイクロサービス",
    "Container": "コンテナ",
    "Kubernetes": "Kubernetes",
    "Docker": "Docker",
    "REST API": "REST API",
    "GraphQL": "GraphQL",
    "NoSQL": "NoSQL"
}
```

### Usage with Glossary

```bash
# Use business terminology glossary
python docx_translator8.py report.docx --target-lang ja --context-file business-terms.json

# Combine with smart context for better results
python docx_translator8.py technical-doc.docx --target-lang ja --context-file tech-glossary.json --smart-context
```

## Smart Context Feature

The `--smart-context` option enhances translation quality by:

- **Document Structure Analysis**: Extracts document title and section headings
- **Content Sampling**: Analyzes first few paragraphs for context
- **Table Context**: Understands table headers and relationships
- **Cross-Reference Awareness**: Maintains consistency across document sections

**When to use Smart Context:**
- Technical documentation
- Business reports with specific terminology
- Documents with tables and structured content
- Long documents with multiple sections

## Table of Contents (TOC) Handling

The translator automatically handles Table of Contents:

1. **Detection**: Automatically finds TOC fields in documents
2. **Marking**: Marks TOC fields for update after translation
3. **User Experience**: Word prompts to update TOC when opening translated document
4. **One-Time Update**: Click "Yes" once, then normal document behavior

### What You'll See

During translation:
```
📑 Marked 1 TOC field(s) for update
💡 Word will prompt to update the Table of Contents when document is opened
```

When opening translated document in Word:
- Prompt: *"This document contains fields that may refer to other files. Do you want to update the fields in this document?"*
- Click **"Yes"** to update TOC with translated headings

## Examples

### Single File Translation

```bash
# Basic translation to Japanese
python docx_translator8.py document.docx --target-lang ja

# Business style with smart context
python docx_translator8.py report.docx --target-lang ja --style-prompt business --smart-context

# Technical document with glossary
python docx_translator8.py manual.docx --target-lang ja --style-prompt technical --context-file tech-terms.json
```

### Batch Processing

```bash
# Translate all documents in a directory
python docx_translator8.py --input-files ./documents/ --output-files ./translated/ --target-lang ja

# Multiple languages with custom output directory
python docx_translator8.py --input-files ./reports/ --output-files ./international/ --target-langs ja,es,fr,de

# High-quality batch processing
python docx_translator8.py --input-files ./docs/ --output-files ./translated/ --target-langs ja,ko --smart-context --style-prompt business
```

### Performance Optimization

```bash
# Faster processing (more concurrent calls)
python docx_translator8.py large-doc.docx --target-lang ja --concurrency 8 --batch-size 20

# Conservative processing (for rate limits)
python docx_translator8.py document.docx --target-lang ja --concurrency 2 --batch-size 5
```

## Output Structure

### Single File Mode
```
original_document.docx → original_document_ja.docx
```

### Directory Mode
```
output_directory/
├── document1_ja.docx
├── document1_es.docx
├── document2_ja.docx
└── document2_es.docx
```

## Requirements

- Python 3.8+
- Google Gemini API key (set as `GOOGLE_API_KEY` environment variable)
- Required packages: `google-generativeai`, `python-docx`, `colorama`, `tqdm`

### Setup

```bash
# Set API key
export GOOGLE_API_KEY="your-api-key-here"

# Install dependencies
pip install -r requirements.txt
```

## Best Practices

1. **Use Smart Context** for technical or structured documents
2. **Create Glossaries** for consistent terminology in specialized domains
3. **Choose Appropriate Style** based on document type and audience
4. **Test with Single Files** before batch processing large directories
5. **Use Conservative Settings** if experiencing rate limits
6. **Always Update TOC** when prompted in Word for translated documents

## Troubleshooting

### Common Issues

**API Key Error**
```
Error: GOOGLE_API_KEY environment variable not set
```
Solution: Set your Google API key as an environment variable.

**Rate Limiting**
```
Translation attempt failed: Resource exhausted
```
Solution: Reduce `--concurrency` or `--batch-size` values.

**No Files Found**
```
No DOCX files found in 'directory'
```
Solution: Check directory path and ensure it contains .docx files.

### Debug Mode

Enable detailed logging for troubleshooting:
```bash
python docx_translator8.py document.docx --target-lang ja --debug
```

This provides detailed information about:
- File processing steps
- API request/response details
- Translation mapping process
- Error diagnostics