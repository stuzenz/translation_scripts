# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a machine translation toolkit that uses Large Language Models (LLMs) to translate various document formats while preserving formatting and structure. The project has evolved from supporting local models to focusing on commercial API providers that don't use data for training.

## Core Architecture

### Translation Scripts by Format

The repository contains specialized translators for different document formats:

- **DOCX**: `docx_translator8.py` - Word documents with formatting preservation
- **PPTX**: `pptx-translator-api4.py` - PowerPoint presentations
- **XLSX**: `excel_translator_21.py` - Excel spreadsheets with formula preservation
- **VTT**: `vtt-translator9.py` - Video subtitle files with language detection
- **Excalidraw**: `excalidraw_translate2.py` - Diagram files

### Google AI SDK Migration

The codebase is transitioning from the deprecated `google-generativeai` package to the new `google-genai` SDK (v1.27+):

**Legacy import pattern** (being phased out):
```python
import google.generativeai as genai
genai.configure(api_key=api_key)
```

**New SDK pattern** (preferred):
```python
from google import genai
from google.genai import types
client = genai.Client(api_key=api_key)
```

### Common Translation Features

All translators share these core capabilities:
- Batch processing for API efficiency
- Context-aware translation using glossaries
- Style-based translation (business, technical, casual, etc.)
- Smart context extraction from surrounding content
- Concurrent API calls for performance
- Robust JSON parsing for LLM responses

## Development Environment

### Setup Commands

```bash
# Enter development environment (uses devenv.sh)
devenv shell

# Install Python dependencies
pip install -r requirements.txt

# Run tests
python src/test.py
```

### Dependencies

Primary dependencies managed in `requirements.txt`:
- `google-genai>=1.27.0` (new SDK, preferred)
- `google-generativeai` (legacy, deprecated Aug 2025)
- Document processing: `python-docx`, `python-pptx`, `openpyxl`, `vsdx`
- Utilities: `langdetect`, `colorama`, `tqdm`, `pandas`

## Common Development Tasks

### Running Translations

Standard command pattern across all translators:
```bash
python script_name.py input_file --target-lang LANG_CODE [options]
```

Common flags:
- `--target-lang ja` / `--target-langs en,ja,es` - Target language(s)
- `--source-lang en` - Source language (auto-detected if not specified)
- `--batch-size 20` - API batch size for performance tuning
- `--concurrency 4` - Concurrent API calls
- `--style-prompt business` - Translation style
- `--context-file glossary.json` - Terminology glossary
- `--analyze` - Analyze file structure without translating

### Environment Variables

Required for translation operations:
```bash
export GOOGLE_API_KEY="your-api-key-here"
```

### File Analysis

Most translators support `--analyze` mode to inspect document structure before translation:
```bash
python excel_translator_21.py complex_file.xlsx --analyze
```

## Code Architecture Patterns

### Translation Pipeline

1. **File Loading**: Format-specific document parsing
2. **Content Extraction**: Text extraction while preserving metadata
3. **Context Building**: Smart context from headers, surrounding content
4. **API Translation**: Batch processing with retry logic
5. **Content Replacement**: Preserving original formatting
6. **File Saving**: Format-specific output generation

### Error Handling

Translators implement robust error handling:
- API rate limiting with exponential backoff
- JSON parsing fallbacks for malformed LLM responses
- File corruption detection and reporting
- Graceful degradation for unsupported features

### Batch Processing

Scripts support both single file and directory processing:
- `input_file.ext` - Single file mode
- `--source-location ./directory` - Batch mode with subdirectory structure preservation

## Testing

Limited test coverage exists:
- `src/test.py` - Basic translation functionality test
- Use `--analyze` mode for debugging file structure issues
- Enable `--debug` flag for detailed logging

## File Organization

- `src/` - All translation scripts and utilities
- `input_files/` - Sample input files for testing
- `output_files/` - Generated translations
- `requirements.txt` - Python dependencies
- `devenv.nix` - Development environment configuration