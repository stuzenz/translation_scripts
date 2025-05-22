# xlsx-translator Usage Guide

A comprehensive guide for using the Excel Translation Utility with Gemini API.

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [Basic Usage](#basic-usage)
- [Advanced Features](#advanced-features)
- [Command Line Options](#command-line-options)
- [Examples](#examples)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)

## Installation

### Prerequisites

1. Python 3.8 or higher
2. Google Gemini API key (not required for `--analyze` mode)

### Setup

1. **Install required packages:**
   ```bash
   pip install openpyxl google-generativeai colorama tqdm
   ```

2. **Set up your Gemini API key (for translation):**
   ```bash
   # Linux/Mac
   export GOOGLE_API_KEY="your-api-key-here"
   
   # Windows (Command Prompt)
   set GOOGLE_API_KEY=your-api-key-here
   
   # Windows (PowerShell)
   $env:GOOGLE_API_KEY="your-api-key-here"
   ```

3. **Make the script executable (Linux/Mac):**
   ```bash
   chmod +x xlsx-translator.py
   ```

## Quick Start

### Translate a single Excel file to Japanese:
```bash
python xlsx-translator.py report.xlsx --target-lang ja
```

### Translate all Excel files in a directory to multiple languages:
```bash
python xlsx-translator.py --source-location ./excel_files --target-langs en,es,fr
```

### Analyze file structure before translating:
```bash
python xlsx-translator.py complex_data.xlsx --analyze
```

**Note:** The `--analyze` flag only analyzes files without translating. Don't combine it with `--target-lang` or `--target-langs` as they will be ignored.

## Basic Usage

### Single File Translation

Translate a single Excel file to a target language:

```bash
# Basic translation
python xlsx-translator.py input.xlsx --target-lang es

# Specify source language (auto-detect by default)
python xlsx-translator.py input.xlsx --source-lang en --target-lang ja

# Save to specific directory
python xlsx-translator.py input.xlsx --target-lang fr --output-dir ./translations
```

### Directory Processing

Process all Excel files in a directory:

```bash
# Translate all .xlsx and .xlsm files
python xlsx-translator.py --source-location ./reports --target-lang de

# Multiple target languages with subdirectories
python xlsx-translator.py --source-location ./data --target-langs en,ja,ko,zh
```

### Output Structure

- Single file: `original_name_{lang}.xlsx`
- Directory mode: Creates subdirectories per language
  ```
  output_dir/
  ├── en/
  │   ├── report1_en.xlsx
  │   └── report2_en.xlsx
  ├── ja/
  │   ├── report1_ja.xlsx
  │   └── report2_ja.xlsx
  ```

## Advanced Features

### Smart Context Translation

Enable enhanced context extraction for better translation quality:

```bash
# Use smart context mode
python xlsx-translator.py data.xlsx --target-lang es --smart-context

# Smart context analyzes:
# - Column headers
# - Sample data from columns
# - Neighboring cells
# - Cell types and patterns
```

### Sheet and Cell Filtering

Process specific parts of Excel files:

```bash
# Translate only specific sheet by name
python xlsx-translator.py workbook.xlsx --sheet-name "Sales Data" --target-lang fr

# Translate sheet by index (0-based)
python xlsx-translator.py workbook.xlsx --sheet-index 2 --target-lang de

# Translate specific columns
python xlsx-translator.py data.xlsx --columns "A,C:E,G" --target-lang ja

# Translate specific rows
python xlsx-translator.py data.xlsx --rows "1,5-10,15-20" --target-lang es

# Combine filters
python xlsx-translator.py report.xlsx --sheet-name "Q4" --columns "B:D" --rows "1-50" --target-lang zh
```

### Context Files and Glossaries

Use a glossary for consistent terminology:

```bash
# With glossary file
python xlsx-translator.py technical.xlsx --target-lang ja --context-file glossary.json

# Limit context items
python xlsx-translator.py doc.xlsx --target-lang es --context-file terms.json --max-context-items 20
```

**Glossary file format (glossary.json):**
```json
{
    "Revenue": "収益",
    "Operating Expenses": "営業費用",
    "Net Profit": "純利益",
    "Q1": "第1四半期",
    "FY2024": "2024年度"
}
```

### Style Prompts

Choose translation style based on content type:

```bash
# Available styles: business, casual, technical, marketing, academic, legal, medical

# Business (default)
python xlsx-translator.py report.xlsx --target-lang ja --style-prompt business

# Technical documentation
python xlsx-translator.py specs.xlsx --target-lang de --style-prompt technical

# Marketing materials
python xlsx-translator.py campaign.xlsx --target-lang es --style-prompt marketing

# List all available styles
python xlsx-translator.py --list-styles
```

### File Analysis

Analyze Excel structure without translating:

```bash
# Analyze single file
python xlsx-translator.py complex_data.xlsx --analyze

# Analyze with debug information for troubleshooting
python xlsx-translator.py problem_file.xlsx --analyze --debug

# Analyze all files in directory
python xlsx-translator.py --source-location ./reports --analyze
```

Analysis shows:
- Sheet names and structure
- Sheet dimensions (rows × columns)
- Total cells, text cells, formula cells
- Detected headers
- Column information with sample data
- Table detection
- Any errors or issues with the file

The analyze mode is particularly useful for:
- Understanding file structure before translation
- Debugging problematic files
- Planning which sheets/columns to translate
- Identifying empty or corrupted sheets

### Performance Optimization

Control API usage and processing speed:

```bash
# Increase batch size (more cells per API call)
python xlsx-translator.py large_file.xlsx --target-lang es --batch-size 50

# Increase concurrent API calls
python xlsx-translator.py data.xlsx --target-lang ja --concurrency 8

# Optimize for large files
python xlsx-translator.py huge_dataset.xlsx --target-lang fr --batch-size 100 --concurrency 10
```

## Command Line Options

### Required Arguments

| Option | Description |
|--------|-------------|
| `input_file` | Path to Excel file (.xlsx or .xlsm) |
| `--source-location` | Directory containing Excel files (alternative to input_file) |

### Language Options

| Option | Description | Default |
|--------|-------------|---------|
| `--source-lang` | Source language code (e.g., en, ja, es) | auto |
| `--target-lang` | Single target language code | - |
| `--target-langs` | Multiple target languages (comma-separated) | - |

### Filtering Options

| Option | Description | Example |
|--------|-------------|---------|
| `--sheet-name` | Process only this sheet by name | "Sales Data" |
| `--sheet-index` | Process only this sheet by index (0-based) | 0 |
| `--columns` | Columns to translate | "A,C:E,G" |
| `--rows` | Rows to translate | "1,5-10,15" |

### Processing Options

| Option | Description | Default |
|--------|-------------|---------|
| `--preserve-formatting` | Preserve cell formatting | True |
| `--preserve-formulas` | Skip cells with formulas | True |
| `--smart-context` | Use enhanced context extraction | False |
| `--style-prompt` | Translation style | business |
| `--context-file` | Path to glossary JSON file | None |
| `--max-context-items` | Max glossary items to use | 10 |

### Output Options

| Option | Description | Default |
|--------|-------------|---------|
| `--output-dir` | Output directory | . (current) |
| `--model` | Gemini model to use | gemini-2.0-flash |
| `--batch-size` | Cells per API call | 20 |
| `--concurrency` | Concurrent API calls | 4 |

### Information Options

| Option | Description |
|--------|-------------|
| `--analyze` | Analyze file structure without translating |
| `--list-languages` | Show language code information |
| `--list-styles` | List available style options |
| `--debug` | Enable debug logging |
| `--help` | Show help message |

## Examples

### Example 1: Financial Report Translation

Translate a financial report with business terminology:

```bash
python xlsx-translator.py financial_report_2024.xlsx \
    --target-lang ja \
    --style-prompt business \
    --context-file financial_terms.json \
    --smart-context
```

### Example 2: Multi-Language Product Catalog

Translate product catalog to multiple languages:

```bash
python xlsx-translator.py --source-location ./catalogs \
    --target-langs es,fr,de,it,pt \
    --style-prompt marketing \
    --preserve-formatting \
    --output-dir ./international_catalogs
```

### Example 3: Technical Specification Sheets

Translate only data sheets, preserving formulas:

```bash
python xlsx-translator.py tech_specs.xlsx \
    --sheet-name "Specifications" \
    --columns "B:F" \
    --target-lang de \
    --style-prompt technical \
    --preserve-formulas
```

### Example 4: Survey Data with Smart Context

Translate survey responses with contextual understanding:

```bash
python xlsx-translator.py survey_results.xlsx \
    --target-lang es \
    --smart-context \
    --rows "2-1000" \
    --batch-size 50
```

### Example 5: Batch Processing with Analysis

Analyze and translate a directory of reports:

```bash
# First, analyze the structure
python xlsx-translator.py --source-location ./monthly_reports --analyze

# Then translate with optimized settings
python xlsx-translator.py --source-location ./monthly_reports \
    --target-langs en,zh,ja \
    --smart-context \
    --concurrency 6 \
    --output-dir ./translated_reports
```

### Example 6: Pre-Translation Analysis Workflow

Best practice workflow with analysis:

```bash
# Step 1: Analyze the file structure
python xlsx-translator.py complex_report.xlsx --analyze --debug

# Step 2: Based on analysis, translate specific sheets
python xlsx-translator.py complex_report.xlsx \
    --sheet-name "Financial Data" \
    --target-lang ja \
    --smart-context \
    --preserve-formulas

# Step 3: For problematic files, use debug mode
python xlsx-translator.py problem_file.xlsx --analyze --debug
```

Note: When using `--analyze`, target language options are ignored as no translation occurs.

## Best Practices

### 1. Pre-Translation Analysis

Always analyze complex files first:
```bash
python xlsx-translator.py complex_file.xlsx --analyze
```

### 2. Use Smart Context for Data Files

Enable smart context for files with:
- Multiple related columns
- Headers and data rows
- Technical or domain-specific content

### 3. Optimize Batch Sizes

- Small files (< 1000 cells): Use default settings
- Medium files (1000-10000 cells): `--batch-size 50`
- Large files (> 10000 cells): `--batch-size 100 --concurrency 8`

### 4. Create Glossaries for Consistency

For technical or specialized content, create a glossary:
```json
{
    "API": "API",
    "SDK": "SDK",
    "Revenue": "収益",
    "Dashboard": "ダッシュボード"
}
```

### 5. Handle Formula-Heavy Files

For files with many formulas:
```bash
python xlsx-translator.py calculations.xlsx \
    --target-lang es \
    --preserve-formulas \
    --columns "A:C"  # Only translate label columns
```

### 6. Language Codes

Common language codes:
- English: `en`
- Spanish: `es`
- French: `fr`
- German: `de`
- Italian: `it`
- Portuguese: `pt`
- Russian: `ru`
- Japanese: `ja`
- Korean: `ko`
- Chinese (Simplified): `zh`
- Chinese (Traditional): `zh-TW`
- Arabic: `ar`
- Hindi: `hi`

For full list: `python xlsx-translator.py --list-languages`

## Troubleshooting

### Common Issues

#### 1. API Key Not Found
```
Error: GOOGLE_API_KEY environment variable not set.
```
**Solution:** Set your API key as shown in the Installation section.
**Note:** API key is not required for `--analyze` mode.

#### 2. Rate Limiting
```
Translation attempt X failed: Resource exhausted
```
**Solution:** Reduce concurrency or add delays:
```bash
python xlsx-translator.py file.xlsx --target-lang es --concurrency 2
```

#### 3. Memory Issues with Large Files
**Solution:** Process in chunks:
```bash
# Process specific sheets
python xlsx-translator.py large_file.xlsx --sheet-index 0 --target-lang ja
python xlsx-translator.py large_file.xlsx --sheet-index 1 --target-lang ja
```

#### 4. Formula Corruption
**Solution:** Always use `--preserve-formulas` for formula-heavy files.

#### 5. Character Encoding Issues
**Solution:** The tool handles UTF-8 by default. For special characters, ensure your terminal supports UTF-8.

#### 6. "list index out of range" or Similar Errors
**Solution:** Use analyze mode with debug to diagnose:
```bash
# Analyze the problematic file
python xlsx-translator.py problem_file.xlsx --analyze --debug

# Common causes:
# - Empty sheets
# - Corrupted file structure
# - Password-protected files
# - Unsupported Excel features
```

#### 7. Empty or Corrupted Sheets
The analyzer will detect and report:
- Empty sheets (no rows or columns)
- Sheets with dimension issues
- Inaccessible cells
- File format problems

### Debug Mode

Enable debug logging for detailed information:
```bash
# Debug translation issues
python xlsx-translator.py file.xlsx --target-lang es --debug

# Debug analysis issues
python xlsx-translator.py file.xlsx --analyze --debug

# Debug with specific filters
python xlsx-translator.py file.xlsx --sheet-name "Data" --debug --analyze
```

Debug mode provides:
- Detailed error messages with line numbers
- Sheet-by-sheet processing information
- Cell-level error reporting
- API request/response logs (for translation)

### Getting Help

```bash
# Show all options
python xlsx-translator.py --help

# List supported styles
python xlsx-translator.py --list-styles

# Show language information
python xlsx-translator.py --list-languages
```

## Performance Tips

1. **For fastest processing:**
   ```bash
   python xlsx-translator.py file.xlsx --target-lang es --batch-size 100 --concurrency 10
   ```

2. **For highest quality:**
   ```bash
   python xlsx-translator.py file.xlsx --target-lang es --smart-context --batch-size 20
   ```

3. **For large directories:**
   ```bash
   python xlsx-translator.py --source-location ./files --target-langs en,es \
       --concurrency 6 --batch-size 50
   ```

## License and Support

This tool is provided as-is. For issues or feature requests, please refer to the project repository.

Remember to respect API usage limits and terms of service for the Google Gemini API.