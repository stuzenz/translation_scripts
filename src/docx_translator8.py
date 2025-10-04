#!/usr/bin/env python3
"""
docx-translator - Microsoft Word Document Translation Utility with Google Generative AI

A command-line tool for translating Microsoft Word documents while preserving
formatting, styles, tables, and document structure. Enhanced with improved JSON
parsing and error handling.

Features:
- Uses google-generativeai SDK with robust error handling
- Translate single files or entire directories
- Preserve document formatting, styles, and structure
- Support for tables with context-aware translation
- Smart context extraction for better translation quality
- Multiple target languages support
- Style-based translation (business, casual, technical, etc.)
- Robust JSON parsing to handle API response variations
- Optimized for Japanese and other languages

Usage Examples:
    # Translate a single file to Japanese
    python docx-translator.py document.docx --target-lang ja

    # Translate with specific style
    python docx-translator.py report.docx --target-lang ja --style-prompt business

    # Translate to multiple languages
    python docx-translator.py document.docx --target-langs en,ja,es,zh

    # Use glossary for consistent terminology
    python docx-translator.py technical.docx --target-lang ja --context-file glossary.json
"""

import argparse
import json
import logging
import os
import re
import sys
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor, as_completed

# Use existing google-generativeai package
import google.generativeai as genai
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table, _Cell
from docx.shared import RGBColor
from colorama import Fore, Style, init
from tqdm import tqdm

# Initialize colorama for colored output
init(autoreset=True)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Constants
DEFAULT_MODEL = 'gemini-flash-latest'
DEFAULT_BATCH_SIZE = 10
DEFAULT_CONCURRENCY = 4
MAX_RETRIES = 3
RETRY_DELAY = 2

# Style prompts optimized for translations
STYLE_PROMPTS = {
    'business': {
        'ja': 'ビジネス文書として適切な敬語と専門用語を使用し、フォーマルで明確な日本語に翻訳してください。',
        'default': 'Translate using formal business language appropriate for professional documents.'
    },
    'casual': {
        'ja': '自然で親しみやすい日本語に翻訳し、適切な丁寧さを保ってください。',
        'default': 'Translate using conversational, friendly language suitable for general communication.'
    },
    'technical': {
        'ja': '技術的な正確性を優先し、専門用語は適切な日本語の技術用語を使用してください。',
        'default': 'Translate using precise technical terminology, maintaining accuracy for specialized content.'
    },
    'academic': {
        'ja': '学術的な文体で、正確性と論理性を重視した日本語に翻訳してください。',
        'default': 'Translate using scholarly language appropriate for academic contexts.'
    },
    'marketing': {
        'ja': '説得力があり、読者の心に響く魅力的な日本語表現を使用してください。',
        'default': 'Translate using persuasive, engaging language suitable for marketing materials.'
    }
}


@dataclass
class TranslationConfig:
    """Configuration for translation job"""
    source_lang: str
    target_lang: str
    model_name: str
    batch_size: int
    concurrency: int
    style_prompt: str
    context_file: Optional[str]
    smart_context: bool
    debug: bool


class ImprovedJSONExtractor:
    """Enhanced JSON extraction with multiple fallback strategies"""
    
    @staticmethod
    def extract_json(response_text: str) -> Dict[str, Any]:
        """Extract JSON from response with multiple strategies"""
        # Strategy 1: Try direct JSON parsing
        try:
            return json.loads(response_text)
        except json.JSONDecodeError:
            pass
        
        # Strategy 2: Extract from markdown code blocks
        code_block_patterns = [
            r'```json\s*\n(.*?)\n```',
            r'```\s*\n(.*?)\n```',
            r'`(.*?)`'
        ]
        
        for pattern in code_block_patterns:
            matches = re.findall(pattern, response_text, re.DOTALL)
            if matches:
                for match in matches:
                    try:
                        return json.loads(match)
                    except json.JSONDecodeError:
                        continue
        
        # Strategy 3: Find JSON-like structures
        json_patterns = [
            r'\{[^{}]*"translations"[^{}]*\}',  # Simple JSON with translations
            r'\{.*?"translations".*?\}',         # More flexible
            r'\{.*\}',                           # Any JSON object
        ]
        
        for pattern in json_patterns:
            matches = re.findall(pattern, response_text, re.DOTALL)
            if matches:
                # Try the longest match first
                for match in sorted(matches, key=len, reverse=True):
                    try:
                        # Clean up common issues
                        cleaned = match
                        # Remove trailing commas
                        cleaned = re.sub(r',\s*}', '}', cleaned)
                        cleaned = re.sub(r',\s*]', ']', cleaned)
                        # Fix escaped quotes that shouldn't be escaped
                        cleaned = cleaned.replace('\\"', '"')
                        # Remove any BOM or zero-width spaces
                        cleaned = cleaned.replace('\ufeff', '').replace('\u200b', '')
                        
                        result = json.loads(cleaned)
                        if isinstance(result, dict):
                            return result
                    except json.JSONDecodeError:
                        continue
        
        # Strategy 4: Try to fix common JSON issues
        cleaned_text = response_text
        # Remove any text before the first {
        first_brace = cleaned_text.find('{')
        if first_brace != -1:
            cleaned_text = cleaned_text[first_brace:]
        # Remove any text after the last }
        last_brace = cleaned_text.rfind('}')
        if last_brace != -1:
            cleaned_text = cleaned_text[:last_brace + 1]
        
        try:
            return json.loads(cleaned_text)
        except json.JSONDecodeError:
            pass
        
        # If all strategies fail, raise an exception with helpful info
        raise ValueError(f"Could not extract valid JSON from response. First 200 chars: {response_text[:200]}")


def get_docx_files_from_directory(directory):
    """Get all DOCX files from a directory"""
    dir_path = Path(directory)
    if not dir_path.exists():
        raise FileNotFoundError(f"Directory '{directory}' does not exist")

    # Get all .docx files, excluding temporary files (starting with ~$)
    docx_files = [f for f in dir_path.glob("*.docx") if not f.name.startswith("~$")]

    if not docx_files:
        print(f"{Fore.YELLOW}⚠️ No DOCX files found in '{directory}'{Style.RESET_ALL}")
    else:
        print(f"{Fore.CYAN}📁 Found {len(docx_files)} DOCX files in '{directory}'{Style.RESET_ALL}")
        for file in docx_files:
            print(f"   📄 {file.name}")

    return docx_files


class DocxTranslator:
    """Main translator class for Word documents using google-genai 1.27.0"""
    
    def __init__(self, config: TranslationConfig):
        self.config = config
        self.json_extractor = ImprovedJSONExtractor()
        self._setup_logging()
        self._configure_api()
        self.context_data = self._load_context()
        
    def _setup_logging(self):
        """Configure logging based on debug flag"""
        level = logging.DEBUG if self.config.debug else logging.INFO
        logging.getLogger().setLevel(level)
        self.logger = logging.getLogger(__name__)
        
    def _configure_api(self):
        """Configure Google Generative AI API"""
        api_key = os.getenv('GOOGLE_API_KEY')
        if not api_key:
            raise ValueError(
                f"{Fore.RED}Error: GOOGLE_API_KEY environment variable not set.{Style.RESET_ALL}\n"
                f"{Fore.YELLOW}Please set your Google API key as an environment variable.{Style.RESET_ALL}"
            )

        # Configure the API
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel(self.config.model_name)
        self.logger.info(f"Initialized google-generativeai with model: {self.config.model_name}")

        # Test the API connection
        if not self._test_api_connection():
            raise ValueError(f"{Fore.RED}Failed to connect to Google Generative AI API{Style.RESET_ALL}")
    
    def _test_api_connection(self) -> bool:
        """Test API connection with a simple request"""
        try:
            test_prompt = 'Respond with only the JSON: {"status": "ok"}'
            response = self.model.generate_content(test_prompt)

            # Check if we got a response
            if response and response.text:
                self.logger.info("API connection test successful")
                return True
            else:
                self.logger.error("API connection test failed - no response text")
                return False

        except Exception as e:
            self.logger.error(f"API connection test failed: {str(e)}")
            return False
        
    def _load_context(self) -> Dict:
        """Load context/glossary file if provided"""
        if not self.config.context_file:
            return {}
            
        try:
            with open(self.config.context_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.logger.info(f"Loaded context from {self.config.context_file}")
                return data
        except Exception as e:
            self.logger.warning(f"Failed to load context file: {e}")
            return {}
    
    def _get_style_prompt(self) -> str:
        """Get appropriate style prompt for the target language"""
        style_dict = STYLE_PROMPTS.get(self.config.style_prompt, {})
        return style_dict.get(self.config.target_lang, style_dict.get('default', ''))
    
    def _build_translation_prompt(self, texts: List[Dict[str, Any]], context: List[str] = None) -> str:
        """Build simple prompt matching the working version"""

        # Create batch data in the same format as working version
        batch_data = [{"id": i, "text": text_obj["text"]} for i, text_obj in enumerate(texts)]

        # Simple prompt that matches the working docx_translator6.py exactly
        prompt = f"""
        Translate from {self.config.source_lang} to {self.config.target_lang}. Maintain formatting EXACTLY.
        Return ONLY VALID JSON using this format:
        {{
            "translations": [
                {{"id": <original_id>, "translation": "<translated_text>"}}
            ]
        }}
        DO NOT USE MARKDOWN. Ensure proper JSON escaping.
        The topic material for translation is a terms of reference document for an interim state architecture.
        Input: {json.dumps(batch_data, ensure_ascii=False)}
        """

        return prompt
    
    # Removed async method - using sync method only for stability
    
    def translate_batch(self, texts: List[Dict[str, Any]], context: List[str] = None) -> Dict[int, str]:
        """Translate a batch using sync API (fallback method)"""
        if not texts:
            return {}
        
        prompt = self._build_translation_prompt(texts, context)
        
        for attempt in range(MAX_RETRIES):
            try:
                # Use the simple google-generativeai API call (same as working version)
                response = self.model.generate_content(prompt)

                # Extract text from the response
                response_text = response.text

                if not response_text:
                    self.logger.error(f"Empty response. Full response: {str(response)[:1000]}")
                    raise ValueError("Received empty response from API")
                
                self.logger.debug(f"Response text (first 500 chars): {response_text[:500]}")
                
                # Extract and parse JSON
                result = self.json_extractor.extract_json(response_text)
                
                # Validate and build translation map
                if 'translations' not in result:
                    self.logger.error(f"No translations key in result: {result}")
                    raise ValueError("Missing 'translations' key in response")
                
                translation_map = {}
                for item in result['translations']:
                    idx = item.get('id')
                    translation = item.get('translation', '')
                    if idx is not None:
                        translation_map[idx] = translation
                
                # Add fallbacks for missing translations
                for i in range(len(texts)):
                    if i not in translation_map:
                        self.logger.warning(f"Missing translation for ID {i}, using original")
                        translation_map[i] = texts[i]['text']
                
                return translation_map
                
            except Exception as e:
                self.logger.warning(f"Translation attempt {attempt + 1} failed: {str(e)}")
                
                # Log more details on last attempt
                if attempt == MAX_RETRIES - 1:
                    self.logger.error(f"Final attempt failed with error: {str(e)}")
                    self.logger.error(f"Prompt was: {prompt[:500]}...")
                
                if "429" in str(e) or "quota" in str(e).lower() or "rate" in str(e).lower():
                    wait_time = RETRY_DELAY * (2 ** (attempt + 1))
                    self.logger.info(f"Rate limit detected, waiting {wait_time} seconds...")
                    time.sleep(wait_time)
                elif attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY * (attempt + 1))
                else:
                    # Return original texts as fallback
                    self.logger.error("All translation attempts failed, using original text")
                    return {i: text['text'] for i, text in enumerate(texts)}
        
        # Fallback
        return {i: text['text'] for i, text in enumerate(texts)}
    
    def _extract_document_context(self, doc: Document) -> List[str]:
        """Extract context from the document"""
        context = []
        
        # Get document title from first heading
        for para in doc.paragraphs[:5]:
            if para.style.name.startswith('Heading') or para.style.name == 'Title':
                if para.text.strip():
                    context.append(f"[Document Title] {para.text.strip()}")
                    break
        
        # Get section headings
        heading_count = 0
        for para in doc.paragraphs:
            if para.style.name.startswith('Heading') and para.text.strip():
                context.append(f"[Section] {para.text.strip()}")
                heading_count += 1
                if heading_count >= 10:
                    break
        
        # Get first few paragraphs for additional context
        para_count = 0
        for para in doc.paragraphs:
            if para.text.strip() and not para.style.name.startswith('Heading'):
                context.append(f"[Content Sample] {para.text.strip()[:100]}...")
                para_count += 1
                if para_count >= 3:
                    break
        
        return context
    
    def _process_paragraph(self, para: Paragraph, translations: Dict[int, str], text_index: int) -> int:
        """Process a single paragraph"""
        for run in para.runs:
            if run.text.strip():
                if text_index in translations:
                    # Preserve formatting while updating text
                    original_text = run.text
                    translated_text = translations[text_index]
                    
                    # Preserve leading/trailing whitespace
                    leading_space = len(original_text) - len(original_text.lstrip())
                    trailing_space = len(original_text) - len(original_text.rstrip())
                    
                    run.text = ' ' * leading_space + translated_text.strip() + ' ' * trailing_space
                text_index += 1
        return text_index
    
    def _process_table_cell(self, cell: _Cell, translations: Dict[int, str], text_index: int) -> int:
        """Process a table cell"""
        for para in cell.paragraphs:
            text_index = self._process_paragraph(para, translations, text_index)
        return text_index

    def _mark_toc_for_update(self, doc: Document):
        """Mark TOC fields to be updated when the document is opened in Word"""
        try:
            # Access the document part to modify field update settings
            doc_part = doc.part

            # Get the document element
            doc_element = doc_part.element

            # Look for field elements in the document
            from docx.oxml.ns import qn

            # Find all field elements (fldSimple and fldChar)
            fields = doc_element.xpath('.//w:fldSimple[@w:instr]',
                                    namespaces=doc_element.nsmap)

            toc_fields_found = 0
            for field in fields:
                instr = field.get(qn('w:instr'))
                if instr and 'TOC' in instr.upper():
                    # Mark field as dirty (needs update)
                    field.set(qn('w:dirty'), 'true')
                    toc_fields_found += 1

            # Also look for complex fields (fldChar approach)
            fld_chars = doc_element.xpath('.//w:fldChar[@w:fldCharType="begin"]',
                                        namespaces=doc_element.nsmap)

            for fld_char in fld_chars:
                # Check if this is followed by a TOC instruction
                next_element = fld_char.getnext()
                while next_element is not None:
                    if next_element.tag.endswith('}instrText'):
                        if 'TOC' in next_element.text.upper():
                            # Mark the field as dirty
                            fld_char.set(qn('w:dirty'), 'true')
                            toc_fields_found += 1
                            break
                    elif next_element.tag.endswith('}fldChar'):
                        break
                    next_element = next_element.getnext()

            if toc_fields_found > 0:
                self.logger.info(f"📑 Marked {toc_fields_found} TOC field(s) for update")
                self.logger.info("💡 Word will prompt to update the Table of Contents when document is opened")
            else:
                self.logger.debug("No TOC fields found in document")

        except Exception as e:
            self.logger.warning(f"Could not mark TOC for update: {e}")
            # This is not critical, so we continue
    
    def translate_document(self, input_path: Path, output_path: Path):
        """Translate a Word document"""
        self.logger.info(f"Processing: {input_path}")
        
        try:
            # Load document
            doc = Document(input_path)
            
            # Extract document context if smart context is enabled
            doc_context = []
            if self.config.smart_context:
                doc_context = self._extract_document_context(doc)
                self.logger.info(f"Extracted {len(doc_context)} context items")
            
            # Collect all texts to translate
            texts_to_translate = []
            
            # Process paragraphs
            for para_idx, para in enumerate(doc.paragraphs):
                for run_idx, run in enumerate(para.runs):
                    if run.text.strip():
                        texts_to_translate.append({
                            'id': len(texts_to_translate),
                            'text': run.text,
                            'type': 'paragraph',
                            'style': para.style.name,
                            'location': f'Paragraph {para_idx}, Run {run_idx}'
                        })
            
            # Process tables
            for table_idx, table in enumerate(doc.tables):
                for row_idx, row in enumerate(table.rows):
                    for cell_idx, cell in enumerate(row.cells):
                        for para_idx, para in enumerate(cell.paragraphs):
                            for run_idx, run in enumerate(para.runs):
                                if run.text.strip():
                                    texts_to_translate.append({
                                        'id': len(texts_to_translate),
                                        'text': run.text,
                                        'type': 'table_cell',
                                        'location': f'Table {table_idx}, Row {row_idx}, Cell {cell_idx}'
                                    })
            
            if not texts_to_translate:
                self.logger.warning(f"No text found to translate in {input_path}")
                doc.save(output_path)
                return
            
            self.logger.info(f"Found {len(texts_to_translate)} text segments to translate")
            
            # Translate in batches with progress bar
            all_translations = {}
            
            with tqdm(total=len(texts_to_translate), desc="Translating") as pbar:
                for i in range(0, len(texts_to_translate), self.config.batch_size):
                    batch = texts_to_translate[i:i + self.config.batch_size]
                    
                    # Prepare batch for API (only id and text)
                    batch_for_api = [{'id': j, 'text': item['text']} 
                                     for j, item in enumerate(batch)]
                    
                    # Add batch context info
                    batch_context = doc_context.copy()
                    if batch:
                        batch_context.append(f"[Batch Info] Translating {batch[0]['type']} elements")
                    
                    # Translate batch
                    batch_translations = self.translate_batch(batch_for_api, batch_context)
                    
                    # Map back to original indices
                    for j, translation in batch_translations.items():
                        original_idx = i + j
                        all_translations[original_idx] = translation
                    
                    pbar.update(len(batch))
            
            # Apply translations back to document
            text_index = 0
            
            # Apply to paragraphs
            for para in doc.paragraphs:
                text_index = self._process_paragraph(para, all_translations, text_index)
            
            # Apply to tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text_index = self._process_table_cell(cell, all_translations, text_index)
            
            # Mark TOC fields for update before saving
            self._mark_toc_for_update(doc)

            # Save translated document
            output_path.parent.mkdir(parents=True, exist_ok=True)
            doc.save(output_path)
            self.logger.info(f"✅ Saved translated document: {output_path}")
            
        except Exception as e:
            self.logger.error(f"Failed to process {input_path}: {e}")
            if self.config.debug:
                import traceback
                traceback.print_exc()
            raise


def main():
    """Main entry point"""
    parser = argparse.ArgumentParser(
        description="Word Document Translation Utility with Google Generative AI",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    # Input options
    parser.add_argument("input_file", nargs='?', help="Single input DOCX file path")
    parser.add_argument("--input-files", help="Directory containing DOCX files to process")

    # Language options
    parser.add_argument('--source-lang', default='auto',
                       help='Source language code (default: auto-detect)')

    parser.add_argument('--target-lang', help='Target language code (e.g., ja)')
    parser.add_argument('--target-langs', help='Multiple target languages (comma-separated)')

    # Translation options
    parser.add_argument('--style-prompt', choices=list(STYLE_PROMPTS.keys()),
                       default='business', help='Translation style')
    parser.add_argument('--context-file', help='Path to glossary/context JSON file')
    parser.add_argument('--smart-context', action='store_true',
                       help='Use document structure for better context')

    # Output options
    parser.add_argument('--output-dir', default='.', help='Output directory')
    parser.add_argument("--output-files", help="Directory to save translated files (same as --output-dir)")
    parser.add_argument('--model', default=DEFAULT_MODEL, help='Gemini model name')
    parser.add_argument('--batch-size', type=int, default=DEFAULT_BATCH_SIZE,
                       help=f'Texts per API call (default: {DEFAULT_BATCH_SIZE})')
    parser.add_argument('--concurrency', type=int, default=DEFAULT_CONCURRENCY,
                       help=f'Concurrent API calls (default: {DEFAULT_CONCURRENCY})')
    
    # Other options
    parser.add_argument('--debug', action='store_true', help='Enable debug logging')
    parser.add_argument('--list-styles', action='store_true', help='List available styles')
    parser.add_argument('--version', action='store_true', help='Show version information')
    
    args = parser.parse_args()
    
    # Handle version request
    if args.version:
        print(f"{Fore.CYAN}DOCX Translator using google-generativeai{Style.RESET_ALL}")
        print(f"Model: {DEFAULT_MODEL}")
        sys.exit(0)
    
    # Handle info requests
    if args.list_styles:
        print(f"\n{Fore.CYAN}Available translation styles:{Style.RESET_ALL}")
        for style, prompts in STYLE_PROMPTS.items():
            print(f"\n  {Fore.GREEN}{style}:{Style.RESET_ALL}")
            for lang, prompt in prompts.items():
                if lang != 'default':
                    lang_name = 'Japanese' if lang == 'ja' else lang.upper()
                    print(f"    {lang_name}: {prompt[:60]}...")
                else:
                    print(f"    Default: {prompt[:60]}...")
        sys.exit(0)
    
    # Validate input arguments
    if not args.input_file and not args.input_files:
        parser.error("Either provide a single input file or use --input-files for directory processing")

    if args.input_file and args.input_files:
        parser.error("Cannot use both single input file and --input-files together")

    # Handle input files
    input_files = []
    if args.input_file:
        input_path = Path(args.input_file)
        if not input_path.exists():
            print(f"{Fore.RED}❌ Input file '{args.input_file}' does not exist{Style.RESET_ALL}")
            sys.exit(1)
        if not input_path.suffix.lower() == '.docx':
            print(f"{Fore.RED}Error: Input file must be a .docx file{Style.RESET_ALL}")
            sys.exit(1)
        input_files = [args.input_file]
    elif args.input_files:
        try:
            input_files = get_docx_files_from_directory(args.input_files)
            if not input_files:
                sys.exit(1)
        except FileNotFoundError as e:
            print(f"{Fore.RED}❌ {str(e)}{Style.RESET_ALL}")
            sys.exit(1)

    # Handle target languages
    target_langs = []
    if args.target_langs:
        target_langs = [lang.strip() for lang in args.target_langs.split(',')]
    elif args.target_lang:
        target_langs = [args.target_lang]
    else:
        parser.error("At least one target language must be specified using --target-lang or --target-langs")

    # Handle output directory
    output_dir = Path(args.output_files) if args.output_files else Path(args.output_dir)

    # Show header
    print(f"\n{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}DOCX Translator - Google Generative AI{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    print(f"Model: {args.model}")
    print(f"Source: {args.source_lang}")
    print(f"Targets: {', '.join(target_langs)}")
    print(f"Style: {args.style_prompt}")
    print(f"Files: {len(input_files)}")
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}\n")

    # Process all files
    all_output_files = []
    total_files = len(input_files)
    
    for i, input_file in enumerate(input_files, 1):
        print(f"\n{Fore.BLUE}📊 Processing file {i}/{total_files}: {Path(input_file).name}{Style.RESET_ALL}")
        input_path = Path(input_file)

        for target_lang in target_langs:
            print(f"\n{Fore.CYAN}Translating to {target_lang}...{Style.RESET_ALL}")

            # Create configuration
            config = TranslationConfig(
                source_lang=args.source_lang,
                target_lang=target_lang,
                model_name=args.model,
                batch_size=args.batch_size,
                concurrency=args.concurrency,
                style_prompt=args.style_prompt,
                context_file=args.context_file,
                smart_context=args.smart_context,
                debug=args.debug
            )

            # Create translator
            try:
                translator = DocxTranslator(config)
            except ValueError as e:
                print(str(e))
                sys.exit(1)

            # Generate output filename
            output_name = f"{input_path.stem}_{target_lang}{input_path.suffix}"
            output_path = output_dir / output_name

            # Translate document
            try:
                start_time = time.time()
                translator.translate_document(input_path, output_path)
                elapsed_time = time.time() - start_time
                print(f"{Fore.GREEN}✅ Successfully created: {output_path}{Style.RESET_ALL}")
                print(f"{Fore.CYAN}   Time taken: {elapsed_time:.2f} seconds{Style.RESET_ALL}")
                all_output_files.append(output_path)
            except Exception as e:
                print(f"{Fore.RED}❌ Translation failed: {e}{Style.RESET_ALL}")
                if args.debug:
                    import traceback
                    traceback.print_exc()
                continue

    # Summary
    print(f"\n{Fore.GREEN}🎉 Translation completed!{Style.RESET_ALL}")
    print(f"{Fore.CYAN}📊 Summary:{Style.RESET_ALL}")
    print(f"   📄 Input files processed: {total_files}")
    print(f"   🌍 Target languages: {len(target_langs)} ({', '.join(target_langs)})")
    print(f"   📁 Output files created: {len(all_output_files)}")

    if all_output_files:
        print(f"\n{Fore.CYAN}📁 Output files:{Style.RESET_ALL}")
        for output_file in all_output_files:
            print(f"   ✅ {output_file}")

        print(f"\n{Fore.YELLOW}📑 Table of Contents Note:{Style.RESET_ALL}")
        print(f"   If your documents contain Table of Contents, Word will prompt")
        print(f"   to update them when you open the translated files. Click 'Yes'")
        print(f"   to refresh the TOC with translated headings.")


if __name__ == "__main__":
    main()
