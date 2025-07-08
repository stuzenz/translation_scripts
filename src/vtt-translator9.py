# vtt-translator8.py - Enhanced with language detection and same-language handling

import argparse
import re
import google.generativeai as genai
import os
import json
import time
import logging
from typing import List, Dict, Optional, Tuple, Any, Set
import sys
import asyncio
from dataclasses import dataclass

# Language detection
try:
    from langdetect import detect, DetectorFactory, LangDetectError
    from langdetect.lang_detect_exception import LangDetectException
    # Set seed for consistent results
    DetectorFactory.seed = 0
    LANG_DETECT_AVAILABLE = True
    logging.info("Language detection available via langdetect")
except ImportError:
    LANG_DETECT_AVAILABLE = False
    logging.warning("langdetect not available. Install with: pip install langdetect")

# Configure logging
log_format = '%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s'
logging.basicConfig(level=logging.INFO, format=log_format)

# Initialize genai
genai = None
try:
    import google.generativeai as genai_imported
    try:
        api_key = os.environ.get("GOOGLE_API_KEY")
        if not api_key:
            raise KeyError("GOOGLE_API_KEY environment variable not set.")
        genai_imported.configure(api_key=api_key)
        genai = genai_imported
        logging.info("Google Generative AI SDK configured successfully.")
    except KeyError as e:
        logging.error(f"FATAL: {e}")
    except Exception as e:
        logging.error(f"FATAL: Error configuring Google Generative AI SDK: {e}")
except ImportError:
    logging.error("FATAL: google-generativeai library not found. Please install it using: pip install google-generativeai")
except Exception as e:
    logging.error(f"FATAL: An unexpected error occurred during genai import/configuration: {e}", exc_info=True)

# Constants
DEFAULT_MODEL = 'gemini-2.0-flash'
BATCH_SIZE = 25
RETRY_DELAY = 5
MAX_RETRIES = 3
VTT_EXTENSION = '.vtt'
DEFAULT_CONCURRENCY = 10
DEFAULT_DETECT_THRESHOLD = 0.7

# Regex Patterns
GUID_PATTERN = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}-\d+$")
TIMESTAMP_PATTERN = re.compile(
    r"^\s*"
    r"(?P<start>(?:\d{1,2}:)?\d{2}:\d{2}\.\d{3})"
    r"\s*-->\s*"
    r"(?P<end>(?:\d{1,2}:)?\d{2}:\d{2}\.\d{3})"
    r"(?:\s+.*)?$"
)

# Type Hinting
VttEntry = Dict[str, str]
VttData = List[VttEntry]
TranslationMap = Dict[str, str]
ParsedVttFileData = Dict[str, Tuple[Optional[VttData], Optional[str]]]

@dataclass
class LanguageDetectionResult:
    """Result of language detection for a text snippet"""
    text: str
    detected_lang: Optional[str]
    confidence: float
    is_target_language: bool
    should_translate: bool
    reason: str

@dataclass
class TranslationConfig:
    """Configuration for translation behavior"""
    skip_same_language: bool = False
    clean_same_language: bool = True
    detect_threshold: float = DEFAULT_DETECT_THRESHOLD
    force_translation: bool = False

# Language code mappings (common aliases)
LANGUAGE_CODE_MAPPINGS = {
    'en': ['en', 'english'],
    'es': ['es', 'spanish', 'spa'],
    'fr': ['fr', 'french', 'fra'],
    'de': ['de', 'german', 'deu'],
    'it': ['it', 'italian', 'ita'],
    'th': ['th', 'thai', 'tha'],
    'vi': ['vi', 'vietnamese', 'vie'],
    'pt': ['pt', 'portuguese', 'por'],
    'ru': ['ru', 'russian', 'rus'],
    'ja': ['ja', 'japanese', 'jpn'],
    'ko': ['ko', 'korean', 'kor'],
    'zh': ['zh', 'chinese', 'zho', 'zh-cn'],
    'zh-tw': ['zh-tw', 'zh-hant', 'traditional chinese'],
    'ar': ['ar', 'arabic', 'ara'],
    'hi': ['hi', 'hindi', 'hin'],
    'tr': ['tr', 'turkish', 'tur'],
    'pl': ['pl', 'polish', 'pol'],
    'nl': ['nl', 'dutch', 'nld'],
    'sv': ['sv', 'swedish', 'swe'],
    'da': ['da', 'danish', 'dan'],
    'no': ['no', 'norwegian', 'nor'],
    'fi': ['fi', 'finnish', 'fin']
}

def normalize_language_code(lang_code: str) -> str:
    """Normalize language code to standard format"""
    lang_code = lang_code.lower().strip()
    
    # Check direct mappings
    for standard, aliases in LANGUAGE_CODE_MAPPINGS.items():
        if lang_code in aliases:
            return standard
    
    # Return original if no mapping found
    return lang_code

def detect_text_language(text: str, target_lang: str, config: TranslationConfig) -> LanguageDetectionResult:
    """
    Detect the language of text and determine if translation is needed
    """
    if not LANG_DETECT_AVAILABLE:
        return LanguageDetectionResult(
            text=text,
            detected_lang=None,
            confidence=0.0,
            is_target_language=False,
            should_translate=True,
            reason="Language detection not available - will translate"
        )
    
    if not text or text.isspace():
        return LanguageDetectionResult(
            text=text,
            detected_lang=None,
            confidence=0.0,
            is_target_language=False,
            should_translate=False,
            reason="Empty or whitespace text"
        )
    
    # Clean text for better detection
    clean_text = re.sub(r'[^\w\s]', ' ', text).strip()
    if len(clean_text) < 3:
        return LanguageDetectionResult(
            text=text,
            detected_lang=None,
            confidence=0.0,
            is_target_language=False,
            should_translate=True,
            reason="Text too short for reliable detection - will translate"
        )
    
    try:
        # Detect language
        detected_lang = detect(clean_text)
        detected_lang = normalize_language_code(detected_lang)
        target_lang_normalized = normalize_language_code(target_lang)
        
        # For confidence, we use a simple heuristic since langdetect doesn't provide confidence
        # We could use detect_langs() for probabilities, but it's more complex
        confidence = 0.8 if len(clean_text) > 20 else 0.6
        
        is_target_language = detected_lang == target_lang_normalized
        
        # Determine if we should translate based on configuration
        if is_target_language:
            if config.skip_same_language:
                should_translate = False
                reason = f"Detected as target language ({detected_lang}) - skipping translation"
            elif config.clean_same_language:
                should_translate = True
                reason = f"Detected as target language ({detected_lang}) - will clean via translation"
            else:
                should_translate = True
                reason = f"Detected as target language ({detected_lang}) - default behavior"
        else:
            should_translate = True
            reason = f"Detected as different language ({detected_lang}) - will translate"
        
        if config.force_translation:
            should_translate = True
            reason += " (forced by configuration)"
        
        return LanguageDetectionResult(
            text=text,
            detected_lang=detected_lang,
            confidence=confidence,
            is_target_language=is_target_language,
            should_translate=should_translate,
            reason=reason
        )
        
    except (LangDetectError, LangDetectException) as e:
        logging.debug(f"Language detection failed for text '{text[:50]}...': {e}")
        return LanguageDetectionResult(
            text=text,
            detected_lang=None,
            confidence=0.0,
            is_target_language=False,
            should_translate=True,
            reason="Language detection failed - will translate"
        )

def load_context(filepath: Optional[str]) -> str:
    context = ""
    if filepath:
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                context = f.read()
            logging.info(f"Loaded context from {filepath}")
        except FileNotFoundError:
            logging.warning(f"Context file not found: {filepath}")
        except Exception as e:
            logging.error(f"Error reading context file {filepath}: {e}")
    return context

def parse_vtt(filepath: str) -> Tuple[Optional[VttData], Optional[str]]:
    """
    Parses a VTT file, automatically detecting GUID-based or standard format.
    """
    logging.info(f"Attempting to parse file: {filepath}")
    
    if not os.path.exists(filepath):
        logging.error(f"Input file not found during parsing: {filepath}")
        return None, None
    if not os.path.isfile(filepath):
        logging.error(f"Input path is not a file: {filepath}")
        return None, None

    vtt_data: VttData = []
    header: Optional[str] = None
    lines: List[str] = []

    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except Exception as e:
        logging.error(f"Error reading VTT file {filepath}: {e}", exc_info=True)
        return None, None

    if not lines or not lines[0].strip().upper().startswith("WEBVTT"):
        logging.error(f"Invalid VTT file: Missing WEBVTT header in {filepath}")
        return None, None
    header = lines[0].strip()
    logging.debug(f"Read {len(lines)} lines from {filepath}. Header found: '{header}'")

    # Format Detection
    is_guid_format: Optional[bool] = None
    line_index = 1
    while line_index < len(lines):
        line = lines[line_index].strip()
        if not line:
            line_index += 1
            continue
        if TIMESTAMP_PATTERN.match(line):
            logging.debug(f"Format Detection: Found potential timestamp at line {line_index+1}: '{line}'")
            prev_line_index = line_index - 1
            while prev_line_index > 0 and not lines[prev_line_index].strip():
                prev_line_index -= 1
            if prev_line_index > 0:
                potential_guid = lines[prev_line_index].strip()
                logging.debug(f"Format Detection: Checking preceding line {prev_line_index+1} for GUID: '{potential_guid}'")
                if GUID_PATTERN.match(potential_guid):
                    is_guid_format = True
                    logging.info(f"Detected GUID-based VTT format for {filepath}")
                else:
                    is_guid_format = False
                    logging.info(f"Detected standard (non-GUID) VTT format for {filepath}")
            else:
                is_guid_format = False
                logging.info(f"Detected standard (non-GUID) VTT format for {filepath} (timestamp follows header).")
            break
        line_index += 1

    if is_guid_format is None:
        has_content = any(line.strip() for line in lines[1:])
        if has_content:
            logging.error(f"Could not determine VTT format for {filepath}. Content found, but no timestamp lines detected after header.")
        else:
            logging.warning(f"VTT file '{filepath}' seems empty after header or contains only header. No entries to parse.")
        return ([], header) if header else (None, None)

    # Parsing Loop
    current_guid: Optional[str] = None
    current_timestamp: Optional[str] = None
    current_text_lines: List[str] = []
    entry_index = 0

    def save_buffered_entry(line_num_for_debug: int):
        nonlocal current_timestamp, current_text_lines, current_guid
        nonlocal vtt_data
        saved = False
        logging.debug(f"SAVE Check (Line ~{line_num_for_debug}): Attempting save. State: GUID='{current_guid}', TS='{current_timestamp}', TextLines Count={len(current_text_lines)}")

        if current_guid and current_timestamp and current_text_lines:
            text_content = "\n".join(current_text_lines).strip()
            if text_content:
                entry_to_add = {
                    "guid": current_guid,
                    "timestamp": current_timestamp,
                    "text": text_content
                }
                logging.debug(f"SAVE ACTION (Line ~{line_num_for_debug}): ADDING Entry: GUID={entry_to_add['guid']}, TS={entry_to_add['timestamp']}, Text='{text_content[:60].replace(chr(10),'/')}{'...' if len(text_content)>60 else ''}'")
                vtt_data.append(entry_to_add)
                saved = True
            else:
                logging.debug(f"SAVE ACTION (Line ~{line_num_for_debug}): SKIPPING Save for GUID='{current_guid}', TS='{current_timestamp}' because effective text was empty after strip.")

        if current_guid and current_timestamp:
            logging.debug(f"SAVE Check (Line ~{line_num_for_debug}): Resetting TS and TextLines state (GUID='{current_guid}' was present).")
            current_timestamp = None
            current_text_lines = []

        return saved

    logging.debug(f"Starting parsing loop. Format detected: {'GUID' if is_guid_format else 'Standard'}")
    
    for i, line_raw in enumerate(lines[1:]):
        line_num = i + 2
        line = line_raw.strip()
        logging.debug(f"PARSER (Line {line_num}): Processing line: '{line[:100].replace(chr(10),'/')}{'...' if len(line)>100 else ''}'")

        if not line:
            logging.debug(f"PARSER (Line {line_num}): Blank line encountered. Attempting save of previous block.")
            save_buffered_entry(line_num)
            continue

        if is_guid_format:
            is_guid_match = GUID_PATTERN.match(line)
            if is_guid_match:
                logging.debug(f"PARSER (Line {line_num}): Found GUID.")
                save_buffered_entry(line_num)
                current_guid = line
                current_timestamp = None
                current_text_lines = []
                continue

            if current_guid:
                is_timestamp_match = TIMESTAMP_PATTERN.match(line)
                if is_timestamp_match:
                    if current_timestamp:
                        logging.warning(f"PARSER (Line {line_num}): Overwriting existing timestamp '{current_timestamp}' with new one '{line}' for GUID '{current_guid}'.")
                    logging.debug(f"PARSER (Line {line_num}): Found Timestamp for GUID {current_guid}.")
                    current_timestamp = line
                    current_text_lines = []
                    continue

                if current_timestamp:
                    logging.debug(f"PARSER (Line {line_num}): Appending text for GUID {current_guid}.")
                    current_text_lines.append(line_raw.rstrip('\n\r'))
                    continue
                else:
                    logging.debug(f"PARSER (Line {line_num}): Ignoring line (have GUID '{current_guid}', but no TS yet): '{line[:100]}...'")
                    continue
            else:
                logging.debug(f"PARSER (Line {line_num}): Ignoring line (expecting GUID): '{line[:100]}...'")
                continue

        else:  # Standard format
            is_timestamp_match = TIMESTAMP_PATTERN.match(line)
            if is_timestamp_match:
                logging.debug(f"PARSER (Line {line_num}): Found Timestamp (Standard Format).")
                save_buffered_entry(line_num)
                current_timestamp = line
                current_guid = f"entry-{entry_index}"
                logging.debug(f"PARSER (Line {line_num}): Assigned new state: GUID='{current_guid}', TS='{current_timestamp}'")
                entry_index += 1
                current_text_lines = []
                continue

            if current_timestamp:
                logging.debug(f"PARSER (Line {line_num}): Appending text for entry {current_guid}.")
                current_text_lines.append(line_raw.rstrip('\n\r'))
                continue
            else:
                logging.debug(f"PARSER (Line {line_num}): Ignoring line (no timestamp context yet): '{line[:100]}...'")
                continue

    logging.debug("PARSER: End of file reached. Attempting final save.")
    save_buffered_entry(len(lines) + 1)

    logging.debug(f"PARSER: Finished parsing loop for {filepath}. Total entries collected: {len(vtt_data)}")
    if vtt_data:
        logging.info(f"Successfully parsed {len(vtt_data)} entries from {filepath} (Format: {'GUID' if is_guid_format else 'Standard'}). First TS: {vtt_data[0].get('timestamp')}, Last TS: {vtt_data[-1].get('timestamp')}")
    elif header:
        logging.info(f"Successfully parsed {filepath} but found 0 valid entries (Format: {'GUID' if is_guid_format else 'Standard'}).")
    else:
        logging.error(f"Parsing completed for {filepath} but result is invalid (no data, no header?).")
        return None, None

    return vtt_data, header

async def translate_batch_with_detection(
    texts: List[str],
    target_lang: str,
    context: str,
    model_name: str,
    semaphore: asyncio.Semaphore,
    config: TranslationConfig
) -> Optional[TranslationMap]:
    """
    Enhanced translation function that handles language detection
    """
    if not genai:
        logging.error("Generative AI client is not initialized. Cannot translate.")
        return None
    
    if not texts:
        logging.debug("Translate batch called with empty list of texts.")
        return {}

    # Analyze each text for language detection
    detection_results = []
    texts_to_translate = []
    texts_to_skip = []
    final_translation_map: TranslationMap = {}

    for text in texts:
        if not text or text.isspace():
            final_translation_map[text] = text
            continue
            
        detection = detect_text_language(text, target_lang, config)
        detection_results.append(detection)
        
        if detection.should_translate:
            texts_to_translate.append(text)
        else:
            texts_to_skip.append(text)
            final_translation_map[text] = text  # Keep original
            logging.debug(f"Skipping translation: {detection.reason}")

    # Log detection summary
    if detection_results:
        same_lang_count = sum(1 for d in detection_results if d.is_target_language)
        different_lang_count = len(detection_results) - same_lang_count
        logging.info(f"Language detection: {same_lang_count} already in {target_lang}, {different_lang_count} in other languages, {len(texts_to_skip)} skipped, {len(texts_to_translate)} to translate")

    # If no texts need translation, return early
    if not texts_to_translate:
        logging.debug("No texts require translation after language detection.")
        return final_translation_map

    # Enhanced prompt for mixed-language scenarios
    input_snippets_json = json.dumps(texts_to_translate, indent=2, ensure_ascii=False)
    
    prompt = f"""Translate the following text snippets to {target_lang}. This content may contain mixed languages.

INSTRUCTIONS:
- If a snippet is already in {target_lang}, you may either:
  - Return it unchanged if it's already well-formed
  - Return a cleaned/improved version if it contains errors or informal language
- For snippets in other languages, translate them accurately to {target_lang}
- Preserve original line breaks (\\n) within each snippet
- Maintain the original tone and meaning
- Do NOT add explanatory text or comments

{f"Use this context/glossary for specialist terms: {context}" if context else ""}

Return ONLY a valid JSON object mapping each original snippet (key) to its corresponding {target_lang} translation or cleaned version (value). Keys must EXACTLY match the input snippets.

Input Snippets (JSON Array):
{input_snippets_json}

Required JSON Output (Map<String, String>):
"""

    # Create model and make API call
    model = genai.GenerativeModel(model_name)
    retries = 0
    last_exception = None
    response_text = ""

    async with semaphore:
        logging.debug(f"Semaphore acquired for {target_lang} batch ({len(texts_to_translate)} snippets). Concurrency active.")
        while retries <= MAX_RETRIES:
            api_response_map: Optional[Dict] = None
            try:
                logging.debug(f"Attempt {retries+1}/{MAX_RETRIES+1}: Sending batch to {model_name} for {target_lang} translation.")
                
                response = await model.generate_content_async(prompt)
                response_text = response.text

                try:
                    api_response_map = json.loads(response_text)
                    logging.debug(f"Successfully parsed JSON directly for {target_lang} batch.")
                except json.JSONDecodeError:
                    import re
                    json_match = re.search(r'\{.*\}', response_text, re.DOTALL)
                    if json_match:
                        try:
                            json_str = json_match.group(0)
                            api_response_map = json.loads(json_str)
                            logging.debug(f"Successfully extracted JSON using regex for {target_lang} batch.")
                        except Exception as e:
                            logging.warning(f"Failed to parse extracted JSON for {target_lang}: {e}")
                            raise ValueError(f"Could not extract valid JSON. Content: {response_text[:200]}...")
                    else:
                        raise ValueError(f"No JSON-like content found in response: {response_text[:200]}...")

                if not isinstance(api_response_map, dict):
                    raise ValueError(f"LLM response was valid JSON but not an object/map. Got: {type(api_response_map)}")

                missing_keys = []
                processed_keys_map = {}
                for original_text in texts_to_translate:
                    if original_text in api_response_map:
                        processed_keys_map[original_text] = api_response_map[original_text]
                    else:
                        missing_keys.append(original_text)
                        processed_keys_map[original_text] = original_text
                        logging.warning(f"API response for {target_lang} (Attempt {retries+1}) missing key: '{original_text[:50]}...'. Applying fallback.")

                if missing_keys:
                    logging.warning(f"API response for {target_lang} (Attempt {retries+1}) was incomplete: Missing {len(missing_keys)}/{len(texts_to_translate)} keys. Fallbacks applied.")

                final_translation_map.update(processed_keys_map)
                logging.debug(f"Successfully processed batch for {target_lang} (Attempt {retries+1}).")
                return final_translation_map

            except json.JSONDecodeError as e:
                logging.warning(f"JSON decode failed for {target_lang} (Attempt {retries+1}/{MAX_RETRIES+1}): {e}. Raw response snippet: '{response_text[:200]}...'")
                last_exception = e
            except ValueError as e_val:
                logging.warning(f"Data validation error for {target_lang} (Attempt {retries+1}/{MAX_RETRIES+1}): {e_val}. Raw response snippet: '{response_text[:200]}...'")
                last_exception = e_val
            except Exception as e_api:
                logging.error(f"Error during LLM call for {target_lang} (Attempt {retries+1}/{MAX_RETRIES+1}): {type(e_api).__name__} - {e_api}", exc_info=False)
                last_exception = e_api
                is_rate_limit = "429" in str(e_api) or "Resource has been exhausted" in str(e_api) or "quota" in str(e_api).lower()
                if is_rate_limit:
                    logging.warning(f"Rate limit likely hit for {target_lang} (Attempt {retries+1}). Retrying after longer delay...")
                    await asyncio.sleep(RETRY_DELAY * (retries + 2))

            retries += 1
            if retries <= MAX_RETRIES:
                logging.info(f"Retrying batch for {target_lang} (Attempt {retries+1}/{MAX_RETRIES+1}) after {RETRY_DELAY} seconds due to: {type(last_exception).__name__}")
                await asyncio.sleep(RETRY_DELAY)
            else:
                logging.error(f"Max retries ({MAX_RETRIES}) reached for {target_lang} batch. Failing batch permanently.")
                logging.error(f"Final error for {target_lang}: {type(last_exception).__name__} - {last_exception}")
                if response_text:
                    logging.error(f"Last raw response text received for {target_lang} batch: '{response_text[:500]}...'")
                logging.warning(f"Applying fallback (original text) for all {len(texts_to_translate)} snippets in {target_lang} batch after max retries.")
                for key in texts_to_translate:
                    final_translation_map[key] = key
                return final_translation_map

    logging.error(f"Translate batch function terminated unexpectedly after semaphore release for {target_lang}. Applying fallback.")
    for key in texts_to_translate:
        if key not in final_translation_map:
            final_translation_map[key] = key
    return final_translation_map

def write_vtt(output_filepath: str, vtt_data: VttData, header: str, translations: TranslationMap):
    """Writes the translated VTT data to a file."""
    try:
        output_dir = os.path.dirname(output_filepath)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        with open(output_filepath, 'w', encoding='utf-8') as f:
            f.write(header + "\n\n")
            entry_count = 0
            for entry in vtt_data:
                if not all(k in entry for k in ["guid", "timestamp", "text"]):
                    logging.warning(f"Skipping invalid entry in write_vtt (missing keys): {entry} in {output_filepath}")
                    continue

                original_text = entry["text"]
                translated_text = translations.get(original_text, original_text)
                translated_text = translated_text if translated_text is not None else ""

                f.write(entry["guid"] + "\n")
                f.write(entry["timestamp"] + "\n")
                f.write(translated_text + "\n\n")
                entry_count += 1
        logging.info(f"Successfully wrote {entry_count} entries to translated VTT: {output_filepath}")
    except Exception as e:
        logging.error(f"Error writing VTT file {output_filepath}: {e}", exc_info=True)

async def main():
    if not genai:
        logging.error("FATAL: Google Generative AI client failed to initialize. Exiting.")
        sys.exit(1)

    parser = argparse.ArgumentParser(description="Translate WebVTT files with language detection and mixed-language support.")

    # Input arguments
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument("input_vtt", nargs='?', default=None, help="Path to a single input VTT file")
    input_group.add_argument("--source-location", help="Directory containing VTT files to process")
    
    # Language arguments
    lang_group = parser.add_mutually_exclusive_group(required=True)
    lang_group.add_argument("--target-lang", help="Single target language code (e.g., 'en', 'ja')")
    lang_group.add_argument("--target-langs", nargs='+', help="List of target language codes")
    
    # Translation behavior
    behavior_group = parser.add_mutually_exclusive_group()
    behavior_group.add_argument("--skip-same-language", action="store_true", default=True, 
                               help="Skip translation for text already in target language (preserve original)")
    behavior_group.add_argument("--clean-same-language", action="store_true",
                               help="Apply translation to same-language text for cleaning (default)")
    
    # Additional options
    parser.add_argument("--force-translation", action="store_true",
                       help="Force translation of all text regardless of detected language")
    parser.add_argument("--detect-threshold", type=float, default=DEFAULT_DETECT_THRESHOLD,
                       help=f"Confidence threshold for language detection (default: {DEFAULT_DETECT_THRESHOLD})")
    parser.add_argument("--context-file", help="Path to a glossary or context file")
    parser.add_argument("--output-dir", default=".", help="Directory to save output files")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"Gemini model to use (default: {DEFAULT_MODEL})")
    parser.add_argument("--batch-size", type=int, default=BATCH_SIZE, help=f"Number of text snippets per API call (default: {BATCH_SIZE})")
    parser.add_argument("--concurrency", type=int, default=DEFAULT_CONCURRENCY, help=f"Max concurrent API calls (default: {DEFAULT_CONCURRENCY})")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")

    args = parser.parse_args()

    # Set logging level
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
        logging.debug("Debug logging enabled.")
    else:
        logging.getLogger().setLevel(logging.INFO)

    # Check language detection availability
    if not LANG_DETECT_AVAILABLE and (args.skip_same_language or not args.force_translation):
        logging.warning("Language detection not available. Install with: pip install langdetect")
        logging.warning("Will translate all text segments without language detection.")

    # Create translation config
    translation_config = TranslationConfig(
        skip_same_language=args.skip_same_language,
        clean_same_language=args.clean_same_language,
        detect_threshold=args.detect_threshold,
        force_translation=args.force_translation
    )

    # Determine input files
    source_files = []
    if args.source_location:
        source_dir = args.source_location
        logging.info(f"Processing VTT files from directory: {source_dir}")
        if not os.path.isdir(source_dir):
            logging.error(f"FATAL: Provided source location is not a valid directory: {source_dir}")
            sys.exit(1)
        
        try:
            found_files = 0
            for filename in os.listdir(source_dir):
                if filename.lower().endswith(VTT_EXTENSION):
                    full_path = os.path.join(source_dir, filename)
                    if os.path.isfile(full_path):
                        source_files.append(full_path)
                        found_files += 1
            logging.info(f"Found {found_files} VTT files in '{source_dir}'.")
            if found_files == 0:
                logging.warning(f"No VTT files found in the specified directory: {source_dir}")
        except OSError as e:
            logging.error(f"FATAL: Error accessing source directory {source_dir}: {e}")
            sys.exit(1)
    elif args.input_vtt:
        if not os.path.isfile(args.input_vtt):
            logging.error(f"FATAL: Specified input VTT file not found: {args.input_vtt}")
            sys.exit(1)
        source_files = [args.input_vtt]
        logging.info(f"Processing single source file: {args.input_vtt}")

    if not source_files:
        logging.warning("No source VTT files found. Exiting.")
        sys.exit(0)

    # Load context
    context = load_context(args.context_file)

    # Parse all source files
    logging.info("--- Starting VTT Parsing Phase ---")
    parsed_data: ParsedVttFileData = {}
    valid_files_count = 0
    total_files = len(source_files)
    parsing_start_time = time.time()
    
    for idx, filepath in enumerate(source_files):
        logging.info(f"Parsing file {idx+1}/{total_files}: {filepath}")
        vtt_data, vtt_header = parse_vtt(filepath)
        if vtt_data is not None and vtt_header is not None:
            parsed_data[filepath] = (vtt_data, vtt_header)
            valid_files_count += 1
        else:
            logging.warning(f"Skipping file {idx+1}/{total_files} due to parsing errors: {filepath}")
    
    parsing_end_time = time.time()
    logging.info(f"--- Finished VTT Parsing Phase ({parsing_end_time - parsing_start_time:.2f} seconds) ---")

    if valid_files_count == 0:
        logging.error("No VTT files could be successfully parsed. Exiting.")
        sys.exit(1)
    logging.info(f"Finished parsing. {valid_files_count} out of {total_files} files parsed successfully.")

    # Collect unique snippets
    logging.info("--- Collecting Unique Text Snippets for Translation ---")
    all_texts_to_translate_set = set()
    total_entries_across_files = 0
    
    for filepath, (vtt_data, _) in parsed_data.items():
        if vtt_data:
            total_entries_across_files += len(vtt_data)
            for entry_idx, entry in enumerate(vtt_data):
                text = entry.get("text")
                if text and not text.isspace():
                    all_texts_to_translate_set.add(text)
                elif text is None:
                    logging.warning(f"Found entry with missing 'text' key in parsed data (File: {filepath}, GUID/ID: {entry.get('guid')}, Index: {entry_idx}). Skipping snippet.")

    all_texts_to_translate = sorted(list(all_texts_to_translate_set))
    logging.info(f"Found {len(all_texts_to_translate)} unique non-empty text snippets across {total_entries_across_files} total entries from {valid_files_count} files.")

    # Determine target languages
    target_languages = []
    if args.target_lang:
        target_languages.append(args.target_lang)
    elif args.target_langs:
        target_languages = args.target_langs

    if not target_languages:
        logging.error("FATAL: No target languages specified.")
        sys.exit(1)
    logging.info(f"Target languages: {', '.join(target_languages)}")

    # Setup concurrency and ensure output directory exists
    semaphore = asyncio.Semaphore(args.concurrency)
    logging.info(f"Concurrency limit set to {args.concurrency}")
    overall_start_time = time.time()

    output_base_dir = args.output_dir
    try:
        os.makedirs(output_base_dir, exist_ok=True)
        logging.info(f"Ensured output directory exists: {output_base_dir}")
    except OSError as e:
        logging.error(f"FATAL: Cannot create output directory: {output_base_dir}. Error: {e}")
        sys.exit(1)

    # Process each language
    logging.info("--- Starting Translation and Writing Phase ---")
    for lang in target_languages:
        lang_start_time = time.time()
        logging.info(f"--- Processing Language: {lang} ---")
        master_translation_map_for_lang: TranslationMap = {}

        if all_texts_to_translate:
            # Prepare and run translation tasks
            tasks = []
            all_batches_original_texts = []
            total_batches = (len(all_texts_to_translate) + args.batch_size - 1) // args.batch_size
            logging.info(f"Preparing {total_batches} translation batches for '{lang}'...")

            for i in range(0, len(all_texts_to_translate), args.batch_size):
                batch_num = i // args.batch_size + 1
                batch_texts = all_texts_to_translate[i : i + args.batch_size]
                all_batches_original_texts.append(batch_texts)
                logging.debug(f"Creating task for batch {batch_num}/{total_batches} ({len(batch_texts)} snippets) for '{lang}'")
                task = translate_batch_with_detection(batch_texts, lang, context, args.model, semaphore, translation_config)
                tasks.append(task)

            logging.info(f"Starting concurrent API calls for {len(tasks)} batches into '{lang}'...")
            results = await asyncio.gather(*tasks, return_exceptions=True)
            logging.info(f"Finished API calls for '{lang}'. Processing {len(results)} batch results...")

            # Process translation results
            successful_batches = 0
            failed_batches = 0
            total_snippets_processed = 0

            for i, result in enumerate(results):
                batch_num = i + 1
                batch_original_texts = all_batches_original_texts[i]
                total_snippets_processed += len(batch_original_texts)

                if isinstance(result, Exception):
                    failed_batches += 1
                    logging.error(f"Batch {batch_num}/{total_batches} for '{lang}' failed entirely with exception: {result}")
                    for text in batch_original_texts:
                        master_translation_map_for_lang[text] = text
                elif result is None:
                    failed_batches += 1
                    logging.error(f"Batch {batch_num}/{total_batches} for '{lang}' returned None unexpectedly.")
                    for text in batch_original_texts:
                        master_translation_map_for_lang[text] = text
                else:
                    successful_batches += 1
                    logging.debug(f"Processing successful result for batch {batch_num}/{total_batches} for '{lang}'.")
                    master_translation_map_for_lang.update(result)

            logging.info(f"Processed translation results for '{lang}': {successful_batches} successful batches, {failed_batches} failed batches.")

        else:
            logging.info(f"No text snippets required translation for language '{lang}'.")

        # Write output files for this language
        logging.info(f"Starting to write {len(parsed_data)} output files for language '{lang}'...")
        files_written_for_lang = 0

        for input_filepath, parse_result in parsed_data.items():
            vtt_data, vtt_header = parse_result
            if vtt_data is None or vtt_header is None:
                logging.warning(f"Skipping writing for '{input_filepath}' (lang: {lang}) due to earlier parsing issues.")
                continue

            base_filename = os.path.splitext(os.path.basename(input_filepath))[0]
            output_filename = f"{base_filename}_{lang}{VTT_EXTENSION}"
            output_filepath = os.path.join(output_base_dir, output_filename)

            logging.debug(f"Writing file for [{lang}] derived from '{input_filepath}' to: {output_filepath}")
            write_vtt(output_filepath, vtt_data, vtt_header, master_translation_map_for_lang)
            files_written_for_lang += 1

        lang_end_time = time.time()
        logging.info(f"Finished writing {files_written_for_lang} files for language '{lang}'. Time taken: {lang_end_time - lang_start_time:.2f} seconds.")
        logging.info(f"--- Finished Processing for Language: {lang} ---")

    overall_end_time = time.time()
    logging.info("--- Batch translation process completed for all specified languages. ---")
    logging.info(f"Total execution time: {overall_end_time - overall_start_time:.2f} seconds.")

# Entry Point
if __name__ == "__main__":
    if genai:
        try:
            print(f"Google Generative AI SDK version: {genai.__version__}")
            models = genai.list_models()
            print(f"Available models: {[model.name for model in models]}")
        except Exception as e:
            print(f"Error during API inspection: {e}")
    
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logging.info("Process interrupted by user (Ctrl+C). Exiting.")
        sys.exit(130)
    except Exception as e:
        logging.error(f"FATAL: An unhandled error occurred during execution: {e}", exc_info=True)
        sys.exit(1)