# vtt-translator7.py

# ... (Keep all imports and setup from previous version) ...
# Import necessary libraries
import argparse
import re
# from google import genai
import google.generativeai as genai
import os
import json
import time
import logging
from typing import List, Dict, Optional, Tuple, Any
import sys
import asyncio
# from google import genai
# Configure logging
# Add filename and line number to debug logging for easier tracing
log_format = '%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s'
# Set default level to INFO, can be overridden by --debug
logging.basicConfig(level=logging.INFO, format=log_format)


# Initialize genai to None to ensure the name always exists
genai = None
try:
    import google.generativeai as genai_imported # Use temporary name first
    try:
        api_key = os.environ.get("GOOGLE_API_KEY")
        if not api_key:
            raise KeyError("GOOGLE_API_KEY environment variable not set.")
        genai_imported.configure(api_key=api_key)
        genai = genai_imported # Assign only if configure succeeds
        logging.info("Google Generative AI SDK configured successfully.")
    except KeyError as e:
        logging.error(f"FATAL: {e}")
    except Exception as e:
        logging.error(f"FATAL: Error configuring Google Generative AI SDK: {e}")
except ImportError:
    logging.error("FATAL: google-generativeai library not found. Please install it using: pip install google-generativeai")
except Exception as e:
    logging.error(f"FATAL: An unexpected error occurred during genai import/configuration: {e}", exc_info=True)


# --- Constants ---
DEFAULT_MODEL = 'gemini-2.0-flash'
BATCH_SIZE = 25
RETRY_DELAY = 5
MAX_RETRIES = 3
VTT_EXTENSION = '.vtt'
DEFAULT_CONCURRENCY = 10 # Default limit for concurrent API calls

# --- Regex Patterns ---
GUID_PATTERN = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}-\d+$")
# Incorporating user's tested regex structure with named groups, anchors, and optional parts
TIMESTAMP_PATTERN = re.compile(
    r"^\s*"  # Optional leading whitespace (from previous fix)
    r"(?P<start>(?:\d{1,2}:)?\d{2}:\d{2}\.\d{3})"  # Named group 'start', optional HH:, required MM:SS.ms
    r"\s*-->\s*"  # Separator with optional whitespace
    r"(?P<end>(?:\d{1,2}:)?\d{2}:\d{2}\.\d{3})"    # Named group 'end', optional HH:, required MM:SS.ms
    r"(?:\s+.*)?$" # Optional trailing style info (like align:start) until end of line
)
# --- Type Hinting ---
VttEntry = Dict[str, str]
VttData = List[VttEntry]
TranslationMap = Dict[str, str]
ParsedVttFileData = Dict[str, Tuple[Optional[VttData], Optional[str]]]

# --- Core Functions ---

def load_context(filepath: Optional[str]) -> str:
    # ... (keep existing implementation) ...
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
    Enhanced debugging for save logic.
    """
    logging.info(f"Attempting to parse file: {filepath}")
    # Basic file checks
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

    # Header check
    if not lines or not lines[0].strip().upper().startswith("WEBVTT"):
        logging.error(f"Invalid VTT file: Missing WEBVTT header in {filepath}")
        return None, None
    header = lines[0].strip()
    logging.debug(f"Read {len(lines)} lines from {filepath}. Header found: '{header}'")

    # --- Format Detection ---
    # (Keep existing format detection logic)
    is_guid_format: Optional[bool] = None
    line_index = 1
    while line_index < len(lines):
        line = lines[line_index].strip()
        if not line: # Skip blank lines during detection
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
        return ([], header) if header else (None, None) # Return empty list if header was found, else fail

    # --- Parsing Loop ---
    current_guid: Optional[str] = None
    current_timestamp: Optional[str] = None
    current_text_lines: List[str] = []
    entry_index = 0 # Used for generating IDs in standard format

    # --- Enhanced Debug Helper ---
    def save_buffered_entry(line_num_for_debug: int):
        nonlocal current_timestamp, current_text_lines, current_guid # State vars
        nonlocal vtt_data # The list we append to
        saved = False
        # Log current state *before* attempting save
        logging.debug(f"SAVE Check (Line ~{line_num_for_debug}): Attempting save. State: GUID='{current_guid}', TS='{current_timestamp}', TextLines Count={len(current_text_lines)}")

        # Check conditions for saving
        if current_guid and current_timestamp and current_text_lines:
            text_content = "\n".join(current_text_lines).strip()
            if text_content:
                entry_to_add = {
                    "guid": current_guid,
                    "timestamp": current_timestamp,
                    "text": text_content
                }
                # Log exactly what is being added
                logging.debug(f"SAVE ACTION (Line ~{line_num_for_debug}): ADDING Entry: GUID={entry_to_add['guid']}, TS={entry_to_add['timestamp']}, Text='{text_content[:60].replace(chr(10),'/')}{'...' if len(text_content)>60 else ''}'")
                vtt_data.append(entry_to_add)
                saved = True
            else:
                 # Log why saving didn't happen (empty text)
                 logging.debug(f"SAVE ACTION (Line ~{line_num_for_debug}): SKIPPING Save for GUID='{current_guid}', TS='{current_timestamp}' because effective text was empty after strip.")
                 # Still consider this block processed, state will be reset below
        else:
            # Log exactly which condition failed if save wasn't attempted
            fail_reason = []
            if not current_guid: fail_reason.append("GUID is missing")
            if not current_timestamp: fail_reason.append("Timestamp is missing")
            if not current_text_lines: fail_reason.append("TextLines is empty")
            # Only log failure if at least one component was present (avoid spamming at start)
            if current_guid or current_timestamp or current_text_lines:
                 logging.debug(f"SAVE Check (Line ~{line_num_for_debug}): Conditions NOT MET for save. Reason(s): {', '.join(fail_reason)}.")

        # Reset state *if* we determined an entry block structure was present (GUID+TS),
        # regardless of whether it was saved (due to empty text) or not (due to missing text lines).
        # This prevents carrying over old timestamps/guids incorrectly.
        if current_guid and current_timestamp:
            logging.debug(f"SAVE Check (Line ~{line_num_for_debug}): Resetting TS and TextLines state (GUID='{current_guid}' was present).")
            current_timestamp = None
            current_text_lines = []
            # Don't reset current_guid here; it's set by the next identifier line.

        return saved

    logging.debug(f"Starting parsing loop. Format detected: {'GUID' if is_guid_format else 'Standard'}")
    # Iterate through lines *with index* for better logging
    for i, line_raw in enumerate(lines[1:]): # Skip header line
        line_num = i + 2 # Adjust for 1-based indexing and skipping header
        line = line_raw.strip()
        logging.debug(f"PARSER (Line {line_num}): Processing line: '{line[:100].replace(chr(10),'/')}{'...' if len(line)>100 else ''}'")

        # --- Handle Blank Lines ---
        if not line:
            logging.debug(f"PARSER (Line {line_num}): Blank line encountered. Attempting save of previous block.")
            save_buffered_entry(line_num) # Attempt to save the buffered entry
            continue

        # --- GUID Format Logic ---
        if is_guid_format:
            is_guid_match = GUID_PATTERN.match(line)
            if is_guid_match:
                logging.debug(f"PARSER (Line {line_num}): Found GUID.")
                save_buffered_entry(line_num) # Save previous entry first
                current_guid = line
                current_timestamp = None # Reset timestamp for the new GUID block
                current_text_lines = []
                continue

            if current_guid: # Only process lines if we are within a GUID block
                is_timestamp_match = TIMESTAMP_PATTERN.match(line)
                if is_timestamp_match:
                    if current_timestamp: # Check if overwriting timestamp - unusual
                         logging.warning(f"PARSER (Line {line_num}): Overwriting existing timestamp '{current_timestamp}' with new one '{line}' for GUID '{current_guid}'.")
                    logging.debug(f"PARSER (Line {line_num}): Found Timestamp for GUID {current_guid}.")
                    current_timestamp = line
                    current_text_lines = [] # Reset text for the new timestamp
                    continue

                if current_timestamp: # If we have GUID and TS, append text
                    logging.debug(f"PARSER (Line {line_num}): Appending text for GUID {current_guid}.")
                    current_text_lines.append(line_raw.rstrip('\n\r'))
                    continue
                else: # Line after GUID but before TS
                    logging.debug(f"PARSER (Line {line_num}): Ignoring line (have GUID '{current_guid}', but no TS yet): '{line[:100]}...'")
                    continue
            else: # Line before the first GUID (or between saved blocks)
                 logging.debug(f"PARSER (Line {line_num}): Ignoring line (expecting GUID): '{line[:100]}...'")
                 continue

        # --- Standard Format Logic ---
        else: # is_guid_format is False
            is_timestamp_match = TIMESTAMP_PATTERN.match(line)
            if is_timestamp_match:
                logging.debug(f"PARSER (Line {line_num}): Found Timestamp (Standard Format).")
                save_buffered_entry(line_num) # Save previous entry first
                # Start new entry state
                current_timestamp = line
                current_guid = f"entry-{entry_index}" # Generate GUID
                logging.debug(f"PARSER (Line {line_num}): Assigned new state: GUID='{current_guid}', TS='{current_timestamp}'")
                entry_index += 1
                current_text_lines = []
                continue

            if current_timestamp: # If we have a timestamp, append following non-empty lines as text
                logging.debug(f"PARSER (Line {line_num}): Appending text for entry {current_guid}.")
                current_text_lines.append(line_raw.rstrip('\n\r'))
                continue
            else: # Line before the first timestamp (or between blocks if format is weird)
                 logging.debug(f"PARSER (Line {line_num}): Ignoring line (no timestamp context yet): '{line[:100]}...'")
                 continue

    # --- End of File ---
    logging.debug("PARSER: End of file reached. Attempting final save.")
    save_buffered_entry(len(lines) + 1) # Save any remaining buffered entry

    logging.debug(f"PARSER: Finished parsing loop for {filepath}. Total entries collected: {len(vtt_data)}")
    if vtt_data:
        logging.info(f"Successfully parsed {len(vtt_data)} entries from {filepath} (Format: {'GUID' if is_guid_format else 'Standard'}). First TS: {vtt_data[0].get('timestamp')}, Last TS: {vtt_data[-1].get('timestamp')}")
    elif header: # Parsed header but no entries
         logging.info(f"Successfully parsed {filepath} but found 0 valid entries (Format: {'GUID' if is_guid_format else 'Standard'}).")
    else: # Should not happen if header check passed
         logging.error(f"Parsing completed for {filepath} but result is invalid (no data, no header?).")
         return None, None

    # Final check for integrity (optional, but good practice)
    # for idx, entry in enumerate(vtt_data):
    #     if not all(k in entry and entry[k] is not None for k in ["guid", "timestamp", "text"]):
    #          logging.error(f"PARSER: Corrupted entry found at index {idx} in {filepath}: {entry}")
    #          # Decide how to handle: return None, filter, or just warn
    #          # return None, None # Safest option

    return vtt_data, header

async def translate_batch(
    texts: List[str],
    target_lang: str,
    context: str,
    model_name: str,
    semaphore: asyncio.Semaphore
    ) -> Optional[TranslationMap]:
    """
    Asynchronously translates a batch of text snippets using the Generative AI model,
    controlling concurrency with a semaphore. Handles retries and JSON issues.
    """
    # --- Start of Function ---
    if not genai:
        logging.error("Generative AI client is not initialized. Cannot translate.")
        return None # Cannot proceed
    if not texts:
        logging.debug("Translate batch called with empty list of texts.")
        return {}

    # --- Separate Valid Texts from Empty/Whitespace ---
    original_texts = texts
    valid_texts_for_api = [t for t in original_texts if t and not t.isspace()]
    final_translation_map: TranslationMap = {t: t for t in original_texts if not t or t.isspace()}

    if not valid_texts_for_api:
        logging.debug("Translate batch: No valid (non-empty) texts to send to API.")
        return final_translation_map # Return map with only identity for empty/whitespace

    # --- Prepare API Call ---
    input_snippets_json = json.dumps(valid_texts_for_api, indent=2, ensure_ascii=False)

    prompt = f"""Translate the following text snippets to {target_lang}. The original snippets might be from a transcription and contain mixed languages (e.g., English and Japanese). Focus on translating the primary language of the transcription or follow standard translation practice if mixed. If a snippet is already entirely in {target_lang}, return it unchanged.
Ensure the translation accurately reflects the original meaning and tone.
Preserve the original line breaks (\\n) within each translated snippet.
Do NOT add extra explanation or introductory text.

{f"Use this context/glossary for specialist terms: {context}" if context else ""}

Return ONLY a valid JSON object mapping each original snippet (key) to its corresponding {target_lang} translation (value). The keys in the JSON MUST EXACTLY match the input snippets provided below, including whitespace and line breaks.

Input Snippets (JSON Array):
{input_snippets_json}

Required JSON Output (Map<String, String>):
"""

    # Create model without any configuration options
    model = genai.GenerativeModel(model_name)
    retries = 0
    last_exception = None
    response_text = "" # Store last response text for debugging

    # --- API Call with Semaphore and Retries ---
    async with semaphore:
        logging.debug(f"Semaphore acquired for {target_lang} batch ({len(valid_texts_for_api)} snippets). Concurrency active.")
        while retries <= MAX_RETRIES:
            api_response_map : Optional[Dict] = None # Map returned by API
            try:
                logging.debug(f"Attempt {retries+1}/{MAX_RETRIES+1}: Sending batch to {model_name} for {target_lang} translation.")
                
                # SIMPLEST POSSIBLE API CALL with NO extra parameters
                response = await model.generate_content_async(prompt)
                response_text = response.text # Store raw text immediately

                # Attempt to parse JSON - we'll try multiple approaches
                try:
                    # First try direct JSON parsing
                    api_response_map = json.loads(response_text)
                    logging.debug(f"Successfully parsed JSON directly for {target_lang} batch.")
                except json.JSONDecodeError:
                    # If that fails, try to extract JSON using regex
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

                # --- Validate Response ---
                if not isinstance(api_response_map, dict):
                     raise ValueError(f"LLM response was valid JSON but not an object/map. Got: {type(api_response_map)}")

                missing_keys = []
                processed_keys_map = {} # Holds translations from this successful API call
                for original_text in valid_texts_for_api:
                    if original_text in api_response_map:
                        processed_keys_map[original_text] = api_response_map[original_text]
                    else:
                        missing_keys.append(original_text)
                        processed_keys_map[original_text] = original_text # Fallback for this attempt
                        logging.warning(f"API response for {target_lang} (Attempt {retries+1}) missing key: '{original_text[:50]}...'. Applying fallback.")

                if missing_keys:
                     logging.warning(f"API response for {target_lang} (Attempt {retries+1}) was incomplete: Missing {len(missing_keys)}/{len(valid_texts_for_api)} keys. Fallbacks applied.")

                # Merge results from this attempt into the final map
                final_translation_map.update(processed_keys_map)
                logging.debug(f"Successfully processed batch for {target_lang} (Attempt {retries+1}).")
                return final_translation_map # SUCCESS for this batch

            # --- Exception Handling ---
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

            # --- Retry Logic ---
            retries += 1
            if retries <= MAX_RETRIES:
                logging.info(f"Retrying batch for {target_lang} (Attempt {retries+1}/{MAX_RETRIES+1}) after {RETRY_DELAY} seconds due to: {type(last_exception).__name__}")
                await asyncio.sleep(RETRY_DELAY)
            else:
                # --- MAX RETRIES REACHED ---
                logging.error(f"Max retries ({MAX_RETRIES}) reached for {target_lang} batch. Failing batch permanently.")
                logging.error(f"Final error for {target_lang}: {type(last_exception).__name__} - {last_exception}")
                if response_text:
                     logging.error(f"Last raw response text received for {target_lang} batch: '{response_text[:500]}...'")
                logging.warning(f"Applying fallback (original text) for all {len(valid_texts_for_api)} snippets in {target_lang} batch after max retries.")
                for key in valid_texts_for_api:
                     final_translation_map[key] = key # Use original text as fallback
                return final_translation_map # Return map with fallbacks

    # --- Fallback if Loop Exits Unexpectedly ---
    logging.error(f"Translate batch function terminated unexpectedly after semaphore release for {target_lang}. Applying fallback.")
    for key in valid_texts_for_api:
         if key not in final_translation_map:
             final_translation_map[key] = key
    return final_translation_map

# ... (write_vtt function remains the same) ...
def write_vtt(output_filepath: str, vtt_data: VttData, header: str, translations: TranslationMap):
    """Writes the translated VTT data to a file."""
    try:
        output_dir = os.path.dirname(output_filepath)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        with open(output_filepath, 'w', encoding='utf-8') as f:
            f.write(header + "\n\n") # Ensure blank line after header
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
                f.write(translated_text + "\n\n") # Ensure blank line after each entry
                entry_count += 1
        logging.info(f"Successfully wrote {entry_count} entries to translated VTT: {output_filepath}")
    except Exception as e:
        logging.error(f"Error writing VTT file {output_filepath}: {e}", exc_info=True)


# --- Main Execution (Async) ---
async def main():
    if not genai:
        logging.error("FATAL: Google Generative AI client failed to initialize. Exiting.")
        sys.exit(1)

    parser = argparse.ArgumentParser(description="Translate WebVTT files (GUID or standard format) using Google Generative AI (Async).")

    # Arguments (including --debug)
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument("input_vtt", nargs='?', default=None, help="Path to a single input VTT file (use this OR --source-location).")
    input_group.add_argument("--source-location", help="Directory containing VTT files to process (use this OR the single positional argument).")
    lang_group = parser.add_mutually_exclusive_group(required=True)
    lang_group.add_argument("--target-lang", help="Single target language code (e.g., 'en', 'ja').")
    lang_group.add_argument("--target-langs", nargs='+', help="List of target language codes (e.g., 'en' 'ja').")
    parser.add_argument("--context-file", help="Path to a glossary or context file (optional).")
    parser.add_argument("--output-dir", default=".", help="Base directory to save output files (default: current directory). Creates language subfolders.")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"Name of the Gemini model to use (default: {DEFAULT_MODEL}).")
    parser.add_argument("--batch-size", type=int, default=BATCH_SIZE, help=f"Number of text snippets per API call (default: {BATCH_SIZE}).")
    parser.add_argument("--concurrency", type=int, default=DEFAULT_CONCURRENCY,
                        help=f"Max number of concurrent API calls (default: {DEFAULT_CONCURRENCY}). Adjust based on API limits.")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging.")


    args = parser.parse_args()

    # Set logging level based on debug flag
    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
        logging.debug("Debug logging enabled.")
    else:
        # Ensure logger is set to INFO if debug is not enabled
        logging.getLogger().setLevel(logging.INFO)


    # --- Determine Input Files ---
    source_files = []
    source_location_used = False
    # ... (file discovery logic remains the same) ...
    if args.source_location:
        source_dir = args.source_location
        source_location_used = True
        logging.info(f"Processing VTT files from directory specified via --source-location: {source_dir}")
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
                    else:
                         logging.warning(f"Found item ending with {VTT_EXTENSION} but it's not a file, skipping: {filename}")
            logging.info(f"Found {found_files} VTT files in '{source_dir}'.")
            if found_files == 0:
                logging.warning(f"No VTT files found in the specified directory: {source_dir}")
        except OSError as e:
            logging.error(f"FATAL: Error accessing source directory {source_dir}: {e}")
            sys.exit(1)
    elif args.input_vtt:
        if not os.path.isfile(args.input_vtt):
             logging.error(f"FATAL: Specified input VTT file not found or is not a file: {args.input_vtt}")
             sys.exit(1)
        source_files = [args.input_vtt]
        logging.info(f"Processing single source file: {args.input_vtt}")
    else:
        logging.error("FATAL: No input VTT file or source location specified.")
        parser.print_help()
        sys.exit(1)

    if not source_files:
         if source_location_used:
              logging.warning("No source VTT files found in the specified directory. Exiting.")
         else:
              logging.error("FATAL: No valid source VTT files to process.")
         sys.exit(0)


    # --- Preparations ---
    context = load_context(args.context_file)

    # --- Parse all source files ---
    logging.info("--- Starting VTT Parsing Phase ---")
    parsed_data: ParsedVttFileData = {}
    valid_files_count = 0
    total_files = len(source_files)
    parsing_start_time = time.time()
    for idx, filepath in enumerate(source_files):
        logging.info(f"Parsing file {idx+1}/{total_files}: {filepath}")
        # Call the parser with enhanced debugging
        vtt_data, vtt_header = parse_vtt(filepath)
        if vtt_data is not None and vtt_header is not None:
            parsed_data[filepath] = (vtt_data, vtt_header)
            valid_files_count += 1
            # Log success summary (INFO level) moved to end of parse_vtt
        else:
            logging.warning(f"Skipping file {idx+1}/{total_files} due to parsing errors: {filepath}")
    parsing_end_time = time.time()
    logging.info(f"--- Finished VTT Parsing Phase ({parsing_end_time - parsing_start_time:.2f} seconds) ---")

    if valid_files_count == 0:
        logging.error("No VTT files could be successfully parsed (check logs for details). Exiting.")
        sys.exit(1)
    logging.info(f"Finished parsing. {valid_files_count} out of {total_files} files parsed successfully overall.")


    # --- Collect unique snippets ---
    logging.info("--- Collecting Unique Text Snippets for Translation ---")
    all_texts_to_translate_set = set()
    total_entries_across_files = 0
    # ... (snippet collection logic remains the same) ...
    for filepath, (vtt_data, _) in parsed_data.items():
         if vtt_data: # Check if list is not None
            total_entries_across_files += len(vtt_data)
            for entry_idx, entry in enumerate(vtt_data):
                text = entry.get("text")
                if text and not text.isspace():
                    all_texts_to_translate_set.add(text)
                elif text is None:
                     logging.warning(f"Found entry with missing 'text' key in parsed data (File: {filepath}, GUID/ID: {entry.get('guid')}, Index: {entry_idx}). Skipping snippet.")

    all_texts_to_translate = sorted(list(all_texts_to_translate_set))
    logging.info(f"Found {len(all_texts_to_translate)} unique non-empty text snippets across {total_entries_across_files} total entries from {valid_files_count} files requiring translation.")

    if not all_texts_to_translate:
         logging.warning("No text snippets found requiring translation. Output files will contain original text.")


    # --- Determine Target Languages ---
    target_languages = []
    # ... (language determination logic remains the same) ...
    if args.target_lang:
        target_languages.append(args.target_lang)
    elif args.target_langs:
        target_languages = args.target_langs

    if not target_languages:
        logging.error("FATAL: No target languages specified.")
        sys.exit(1)
    logging.info(f"Target languages: {', '.join(target_languages)}")


    # --- Setup Concurrency and Output Dirs ---
    semaphore = asyncio.Semaphore(args.concurrency)
    logging.info(f"Concurrency limit set to {args.concurrency}")
    overall_start_time = time.time() # Reset start time to just before translation

    output_base_dir = args.output_dir
    lang_output_dirs = {}
    valid_target_languages = []
    # ... (output dir creation logic remains the same) ...
    for lang in target_languages:
        lang_dir = os.path.join(output_base_dir, lang)
        try:
            os.makedirs(lang_dir, exist_ok=True)
            lang_output_dirs[lang] = lang_dir
            valid_target_languages.append(lang)
            logging.info(f"Ensured output directory exists: {lang_dir}")
        except OSError as e:
            logging.error(f"Cannot create output directory for language '{lang}': {lang_dir}. Error: {e}. Skipping this language.")

    if not valid_target_languages:
        logging.error("FATAL: Could not create any output directories for the specified languages. Exiting.")
        sys.exit(1)
    if len(valid_target_languages) < len(target_languages):
        skipped_langs = set(target_languages) - set(valid_target_languages)
        logging.warning(f"Will not process languages due to output directory errors: {', '.join(skipped_langs)}")


    # --- Async Translation and Writing Process ---
    logging.info("--- Starting Translation and Writing Phase ---")
    # ... (translation loop remains the same) ...
    for lang in valid_target_languages:
        lang_start_time = time.time()
        logging.info(f"--- Processing Language: {lang} ---")
        master_translation_map_for_lang: TranslationMap = {}

        if all_texts_to_translate:
            # --- Prepare and Run Translation Tasks ---
            tasks = []
            all_batches_original_texts = [] # For error reporting/fallback check
            total_batches = (len(all_texts_to_translate) + args.batch_size - 1) // args.batch_size
            logging.info(f"Preparing {total_batches} translation batches for '{lang}'...")

            for i in range(0, len(all_texts_to_translate), args.batch_size):
                batch_num = i // args.batch_size + 1
                batch_texts = all_texts_to_translate[i : i + args.batch_size]
                all_batches_original_texts.append(batch_texts)
                logging.debug(f"Creating task for batch {batch_num}/{total_batches} ({len(batch_texts)} snippets) for '{lang}'")
                task = translate_batch(batch_texts, lang, context, args.model, semaphore)
                tasks.append(task)

            logging.info(f"Starting concurrent API calls for {len(tasks)} batches into '{lang}'...")
            results = await asyncio.gather(*tasks, return_exceptions=True)
            logging.info(f"Finished API calls for '{lang}'. Processing {len(results)} batch results...")

            # --- Process Translation Results ---
            successful_batches = 0
            failed_batches = 0
            total_snippets_processed = 0
            # snippets_with_fallback = 0 # Hard to track accurately here

            for i, result in enumerate(results):
                batch_num = i + 1
                batch_original_texts = all_batches_original_texts[i]
                total_snippets_processed += len(batch_original_texts)

                if isinstance(result, Exception):
                    failed_batches += 1
                    logging.error(f"Batch {batch_num}/{total_batches} for '{lang}' failed entirely with exception: {result}")
                    for text in batch_original_texts:
                        master_translation_map_for_lang[text] = text # Fallback
                        # snippets_with_fallback += 1
                elif result is None:
                    failed_batches += 1
                    logging.error(f"Batch {batch_num}/{total_batches} for '{lang}' returned None unexpectedly.")
                    for text in batch_original_texts:
                         master_translation_map_for_lang[text] = text # Fallback
                         # snippets_with_fallback += 1
                else:
                    successful_batches += 1
                    logging.debug(f"Processing successful result for batch {batch_num}/{total_batches} for '{lang}'.")
                    master_translation_map_for_lang.update(result)
                    # Check for internal fallbacks is omitted for simplicity here

            logging.info(f"Processed translation results for '{lang}': {successful_batches} successful batches, {failed_batches} failed batches.")

        else:
             logging.info(f"No text snippets required translation for language '{lang}'.")

        # --- Write output files for this specific language ---
        logging.info(f"Starting to write {len(parsed_data)} output files for language '{lang}'...")
        files_written_for_lang = 0
        output_dir_for_lang = lang_output_dirs[lang] # Get pre-created dir path

        for input_filepath, parse_result in parsed_data.items():
            vtt_data, vtt_header = parse_result
            if vtt_data is None or vtt_header is None:
                logging.warning(f"Skipping writing for '{input_filepath}' (lang: {lang}) due to earlier parsing issues.")
                continue

            base_filename = os.path.splitext(os.path.basename(input_filepath))[0]
            output_filename = f"{base_filename}{VTT_EXTENSION}"
            output_filepath = os.path.join(output_dir_for_lang, output_filename)

            logging.debug(f"Writing file for [{lang}] derived from '{input_filepath}' to: {output_filepath}")
            write_vtt(output_filepath, vtt_data, vtt_header, master_translation_map_for_lang)
            files_written_for_lang +=1

        lang_end_time = time.time()
        logging.info(f"Finished writing {files_written_for_lang} files for language '{lang}'. Time taken: {lang_end_time - lang_start_time:.2f} seconds.")
        logging.info(f"--- Finished Processing for Language: {lang} ---")

    overall_end_time = time.time()
    logging.info("--- Batch translation process completed for all specified languages. ---")
    logging.info(f"Total execution time: {overall_end_time - overall_start_time:.2f} seconds.")


# --- Entry Point ---
if __name__ == "__main__":
    if genai:
        try:
            print(f"Google Generative AI SDK version: {genai.__version__}")
            
            # Test available models
            models = genai.list_models()
            print(f"Available models: {[model.name for model in models]}")
            
            # Check GenerationConfig parameters
            import inspect
            params = inspect.signature(genai.types.GenerationConfig.__init__).parameters
            print(f"Available GenerationConfig parameters: {list(params.keys())}")
        except Exception as e:
            print(f"Error during API inspection: {e}")
    # Ensure platform compatibility for asyncio event loop policy
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
