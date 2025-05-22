import json
import argparse
import os  # Import os for environment variables
import re  # ADD THIS LINE: Import the 're' module for regular expressions

import google.generativeai as genai  # Import the Gemini API client


def extract_text_elements(filepath):
    """
    Extracts text elements from an Excalidraw JSON file.
    (No changes in this function)
    """
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)

        if not isinstance(data, dict) or data.get('type') != 'excalidraw':
            print(f"Error: File '{filepath}' is not a valid Excalidraw JSON file.")
            return None

        text_elements = {}
        for element in data.get('elements', []):
            if element.get('type') == 'text':
                text_elements[element['id']] = element['text']
        return text_elements

    except FileNotFoundError:
        print(f"Error: File not found: {filepath}")
        return None
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON format in file: {filepath}")
        return None


def extract_json(response_text):
    """Extract JSON from markdown code blocks or raw text (from DOCX example)"""
    # Try to find JSON code blocks
    matches = re.findall(r'```(?:json)?\n(.*?)\n```', response_text, re.DOTALL)
    if matches:
        return matches[0]
    # If no code blocks, try to find first JSON structure
    match = re.search(r'{(.*?)}', response_text, re.DOTALL)
    if match:
        return f'{{{match.group(1)}}}'
    return response_text


def translate_text_batch_gemini(text_elements_dict, source_lang, target_lang, batch_size=5, model_name="gemini-1.5-flash"):
    """
    Translates a dictionary of text elements in batches using Gemini API.
    (No changes in this function except for error message improvement)
    """
    translated_texts_dict = {}
    text_items = list(text_elements_dict.items()) # Get (id, text) pairs

    # Initialize Gemini API (using your working DOCX code's initialization)
    try:
        genai.configure(api_key=os.getenv('GOOGLE_API_KEY')) # Ensure GOOGLE_API_KEY is set in environment
        model = genai.GenerativeModel(model_name)
    except Exception as e:
        print(f"Error initializing Gemini client: {e}. Make sure GOOGLE_API_KEY is set correctly.") # Improved error message
        return None

    print(f"🚀 Processing {len(text_items)} text elements in {len(text_items)//batch_size + 1} batches")


    for batch_start in range(0, len(text_items), batch_size):
        batch_items = text_items[batch_start:batch_start + batch_size] # Get batch of (id, text) tuples
        batch_data = [{"id": element_id, "text": text} for element_id, text in batch_items] # Prepare batch data for prompt
        batch_num = (batch_start // batch_size) + 1
        print(f"🔧 Batch {batch_num} ({len(batch_data)} elements)")


        prompt = f"""
        Translate from {source_lang} to {target_lang}. Maintain formatting EXACTLY. If a \\n exists persist the \\n in the result
        Return ONLY VALID JSON using this format:
        {{
            "translations": [
                {{"id": <original_id>, "translation": "<translated_text>"}}
            ]
        }}
        DO NOT USE MARKDOWN. Ensure proper JSON escaping.
        Input: {json.dumps(batch_data, ensure_ascii=False)}
        """

        try:
            response = model.generate_content(prompt)
            print(f"📥 Raw Response: {response.text[:150]}...") # Print snippet of raw response

            cleaned_response = extract_json(response.text) # Extract JSON from response
            print(f"🧹 Cleaned Response: {cleaned_response[:150]}...") # Print snippet of cleaned response

            try:
                result = json.loads(cleaned_response) # Parse JSON response
            except json.JSONDecodeError as e:
                print(f"❌ JSON Error: {str(e)}")
                print(f"🧬 Response Fragment: {cleaned_response[:500]}")
                continue  # Skip to next batch on JSON error

            if 'translations' not in result:
                print(f"⚠️ Missing 'translations' key in response")
                continue # Skip to next batch if 'translations' key is missing

            success = 0
            for item in result['translations']:
                element_id = item['id'] # Get element ID from translation item
                translated_text = item.get('translation', '') # Get translation, default to empty string on missing
                translated_texts_dict[element_id] = translated_text # Store translation in dict with element ID as key
                success += 1

            print(f"✅ Applied {success}/{len(batch_data)} translations in batch {batch_num}")


        except Exception as e:
            print(f"💥 Batch Error: {str(e)}")
            print(f"📋 Failed Batch Input: {batch_data}")
            return None # Indicate translation failure for the whole process on batch error

    return translated_texts_dict # Return the dictionary of translated texts (ID: translation)


def map_translations_to_elements(original_elements, translations):
    """
    Maps translations back to the original text element IDs, maintaining order.
    (No longer needed as translate_text_batch_gemini now returns a dict)
    """
    return translations # Translations are already in a dict with element IDs as keys


def save_translations_json(translations_dict, output_filepath):
    """
    Saves the translations to a JSON file.
    (No changes in this function)
    """
    try:
        with open(output_filepath, 'w', encoding='utf-8') as outfile:
            json.dump(translations_dict, outfile, indent=4, ensure_ascii=False)
        print(f"Translations saved to: {output_filepath}")
    except Exception as e:
        print(f"Error saving translations to JSON: {e}")



def main():
    parser = argparse.ArgumentParser(description="Translate text elements in Excalidraw files using Gemini 2.0 Flash.")
    parser.add_argument("source_file", help="Path to the source Excalidraw JSON file")
    parser.add_argument("--source-lang", default="en", help="Source language code (default: en)")
    parser.add_argument("--target-lang", default="de", help="Target language code (default: de)") # Default target lang
    parser.add_argument("--batch-size", type=int, default=15, help="Batch size for translation (default: 15)") # Increased default batch size
    parser.add_argument("--output", default="translations.json", help="Output JSON file for translations (default: translations.json)")
    parser.add_argument("--model", default="gemini-2.0-flash", help="Gemini model name (default: gemini-1.5-flash)") # Model name argument


    args = parser.parse_args()

    text_elements = extract_text_elements(args.source_file)
    if text_elements is None:
        return

    if not text_elements:
        print("No text elements found in the Excalidraw file.")
        return

    # Modified to pass text_elements dict directly
    translations_dict = translate_text_batch_gemini(
        text_elements,
        args.source_lang,
        args.target_lang,
        args.batch_size,
        args.model
    )

    if translations_dict: # Check if translations were returned successfully
        # map_translations_to_elements is now redundant as translation returns a dict
        save_translations_json(translations_dict, args.output)
    else:
        print("Translation process failed.")


if __name__ == "__main__":
    main()