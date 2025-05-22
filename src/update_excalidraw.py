import json
import argparse

def load_json(filepath):
    """Loads JSON data from a file."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Error: File not found: {filepath}")
        return None
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON format in file: {filepath}")
        return None

def update_excalidraw_with_translations(excalidraw_filepath, translations_filepath, output_filepath=None):
    """
    Updates an Excalidraw JSON file with translations from a translations JSON file.
    Also updates the 'originalText' field with the translated text.

    Args:
        excalidraw_filepath (str): Path to the original Excalidraw JSON file.
        translations_filepath (str): Path to the JSON file containing translations.
        output_filepath (str, optional): Path to save the updated Excalidraw file.
                                         If None, a new file with "_translated" suffix is created.
    """
    excalidraw_data = load_json(excalidraw_filepath)
    translations_data = load_json(translations_filepath)

    if excalidraw_data is None or translations_data is None:
        return

    if not isinstance(excalidraw_data, dict) or excalidraw_data.get('type') != 'excalidraw':
        print(f"Error: File '{excalidraw_filepath}' is not a valid Excalidraw JSON file.")
        return

    if not isinstance(translations_data, dict):
        print(f"Error: File '{translations_filepath}' is not a valid translations JSON file (expecting a dictionary).")
        return

    updated_elements = []
    for element in excalidraw_data.get('elements', []):
        if element.get('type') == 'text':
            element_id = element['id']
            if element_id in translations_data:
                translated_text = translations_data[element_id]
                element['text'] = translated_text  # Update text element with translation
                element['originalText'] = translated_text # Update originalText as well
        updated_elements.append(element) # Keep all elements, translated or not

    excalidraw_data['elements'] = updated_elements

    if output_filepath is None:
        output_filepath = excalidraw_filepath.replace(".json", "_translated.json")

    try:
        with open(output_filepath, 'w', encoding='utf-8') as outfile:
            json.dump(excalidraw_data, outfile, indent=2, ensure_ascii=False)
        print(f"Updated Excalidraw file saved to: {output_filepath}")
    except Exception as e:
        print(f"Error saving updated Excalidraw file: {e}")


def main():
    parser = argparse.ArgumentParser(description="Update Excalidraw JSON with translations from a JSON file.")
    parser.add_argument("excalidraw_file", help="Path to the original Excalidraw JSON file")
    parser.add_argument("translations_file", help="Path to the JSON file containing translations")
    parser.add_argument("--output", help="Output file path for the updated Excalidraw file (optional)")

    args = parser.parse_args()

    update_excalidraw_with_translations(
        args.excalidraw_file,
        args.translations_file,
        args.output
    )

if __name__ == "__main__":
    main()