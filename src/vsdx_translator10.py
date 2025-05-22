import argparse
import json
import os
import shutil
from vsdx import VisioFile
import google.generativeai as genai
from jsonschema import validate, ValidationError

def get_output_filename(input_file, output_pdf=False):
    """Generate output filename with _translated suffix, optionally for PDF"""
    base, ext = os.path.splitext(input_file)
    suffix = "_translated"
    if output_pdf:
        return f"{base}{suffix}.pdf"
    else:
        return f"{base}{suffix}{ext}"

def translate_batch(batch_texts, source_lang, target_lang, model_name):
    """Translate a batch of texts using Gemini API with strict JSON formatting and schema validation"""
    print(f"Entering translate_batch with {len(batch_texts)} texts, source_lang: {source_lang}, target_lang: {target_lang}, model: {model_name}")
    genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
    model = genai.GenerativeModel(model_name)

    translation_schema = {
        "type": "array",
        "items": {
            "type": "string"
        }
    }

    system_instruction = (
        f"Translate the following {len(batch_texts)} texts from {source_lang} to {target_lang}. "
        "Return your response as a JSON array that strictly adheres to the following schema:\n"
        f"{json.dumps(translation_schema)}\n"
        "The JSON array should contain only the translated texts in the same order as the input. "
        "Do not include any additional text, explanations, or formatting outside of the JSON array."
        "For example, if the input is ['Hello', 'World'], the output should be: ['こんにちは', '世界']\n\n"
        "Texts to translate:\n" # Added a clearer separator
    )

    full_prompt = system_instruction + json.dumps(batch_texts, ensure_ascii=False) # Combine instruction and texts

    try:
        print(f"Sending batch to Gemini API. Texts: {batch_texts}")
        response = model.generate_content(
            contents=[
                {'role': 'user', 'parts': [{'text': full_prompt}]} # User role with combined prompt
            ],
            generation_config={
                "temperature": 0.1,
                "max_output_tokens": 5000,
            }
        )

        # Print raw response for debugging
        print(f"Raw API Response Text: {response.text}")

        # Clean response and parse JSON
        response_text = response.text.strip()
        if response_text.startswith('```json'):
            response_text = response_text[len('```json'):]
        if response_text.endswith('```'):
            response_text = response_text[:-len('```')]
        response_text = response_text.strip()

        print(f"Cleaned API Response Text for JSON Parsing: {response_text}")
        translated_list = json.loads(response_text)

        # Validate against schema
        print(f"Validating response against schema: {translation_schema}")
        validate(instance=translated_list, schema=translation_schema)
        print("Schema validation successful.")

        if not isinstance(translated_list, list) or len(translated_list) != len(batch_texts):
            raise ValueError(f"Expected {len(batch_texts)} translations, but response contains {len(translated_list)} items.")

        print("Exiting translate_batch successfully.")
        return translated_list

    except json.JSONDecodeError as e:
        error_message = f"JSON Decode Error during translation: Invalid JSON response from model: {response.text}. Error: {str(e)}"
        print(error_message)
        raise RuntimeError(error_message) from e
    except ValidationError as e:
        error_message = f"Schema Validation Error: Gemini response does not conform to the expected schema: {response.text}. Error: {str(e)}"
        print(error_message)
        raise RuntimeError(error_message) from e
    except Exception as e:
        error_message = f"Translation failed: {str(e)}"
        print(error_message)
        raise RuntimeError(error_message) from e

def convert_to_pdf(input_file, output_file):
    """
    Convert Visio file to PDF using Aspose.Words for Python via .NET.
    Requires aspose-words package and valid license if using Aspose commercially.
    """
    try:
        import aspose.words as aw
        doc = aw.Document(input_file)
        doc.save(output_file)
        print(f"Successfully converted '{input_file}' to PDF '{output_file}' using Aspose.Words.")
    except ImportError:
        print("Aspose.Words for Python via .NET is not installed. Please install it to enable PDF conversion.")
        print("You can install it using pip: pip install aspose-words")
        print("Alternatively, you can explore other Visio to PDF conversion methods or libraries.")
    except Exception as e:
        print(f"Error during PDF conversion using Aspose.Words: {e}")


def process_shapes(visio_file, process_fn):
    """Process all shapes in a Visio file including child_shapes (groups and containers)"""
    print("Entering process_shapes")

    def process_shape(shape):
        print(f"Processing shape with text: '{shape.text if shape.text else 'No Text'}' and ID: {shape.ID}")
        process_fn(shape)
        if hasattr(shape, 'child_shapes'): # Check if the shape has child_shapes
            print(f"Shape ID: {shape.ID} has child_shapes. Processing child_shapes...")
            try:
                child_shapes = shape.child_shapes # Access child_shapes as a property
                for child_shape in child_shapes: # Iterate through child_shapes
                    process_shape(child_shape)
            except Exception as e:
                print(f"Error iterating child_shapes for Shape ID: {shape.ID}. Error: {e}")
        else:
            print(f"Shape ID: {shape.ID} has no child_shapes.")

    for page in visio_file.pages:
        print(f"Processing page: {page.name}")
        # Use page.child_shapes for top-level shapes as per documentation
        for shape in page.child_shapes: # Top level shapes of the page
            process_shape(shape)


    print("Exiting process_shapes")


def translate_visio_file(input_file, source_lang, target_lang, model_name, batch_size, output_pdf=False, dual_language=False):
    """Main translation workflow with PDF output and dual language options"""
    output_vsdx_file = get_output_filename(input_file)
    print(f"Entering translate_visio_file with input_file: {input_file}, output_vsdx_file: {output_vsdx_file}, target_lang: {target_lang}, model: {model_name}, batch_size: {batch_size}, output_pdf: {output_pdf}, dual_language: {dual_language}")

    shutil.copyfile(input_file, output_vsdx_file) # 1. Copy source to target
    print(f"Copied input file to: {output_vsdx_file}")

    with VisioFile(output_vsdx_file) as vis: # Open the *copied* file
        # 2. Parse and Translate Text
        text_elements = []
        def collect_texts(shape):
            if shape.text and shape.text.strip():
                text_elements.append({
                    "shape": shape,
                    "original": shape.text.strip()
                })

        print("Collecting texts from shapes...")
        process_shapes(vis, collect_texts)
        print(f"Collected {len(text_elements)} text elements for translation.")

        total = len(text_elements)
        if not text_elements:
            print("\n************************************************")
            print("No translatable text found in the Visio file.")
            print("************************************************\n")
            return

        for i in range(0, total, batch_size):
            batch = text_elements[i:i+batch_size]
            texts = [item["original"] for item in batch]

            print(f"\n------------------------------------------------")
            print(f"Translating batch {i//batch_size + 1}/{(total + batch_size - 1)//batch_size}")
            print(f"Batch texts: {texts}")

            try:
                translated = translate_batch(texts, source_lang, target_lang, model_name)
                print(f"Batch translation successful. Translated texts: {translated}")

                for j, text in enumerate(translated):
                    if dual_language:
                        new_text = f"{text}\n{batch[j]['original']}" # Translated above original
                        batch[j]["shape"].text = new_text
                    else:
                        batch[j]["shape"].text = text # Simple text replacement

                    # Attempt to enable AutoSize - simpler approach, if available
                    try:
                        batch[j]["shape"].set_cell_value("AutoSize", 1) # Try simple AutoSize
                        print(f"AutoSize enabled for shape ID: {batch[j]['shape'].ID}")
                    except Exception as autosize_e:
                        print(f"Error setting AutoSize for shape ID: {batch[j]['shape'].ID}: {autosize_e}")


                print(f"Batch {i//batch_size + 1} text replacement in Visio shapes complete.")

            except RuntimeError as e:
                print(f"************************************************")
                print(f"Error during translation of batch {i//batch_size + 1}: {e}")
                print("Skipping batch and proceeding. The translated Visio file will be incomplete.")
                print("************************************************\n")
                continue

        # 3. Save target file
        print(f"Output VSDX filename before save_vsdx: {output_vsdx_file}") # Debug print
        try:
            vis.save_vsdx(output_vsdx_file)
            print(f"Visio file saved with translations to: {output_vsdx_file}")
        except Exception as save_e:
            print(f"Error during vis.save_vsdx: {save_e}")
            vis.save_vsdx() # Fallback to default save

    if output_pdf:
        output_pdf_file = get_output_filename(input_file, output_pdf=True)
        print(f"Attempting to convert '{output_vsdx_file}' to PDF '{output_pdf_file}'")
        convert_to_pdf(output_vsdx_file, output_pdf_file)
    else:
        output_pdf_file = None # Not generated

    print(f"\n************************************************")
    print(f"Successfully translated file saved to: {output_vsdx_file}")
    if output_pdf_file:
        print(f"Successfully converted to PDF: {output_pdf_file}")
    print("************************************************\n")
    print("Exiting translate_visio_file")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Optimized Visio Translator")
    parser.add_argument("input_file", help="Visio file to translate")
    parser.add_argument("--source-lang", default="auto", help="Source language code")
    parser.add_argument("--target-lang", required=True, help="Target language code")
    parser.add_argument("--model", default="gemini-2.0-flash", help="Gemini model")
    parser.add_argument("--batch-size", type=int, default=30, help="Translation batch size") # Reduced default batch size for potentially better error handling
    parser.add_argument("--output-pdf", action="store_true", help="Convert the translated Visio file to PDF using Aspose.Words (if installed)")
    parser.add_argument("--dual-language", action="store_true", help="Output Visio file with both translated and source language text in shapes (translated above source)")


    args = parser.parse_args()

    if not os.getenv("GOOGLE_API_KEY"):
        raise ValueError("GOOGLE_API_KEY environment variable not set")

    print("\n-------------------- START TRANSLATION PROCESS --------------------\n")
    translate_visio_file(
        args.input_file,
        args.source_lang,
        args.target_lang,
        args.model,
        args.batch_size,
        args.output_pdf,
        args.dual_language
    )
    print("\n--------------------- END TRANSLATION PROCESS ---------------------\n")