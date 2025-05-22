How to Use:

    Save: Save the code above as a Python file (e.g., excel_translator_gemini_batch.py).

    API Key: Make sure your Google Gemini API key is set as an environment variable GOOGLE_API_KEY or pass it using the --api-key argument.

          
    export GOOGLE_API_KEY="YOUR_API_KEY_HERE"
    # or use --api-key YOUR_API_KEY_HERE in the command

        

    IGNORE_WHEN_COPYING_START

Use code with caution.Bash
IGNORE_WHEN_COPYING_END

Run from Command Line:

```bash 
python excel_translator_gemini_10.py \
    --source-location ./path/to/your/excel/files \
    --target-langs ja ko fr \
    --output-location ./path/to/save/translations \
    --ignore-font "Times New Roman" Arial \
    --concurrency 8 \
    --model gemini-1.5-pro # Optional: specify a different model
    # --api-key YOUR_API_KEY_HERE # Optional: if not using env var
```
    

IGNORE_WHEN_COPYING_START

    Use code with caution.Bash
    IGNORE_WHEN_COPYING_END

Explanation of Arguments:

    --source-location: The folder containing the .xlsx and .xlsm files you want to translate. It will search inside subfolders too.

    --target-langs: A list of language codes (like ja, ko, es, de, en) separated by spaces. The script will create a version of each input file for each of these languages.

    --output-location: The main folder where the translated files will be saved. The script will create subfolders inside this location to match the structure found in the --source-location.

    --ignore-font (Optional): A list of font names (e.g., Arial, "Times New Roman", Calibri). Any text formatted only with these fonts will not be sent for translation and will keep its original value. Font names are case-insensitive. If a font name has spaces, enclose it in quotes.

    --concurrency (Optional): How many translation tasks (file + language combination) to run at the same time. Default is 4. Increase this based on your system's capability and API rate limits.

    --model (Optional): The specific Gemini model to use (default is gemini-1.5-flash).

    --batch-size (Optional): How many cells are sent to the API in a single request (default is 15).

    --api-key (Optional): Provide your API key directly if you prefer not to use environment variables.