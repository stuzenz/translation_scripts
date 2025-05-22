import asyncio
import google.generativeai as genai
import os
import json

# Configure API
api_key = os.environ.get("GOOGLE_API_KEY")
if not api_key:
    print("GOOGLE_API_KEY environment variable not set.")
    exit(1)
    
genai.configure(api_key=api_key)
print(f"Google Generative AI SDK version: {genai.__version__}")

available_models = genai.list_models()
print(f"Available models: {[model.name for model in available_models]}")

async def test_translation():
    model_name = "gemini-1.5-pro-latest"  # Try different models
    model = genai.GenerativeModel(model_name)
    
    print(f"Testing translation with model: {model_name}")
    
    # Simple text to translate
    test_text = "こんにちは、世界！これは翻訳テストです。"
    
    # Create prompt
    prompt = f"""
    Translate the following Japanese text to English:
    
    {test_text}
    
    Return ONLY the English translation.
    """
    
    try:
        # Basic API call with no extra parameters
        response = await model.generate_content_async(prompt)
        print(f"Translation result: {response.text}")
        print("Basic translation test succeeded!")
        
        # Now test with JSON output
        test_texts = ["こんにちは、世界！", "私の名前はボブです。", "今日は良い天気ですね。"]
        
        json_prompt = f"""
        Translate these Japanese snippets to English:
        
        {json.dumps(test_texts, ensure_ascii=False)}
        
        Return ONLY a valid JSON object mapping each Japanese text to its English translation.
        """
        
        json_response = await model.generate_content_async(json_prompt)
        print(f"JSON response: {json_response.text}")
        
        # Try to parse the JSON
        try:
            parsed_json = json.loads(json_response.text)
            print("Successfully parsed JSON response!")
            print(f"Parsed result: {parsed_json}")
            return True
        except json.JSONDecodeError:
            print("JSON parsing failed. Raw response:")
            print(json_response.text)
            return False
        
    except Exception as e:
        print(f"Error during translation test: {type(e).__name__} - {e}")
        return False

# Run the async test
async def main():
    await test_translation()

if __name__ == "__main__":
    asyncio.run(main())
