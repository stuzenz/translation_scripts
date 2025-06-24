import os
from google import genai

def generate_japanese_srt(audio_file_path, output_srt_path):
    """
    Generate Japanese SRT subtitles from an audio file using Gemini 2.5 Flash
    
    Args:
        audio_file_path (str): Path to the input audio file
        output_srt_path (str): Path where the SRT file will be saved
    """
    
    # Initialize the client
    client = genai.Client(api_key=os.getenv('GOOGLE_API_KEY'))
    
    # Upload the audio file
    print("Uploading audio file...")
    myfile = client.files.upload(file=audio_file_path)
    print(f"Uploaded file: {myfile.name}")
    
    # Crafted prompt for high-quality Japanese SRT generation
    prompt = """You're a professional transcriber. You take an audio file and MUST output the transcription in Japanese using kanji where appropriate. You will return an accurate, high-quality SubRip Subtitle (SRT) file.

CRITICAL REQUIREMENTS:
1. You MUST output ONLY the SRT content with no additional text or markdown.
2. Every timestamp MUST be in valid SRT format: 00:00:00,000 --> 00:00:00,000
3. Each segment should be 1-2 lines and maximum 5 seconds duration.
4. Every subtitle entry MUST have:
   - A sequential number starting from 1
   - A timestamp line (start --> end)
   - 1-2 lines of Japanese text with proper kanji usage
   - A blank line between entries

Example format:
1
00:00:00,000 --> 00:00:05,000
こんにちは、元気ですか？

2
00:00:05,000 --> 00:00:10,000
はい、元気です。ありがとうございます。

IMPORTANT: Use natural Japanese with appropriate kanji, hiragana, and katakana. Ensure timestamps are accurate and segments are well-timed for readability."""

    print("Generating Japanese SRT transcription...")
    
    # Generate the SRT content
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents=[prompt, myfile]
    )
    
    # Save the SRT content to file
    if response.text:
        with open(output_srt_path, 'w', encoding='utf-8') as f:
            f.write(response.text.strip())
        print(f"SRT file saved successfully: {output_srt_path}")
        
        # Display first few lines for verification
        lines = response.text.strip().split('\n')[:10]
        print("\nFirst few lines of generated SRT:")
        print('\n'.join(lines))
        
    else:
        print("No response received from Gemini")
    
    return response.text if response.text else None

def validate_srt_format(srt_content):
    """
    Basic validation of SRT format
    
    Args:
        srt_content (str): The SRT file content
        
    Returns:
        bool: True if format appears valid
    """
    lines = srt_content.strip().split('\n')
    
    # Check if it starts with a number
    if not lines[0].isdigit():
        return False
    
    # Check for timestamp format in second line
    if '-->' not in lines[1]:
        return False
        
    print("SRT format validation: PASSED")
    return True

# Main execution
if __name__ == "__main__":
    # Configuration
    input_audio = "../input_files/F1.mp3"  # Your audio file path
    output_srt = "../output_files/F1_source.srt"  # Output SRT file path
    
    # Create output directory if it doesn't exist
    os.makedirs(os.path.dirname(output_srt), exist_ok=True)
    
    try:
        # Generate the SRT file
        srt_content = generate_japanese_srt(input_audio, output_srt)
        
        if srt_content:
            # Validate the format
            if validate_srt_format(srt_content):
                print(f"\n✅ Success! Japanese SRT file created: {output_srt}")
            else:
                print("\n⚠️  Warning: Generated content may not be in proper SRT format")
        
    except Exception as e:
        print(f"Error: {e}")