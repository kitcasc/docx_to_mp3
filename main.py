from gtts import gTTS
from docx import Document

# Read Word file and extract text
def read_word_file(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

# Convert text to speech and save as MP3
def text_to_speech(text, output_file):
    tts = gTTS(text)
    tts.save(output_file)

# main function
def convert_word_to_audio(word_file, audio_file):
    text = read_word_file(word_file)  # Read the Word file
    text_to_speech(text, audio_file)  # Convert text to speech and save as MP3

# Example usage
# word_file = r'C:\Users\USER\Desktop\Speech draft.docx'  # 
word_file = input('Enter the directory of the Word file: ')  # Prompt user to input the file path
audio_file = input('Enter file name for output(ex. audio.mp3): ')  # Output the file name which is input by user

convert_word_to_audio(word_file, audio_file)

print(f"Text successfully converted to audio: {audio_file}")