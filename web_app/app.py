from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
from gtts import gTTS
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'docx'}

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_word_file(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def text_to_speech(text, output_file):
    tts = gTTS(text)
    tts.save(output_file)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('index.html', error='No file part')
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', error='No selected file')
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Process the file
            text = read_word_file(filepath)
            output_filename = filename.rsplit('.', 1)[0] + '.mp3'
            output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
            text_to_speech(text, output_path)
            
            return render_template('index.html', download_link=output_filename)
        else:
            return render_template('index.html', error='File type not allowed. Please upload a .docx file')
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    response = send_file(file_path, as_attachment=True, download_name=filename, mimetype='audio/mpeg')
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    from urllib.parse import quote
    response.headers['Content-Disposition'] = f'attachment; filename="{quote(filename)}"'
    return response

if __name__ == '__main__':
    app.run(debug=True, port=5000)
    print("Flask server started on http://localhost:5000")
