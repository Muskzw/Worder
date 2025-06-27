import uuid
from flask import Flask, request, send_file, render_template_string, redirect, url_for, flash
from pdf2docx import Converter
import os
import pytesseract
from PIL import Image
from docx import Document
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB limit
app.secret_key = 'your_secret_key'  # Needed for flashing messages

HTML_FORM = '''
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>File to Word Converter</title>
</head>
<body>
  <h2>Convert File to Word (.docx)</h2>
  {% with messages = get_flashed_messages() %}
    {% if messages %}
      <ul>
      {% for message in messages %}
        <li style="color:red;">{{ message }}</li>
      {% endfor %}
      </ul>
    {% endif %}
  {% endwith %}
  <form method="post" action="/convert" enctype="multipart/form-data">
    <input type="file" name="file" required><br><br>
    <select name="lang">
      <option value="eng">English</option>
      <option value="spa">Spanish</option>
    </select><br><br>
    <button type="submit">Convert</button>
  </form>
</body>
</html>
'''

@app.route('/')
def home():
    return render_template_string(HTML_FORM)

@app.route('/convert', methods=['POST'])
def convert():
    uploaded_file = request.files['file']
    if not uploaded_file:
        flash("No file uploaded")
        return redirect(url_for('home'))

    filename = secure_filename(uploaded_file.filename)
    if not filename.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg', '.txt')):
        flash("Unsupported file type")
        return redirect(url_for('home'))

    unique_id = str(uuid.uuid4())
    filepath = os.path.join(UPLOAD_FOLDER, unique_id + "_" + filename)
    uploaded_file.save(filepath)

    output_filename = unique_id + "_converted.docx"
    output_path = os.path.join(UPLOAD_FOLDER, output_filename)

    try:
        if filename.endswith('.pdf'):
            cv = Converter(filepath)
            cv.convert(output_path, start=0, end=None)
            cv.close()
        elif filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            lang = request.form.get('lang', 'eng')
            text = pytesseract.image_to_string(Image.open(filepath), lang=lang)
            doc = Document()
            doc.add_paragraph(text)
            doc.save(output_path)
        elif filename.endswith('.txt'):
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            doc = Document()
            doc.add_paragraph(content)
            doc.save(output_path)
    except Exception as e:
        flash(f"Conversion failed: {e}")
        if os.path.exists(filepath):
            os.remove(filepath)
        return redirect(url_for('home'))

    # Remove the uploaded file after conversion
    if os.path.exists(filepath):
        os.remove(filepath)

    # Show a download link instead of immediate download
    return render_template_string('''
        <h3>Conversion successful!</h3>
        <a href="{{ url_for('download', filename=filename) }}">Download your Word file</a>
    ''', filename=output_filename)

@app.route('/download/<filename>')
def download(filename):
    output_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(output_path):
        return "File not found", 404
    response = send_file(output_path, as_attachment=True)
    # Remove the file after sending
    try:
        os.remove(output_path)
    except Exception:
        pass
    return response

if __name__ == '__main__':
    app.run(debug=True)
