from flask import Flask, request, send_file, render_template_string
from pdf2docx import Converter
import os
import pytesseract
from PIL import Image
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML_FORM = '''
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>File to Word Converter</title>
</head>
<body>
  <h2>Convert File to Word (.docx)</h2>
  <form method="post" action="/convert" enctype="multipart/form-data">
    <input type="file" name="file" required><br><br>
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
        return "No file uploaded", 400

    filename = uploaded_file.filename
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    uploaded_file.save(filepath)

    output_path = os.path.join(UPLOAD_FOLDER, 'converted.docx')

    if filename.endswith('.pdf'):
        cv = Converter(filepath)
        cv.convert(output_path, start=0, end=None)
        cv.close()
    elif filename.lower().endswith(('.png', '.jpg', '.jpeg')):
        text = pytesseract.image_to_string(Image.open(filepath))
        doc = Document()
        doc.add_paragraph(text)
        doc.save(output_path)
    elif filename.endswith('.txt'):
        with open(filepath, 'r') as f:
            content = f.read()
        doc = Document()
        doc.add_paragraph(content)
        doc.save(output_path)
    else:
        return "Unsupported file type", 415

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
