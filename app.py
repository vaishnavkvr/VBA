from flask import Flask, request, render_template, send_file
from extract_vba import extract_vba_from_excel
from generate_documentation import generate_documentation
from transform_vba_to_python import transform_vba_to_python
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    
    file_path = os.path.join('uploads', file.filename)
    file.save(file_path)
    
    vba_code = extract_vba_from_excel(file_path)
    documentation = generate_documentation(vba_code)
    python_code = transform_vba_to_python(vba_code)
    
    with open("documentation.md", "w") as doc_file:
        doc_file.write(documentation)
    
    with open("transformed_code.py", "w") as code_file:
        code_file.write(python_code)
    
    return render_template('results.html', documentation=documentation, python_code=python_code)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    app.run(debug=True)
