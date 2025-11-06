from flask import Flask, render_template, request, send_file, redirect, url_for
import os
from werkzeug.utils import secure_filename
import pandas as pd

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['REPORT_FOLDER'] = 'reports'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORT_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files['file']
    if file and file.filename.endswith('.xlsx'):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        # Placeholder for analysis logic
        df = pd.read_excel(filepath)
        summary = df.describe().to_string()
        report_path = os.path.join(app.config['REPORT_FOLDER'], filename.replace('.xlsx', '.txt'))
        with open(report_path, 'w') as f:
            f.write("Simple Audit Summary:\n\n")
            f.write(summary)
        return send_file(report_path, as_attachment=True)
    return "Invalid file type. Please upload an Excel file."

@app.route('/reports')
def reports():
    files = os.listdir(app.config['REPORT_FOLDER'])
    return render_template('index.html', files=files)

if __name__ == "__main__":
    app.run(debug=True)


