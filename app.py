from flask import Flask, render_template, request, send_from_directory
import os
from werkzeug.utils import secure_filename
from exam_analysis import process_and_analyze
import pandas as pd
from fpdf import FPDF

app = Flask(__name__)

# Define folders
BASE_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "ExamAnalysis")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "Uploads")
PROCESSED_FOLDER = os.path.join(BASE_DIR, "Reports")

# Ensure required directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    report_files = []
    tables = {}
    
    if request.method == 'POST':
        files = request.files.getlist('files')
        exam_type = request.form.get('exam_type')

        for file in files:
            if file.filename.endswith('.xlsx'):
                filename = secure_filename(file.filename)
                file_path = os.path.join(UPLOAD_FOLDER, filename)
                file.save(file_path)
                
                # Process and analyze the file
                report_path = process_and_analyze(file_path, exam_type, PROCESSED_FOLDER)
                report_files.append(os.path.basename(report_path))
    
    # Load only analyzed reports (ignore cleaned files)
    for report in os.listdir(PROCESSED_FOLDER):
        if report.endswith('.xlsx') and not report.endswith('_cleaned.xlsx'):
            report_files.append(report)
            tables[report] = pd.read_excel(os.path.join(PROCESSED_FOLDER, report), sheet_name=None)
    
    return render_template('index.html', tables=tables, report_files=report_files if report_files else None)

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(PROCESSED_FOLDER, filename, as_attachment=True)

@app.route('/export_pdf/<filename>')
def export_pdf(filename):
    file_path = os.path.join(PROCESSED_FOLDER, filename)
    df = pd.read_excel(file_path, sheet_name=None)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    for sheet, data in df.items():
        pdf.add_page()
        pdf.set_font("Arial", style='B', size=16)
        pdf.cell(200, 10, sheet, ln=True, align='C')
        pdf.set_font("Arial", size=10)
        
        for row in data.itertuples(index=False):
            pdf.cell(0, 10, " | ".join(str(x) for x in row), ln=True)

    pdf_path = os.path.join(PROCESSED_FOLDER, filename.replace('.xlsx', '.pdf'))
    pdf.output(pdf_path)
    return send_from_directory(PROCESSED_FOLDER, os.path.basename(pdf_path), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
