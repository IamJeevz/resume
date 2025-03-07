import os
import re
import pdfplumber
import docx
import openpyxl
from flask import Flask, request, jsonify, send_file, render_template
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configure the file upload folder
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Function to extract email
def extract_email(text):
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    email = re.findall(email_pattern, text)
    return email[0] if email else None

# Function to extract phone number
def extract_phone(text):
    phone_pattern = r'\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}|\d{10,15}'
    phone = re.findall(phone_pattern, text)
    return phone[0] if phone else None

# Function to extract name
def extract_name(text):
    lines = text.splitlines()
    for line in lines:
        if line.strip():  # Check if the line is not empty
            return line.strip()
    return None

# Function to read PDF file
def read_pdf(file_path):
    with pdfplumber.open(file_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    return text

# Function to read DOCX file
def read_docx(file_path):
    doc = docx.Document(file_path)
    text = ''
    for para in doc.paragraphs:
        text += para.text
    return text

# Function to process resumes and extract data
def process_resume(file_path):
    if file_path.endswith('.pdf'):
        text = read_pdf(file_path)
    elif file_path.endswith('.docx'):
        text = read_docx(file_path)
    else:
        return None

    name = extract_name(text)
    email = extract_email(text)
    phone = extract_phone(text)

    return [name, email, phone]

# Function to create and save data into an Excel file
def create_excel(data, output_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resume Data"
    headers = ["Name", "Email", "Phone Number"]
    ws.append(headers)
    for row in data:
        ws.append(row)

    wb.save(output_file)

# Flask route to upload file(s) and process resume
@app.route('/upload', methods=['POST'])
def upload_file():
    files = request.files.getlist('file')
    
    if not files:
        return jsonify({"error": "No files selected"}), 400

    resume_data = []

    # Process each file
    for file in files:
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        data = process_resume(file_path)
        
        if data:
            resume_data.append(data)
        else:
            return jsonify({"error": f"Unable to process the file {filename}"}), 400

    output_file = 'resume_data.xlsx'
    create_excel(resume_data, output_file)
    return send_file(output_file, as_attachment=True)

# Flask route to serve the front-end page
@app.route('/')
def index():
    return render_template('index.html')

if __name__ == "__main__":
    # Make sure to create an 'uploads' folder
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    
    app.run(debug=True)
