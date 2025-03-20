import os
import tempfile
from flask import Flask, request, render_template, send_file
import pdfplumber
import docx
import re
import openpyxl
from datetime import datetime
from difflib import SequenceMatcher

# Mapping country names to nationalities
country_to_nationality = {
    "United States": "American", "USA": "American", "India": "Indian", "Canada": "Canadian",
    "United Kingdom": "British", "UK": "British", "Australia": "Australian", "Germany": "German",
    "France": "French", "Spain": "Spanish", "Italy": "Italian", "China": "Chinese", "Japan": "Japanese",
    "South Korea": "Korean", "Brazil": "Brazilian", "Mexico": "Mexican", "Russia": "Russian",
    "Netherlands": "Dutch", "Turkey": "Turkish", "Sweden": "Swedish", "Norway": "Norwegian",
    "Denmark": "Danish", "Finland": "Finnish", "Switzerland": "Swiss", "South Africa": "South African",
    "Argentina": "Argentinian", "Egypt": "Egyptian", "Saudi Arabia": "Saudi", 
    "United Arab Emirates": "Emirati", "UAE": "Emirati"
}

app = Flask(__name__)

# Use a temporary directory for file uploads
UPLOAD_FOLDER = tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Words to ignore in filename
IGNORE_WORDS = {"resume", "cv", "profile"}

def clean_filename(file_name):
    """
    Removes ignored words and numbers from the filename.
    Returns the cleaned name if it contains more than 3 letters.
    """
    words = re.split(r'[\s\W_]+', file_name)  # Split by space, special characters, and underscores
    cleaned_words = [word for word in words if word.lower() not in IGNORE_WORDS and not word.isdigit()]
    
    cleaned_name = " ".join(cleaned_words)  # Rejoin the words
    return cleaned_name if len(cleaned_name) > 3 else None  # Return only if length > 3

def name_similarity(extracted_name, file_name):
    """
    Checks the similarity between the extracted name and filename.
    Implements logic for handling similarity scores and ignored words.
    """
    file_name_base = os.path.splitext(file_name)[0]  # Remove file extension
    file_name_cleaned = clean_filename(file_name_base)  # Clean filename

    if not extracted_name:  # If no name extracted, use cleaned filename if available
        return file_name_cleaned if file_name_cleaned else "Unknown"

    similarity_score = SequenceMatcher(None, extracted_name.lower(), file_name_base.lower()).ratio()

    if similarity_score > 0.5:
        return extracted_name
    
    # If similarity is low, check the cleaned file name logic
    if file_name_cleaned:
        return file_name_cleaned  # Use cleaned filename if valid
    
    # If filename is invalid, check if extracted name has numbers
    if not re.search(r'\d', extracted_name):  # If extracted name has no numbers, keep it
        return extracted_name
    
    return "Unknown"  # If both fail, return "Unknown"

# Function to extract email
def extract_email(text):
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    email = re.findall(email_pattern, text)
    return email[0] if email else None

# Function to extract phone number
def extract_phone(text):
    phone_pattern = r'\+?\d{1,3}[-.\s]?\d{1,4}[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{2,4}'
    phone_matches = re.findall(phone_pattern, text)
    
    if phone_matches:
        clean_phone = re.sub(r'\D', '', phone_matches[0])  # Remove non-digit characters
        if len(clean_phone) > 14:
            return "Not Found"
        return phone_matches[0]
    
    return None

# Function to extract name (assuming it is in the first non-empty line)
def extract_name(text):
    lines = text.splitlines()
    for line in lines:
        if line.strip():
            return line.strip()
    return None

# Function to extract nationality based on country mention
def extract_nationality(text):
    found_countries = []
    for country, nationality in country_to_nationality.items():
        if re.search(rf'\b{country}\b', text, re.IGNORECASE):
            found_countries.append(nationality)

    nationality_match = re.search(r'Nationality[:\-]?\s*(\w+)', text, re.IGNORECASE)
    if nationality_match:
        return nationality_match.group(1).capitalize()

    if len(found_countries) > 1:
        return ", ".join(set(found_countries))

    return found_countries[0] if found_countries else "Not Found"

# Function to read PDF file
def read_pdf(file_path):
    with pdfplumber.open(file_path) as pdf:
        text = ''.join(page.extract_text() for page in pdf.pages if page.extract_text())
    return text

# Function to read DOCX file
def read_docx(file_path):
    doc = docx.Document(file_path)
    return '\n'.join(para.text for para in doc.paragraphs)

# Function to process a single resume and extract data
def process_resume(file_path):
    file_name = os.path.basename(file_path)  # Get filename with extension

    # Read the file and extract text
    text = ''
    if file_path.endswith('.pdf'):
        text = read_pdf(file_path)
    elif file_path.endswith('.docx'):
        text = read_docx(file_path)
    else:
        print(f"Unsupported file type: {file_path}")
        return []

    # Extract details
    extracted_name = extract_name(text)
    email = extract_email(text)
    phone = extract_phone(text)
    nationality = extract_nationality(text)

    # Determine final name based on similarity logic
    final_name = name_similarity(extracted_name, file_name)

    # Return extracted data
    return [(final_name, email, phone, nationality)]

# Function to create and save data into an Excel file
def create_excel(data, output_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resume Data"

    headers = ["Name", "Email", "Phone Number", "Nationality"]
    ws.append(headers)
    for row in data:
        ws.append(row)

    wb.save(output_file)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        uploaded_files = request.files.getlist('file')
        all_resume_data = []

        for file in uploaded_files:
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)
            
            # Process file and extract data
            resume_data = process_resume(file_path)
            all_resume_data.extend(resume_data)

        # Generate unique output filename
        current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_file = os.path.join(app.config['UPLOAD_FOLDER'], f'resumedata_{current_datetime}.xlsx')

        # Save extracted data into Excel
        create_excel(all_resume_data, output_file)

        return send_file(output_file, as_attachment=True)

    return render_template('index.html')

if __name__ == "__main__":
    port = os.getenv("PORT", 5000)  # Use Render's port or default to 5000
    app.run(host="0.0.0.0", port=int(port), debug=True)
