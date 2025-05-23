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


# Define a list of common job title keywords
job_keywords = [
    "Manager", "Engineer", "Developer", "Doctor", "Consultant", "Coordinator",
    "Specialist", "Analyst", "Nurse", "Architect", "Technician", "Lead", "Director", "Executive", "Trainer", "Scientist",
    "Assistant", "Supervisor", "Administrator", "Clerk", "Operator", "Officer", "Designer", "Trainer", "Technologist",
    "Chef", "Sales", "Accountant", "Business Analyst", "Project Manager", "Program Manager", "Product Manager", "Legal Advisor",
    "Social Worker", "Researcher", "Marketing", "HR", "Director", "Chief", "Chief Executive Officer", "CFO", "COO", "CTO",
    "Software Engineer", "Web Developer", "Data Scientist", "System Analyst", "IT Manager", "Business Development", "Chief Marketing Officer",
    "UX Designer", "Product Designer", "Data Analyst", "Business Development Manager", "Digital Marketing", "Account Executive",
    "Financial Analyst", "Security Specialist", "HR Manager", "Operations Manager", "Quality Analyst", "Risk Manager", "IT Specialist",
    "Sales Manager", "Customer Support", "Logistics Manager", "Project Coordinator", "Public Relations", "Copywriter", "Content Writer",
    "Photographer", "Videographer", "Consulting Analyst", "Security Consultant", "Healthcare Consultant", "Marketing Consultant",
    "SEO Specialist", "UX/UI Designer", "Event Coordinator", "Facilities Manager", "Office Manager", "Customer Service Representative",
    "Research Analyst", "Teacher", "Instructor", "Professor", "Lecturer", "Academic Advisor", "Instructional Designer", "Counselor",
    "Chief Information Officer", "Software Developer", "Field Engineer", "Maintenance Engineer", "Systems Administrator", "Network Engineer",
    "Recruiter", "Event Planner", "Data Entry", "Technician", "Help Desk", "Support Engineer", "Financial Controller", "Health Educator",
    "Project Director", "Creative Director", "Brand Manager", "Talent Manager", "Business Partner", "Product Specialist", "SEO Manager"
]





app = Flask(__name__)

# Use a temporary directory for file uploads
UPLOAD_FOLDER = tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Words to ignore in filename
IGNORE_WORDS = {'resume', 'cv', 'curriculum', 'vitae', 'application', 'letter'}  # Add more as needed



def clean_filename(file_name):
    """
    Removes ignored words and numbers from the filename.
    Returns the cleaned name if it contains more than 3 letters.
    """
    words = re.split(r'[\s\W_]+', file_name)  # Split by space, special characters, and underscores
    cleaned_words = [word for word in words if word.lower() not in IGNORE_WORDS and not word.isdigit()]
    cleaned_name = " ".join(cleaned_words)  # Rejoin the words
    return cleaned_name if len(cleaned_name) > 3 else None

def name_similarity(extracted_name, file_name):
    """
    Compares the extracted name with the filename based on multiple conditions.
    If the similarity score is greater than 0.5, return extracted_name.
    If filename contains extracted_name or vice versa, return extracted_name.
    If the extracted name contains any numbers, return the file name.
    If the file name appears in the entire file content, return the filename.
    Otherwise, return "Not Found".
    """
    if not extracted_name and not file_name:
        return "Not Found"  # Return "Not Found" if both name and file name are missing

    file_name_base = os.path.splitext(file_name)[0]  # Remove file extension
    file_name_cleaned = clean_filename(file_name_base)  # Clean filename

    # 1. If extracted_name contains any number, return file_name
    if re.search(r'\d', extracted_name):  # If extracted_name contains a digit
        return file_name_cleaned

    # 2. Compare extracted_name and file_name similarity score
    similarity_score = SequenceMatcher(None, extracted_name.lower(), file_name_base.lower()).ratio()
    if similarity_score > 0.5:
        return extracted_name

    # 3. Check if file_name contains extracted_name or extracted_name contains file_name
    if extracted_name.lower() in file_name_base.lower() or file_name_base.lower() in extracted_name.lower():
        return extracted_name

    # 4. Check if the entire file contains the file_name (use the cleaned version for comparison)
    if file_name_cleaned and file_name_cleaned in extracted_name.lower():
        return file_name_cleaned

    # 5. If none of the above, return extracted_name (default behavior)
    return extracted_name

# Function to extract email
def extract_email(text):
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    email = re.findall(email_pattern, text)
    return email[0] if email else "Not Found"

# Function to extract phone number
def extract_phone(text):
    phone_pattern = r'\+?\d{1,3}[-.\s]?\d{1,4}[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{2,4}'
    phone_matches = re.findall(phone_pattern, text)
    if phone_matches:
        clean_phone = re.sub(r'\D', '', phone_matches[0])  # Remove non-digit characters
        if len(clean_phone) > 14:
            return "Not Found"
        return phone_matches[0]
    return "Not Found"

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

def extract_designation_simple(text):
    """
    Extracts job titles/designations from the given text using predefined job-related keywords.
    """
    # Use regex to find occurrences of job keywords in the text
    pattern = r'\b(?:' + '|'.join(job_keywords) + r')\b'
    job_titles = re.findall(pattern, text, re.IGNORECASE)
    
    # Remove duplicates by converting to a set and back to a list
    job_titles = list(set([title.capitalize() for title in job_titles]))
    
    # If no job title is found, return "Not Found"
    if job_titles:
        return ", ".join(job_titles)
    return "Not Found"

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
    designation = extract_designation_simple(text)  # Extract the designation/job title

    # Determine final name based on similarity logic
    final_name = name_similarity(extracted_name, file_name)

    # Return extracted data
    return [(final_name, email, phone, nationality, designation)]

# Function to create and save data into an Excel file
def create_excel(data, output_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resume Data"

    headers = ["Name", "Email", "Phone Number", "Nationality", "Designation"]
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
    app.run(host="0.0.0.0", port=int(port), debug=True)  # Start the Flask app
