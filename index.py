import pdfplumber
import docx
import re
import openpyxl
import os

# Mapping country names to nationalities
country_to_nationality = {
    "United States": "American",
    "USA": "American",
    "India": "Indian",
    "Canada": "Canadian",
    "United Kingdom": "British",
    "UK": "British",
    "Australia": "Australian",
    "Germany": "German",
    "France": "French",
    "Spain": "Spanish",
    "Italy": "Italian",
    "China": "Chinese",
    "Japan": "Japanese",
    "South Korea": "Korean",
    "Brazil": "Brazilian",
    "Mexico": "Mexican",
    "Russia": "Russian",
    "Netherlands": "Dutch",
    "Turkey": "Turkish",
    "Sweden": "Swedish",
    "Norway": "Norwegian",
    "Denmark": "Danish",
    "Finland": "Finnish",
    "Switzerland": "Swiss",
    "South Africa": "South African",
    "Argentina": "Argentinian",
    "Egypt": "Egyptian",
    "Saudi Arabia": "Saudi",
    "United Arab Emirates": "Emirati",
    "UAE": "Emirati",
}

# Function to extract email
def extract_email(text):
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    emails = re.findall(email_pattern, text)
    
    valid_emails = [email for email in emails if "." in email.split("@")[-1]]
    
    return valid_emails[0] if valid_emails else "Not Found"

# Function to extract nationality based on country name
def extract_nationality(text):
    for country, nationality in country_to_nationality.items():
        if re.search(rf'\b{country}\b', text, re.IGNORECASE):  # Ensures whole-word match
            return nationality
    return "Not Found"

# Function to read PDF file
def read_pdf(file_path):
    text = ''
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text += page.extract_text() or ""  # Handle None values
    except Exception as e:
        print(f"⚠️ Error reading PDF: {file_path} → {e}")
    return text

# Function to read DOCX file
def read_docx(file_path):
    text = ''
    try:
        doc = docx.Document(file_path)
        for para in doc.paragraphs:
            text += para.text + " "
    except Exception as e:
        print(f"⚠️ Error reading DOCX: {file_path} → {e}")
    return text

# Function to process resumes and extract data
def process_resumes(folder_path):
    resume_data = []
    
    if not os.path.exists(folder_path):
        print(f"❌ Folder not found: {folder_path}")
        return []
    
    file_count = 0
    
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)

        if file_name.endswith('.pdf'):
            text = read_pdf(file_path)
        elif file_name.endswith('.docx'):
            text = read_docx(file_path)
        else:
            print(f"⚠️ Skipping unsupported file: {file_name}")
            continue

        if not text.strip():  # Skip empty files
            print(f"⚠️ Empty or unreadable file: {file_name}")
            continue

        # Extract email and nationality
        email = extract_email(text)
        nationality = extract_nationality(text)

        resume_data.append([file_name, email, nationality])
        file_count += 1

    print(f"✅ Processed {file_count} resumes.")
    return resume_data

# Function to create and save data into an Excel file
def create_excel(data, output_file):
    if not data:
        print("⚠️ No data to save. Excel file not created.")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resume Data"

    headers = ["Filename", "Email", "Nationality"]
    ws.append(headers)
    for row in data:
        ws.append(row)

    # Auto-adjust column width
    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col) + 2
        ws.column_dimensions[col[0].column_letter].width = max_length

    wb.save(output_file)
    print(f"✅ Data saved successfully to {output_file}")

# Main function
def main():
    folder_path = r"D:\Python Projects\resume_full\cv"
    output_file = "resume_data.xlsx"

    resume_data = process_resumes(folder_path)
    create_excel(resume_data, output_file)

if __name__ == "__main__":
    main()
