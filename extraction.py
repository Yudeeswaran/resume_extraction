import PyPDF2
import pandas as pd

# Function to extract text from a PDF
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text() or ""  # Ensure we avoid None type if no text is extracted
    return text

# Function to format the extracted text into sections
def format_resume_text(raw_text):
    # Split the text into lines for easier manipulation
    lines = raw_text.split('\n')
    
    # Initialize structured output
    resume_data = {
        "Section": [],
        "Content": []
    }
    
    # Initialize flags for different sections
    current_section = None

    for line in lines:
        # Clean up empty lines
        if line.strip() == "":
            continue

        # Detect common sections by keywords and format them
        line_lower = line.lower()  # Convert line to lower case for comparison

        # Check for section headers
        if 'skills' in line_lower and current_section != 'Skills':
            current_section = 'Skills'
            resume_data["Section"].append(current_section)
            resume_data["Content"].append("")  # Prepare to collect content
            continue
        elif 'interests' in line_lower and current_section != 'Interests':
            current_section = 'Interests'
            resume_data["Section"].append(current_section)
            resume_data["Content"].append("")  # Prepare to collect content
            continue
        elif 'education' in line_lower and current_section != 'Education':
            current_section = 'Education'
            resume_data["Section"].append(current_section)
            resume_data["Content"].append("")  # Prepare to collect content
            continue
        elif 'industrial project' in line_lower and current_section != 'Industrial Project':
            current_section = 'Industrial Project'
            resume_data["Section"].append(current_section)
            resume_data["Content"].append("")  # Prepare to collect content
            continue
        elif 'achievements' in line_lower and current_section != 'Achievements':
            current_section = 'Achievements'
            resume_data["Section"].append(current_section)
            resume_data["Content"].append("")  # Prepare to collect content
            continue
        elif 'organization' in line_lower and current_section != 'Organization':
            current_section = 'Organization'
            resume_data["Section"].append(current_section)
            resume_data["Content"].append("")  # Prepare to collect content
            continue
        elif 'certificates' in line_lower and current_section != 'Certificates':
            current_section = 'Certificates'
            resume_data["Section"].append(current_section)
            resume_data["Content"].append("")  # Prepare to collect content
            continue
        elif 'projects' in line_lower and current_section != 'Projects':
            current_section = 'Projects'
            resume_data["Section"].append(current_section)
            resume_data["Content"].append("")  # Prepare to collect content
            continue
        
        # Append the line to the appropriate section content
        if current_section:
            resume_data["Content"][-1] += line.strip() + "\n"
    
    return resume_data

# Function to save the formatted resume data to an Excel file
def save_to_excel(resume_data, output_path):
    df = pd.DataFrame(resume_data)
    df.to_excel(output_path, index=False, engine='openpyxl')

# Main part: extracting and formatting resume
pdf_path = 'Final.pdf'  # Replace with your PDF file path
output_path = 'formatted_resume.xlsx'  # Define the output Excel file path

# Extract the text from the PDF resume
resume_text = extract_text_from_pdf(pdf_path)

# Format the extracted resume text
formatted_resume = format_resume_text(resume_text)

# Save the formatted resume to an Excel file
save_to_excel(formatted_resume, output_path)

print(f"\nFormatted resume has been saved to {output_path}.")
