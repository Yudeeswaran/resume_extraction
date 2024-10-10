import PyPDF2
from docx import Document
from docx.shared import Inches

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
    lines = raw_text.split('\n')
    
    # Initialize structured output
    resume_data = {
        "Skills": "",
        "Interests": "",
        "Education": "",
        "Industrial Project": "",
        "Achievements": "",
        "Organization": "",
        "Certificates": "",
        "Projects": ""
    }
    
    current_section = None

    for line in lines:
        if line.strip() == "":
            continue

        line_lower = line.lower()  # Convert line to lower case for comparison

        # Check for section headers
        if 'skills' in line_lower:
            current_section = 'Skills'
            continue
        elif 'interests' in line_lower:
            current_section = 'Interests'
            continue
        elif 'education' in line_lower:
            current_section = 'Education'
            continue
        elif 'industrial project' in line_lower:
            current_section = 'Industrial Project'
            continue
        elif 'achievements' in line_lower:
            current_section = 'Achievements'
            continue
        elif 'organization' in line_lower:
            current_section = 'Organization'
            continue
        elif 'certificates' in line_lower:
            current_section = 'Certificates'
            continue
        elif 'projects' in line_lower:
            current_section = 'Projects'
            continue

        # Append the line to the appropriate section content
        if current_section:
            resume_data[current_section] += line.strip() + "\n"
    
    return resume_data

# Function to insert a table after a specific paragraph without adding headers again
def insert_table_after_paragraph(doc, para, left_content, right_content):
    table = doc.add_table(rows=1, cols=2)  # Create the table after the paragraph
    table.autofit = False
    table.columns[0].width = Inches(3)  # Left side (half page)
    table.columns[1].width = Inches(3)  # Right side (half page)

    # Add left content (no header)
    table.cell(0, 0).text = left_content

    # Add right content (no header)
    table.cell(0, 1).text = right_content

    # Move the table after the paragraph
    para._element.addnext(table._element)

# Function to fill the Word template
def fill_docx_template(template_path, output_path, resume_data):
    doc = Document(template_path)

    # Traverse the paragraphs in the document
    for para in doc.paragraphs:
        found_headers = [key for key in resume_data if f'{{{key}}}' in para.text]
        
        if len(found_headers) == 2:  # If two placeholders are found on the same line
            left_header, right_header = found_headers
            left_content = resume_data[left_header].strip()
            right_content = resume_data[right_header].strip()

            # Clear the placeholders, but do not repeat the headers
            para.clear()  # Clear the paragraph content but keep the headers intact
            insert_table_after_paragraph(doc, para, left_content, right_content)  # Insert content

        # For single header, replace the placeholder with content
        for key, value in resume_data.items():
            placeholder = '{' + key + '}'  # For example, {Skills}
            if placeholder in para.text:
                # Replace only the placeholder, leave the header intact
                para.text = para.text.replace(placeholder, value.strip())

    # Save the updated document
    doc.save(output_path)

# Function to check and validate the structure of the template
def validate_template(template_path, resume_data):
    doc = Document(template_path)

    missing_placeholders = []
    # Ensure all required placeholders are present in the document
    for key in resume_data.keys():
        found = False
        for para in doc.paragraphs:
            placeholder = '{' + key + '}'
            if placeholder in para.text:
                found = True
                break
        if not found:
            missing_placeholders.append(key)

    if missing_placeholders:
        print(f"Missing placeholders: {missing_placeholders}")
    else:
        print("All placeholders are correctly placed in the template.")

# Main part: extracting and populating resume
pdf_path = 'Final.pdf'  # Replace with your PDF file path
template_path = 'template.docx'  # Path to your Word template file
output_path = 'populated_resume.docx'  # Define the output Word file path

# Extract the text from the PDF resume
resume_text = extract_text_from_pdf(pdf_path)

# Format the extracted resume text
formatted_resume = format_resume_text(resume_text)

# Check and validate the template to ensure placeholders are correct
validate_template(template_path, formatted_resume)

# Fill the Word template with the formatted resume data
fill_docx_template(template_path, output_path, formatted_resume)

print(f"\nPopulated resume has been saved to {output_path}.")
