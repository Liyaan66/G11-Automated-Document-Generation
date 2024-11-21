import os
import openpyxl
from docxtpl import DocxTemplate
from docx2pdf import convert
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

def generate_documents(excel_file, template_file, output_folder):

     # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")   
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    data = list(sheet.values) 

    #folder to store file PDF
    pdf_folder = os.path.join(output_folder, "pdfs")
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)
        print(f"Created PDF folder: {pdf_folder}")
    
    generated_files = []

    #enter data from excel to word file
    student_template = DocxTemplate(template_file)
    for student in data[1:]:
        student_template.render({
            "id_kh": student[0],
            "id_e": student[1],
            "name_kh": student[2],
            "name_e": student[3],
            "g1": student[4],
            "g2": student[5],
            "dob_kh": student[6],
            "dob_e": student[7],
            "pro_kh": student[8],
            "pro_e": student[9],
            "ed_kh": student[10],
            "ed_e": student[11]
        })

        #save the word document
        docx_filename = os.path.join(output_folder, f"{student[0]}.docx")
        student_template.save(docx_filename)
        print(f"Generated document: {docx_filename}")
        generated_files.append(docx_filename)
        
        # Convert the Word document to PDF
        pdf_filename = os.path.join(pdf_folder, f"{student[0]}.pdf")
        convert(docx_filename, pdf_filename)
        print(f"Converted to PDF: {pdf_filename}")
        generated_files.append(pdf_filename)

generate_documents("associate_degree.xlsx", "WEP_temporary_ template.docx", "output_documents")

    sheet = workbook.active 
    student_data = list(sheet.values)[1:]  
    student_info = DocxTemplate(doc_template)
    for student in student_data:
        student_info.render({
            "student_data": student[0] 
        })      
        student_name = student[0]  
        my_file = f"{student_name}.docx"
        student_info.save(my_file)
    print("Documents generated successfully.")
generate_student_docs("students_data.xlsx", "Doc1.docx")


#Generation Certificate---------------------------------

# import pandas as pd
# import os
# from PIL import Image, ImageDraw, ImageFont

# Input data and template file 
excel_file = "students_data.xlsx"
template = "template.png"

# Output folder
output_folder = "generated_certificate"

# Load the Excel file
data = pd.read_excel(excel_file)

# Create the output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Font settings
bold_font = "calibrib.ttf"
font_name = ImageFont.truetype(bold_font, 100)

# Generate certificates
for index, row in data.iterrows():
    name = row["Name"]
    certificate = Image.open(template)
    draw = ImageDraw.Draw(certificate)
    
    # Measure the width and height of the name text
    text_bbox = draw.textbbox((0, 0), name, font=font_name)  # Get bounding box for text
    text_width = text_bbox[2] - text_bbox[0]  # Calculate text width
    text_height = text_bbox[3] - text_bbox[1]  # Calculate text height
    
    # Calculate the horizontal position for centering the text
    image_width = certificate.width
    x_position = (image_width - text_width) //2 # Center horizontally
    
    # Fixed vertical position
    y_position = 625  # Adjust as needed
    # Draw the name on the certificate
    draw.text((x_position, y_position), name, fill="orange", font=font_name)

    # Save the certificate to the output folder
    output_path = os.path.join(output_folder, f"{name}.png")
    certificate.save(output_path)

    # Print confirmation
    print(f"Certificate generated for {name} and saved to {output_path}")

print("All certificates have been generated!")
