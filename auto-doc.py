import openpyxl
from docxtpl import DocxTemplate

def generate_student_docs(excel_file, doc_template):
    workbook = openpyxl.load_workbook(excel_file)
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

import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont

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
    text_bbox = draw.textbbox((0, 0), name, font=font_name)  # Box for text
    text_width = text_bbox[2] - text_bbox[0]  # Text width
    text_height = text_bbox[3] - text_bbox[1] # Text height
    
    # position for centering the text
    image_width = certificate.width
    x_position = (image_width - text_width) //2 # Center horizontally
    
    # Fixed vertical position
    y_position = 629  # Adjust as needed

    # Put the name on the certificates
    draw.text((x_position, y_position), name, fill="orange", font=font_name)

    # Save the certificates to the output folder
    output_path = os.path.join(output_folder, f"{name}.png")
    certificate.save(output_path)

    # Print confirmation
    print(f"Certificate generated for {name} and saved to {output_path}")

print("All certificates have been generated!")
