import pandas as pd
import os
from PIL import Image, ImageDraw, ImageFont
#Generation Certificate---------------------------------
def generate_certificates(excel_file, template, output_folder, font_file, font_size=100, text_color="orange", y_position=629):
    # Load the Excel file
    data = pd.read_excel(excel_file)

    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Load the font
    font = ImageFont.truetype(font_file, font_size)

    # Generate certificates
    for index, row in data.iterrows():
        name = row["Name"]
        certificate = Image.open(template)
        draw = ImageDraw.Draw(certificate)

        # Measure the width and height of the name text
        text_bbox = draw.textbbox((0, 0), name, font=font)  # Bounding box for text
        text_width = text_bbox[2] - text_bbox[0]  # Text width

        # Calculate position for centering the text
        image_width = certificate.width
        x_position = (image_width - text_width) // 2  # Center horizontally

        # Draw the name on the certificate
        draw.text((x_position, y_position), name, fill=text_color, font=font)

        # Save the certificate to the output folder
        output_path = os.path.join(output_folder, f"{name}.png")
        certificate.save(output_path)

        # Print confirmation
        print(f"Certificate generated for {name} and saved to {output_path}")

# Example usage
generate_certificates(
    excel_file="students_data.xlsx",
    template="template.png",
    output_folder="generated_certificate",
    font_file="calibrib.ttf"
)
import openpyxl
from docxtpl import DocxTemplate
from docx2pdf import convert

def generate_template_documents(excel_file, template_pnc_file, output_template_folder):

     
    if not os.path.exists(output_template_folder):
        os.makedirs(output_template_folder)
        print(f"Created output folder: {output_template_folder}")   
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    data = list(sheet.values) 

    
    template_pdf_folder = os.path.join(output_template_folder, "pdfs")
    if not os.path.exists(template_pdf_folder):
        os.makedirs(template_pdf_folder)
        print(f"Created PDF folder: {template_pdf_folder}")
    
    generated_files = []

    
    student_data = DocxTemplate(template_pnc_file)
    for student in data[1:]:
        student_data.render({
           "student_id": student[0],
            "first_name": student[1],
            "last_name": student[2],
            "logic": student[3],
            "l_g": student[4],
            "bcum": student[5],
            "bc_g": student[6],
            "design": student[7],
            "d_g": student[8],
            "p1": student[9],
            "p1_g": student[10],
            "e1": student[11],
            "wd": student[12],
            "wd_g": student[13],
            "algo": student[14],
            "al_g": student[15],
            "p2": student[16],
            "p2_g": student[17],
            "e2": student[18],
            "e2_g": student[19],
            "sd": student[20],
            "sd_g": student[21],
            "js": student[22],
            "js_g": student[23],
            "php": student[24],
            "ph_g": student[25],
            "db": student[26],
            "db_g": student[27],
            "vc1": student[28],
            "v1_g": student[29],
            "node": student[30],
            "no_g": student[31],
            "e3": student[32],
            "e3_g": student[33],
            "p3": student[34],
            "p3_g": student[35],
            "oop": student[36],
            "op_g": student[37],
            "lar": student[38],
            "lar_g": student[39],
            "vue": student[40],
            "vu_g": student[41],
            "vc2": student[42],
            "v2_g": student[43],
            "e4": student[44],
            "e4_g": student[45],
            "p4": student[46],
            "p4_g": student[47],
            "int": student[48],
            "in_g": student[49],
            "cur_date": student[50]
        })

        
        docx_data_filename = os.path.join(output_template_folder, f"{student[0]}.docx")
        student_data.save(docx_data_filename)
        print(f"Generated document: {docx_data_filename}")
        generated_files.append(docx_data_filename)
        
        
        pdf_filename = os.path.join(template_pdf_folder, f"{student[0]}.pdf")
        convert(docx_data_filename, pdf_filename)
        print(f"Converted to PDF: {pdf_filename}")
        generated_files.append(pdf_filename)
generate_template_documents("data.xlsx", "template-pnc.docx", "output_data_documents")


#funtion for generate associate
def generate_documents(excel_file, template_file, output_folder):

     # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")   
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    template = list(sheet.values) 

    #folder to store file PDF
    pdf_folder = os.path.join(output_folder, "pdfs")
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)
        print(f"Created PDF folder: {pdf_folder}")
    
    generated_files = []

    #enter template from excel to word file
    student_template = DocxTemplate(template_file)
    for student in template[1:]:
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
            "ed_e": student[11],
            "cur_date": student[12]
        })

        #save the word document
        docx_filename = os.path.join(output_folder, f"{student[0]}.docx")
        student_template.save(docx_filename)
        print(f"Generated document: {docx_filename}")
        generated_files.append(docx_filename)
        
        # Convert the Word document to PDF
        pdf_associate_filename = os.path.join(pdf_folder, f"{student[0]}.pdf")
        convert(docx_filename, pdf_associate_filename)
        print(f"Converted to PDF: {pdf_associate_filename}")
        generated_files.append(pdf_associate_filename)

generate_documents("associate_degree.xlsx", "WEP_temporary_ template.docx", "output_documents")


