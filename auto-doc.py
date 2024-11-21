import os
import openpyxl
from docxtpl import DocxTemplate
from docx2pdf import convert

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

