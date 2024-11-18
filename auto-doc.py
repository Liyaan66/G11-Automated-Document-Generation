import openpyxl
from docxtpl import DocxTemplate

def generate_student_docs(excel_file, doc_template):
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active 
    student_data = list(sheet.values)[1:]  
    student_info = DocxTemplate(doc_template)
    for student in student_data:
        student_info.render({
            "student_name": student[0] 
        })      
        student_name = student[0]  
        my_file = f"{student_name}.docx"
        student_info.save(my_file)
    print("Documents generated successfully.")
generate_student_docs("students_data.xlsx", "Doc1.docx")
