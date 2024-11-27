import tkinter as tk
from tkinter import messagebox

# Functions from the provided script
def generate_certificates(excel_file, template, output_folder, font_file, font_size=100, text_color="orange", y_position=629):
    import os
    import pandas as pd
    from PIL import Image, ImageDraw, ImageFont

    data = pd.read_excel(excel_file)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    font = ImageFont.truetype(font_file, font_size)

    for index, row in data.iterrows():
        name = row["Name"]
        certificate = Image.open(template)
        draw = ImageDraw.Draw(certificate)
        text_bbox = draw.textbbox((0, 0), name, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        x_position = (certificate.width - text_width) // 2
        draw.text((x_position, y_position), name, fill=text_color, font=font)
        output_path = os.path.join(output_folder, f"{name}.png")
        certificate.save(output_path)
        print(f"Certificate generated for {name} and saved to {output_path}")

def generate_template_documents(excel_file, template_pnc_file, output_template_folder):
    import os
    import openpyxl
    from docxtpl import DocxTemplate
    from docx2pdf import convert

    if not os.path.exists(output_template_folder):
        os.makedirs(output_template_folder)

    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    data = list(sheet.values)

    template_pdf_folder = os.path.join(output_template_folder, "pdfs")
    if not os.path.exists(template_pdf_folder):
        os.makedirs(template_pdf_folder)

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
            "e1_g": student[12],
            "wd": student[13],
            "wd_g": student[14],
            "algo": student[15],
            "al_g": student[16],
            "p2": student[17],
            "p2_g": student[18],
            "e2": student[19],
            "e2_g": student[20],
            "sd": student[21],
            "sd_g": student[22],
            "js": student[23],
            "js_g": student[24],
            "php": student[25],
            "ph_g": student[26],
            "db": student[27],
            "db_g": student[28],
            "vc1": student[29],
            "v1_g": student[30],
            "node": student[31],
            "no_g": student[32],
            "e3": student[33],
            "e3_g": student[34],
            "p3": student[35],
            "p3_g": student[36],
            "oop": student[37],
            "op_g": student[38],
            "lar": student[39],
            "lar_g": student[40],
            "vue": student[41],
            "vu_g": student[42],
            "vc2": student[43],
            "v2_g": student[44],
            "e4": student[45],
            "e4_g": student[46],
            "p4": student[47],
            "p4_g": student[48],
            "int": student[49],
            "in_g": student[50],
            "cur_date": student[51],
        })

        docx_data_filename = os.path.join(output_template_folder, f"{student[0]}.docx")
        student_data.save(docx_data_filename)
        print(f"Generated document: {docx_data_filename}")

        pdf_filename = os.path.join(template_pdf_folder, f"{student[0]}.pdf")
        convert(docx_data_filename, pdf_filename)
        print(f"Converted to PDF: {pdf_filename}")

def generate_documents(excel_file, template_file, output_folder):
    import os
    import openpyxl
    from docxtpl import DocxTemplate
    from docx2pdf import convert

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    data = list(sheet.values)

    pdf_folder = os.path.join(output_folder, "pdfs")
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

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
            "ed_e": student[11],
            "cur_date": student[12],
        })

        docx_filename = os.path.join(output_folder, f"{student[0]}.docx")
        student_template.save(docx_filename)
        print(f"Generated document: {docx_filename}")

        pdf_filename = os.path.join(pdf_folder, f"{student[0]}.pdf")
        convert(docx_filename, pdf_filename)
        print(f"Converted to PDF: {pdf_filename}")

# Interface
def generate(option):
    try:
        if option == "certificates":
            generate_certificates(
                excel_file="certificate_data.xlsx",
                template="template.png",
                output_folder="generated_certificate",
                font_file="calibrib.ttf",
            )
        elif option == "transcripts":
            generate_template_documents(
                excel_file="trainscript_data.xlsx",
                template_pnc_file="trainscript_template.docx",
                output_template_folder="output_data_documents",
            )
        elif option == "associates":
            generate_documents(
                excel_file="associate_degree.xlsx",
                template_file="associate_template.docx",
                output_folder="output_documents",
            )
        elif option == "all":
            generate_certificates(
                excel_file="certificate_data.xlsx",
                template="template.png",
                output_folder="generated_certificate",
                font_file="calibrib.ttf",
            )
            generate_template_documents(
                excel_file="trainscript_data.xlsx",
                template_pnc_file="trainscript_template.docx",
                output_template_folder="output_data_documents",
            )
            generate_documents(
                excel_file="associate_degree.xlsx",
                template_file="associate_template.docx",
                output_folder="output_documents",
            )
        messagebox.showinfo("Success", f"{option.capitalize()} generated successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def create_interface():
    root = tk.Tk()
    root.title("Document Generator")
    root.geometry("400x400")

    tk.Label(root, text="Select Document to Generate", font=("Arial", 25)).pack(pady=50)

    tk.Button(root, text="Generate Certificates", command=lambda: generate("certificates"), width=50, bg="lightblue", font=("Arial", 15)).pack(pady=10, ipady=10)
    tk.Button(root, text="Generate Transcripts", command=lambda: generate("transcripts"), width=50, bg="lightblue", font=("Arial", 15)).pack(pady=10, ipady=10)
    tk.Button(root, text="Generate Associate", command=lambda: generate("associates"), width=50, bg="lightblue", font=("Arial", 15)).pack(pady=10, ipady=10)
    tk.Button(root, text="Generate All", command=lambda: generate("all"), width=50, bg="lightgreen", font=("Arial", 15)).pack(pady=10 , ipady=10)

    root.mainloop()

create_interface()
