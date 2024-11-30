import tkinter as tk
from tkinter import messagebox

# Function to generate certificates using the given template and input data
def generate_certificates(excel_file, template, output_folder, font_file, font_size=100, text_color="orange", y_position=629):
    import os
    import pandas as pd
    from PIL import Image, ImageDraw, ImageFont

    # Load data from the Excel file
    data = pd.read_excel(excel_file)

    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Load the font file
    font = ImageFont.truetype(font_file, font_size)

    # Generate certificates for each name in the data
    for index, row in data.iterrows():
        name = row["Name"]
        certificate = Image.open(template)  # Open the template image
        draw = ImageDraw.Draw(certificate)  # Draw on the image

        # Center the text horizontally
        text_bbox = draw.textbbox((0, 0), name, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        x_position = (certificate.width - text_width) // 2

        # Add the name to the certificate
        draw.text((x_position, y_position), name, fill=text_color, font=font)

        # Save the generated certificate
        output_path = os.path.join(output_folder, f"{name}.png")
        certificate.save(output_path)
        print(f"Certificate generated for {name} and saved to {output_path}")

# Function to generate transcripts based on the template
def generate_template_documents(excel_file, template_pnc_file, output_template_folder):
    import os
    import openpyxl
    from docxtpl import DocxTemplate
    from docx2pdf import convert

    # Ensure the output folder exists
    if not os.path.exists(output_template_folder):
        os.makedirs(output_template_folder)

    # Load data from the Excel file
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    data = list(sheet.values)

    # Create a folder for PDFs if it doesn't exist
    template_pdf_folder = os.path.join(output_template_folder, "pdfs")
    if not os.path.exists(template_pdf_folder):
        os.makedirs(template_pdf_folder)

    # Load the Word document template
    student_data = DocxTemplate(template_pnc_file)

    # Generate and save a transcript for each student
    for student in data[1:]:
        # Fill the template with student data
        student_data.render({
            "student_id": student[0],
            "first_name": student[1],
            "last_name": student[2],
            # Additional data mappings go here...
            "cur_date": student[51],
        })

        # Save the document as a Word file
        docx_data_filename = os.path.join(output_template_folder, f"{student[0]}.docx")
        student_data.save(docx_data_filename)
        print(f"Generated document: {docx_data_filename}")

        # Convert the Word file to a PDF
        pdf_filename = os.path.join(template_pdf_folder, f"{student[0]}.pdf")
        convert(docx_data_filename, pdf_filename)
        print(f"Converted to PDF: {pdf_filename}")

# Function to generate associate degree documents
def generate_documents(excel_file, template_file, output_folder):
    import os
    import openpyxl
    from docxtpl import DocxTemplate
    from docx2pdf import convert

    # Ensure the output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Load data from the Excel file
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    data = list(sheet.values)

    # Create a folder for PDFs if it doesn't exist
    pdf_folder = os.path.join(output_folder, "pdfs")
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    # Load the Word document template
    student_template = DocxTemplate(template_file)

    # Generate and save documents for each student
    for student in data[1:]:
        # Fill the template with student data
        student_template.render({
            "id_kh": student[0],
            "id_e": student[1],
            # Additional data mappings go here...
            "cur_date": student[12],
        })

        # Save the document as a Word file
        docx_filename = os.path.join(output_folder, f"{student[0]}.docx")
        student_template.save(docx_filename)
        print(f"Generated document: {docx_filename}")

        # Convert the Word file to a PDF
        pdf_filename = os.path.join(pdf_folder, f"{student[0]}.pdf")
        convert(docx_filename, pdf_filename)
        print(f"Converted to PDF: {pdf_filename}")

# Function to determine which document type to generate
def generate(option):
    try:
        if option == "certificates":
            # Generate certificates
            generate_certificates(
                excel_file="certificate_data.xlsx",
                template="template.png",
                output_folder="generated_certificate",
                font_file="calibrib.ttf",
            )
        elif option == "transcripts":
            # Generate transcripts
            generate_template_documents(
                excel_file="trainscript_data.xlsx",
                template_pnc_file="trainscript_template.docx",
                output_template_folder="output_data_documents",
            )
        elif option == "associates":
            # Generate associate degree documents
            generate_documents(
                excel_file="associate_degree.xlsx",
                template_file="associate_template.docx",
                output_folder="output_documents",
            )
        elif option == "all":
            # Generate all document types
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
        # Show success message
        messagebox.showinfo("Success", f"{option.capitalize()} generated successfully!")
    except Exception as e:
        # Show error message
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to create the GUI interface
def create_interface():
    root = tk.Tk()
    root.title("Document Generator")
    root.geometry("400x400")

    # Title label
    tk.Label(root, text="Select Document to Generate", font=("Arial", 25)).pack(pady=50)

    # Buttons for each document type
    tk.Button(root, text="Generate Certificates", command=lambda: generate("certificates"), width=50, bg="lightblue", font=("Arial", 15)).pack(pady=10, ipady=10)
    tk.Button(root, text="Generate Transcripts", command=lambda: generate("transcripts"), width=50, bg="lightblue", font=("Arial", 15)).pack(pady=10, ipady=10)
    tk.Button(root, text="Generate Associate", command=lambda: generate("associates"), width=50, bg="lightblue", font=("Arial", 15)).pack(pady=10, ipady=10)
    tk.Button(root, text="Generate All", command=lambda: generate("all"), width=50, bg="lightgreen", font=("Arial", 15)).pack(pady=10 , ipady=10)

    # Run the application
    root.mainloop()

# Start the application
create_interface()
