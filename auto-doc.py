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
