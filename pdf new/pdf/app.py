import os
import re
from pdf2image import convert_from_path
import pytesseract
from openpyxl import Workbook
from openpyxl.styles import Alignment

def convert_pdf_to_images(pdf_path, output_dir, dpi=300):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    images = convert_from_path(pdf_path, dpi=dpi)
    
    image_paths = []
    for i, image in enumerate(images):
        image_path = os.path.join(output_dir, f'page_{i + 1}.png')
        image.save(image_path, 'PNG')
        image_paths.append(image_path)
        print(f"Saved {image_path}")
    
    return image_paths

def extract_text_from_images(image_paths):
    extracted_texts = []
    for image_path in image_paths:
        text = pytesseract.image_to_string(image_path)
        extracted_texts.append(text)
        print(f"Extracted text from {image_path}")
    return extracted_texts

def clean_and_join_lines(text):
    lines = []
    start_reading = False  

    for line in text.split('\n'):
        
        if re.match(r'^\d+\.\s*|\([a-z]\)\s*', line.strip()):
            start_reading = True

        if start_reading:
            
            cleaned_line = re.sub(r'^\d+\.\s*|\([a-z]\)\s*', '', line.strip())
            if cleaned_line:
                if lines and (cleaned_line[0].islower() or cleaned_line[0] in {'-', '(', '[', '{'}):
                    lines[-1] += ' ' + cleaned_line
                else:
                    lines.append(cleaned_line)
    
    return '\n'.join(lines)

def write_texts_to_excel(texts, excel_path):
    workbook = Workbook()
    worksheet = workbook.active

    row = 1

    for text in texts:
        cleaned_text = clean_and_join_lines(text)
        lines = cleaned_text.split('\n')
        for line in lines:
            if line.strip():
                
                cell = worksheet.cell(row=row, column=1, value=line)
                cell.alignment = Alignment(horizontal='left', vertical='top')
                row += 1
    
    workbook.save(excel_path)
    print(f"Excel file saved at {excel_path}")

def main(pdf_path, output_dir, excel_path):
    image_paths = convert_pdf_to_images(pdf_path, output_dir)
    extracted_texts = extract_text_from_images(image_paths)
    write_texts_to_excel(extracted_texts, excel_path)


pdf_path = "p1.pdf"  
output_dir = "pdf_images"  
excel_path = "output_text.xlsx"  
main(pdf_path, output_dir, excel_path)
