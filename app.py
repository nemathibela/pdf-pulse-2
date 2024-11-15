import io
import os
import zipfile

import PyPDF2
import fitz
import openpyxl
import pandas as pd
import tabula
from flask import Flask, request, send_file
from flask_cors import CORS
from pdf2docx import Converter
from dotenv import load_dotenv

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

load_dotenv()

#Convert PDF to Word
@app.route('/convert', methods=['POST'])
def convert_pdf_to_word():
    pdf_file = request.files['pdf_file']
    output_filename = request.form.get('output_filename', 'output.docx')

    # Saves the uploaded PDF file temporarily
    pdf_path = 'temp.pdf'
    pdf_file.save(pdf_path)

    # Convert PDF to DOCX
    docx_path = output_filename
    converter = Converter(pdf_path)
    converter.convert(docx_path)
    converter.close()

    # Cleanup temporary files
    os.remove(pdf_path)

    return send_file(docx_path, as_attachment=True, download_name=output_filename)

#Convert to Image
@app.route('/convert-image', methods=['POST'])
def convert_pdf_to_image():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    # Use BytesIO to handle the PDF in memory
    pdf_bytes = io.BytesIO(file.read())

    # Convert PDF to images
    image_paths = []
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")

    # Create a BytesIO object for the zip file
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        for page_number in range(len(pdf_document)):
            page = pdf_document.load_page(page_number)
            pix = page.get_pixmap()
            image_data = io.BytesIO()
            # Save the image with a filename to determine the format
            image_filename = f'page_{page_number + 1}.png'
            pix.save(image_filename)  # Save to a file-like object, not in memory

            # Get the bytes of the saved image
            with open(image_filename, 'rb') as img_file:
                zipf.writestr(image_filename, img_file.read())

    # Prepare the zip file for download
    zip_buffer.seek(0)  # Rewind the BytesIO object to the beginning
    return send_file(zip_buffer, as_attachment=True, download_name='converted_images.zip', mimetype='application/zip')

#Excel Conversion
@app.route('/convert_excel', methods=['POST'])

def convert_pdf_to_excel():
    pdf_file = request.files['pdf_file']
    output_filename = request.form.get('output_filename', 'output.xlsx')

    # Save the uploaded PDF file temporarily
    pdf_path = 'temp.pdf'
    pdf_file.save(pdf_path)
    excel_path = output_filename

    try:
        # Read tables from PDF and convert to Excel
        tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
        with pd.ExcelWriter(excel_path) as writer:
            for i, table in enumerate(tables):
                table.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)

        # Cleanup temporary files
        os.remove(pdf_path)
        return send_file(excel_path, as_attachment=True, download_name=output_filename)
    except Exception as e:
        return {"error": str(e)}, 500

#Converting to Excel
@app.route('/convert-to-excel', methods=['POST'])
def convert_to_excel():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    # Use BytesIO to handle the FileStorage object
    pdf_reader = PyPDF2.PdfReader(io.BytesIO(file.read()))
    extracted_text = []

    # Extract text from each page
    for page in pdf_reader.pages:
        extracted_text.append(page.extract_text())

    # Create a new Excel workbook and add a worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write the extracted text to the Excel sheet
    for row_index, text in enumerate(extracted_text):
        sheet.cell(row=row_index + 1, column=1, value=text)

    # Save the workbook to a BytesIO object
    excel_buffer = io.BytesIO()
    workbook.save(excel_buffer)
    excel_buffer.seek(0)  # Rewind the buffer for reading

    # Send the Excel file as a response
    return send_file(excel_buffer, as_attachment=True, download_name='converted_file.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    port = int(os.environ.get("PORT"))
    app.run(host='0.0.0.0', port=5000)