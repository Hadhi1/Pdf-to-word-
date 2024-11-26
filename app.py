import subprocess
import os
from io import BytesIO

def word_to_pdf(input_stream):
    """
    Convert a Word document to a PDF and return the PDF as a BytesIO object.
    Uses LibreOffice for conversion (Linux and cross-platform).
    """
    # Save the input stream as a temporary Word file
    temp_input = 'temp.docx'
    temp_output = 'temp.pdf'

    # Save the DOCX to a temporary file
    with open(temp_input, 'wb') as temp_file:
        temp_file.write(input_stream.read())

    # Run LibreOffice in headless mode to convert DOCX to PDF
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', temp_input])

    # Load the PDF into a BytesIO object
    pdf_stream = BytesIO()
    with open(temp_output, 'rb') as pdf_file:
        pdf_stream.write(pdf_file.read())

    # Clean up temporary files
    os.remove(temp_input)
    os.remove(temp_output)

    pdf_stream.seek(0)
    return pdf_stream
