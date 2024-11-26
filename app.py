from flask import Flask, request, jsonify, send_file
from io import BytesIO
from docx import Document
from PyPDF2 import PdfReader, PdfWriter
import pythoncom  # Required for comtypes
import comtypes.client  # For Word-to-PDF conversion (Windows only)

app = Flask(__name__)


def word_to_pdf(input_stream):
    """
    Convert a Word document to a PDF and return the PDF as a BytesIO object.
    Requires Microsoft Word installed on the system (Windows only).
    """
    pythoncom.CoInitialize()
    word = comtypes.client.CreateObject("Word.Application")
    temp_input = 'temp.docx'
    temp_output = 'temp.pdf'

    # Save the input_stream as a temporary Word file
    with open(temp_input, 'wb') as temp_file:
        temp_file.write(input_stream.read())

    # Convert to PDF
    doc = word.Documents.Open(temp_input)
    doc.SaveAs(temp_output, FileFormat=17)  # FileFormat=17 for PDF
    doc.Close()
    word.Quit()

    # Load the PDF into a BytesIO object
    pdf_stream = BytesIO()
    with open(temp_output, 'rb') as pdf_file:
        pdf_stream.write(pdf_file.read())

    # Clean up temporary files
    import os
    os.remove(temp_input)
    os.remove(temp_output)

    pdf_stream.seek(0)
    return pdf_stream


def pdf_to_word(input_stream):
    """
    Convert a PDF to a Word document and return the Word document as a BytesIO object.
    """
    pdf_reader = PdfReader(input_stream)
    doc = Document()

    for page in pdf_reader.pages:
        text = page.extract_text()
        if text:
            doc.add_paragraph(text)

    word_stream = BytesIO()
    doc.save(word_stream)
    word_stream.seek(0)
    return word_stream


@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']

    if file:
        filename = file.filename
        file_extension = filename.rsplit('.', 1)[1].lower()

        try:
            if file_extension == 'docx':
                pdf_stream = word_to_pdf(file.stream)
                return send_file(pdf_stream, download_name='converted.pdf', as_attachment=True)

            elif file_extension == 'pdf':
                word_stream = pdf_to_word(file.stream)
                return send_file(word_stream, download_name='converted.docx', as_attachment=True)

            else:
                return jsonify({'error': 'Unsupported file type'}), 400

        except Exception as e:
            return jsonify({'error': str(e)}), 500

    return jsonify({'error': 'Invalid file'}), 400


if __name__ == '__main__':
    app.run(debug=True)
