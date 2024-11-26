import pypandoc

def word_to_pdf(input_stream):
    input_file = 'temp.docx'
    output_file = 'temp.pdf'

    # Save the input stream as a temporary Word file
    with open(input_file, 'wb') as f:
        f.write(input_stream.read())

    # Use pypandoc to convert the DOCX to PDF
    output = pypandoc.convert_file(input_file, 'pdf', outputfile=output_file)

    with open(output_file, 'rb') as f:
        pdf_stream = BytesIO(f.read())

    # Clean up temporary files
    os.remove(input_file)
    os.remove(output_file)

    pdf_stream.seek(0)
    return pdf_stream
