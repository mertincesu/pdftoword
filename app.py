import streamlit as st
import fitz  # PyMuPDF
from pdf2docx import Converter
from io import BytesIO
import tempfile

def pdf_to_text(pdf_file):
    # Open the PDF file
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    
    text = ""
    # Iterate through the pages
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)  # Load each page
        text += page.get_text()  # Extract text from the page

    return text

def pdf_to_word(pdf_bytes):
    # Save the PDF bytes to a temporary file
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
        temp_pdf.write(pdf_bytes)
        temp_pdf_path = temp_pdf.name
    
    # Prepare a BytesIO object to receive the Word data
    docx_file = BytesIO()
    
    # Use pdf2docx to convert PDF to DOCX
    cv = Converter(temp_pdf_path)
    cv.convert(docx_file)  # convert_method accepts file streams
    cv.close()
    
    docx_file.seek(0)
    
    # Clean up the temporary PDF file
    temp_pdf.close()

    return docx_file

def text_to_txt(text):
    # Create a text file in a BytesIO object
    txt_file = BytesIO()
    txt_file.write(text.encode('utf-8'))
    txt_file.seek(0)
    
    return txt_file

st.title("PDF to Word or TXT Converter")

# Upload PDF
uploaded_pdf = st.file_uploader("Upload your PDF file", type="pdf")

# Selection for output format
output_format = st.selectbox("Select output format", ["Word Document", "Text File"])

if uploaded_pdf:
    if output_format == "Word Document":
        st.write("Converting PDF to Word document...")
        # Read the uploaded PDF file as bytes
        pdf_bytes = uploaded_pdf.read()
        # Convert PDF to Word
        try:
            word_file = pdf_to_word(pdf_bytes)
            st.write("Conversion to Word document completed.")
            
            # Provide download button for the Word file
            st.download_button(
                label="Download Word Document",
                data=word_file,
                file_name="converted_document.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"An error occurred during conversion: {e}")
    elif output_format == "Text File":
        st.write("Extracting text from PDF...")
        # Convert PDF to text
        try:
            # Since we've read the PDF bytes for the word conversion, we need to reinitialize the BytesIO object
            uploaded_pdf.seek(0)
            extracted_text = pdf_to_text(uploaded_pdf)
            st.write("Text extraction completed.")
            
            # Convert text to TXT
            txt_file = text_to_txt(extracted_text)
            st.write("Conversion to TXT file completed.")
            
            # Provide download button for the TXT file
            st.download_button(
                label="Download Text File",
                data=txt_file,
                file_name="extracted_text.txt",
                mime="text/plain"
            )
        except Exception as e:
            st.error(f"An error occurred during text extraction: {e}")
