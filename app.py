import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from io import BytesIO

def pdf_to_text(pdf_file):
    # Open the PDF file
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    
    text = ""
    # Iterate through the pages
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)  # Load each page
        text += page.get_text()  # Extract text from the page

    return text

def text_to_word(text):
    # Create a Word document
    doc = Document()
    # Add text to the document
    for paragraph in text.split('\n'):
        doc.add_paragraph(paragraph)
    
    # Save the document to a BytesIO object
    word_file = BytesIO()
    doc.save(word_file)
    word_file.seek(0)
    
    return word_file

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
    # Convert PDF to text
    st.write("Converting PDF to text...")
    extracted_text = pdf_to_text(uploaded_pdf)
    st.write("Text extraction completed.")

    if output_format == "Word Document":
        # Convert text to Word
        st.write("Converting text to Word document...")
        file = text_to_word(extracted_text)
        file_extension = "docx"
        mime_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        file_name = "converted_document.docx"
    else:
        # Convert text to TXT
        st.write("Converting text to TXT file...")
        file = text_to_txt(extracted_text)
        file_extension = "txt"
        mime_type = "text/plain"
        file_name = "converted_document.txt"

    st.write(f"Conversion to {output_format} completed.")

    # Provide download button for the file
    st.download_button(
        label=f"Download {output_format}",
        data=file,
        file_name=file_name,
        mime=mime_type
    )
