import streamlit as st
from pdf2docx import Converter
from io import BytesIO
import tempfile
from docx import Document
from PIL import Image
from io import BytesIO as PILBytesIO
import fitz  # PyMuPDF

def pdf_to_text(pdf_file):
    doc = fitz.open(stream=pdf_file, filetype="pdf")
    text = ""
    for page_num in range(doc.page_count):
        if 'uploaded_pdf' not in st.session_state or st.session_state.uploaded_pdf is None:
            raise Exception("File upload canceled during conversion.")
        page = doc.load_page(page_num)
        text += page.get_text()
    return text

def compress_image(image):
    if image.mode == 'RGBA':
        image = image.convert('RGB')
    with PILBytesIO() as output:
        image.save(output, format="JPEG", quality=50)
        compressed_image = output.getvalue()
    return compressed_image

def optimize_images_in_docx(docx_file):
    doc = Document(docx_file)
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image = Image.open(BytesIO(rel.target_part.blob))
            compressed_image = compress_image(image)
            rel.target_part._blob = compressed_image
    optimized_docx = BytesIO()
    doc.save(optimized_docx)
    optimized_docx.seek(0)
    return optimized_docx

def pdf_to_word(pdf_bytes):
    # Save the PDF bytes to a temporary file
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
        temp_pdf.write(pdf_bytes)
        temp_pdf_path = temp_pdf.name
    
    # Prepare a BytesIO object to receive the Word data
    docx_file = BytesIO()
    
    # Use pdf2docx to convert PDF to DOCX
    cv = Converter(temp_pdf_path)
    cv.convert(docx_file)
    cv.close()
    
    docx_file.seek(0)
    
    # Optimize images in the DOCX file
    optimized_docx = optimize_images_in_docx(docx_file)
    
    return optimized_docx

def text_to_txt(text):
    txt_file = BytesIO()
    txt_file.write(text.encode('utf-8'))
    txt_file.seek(0)
    return txt_file

st.title("PDF Converter: Word, Text, and More")

# Session state to track uploaded file and conversion status
if 'uploaded_pdf' not in st.session_state:
    st.session_state.uploaded_pdf = None

if 'conversion_done' not in st.session_state:
    st.session_state.conversion_done = False

# Upload PDF
uploaded_pdf = st.file_uploader("Upload your PDF file", type="pdf")

if uploaded_pdf:
    st.session_state.uploaded_pdf = uploaded_pdf.getvalue()
    st.session_state.conversion_done = False  # Reset conversion status when a new file is uploaded
else:
    st.session_state.uploaded_pdf = None

# Select output format (visible when the app starts)
output_format = st.selectbox("Select output format", ["Word Document", "Text File", "Extracted Text"])

# Initialize empty variables for holding file data
download_data = None
download_label = ""
download_mime = ""

# Start conversion button
if st.button("Start Conversion"):
    if st.session_state.uploaded_pdf is None:
        st.error("No file uploaded. Please upload a PDF file first.")
    elif not st.session_state.conversion_done:
        try:
            if output_format == "Word Document":
                st.write("Converting PDF to Word document...")
                word_file = pdf_to_word(st.session_state.uploaded_pdf)
                download_data = word_file
                download_label = "Download Word Document"
                download_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                download_filename = "converted_document.docx"

            elif output_format == "Text File":
                st.write("Extracting text from PDF...")
                extracted_text = pdf_to_text(BytesIO(st.session_state.uploaded_pdf))
                txt_file = text_to_txt(extracted_text)
                download_data = txt_file
                download_label = "Download Text File"
                download_mime = "text/plain"
                download_filename = "extracted_text.txt"

            elif output_format == "Extracted Text":
                st.write("Extracting text from PDF...")
                extracted_text = pdf_to_text(BytesIO(st.session_state.uploaded_pdf))
                st.text_area("Extracted Text", value=extracted_text, height=300)

            st.session_state.conversion_done = True  # Mark conversion as done
        except Exception as e:
            st.error(f"An error occurred during conversion: {e}")
            print(f"Error details: {e}")
    else:
        st.warning("File already converted. Please upload a new file or press the button again to reconvert.")

# Show download button if data is ready
if download_data:
    st.download_button(
        label=download_label,
        data=download_data,
        file_name=download_filename,
        mime=download_mime
    )
