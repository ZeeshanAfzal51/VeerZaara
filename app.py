import streamlit as st
import fitz  # PyMuPDF
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
import os
from openpyxl import load_workbook
import google.generativeai as genai

# Set up Google Generative AI client
os.environ["GEMINI_API_KEY"] = "AIzaSyDI2DelJZlGyXEPG3_b-Szo-ixRvaB0ydY"
genai.configure(api_key=os.environ["GEMINI_API_KEY"])

# Uploading multiple PDF files
st.markdown("**Upload the Invoice PDFs**")
uploaded_pdfs = st.file_uploader("", type="pdf", accept_multiple_files=True)

# Snippet to attach local Excel file
st.markdown("**Upload the Local Master Excel File**")
uploaded_excel = st.file_uploader("", type="xlsx")

if uploaded_pdfs and uploaded_excel:
    # Loading the workbook and selecting the active sheet
    workbook = load_workbook(uploaded_excel)
    worksheet = workbook.active

    def extract_text_from_pdf(pdf_stream):
        doc = fitz.open(stream=pdf_stream.read(), filetype="pdf")
        text_data = []
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text("text")
            text_data.append(text)
        return text_data

    def convert_pdf_to_images_and_ocr(pdf_stream):
        images = convert_from_path(pdf_stream)
        ocr_results = [pytesseract.image_to_string(image) for image in images]
        return ocr_results

    def combine_text_and_ocr_results(text_data, ocr_results):
        combined_results = []
        for text, ocr_text in zip(text_data, ocr_results):
            combined_results.append(text + "\n" + ocr_text)
        combined_text = "\n".join(combined_results)
        return combined_text

    def extract_parameters_from_response(response_text):
        def sanitize_value(value):
            return value.strip().replace('"', '').replace(',', '')

        parameters = {
            "PO Number": "NA",
            "Invoice Number": "NA",
            "Invoice Amount": "NA",
            "Invoice Date": "NA",
            "CGST Amount": "NA",
            "SGST Amount": "NA",
            "IGST Amount": "NA",
            "Total Tax Amount": "NA",
            "Taxable Amount": "NA",
            "TCS Amount": "NA",
            "IRN Number": "NA",
            "Receiver GSTIN": "NA",
            "Receiver Name": "NA",
            "Vendor GSTIN": "NA",
            "Vendor Name": "NA",
            "Remarks": "NA",
            "Vendor Code": "NA"
        }
        lines = response_text.splitlines()
        for line in lines:
            for key in parameters.keys():
                if key in line:
                    value = sanitize_value(line.split(":")[-1].strip())
                    parameters[key] = value
        return parameters

    # The prompt to send
    prompt = (
        "The following is OCR extracted text from a single invoice PDF. "
        "Please use the OCR extracted text to give a structured summary. "
        "The structured summary should consider information such as PO Number, Invoice Number, Invoice Amount, Invoice Date, "
        "CGST Amount, SGST Amount, IGST Amount, Total Tax Amount, Taxable Amount, TCS Amount, IRN Number, Receiver GSTIN, "
        "Receiver Name, Vendor GSTIN, Vendor Name, Remarks, and Vendor Code. If any of this information is not available or present, "
        "then NA must be denoted next to the value. Please do not give any additional information."
    )

    success_messages = []
    structured_summaries = []

    # Process each PDF and send data to Excel
    for pdf_file in uploaded_pdfs:
        pdf_name = pdf_file.name
        text_data = extract_text_from_pdf(pdf_file)
        ocr_results = convert_pdf_to_images_and_ocr(pdf_file)
        combined_text = combine_text_and_ocr_results(text_data, ocr_results)

        input_text = f"{prompt}\n\n{combined_text}"

        # Generate the response using Google Generative AI
        response = genai.generate(input_text)
        parameters = extract_parameters_from_response(response.text)

        # Add data to the Excel file
        row_data = [parameters[key] for key in parameters.keys()]
        worksheet.append(row_data)

        # Collect success message
        success_messages.append(f"Data from {pdf_name} has been successfully added to the Excel file")

        # Collect structured summary
        structured_summaries.append((pdf_name, parameters))

    # Display all success messages
    for message in success_messages:
        st.write(message)

    # Print the structured summaries after processing all PDFs
    st.write("### Structured Summaries")
    for pdf_name, parameters in structured_summaries:
        st.markdown(f"**{pdf_name} Structured Summary:**")
        st.table(parameters)

    # Save the updated Excel file
    workbook.save(uploaded_excel.name)
    st.write("Excel file has been updated.")

