import streamlit as st
import pdfplumber
import pandas as pd
from docx import Document
from io import BytesIO

def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ''
        for page in pdf.pages:
            text += page.extract_text()
    return text

def extract_text_from_doc(doc_path):
    doc = Document(doc_path)
    text = ' '.join([paragraph.text for paragraph in doc.paragraphs])
    return text

def save_to_excel_and_csv(data, excel_path, csv_path):
    df = pd.DataFrame(data, columns=['Name', 'Email', 'Contact', 'Text'])
    # Save to Excel using BytesIO
    excel_data = BytesIO()
    df.to_excel(excel_data, index=False)
    excel_data.seek(0) # Move to the beginning of the BytesIO object
    # Save to CSV
    df.to_csv(csv_path, index=False)
    return excel_data

def process_files(files):
    data = []
    for file in files:
        if file.type == "application/pdf":
            text = extract_text_from_pdf(file)
        elif file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            text = extract_text_from_doc(file)
        else:
            st.error(f"Unsupported file type: {file.type}. Please upload PDF or DOCX files.")
            continue

        text_parts = text.split('\n')
        if len(text_parts) < 3:
            st.error("Insufficient data in the file. Please ensure the CV contains a name, email, and contact number.")
            continue

        name = text_parts[0]
        email = text_parts[1]
        contact = text_parts[2]

        data.append([name, email, contact, text])
    return data

if __name__ == "__main__":
    st.title("CV Processor")

    uploaded_files = st.file_uploader("Upload your CVs (PDF or DOCX)", type=["pdf", "docx"], accept_multiple_files=True)
    if uploaded_files:
        data = process_files(uploaded_files)
        if data:
            st.success("Files processed successfully.")
            st.write(data)

            # Save the data to a file
            excel_path = "./output.xlsx"
            csv_path = "./output.csv"
            excel_data = save_to_excel_and_csv(data, excel_path, csv_path)

            # Provide download links
            st.markdown(f"Download the processed data as:")
            # Use st.download_button for Excel file
            st.download_button(
                label="Download Excel",
                data=excel_data,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Failed to process the files. Please check the file formats and contents.")
