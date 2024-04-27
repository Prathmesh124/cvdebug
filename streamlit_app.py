import streamlit as st
import pdfplumber
import pandas as pd
from docx import Document

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
    df.to_excel(excel_path, index=False)
    df.to_csv(csv_path, index=False)

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

        # Assuming the name is at the beginning of the text, followed by email and contact number
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
            # Display the extracted data
            st.write(data)

            # Save the data to a file
            excel_path = "output.xlsx"
            csv_path = "output.csv"
            save_to_excel_and_csv(data, excel_path, csv_path)

            # Provide download links
            st.markdown(f"Download the processed data as:")
            st.markdown(f"- [Excel]({excel_path})")
            st.markdown(f"- [CSV]({csv_path})")
        else:
            st.error("Failed to process the files. Please check the file formats and contents.")
