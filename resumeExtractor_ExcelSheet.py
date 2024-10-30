import os
import pandas as pd
import streamlit as st
import cohere
from dotenv import load_dotenv
from zipfile import ZipFile
import io
import fitz  # PyMuPDF

def pdf_to_text(pdf_path):
    try:
        # Open the PDF document
        doc = fitz.open(pdf_path)
        text = ''
        for page in doc:
            text += page.get_text()
        return text
    except Exception as e:
        print(f"Error processing file {pdf_path}: {e}")
        return None

def write_files(input_dir, output_path):
    if not os.path.isdir(input_dir):
        st.write(f"Input directory '{input_dir}' does not exist.")
        return
    
    os.makedirs(output_path, exist_ok=True)
    
    for file in os.listdir(input_dir):
       
        input_file = os.path.join(input_dir, file)
        if os.path.isdir(input_file):
            continue  # Skip directories

        pdf_text = pdf_to_text(input_file)
        if pdf_text is None:
            continue  # Skip if the PDF could not be read

        output_file = os.path.join(output_path, f"{os.path.splitext(file)[0]}.txt")
        with open(output_file, 'w') as f:
            f.write(pdf_text)

# Load environment variables
load_dotenv()
api_key = os.getenv("COHERE_API_KEY")
cohere_client = cohere.Client(api_key)

def extract_info(document_text):
    prompt = f"""
    Extract the name, email id, phone number, age/date of birth, educational qualification
    from the following document. If any information is missing, 
    write 'N/A' in that cell. Include the current company the candidate is working at,
    the current designation, and the last company worked at before the current one.

    Document: 
    {document_text}

    Format the output as follows:
    | <Candidate Name> | <Email ID> | <Phone Number> | <Age/Date of birth> | <Educational Qualification> | <Current company> | <Current Designation> | <Last Company> |
    """
    response = cohere_client.generate(
        model='command-xlarge-nightly',
        prompt=prompt,
        max_tokens=1500,
        temperature=0.7
    )
    return response.generations[0].text
def process_files(uploaded_files):
    data = []
    num_files = len(uploaded_files)  # Get the number of files to process
    progress_bar = st.progress(0)  # Initialize the progress bar

    for i, uploaded_file in enumerate(uploaded_files):
        document_text = uploaded_file.read().decode("utf-8")
        extracted_text = extract_info(document_text)
        data_row = [cell.strip() for cell in extracted_text.split('|')[1:9]]
        data.append(data_row)

        # Update the progress bar
        progress = (i + 1) / num_files
        progress_bar.progress(progress)

        # Display the number of files completed
        st.write(f"Files completed: {i + 1} / {num_files}")

    columns = ["Name", "Email ID", "Phone Number", "Age/Date of birth", "Educational Qualification",
               "Current Company", "Current Designation", "Last Company"]
    df = pd.DataFrame(data, columns=columns)
    return df


st.title("Resume Information Extractor")
st.write("Upload a folder of resumes (in a .zip file), and receive an Excel sheet with the extracted details.")

uploaded_zip = st.file_uploader("Upload a .zip file of resume text files", type=["zip"])

if uploaded_zip is not None:
    try:
        with ZipFile(io.BytesIO(uploaded_zip.read())) as zip_file:
            # List the contents of the zip file
            zip_contents = zip_file.namelist()
            # st.write("Files in the uploaded zip:", zip_contents)

            # Ensure we have files to extract
            if not zip_contents:
                st.error("The uploaded zip file is empty.")
            else:
                extracted_dir = "extracted_resumes"
                os.makedirs(extracted_dir, exist_ok=True)

                # Extract files
                # zip_file.extractall(extracted_dir)
                for file_name in zip_contents:
                    # Extract the file name only (without parent folders)
                    extracted_file_name = os.path.basename(file_name)
                    extracted_file_path = os.path.join(extracted_dir, extracted_file_name)
                    
                    with zip_file.open(file_name) as source_file:
                        with open(extracted_file_path, 'wb') as dest_file:
                            dest_file.write(source_file.read())

                # Convert all PDFs in the extracted directory to text files
                text_output_dir = "text_resumes"
                write_files(extracted_dir, text_output_dir)

                # Check if text files were created
                text_files = [f for f in os.listdir(text_output_dir) if f.endswith('.txt')]
                if not text_files:
                    st.error("No text files were created from the resumes.")
                else:

                    # Process extracted text files and create the DataFrame
                    uploaded_files = [open(os.path.join(text_output_dir, file), 'rb') for file in text_files]
                    df = process_files(uploaded_files)

                    # Save DataFrame to Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name="Extracted_Info")
                    output.seek(0)

                    # Allow user to download the generated Excel file
                    st.download_button(
                        label="Download Extracted Information",
                        data=output,
                        file_name="Tracker_DishaOutsourcing.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("File processed successfully! Click the button to download.")
    except Exception as e:
        st.error(f"Error processing the uploaded zip file: {e}")
