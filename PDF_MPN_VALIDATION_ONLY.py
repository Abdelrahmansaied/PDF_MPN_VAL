import re
import pandas as pd
import requests
import io
import streamlit as st
from concurrent.futures import ThreadPoolExecutor
import difflib as dlb
import fitz  # PyMuPDF
import traceback

def clean_string(s):
    """Remove illegal characters from a string."""
    if isinstance(s, str):
        return re.sub(r'[\x00-\x1F\x7F]', '', s)
    return s

def get_pdf_response(pdf):
    """Get PDF response from URL and return content."""
    try:
        response = requests.get(pdf, timeout=10)
        response.raise_for_status()
        return pdf, io.BytesIO(response.content)
    except Exception as e:
        print(f"Error fetching PDF {pdf}: {e}")
        return pdf, None

def get_pdf_text(pdfs):
    """Extract text from a list of PDF URLs."""
    pdf_data = {}
    chunks = [pdfs[i:i + 100] for i in range(0, len(pdfs), 100)]

    for chunk in chunks:
        with ThreadPoolExecutor() as executor:
            results = list(executor.map(get_pdf_response, chunk))

        for pdf, byt in results:
            if byt is not None:
                try:
                    with fitz.open(stream=byt, filetype='pdf') as doc:
                        pdf_data[pdf] = '\n'.join(page.get_text() for page in doc)
                except Exception as e:
                    print(f"Error reading PDF {pdf}: {e}")

    return pdf_data

def pn_validation(pdf_data, part_col, pdf_col, data):
    """Validate parts against extracted PDF data."""
    data['STATUS'] = None
    data['EQUIVALENT'] = None
    data['SIMILARS'] = None

    def set_desc(index):
        part = data[part_col][index]
        pdf_url = data[pdf_col][index]
        if pdf_url not in pdf_data:
            data['STATUS'][index] = 'May be Broken'
            return

        values = pdf_data[pdf_url]
        if len(values) <= 100:
            data['STATUS'][index] = 'OCR'
            return

        exact = re.search(re.escape(part), values, flags=re.IGNORECASE)
        if exact:
            data['STATUS'][index] = 'Exact'
            data['EQUIVALENT'][index] = exact.group(0)
            semi_regex = {
                match.strip() for match in re.findall(r'\b\w*' + re.escape(part) + r'\w*\b', values, flags=re.IGNORECASE)
            }
            if semi_regex:
                data['SIMILARS'][index] = '|'.join(semi_regex)
            return

        dlb_match = dlb.get_close_matches(part, re.split('[ \n]', values), n=1, cutoff=0.65)
        if dlb_match:
            pdf_part = dlb_match[0]
            data['STATUS'][index] = 'Includes or Missed Suffixes'
            data['EQUIVALENT'][index] = pdf_part
            return

        data['STATUS'][index] = 'Not Found'

    with ThreadPoolExecutor() as executor:
        executor.map(set_desc, data.index)

    return data

def search_mpns_in_pdfs(mpns, pdf_data):
    """Search for MPNs in provided PDFs and return matches."""
    found_pdfs = []
    
    for mpn in mpns:
        for pdf_url, content in pdf_data.items():
            if re.search(re.escape(mpn), content, flags=re.IGNORECASE):
                found_pdfs.append({"MPN": mpn, "PDF_URL": pdf_url})
    
    return found_pdfs

def main():
    st.title("MPN PDF Validation App 🛠️")

    # Section for uploading a single file with MPNs and PDF URLs
    st.subheader("Upload Excel file with MPN and PDF URL")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            data = pd.read_excel(uploaded_file)
            st.write("### Uploaded Data:")
            st.dataframe(data)

            if all(col in data.columns for col in ['MPN', 'PDF']):
                pdfs = data['PDF'].tolist()
                pdf_data = get_pdf_text(pdfs)
                result_data = pn_validation(pdf_data, 'MPN', 'PDF', data)

                # Clean the output data
                for col in ['MPN', 'PDF', 'STATUS', 'EQUIVALENT', 'SIMILARS']:
                    result_data[col] = result_data[col].apply(clean_string)

                # Display validation results
                st.subheader("Validation Results")
                STATUS_color = {
                    'Exact': 'green',
                    'Includes or Missed Suffixes': 'orange',
                    'Not Found': 'red',
                    'May be Broken': 'gray'
                }

                for index, row in result_data.iterrows():
                    color = STATUS_color.get(row['STATUS'], 'black')
                    st.markdown(f"<div style='color: {color};'>{row['MPN']} - {row['STATUS']} - {row['EQUIVALENT']} - {row['SIMILARS']}</div>", unsafe_allow_html=True)

                output_file = "MPN_Validation_Result.xlsx"
                result_data.to_excel(output_file, index=False, engine='openpyxl')
                st.sidebar.download_button("Download Results 📥", data=open(output_file, "rb"), file_name=output_file)

            else:
                st.error("The uploaded file must contain 'MPN' and 'PDF' columns.")
        except Exception as e:
            st.error(f"An error occurred while processing: {e}")
            st.error(traceback.format_exc())

    # Section for searching MPNs in separate PDF files
    st.subheader("Upload Excel file with MPNs and another with PDF URLs")
    uploaded_mpn_file = st.file_uploader("Upload Excel file with MPNs", type=["xlsx"], key='mpn')
    uploaded_pdf_file = st.file_uploader("Upload Excel file with PDF URLs", type=["xlsx"], key='pdf')

    if uploaded_mpn_file is not None and uploaded_pdf_file is not None:
        try:
            mpn_data = pd.read_excel(uploaded_mpn_file)
            pdf_data_frame = pd.read_excel(uploaded_pdf_file)

            st.write("### Uploaded MPN Data:")
            st.dataframe(mpn_data)

            st.write("### Uploaded PDF Data:")
            st.dataframe(pdf_data_frame)

            if 'MPN' in mpn_data.columns and 'PDF' in pdf_data_frame.columns:
                mpns = mpn_data['MPN'].tolist()
                pdfs = pdf_data_frame['PDF'].tolist()

                # Get PDF text content
                pdf_data = get_pdf_text(pdfs)

                # Search MPNs in PDFs
                found_pdfs = search_mpns_in_pdfs(mpns, pdf_data)

                # Display results
                if found_pdfs:
                    st.subheader("Found PDFs:")
                    st.write(pd.DataFrame(found_pdfs))
                else:
                    st.write("No PDFs found containing the provided MPNs.")
            else:
                st.error("Both files must contain an 'MPN' and a 'PDF' column.")
        except Exception as e:
            st.error(f"An error occurred while processing: {e}")
            st.error(traceback.format_exc())

if __name__ == "__main__":
    main()
