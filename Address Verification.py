import streamlit as st
import pandas as pd
from datetime import datetime
import os
from fuzzywuzzy import fuzz

def generate_output_filename():
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    current_datetime = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(desktop_path, f"output_{current_datetime}.xlsx")

def preprocess_address(address):
    if pd.notna(address):
        tokens = sorted(''.join(e for e in str(address).lower() if e.isalnum() or e.isspace()).split())
        return " ".join(tokens)
    else:
        return ""

def address_similarity(address1, address2):
    return fuzz.ratio(address1, address2)

def add_remarks(row):
    if pd.notna(row['Are you currently residing in Singapore']) and row['Are you currently residing in Singapore'].strip().lower() == 'no':
        if pd.notna(row['Country key']) and row['Country key'].strip().lower() == 'singapore':
            return "Mismatch as SAP has record of Singapore Address"
        else:
            return "Selected as not residing in Singapore"
    elif pd.notna(row['Are you currently residing in Singapore']) and row['Are you currently residing in Singapore'].strip().lower() == 'yes':
        if pd.notna(row['Country key']):
            country_key = row['Country key'].strip().lower()
            if country_key == 'malaysia':
                return "Mismatch as SAP is recorded as Malaysia Address"
            elif country_key == 'china':
                return "Mismatch as SAP is recorded as China Address"
    if row['Match']:
        return "Matched"
    elif pd.isnull(row['Full Address_sap']) or pd.isnull(row['Full Address_form']):
        return "Missing Address"
    else:
        return "Mismatched"

def process_files(sap_file, form_file):
    try:
        sap_df = pd.read_excel(sap_file)
        form_df = pd.read_excel(form_file)

        sap_df = sap_df.rename(columns={'Staff ID': 'Staff ID', 'Full Address': 'Full Address_sap', 'Country key': 'Country key'})
        form_df = form_df.rename(columns={'Staff ID': 'Staff ID', 'Full Address': 'Full Address_form', 'Are you currently residing in Singapore': 'Are you currently residing in Singapore'})

        merged_df = pd.merge(sap_df, form_df, on='Staff ID', how='outer')

        merged_df['Match'] = merged_df.apply(lambda row: address_similarity(preprocess_address(row['Full Address_sap']), preprocess_address(row['Full Address_form'])) > 60, axis=1)

        merged_df['Remarks'] = merged_df.apply(add_remarks, axis=1)

        output_filename = generate_output_filename()
        merged_df.to_excel(output_filename, columns=['Staff ID', 'Full Address_sap', 'Full Address_form', 'Remarks'], index=False)

        st.success(f"Processing complete. Results saved to {output_filename}")
        st.download_button(label="Download Output", data=open(output_filename, "rb"), file_name=os.path.basename(output_filename))

    except Exception as e:
        error_message = f"Error: {str(e)}"
        st.error(error_message)
        error_df = pd.DataFrame({'Error Remarks': [error_message]})
        error_df.to_excel(generate_output_filename(), index=False)

# Streamlit App UI
st.title("Address Matcher")

# File upload widgets
sap_file = st.file_uploader("Upload SAP Excel File", type="xlsx")
form_file = st.file_uploader("Upload Form Excel File", type="xlsx")

# Process button
if st.button("Process & Match"):
    if sap_file is not None and form_file is not None:
        process_files(sap_file, form_file)
    else:
        st.error("Please upload both SAP and Form files")
