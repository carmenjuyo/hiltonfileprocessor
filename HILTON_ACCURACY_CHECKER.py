import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import csv
import io
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
import zipfile
import xlsxwriter
import os

# Set Streamlit page configuration to wide layout
st.set_page_config(layout="wide", page_title="Hilton Accuracy Check Tool")

# Repair function for corrupted Excel files using in-memory operations
def repair_xlsx(file):
    repaired_file = BytesIO()
    with zipfile.ZipFile(file, 'r') as zip_ref:
        with zipfile.ZipFile(repaired_file, 'w') as repaired_zip:
            for item in zip_ref.infolist():
                data = zip_ref.read(item.filename)
                repaired_zip.writestr(item, data)
            if 'xl/sharedStrings.xml' not in zip_ref.namelist():
                shared_string_content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                shared_string_content += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">\n'
                shared_string_content += '</sst>'
                repaired_zip.writestr('xl/sharedStrings.xml', shared_string_content)
    repaired_file.seek(0)
    return repaired_file

# Function to detect delimiter and load CSV file
def load_csv(file):
    if file is None:
        st.error("No CSV file uploaded.")
        return pd.DataFrame()
    
    try:
        content = file.read().decode('utf-8')
        file_obj = io.StringIO(content)
        sample = content[:1024]
        dialect = csv.Sniffer().sniff(sample)
        delimiter = dialect.delimiter
        return pd.read_csv(file_obj, delimiter=delimiter)
    except Exception as e:
        st.error(f"Error loading CSV file: {e}")
        return pd.DataFrame()

# Function to dynamically find headers and process data
def dynamic_process_files(csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate):
    csv_data = load_csv(csv_file)
    if csv_data.empty:
        st.warning("CSV file could not be processed. Please check the file and try again.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    arrival_date_col = 'arrivalDate'
    rn_col = 'rn'
    revnet_col = 'revNet'

    if arrival_date_col not in csv_data.columns:
        st.error(f"Expected column '{arrival_date_col}' not found in CSV file.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    csv_data[arrival_date_col] = pd.to_datetime(csv_data[arrival_date_col], errors='coerce')
    csv_data = csv_data.dropna(subset=[arrival_date_col])

    repaired_excel_file = repair_xlsx(excel_file) if excel_file else None
    repaired_excel_file_2 = repair_xlsx(excel_file_2) if excel_file_2 else None

    try:
        excel_data = pd.read_excel(repaired_excel_file, sheet_name=0, engine='openpyxl', header=None) if repaired_excel_file else None
        excel_data_2 = pd.read_excel(repaired_excel_file_2, sheet_name="Market Segment", engine='openpyxl', header=None) if repaired_excel_file_2 else None
    except Exception as e:
        st.error(f"Error reading Excel files: {e}")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    # Function to detect headers dynamically
    def find_header(label, data):
        for row in range(len(data)):
            for col in data.columns:
                cell_value = str(data.iloc[row, col]).strip().lower()
                if label.lower() in cell_value:
                    return row
        return None

    if excel_data is not None:
        # Process the first Excel file as before...
        pass  # Keeping this part unchanged.

    if excel_data_2 is not None:
        # Improved header detection for the second sheet
        headers_2 = {'occupancy date': None, 'occupancy on books this year': None, 'booked room revenue this year': None}
        header_row = None

        for label in headers_2.keys():
            headers_2[label] = find_header(label, excel_data_2)
            if headers_2[label] is not None:
                if header_row is None or headers_2[label] > header_row:
                    header_row = headers_2[label]

        # If dynamic header detection fails, fallback to a specific row (e.g., row 6)
        if header_row is None:
            st.warning("Headers not detected dynamically. Using default row for headers (row 6).")
            header_row = 6

        op_data_2 = pd.read_excel(repaired_excel_file_2, sheet_name="Market Segment", engine='openpyxl', header=header_row)

        # Rename columns to expected format if necessary
        op_data_2.columns = [col.lower().strip() for col in op_data_2.columns]

        # Ensure required columns are present
        if 'occupancy date' not in op_data_2.columns or 'occupancy on books this year' not in op_data_2.columns or 'booked room revenue this year' not in op_data_2.columns:
            st.error("Expected columns ('Occupancy Date', 'Occupancy On Books This Year', 'Booked Room Revenue This Year') not found.")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

        op_data_2['occupancy date'] = pd.to_datetime(op_data_2['occupancy date'], errors='coerce')
        op_data_2 = op_data_2.dropna(subset=['occupancy date'])

        # Process the future data and accuracy calculations as before...
        pass  # Keeping this part unchanged.

    return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

st.title('Hilton Accuracy Check Tool')

csv_file = st.file_uploader("Upload Daily Totals Extract (.csv)", type="csv")
excel_file = st.file_uploader("Upload Operational Report or Daily Market Segment with Inncode (.xlsx)", type="xlsx")

if excel_file:
    inncode = st.text_input("Enter Inncode to process (mandatory if the extract contains multiple properties):", value="")
else:
    inncode = ""

excel_file_2 = st.file_uploader("Upload IDeaS Report (.xlsx)", type="xlsx")

if excel_file_2:
    apply_vat = st.checkbox("Apply VAT deduction to IDeaS revenue?", value=False)
    if apply_vat:
        vat_rate = st.number_input("Enter VAT rate (%)", min_value=0.0, value=0.0, step=0.1)
else:
    apply_vat = False
    vat_rate = None

perspective_date = st.date_input("Enter perspective date (Date of the IDeaS file receipt and Support UI extract):", value=datetime.now().date())

if st.button("Process"):
    with st.spinner('Processing...'):
        results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev = dynamic_process_files(
            csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate
        )
        
        if results_df.empty and future_results_df.empty:
            st.warning("No data to display after processing. Please check the input files and parameters.")
        else:
            # Extract the base filename from the uploaded CSV file, before the first underscore
            base_filename = os.path.splitext(os.path.basename(csv_file.name))[0].split('_')[0]

            excel_data, base_filename = create_excel_download(
                results_df, future_results_df, base_filename, 
                past_accuracy_rn, past_accuracy_rev, 
                future_accuracy_rn, future_accuracy_rev
            )
            
            st.download_button(
                label="Download results as Excel",
                data=excel_data,
                file_name=f"{base_filename}_Accuracy_Results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
