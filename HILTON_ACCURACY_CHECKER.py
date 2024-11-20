import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import csv
import io
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

# Function to find column headers dynamically in the second Excel file
def find_headers(sheet_data, required_headers):
    """
    Finds headers in a given DataFrame `sheet_data` by scanning row-wise up to column Z.
    Returns a dictionary mapping required headers to their actual columns in the DataFrame.
    """
    header_mapping = {header: None for header in required_headers}

    for col in sheet_data.columns[:26]:  # Iterate up to column Z
        for row in range(sheet_data.shape[0]):
            cell_value = str(sheet_data.iloc[row, col]).strip().lower()
            for header in required_headers:
                if header.lower() == cell_value:
                    header_mapping[header] = col
                    break

    # Ensure all headers were found
    missing_headers = [header for header, value in header_mapping.items() if value is None]
    if missing_headers:
        raise ValueError(f"Missing required columns in the Market Segment data: {', '.join(missing_headers)}")

    return header_mapping

# Function to dynamically process files
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
        excel_data_2 = pd.read_excel(repaired_excel_file_2, sheet_name="Market Segment", engine='openpyxl', header=None) if repaired_excel_file_2 else None
    except Exception as e:
        st.error(f"Error reading Excel files: {e}")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    # Process the "Market Segment" sheet
    if excel_data_2 is not None:
        try:
            required_headers = ['Occupancy Date', 'Occupancy On Books This Year', 'Booked Room Revenue This Year']
            header_mapping = find_headers(excel_data_2, required_headers)

            # Extract data starting from the header row
            op_data_2 = pd.read_excel(
                repaired_excel_file_2,
                sheet_name="Market Segment",
                engine='openpyxl',
                skiprows=header_mapping[required_headers[0]]
            )

            # Normalize column names
            op_data_2.columns = [col.lower().strip() for col in op_data_2.columns]
            op_data_2.rename(columns={
                header_mapping['Occupancy Date']: 'occupancy_date',
                header_mapping['Occupancy On Books This Year']: 'occupancy_this_year',
                header_mapping['Booked Room Revenue This Year']: 'revenue_this_year'
            }, inplace=True)

            # Convert date and filter data by perspective date
            op_data_2['occupancy_date'] = pd.to_datetime(op_data_2['occupancy_date'], errors='coerce')
            op_data_2 = op_data_2.dropna(subset=['occupancy_date'])
            if perspective_date:
                perspective_date = pd.to_datetime(perspective_date)
                op_data_2 = op_data_2[op_data_2['occupancy_date'] > perspective_date]

            # Apply VAT if needed
            if apply_vat:
                op_data_2['revenue_this_year'] /= (1 + vat_rate / 100)

            st.success("Market Segment sheet processed successfully.")
        except Exception as e:
            st.error(f"Error processing the 'Market Segment' sheet: {e}")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0
    else:
        st.warning("No data found in the 'Market Segment' sheet.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    # Return processed results
    # Simulating some calculations for demonstration purposes
    future_results_df = pd.DataFrame(op_data_2)
    past_accuracy_rn, past_accuracy_rev, future_accuracy_rn, future_accuracy_rev = 0, 0, 98.5, 95.6  # Example

    return pd.DataFrame(), past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

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
            st.success("Data processed successfully!")
            # Further processing like downloading Excel or displaying data can follow...
