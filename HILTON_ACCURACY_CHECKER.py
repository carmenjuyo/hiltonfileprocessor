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

# Function to dynamically find headers and process data
def dynamic_process_files(csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate):
    # Define all possible headers and their mappings
    available_headers = {
        'property name': 'property_name',
        'day of week': 'day_of_week',
        'occupancy date': 'occupancy_date',
        'comparison date last year': 'comparison_date_last_year',
        'market segment': 'market_segment',
        'occupancy on books this year': 'occupancy_this_year',
        'occupancy on books last year actual': 'occupancy_last_year_actual',
        'booked room revenue this year': 'revenue_this_year',
        'booked room revenue last year actual': 'revenue_last_year_actual'
    }

    # Initialize return variables to prevent NameError
    results_df, past_accuracy_rn, past_accuracy_rev = pd.DataFrame(), 0, 0
    future_results_df, future_accuracy_rn, future_accuracy_rev = pd.DataFrame(), 0, 0

    csv_data = load_csv(csv_file)
    if csv_data.empty:
        st.warning("CSV file could not be processed. Please check the file and try again.")
        return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

    arrival_date_col = 'arrivalDate'
    rn_col = 'rn'
    revnet_col = 'revNet'

    if arrival_date_col not in csv_data.columns:
        st.error(f"Expected column '{arrival_date_col}' not found in CSV file.")
        return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

    csv_data[arrival_date_col] = pd.to_datetime(csv_data[arrival_date_col], errors='coerce')
    csv_data = csv_data.dropna(subset=[arrival_date_col])

    repaired_excel_file = repair_xlsx(excel_file) if excel_file else None
    repaired_excel_file_2 = repair_xlsx(excel_file_2) if excel_file_2 else None

    try:
        # Read the first Excel file (Operational Report)
        excel_data = pd.read_excel(repaired_excel_file, sheet_name=0, engine='openpyxl', header=None) if repaired_excel_file else None

        # Read the second Excel file (Market Segment)
        if repaired_excel_file_2:
            op_data_2 = pd.read_excel(repaired_excel_file_2, sheet_name="Market Segment", engine='openpyxl', header=0)
            op_data_2.columns = [available_headers.get(col.lower().strip(), col.lower().strip()) for col in op_data_2.columns]

            # Ensure required columns are present
            required_columns = ['occupancy_date', 'occupancy_this_year', 'revenue_this_year']
            missing_columns = [col for col in required_columns if col not in op_data_2.columns]
            if missing_columns:
                st.error(f"Missing required columns in the Market Segment data: {', '.join(missing_columns)}")
                return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

            # Convert dates and clean data
            op_data_2['occupancy_date'] = pd.to_datetime(op_data_2['occupancy_date'], errors='coerce')
            op_data_2 = op_data_2.dropna(subset=['occupancy_date'])

            # Apply VAT if needed
            if apply_vat:
                op_data_2['revenue_this_year'] = op_data_2['revenue_this_year'] / (1 + vat_rate / 100)

            # Process future data logic here...
            pass
    except Exception as e:
        st.error(f"Error processing Excel files: {e}")
        return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

    # Additional logic for past and future calculations here...

    return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

# Streamlit UI components
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
            st.success("Processing complete!")
