
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import csv
import io
from io import BytesIO
import zipfile
import xlsxwriter

st.set_page_config(layout="wide", page_title="Hilton Accuracy Check Tool")

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

def find_header(label, data):
    for col in data.columns:
        for row in range(len(data)):
            cell_value = str(data[col][row]).strip().lower()
            if label in cell_value:
                return (row, col)
    return None

def find_header_excel2(label, data):
    max_cols = 26
    max_rows = len(data)
    for col in range(max_cols):
        for row in range(max_rows):
            cell_value = str(data.iloc[row, col]).strip().lower()
            if cell_value == label.lower():
                return (row, col)
    return None

def color_scale(val):
    if isinstance(val, str) and '%' in val:
        val = float(val.strip('%'))
        if val >= 98:
            return 'background-color: #469798'
        elif 95 <= val < 98:
            return 'background-color: #F2A541'
        else:
            return 'background-color: #BF3100'
    return ''

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

    if excel_data_2 is not None:
        try:
            headers_2 = {
                'occupancy date': None,
                'occupancy on books this year': None,
                'booked room revenue this year': None
            }
            row_start_2 = None

            for label in headers_2.keys():
                headers_2[label] = find_header_excel2(label, excel_data_2)
                if headers_2[label]:
                    if row_start_2 is None or headers_2[label][0] > row_start_2:
                        row_start_2 = headers_2[label][0]

            if not all(headers_2.values()):
                missing_headers = [label for label, pos in headers_2.items() if pos is None]
                st.error(f"Missing required headers in Excel File 2: {', '.join(missing_headers)}")
                return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

            col_mapping = {headers_2[label][1]: label.lower() for label in headers_2.keys()}
            data_rows = excel_data_2.iloc[row_start_2 + 1:].reset_index(drop=True)

            op_data_2 = pd.DataFrame()
            for col_index, col_name in col_mapping.items():
                op_data_2[col_name] = data_rows.iloc[:, col_index]

            op_data_2.columns = [col.lower().strip() for col in op_data_2.columns]
            op_data_2['occupancy date'] = pd.to_datetime(op_data_2['occupancy date'], errors='coerce')
            op_data_2 = op_data_2.dropna(subset=['occupancy date'])

        except Exception as e:
            st.error(f"Error processing Excel File 2: {e}")
            op_data_2 = pd.DataFrame()

    return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0


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
