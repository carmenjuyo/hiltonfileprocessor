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

    def find_header(label, data):
        for col in data.columns:
            for row in range(len(data)):
                cell_value = str(data[col][row]).strip().lower()
                if label in cell_value:
                    return (row, col)
        return None

    if excel_data_2 is not None:
        try:
            # Function to find headers dynamically up to column Z
            def find_headers_in_sheet(data, headers_to_find):
                found_headers = {}
                for col in data.columns[:26]:  # Restrict to columns A-Z
                    for row in range(len(data)):
                        cell_value = str(data[col][row]).strip().lower()
                        for header in headers_to_find:
                            if header in cell_value:
                                found_headers[header] = (row, col)
                return found_headers

            headers_to_find = ['occupancy date', 'occupancy on books this year', 'booked room revenue this year']
            headers_found = find_headers_in_sheet(excel_data_2, headers_to_find)

            # Check if all headers are found
            if len(headers_found) < len(headers_to_find):
                st.error("Could not find all required headers ('Occupancy Date', 'Occupancy On Books This Year', 'Booked Room Revenue This Year') in the second Excel file.")
                return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

            # Use the highest row number among the headers found as the header row
            row_start_2 = max([position[0] for position in headers_found.values()])
            op_data_2 = pd.read_excel(repaired_excel_file_2, sheet_name="Market Segment", engine='openpyxl', header=row_start_2)
            op_data_2.columns = [col.lower().strip() for col in op_data_2.columns]

            if 'occupancy date' not in op_data_2.columns or 'occupancy on books this year' not in op_data_2.columns or 'booked room revenue this year' not in op_data_2.columns:
                st.error("Expected columns 'Occupancy Date', 'Occupancy On Books This Year', or 'Booked Room Revenue This Year' not found in the second Excel file.")
                return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

            op_data_2['occupancy date'] = pd.to_datetime(op_data_2['occupancy date'], errors='coerce')
            op_data_2 = op_data_2.dropna(subset=['occupancy date'])

            if perspective_date:
                end_date = pd.to_datetime(perspective_date)
            else:
                end_date = datetime.now() - timedelta(days=1)

            future_data = csv_data[csv_data[arrival_date_col] > end_date]
            future_data_2 = op_data_2[op_data_2['occupancy date'] > end_date]

            future_common_dates = set(future_data[arrival_date_col]).intersection(set(future_data_2['occupancy date']))

            grouped_data_2 = future_data_2.groupby('occupancy date').agg({'occupancy on books this year': 'sum', 'booked room revenue this year': 'sum'}).reset_index()

            future_results = []
            for _, row in future_data.iterrows():
                occupancy_date = row[arrival_date_col]
                if occupancy_date not in future_common_dates:
                    continue
                rn = row[rn_col]
                revnet = row[revnet_col]

                excel_row = grouped_data_2[grouped_data_2['occupancy date'] == occupancy_date]
                if excel_row.empty:
                    continue

                occupancy_sum = excel_row['occupancy on books this year'].values[0]
                booked_revenue_sum = excel_row['booked room revenue this year'].values[0]

                if apply_vat:
                    booked_revenue_sum /= (1 + vat_rate / 100)

                rn_diff = rn - occupancy_sum
                rev_diff = revnet - booked_revenue_sum

                rn_percentage = 100 if rn == 0 else 100 - (abs(rn_diff) / rn) * 100
                rev_percentage = 100 if revnet == 0 else 100 - (abs(rev_diff) / revnet) * 100

                future_results.append({
                    'Business Date': occupancy_date,
                    'Juyo RN': int(rn),
                    'IDeaS RN': int(occupancy_sum),
                    'RN Difference': int(rn_diff),
                    'RN Percentage': rn_percentage / 100,  # Store as decimal for Excel
                    'Juyo Rev': revnet,
                    'IDeaS Rev': booked_revenue_sum,
                    'Rev Difference': rev_diff,
                    'Rev Percentage': rev_percentage / 100  # Store as decimal for Excel
                })

            future_results_df = pd.DataFrame(future_results)

            future_accuracy_rn = future_results_df['RN Percentage'].mean() * 100  # Convert back to percentage for display
            future_accuracy_rev = future_results_df['Rev Percentage'].mean() * 100  # Convert back to percentage for display
        except Exception as e:
            st.error(f"Error processing the second Excel file: {e}")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0
    else:
        future_results_df, future_accuracy_rn, future_accuracy_rev = pd.DataFrame(), 0, 0

# Function to create an Excel file for download
def create_excel_download(results_df, future_results_df, base_filename, past_accuracy_rn, past_accuracy_rev, future_accuracy_rn, future_accuracy_rev):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        accuracy_matrix = pd.DataFrame({
            'Metric': ['RNs', 'Revenue'],
            'Past': [past_accuracy_rn / 100, past_accuracy_rev / 100],
            'Future': [future_accuracy_rn / 100, future_accuracy_rev / 100]
        })
        accuracy_matrix.to_excel(writer, sheet_name='Accuracy Matrix', index=False, startrow=1)
        worksheet = writer.sheets['Accuracy Matrix']

        format_percent = workbook.add_format({'num_format': '0.00%'})
        worksheet.set_column('B:C', None, format_percent)

        if not future_results_df.empty:
            future_results_df.to_excel(writer, sheet_name='Future Accuracy', index=False)

    output.seek(0)
    return output, base_filename

# Streamlit app UI
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
