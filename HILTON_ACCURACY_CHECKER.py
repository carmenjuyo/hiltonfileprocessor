import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import csv
import io
from io import BytesIO
import zipfile
import xlsxwriter
import os

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
            def find_headers_in_sheet(data, headers_to_find):
                found_headers = {}
                for col in data.columns[:26]:
                    for row in range(len(data)):
                        cell_value = str(data[col][row]).strip().lower()
                        for header in headers_to_find:
                            if header in cell_value:
                                found_headers[header] = (row, col)
                return found_headers

            headers_to_find = ['occupancy date', 'occupancy on books this year', 'booked room revenue this year']
            headers_found = find_headers_in_sheet(excel_data_2, headers_to_find)

            if len(headers_found) < len(headers_to_find):
                st.error("Could not find all required headers ('Occupancy Date', 'Occupancy On Books This Year', 'Booked Room Revenue This Year') in the second Excel file.")
                return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

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
                    'RN Percentage': rn_percentage / 100,
                    'Juyo Rev': revnet,
                    'IDeaS Rev': booked_revenue_sum,
                    'Rev Difference': rev_diff,
                    'Rev Percentage': rev_percentage / 100
                })

            future_results_df = pd.DataFrame(future_results)

            future_accuracy_rn = future_results_df['RN Percentage'].mean() * 100
            future_accuracy_rev = future_results_df['Rev Percentage'].mean() * 100
        except Exception as e:
            st.error(f"Error processing the second Excel file: {e}")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0
    else:
        future_results_df, future_accuracy_rn, future_accuracy_rev = pd.DataFrame(), 0, 0

    if not results_df.empty or not future_results_df.empty:
        accuracy_matrix = pd.DataFrame({
            'Metric': ['RNs', 'Revenue'],
            'Past': [f'{past_accuracy_rn:.2f}%', f'{past_accuracy_rev:.2f}%'] if not results_df.empty else ['N/A', 'N/A'],
            'Future': [f'{future_accuracy_rn:.2f}%', f'{future_accuracy_rev:.2f}%'] if not future_results_df.empty else ['N/A', 'N/A']
        })

    return future_results_df, future_accuracy_rn, future_accuracy_rev
