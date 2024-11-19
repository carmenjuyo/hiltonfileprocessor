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

# Helper function to find a column with case-insensitive matching
def find_column(name, columns):
    for col in columns:
        if name.lower() in col.lower():
            return col  # Return the original column name
    return None

# Function to dynamically find headers and process data
def dynamic_process_files(csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate):
    csv_data = load_csv(csv_file)
    if csv_data.empty:
        st.warning("CSV file could not be processed. Please check the file and try again.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    arrival_date_col = find_column('arrivaldate', csv_data.columns)
    rn_col = find_column('rn', csv_data.columns)
    revnet_col = find_column('revnet', csv_data.columns)

    if not arrival_date_col or not rn_col or not revnet_col:
        st.error("Required columns ('arrivalDate', 'rn', 'revNet') not found in the CSV file.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    csv_data[arrival_date_col] = pd.to_datetime(csv_data[arrival_date_col], errors='coerce')
    csv_data = csv_data.dropna(subset=[arrival_date_col])

    repaired_excel_file = repair_xlsx(excel_file) if excel_file else None
    repaired_excel_file_2 = repair_xlsx(excel_file_2) if excel_file_2 else None

    try:
        excel_data = pd.read_excel(repaired_excel_file, sheet_name=0, engine='openpyxl') if repaired_excel_file else None
        excel_data_2 = pd.read_excel(repaired_excel_file_2, sheet_name="Market Segment", engine='openpyxl') if repaired_excel_file_2 else None
    except Exception as e:
        st.error(f"Error reading Excel files: {e}")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    if excel_data is not None:
        excel_data.columns = [col.lower().strip() for col in excel_data.columns]
        business_date_col = find_column('business date', excel_data.columns)
        sold_col = find_column('sold', excel_data.columns)
        rev_col = find_column('rev', excel_data.columns) or find_column('revenue', excel_data.columns)

        if not business_date_col or not sold_col or not rev_col:
            st.error("Required columns ('Business Date', 'Sold', 'Rev' or 'Revenue') not found in the first Excel file.")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

        excel_data[business_date_col] = pd.to_datetime(excel_data[business_date_col], errors='coerce')
        excel_data = excel_data.dropna(subset=[business_date_col])

        if perspective_date:
            end_date = pd.to_datetime(perspective_date)
        else:
            end_date = datetime.now() - timedelta(days=1)

        filtered_data = excel_data[excel_data[business_date_col] <= end_date]
        csv_data_past = csv_data[csv_data[arrival_date_col] <= end_date]

        grouped_data = filtered_data.groupby(business_date_col).agg({sold_col: 'sum', rev_col: 'sum'}).reset_index()
        results = []

        for _, row in csv_data_past.iterrows():
            business_date = row[arrival_date_col]
            rn = row[rn_col]
            revnet = row[revnet_col]

            excel_row = grouped_data[grouped_data[business_date_col] == business_date]
            if excel_row.empty:
                continue

            sold_sum = excel_row[sold_col].values[0]
            rev_sum = excel_row[rev_col].values[0]

            rn_diff = rn - sold_sum
            rev_diff = revnet - rev_sum

            rn_percentage = 100 if rn == 0 else 100 - (abs(rn_diff) / rn) * 100
            rev_percentage = 100 if revnet == 0 else 100 - (abs(rev_diff) / revnet) * 100

            results.append({
                'Business Date': business_date,
                'Juyo RN': int(rn),
                'Hilton RN': int(sold_sum),
                'RN Difference': int(rn_diff),
                'RN Percentage': rn_percentage / 100,
                'Juyo Rev': revnet,
                'Hilton Rev': rev_sum,
                'Rev Difference': rev_diff,
                'Rev Percentage': rev_percentage / 100
            })

        results_df = pd.DataFrame(results)
        past_accuracy_rn = results_df['RN Percentage'].mean() * 100
        past_accuracy_rev = results_df['Rev Percentage'].mean() * 100
    else:
        results_df, past_accuracy_rn, past_accuracy_rev = pd.DataFrame(), 0, 0

    if excel_data_2 is not None:
        excel_data_2.columns = [col.lower().strip() for col in excel_data_2.columns]
        occupancy_date_col = find_column('occupancy date', excel_data_2.columns)
        occupancy_books_col = find_column('occupancy on books this year', excel_data_2.columns)
        booked_revenue_col = find_column('booked room revenue this year', excel_data_2.columns)

        if not occupancy_date_col or not occupancy_books_col or not booked_revenue_col:
            st.error("Required columns ('Occupancy Date', 'Occupancy On Books This Year', 'Booked Room Revenue This Year') not found in the IDeaS file.")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

        excel_data_2[occupancy_date_col] = pd.to_datetime(excel_data_2[occupancy_date_col], errors='coerce')
        excel_data_2 = excel_data_2.dropna(subset=[occupancy_date_col])

        if perspective_date:
            end_date = pd.to_datetime(perspective_date)
        else:
            end_date = datetime.now() - timedelta(days=1)

        future_data = csv_data[csv_data[arrival_date_col] > end_date]
        future_data_2 = excel_data_2[excel_data_2[occupancy_date_col] > end_date]

        grouped_data_2 = future_data_2.groupby(occupancy_date_col).agg({
            occupancy_books_col: 'sum',
            booked_revenue_col: 'sum'
        }).reset_index()

        future_results = []

        for _, row in future_data.iterrows():
            occupancy_date = row[arrival_date_col]
            rn = row[rn_col]
            revnet = row[revnet_col]

            excel_row = grouped_data_2[grouped_data_2[occupancy_date_col] == occupancy_date]
            if excel_row.empty:
                continue

            occupancy_sum = excel_row[occupancy_books_col].values[0]
            booked_revenue_sum = excel_row[booked_revenue_col].values[0]

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
    else:
        future_results_df, future_accuracy_rn, future_accuracy_rev = pd.DataFrame(), 0, 0

    return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev
