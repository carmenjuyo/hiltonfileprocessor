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

    # Handle Excel file 2 with specific requirements
    if repaired_excel_file_2:
        try:
            excel_data_2 = pd.read_excel(repaired_excel_file_2, sheet_name="Market Segment", engine='openpyxl', header=None)
        except Exception as e:
            st.error(f"Error reading Excel file 2: {e}")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

        # Find headers in the specified sheet up to column Z
        headers_2 = ['occupancy date', 'occupancy on books this year', 'booked room revenue this year']
        header_locations = {}
        for col in range(26):  # Columns A to Z
            for row in range(len(excel_data_2)):
                value = str(excel_data_2.iloc[row, col]).strip().lower()
                if value in headers_2:
                    header_locations[value] = (row, col)
                    if len(header_locations) == len(headers_2):
                        break
            if len(header_locations) == len(headers_2):
                break

        if len(header_locations) != len(headers_2):
            st.error("Required headers not found in the 'Market Segment' sheet.")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

        # Determine the header row and process data
        header_row = max([loc[0] for loc in header_locations.values()])
        data_start = header_row + 1
        op_data_2 = pd.read_excel(repaired_excel_file_2, sheet_name="Market Segment", engine='openpyxl', skiprows=data_start)
        op_data_2.columns = [col.lower().strip() for col in op_data_2.columns]

        if not all(h in op_data_2.columns for h in headers_2):
            st.error("Required columns not found after processing the 'Market Segment' sheet.")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

        op_data_2['occupancy date'] = pd.to_datetime(op_data_2['occupancy date'], errors='coerce')
        op_data_2 = op_data_2.dropna(subset=['occupancy date'])

        if perspective_date:
            end_date = pd.to_datetime(perspective_date)
        else:
            end_date = datetime.now() - timedelta(days=1)

        # Process data (similar to original future processing logic)
        future_data = csv_data[csv_data[arrival_date_col] > end_date]
        future_data_2 = op_data_2[op_data_2['occupancy date'] > end_date]

        future_common_dates = set(future_data[arrival_date_col]).intersection(set(future_data_2['occupancy date']))

        grouped_data_2 = future_data_2.groupby('occupancy date').agg({
            'occupancy on books this year': 'sum',
            'booked room revenue this year': 'sum'
        }).reset_index()

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

    else:
        future_results_df, future_accuracy_rn, future_accuracy_rev = pd.DataFrame(), 0, 0

    # Process past data if the first Excel file is provided
    if repaired_excel_file:
        try:
            excel_data = pd.read_excel(repaired_excel_file, sheet_name=0, engine='openpyxl', header=None)
        except Exception as e:
            st.error(f"Error reading Excel file 1: {e}")
            return pd.DataFrame(), 0, 0, future_results_df, future_accuracy_rn, future_accuracy_rev

        headers = {'business date': None, 'inncode': None, 'sold': None, 'rev': None, 'revenue': None, 'hotel name': None}
        header_locations = {}
        for col in range(len(excel_data.columns)):
            for row in range(len(excel_data)):
                value = str(excel_data.iloc[row, col]).strip().lower()
                if value in headers and headers[value] is None:
                    header_locations[value] = (row, col)
                    headers[value] = (row, col)

        if not headers['business date'] or not headers['sold'] or not (headers['rev'] or headers['revenue']):
            st.error("Could not find all required headers ('Business Date', 'Sold', 'Rev' or 'Revenue') in the first Excel file.")
            return pd.DataFrame(), 0, 0, future_results_df, future_accuracy_rn, future_accuracy_rev

        header_row = max([loc[0] for loc in header_locations.values()])
        op_data = pd.read_excel(repaired_excel_file, sheet_name=0, engine='openpyxl', skiprows=header_row + 1)
        op_data.columns = [col.lower().strip() for col in op_data.columns]

        if 'business date' not in op_data.columns or (inncode and 'inncode' not in op_data.columns):
            st.error("Expected columns 'Business Date' or 'Inncode' not found in the first Excel file.")
            return pd.DataFrame(), 0, 0, future_results_df, future_accuracy_rn, future_accuracy_rev

        if inncode:
            filtered_data = op_data[op_data['inncode'] == inncode]
        else:
            filtered_data = op_data

        if 'hotel name' in filtered_data.columns:
            filtered_data = filtered_data[filtered_data['hotel name'].str.lower() != 'total']

        if filtered_data.empty:
            st.warning("No data found for the given Inncode in the first Excel file.")
            return pd.DataFrame(), 0, 0, future_results_df, future_accuracy_rn, future_accuracy_rev

        filtered_data['business date'] = pd.to_datetime(filtered_data['business date'], errors='coerce')
        filtered_data = filtered_data.dropna(subset=['business date'])

        if perspective_date:
            end_date = pd.to_datetime(perspective_date)
        else:
            end_date = datetime.now() - timedelta(days=1)

        filtered_data = filtered_data[filtered_data['business date'] <= end_date]
        csv_data_past = csv_data[csv_data[arrival_date_col] <= end_date]

        common_dates = set(csv_data_past[arrival_date_col]).intersection(set(filtered_data['business date']))

        rev_col = 'rev' if 'rev' in filtered_data.columns else 'revenue'
        grouped_data = filtered_data.groupby('business date').agg({'sold': 'sum', rev_col: 'sum'}).reset_index()

        results = []
        for _, row in csv_data_past.iterrows():
            business_date = row[arrival_date_col]
            if business_date not in common_dates:
                continue
            rn = row[rn_col]
            revnet = row[revnet_col]

            excel_row = grouped_data[grouped_data['business date'] == business_date]
            if excel_row.empty:
                continue

            sold_sum = excel_row['sold'].values[0]
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

    # Return both past and future results
    return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

# Function to create Excel file for download with color formatting and accuracy matrix
def create_excel_download(results_df, future_results_df, base_filename, past_accuracy_rn, past_accuracy_rev, future_accuracy_rn, future_accuracy_rev):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Write the Accuracy Matrix
        accuracy_matrix = pd.DataFrame({
            'Metric': ['RNs', 'Revenue'],
            'Past': [past_accuracy_rn / 100, past_accuracy_rev / 100],
            'Future': [future_accuracy_rn / 100, future_accuracy_rev / 100]
        })
        
        accuracy_matrix.to_excel(writer, sheet_name='Accuracy Matrix', index=False, startrow=1)
        worksheet = writer.sheets['Accuracy Matrix']

        format_percent = workbook.add_format({'num_format': '0.00%'})
        worksheet.set_column('B:C', None, format_percent)

        if not results_df.empty:
            results_df.to_excel(writer, sheet_name='Past Accuracy', index=False)
        if not future_results_df.empty:
            future_results_df.to_excel(writer, sheet_name='Future Accuracy', index=False)

    output.seek(0)
    return output, base_filename

st.title('Hilton Accuracy Check Tool')

csv_file = st.file_uploader("Upload Daily Totals Extract (.csv)", type="csv")
excel_file = st.file_uploader("Upload Operational Report or Daily Market Segment with Inncode (.xlsx)", type="xlsx")
excel_file_2 = st.file_uploader("Upload IDeaS Report (.xlsx)", type="xlsx")
inncode = st.text_input("Enter Inncode (if required):", value="")
perspective_date = st.date_input("Enter perspective date:", value=datetime.now().date())
apply_vat = st.checkbox("Apply VAT deduction to IDeaS revenue?")
vat_rate = st.number_input("Enter VAT rate:", min_value=0.0, value=0.0) if apply_vat else None

if st.button("Process"):
    results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev = dynamic_process_files(
        csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate
    )
    if not results_df.empty or not future_results_df.empty:
        excel_data, base_filename = create_excel_download(
            results_df, future_results_df, "Results", past_accuracy_rn, past_accuracy_rev, future_accuracy_rn, future_accuracy_rev
        )
        st.download_button("Download Excel", excel_data, file_name=f"{base_filename}.xlsx")
