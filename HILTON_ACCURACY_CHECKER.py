import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import csv
import io

# Function to detect delimiter and load CSV file
def load_csv(file):
    content = file.read().decode('utf-8')
    file_obj = io.StringIO(content)
    sample = content[:1024]
    dialect = csv.Sniffer().sniff(sample)
    delimiter = dialect.delimiter
    return pd.read_csv(file_obj, delimiter=delimiter)

# Function to dynamically find headers and process data
def dynamic_process_files(csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate):
    csv_data = load_csv(csv_file)
    arrival_date_col = 'arrivalDate'
    rn_col = 'rn'
    revnet_col = 'revNet'

    if arrival_date_col not in csv_data.columns:
        st.error(f"Expected column '{arrival_date_col}' not found in CSV file.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    csv_data[arrival_date_col] = pd.to_datetime(csv_data[arrival_date_col])

    try:
        excel_data = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl', header=None)
        excel_data_2 = pd.read_excel(excel_file_2, sheet_name="Market Segment", engine='openpyxl', header=None)
    except Exception as e:
        st.error(f"Error reading Excel files: {e}")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    headers = {'business date': None, 'inncode': None, 'sold': None, 'rev': None}
    headers_2 = {'occupancy date': None, 'occupancy on books this year': None, 'booked room revenue this year': None}
    row_start = None
    row_start_2 = None

    def find_header(label, data):
        for col in data.columns:
            for row in range(len(data)):
                cell_value = str(data[col][row]).strip().lower()
                if label in cell_value:
                    return (row, col)
        return None

    for label in headers.keys():
        headers[label] = find_header(label, excel_data)
        if headers[label]:
            if row_start is None or headers[label][0] > row_start:
                row_start = headers[label][0]

    for label in headers_2.keys():
        headers_2[label] = find_header(label, excel_data_2)
        if headers_2[label]:
            if row_start_2 is None or headers_2[label][0] > row_start_2:
                row_start_2 = headers_2[label][0]

    if not all(headers.values()):
        st.error("Could not find all required headers ('Business Date', 'Inncode', 'SOLD', 'Rev') in the first Excel file.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0
    if not all(headers_2.values()):
        st.error("Could not find all required headers ('Occupancy Date', 'Occupancy On Books This Year', 'Booked Room Revenue This Year') in the second Excel file.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    op_data = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl', header=row_start)
    op_data_2 = pd.read_excel(excel_file_2, sheet_name="Market Segment", engine='openpyxl', header=row_start_2)

    op_data.columns = [col.lower().strip() for col in op_data.columns]
    op_data_2.columns = [col.lower().strip() for col in op_data_2.columns]

    if 'inncode' not in op_data.columns or 'business date' not in op_data.columns:
        st.error("Expected columns 'Inncode' or 'Business Date' not found in the first Excel file.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0
    if 'occupancy date' not in op_data_2.columns or 'occupancy on books this year' not in op_data_2.columns or 'booked room revenue this year' not in op_data_2.columns:
        st.error("Expected columns 'Occupancy Date', 'Occupancy On Books This Year', or 'Booked Room Revenue This Year' not found in the second Excel file.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    filtered_data = op_data[op_data['inncode'] == inncode]

    if filtered_data.empty:
        st.warning("No data found for the given Inncode in the first Excel file.")
        return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

    filtered_data['business date'] = pd.to_datetime(filtered_data['business date'])
    op_data_2['occupancy date'] = pd.to_datetime(op_data_2['occupancy date'])

    if perspective_date:
        end_date = pd.to_datetime(perspective_date)
    else:
        end_date = datetime.now() - timedelta(days=1)

    filtered_data = filtered_data[filtered_data['business date'] <= end_date]
    csv_data_past = csv_data[csv_data[arrival_date_col] <= end_date]
    future_data = csv_data[csv_data[arrival_date_col] > end_date]
    future_data_2 = op_data_2[op_data_2['occupancy date'] > end_date]

    common_dates = set(csv_data_past[arrival_date_col]).intersection(set(filtered_data['business date']))
    future_common_dates = set(future_data[arrival_date_col]).intersection(set(future_data_2['occupancy date']))

    grouped_data = filtered_data.groupby('business date').agg({'sold': 'sum', 'rev': 'sum'}).reset_index()
    grouped_data_2 = future_data_2.groupby('occupancy date').agg({'occupancy on books this year': 'sum', 'booked room revenue this year': 'sum'}).reset_index()

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
        rev_sum = excel_row['rev'].values[0]

        rn_diff = rn - sold_sum
        rev_diff = revnet - rev_sum

        rn_percentage = 100 if rn == 0 else 100 - (abs(rn_diff) / rn) * 100
        rev_percentage = 100 if revnet == 0 else 100 - (abs(rev_diff) / revnet) * 100

        results.append({
            'Business Date': business_date,
            'Juyo RN': rn,
            'Hilton RN': sold_sum,
            'RN Difference': rn_diff,
            'RN Percentage': f"{rn_percentage:.2f}%",
            'Juyo Rev': revnet,
            'Hilton Rev': rev_sum,
            'Rev Difference': rev_diff,
            'Rev Percentage': f"{rev_percentage:.2f}%"
        })

    results_df = pd.DataFrame(results)

    past_accuracy_rn = results_df['RN Percentage'].apply(lambda x: float(x.strip('%'))).mean()
    past_accuracy_rev = results_df['Rev Percentage'].apply(lambda x: float(x.strip('%'))).mean()

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
            'Juyo RN': rn,
            'IDeaS RN': occupancy_sum,
            'RN Difference': rn_diff,
            'RN Percentage': f"{rn_percentage:.2f}%",
            'Juyo Rev': revnet,
            'IDeaS Rev': booked_revenue_sum,
            'Rev Difference': rev_diff,
            'Rev Percentage': f"{rev_percentage:.2f}%"
        })

    future_results_df = pd.DataFrame(future_results)

    future_accuracy_rn = future_results_df['RN Percentage'].apply(lambda x: float(x.strip('%'))).mean()
    future_accuracy_rev = future_results_df['Rev Percentage'].apply(lambda x: float(x.strip('%'))).mean()

    st.subheader('Comparison Results (Past):')
    st.dataframe(results_df)

    st.subheader('Comparison Results (Future):')
    st.dataframe(future_results_df)

    st.subheader('Accuracy Checks:')
    st.write(f"Past RN Accuracy: {past_accuracy_rn:.2f}%")
    st.write(f"Past Rev Accuracy: {past_accuracy_rev:.2f}%")
    st.write(f"Future RN Accuracy: {future_accuracy_rn:.2f}%")
    st.write(f"Future Rev Accuracy: {future_accuracy_rev:.2f}%")

    return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

# Streamlit app layout
st.title('Operational and Revenue Report Comparison Tool')

csv_file = st.file_uploader("Upload Daily Totals Extract CSV", type="csv")
excel_file = st.file_uploader("Upload Operational Report Excel", type="xlsx")
excel_file_2 = st.file_uploader("Upload Market Segment Excel", type="xlsx")
inncode = st.text_input("Enter Inncode to process:", value="")

# VAT options
apply_vat = st.checkbox("Apply VAT deduction to future revenue?")
vat_rate = None
if apply_vat:
    vat_rate = st.number_input("Enter VAT rate (%)", min_value=0.0, value=20.0, step=0.1)

perspective_date = st.date_input("Enter perspective date (optional):", value=datetime.now().date())

if st.button("Process"):
    if not csv_file or not excel_file or not excel_file_2 or not inncode:
        st.error("Please upload all files and enter the Inncode to process.")
    else:
        results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev = dynamic_process_files(
            csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate
        )
