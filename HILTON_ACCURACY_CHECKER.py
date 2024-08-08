import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import csv
import io

# Function to detect delimiter and load CSV file
def load_csv(file):
    # Read the content of the file
    content = file.read().decode('utf-8')
    # Use StringIO to simulate a file object
    file_obj = io.StringIO(content)
    # Read the first few lines to detect the delimiter
    sample = content[:1024]
    dialect = csv.Sniffer().sniff(sample)
    delimiter = dialect.delimiter

    # Load the CSV with the detected delimiter
    return pd.read_csv(file_obj, delimiter=delimiter)

# Function to dynamically find headers and process data
def dynamic_process_files(csv_file, excel_file, inncode):
    # Load CSV file with automatic delimiter detection
    csv_data = load_csv(csv_file)

    # Display CSV columns for inspection
    st.write("CSV Columns:")
    st.write(csv_data.columns)

    # Identify correct column names based on inspection
    arrival_date_col = 'arrivalDate'  # Adjust this based on actual column name
    rn_col = 'rn'                    # Adjust this based on actual column name
    revnet_col = 'revNet'            # Adjust this based on actual column name

    # Assuming the CSV file has the columns 'arrivalDate', 'rn', 'revNet'
    if arrival_date_col not in csv_data.columns:
        st.error(f"Expected column '{arrival_date_col}' not found in CSV file.")
        return pd.DataFrame()

    # Convert arrivalDate in CSV to datetime
    csv_data[arrival_date_col] = pd.to_datetime(csv_data[arrival_date_col])

    # Load Excel file using openpyxl and access the first sheet
    excel_data = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl', header=None)

    # Display available sheet names for debugging
    st.write("Excel Data Sample:")
    st.dataframe(excel_data.head(20))

    # Initialize variables to hold header indices
    headers = {'business date': None, 'inncode': None, 'sold': None, 'rev': None}
    row_start = None

    # Function to find the header row and column
    def find_header(label):
        for col in excel_data.columns:
            for row in range(len(excel_data)):
                cell_value = str(excel_data[col][row]).strip().lower()
                if label in cell_value:
                    return (row, col)
        return None

    # Search for each header
    for label in headers.keys():
        headers[label] = find_header(label)
        if headers[label]:
            st.write(f"'{label.capitalize()}' found at row {headers[label][0]} column {headers[label][1]}")
            if row_start is None or headers[label][0] > row_start:
                row_start = headers[label][0]

    # Check if all required headers were found
    if not all(headers.values()):
        st.error("Could not find all required headers ('Business Date', 'Inncode', 'SOLD', 'Rev').")
        return pd.DataFrame()

    # Extract data using the identified headers
    op_data = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl', header=row_start)

    # Rename columns to standard names
    op_data.columns = [col.lower().strip() for col in op_data.columns]

    # Display adjusted Excel columns for debugging
    st.write("Adjusted Excel Columns:")
    st.dataframe(op_data.head())

    # Ensure the key columns are present after manual adjustment
    if 'inncode' not in op_data.columns or 'business date' not in op_data.columns:
        st.error("Expected columns 'Inncode' or 'Business Date' not found in Excel file.")
        return pd.DataFrame()

    # Filter Excel data by Inncode
    filtered_data = op_data[op_data['inncode'] == inncode]

    # Check if filtering results in any data
    if filtered_data.empty:
        st.warning("No data found for the given Inncode.")
        return pd.DataFrame()

    # Convert business date in filtered data to datetime
    filtered_data['business date'] = pd.to_datetime(filtered_data['business date'])

    # Get yesterday's date
    yesterday = datetime.now() - timedelta(days=1)

    # Filter out future dates
    filtered_data = filtered_data[filtered_data['business date'] <= yesterday]
    csv_data = csv_data[csv_data[arrival_date_col] <= yesterday]

    # Find common dates in both files
    common_dates = set(csv_data[arrival_date_col]).intersection(set(filtered_data['business date']))

    # Group Excel data by Business Date
    grouped_data = filtered_data.groupby('business date').agg({'sold': 'sum', 'rev': 'sum'}).reset_index()

    # Prepare comparison results
    results = []
    for _, row in csv_data.iterrows():
        business_date = row[arrival_date_col]
        if business_date not in common_dates:
            continue  # Skip dates not common to both files
        rn = row[rn_col]
        revnet = row[revnet_col]

        # Find corresponding data in Excel
        excel_row = grouped_data[grouped_data['business date'] == business_date]
        if excel_row.empty:
            continue  # Skip mismatched dates

        sold_sum = excel_row['sold'].values[0]
        rev_sum = excel_row['rev'].values[0]

        # Calculate differences
        rn_diff = rn - sold_sum
        rev_diff = revnet - rev_sum

        # Calculate percentages
        rn_percentage = 100 - (abs(rn_diff) / rn) * 100 if rn != 0 else 100
        rev_percentage = 100 - (abs(rev_diff) / revnet) * 100 if revnet != 0 else 100

        # Append results
        results.append({
            'Business Date': business_date,
            'CSV RN': rn,
            'Excel SOLD Sum': sold_sum,
            'RN Difference': rn_diff,
            'RN Percentage': f"{rn_percentage:.2f}%",
            'CSV RevNET': revnet,
            'Excel Rev Sum': rev_sum,
            'Rev Difference': rev_diff,
            'Rev Percentage': f"{rev_percentage:.2f}%"
        })

    # Convert results to DataFrame
    results_df = pd.DataFrame(results)

    # Calculate past accuracy
    past_accuracy_rn = results_df['RN Percentage'].apply(lambda x: float(x.strip('%'))).mean()
    past_accuracy_rev = results_df['Rev Percentage'].apply(lambda x: float(x.strip('%'))).mean()

    return results_df, past_accuracy_rn, past_accuracy_rev

# Streamlit app layout
st.title("Operational and Revenue Report Comparison Tool")

# File uploads
st.sidebar.title("Upload Files")
csv_file = st.sidebar.file_uploader("Upload Daily Totals Extract CSV", type='csv')
excel_file = st.sidebar.file_uploader("Upload Operational Report Excel", type='xlsx')

# Inncode input
inncode = st.sidebar.text_input("Enter Inncode to process:")

# Process and display results
if csv_file and excel_file and inncode:
    st.write("Processing...")
    try:
        results_df, past_accuracy_rn, past_accuracy_rev = dynamic_process_files(csv_file, excel_file, inncode)
        if not results_df.empty:
            # Display only the comparison results and accuracy checks
            st.write("Comparison Results:")
            st.dataframe(results_df, height=600)

            st.write("Accuracy Checks:")
            accuracy_check = results_df[['RN Difference', 'Rev Difference']].abs().sum()
            st.write(f"RN Difference: {accuracy_check['RN Difference']}")
            st.write(f"Rev Difference: {accuracy_check['Rev Difference']}")

            st.write("Past Accuracy:")
            st.write(f"RN Percentage Accuracy: {past_accuracy_rn:.2f}%")
            st.write(f"Rev Percentage Accuracy: {past_accuracy_rev:.2f}%")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.write("Please upload both files and enter an Inncode.")
