import streamlit as st
import pandas as pd

# Function to dynamically find headers and process data
def dynamic_process_files(csv_file, excel_file, inncode):
    # Load CSV file
    csv_data = pd.read_csv(csv_file)

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

    # Group by Business Date
    grouped_data = filtered_data.groupby('business date').agg({'sold': 'sum', 'rev': 'sum'}).reset_index()

    # Prepare comparison results
    results = []
    for _, row in csv_data.iterrows():
        business_date = row['arrivalDate']  # Adjust to actual CSV column name
        rn = row['rn']                      # Adjust to actual CSV column name
        revnet = row['revNet']              # Adjust to actual CSV column name

        # Find corresponding data in Excel
        excel_row = grouped_data[grouped_data['business date'] == business_date]
        if not excel_row.empty:
            sold_sum = excel_row['sold'].values[0]
            rev_sum = excel_row['rev'].values[0]

            # Calculate differences
            rn_diff = rn - sold_sum
            rev_diff = revnet - rev_sum

            # Calculate percentages
            rn_percentage = (rn_diff / rn) * 100 if rn != 0 else 0
            rev_percentage = (rev_diff / revnet) * 100 if revnet != 0 else 0

            # Append results
            results.append({
                'Business Date': business_date,
                'CSV RN': rn,
                'Excel SOLD Sum': sold_sum,
                'RN Difference': rn_diff,
                'RN Percentage': rn_percentage,
                'CSV RevNET': revnet,
                'Excel Rev Sum': rev_sum,
                'Rev Difference': rev_diff,
                'Rev Percentage': rev_percentage
            })

    # Convert results to DataFrame
    results_df = pd.DataFrame(results)
    return results_df

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
        results_df = dynamic_process_files(csv_file, excel_file, inncode)
        if not results_df.empty:
            st.write("Comparison Results:")
            st.dataframe(results_df)
            st.write("Accuracy Checks:")
            accuracy_check = results_df[['RN Difference', 'Rev Difference']].abs().sum()
            st.write(accuracy_check)
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.write("Please upload both files and enter an Inncode.")
