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
    business_date_idx = None
    inncode_idx = None
    sold_idx = None
    rev_idx = None

    # Find headers dynamically
    for col in excel_data.columns:
        for row in range(len(excel_data)):
            cell_value = str(excel_data[col][row]).strip().lower()

            # Check for 'Business Date' header
            if business_date_idx is None and 'business date' in cell_value:
                business_date_idx = (row, col)
                st.write(f"'Business Date' found at row {row} column {col}")

            # Check for 'Inncode' header
            if inncode_idx is None and 'inncode' in cell_value:
                inncode_idx = (row, col)
                st.write(f"'Inncode' found at row {row} column {col}")

            # Check for 'SOLD' header
            if sold_idx is None and 'sold' in cell_value:
                sold_idx = (row, col)
                st.write(f"'SOLD' found at row {row} column {col}")

            # Check for 'Rev' header
            if rev_idx is None and 'rev' in cell_value:
                rev_idx = (row, col)
                st.write(f"'Rev' found at row {row} column {col}")

        # Stop if all headers have been found
        if business_date_idx and inncode_idx and sold_idx and rev_idx:
            break

    # Check if all required headers were found
    if not all([business_date_idx, inncode_idx, sold_idx, rev_idx]):
        st.error("Could not find all required headers ('Business Date', 'Inncode', 'SOLD', 'Rev').")
        return pd.DataFrame()

    # Extract data using the identified headers
    header_row = max(business_date_idx[0], inncode_idx[0], sold_idx[0], rev_idx[0])
    op_data = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl', header=header_row)

    # Display adjusted Excel columns for debugging
    st.write("Adjusted Excel Columns:")
    st.write(op_data.head())

    # Ensure column names match the actual file
    if 'Inncode' not in op_data.columns or 'Business Date' not in op_data.columns:
        st.error("Expected columns 'Inncode' or 'Business Date' not found in Excel file.")
        return pd.DataFrame()

    # Filter Excel data by Inncode
    filtered_data = op_data[op_data['Inncode'] == inncode]

    # Check if filtering results in any data
    if filtered_data.empty:
        st.warning("No data found for the given Inncode.")
        return pd.DataFrame()

    # Group by Business Date
    grouped_data = filtered_data.groupby('Business Date').agg({'SOLD': 'sum', 'Rev': 'sum'}).reset_index()

    # Prepare comparison results
    results = []
    for _, row in csv_data.iterrows():
        business_date = row['arrivalDate']  # Adjust to actual CSV column name
        rn = row['rn']                      # Adjust to actual CSV column name
        revnet = row['revNet']              # Adjust to actual CSV column name

        # Find corresponding data in Excel
        excel_row = grouped_data[grouped_data['Business Date'] == business_date]
        if not excel_row.empty:
            sold_sum = excel_row['SOLD'].values[0]
            rev_sum = excel_row['Rev'].values[0]

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
