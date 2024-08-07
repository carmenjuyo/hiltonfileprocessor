import streamlit as st
import pandas as pd

# Function to process and compare data
def process_files(csv_file, excel_file, inncode):
    # Load CSV file
    csv_data = pd.read_csv(csv_file)
    
    # Load Excel file
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    op_data = excel_data.get('Sheet1')  # Assuming the relevant sheet is named 'Sheet1'

    # Filter Excel data by Inncode
    filtered_data = op_data[op_data['Inncode'] == inncode]

    # Group by Business Date
    grouped_data = filtered_data.groupby('Business Date').agg({'SOLD': 'sum', 'Rev': 'sum'}).reset_index()

    # Prepare comparison results
    results = []
    for _, row in csv_data.iterrows():
        business_date = row['Business Date']
        rn = row['RN']
        revnet = row['RevNET']

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
    results_df = process_files(csv_file, excel_file, inncode)
    st.write("Comparison Results:")
    st.dataframe(results_df)
    st.write("Accuracy Checks:")
    accuracy_check = results_df[['RN Difference', 'Rev Difference']].abs().sum()
    st.write(accuracy_check)
else:
    st.write("Please upload both files and enter an Inncode.")
