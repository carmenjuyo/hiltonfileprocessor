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

# Main processing function
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

    results_df, past_accuracy_rn, past_accuracy_rev = pd.DataFrame(), 0, 0
    future_results_df, future_accuracy_rn, future_accuracy_rev = pd.DataFrame(), 0, 0

    # Past Data Processing
    if excel_data is not None:
        # Process and compute past data accuracy...
        # Example: Apply your logic here for past data
        pass  # Replace with actual processing logic

    # Future Data Processing
    if excel_data_2 is not None:
        try:
            header_row_index = 6  # Assuming the headers are in row 7 (index 6)
            op_data_2 = pd.read_excel(repaired_excel_file_2, sheet_name="Market Segment", engine="openpyxl", header=None)
            op_data_2.columns = op_data_2.iloc[header_row_index]  # Set headers
            op_data_2 = op_data_2.iloc[header_row_index + 1:]  # Skip the header row

            op_data_2 = op_data_2[['Occupancy Date', 'Occupancy On Books This Year', 'Booked Room Revenue This Year']]
            op_data_2.columns = ['occupancy date', 'occupancy on books this year', 'booked room revenue this year']
            op_data_2['occupancy date'] = pd.to_datetime(op_data_2['occupancy date'], errors='coerce')
            op_data_2 = op_data_2.dropna(subset=['occupancy date'])

            # Filter for perspective date
            end_date = pd.to_datetime(perspective_date)
            future_data = csv_data[csv_data[arrival_date_col] > end_date]
            future_data_2 = op_data_2[op_data_2['occupancy date'] > end_date]

            grouped_data_2 = future_data_2.groupby('occupancy date').agg({
                'occupancy on books this year': 'sum',
                'booked room revenue this year': 'sum'
            }).reset_index()

            future_results = []
            for _, row in future_data.iterrows():
                occupancy_date = row[arrival_date_col]
                excel_row = grouped_data_2[grouped_data_2['occupancy date'] == occupancy_date]
                if not excel_row.empty:
                    # Calculate differences and percentages
                    pass  # Replace with your logic for RN/Rev differences

            future_results_df = pd.DataFrame(future_results)
        except Exception as e:
            st.error(f"Error processing future data: {e}")

    return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

# Function to apply color scale
def color_scale(val):
    if isinstance(val, str) and '%' in val:
        val = float(val.strip('%'))
        if val >= 98:
            return 'background-color: #469798'  # Green
        elif 95 <= val < 98:
            return 'background-color: #F2A541'  # Yellow
        else:
            return 'background-color: #BF3100'  # Red
    return ''

# Display accuracy matrix for past and future data
if not results_df.empty or not future_results_df.empty:
    accuracy_matrix = pd.DataFrame({
        'Metric': ['RNs', 'Revenue'],
        'Past': [f'{past_accuracy_rn:.2f}%', f'{past_accuracy_rev:.2f}%'] if not results_df.empty else ['N/A', 'N/A'],
        'Future': [f'{future_accuracy_rn:.2f}%', f'{future_accuracy_rev:.2f}%'] if not future_results_df.empty else ['N/A', 'N/A']
    })

    accuracy_matrix_styled = accuracy_matrix.style.applymap(color_scale, subset=['Past', 'Future'])
    st.subheader(f'Accuracy Matrix for the hotel with code: {inncode}')
    st.dataframe(accuracy_matrix_styled, use_container_width=True)

# Update display for past results with percentage formatting and color coding
if not results_df.empty:
    st.subheader('Detailed Accuracy Comparison (Past)')

    def color_scale(val):
        if val >= 0.98:
            color = '#469798'  # Green
        elif 0.96 <= val < 0.98:
            color = '#F2A541'  # Yellow
        else:
            color = '#BF3100'  # Red
        return f'background-color: {color}'

    past_styled = results_df.style.format({
        'RN Percentage': '{:.2%}',
        'Rev Percentage': '{:.2%}'
    }).applymap(color_scale, subset=['RN Percentage', 'Rev Percentage'])

    st.dataframe(past_styled, use_container_width=True)

# Update display for future results with percentage formatting and color coding
if not future_results_df.empty:
    st.subheader('Detailed Accuracy Comparison (Future)')

    future_styled = future_results_df.style.format({
        'RN Percentage': '{:.2%}',
        'Rev Percentage': '{:.2%}'
    }).applymap(color_scale, subset=['RN Percentage', 'Rev Percentage'])

    st.dataframe(future_styled, use_container_width=True)

# Function to create Excel file for download with color formatting and accuracy matrix
def create_excel_download(results_df, future_results_df, base_filename, past_accuracy_rn, past_accuracy_rev, future_accuracy_rn, future_accuracy_rev):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Write the Accuracy Matrix
        accuracy_matrix = pd.DataFrame({
            'Metric': ['RNs', 'Revenue'],
            'Past': [past_accuracy_rn / 100, past_accuracy_rev / 100],  # Store as decimal
            'Future': [future_accuracy_rn / 100, future_accuracy_rev / 100]  # Store as decimal
        })
        
        accuracy_matrix.to_excel(writer, sheet_name='Accuracy Matrix', index=False, startrow=1)
        worksheet = writer.sheets['Accuracy Matrix']

        # Define formats
        format_green = workbook.add_format({'bg_color': '#469798', 'font_color': '#FFFFFF'})
        format_yellow = workbook.add_format({'bg_color': '#F2A541', 'font_color': '#FFFFFF'})
        format_red = workbook.add_format({'bg_color': '#BF3100', 'font_color': '#FFFFFF'})
        format_percent = workbook.add_format({'num_format': '0.00%'})  # Percentage format

        # Apply percentage format to the relevant cells
        worksheet.set_column('B:C', None, format_percent)  # Set percentage format

        # Apply simplified conditional formatting for Accuracy Matrix
        worksheet.conditional_format('B3:B4', {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
        worksheet.conditional_format('B3:B4', {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
        worksheet.conditional_format('B3:B4', {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

        worksheet.conditional_format('C3:C4', {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
        worksheet.conditional_format('C3:C4', {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
        worksheet.conditional_format('C3:C4', {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

        # Write past and future results to separate sheets
        if not results_df.empty:
            results_df['RN Percentage'] = results_df['RN Percentage'].astype(float)
            results_df['Rev Percentage'] = results_df['Rev Percentage'].astype(float)

            results_df.to_excel(writer, sheet_name='Past Accuracy', index=False)
            worksheet_past = writer.sheets['Past Accuracy']

            # Define formats
            format_number = workbook.add_format({'num_format': '#,##0.00'})  # Floats
            format_whole = workbook.add_format({'num_format': '0'})  # Whole numbers
            format_percent = workbook.add_format({'num_format': '0.00%'})  # Percentage format

            # Format columns
            worksheet_past.set_column('A:A', None, format_whole)  # Whole numbers
            worksheet_past.set_column('C:C', None, format_whole)  # Whole numbers
            worksheet_past.set_column('D:D', None, format_whole)  # Whole numbers
            worksheet_past.set_column('F:F', None, format_number)  # Floats
            worksheet_past.set_column('G:G', None, format_number)  # Floats
            worksheet_past.set_column('H:H', None, format_number)  # Floats
            worksheet_past.set_column('E:E', None, format_percent)  # Percentage
            worksheet_past.set_column('I:I', None, format_percent)  # Percentage

            worksheet_past.conditional_format('E2:E{}'.format(len(results_df) + 1),
                                              {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
            worksheet_past.conditional_format('E2:E{}'.format(len(results_df) + 1),
                                              {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
            worksheet_past.conditional_format('E2:E{}'.format(len(results_df) + 1),
                                              {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

            worksheet_past.conditional_format('I2:I{}'.format(len(results_df) + 1),
                                              {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
            worksheet_past.conditional_format('I2:I{}'.format(len(results_df) + 1),
                                              {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
            worksheet_past.conditional_format('I2:I{}'.format(len(results_df) + 1),
                                              {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

        if not future_results_df.empty:
            future_results_df['RN Percentage'] = future_results_df['RN Percentage'].astype(float)
            future_results_df['Rev Percentage'] = future_results_df['Rev Percentage'].astype(float)

            future_results_df.to_excel(writer, sheet_name='Future Accuracy', index=False)
            worksheet_future = writer.sheets['Future Accuracy']

            # Format columns
            worksheet_future.set_column('A:A', None, format_whole)  # Whole numbers
            worksheet_future.set_column('C:C', None, format_whole)  # Whole numbers
            worksheet_future.set_column('D:D', None, format_whole)  # Whole numbers
            worksheet_future.set_column('F:F', None, format_number)  # Floats
            worksheet_future.set_column('G:G', None, format_number)  # Floats
            worksheet_future.set_column('H:H', None, format_number)  # Floats
            worksheet_future.set_column('E:E', None, format_percent)  # Percentage
            worksheet_future.set_column('I:I', None, format_percent)  # Percentage

            worksheet_future.conditional_format('E2:E{}'.format(len(future_results_df) + 1),
                                                {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
            worksheet_future.conditional_format('E2:E{}'.format(len(future_results_df) + 1),
                                                {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
            worksheet_future.conditional_format('E2:E{}'.format(len(future_results_df) + 1),
                                                {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

            worksheet_future.conditional_format('I2:I{}'.format(len(future_results_df) + 1),
                                                {'type': 'cell', 'criteria': '<', 'value': 0.96, 'format': format_red})
            worksheet_future.conditional_format('I2:I{}'.format(len(future_results_df) + 1),
                                                {'type': 'cell', 'criteria': 'between', 'minimum': 0.96, 'maximum': 0.9799, 'format': format_yellow})
            worksheet_future.conditional_format('I2:I{}'.format(len(future_results_df) + 1),
                                                {'type': 'cell', 'criteria': '>=', 'value': 0.98, 'format': format_green})

    output.seek(0)  # Make sure to seek to the beginning of the output
    return output, base_filename

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
            # Extract the base filename from the uploaded CSV file, before the first underscore
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
