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
st.set_page_config(layout="wide", page_title="Accuracy Check Tool")

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

    if excel_data is not None:
        headers = {'business date': None, 'inncode': None, 'sold': None, 'rev': None, 'revenue': None}
        row_start = None

        for label in headers.keys():
            headers[label] = find_header(label, excel_data)
            if headers[label]:
                if row_start is None or headers[label][0] > row_start:
                    row_start = headers[label][0]

        if not (headers['business date'] and (not inncode or headers['inncode']) and headers['sold'] and (headers['rev'] or headers['revenue'])):
            st.error("Could not find all required headers ('Business Date', 'Inncode', 'SOLD', 'Rev' or 'Revenue') in the first Excel file.")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

        op_data = pd.read_excel(repaired_excel_file, sheet_name=0, engine='openpyxl', header=row_start)
        op_data.columns = [col.lower().strip() for col in op_data.columns]

        if 'business date' not in op_data.columns or (inncode and 'inncode' not in op_data.columns):
            st.error("Expected columns 'Business Date' or 'Inncode' not found in the first Excel file.")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

        if inncode:
            filtered_data = op_data[op_data['inncode'] == inncode]
        else:
            filtered_data = op_data

        if filtered_data.empty:
            st.warning("No data found for the given Inncode in the first Excel file.")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

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
                'RN Percentage': rn_percentage / 100,  # Store as decimal for Excel
                'Juyo Rev': revnet,
                'Hilton Rev': rev_sum,
                'Rev Difference': rev_diff,
                'Rev Percentage': rev_percentage / 100  # Store as decimal for Excel
            })

        results_df = pd.DataFrame(results)

        past_accuracy_rn = results_df['RN Percentage'].mean() * 100  # Convert back to percentage for display
        past_accuracy_rev = results_df['Rev Percentage'].mean() * 100  # Convert back to percentage for display
    else:
        results_df, past_accuracy_rn, past_accuracy_rev = pd.DataFrame(), 0, 0

    if excel_data_2 is not None:
        headers_2 = {'occupancy date': None, 'occupancy on books this year': None, 'booked room revenue this year': None}
        row_start_2 = None

        for label in headers_2.keys():
            headers_2[label] = find_header(label, excel_data_2)
            if headers_2[label]:
                if row_start_2 is None or headers_2[label][0] > row_start_2:
                    row_start_2 = headers_2[label][0]

        if not all(headers_2.values()):
            st.error("Could not find all required headers ('Occupancy Date', 'Occupancy On Books This Year', 'Booked Room Revenue This Year') in the second Excel file.")
            return pd.DataFrame(), 0, 0, pd.DataFrame(), 0, 0

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
    else:
        future_results_df, future_accuracy_rn, future_accuracy_rev = pd.DataFrame(), 0, 0

    if not results_df.empty or not future_results_df.empty:
        accuracy_matrix = pd.DataFrame({
            'Metric': ['RNs', 'Revenue'],
            'Past': [f'{past_accuracy_rn:.2f}%', f'{past_accuracy_rev:.2f}%'] if not results_df.empty else ['N/A', 'N/A'],
            'Future': [f'{future_accuracy_rn:.2f}%', f'{future_accuracy_rev:.2f}%'] if not future_results_df.empty else ['N/A', 'N/A']
        })

        def color_scale(val):
            if isinstance(val, str) and '%' in val:
                val = float(val.strip('%'))
                if val >= 98:
                    return 'background-color: #469798'
                elif 95 <= val < 98:
                    return 'background-color: #F2A541'
                else:
                    return 'background-color: #BF3100'
            return ''

        accuracy_matrix_styled = accuracy_matrix.style.applymap(color_scale, subset=['Past', 'Future'])
        st.subheader(f'Accuracy Matrix for the hotel with code: {inncode}')
        st.dataframe(accuracy_matrix_styled, use_container_width=True)

    if not results_df.empty or not future_results_df.empty:
        st.subheader('RNs and Revenue Discrepancy Over Time')

        fig = make_subplots(specs=[[{"secondary_y": True}]])

        fig.add_trace(go.Bar(
            x=results_df['Business Date'] if not results_df.empty else future_results_df['Business Date'],
            y=results_df['RN Difference'] if not results_df.empty else future_results_df['RN Difference'],
            name='RNs Discrepancy',
            marker_color='#469798'
        ), secondary_y=False)

        fig.add_trace(go.Scatter(
            x=results_df['Business Date'] if not results_df.empty else future_results_df['Business Date'],
            y=results_df['Rev Difference'] if not results_df.empty else future_results_df['Rev Difference'],
            name='Revenue Discrepancy',
            mode='lines+markers',
            line=dict(color='#BF3100', width=2),
            marker=dict(size=8)
        ), secondary_y=True)

        max_room_discrepancy = results_df['RN Difference'].abs().max() if not results_df.empty else future_results_df['RN Difference'].abs().max()
        max_revenue_discrepancy = results_df['Rev Difference'].abs().max() if not results_df.empty else future_results_df['Rev Difference'].abs().max()

        fig.update_layout(
            height=600,
            title='RNs and Revenue Discrepancy Over Time',
            xaxis_title='Date',
            yaxis_title='RNs Discrepancy',
            yaxis2_title='Revenue Discrepancy',
            yaxis=dict(range=[-max_room_discrepancy, max_room_discrepancy]),
            yaxis2=dict(range=[-max_revenue_discrepancy, max_revenue_discrepancy]),
            legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01)
        )

        fig.update_yaxes(matches=None, showgrid=True, gridwidth=1, gridcolor='grey')

        st.plotly_chart(fig, use_container_width=True)

    if not results_df.empty:
        st.subheader('Detailed Accuracy Comparison (Past)')
        past_styled = results_df.style.applymap(lambda val: color_scale(val), subset=['RN Percentage', 'Rev Percentage'])
        st.dataframe(past_styled, use_container_width=True)

    if not future_results_df.empty:
        st.subheader('Detailed Accuracy Comparison (Future)')
        future_styled = future_results_df.style.applymap(lambda val: color_scale(val), subset=['RN Percentage', 'Rev Percentage'])
        st.dataframe(future_styled, use_container_width=True)

    return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

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
            # Ensure percentage columns are properly formatted as decimals
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

            # Apply simplified conditional formatting to percentages in columns E and I
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
            # Ensure percentage columns are properly formatted as decimals
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

            # Apply simplified conditional formatting to percentages in columns E and I
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
    output.seek(0)
    return output, base_filename

st.title('Accuracy Check Tool')

csv_file = st.file_uploader("Upload Daily Totals Extract (.csv)", type="csv")
excel_file = st.file_uploader("Upload Operational Report (.xlsx)", type="xlsx")

if excel_file:
    inncode = st.text_input("Enter Inncode to process:", value="")
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

perspective_date = st.date_input("Enter perspective date (Date of the IDeaS file receipt):", value=datetime.now().date())

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
