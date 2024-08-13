import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import csv
import io
import plotly.graph_objects as go

# Set Streamlit page configuration to wide layout and dark theme
st.set_page_config(layout="wide", page_title="Hilton Accuracy Check Tool")

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
            'RN Percentage': rn_percentage,
            'Juyo Rev': revnet,
            'Hilton Rev': rev_sum,
            'Rev Difference': rev_diff,
            'Rev Percentage': rev_percentage
        })

    results_df = pd.DataFrame(results)

    past_accuracy_rn = results_df['RN Percentage'].mean()
    past_accuracy_rev = results_df['Rev Percentage'].mean()

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
            'RN Percentage': rn_percentage,
            'Juyo Rev': revnet,
            'IDeaS Rev': booked_revenue_sum,
            'Rev Difference': rev_diff,
            'Rev Percentage': rev_percentage
        })

    future_results_df = pd.DataFrame(future_results)

    future_accuracy_rn = future_results_df['RN Percentage'].mean()
    future_accuracy_rev = future_results_df['Rev Percentage'].mean()

    st.subheader(f'Accuracy Matrix for the hotel with code: {inncode}')

    # Apply color coding to the accuracy matrix
    accuracy_matrix = pd.DataFrame({
        'Metric': ['RNs', 'Revenue'],
        'Past': [f'{past_accuracy_rn:.2f}%', f'{past_accuracy_rev:.2f}%'],
        'Future': [f'{future_accuracy_rn:.2f}%', f'{future_accuracy_rev:.2f}%']
    })

    def color_scale(val):
        """Color scale for percentages."""
        if isinstance(val, str) and '%' in val:
            val = float(val.strip('%'))
            if val >= 98:
                return 'background-color: green'
            elif 95 <= val < 98:
                return 'background-color: darkgoldenrod'
            else:
                return 'background-color: red'
        return ''

    accuracy_matrix_styled = accuracy_matrix.style.applymap(color_scale, subset=['Past', 'Future'])
    st.dataframe(accuracy_matrix_styled, use_container_width=True)

    # Plotting the discrepancy over time using Plotly
    st.subheader('RNs and Revenue Discrepancy Over Time')

    fig = go.Figure()

    # RN Discrepancy (Past)
    fig.add_trace(go.Scatter(
        x=results_df['Business Date'],
        y=results_df['RN Difference'],
        mode='lines+markers',
        name='RNs Discrepancy (Past)',
        line=dict(color='cyan'),
        marker=dict(color='cyan', size=8)
    ))

    # Revenue Discrepancy (Past)
    fig.add_trace(go.Scatter(
        x=results_df['Business Date'],
        y=results_df['Rev Difference'],
        mode='lines+markers',
        name='Revenue Discrepancy (Past)',
        line=dict(color='red'),
        marker=dict(color='red', size=8)
    ))

    # RN Discrepancy (Future)
    fig.add_trace(go.Scatter(
        x=future_results_df['Business Date'],
        y=future_results_df['RN Difference'],
        mode='lines+markers',
        name='RNs Discrepancy (Future)',
        line=dict(color='orange'),
        marker=dict(color='orange', size=8)
    ))

    # Revenue Discrepancy (Future)
    fig.add_trace(go.Scatter(
        x=future_results_df['Business Date'],
        y=future_results_df['Rev Difference'],
        mode='lines+markers',
        name='Revenue Discrepancy (Future)',
        line=dict(color='magenta'),
        marker=dict(color='magenta', size=8)
    ))

    fig.update_layout(
        template='plotly_dark',
        title='RNs and Revenue Discrepancy Over Time',
        xaxis_title='Date',
        yaxis_title='Discrepancy',
        legend=dict(
            x=0,
            y=1.1,
            orientation='h'
        ),
        hovermode='x unified'
    )

    st.plotly_chart(fig, use_container_width=True)

    # Display past and future results as tables after the graph
    st.subheader('Detailed Accuracy Comparison (Past and Future)')

    def selective_color_scale(val, subset):
        """Apply color scale only to percentage columns."""
        if subset in ['RN Percentage', 'Rev Percentage']:
            if val >= 98:
                return 'background-color: green'
            elif 95 <= val < 98:
                return 'background-color: darkgoldenrod'
            else:
                return 'background-color: red'
        return ''

    st.write("Past Comparison:")
    past_styled = results_df.style.applymap(lambda val: selective_color_scale(val, 'RN Percentage'), subset=['RN Percentage']).applymap(lambda val: selective_color_scale(val, 'Rev Percentage'), subset=['Rev Percentage'])
    st.dataframe(past_styled, use_container_width=True)

    st.write("Future Comparison:")
    future_styled = future_results_df.style.applymap(lambda val: selective_color_scale(val, 'RN Percentage'), subset=['RN Percentage']).applymap(lambda val: selective_color_scale(val, 'Rev Percentage'), subset=['Rev Percentage'])
    st.dataframe(future_styled, use_container_width=True)

    return results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev

# Streamlit app layout
st.title('Hilton Accuracy Check Tool')

csv_file = st.file_uploader("Upload Daily Totals Extract (.csv)", type="csv")
excel_file = st.file_uploader("Upload Operational Report (.xlsx)", type="xlsx")
excel_file_2 = st.file_uploader("Upload IDeaS Report (.xlsx)", type="xlsx")
inncode = st.text_input("Enter Inncode to process:", value="")

# VAT options
apply_vat = st.checkbox("Apply VAT deduction to IDeaS revenue?", value=False)
vat_rate = None
if apply_vat:
    vat_rate = st.number_input("Enter VAT rate (%)", min_value=0.0, value=0.0, step=0.1)

perspective_date = st.date_input("Enter perspective date (Date of the IDeaS file receipt):", value=datetime.now().date())

if st.button("Process"):
    if not csv_file or not excel_file or not excel_file_2 or not inncode:
        st.error("Please upload all files and enter the Inncode to process.")
    else:
        results_df, past_accuracy_rn, past_accuracy_rev, future_results_df, future_accuracy_rn, future_accuracy_rev = dynamic_process_files(
            csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate
        )
