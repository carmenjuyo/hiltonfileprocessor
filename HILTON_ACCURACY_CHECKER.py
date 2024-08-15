import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import csv
import io
import plotly.graph_objects as go
from io import BytesIO
import zipfile

# Set Streamlit page configuration to wide layout and dark theme
st.set_page_config(layout="wide", page_title="Hilton Accuracy Check Tool")

# Inject custom CSS to change the icon colors
st.markdown(
    """
    <style>
    /* Make the cloud upload icons cyan */
    .stFileUpload > label div[data-testid="fileUploadDropzone"] svg {
        color: cyan !important;
    }

    /* Make the file icons green */
    .stFileUploadDisplay > div:first-child > svg {
        color: #469798 !important;
    } 
    </style>
    """,
    unsafe_allow_html=True,
)

# Repair function for corrupted Excel files using in-memory operations
def repair_xlsx(file):
    repaired_file = BytesIO()
    with zipfile.ZipFile(file, 'r') as zip_ref:
        with zipfile.ZipFile(repaired_file, 'w') as repaired_zip:
            for item in zip_ref.infolist():
                data = zip_ref.read(item.filename)
                repaired_zip.writestr(item, data)
            # Check and add sharedStrings.xml if missing
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

# Upload files
uploaded_csv = st.file_uploader("Upload CSV", type=["csv"])
uploaded_operational_report = st.file_uploader("Upload Operational Report", type=["xlsx"])
uploaded_ideas_report = st.file_uploader("Upload IDEAs Report", type=["xlsx"])

# Function to process past data and generate reports
def process_past_data(csv_file, operational_report):
    st.write("Processing Past Data...")
    # Add your logic here to process past data, generate variance table, accuracy graph, and summary
    st.write("Past Data Variance Table")
    st.write("Past Accuracy Graph")
    st.write("Past Summary")

# Function to process future data and generate reports
def process_future_data(csv_file, ideas_report):
    st.write("Processing Future Data...")
    # Add your logic here to process future data, generate variance table, accuracy graph, and summary
    st.write("Future Data Variance Table")
    st.write("Future Accuracy Graph")
    st.write("Future Summary")

# Check which files are uploaded
if uploaded_csv is not None and uploaded_operational_report is not None and uploaded_ideas_report is None:
    # Only Past Data
    process_past_data(uploaded_csv, uploaded_operational_report)
    
elif uploaded_csv is not None and uploaded_operational_report is None and uploaded_ideas_report is not None:
    # Only Future Data
    process_future_data(uploaded_csv, uploaded_ideas_report)
    
elif uploaded_csv is not None and uploaded_operational_report is not None and uploaded_ideas_report is not None:
    # Both Past and Future Data
    process_past_data(uploaded_csv, uploaded_operational_report)
    process_future_data(uploaded_csv, uploaded_ideas_report)
    
else:
    st.write("Please upload the required files.")

df, use_container_width=True
        st.write(f"Past RN Accuracy: {past_accuracy_rn:.2f}%")
        st.write(f"Past Revenue Accuracy: {past_accuracy_rev:.2f}%")

    if not future_results_df.empty:
        st.write("Future Data Variance Table")
        st.dataframe(future_results_df, use_container_width=True)
        st.write(f"Future RN Accuracy: {future_accuracy_rn:.2f}%")
        st.write(f"Future Revenue Accuracy: {future_accuracy_rev:.2f}%")


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
    with st.spinner('Processing...'):
        if csv_file and excel_file and not excel_file_2:
            display_past_analysis(csv_file, excel_file, inncode, perspective_date, apply_vat, vat_rate)
        elif csv_file and not excel_file and excel_file_2:
            display_future_analysis(csv_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate)
        elif csv_file and excel_file and excel_file_2:
            display_both_analyses(csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate)
        else:
            st.warning("Please upload the necessary files for analysis.")

        st.dataframe(past_results_df, use_container_width=True)
        st.write(f"Past RN Accuracy: {past_accuracy_rn:.2f}%")
        st.write(f"Past Revenue Accuracy: {past_accuracy_rev:.2f}%")

    if not future_results_df.empty:
        st.write("Future Data Variance Table")
        st.dataframe(future_results_df, use_container_width=True)
        st.write(f"Future RN Accuracy: {future_accuracy_rn:.2f}%")
        st.write(f"Future Revenue Accuracy: {future_accuracy_rev:.2f}%")

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
    with st.spinner('Processing...'):
        if csv_file and excel_file and not excel_file_2:
            display_past_analysis(csv_file, excel_file, inncode, perspective_date, apply_vat, vat_rate)
        elif csv_file and not excel_file and excel_file_2:
            display_future_analysis(csv_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate)
        elif csv_file and excel_file and excel_file_2:
            display_both_analyses(csv_file, excel_file, excel_file_2, inncode, perspective_date, apply_vat, vat_rate)
        else:
            st.warning("Please upload the necessary files for analysis.")
