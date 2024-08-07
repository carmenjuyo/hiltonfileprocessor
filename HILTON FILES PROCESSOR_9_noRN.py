import streamlit as st
import pandas as pd
import json
import os

class FileProcessorApp:
    def __init__(self):
        self.file_paths = []
        self.data_frames = []
        self.merged_data = pd.DataFrame()
        self.room_revenue_data = pd.DataFrame()
    
    def display_header(self):
        st.title("Hilton File Processor")
    
    def upload_files(self):
        # Streamlit file uploader allows multiple files
        uploaded_files = st.file_uploader("Upload JSON files", type="json", accept_multiple_files=True)
        if uploaded_files:
            self.file_paths = uploaded_files
            st.success(f"Uploaded {len(uploaded_files)} files.")

    def process_files(self, filter_criteria, inncode_filter):
        self.data_frames = []
        
        # Iterate over uploaded files
        for uploaded_file in self.file_paths:
            # Read file content
            file_content = uploaded_file.read().decode("utf-8")
            data = json.loads(file_content)
            
            # Normalize JSON data to a DataFrame
            df = pd.json_normalize(data)
            
            # Check for extract_type and process accordingly
            if 'extract_type' in df.columns:
                if df['extract_type'][0] == 'LEDGER':
                    self.process_ledger_file(df)
                elif df['extract_type'][0] == 'STAY':
                    self.process_stay_file(df)

        self.display_data(filter_criteria, inncode_filter)

    def process_ledger_file(self, df):
        # Rename columns for the ledger file
        df.rename(columns={
            "account_id": "Account ID",
            "account_name": "Account Name",
            "accounting_category": "Accounting Category",
            "accounting_id": "Accounting ID",
            "accounting_id_desc": "Accounting ID Desc",
            "accounting_type": "Accounting Type",
            "business_date": "Business Date",
            "charge_routed": "Charge Routed",
            "common_account_identifier": "Common Account Identifier",
            "confirmation_number": "Confirmation Number",
            "crs_inn_code": "CRS Inn Code",
            "employee_id": "Employee ID",
            "entry_currency_code": "Entry Currency Code",
            "entry_datetime": "Entry Datetime",
            "entry_id": "Entry ID",
            "entry_type": "Entry Type",
            "exchange_rate": "Exchange Rate",
            "extract_type": "Extract Type",
            "facility_id": "Facility ID",
            "foreign_amount": "Foreign Amount",
            "gl_account_id": "GL Account ID",
            "gnr": "GNR",
            "hhonors_receipt_ind": "HHonors Receipt Ind",
            "include_in_net_use": "Include in Net Use",
            "inncode": "Inncode",
            "insert_datetime_utc": "Insert Datetime UTC",
            "ledger_entry_amount": "Ledger Entry Amount",
            "original_folio_id": "Original Folio ID",
            "original_receipt_id": "Original Receipt ID",
            "original_stay_id": "Original Stay ID",
            "partition_date": "Partition Date",
            "pms_inn_code": "PMS Inn Code",
            "posting_type_code": "Posting Type Code",
            "rate_plan_id": "Rate Plan ID",
            "rate_plan_type": "Rate Plan Type",
            "receipt_id": "Receipt ID",
            "routed_to_folio": "Routed to Folio",
            "stay_id": "Stay ID",
            "trans_desc": "Trans Desc",
            "trans_id": "Trans ID",
            "version": "Version",
            "charge_category": "Charge Category",
            "group_key": "Group Key",
            "group_name": "Group Name",
            "trans_travel_reason_code": "Trans Travel Reason Code",
            "ar_account_key": "AR Account Key",
            "ar_account_id": "AR Account ID",
            "ar_description": "AR Description",
            "ar_code": "AR Code",
            "ar_type_code": "AR Type Code",
            "ar_type_sub_code": "AR Type Sub Code",
            "house_key": "House Key"
        }, inplace=True)
        self.data_frames.append(df)

    def process_stay_file(self, df):
        # Rename columns for the stay file
        df.rename(columns={
            "account_id": "Account ID",
            "account_name": "Account Name",
            "arrival_date": "Arrival Date",
            "booked_date": "Booked Date",
            "booked_datetime": "Booked Datetime",
            "booking_segment_number": "Booking Segment Number",
            "confirmation_number": "Confirmation Number",
            "crs_inn_code": "CRS Inn Code",
            "departure_date": "Departure Date",
            "extract_type": "Extract Type",
            "facility_id": "Facility ID",
            "filename": "Filename",
            "gnr": "GNR",
            "guarantee_type_code": "Guarantee Type Code",
            "guarantee_type_text": "Guarantee Type Text",
            "inncode": "Inncode",
            "insert_datetime_utc": "Insert Datetime UTC",
            "mcat_code": "MCAT Code",
            "no_show_ind": "No Show Ind",
            "number_of_adults": "Number of Adults",
            "old_transaction_datetime_utc": "Old Transaction Datetime UTC",
            "originating_reservation_center": "Originating Reservation Center",
            "partition_by_date_id": "Partition by Date ID",
            "partition_date": "Partition Date",
            "prop_crs_room_rate": "Prop CRS Room Rate",
            "prop_currency_code": "Prop Currency Code",
            "reservation_status": "Reservation Status",
            "room_type_code": "Room Type Code",
            "srp_code": "SRP Code",
            "srp_name": "SRP Name",
            "srp_type": "SRP Type",
            "stay_date": "Stay Date",
            "tax_calculation_type": "Tax Calculation Type",
            "tax_included_ind": "Tax Included Ind",
            "transaction_datetime_utc": "Transaction Datetime UTC",
            "version": "Version"
        }, inplace=True)
        self.data_frames.append(df)

    def display_data(self, filter_criteria, inncode_filter):
        if self.data_frames:
            self.merged_data = pd.concat(self.data_frames)
            if inncode_filter:
                self.merged_data = self.merged_data[self.merged_data['Inncode'] == inncode_filter]
            st.dataframe(self.merged_data)
        else:
            st.warning("No data matched the filter criteria.")

    def save_to_csv(self):
        if not self.merged_data.empty:
            csv = self.merged_data.to_csv(index=False)
            st.download_button(
                label="Download data as CSV",
                data=csv,
                file_name='processed_data.csv',
                mime='text/csv',
            )
        else:
            st.warning("No data to save.")

    def process_room_revenue(self, filter_criteria, inncode_filter):
        room_revenue_data_frames = []

        # Iterate over uploaded files
        for uploaded_file in self.file_paths:
            file_content = uploaded_file.read().decode("utf-8")
            data = json.loads(file_content)

            df = pd.json_normalize(data)

            if 'extract_type' in df.columns and df['extract_type'][0] == 'LEDGER':
                df['ledger_entry_amount'] = pd.to_numeric(df['ledger_entry_amount'], errors='coerce')

                # Filter for revenue only
                revenue_filter = (df['charge_category'] == 'R') | (df['accounting_category'] == 'RA')
                df_filtered_revenue = df[revenue_filter]

                if inncode_filter:
                    df_filtered_revenue = df_filtered_revenue[df_filtered_revenue['inncode'] == inncode_filter]

                # Group by Business Date and Inncode
                df_agg_revenue = df_filtered_revenue.groupby(['business_date', 'inncode']).agg(
                    Ledger_Entry_Amount=('ledger_entry_amount', 'sum')
                ).reset_index()

                room_revenue_data_frames.append(df_agg_revenue)

        if room_revenue_data_frames:
            self.room_revenue_data = pd.concat(room_revenue_data_frames, ignore_index=True)

            # Sort and ensure unique Business Date and Inncode
            self.room_revenue_data['business_date'] = pd.to_datetime(self.room_revenue_data['business_date'])
            self.room_revenue_data.sort_values(by=['business_date', 'inncode'], inplace=True)

            # Deduplicate by Business Date and Inncode if necessary
            self.room_revenue_data = self.room_revenue_data.drop_duplicates(subset=['business_date', 'inncode'])

            st.dataframe(self.room_revenue_data)
        else:
            st.warning("No data matched the filter criteria or there is no room revenue data.")

    def save_room_revenue_to_csv(self):
        if not self.room_revenue_data.empty:
            csv = self.room_revenue_data.to_csv(index=False)
            st.download_button(
                label="Download room revenue data as CSV",
                data=csv,
                file_name='room_revenue_data.csv',
                mime='text/csv',
            )
        else:
            st.warning("No room revenue data to save.")


# Main Streamlit app
def main():
    app = FileProcessorApp()
    app.display_header()

    st.sidebar.title("Options")
    
    # Sidebar for uploading files
    app.upload_files()
    
    # Filter criteria input
    filter_criteria = st.sidebar.text_input("Name Filter (e.g., LEDGER_Westmont):")
    inncode_filter = st.sidebar.text_input("Enter Inncode (optional):")
    
    if st.sidebar.button("Process Raw Data"):
        app.process_files(filter_criteria, inncode_filter)

    if st.sidebar.button("Process Room Revenue by Day"):
        app.process_room_revenue(filter_criteria, inncode_filter)
        
    # Save buttons
    if st.sidebar.button("Save Processed Data to CSV"):
        app.save_to_csv()
        
    if st.sidebar.button("Save Room Revenue Data to CSV"):
        app.save_room_revenue_to_csv()

if __name__ == "__main__":
    main()
