import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import os


class FileProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Hilton File Processor")
        self.file_paths = []
        self.data_frames = []  # Define data_frames here to be used in multiple functions
        self.create_widgets()

    def create_widgets(self):
        self.root.configure(bg='#121212')

        # Header
        self.header_frame = tk.Frame(self.root, bg='#121212')
        self.header_frame.grid(row=0, column=0, columnspan=2, pady=10)
        self.header_label = tk.Label(self.header_frame, text="Hilton File Processor", font=('Helvetica', 18), fg='white', bg='#121212')
        self.header_label.pack()

        # Separating line under header
        self.separator = tk.Frame(self.root, height=2, bd=1, relief='sunken', bg='white')
        self.separator.grid(row=1, column=0, columnspan=2, sticky='we', padx=5, pady=5)

        # Left frame
        self.left_frame = tk.Frame(self.root, bg='#121212')
        self.left_frame.grid(row=2, column=0, padx=10, pady=10, sticky='nsew')

        # Right frame
        self.right_frame = tk.Frame(self.root, bg='#121212')
        self.right_frame.grid(row=2, column=1, padx=10, pady=10, sticky='nsew')

        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=1)

        # Left frame widgets
        self.option_var = tk.StringVar(value="upload")
        self.upload_radio = tk.Radiobutton(self.left_frame, text="Upload Files", variable=self.option_var, value="upload", command=self.toggle_input_method, bg='#121212', fg='white')
        self.upload_radio.grid(row=0, column=0, sticky='w', pady=5, padx=5)

        self.path_radio = tk.Radiobutton(self.left_frame, text="Specify Directory Path", variable=self.option_var, value="path", command=self.toggle_input_method, bg='#121212', fg='white')
        self.path_radio.grid(row=1, column=0, sticky='w', pady=5, padx=5)

        self.upload_button = tk.Button(self.left_frame, text="Upload", command=self.upload_files, bg='gray', fg='black')
        self.upload_button.grid(row=0, column=1, pady=5, padx=5, sticky='w')

        self.path_entry = tk.Entry(self.left_frame, width=30, bg='black', fg='white')
        self.path_entry.grid(row=1, column=1, pady=5, padx=5, sticky='we')

        self.separator1 = tk.Frame(self.left_frame, height=2, bd=1, relief='sunken', bg='white')
        self.separator1.grid(row=2, column=0, columnspan=2, sticky='we', padx=5, pady=5)

        self.process_button = tk.Button(self.left_frame, text="Process Raw Data", command=self.process_files, bg='gray', fg='black')
        self.process_button.grid(row=3, column=0, columnspan=2, pady=10)

        self.text_area_frame = tk.Frame(self.left_frame)
        self.text_area_frame.grid(row=4, column=0, columnspan=2, pady=10, padx=5, sticky='nsew')
        self.left_frame.grid_rowconfigure(4, weight=1)
        self.left_frame.grid_columnconfigure(0, weight=1)

        self.text_area = tk.Text(self.text_area_frame, wrap='none', bg='black', fg='white')
        self.text_area.pack(side='left', fill='both', expand=True)

        self.scrollbar_y = tk.Scrollbar(self.text_area_frame, orient='vertical', command=self.text_area.yview)
        self.scrollbar_y.pack(side='right', fill='y')
        self.text_area['yscrollcommand'] = self.scrollbar_y.set

        self.scrollbar_x = tk.Scrollbar(self.left_frame, orient='horizontal', command=self.text_area.xview)
        self.scrollbar_x.grid(row=5, column=0, columnspan=2, sticky='ew')
        self.text_area['xscrollcommand'] = self.scrollbar_x.set

        self.save_button = tk.Button(self.left_frame, text="Save to CSV", command=self.save_to_csv, bg='gray', fg='black')
        self.save_button.grid(row=6, column=0, columnspan=2, pady=10)

        # Right frame widgets
        self.filter_label = tk.Label(self.right_frame, text="Name Filter (e.g., LEDGER_Westmont):", bg='#121212', fg='white')
        self.filter_label.grid(row=0, column=0, pady=5, padx=5, sticky='w')

        self.filter_entry = tk.Entry(self.right_frame, width=30, bg='black', fg='white')
        self.filter_entry.grid(row=0, column=1, pady=5, padx=5, sticky='we')

        self.inncode_label = tk.Label(self.right_frame, text="Enter Inncode (optional):", bg='#121212', fg='white')
        self.inncode_label.grid(row=1, column=0, pady=5, padx=5, sticky='w')

        self.inncode_entry = tk.Entry(self.right_frame, width=30, bg='black', fg='white')
        self.inncode_entry.grid(row=1, column=1, pady=5, padx=5, sticky='we')

        self.separator2 = tk.Frame(self.right_frame, height=2, bd=1, relief='sunken', bg='white')
        self.separator2.grid(row=2, column=0, columnspan=2, sticky='we', padx=5, pady=5)

        self.room_revenue_button = tk.Button(self.right_frame, text="Process Room Revenue by Day", command=self.process_room_revenue, bg='gray', fg='black')
        self.room_revenue_button.grid(row=3, column=0, columnspan=2, pady=10)

        self.room_revenue_text_area_frame = tk.Frame(self.right_frame)
        self.room_revenue_text_area_frame.grid(row=4, column=0, columnspan=2, pady=10, padx=5, sticky='nsew')
        self.right_frame.grid_rowconfigure(4, weight=1)
        self.right_frame.grid_columnconfigure(0, weight=1)

        self.room_revenue_text_area = tk.Text(self.room_revenue_text_area_frame, wrap='none', bg='black', fg='white')
        self.room_revenue_text_area.pack(side='left', fill='both', expand=True)

        self.room_revenue_scrollbar_y = tk.Scrollbar(self.room_revenue_text_area_frame, orient='vertical', command=self.room_revenue_text_area.yview)
        self.room_revenue_scrollbar_y.pack(side='right', fill='y')
        self.room_revenue_text_area['yscrollcommand'] = self.room_revenue_scrollbar_y.set

        self.room_revenue_scrollbar_x = tk.Scrollbar(self.right_frame, orient='horizontal', command=self.room_revenue_text_area.xview)
        self.room_revenue_scrollbar_x.grid(row=5, column=0, columnspan=2, sticky='ew')
        self.room_revenue_text_area['xscrollcommand'] = self.room_revenue_scrollbar_x.set

        self.save_room_revenue_button = tk.Button(self.right_frame, text="Save to CSV", command=self.save_room_revenue_to_csv, bg='gray', fg='black')
        self.save_room_revenue_button.grid(row=6, column=0, columnspan=2, pady=10)

        self.toggle_input_method()

    def toggle_input_method(self):
        if self.option_var.get() == "upload":
            self.upload_button.grid(row=0, column=1, pady=5, padx=5, sticky='w')
            self.path_entry.grid_forget()
        else:
            self.upload_button.grid_forget()
            self.path_entry.grid(row=1, column=1, pady=5, padx=5, sticky='we')

    def upload_files(self):
        self.file_paths = filedialog.askopenfilenames(filetypes=[("JSON files", "*.json")])
        if self.file_paths:
            self.text_area.insert(tk.END, f"Selected files:\n{', '.join(self.file_paths)}\n")

    def process_files(self):
        filter_criteria = self.filter_entry.get().split(',')
        self.data_frames = []

        if self.option_var.get() == "path":
            directory_path = self.path_entry.get()
            self.file_paths = [os.path.join(directory_path, f) for f in os.listdir(directory_path) if f.endswith('.json')]

        for file_path in self.file_paths:
            file_name = os.path.basename(file_path)
            if any(criterion.strip() in file_name for criterion in filter_criteria):
                with open(file_path, 'r') as file:
                    data = json.load(file)
                    df = pd.json_normalize(data)
                    if 'extract_type' in df.columns and df['extract_type'][0] == 'LEDGER':
                        self.process_ledger_file(df)
                    elif 'extract_type' in df.columns and df['extract_type'][0] == 'STAY':
                        self.process_stay_file(df)

        self.display_data()

    def process_ledger_file(self, df):
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

    def display_data(self):
        self.text_area.delete('1.0', tk.END)
        if self.data_frames:
            self.merged_data = pd.concat(self.data_frames)
            inncode_filter = self.inncode_entry.get().strip()
            if inncode_filter:
                self.merged_data = self.merged_data[self.merged_data['Inncode'] == inncode_filter]
            self.text_area.insert(tk.END, self.merged_data.to_string())
        else:
            self.text_area.insert(tk.END, "No data matched the filter criteria.\n")

    def save_to_csv(self):
        if hasattr(self, 'merged_data') and not self.merged_data.empty:
            save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if save_path:
                self.merged_data.to_csv(save_path, index=False)
                messagebox.showinfo("Success", f"Data saved to {save_path}")
        else:
            messagebox.showwarning("Warning", "No data to save.")

    def process_room_revenue(self):
        filter_criteria = self.filter_entry.get().split(',')
        inncode_filter = self.inncode_entry.get().strip()
        room_revenue_data_frames = []

        if self.option_var.get() == "path":
            directory_path = self.path_entry.get()
            self.file_paths = [os.path.join(directory_path, f) for f in os.listdir(directory_path) if f.endswith('.json')]

        for file_path in self.file_paths:
            file_name = os.path.basename(file_path)
            if any(criterion.strip() in file_name for criterion in filter_criteria):
                with open(file_path, 'r') as file:
                    data = json.load(file)
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

            # Display in text area
            self.room_revenue_text_area.delete('1.0', tk.END)
            self.room_revenue_text_area.insert(tk.END, self.room_revenue_data.to_string(index=False))
        else:
            messagebox.showwarning("Warning", "No data matched the filter criteria or there is no room revenue data.")

    def save_room_revenue_to_csv(self):
        if hasattr(self, 'room_revenue_data') and not self.room_revenue_data.empty:
            save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if save_path:
                self.room_revenue_data.to_csv(save_path, index=False)
                messagebox.showinfo("Success", f"Room revenue data saved to {save_path}")
        else:
            messagebox.showwarning("Warning", "No room revenue data to save.")


if __name__ == "__main__":
    root = tk.Tk()
    app = FileProcessorApp(root)
    root.mainloop()
