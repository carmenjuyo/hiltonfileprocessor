import streamlit as st
import pandas as pd
from xml.etree import ElementTree
from datetime import datetime
import plotly.graph_objs as go
from plotly.subplots import make_subplots

# Set page layout to wide
st.set_page_config(layout="wide", page_title="Daily Variance and Accuracy Calculator")

# Define the function to parse the XML
def parse_xml(xml_content):
    tree = ElementTree.fromstring(xml_content)
    data = []
    system_time = datetime.strptime(tree.find('G_RESORT/SYSTEM_TIME').text, "%d-%b-%y %H:%M:%S")
    for g_considered_date in tree.iter('G_CONSIDERED_DATE'):
        date = g_considered_date.find('CONSIDERED_DATE').text
        ind_deduct_rooms = int(g_considered_date.find('IND_DEDUCT_ROOMS').text)
        grp_deduct_rooms = int(g_considered_date.find('GRP_DEDUCT_ROOMS').text)
        ind_deduct_revenue = float(g_considered_date.find('IND_DEDUCT_REVENUE').text)
        grp_deduct_revenue = float(g_considered_date.find('GRP_DEDUCT_REVENUE').text)
        date = datetime.strptime(date, "%d-%b-%y").date()  # Assuming date is in 'DD-MMM-YY' format
        data.append({
            'date': date,
            'system_time': system_time,
            'HF RNs': ind_deduct_rooms + grp_deduct_rooms,
            'HF Rev': ind_deduct_revenue + grp_deduct_revenue
        })
    return pd.DataFrame(data)

# Define color coding for accuracy values
def color_scale(val):
    """Color scale for percentages."""
    if isinstance(val, str) and '%' in val:
        val = float(val.strip('%'))
        if val >= 98:
            return 'background-color: #469798; color: white;'  # green
        elif 95 <= val < 98:
            return 'background-color: #F2A541; color: white;'  # yellow
        else:
            return 'background-color: #BF3100; color: white;'  # red
    return ''

# Streamlit application
    def main():
        # Center the title using markdown with HTML
        st.markdown("<h1 style='text-align: center;'>Daily Variance and Accuracy Calculator</h1>", unsafe_allow_html=True)

        # Note/warning box with matching colors
        st.warning("The reference date of the Daily Totals Extract should be equal to the History and Forecast file date.")

        # File uploaders in columns
        col1, col2 = st.columns(2)
        with col1:
            xml_files = st.file_uploader("Upload History and Forecast .xml", type=['xml'], accept_multiple_files=True)
        with col2:
            csv_file = st.file_uploader("Upload Daily Totals Extract from Support UI", type=['csv'])

        # When files are uploaded
        if xml_files and csv_file:
            # Process XML files and combine them into a single DataFrame
            combined_xml_df = pd.DataFrame()
            for xml_file in xml_files:
                xml_df = parse_xml(xml_file.getvalue())
                combined_xml_df = pd.concat([combined_xml_df, xml_df])

            # Keep only the latest entry for each considered date
            combined_xml_df = combined_xml_df.sort_values(by=['date', 'system_time'], ascending=[True, False])
            combined_xml_df = combined_xml_df.drop_duplicates(subset=['date'], keep='first')

            # Process CSV
            csv_df = pd.read_csv(csv_file, delimiter=';', quotechar='"')
            csv_df.columns = [col.replace('"', '').strip() for col in csv_df.columns]
            csv_df['arrivalDate'] = pd.to_datetime(csv_df['arrivalDate'], errors='coerce')
            csv_df['Juyo RN'] = csv_df['rn'].astype(int)
            csv_df['Juyo Rev'] = csv_df['revNet'].astype(float)

            # Ensure both date columns are of type datetime64[ns] before merging
            combined_xml_df['date'] = pd.to_datetime(combined_xml_df['date'], errors='coerce')
            csv_df['arrivalDate'] = pd.to_datetime(csv_df['arrivalDate'], errors='coerce')

            # Merge data
            merged_df = pd.merge(combined_xml_df, csv_df, left_on='date', right_on='arrivalDate')

            # Calculate discrepancies for rooms and revenue
            merged_df['RN Diff'] = merged_df['HF RNs'] - merged_df['Juyo RN']
            merged_df['Rev Diff'] = merged_df['HF Rev'] - merged_df['Juyo Rev']

            # Calculate accuracies
            current_date = pd.to_datetime('today').normalize()  # Get the current date without the time part
            past_mask = merged_df['date'] < current_date
            future_mask = merged_df['date'] >= current_date

            past_rooms_accuracy = (1 - (abs(merged_df.loc[past_mask, 'RN Diff']).sum() / merged_df.loc[past_mask, 'HF RNs'].sum())) * 100
            past_revenue_accuracy = (1 - (abs(merged_df.loc[past_mask, 'Rev Diff']).sum() / merged_df.loc[past_mask, 'HF Rev'].sum())) * 100
            future_rooms_accuracy = (1 - (abs(merged_df.loc[future_mask, 'RN Diff']).sum() / merged_df.loc[future_mask, 'HF RNs'].sum())) * 100
            future_revenue_accuracy = (1 - (abs(merged_df.loc[future_mask, 'Rev Diff']).sum() / merged_df.loc[future_mask, 'HF Rev'].sum())) * 100

            # Display accuracy matrix in a table within a container for width control
            accuracy_data = {
                "RNs": [f"{past_rooms_accuracy:.2f}%", f"{future_rooms_accuracy:.2f}%"],
                "Revenue": [f"{past_revenue_accuracy:.2f}%", f"{future_revenue_accuracy:.2f}%"]
            }
            accuracy_df = pd.DataFrame(accuracy_data, index=["Past", "Future"])

            # Center the accuracy matrix table
            with st.container():
                st.table(accuracy_df.style.applymap(color_scale).set_table_styles([{"selector": "th", "props": [("backgroundColor", "#f0f2f6")]}]))

            # Warning about future discrepancies with matching colors
            st.warning("Future discrepancies might be a result of timing discrepancies between the moment that the data was received and the moment that the history and forecast file was received.")

            # Interactive bar and line chart for visualizing discrepancies
            fig = make_subplots(specs=[[{"secondary_y": True}]])

            # RN Discrepancies - Bar chart
            fig.add_trace(go.Bar(
                x=merged_df['date'],
                y=merged_df['RN Diff'],
                name='RNs Discrepancy',
                marker_color='blue'
            ), secondary_y=False)

            # Revenue Discrepancies - Line chart
            fig.add_trace(go.Scatter(
                x=merged_df['date'],
                y=merged_df['Rev Diff'],
                name='Revenue Discrepancy',
                mode='lines+markers',
                line=dict(color='red', width=2),
                marker=dict(size=8)
            ), secondary_y=True)

            # Update plot layout for dynamic axis scaling and increased height
            max_room_discrepancy = merged_df['RN Diff'].abs().max()
            max_revenue_discrepancy = merged_df['Rev Diff'].abs().max()

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

            # Align grid lines
            fig.update_yaxes(matches=None, showgrid=True, gridwidth=1, gridcolor='grey')

            st.plotly_chart(fig, use_container_width=True)

            # Display daily variance detail in a table
            st.markdown("### Daily Variance Detail", unsafe_allow_html=True)
            detail_container = st.container()
            with detail_container:
                formatted_df = merged_df[['date', 'HF RNs', 'HF Rev', 'Juyo RN', 'Juyo Rev', 'RN Diff', 'Rev Diff']]
                styled_df = formatted_df.style.format({'RN Diff': "{:+.0f}", 'Rev Diff': "{:+.2f}"}).applymap(color_scale, subset=['RN Diff', 'Rev Diff']).set_properties(**{'text-align': 'center'})
                st.table(styled_df)

    if __name__ == "__main__":
        main()
