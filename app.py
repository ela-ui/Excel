import streamlit as st
import zipfile
import pandas as pd
from io import BytesIO
from datetime import datetime

# Streamlit app title
st.title("Excel Combiner from ZIP")

# Step 1: Upload ZIP file
uploaded_zip = st.file_uploader("Upload a ZIP file containing Excel files", type="zip")

# Step 2: Date selection for subtraction
selected_date = st.date_input("Select a date to subtract from Date of Disbursement")

if uploaded_zip:
    # Step 3: Read the ZIP file
    with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
        # Step 4: Get all Excel file names within the ZIP archive
        excel_files = [f for f in zip_ref.namelist() if f.endswith('.xlsx')]

        # Initialize an empty DataFrame for combining data
        combined_df = pd.DataFrame()

        # Step 5: Process each Excel file in the ZIP
        for excel_file in excel_files:
            with zip_ref.open(excel_file) as file:
                # Read all sheets from the Excel file
                excel_df = pd.read_excel(file, sheet_name=None)

                # Flatten sheets and concatenate them into a single DataFrame
                for sheet_name, sheet_df in excel_df.items():
                    # Check if the 'Date of Disbursement' column exists
                    if 'Date of Disbursement' in sheet_df.columns and 'Ageing' in sheet_df.columns:
                        # Convert 'Date of Disbursement' to datetime if it's not already
                        sheet_df['Date of Disbursement'] = pd.to_datetime(sheet_df['Date of Disbursement'], errors='coerce')

                        # Convert selected_date to datetime
                        selected_date = pd.to_datetime(selected_date)

                        # Subtract the selected date from 'Date of Disbursement' to calculate the 'Ageing'
                        sheet_df['Ageing'] = (selected_date - sheet_df['Date of Disbursement']).dt.days

                        # Update 'Slab' column based on 'Ageing' value
                        slab_conditions = [
                            (sheet_df['Ageing'] > 60),
                            (sheet_df['Ageing'] > 90),
                            (sheet_df['Ageing'] > 180),
                            (sheet_df['Ageing'] > 365)
                        ]
                        slab_values = ['>60', '>90', '>180', '>365']

                        # Apply conditions to update 'Slab' based on Ageing value
                        # Initially set all Slab to a default value, e.g., '<60'
                        sheet_df['Slab'] = '<60'

                        for condition, slab in zip(slab_conditions, slab_values):
                            sheet_df.loc[condition, 'Slab'] = slab

                    # Add 'State_Count' column with value 1
                    sheet_df['State_Count'] = 1

                    # Concatenate the current sheet's data with the combined DataFrame
                    combined_df = pd.concat([combined_df, sheet_df], ignore_index=True)

        # Step 6: Save combined DataFrame to an Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Combined Data')

        # Step 7: Prepare the download of the combined file
        st.success("Excel files have been successfully combined and processed!")
        st.download_button(
            label="Download Combined Excel File",
            data=output.getvalue(),
            file_name="combined_output_with_state_count_and_slab.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
