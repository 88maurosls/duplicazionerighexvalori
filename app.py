import streamlit as st
import pandas as pd

st.title('Excel File Previewer')

# Upload the Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Load the Excel file
    excel_data = pd.ExcelFile(uploaded_file)
    
    # Display sheet names
    sheet_names = excel_data.sheet_names
    st.write("Available Sheets:")
    st.write(sheet_names)
    
    # Select a sheet
    selected_sheet = st.selectbox("Select a sheet to preview", sheet_names)
    
    # Display preview of selected sheet
    if selected_sheet:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        st.write("Preview of the selected sheet:")
        st.dataframe(df)
