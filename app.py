import streamlit as st
import pandas as pd
import io
import os

st.title('Excel File Previewer and Modifier')

# Check if a new file is uploaded
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
    st.session_state.header_row = 0

# Upload the Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Check if the uploaded file is new
    if uploaded_file != st.session_state.uploaded_file:
        st.session_state.uploaded_file = uploaded_file
        st.session_state.header_row = 0
    
    # Load the Excel file
    excel_data = pd.ExcelFile(uploaded_file)
    
    # Display sheet names
    sheet_names = excel_data.sheet_names
    
    # Select a sheet
    selected_sheet = st.selectbox("Select a sheet to preview", sheet_names)

    # Input header row
    header_row = st.number_input("Enter the header row number (starting from 0)", min_value=0, value=st.session_state.header_row, key="header_row")
    
    # Display preview of selected sheet
    if selected_sheet:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row, dtype=str)
        st.write("Preview of the selected sheet:")
        st.dataframe(df)
        
        # Select columns for the operation
        columns = df.columns.tolist()
        col1 = st.selectbox("Select the column with values to duplicate", columns, key="col1")
        col2 = st.selectbox("Select the column with number of duplications", columns, key="col2")
        
        if st.button("Modify and Export"):
            # Create the modified dataframe
            new_rows = []
            for index, row in df.iterrows():
                value = row[col1]
                repeat_times = row[col2]
                if pd.isna(repeat_times) or repeat_times.strip() == '' or not repeat_times.isdigit() or int(repeat_times) == 0:
                    new_rows.append(row)
                else:
                    repeat_times = int(repeat_times)
                    for _ in range(repeat_times):
                        new_row = row.copy()
                        new_row[col2] = '1'  # Set the duplication column value to '1'
                        new_rows.append(new_row)
            
            modified_df = pd.DataFrame(new_rows)
            
            # Save to a new Excel file in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet in sheet_names:
                    if sheet == selected_sheet:
                        modified_df.to_excel(writer, index=False, sheet_name=sheet)
                    else:
                        pd.read_excel(uploaded_file, sheet_name=sheet, dtype=str).to_excel(writer, index=False, sheet_name=sheet)
            output.seek(0)
            
            # Get the original file name and add "_modificato"
            original_filename = uploaded_file.name
            new_filename = os.path.splitext(original_filename)[0] + "_modificato.xlsx"
            
            st.success(f"File exported successfully. You can download it below as {new_filename}.")
            st.download_button(label="Download modified Excel file", data=output, file_name=new_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            st.write("Preview of the modified sheet:")
            st.dataframe(modified_df)
