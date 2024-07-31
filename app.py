import streamlit as st
import pandas as pd
import io

st.title('Excel File Previewer and Modifier')

# Upload the Excel file
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Load the Excel file
    excel_data = pd.ExcelFile(uploaded_file)
    
    # Display sheet names
    sheet_names = excel_data.sheet_names
    
    # Select a sheet
    selected_sheet = st.selectbox("Select a sheet to preview", sheet_names)
    
    # Display preview of selected sheet
    if selected_sheet:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        st.write("Preview of the selected sheet:")
        st.dataframe(df)
        
        # Select columns for the operation
        columns = df.columns.tolist()
        col1 = st.selectbox("Select the column with values to duplicate", columns)
        col2 = st.selectbox("Select the column with number of duplications", columns)
        
        if st.button("Modify and Export"):
            # Create the modified dataframe
            new_rows = []
            for index, row in df.iterrows():
                value = row[col1]
                repeat_times = int(row[col2])
                for _ in range(repeat_times):
                    new_rows.append(row)
            
            modified_df = pd.DataFrame(new_rows)
            
            # Save to a new Excel file in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                modified_df.to_excel(writer, index=False, sheet_name='Modified')
            output.seek(0)
            
            st.success("File exported successfully. You can download it below.")
            st.download_button(label="Download modified Excel file", data=output, file_name="modified_file.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            st.dataframe(modified_df)
