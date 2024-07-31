import streamlit as st
import pandas as pd
import io

st.title('Duplicatore Dope righe Excel')

# Controlla se un nuovo file è caricato
if 'uploaded_file' not in st.session_state:
    st.session_state.uploaded_file = None
    st.session_state.header_row = 0

# Carica il file Excel
uploaded_file = st.file_uploader("Scegli un file Excel", type="xlsx")

if uploaded_file is not None:
    # Controlla se il file caricato è nuovo
    if uploaded_file != st.session_state.uploaded_file:
        st.session_state.uploaded_file = uploaded_file
        st.session_state.header_row = 0
    
    # Input per la riga di intestazione
    header_row = st.number_input("Inserisci il numero della riga di intestazione", min_value=0, value=st.session_state.header_row, key="header_row_input")
    st.session_state.header_row = header_row
    
    # Carica il file Excel
    excel_data = pd.ExcelFile(uploaded_file)
    
    # Mostra i nomi dei fogli
    sheet_names = excel_data.sheet_names
    
    # Seleziona un foglio
    selected_sheet = st.selectbox("Seleziona un foglio da visualizzare", sheet_names)

    # Visualizza l'anteprima del foglio selezionato
    if selected_sheet:
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=st.session_state.header_row, dtype=str)
        st.write("Anteprima del foglio selezionato:")
        st.dataframe(df)
        
        # Seleziona le colonne per l'operazione
        columns = df.columns.tolist()
        col1 = st.selectbox("Seleziona la colonna con i valori da duplicare", columns, key="col1")
        col2 = st.selectbox("Seleziona la colonna con il numero di duplicazioni", columns, key="col2")
        
        if st.button("Modifica ed Esporta"):
            # Crea il dataframe modificato
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
                        new_row[col2] = '1'  # Imposta il valore della colonna di duplicazione a '1'
                        new_rows.append(new_row)
            
            modified_df = pd.DataFrame(new_rows)
            
            # Salva il file Excel modificato in memoria
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet in sheet_names:
                    if sheet == selected_sheet:
                        modified_df.to_excel(writer, index=False, sheet_name=sheet)
                    else:
                        pd.read_excel(uploaded_file, sheet_name=sheet, dtype=str).to_excel(writer, index=False, sheet_name=sheet)
            output.seek(0)
            
            st.success("File modificato con successo. Puoi scaricarlo qui sotto.")
            st.download_button(label="Scarica il file Excel modificato", data=output, file_name="file_modificato.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
            st.write("Anteprima del foglio modificato:")
            st.dataframe(modified_df)
