import streamlit as st
import pandas as pd
import os
import io
import re
import zipfile
from datetime import datetime

st.set_page_config(page_title="Generatore File Excel per Classi di Concorso", page_icon="📊")

st.title("Generatore File Excel per Classi di Concorso")

st.write("""
Questa applicazione genera file Excel separati per ogni classe di concorso a partire da un file CSV o Excel.
""")

uploaded_file = st.file_uploader("Carica il file CSV o Excel con gli iscritti", type=["csv", "xlsx", "xls"])

if uploaded_file is not None:
    # Determina il tipo di file e legge i dati
    file_extension = uploaded_file.name.split(".")[-1]
    
    try:
        if file_extension == "csv":
            # Resetta il puntatore del file prima della lettura
            uploaded_file.seek(0)
            
            # Prova diversi encodings per leggere il contenuto
            encodings_to_try = ['latin1', 'iso-8859-1', 'cp1252', 'utf-8-sig', 'utf-8']
            content = None
            
            for enc in encodings_to_try:
                try:
                    uploaded_file.seek(0)
                    content = uploaded_file.read().decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            
            if content is None:
                raise Exception("Impossibile decodificare il file con gli encodings disponibili")
            
            uploaded_file.seek(0)
            
            # Determina il separatore basandosi sul contenuto
            if ";" in content:
                separator = ";"
            elif "," in content:
                separator = ","
            else:
                separator = "\t"
            
            # Prova a leggere il file con lo stesso encoding usato per il contenuto
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, sep=separator, encoding=enc)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success(f"File caricato con successo: {uploaded_file.name}")
        
        # Visualizza l'anteprima dei dati
        st.subheader("Anteprima dei dati")
        st.dataframe(df.head())
        
        # Verifica che esista la colonna "Classe"
        if "Classe" not in df.columns:
            st.error("Il file non contiene una colonna 'Classe'. Assicurati che il file contenga questa colonna.")
            st.stop()
        
        # Funzione per estrarre il codice classe di concorso (es. A-01) dalla descrizione
        def estrai_codice_classe(nome_classe):
            match = re.search(r'\(([A-Z]-\d+)\)', nome_classe)
            if match:
                return match.group(1)  # Restituisce solo il codice, es. A-01
            return nome_classe  # Fallback al nome completo se non trova il codice
        
        # Crea una nuova colonna con il codice della classe di concorso
        df['Codice_Classe'] = df['Classe'].apply(estrai_codice_classe)
        
        # Ottieni tutti i codici classe unici
        codici_classe = df['Codice_Classe'].unique()
        st.write(f"Sono state trovate {len(codici_classe)} classi di concorso.")
        
        # Visualizza classi e conteggio studenti
        st.subheader("Classi di concorso trovate:")
        for codice in codici_classe:
            count = len(df[df['Codice_Classe'] == codice])
            # Mostra un esempio di nome completo della classe
            esempio_classe = df[df['Codice_Classe'] == codice]['Classe'].iloc[0]
            st.write(f"- {codice} ({esempio_classe.split('(')[0].strip()}): {count} iscritti")
            
        use_timestamp = st.checkbox("Aggiungi timestamp ai nomi dei file", value=False, 
                                  help="Se selezionato, verrà aggiunto un timestamp ai nomi dei file per evitare sovrascritture")
            
        if st.button("Genera file Excel per classi di concorso"):
            # Crea una directory per i file di output se non esiste
            output_dir = "file_classi_concorso"
            os.makedirs(output_dir, exist_ok=True)
            
            # Timestamp per i nomi dei file (opzionale)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S") if use_timestamp else ""
            timestamp_suffix = f"_{timestamp}" if timestamp else ""
            
            # Raggruppa per codice classe e genera file Excel
            generated_files = []
            
            # Per download multiplo, prepara un dizionario di buffer di memoria
            excel_files = {}
            
            for codice in codici_classe:
                # Filtra il dataframe per codice classe
                df_classe = df[df['Codice_Classe'] == codice].copy()
                
                # Usa il codice classe direttamente come nome del file
                safe_code = codice.replace("/", "_").replace("\\", "_")
                
                # Crea il file nella directory
                file_name = f"{safe_code}{timestamp_suffix}.xlsx"
                file_path = os.path.join(output_dir, file_name)
                
                # Salva su disco (rimuovi la colonna Codice_Classe che abbiamo aggiunto)
                df_output = df_classe.drop(columns=['Codice_Classe'])
                df_output.to_excel(file_path, index=False)
                
                # Prepara il buffer per il download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_output.to_excel(writer, index=False)
                excel_files[file_name] = output.getvalue()
                
                generated_files.append(file_path)
            
            st.success(f"Generati {len(generated_files)} file Excel nella cartella '{output_dir}'")
            
            # Offri opzioni per il download dei file
            st.subheader("Scarica i file generati:")
            
            # Aggiungi checkbox per selezionare quali file scaricare
            st.write("Seleziona i file da scaricare:")
            
            # Dizionario per tenere traccia delle selezioni
            file_selections = {}
            
            # Crea tre colonne per organizzare meglio le checkbox
            cols = st.columns(3)
            
            # Distribuisci le checkbox tra le colonne
            for i, (file_name, file_data) in enumerate(excel_files.items()):
                codice = file_name.split("_")[0] if timestamp else file_name.replace(".xlsx", "")
                esempio_classe = df[df['Codice_Classe'] == codice]['Classe'].iloc[0].split('(')[0].strip()
                display_name = f"{codice} - {esempio_classe}"
                
                # Distribuisci in modo uniforme tra le colonne
                col_idx = i % 3
                with cols[col_idx]:
                    file_selections[file_name] = st.checkbox(display_name, value=True)
            
            # Pulsanti per selezionare/deselezionare tutti
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Seleziona tutti"):
                    for key in file_selections:
                        file_selections[key] = True
                    st.experimental_rerun()
            
            with col2:
                if st.button("Deseleziona tutti"):
                    for key in file_selections:
                        file_selections[key] = False
                    st.experimental_rerun()
            
            # Lista dei file selezionati
            selected_files = {name: data for name, data in excel_files.items() if file_selections.get(name, False)}
            
            if not selected_files:
                st.warning("Nessun file selezionato per il download.")
            else:
                st.write(f"{len(selected_files)} file selezionati per il download.")
                
                # Download dei singoli file
                for file_name, file_data in selected_files.items():
                    codice = file_name.split("_")[0] if timestamp else file_name.replace(".xlsx", "")
                    esempio_classe = df[df['Codice_Classe'] == codice]['Classe'].iloc[0].split('(')[0].strip()
                    st.download_button(
                        label=f"Scarica {codice} - {esempio_classe}",
                        data=file_data,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"single_{file_name}"
                    )
                
                # Crea un file ZIP con tutti i file selezionati
                if len(selected_files) > 1:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for file_name, file_data in selected_files.items():
                            zip_file.writestr(file_name, file_data)
                    
                    # Offri il download del file ZIP
                    st.download_button(
                        label=f"Scarica tutti i file selezionati ({len(selected_files)} file) in ZIP",
                        data=zip_buffer.getvalue(),
                        file_name=f"classi_concorso_{timestamp if timestamp else datetime.now().strftime('%Y%m%d')}.zip",
                        mime="application/zip"
                    )
            
    except Exception as e:
        st.error(f"Si è verificato un errore durante la lettura del file: {e}")
        st.exception(e)  # Mostra l'errore dettagliato per il debug
        
# Informazioni aggiuntive
st.markdown("---")
st.markdown("### Come utilizzare questa applicazione:")
st.markdown("""
1. Carica un file CSV o Excel contenente gli iscritti
2. Verifica i dati nell'anteprima
3. Decidi se aggiungere il timestamp ai nomi dei file
4. Clicca su 'Genera file Excel per classi di concorso' per creare un file Excel separato per ogni classe
5. Scarica i file generati o trovali nella cartella 'file_classi_concorso'
""")
