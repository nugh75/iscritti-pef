import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime

st.set_page_config(page_title="Generatore File Excel per Classi", page_icon="📊")

st.title("Generatore File Excel per Classi di Corso")

st.write("""
Questa applicazione genera file Excel separati per ogni classe di corso a partire da un file CSV o Excel.
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
            
        # Estrai il codice della classe di concorso (es. A-01) dalla colonna "Classe"
        import re
        
        def estrai_codice_classe(nome_classe):
            # Cerca il pattern (X-YY) dove X è una lettera e YY sono numeri
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
            st.write(f"- {codice} ({esempio_classe}): {count} iscritti")
            
        if st.button("Genera file Excel per classi"):
            # Crea una directory per i file di output se non esiste
            output_dir = "file_classi"
            os.makedirs(output_dir, exist_ok=True)
            
            # Timestamp per i nomi dei file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Raggruppa per classe e genera file Excel
            generated_files = []
            
            # Per download multiplo, prepara un dizionario di buffer di memoria
            excel_files = {}
            
            for classe in classi:
                # Filtra il dataframe per classe
                df_classe = df[df["Classe"] == classe].copy()
                
                # Crea un nome file sicuro (rimuovi caratteri problematici)
                safe_class_name = "".join(c if c.isalnum() or c in [" ", "_", "-"] else "_" for c in classe)
                safe_class_name = safe_class_name.replace(" ", "_")
                
                # Crea il file nella directory
                file_name = f"{safe_class_name}_{timestamp}.xlsx"
                file_path = os.path.join(output_dir, file_name)
                
                # Salva su disco
                df_classe.to_excel(file_path, index=False)
                
                # Prepara il buffer per il download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_classe.to_excel(writer, index=False)
                excel_files[file_name] = output.getvalue()
                
                generated_files.append(file_path)
            
            st.success(f"Generati {len(generated_files)} file Excel nella cartella '{output_dir}'")
            
            # Offri il download di ciascun file
            st.subheader("Scarica i file generati:")
            
            for file_name, file_data in excel_files.items():
                classe_nome = file_name.split("_")[0]
                classe_completa = next((c for c in classi if classe_nome.replace("_", " ") in c.replace("(", "").replace(")", "")), classe_nome)
                st.download_button(
                    label=f"Scarica {classe_completa}",
                    data=file_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
    except Exception as e:
        st.error(f"Si è verificato un errore durante la lettura del file: {e}")
        
# Informazioni aggiuntive
st.markdown("---")
st.markdown("### Come utilizzare questa applicazione:")
st.markdown("""
1. Carica un file CSV o Excel contenente gli iscritti
2. Verifica i dati nell'anteprima
3. Clicca su 'Genera file Excel per classi' per creare un file Excel separato per ogni classe
4. Scarica i file generati o trovali nella cartella 'file_classi'
""")
