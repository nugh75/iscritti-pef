import streamlit as st
import pandas as pd
import os
import io
import zipfile
from datetime import datetime

st.set_page_config(page_title="Generatore File Excel per Classi", page_icon="📊")

st.title("Generatore File Excel per Classi di Corso")

st.write("""
Questa applicazione genera file Excel separati per ogni classe di corso a partire da un file CSV o Excel.
""")

# File con gli iscritti attuali
st.subheader("1. Carica il file CSV o Excel con gli iscritti attuali")
uploaded_file = st.file_uploader("Carica il file principale", type=["csv", "xlsx", "xls"])

# File opzionale con gli iscritti precedenti (per confronto)
st.subheader("2. File di confronto (opzionale)")
st.write("""
Se vuoi identificare solo i nuovi iscritti, carica un file con gli iscritti precedenti. 
Il confronto verrà fatto utilizzando il Codice Fiscale (CF) come identificatore univoco.
""")
compare_file = st.file_uploader("Carica il file di confronto", type=["csv", "xlsx", "xls"])

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
            
        st.success(f"File principale caricato con successo: {uploaded_file.name}")
        
        # Visualizza l'anteprima dei dati
        st.subheader("Anteprima dei dati")
        st.dataframe(df.head())
        
        # Verifica che esistano le colonne necessarie
        if "Classe" not in df.columns:
            st.error("Il file non contiene una colonna 'Classe'. Assicurati che il file contenga questa colonna.")
            st.stop()
            
        if "CF" not in df.columns:
            st.warning("Il file non contiene una colonna 'CF'. Il confronto per identificare i nuovi iscritti non sarà possibile.")
            has_cf = False
        else:
            has_cf = True
            
        # Gestione del file di confronto se presente
        df_nuovi = None
        if compare_file is not None:
            try:
                # Determina il tipo di file e legge i dati del file di confronto
                compare_extension = compare_file.name.split(".")[-1]
                
                if compare_extension == "csv":
                    # Riutilizziamo la logica di lettura CSV
                    compare_file.seek(0)
                    
                    # Prova diversi encodings
                    content = None
                    for enc in encodings_to_try:
                        try:
                            compare_file.seek(0)
                            content = compare_file.read().decode(enc)
                            break
                        except UnicodeDecodeError:
                            continue
                    
                    if content is None:
                        raise Exception("Impossibile decodificare il file di confronto con gli encodings disponibili")
                    
                    compare_file.seek(0)
                    
                    # Determina il separatore
                    if ";" in content:
                        separator_compare = ";"
                    elif "," in content:
                        separator_compare = ","
                    else:
                        separator_compare = "\t"
                    
                    compare_file.seek(0)
                    df_compare = pd.read_csv(compare_file, sep=separator_compare, encoding=enc)
                else:
                    df_compare = pd.read_excel(compare_file)
                    
                st.success(f"File di confronto caricato con successo: {compare_file.name}")
                
                # Verifica che il file di confronto contenga la colonna CF
                if "CF" not in df_compare.columns:
                    st.error("Il file di confronto non contiene una colonna 'CF'. Non è possibile effettuare il confronto.")
                elif not has_cf:
                    st.error("Il file principale non contiene una colonna 'CF'. Non è possibile effettuare il confronto.")
                else:
                    # Confronta i due file e trova i nuovi iscritti (presenti nel file principale ma non nel file di confronto)
                    cf_vecchi = set(df_compare["CF"].str.upper())
                    df["CF_upper"] = df["CF"].str.upper()  # Normalizza maiuscole/minuscole
                    df_nuovi = df[~df["CF_upper"].isin(cf_vecchi)].copy()
                    df.drop("CF_upper", axis=1, inplace=True)  # Rimuovi la colonna temporanea
                    
                    num_nuovi = len(df_nuovi)
                    num_totale = len(df)
                    
                    st.subheader("Risultato del confronto:")
                    st.write(f"- Iscritti totali nel file principale: {num_totale}")
                    st.write(f"- Nuovi iscritti (non presenti nel file di confronto): {num_nuovi}")
                    
                    # Mostra un'anteprima dei nuovi iscritti
                    if num_nuovi > 0:
                        st.subheader("Anteprima dei nuovi iscritti:")
                        st.dataframe(df_nuovi.head())
                    
                    # Opzione per utilizzare tutti gli iscritti o solo i nuovi
                    use_only_new = st.checkbox("Genera file Excel solo con i nuovi iscritti", value=True)
                    
                    if use_only_new and num_nuovi > 0:
                        # Se l'utente sceglie di usare solo i nuovi iscritti, sostituisci df con df_nuovi
                        st.info(f"Verranno utilizzati solo i {num_nuovi} nuovi iscritti per generare i file Excel.")
                        df_to_use = df_nuovi
                    else:
                        # Altrimenti usa tutti gli iscritti
                        st.info(f"Verranno utilizzati tutti i {num_totale} iscritti per generare i file Excel.")
                        df_to_use = df
            except Exception as e:
                st.error(f"Si è verificato un errore durante la lettura del file di confronto: {e}")
                df_to_use = df  # In caso di errore, utilizza tutti gli iscritti
        else:
            # Se non è stato caricato un file di confronto, utilizza tutti gli iscritti
            df_to_use = df
            
        # Permetti all'utente di selezionare le colonne da includere
        st.subheader("Seleziona le colonne da includere nei file Excel:")
        
        # La colonna "Classe" è obbligatoria e non può essere deselezionata
        columns_to_select = [col for col in df_to_use.columns if col != "Classe"]
        
        # Crea tre colonne per organizzare meglio le checkbox
        cols_check = st.columns(3)
        
        # Dizionario per tenere traccia delle selezioni delle colonne
        column_selections = {}
        
        # Distribuisci le checkbox tra le colonne
        for i, col in enumerate(columns_to_select):
            col_idx = i % 3
            with cols_check[col_idx]:
                column_selections[col] = st.checkbox(f"{col}", value=True, key=f"col_{col}")
        
        # Pulsanti per selezionare/deselezionare tutte le colonne
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Seleziona tutte le colonne"):
                for key in column_selections:
                    column_selections[key] = True
                st.experimental_rerun()
        
        with col2:
            if st.button("Deseleziona tutte le colonne"):
                for key in column_selections:
                    column_selections[key] = False
                st.experimental_rerun()
        
        # Ottieni le colonne selezionate (la colonna "Classe" è sempre inclusa)
        selected_columns = ["Classe"] + [col for col in columns_to_select if column_selections.get(col, False)]
        
        if len(selected_columns) <= 1:
            st.warning("Devi selezionare almeno una colonna oltre a 'Classe'.")
            selected_columns = df_to_use.columns.tolist()  # Se nessuna colonna è selezionata, usa tutte le colonne
            
        # Estrai il codice della classe di concorso (es. A-01) dalla colonna "Classe"
        import re
        
        def estrai_codice_classe(nome_classe):
            # Cerca il pattern (X-YY) dove X è una lettera e YY sono numeri
            match = re.search(r'\(([A-Z]-\d+)\)', nome_classe)
            if match:
                return match.group(1)  # Restituisce solo il codice, es. A-01
            return nome_classe  # Fallback al nome completo se non trova il codice
        
        # Crea una nuova colonna con il codice della classe di concorso
        df_to_use['Codice_Classe'] = df_to_use['Classe'].apply(estrai_codice_classe)
        
        # Ottieni tutti i codici classe unici
        codici_classe = df_to_use['Codice_Classe'].unique()
        st.write(f"Sono state trovate {len(codici_classe)} classi di concorso.")
        
        # Visualizza classi e conteggio studenti
        st.subheader("Classi di concorso trovate:")
        for codice in codici_classe:
            count = len(df_to_use[df_to_use['Codice_Classe'] == codice])
            # Mostra un esempio di nome completo della classe
            esempio_classe = df_to_use[df_to_use['Codice_Classe'] == codice]['Classe'].iloc[0]
            st.write(f"- {codice} ({esempio_classe}): {count} iscritti")
            
        if st.button("Genera file Excel per classi di concorso"):
            # Crea una directory per i file di output se non esiste
            output_dir = "file_classi_concorso"
            os.makedirs(output_dir, exist_ok=True)
            
            # Timestamp per i nomi dei file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Raggruppa per codice classe e genera file Excel
            generated_files = []
            
            # Per download multiplo, prepara un dizionario di buffer di memoria
            excel_files = {}
            
            for codice in codici_classe:
                # Filtra il dataframe per codice classe
                df_classe = df_to_use[df_to_use['Codice_Classe'] == codice].copy()
                
                # Se non ci sono iscritti in questa classe, salta alla prossima classe
                if len(df_classe) == 0:
                    continue
                    
                # Usa il codice classe direttamente come nome del file
                safe_code = codice.replace("/", "_").replace("\\", "_")
                
                # Crea il file nella directory
                file_name = f"{safe_code}_{timestamp}.xlsx"
                file_path = os.path.join(output_dir, file_name)
                
                # Seleziona solo le colonne scelte dall'utente (rimuovi sempre la colonna Codice_Classe)
                df_output = df_classe[selected_columns].copy()
                if 'Codice_Classe' in df_output.columns:
                    df_output = df_output.drop(columns=['Codice_Classe'])
                
                # Salva su disco
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
                    file_selections[file_name] = st.checkbox(display_name, value=True, key=f"check_{codice}")
            
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
                        key=f"single_{codice}"
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
                        file_name=f"classi_concorso_{timestamp}.zip",
                        mime="application/zip"
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
