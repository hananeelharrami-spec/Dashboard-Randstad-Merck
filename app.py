import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

# --- EN-TÊTE AVEC KPI MANUELS ---
col_title, col_kpis = st.columns([2, 2])
with col_title:
    st.title("📊 Dashboard de Pilotage")
    st.caption("Randstad / Merck")

with col_kpis:
    # Zone de saisie manuelle (côte à côte)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 👥 Effectifs")
        effectif = st.number_input("Intérimaires en poste", value=133, step=1)
    with c2:
        st.markdown("### ⭐ Satisfaction")
        nps = st.number_input("NPS Intérimaire (/10)", value=9.1, step=0.1, format="%.1f")

# --- FONCTION DE NETTOYAGE ROBUSTE ---
def clean_and_scale_data(df):
    # 0. Nettoyage des Noms de Colonnes
    df.columns = df.columns.str.strip()

    # 1. CONVERSION TEXTE -> NOMBRE (Sécurisée)
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                # On travaille sur une copie
                series = df[col].astype(str).str.strip()
                # Enlève les guillemets et espaces insécables
                series = series.str.replace('"', '').str.replace('\u202f', '').str.replace('\xa0', '')
                # Enlève % et remplace virgule
                series = series.str.replace('%', '').str.replace(' ', '').str.replace(',', '.')
                
                # Tente la conversion, remplace les erreurs par NaN
                df[col] = pd.to_numeric(series, errors='coerce')
            except Exception:
                pass

    # 2. SECURISATION DES ANNEES (Anti-Crash)
    if 'Année' in df.columns:
        # On force en numérique
        df['Année'] = pd.to_numeric(df['Année'], errors='coerce')
        # On remplit les NaN par une année par défaut (ex: 2025) pour ne pas perdre la donnée
        df['Année'] = df['Année'].fillna(2025).astype(int)

    # 3. MISE A L'ECHELLE DES POURCENTAGES
    for col in df.columns:
        col_lower = col.lower()
        keywords = ['taux', '%', 'atteinte', 'validation', 'rendement', 'impact']
        if any(x in col_lower for x in keywords):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                # Si max est petit (ex: 0.88), on x100
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
    return df

# --- CHARGEMENT ---
@st.cache_data
def load_data():
    excel_files = glob.glob("*.xlsx")
    if not excel_files:
        return None, None

    found_file = excel_files[0]
    data = {}
    
    expected = {
        "YTD": "CONSOLIDATION_YTD",
        "RECRUT": "Recrutement_Mensuel",
        "ABS": "Absentéisme_Global_Mois",
        "ABS_MOTIF": "Absentéisme_Par_Motif",
        "ABS_SERVICE": "Absentéisme_Par_Service",
        "SOURCE": "KPI_Sourcing_Rendement",
        "PLAN": "Suivi_Plan_Action"
    }
    
    try:
        xls = pd.ExcelFile(found_file)
        all_sheets = xls.sheet_names
        
        for key, sheet_name in expected.items():
            if sheet_name in all_sheets:
                # Lecture brute puis nettoyage
                df_raw = pd.read_excel(found_file, sheet_name=sheet_name)
                data[key] = clean_and_scale_data(df_raw)
        return data, found_file
        
    except Exception as e:
        st.error(f"Erreur technique lors de la lecture : {e}")
        return None, found_file

data, filename = load_data()

if data is None:
    st.error("❌ Aucun fichier Excel trouvé. Uploadez data.xlsx sur GitHub.")
    st.stop()
else:
    st.markdown("---")

# --- BARRE LATÉRALE : FILTRES ---
st.sidebar.header("Filtres")

# Récupération dynamique des années
annees_dispo = set()
for key, df in data.items():
    if 'Année' in df.columns:
        try:
            unique_years = df['Année'].dropna().unique()
            # On ne garde que ce qui ressemble à une année (ex: > 2020)
            valid_years = [int(y) for y in unique_years if y > 2020]
            annees_dispo.update(valid_years)
        except:
            pass

annees_dispo = sorted(list(annees_dispo))
# Par défaut, on met l'année la plus récente si disponible, sinon Toutes
index_default = len(annees_dispo) # Par défaut "Toutes" (qui sera à la fin de la liste d'options)

options_annee =
