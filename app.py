import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

# --- EN-TÊTE ---
col_title, col_plan = st.columns([3, 1])
with col_title:
    st.title("📊 Dashboard de Pilotage - Randstad / Merck")

with col_plan:
    st.markdown("### 👥 Planning / Effectifs")
    effectif = st.number_input("Intérimaires en poste", value=133, step=1)

# --- FONCTION DE NETTOYAGE ULTIME ---
def clean_and_scale_data(df):
    # 0. Nettoyage des Noms de Colonnes (enlève les espaces à la fin)
    df.columns = df.columns.str.strip()

    # 1. CONVERSION TEXTE -> NOMBRE
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                series = df[col].astype(str).str.replace('"', '').str.strip()
                series = series.str.replace('%', '').str.replace(' ', '').str.replace('\u202f', '')
                series = series.str.replace(',', '.')
                df[col] = pd.to_numeric(series, errors='ignore')
            except Exception:
                pass

    # 2. SECURISATION DES ANNEES (Pour voir 2026)
    if 'Année' in df.columns:
        # On force en numérique, les erreurs deviennent NaN
        df['Année'] = pd.to_numeric(df['Année'], errors='coerce')
        # On remplit les vides ou convertit en 0 pour éviter les plantages
        df['Année'] = df['Année'].fillna(0).astype(int)

    # 3. MISE A L'ECHELLE DES POURCENTAGES
    for col in df.columns:
        col_lower = col.lower()
        if any(x in col_lower for x in ['taux', '%', 'atteinte', 'validation', 'rendement', 'impact']):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
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
    try:
        xls = pd.ExcelFile(found_file)
        all_sheets = xls.sheet_names
        
        expected = {
            "YTD": "CONSOLIDATION_YTD",
            "RECRUT": "Recrutement_Mensuel",
            "ABS": "Absentéisme_Global_Mois",
            "ABS_MOTIF": "Absentéisme_Par_Motif",
            "ABS_SERVICE": "Absentéisme_Par_Service",
            "SOURCE": "KPI_Sourcing_Rendement",
            "PLAN": "Suivi_Plan_Action"
        }
        
        for key, sheet_name in expected.items():
            if sheet_name in all_sheets:
                df_raw = pd.read_excel(found_file, sheet_name=sheet_name)
                data[key] = clean_and_scale_data(df_raw)
        return data, found_file
        
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
        return None, found_file

data, filename = load_data()

if data is None:
    st.error("❌ Aucun fichier Excel trouvé. Uploadez data.xlsx sur GitHub.")
    st.stop()
else:
    st.markdown("---")

# --- SIDEBAR : FILTRES GLOBAUX ---
st.sidebar.header("Filtres")

# Récupération des années disponibles dans les données
annees_dispo = set()
for key, df in data.items():
    if 'Année' in df.columns:
        annees_dispo.update(df['Année'].unique())

# On enlève 0 si présent et on trie
annees_dispo = sorted([a for a in annees_dispo if a > 2000])

# Sélecteur d'année
annee_select = st.sidebar.radio("Choisir l'année :", ["Toutes"] + [str(a) for a in annees_dispo])

# FONCTION FILTRE
def filter_year(df):
    if annee_select == "Toutes":
        return df
    if 'Année' in df.columns:
        return df[df['Année'] == int(annee_select)]
    return df

# --- DASHBOARD ---

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Vue Globale", "🤝 Recrutement", "🏥 Absentéisme", "🔍 Sourcing", "✅ Plan d'Action"])

# --- 1. VUE GLOBALE (YTD) ---
with tab1:
    st.header(f"Performance {annee_select}")
    if "YTD" in data:
        df_ytd = filter_year(data["YTD"]) # Filtre appliqué
        
        if not df_ytd.empty:
            col_val = 'Valeur YTD'
            cols = st.columns(4)
            for index, row in df_ytd.iterrows():
                indic = row['Indicateur']
                val = row[col_val]
                val_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                cols[index % 4].metric(label=indic, value=val_str)
        else:
            st.warning(f"Pas de données consolidées pour {annee_select}")

# --- 2. RECRUTEMENT ---
with tab2:
    st.header("Recrutement Mensuel")
    if "RECRUT" in data:
        df_rec = filter_year(data["RECRUT"]) # Filtre appliqué
        
        if not df_rec.empty:
            if 'Mois' in df_rec.columns:
                df_rec = df_rec.sort_values(['Année', 'Mois'])
                # Axe X propre : Mois/Année
                df_rec['Période'] = df_rec['Mois'].astype(str) + "/" + df_rec['Année'].astype(str)

                c1, c2 = st.columns(2)
                with c1:
                    st.subheader("Taux de Service & Transformation")
                    fig = px.line(df_rec, x='Période', y=['Taux Service', 'Taux Transfo'], markers=True)
                    fig.update_
