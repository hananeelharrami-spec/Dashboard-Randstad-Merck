import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

st.title("📊 Dashboard de Pilotage - Randstad / Merck")
st.markdown("---")

# --- FONCTION DE NETTOYAGE (Inchangée) ---
def clean_data(df):
    """Nettoie les chiffres français (virgule, %)"""
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                series = df[col].astype(str).str.replace('"', '').str.strip()
                if series.str.contains('%').any():
                    series = series.str.replace('%', '').str.replace(',', '.').replace(' ', '')
                    df[col] = pd.to_numeric(series, errors='coerce')
                elif series.str.match(r'^-?\d+(?:,\d+)?$').any():
                    series = series.str.replace(',', '.').replace(' ', '').replace('\u202f', '')
                    df[col] = pd.to_numeric(series, errors='coerce')
            except Exception:
                pass
    return df

# --- SIDEBAR : UPLOAD UNIQUE ---
st.sidebar.header("📂 Import des Données")
st.sidebar.info("Téléchargez le Google Sheet en format .xlsx et déposez-le ici.")

uploaded_file = st.sidebar.file_uploader("Fichier Dashboard Global (.xlsx)", type="xlsx")

# Dictionnaire pour stocker les données chargées
data = {}

if uploaded_file:
    try:
        # On charge tout le fichier Excel d'un coup
        xls = pd.ExcelFile(uploaded_file)
        all_sheet_names = xls.sheet_names
        
        # Liste des onglets attendus (ceux générés par ton script Apps Script)
        expected_sheets = {
            "YTD": "CONSOLIDATION_YTD",
            "RECRUT": "Recrutement_Mensuel",
            "ABS": "Absentéisme_Global_Mois",
            "SOURCE": "KPI_Sourcing_Rendement",
            "PLAN": "Suivi_Plan_Action"
        }
        
        # Chargement intelligent
        for key, sheet_name in expected_sheets.items():
            if sheet_name in all_sheet_names:
                data[key] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                data[key] = clean_data(data[key]) # Nettoyage immédiat
            else:
                st.sidebar.warning(f"⚠️ Onglet manquant : {sheet_name}")
                
        st.sidebar.success("Fichier chargé avec succès !")
        
    except Exception as e:
        st.sidebar.error(f"Erreur de lecture du fichier : {e}")

# --- DASHBOARD ---

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Vue Globale (YTD)", "🤝 Recrutement", "🏥 Absentéisme", "🔍 Sourcing", "✅ Plan d'Action"])

# 1. VUE GLOBALE (YTD)
with tab1:
    st.header("Performance Annuelle (Year-To-Date)")
    if "YTD" in data:
        df_ytd = data["YTD"]
        # Nettoyage spécifique valeur YTD
        col_val = 'Valeur YTD'
        if col_val in df_ytd.columns:
            cols = st.columns(4)
            for index, row in df_ytd.iterrows():
                indic = row['Indicateur']
                # On s'assure que c'est une string pour l'affichage
                val = str(row[col_val]) 
                col_idx = index % 4
                cols[col_idx].metric(label=indic, value=val)
        else:
            st.error("Colonne 'Valeur YTD' introuvable.")
    else:
        st.info("En attente du fichier...")

# 2. RECRUTEMENT
with tab2:
    st.header("Performance Recrutement Mensuel")
    if "RECRUT" in data:
        df_rec = data["RECRUT"]
        # Tri chronologique
        if 'Année' in df_rec.columns and 'Mois' in df_rec.columns:
            df_rec = df_rec.sort_values(['Année', 'Mois'])
            
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Taux de Service & Transformation")
                fig_taux = px.line(df_rec, x='Mois', y=['Taux Service', 'Taux Transfo'], 
                                   markers=True, title="Evolution des Taux (%)")
                st.plotly_chart(fig_taux, use_container_width=True)

            with col2:
                st.subheader("Volumes")
                fig_vol = px.bar(df_rec, x='Mois', y=['Nb Requisitions', 'Nb Hired'],
                                 barmode='group', title="Commandes vs Embauches")
                st.plotly_chart(fig_vol, use_container_width=True)
            
            with st.expander("Données détaillées"):
                st.dataframe(df_rec)
    else:
        st.info("En attente du fichier...")

# 3. ABSENTÉISME
with tab3:
    st.header("Suivi de l'Absentéisme")
    if "ABS" in data:
        df_abs = data["ABS"]
        if 'Année' in df_abs.columns and 'Mois' in df_abs.columns:
             df_abs = df_abs.sort_values(['Année', 'Mois'])

        fig_abs = px.area(df_abs, x='Mois', y='Taux Absentéisme', 
                          title="Taux d'Absentéisme Global (%)", markers=True,
                          color_discrete_sequence=['#FF5733'])
        st.plotly_chart(fig_abs, use_container_width=True)
        
        kpi1, kpi2 = st.columns(2)
        if 'Taux Absentéisme' in df_abs.columns:
             val_mean = df_abs['Taux Absentéisme'].mean()
             kpi1.metric("Moyenne Annuelle", f"{val_mean:.2f}%")
        
        kpi1.dataframe(df_abs)
    else:
        st.info("En attente du fichier...")

# 4. SOURCING
with tab4:
    st.header("Entonnoir de Sourcing")
    if "SOURCE" in data:
        df_source = data["SOURCE"]
        
        # Agrégation annuelle par source
        if 'Source' in df_source.columns:
            df_agg = df_source.groupby('Source')[['1. Appels Reçus', '3. Intégrés (Délégués)']].sum().reset_index()
            
            st.subheader("Efficacité par Canal")
            fig_source = px.bar(df_agg, x='Source', y=['1. Appels Reçus', '3. Intégrés (Délégués)'],
                                barmode='group', title="Volume vs Intégration")
            st.plotly_chart(fig_source, use_container_width=True)
            st.dataframe(df_source)
    else:
        st.info("En attente du fichier...")

# 5. PLAN D'ACTION
with tab5:
    st.header("Avancement du Plan d'Action")
    if "PLAN" in data:
        df_plan = data["PLAN"]
        
        # Jauge Global
        row_global = df_plan[df_plan['Catégorie / Section'].astype(str).str.contains('GLOBAL', case=False, na=False)]
        
        if not row_global.empty:
            taux = row_global.iloc[0]['% Atteinte']
            # Si c'est déjà un float grâce à clean_data, parfait, sinon on convertit
            fig_gauge = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = taux if isinstance(taux, (int, float)) else 0,
                title = {'text': "Avancement Global (%)"},
                gauge = {'axis': {'range': [None, 100]}, 'bar': {'color': "green"}}
            ))
            st.plotly_chart(fig_gauge, use_container_width=True)
        
        st.dataframe(df_plan)
    else:
        st.info("En attente du fichier...")
