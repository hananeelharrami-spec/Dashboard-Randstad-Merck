import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

st.title("📊 Dashboard de Pilotage - Randstad / Merck")
st.markdown("---")

# --- CONFIGURATION DU FICHIER SOURCE ---
# C'est ici que ça se joue : on définit le nom du fichier fixe
DATA_FILE = "data.xlsx"

# --- FONCTION DE NETTOYAGE RENFORCÉE ---
def clean_and_scale_data(df):
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

    # 2. MISE A L'ECHELLE DES POURCENTAGES (0.88 -> 88.0)
    for col in df.columns:
        col_lower = col.lower()
        if any(x in col_lower for x in ['taux', '%', 'atteinte', 'validation', 'rendement']):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
    return df

# --- CHARGEMENT DES DONNÉES AUTOMATIQUE ---
@st.cache_data # Garde les données en mémoire pour que ce soit ultra rapide
def load_data():
    if not os.path.exists(DATA_FILE):
        return None
    
    data = {}
    try:
        xls = pd.ExcelFile(DATA_FILE)
        all_sheets = xls.sheet_names
        
        expected = {
            "YTD": "CONSOLIDATION_YTD",
            "RECRUT": "Recrutement_Mensuel",
            "ABS": "Absentéisme_Global_Mois",
            "SOURCE": "KPI_Sourcing_Rendement",
            "PLAN": "Suivi_Plan_Action"
        }
        
        for key, sheet_name in expected.items():
            if sheet_name in all_sheets:
                df_raw = pd.read_excel(DATA_FILE, sheet_name=sheet_name)
                data[key] = clean_and_scale_data(df_raw)
        return data
    except Exception as e:
        st.error(f"Erreur de lecture du fichier : {e}")
        return None

# Exécution du chargement
data = load_data()

if data is None:
    st.error(f"⚠️ Le fichier source '{DATA_FILE}' est introuvable sur le serveur.")
    st.info("Administrateur : Veuillez uploader 'data.xlsx' sur GitHub.")
    st.stop() # Arrête l'app si pas de fichier

# --- DASHBOARD (Code inchangé pour l'affichage) ---

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Vue Globale", "🤝 Recrutement", "🏥 Absentéisme", "🔍 Sourcing", "✅ Plan d'Action"])

# --- 1. VUE GLOBALE (YTD) ---
with tab1:
    st.header("Performance Annuelle (Year-To-Date)")
    if "YTD" in data:
        df_ytd = data["YTD"]
        col_val = 'Valeur YTD'
        if col_val in df_ytd.columns:
            cols = st.columns(4)
            for index, row in df_ytd.iterrows():
                indic = row['Indicateur']
                val = row[col_val]
                val_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                cols[index % 4].metric(label=indic, value=val_str)

# --- 2. RECRUTEMENT ---
with tab2:
    st.header("Recrutement Mensuel")
    if "RECRUT" in data:
        df_rec = data["RECRUT"]
        if 'Mois' in df_rec.columns:
            if 'Année' in df_rec.columns:
                df_rec = df_rec.sort_values(['Année', 'Mois'])
                df_rec['Période'] = df_rec['Mois'].astype(str) + "/" + df_rec['Année'].astype(str)
            else:
                df_rec['Période'] = df_rec['Mois'].astype(str)

            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Taux de Service & Transformation")
                fig = px.line(df_rec, x='Période', y=['Taux Service', 'Taux Transfo'], markers=True)
                fig.update_layout(yaxis_ticksuffix="%")
                st.plotly_chart(fig, use_container_width=True)

            with c2:
                st.subheader("Volume Commandes vs Hired")
                fig_bar = px.bar(df_rec, x='Période', y=['Nb Requisitions', 'Nb Hired'], barmode='group')
                st.plotly_chart(fig_bar, use_container_width=True)
            
            with st.expander("Voir le détail chiffré"):
                df_show = df_rec.copy()
                for col in df_show.columns:
                    if 'Taux' in col and pd.api.types.is_numeric_dtype(df_show[col]):
                        df_show[col] = df_show[col].apply(lambda x: f"{x:.2f}%")
                st.dataframe(df_show)

# --- 3. ABSENTÉISME ---
with tab3:
    st.header("Absentéisme")
    if "ABS" in data:
        df_abs = data["ABS"]
        if 'Mois' in df_abs.columns:
            if 'Année' in df_abs.columns:
                 df_abs = df_abs.sort_values(['Année', 'Mois'])
                 df_abs['Période'] = df_abs['Mois'].astype(str) + "/" + df_abs['Année'].astype(str)
            else:
                 df_abs['Période'] = df_abs['Mois'].astype(str)

        fig_abs = px.area(df_abs, x='Période', y='Taux Absentéisme', 
                          title="Taux Absentéisme Global", markers=True, color_discrete_sequence=['#FF5733'])
        fig_abs.update_layout(yaxis_ticksuffix="%")
        st.plotly_chart(fig_abs, use_container_width=True)
        
        c1, c2 = st.columns(2)
        if 'Taux Absentéisme' in df_abs.columns:
             moy = df_abs['Taux Absentéisme'].mean()
             c1.metric("Moyenne Annuelle", f"{moy:.2f}%")
        
        df_show_abs = df_abs.copy()
        if 'Taux Absentéisme' in df_show_abs.columns:
             df_show_abs['Taux Absentéisme'] = df_show_abs['Taux Absentéisme'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int,float)) else x)
        c1.dataframe(df_show_abs)

# --- 4. SOURCING ---
with tab4:
    st.header("Rendement Sourcing")
    if "SOURCE" in data:
        df_src = data["SOURCE"]
        if 'Source' in df_src.columns:
            df_agg = df_src.groupby('Source', as_index=False)[['1. Appels Reçus', '3. Intégrés (Délégués)']].sum()
            df_agg = df_agg.sort_values('1. Appels Reçus', ascending=False)
            
            st.subheader("Volume vs Intégration par Source")
            fig_src = px.bar(df_agg, x='Source', y=['1. Appels Reçus', '3. Intégrés (Délégués)'], barmode='group')
            st.plotly_chart(fig_src, use_container_width=True)
            
            st.write("Détail Mensuel")
            df_src_show = df_src.copy()
            for col in df_src_show.columns:
                if 'Taux' in col and pd.api.types.is_numeric_dtype(df_src_show[col]):
                    df_src_show[col] = df_src_show[col].apply(lambda x: f"{x:.2f}%")
            st.dataframe(df_src_show)

# --- 5. PLAN D'ACTION ---
with tab5:
    st.header("Plan d'Action")
    if "PLAN" in data:
        df_plan = data["PLAN"]
        row_global = df_plan[df_plan['Catégorie / Section'].astype(str).str.contains('GLOBAL', case=False, na=False)]
        if not row_global.empty:
            val = row_global.iloc[0]['% Atteinte']
            fig_gauge = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = val,
                number = {'suffix': "%"},
                title = {'text': "Avancement Global"},
                gauge = {'axis': {'range': [None, 100]}, 'bar': {'color': "green"}}
            ))
            st.plotly_chart(fig_gauge, use_container_width=True)
        
        df_plan_show = df_plan.copy()
        if '% Atteinte' in df_plan_show.columns:
            df_plan_show['% Atteinte'] = df_plan_show['% Atteinte'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
        st.dataframe(df_plan_show)
