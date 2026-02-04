import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

st.title("📊 Dashboard de Pilotage - Randstad / Merck")

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

# --- CHARGEMENT INTELLIGENT ---
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

# Exécution du chargement
data, filename = load_data()

if data is None:
    st.error("❌ Aucun fichier Excel (.xlsx) trouvé sur le serveur.")
    st.info("Veuillez uploader votre fichier de données (ex: data.xlsx) sur GitHub.")
    st.stop()
else:
    st.toast(f"Données chargées depuis : {filename}", icon="✅")
    st.markdown("---")

# --- DASHBOARD ---

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Vue Globale", "🤝 Recrutement", "🏥 Absentéisme", "🔍 Sourcing (Top 5 & TC)", "✅ Plan d'Action"])

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

# --- 4. SOURCING (MODIFIÉ) ---
with tab4:
    st.header("Performance Sourcing")
    if "SOURCE" in data:
        df_src = data["SOURCE"]
        if 'Source' in df_src.columns:
            # 1. Aggrégation
            df_agg = df_src.groupby('Source', as_index=False)[['1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']].sum()
            
            # 2. FOCUS TALENT CENTER
            st.subheader("🔥 Focus : Efficience Talent Center")
            
            # Recherche de la ligne Talent Center
            # On cherche n'importe quelle source contenant "TALENT CENTER"
            mask_tc = df_agg['Source'].astype(str).str.contains("TALENT CENTER", case=False, na=False)
            df_tc = df_agg[mask_tc]
            
            if not df_tc.empty:
                # On somme au cas où il y ait plusieurs lignes Talent Center
                vol_tc = df_tc['1. Appels Reçus'].sum()
                val_tc = df_tc['2. Validés (Sél.)'].sum()
                int_tc = df_tc['3. Intégrés (Délégués)'].sum()
                
                # Calculs Taux
                taux_valid_tc = (val_tc / vol_tc * 100) if vol_tc > 0 else 0
                taux_transfo_tc = (int_tc / vol_tc * 100) if vol_tc > 0 else 0
                
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Volume Appels (TC)", int(vol_tc))
                k2.metric("Validés (TC)", int(val_tc))
                k3.metric("Intégrés (TC)", int(int_tc))
                k4.metric("Rendement Final (TC)", f"{taux_transfo_tc:.2f}%", delta_color="normal")
            else:
                st.info("Source 'Talent Center' non détectée dans les données.")
            
            st.markdown("---")

            # 3. TOP 5 SOURCES
            st.subheader("🏆 Top 5 des Meilleures Sources")
            
            # Tri par Intégrés (Qualité) puis Volume (Quantité)
            df_top5 = df_agg.sort_values(by=['3. Intégrés (Délégués)', '1. Appels Reçus'], ascending=[False, False]).head(5)
            
            # Ajout d'une colonne couleur pour surbrillance
            def get_color(source_name):
                if "TALENT CENTER" in str(source_name).upper():
                    return "Talent Center" # Sera mappé à une couleur vive
                return "Autres Sources"

            df_top5['Type'] = df_top5['Source'].apply(get_color)
            
            # Graphique avec couleur conditionnelle
            fig_src = px.bar(
                df_top5, 
                x='Source', 
                y=['1. Appels Reçus', '3. Intégrés (Délégués)'], 
                barmode='group',
                title="Volume vs Intégration (Top 5)",
                color='Type', # Utilise la colonne Type pour la couleur
                # Dictionnaire de couleurs : Talent Center en Orange Vif, Autres en Gris/Bleu
                color_discrete_map={
                    "Talent Center": "#FF8C00",   # Orange
                    "Autres Sources": "#636EFA",  # Bleu par défaut Plotly
                    "1. Appels Reçus": "#B0C4DE", # Gris clair (si groupé par variable)
                    "3. Intégrés (Délégués)": "#4682B4" # Bleu acier
                }
            )
            # Petit hack pour que Plotly comprenne bien les couleurs quand on a deux barres par source
            # On refait plus simple : Couleur par variable, mais on note TC dans le titre ou annotations si besoin
            # Le plus simple pour garder la lisibilité "Bar Group" est de ne pas colorer par source mais par métrique
            # MAIS on va utiliser une astuce : trier pour que TC soit visible.
            
            # Alternative visuelle plus propre : Bar chart simple des Intégrés avec couleur
            fig_best = px.bar(
                df_top5,
                x='Source',
                y='3. Intégrés (Délégués)',
                color='Type',
                title="Top 5 Sources par Nombre d'Intégrations",
                text='3. Intégrés (Délégués)',
                color_discrete_map={
                    "Talent Center": "#FF4500", # Orange Rouge
                    "Autres Sources": "#1f77b4" # Bleu
                }
            )
            fig_best.update_traces(textposition='outside')
            st.plotly_chart(fig_best, use_container_width=True)
            
            with st.expander("Voir le détail complet des sources"):
                # Formatage tableau
                df_disp_src = df_agg.copy()
                df_disp_src['Rendement (%)'] = (df_disp_src['3. Intégrés (Délégués)'] / df_disp_src['1. Appels Reçus'] * 100).fillna(0).map('{:.2f}%'.format)
                st.dataframe(df_disp_src)

    else:
        st.info("En attente du fichier...")

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
