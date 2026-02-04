import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import glob

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

st.title("📊 Dashboard de Pilotage - Randstad / Merck")

# --- OUTIL DE DEBUG (Visible seulement si besoin) ---
with st.expander("🛠️ Zone Admin & Debug (Cliquer pour voir les fichiers)"):
    st.write("Dossier actuel :", os.getcwd())
    st.write("Fichiers détectés :", os.listdir('.'))

# --- FONCTION DE NETTOYAGE ---
def clean_and_scale_data(df):
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                series = df[col].astype(str).str.replace('"', '').str.strip()
                series = series.str.replace('%', '').str.replace(' ', '').str.replace('\u202f', '')
                series = series.str.replace(',', '.')
                df[col] = pd.to_numeric(series, errors='ignore')
            except Exception:
                pass

    for col in df.columns:
        col_lower = col.lower()
        if any(x in col_lower for x in ['taux', '%', 'atteinte', 'validation', 'rendement']):
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
    return df

# --- LOGIQUE DE CHARGEMENT ---
# 1. On essaie de trouver un fichier Excel automatiquement
excel_files = glob.glob("*.xlsx")
data = {}
file_source = "Aucun"

# Si on trouve un fichier sur le serveur, on le charge
if len(excel_files) > 0:
    file_path = excel_files[0]
    file_source = f"Automatique ({file_path})"
    try:
        xls = pd.ExcelFile(file_path)
    except Exception as e:
        st.error(f"Erreur lecture auto : {e}")
        xls = None
else:
    xls = None

# 2. Si pas de fichier auto (ou erreur), on affiche l'upload manuel
if xls is None:
    st.warning("⚠️ Aucun fichier Excel trouvé automatiquement sur le serveur.")
    uploaded_file = st.file_uploader("📂 Chargez votre fichier Excel manuellement (Secours)", type="xlsx")
    if uploaded_file:
        file_source = "Manuel (Upload)"
        xls = pd.ExcelFile(uploaded_file)
    else:
        st.stop() # On arrête tout ici si on n'a rien

# 3. Traitement du fichier (qu'il vienne du serveur ou de l'upload)
if xls:
    try:
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
                df_raw = pd.read_excel(xls, sheet_name=sheet_name)
                data[key] = clean_and_scale_data(df_raw)
        
        st.toast(f"Source des données : {file_source}", icon="✅")
        
    except Exception as e:
        st.error(f"Erreur critique lors du traitement : {e}")
        st.stop()

st.markdown("---")

# --- DASHBOARD ---

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

# --- 4. SOURCING (SURBRILLANCE ORANGE + FOCUS TC) ---
with tab4:
    st.header("Performance Sourcing")
    if "SOURCE" in data:
        df_src = data["SOURCE"]
        if 'Source' in df_src.columns:
            df_agg = df_src.groupby('Source', as_index=False)[['1. Appels Reçus', '2. Validés (Sél.)', '3. Intégrés (Délégués)']].sum()
            
            # FOCUS TC
            st.subheader("🔥 Focus : Efficience Talent Center")
            mask_tc = df_agg['Source'].astype(str).str.upper().str.contains("TALENT CENTER")
            df_tc = df_agg[mask_tc]
            
            if not df_tc.empty:
                vol_tc = df_tc['1. Appels Reçus'].sum()
                val_tc = df_tc['2. Validés (Sél.)'].sum()
                int_tc = df_tc['3. Intégrés (Délégués)'].sum()
                taux_transfo_tc = (int_tc / vol_tc * 100) if vol_tc > 0 else 0
                
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Volume Appels (TC)", int(vol_tc))
                k2.metric("Validés (TC)", int(val_tc))
                k3.metric("Intégrés (TC)", int(int_tc))
                k4.metric("Rendement Final (TC)", f"{taux_transfo_tc:.2f}%", delta_color="normal")
            
            st.markdown("---")

            # TOP 5 COLORÉ
            st.subheader("🏆 Top 5 des Meilleures Sources (Intégration)")
            df_top5 = df_agg.sort_values(by=['3. Intégrés (Délégués)', '1. Appels Reçus'], ascending=[False, False]).head(5)
            
            def categorize_source(source_name):
                return "Talent Center" if "TALENT CENTER" in str(source_name).upper() else "Autres Sources"

            df_top5['Catégorie'] = df_top5['Source'].apply(categorize_source)
            
            fig_best = px.bar(
                df_top5,
                x='Source',
                y='3. Intégrés (Délégués)',
                color='Catégorie',
                title="Nombre d'Intégrations par Source",
                text='3. Intégrés (Délégués)',
                color_discrete_map={
                    "Talent Center": "#FF4500", # Orange
                    "Autres Sources": "#1f77b4" # Bleu
                }
            )
            fig_best.update_traces(textposition='outside')
            st.plotly_chart(fig_best, use_container_width=True)
            
            with st.expander("Voir le tableau complet"):
                df_disp_src = df_agg.copy()
                df_disp_src['Rendement (%)'] = (df_disp_src['3. Intégrés (Délégués)'] / df_disp_src['1. Appels Reçus'] * 100).fillna(0).map('{:.2f}%'.format)
                st.dataframe(df_disp_src)

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
