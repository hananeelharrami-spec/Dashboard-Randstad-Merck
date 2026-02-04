import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

st.title("📊 Dashboard de Pilotage - Randstad / Merck")
st.markdown("---")

# --- FONCTION DE NETTOYAGE RENFORCÉE (LE COEUR DU FIX) ---
def clean_and_scale_data(df):
    # 1. CONVERSION TEXTE -> NOMBRE
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                # Nettoyage agressif des caractères parasites
                series = df[col].astype(str).str.replace('"', '').str.strip()
                series = series.str.replace('%', '').str.replace(' ', '').str.replace('\u202f', '') # Espace insécable
                series = series.str.replace(',', '.') # Virgule française
                
                # Conversion en numérique
                df[col] = pd.to_numeric(series, errors='ignore')
            except Exception:
                pass

    # 2. MISE A L'ECHELLE DES POURCENTAGES (0.88 -> 88.0)
    for col in df.columns:
        col_lower = col.lower()
        # Si le nom de la colonne suggère un taux
        if any(x in col_lower for x in ['taux', '%', 'atteinte', 'validation', 'rendement']):
            # On vérifie que la colonne est bien numérique maintenant
            if pd.api.types.is_numeric_dtype(df[col]):
                max_val = df[col].max()
                # LOGIQUE DE SECURITE :
                # Si le max est petit (<= 1.5), c'est un ratio (ex: 0.88). On multiplie par 100.
                # On met 1.5 pour gérer les cas où on dépasse un peu 100% (ex: 1.1 pour 110%)
                if pd.notna(max_val) and -1.5 <= max_val <= 1.5 and max_val != 0:
                    df[col] = df[col] * 100
    
    return df

# --- SIDEBAR : UPLOAD ---
st.sidebar.header("📂 Import des Données")
st.sidebar.info("Déposez le fichier Excel (.xlsx) contenant tous les onglets.")

uploaded_file = st.sidebar.file_uploader("Fichier Dashboard Global", type="xlsx")

data = {}

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_sheets = xls.sheet_names
        
        # Mapping des onglets
        expected = {
            "YTD": "CONSOLIDATION_YTD",
            "RECRUT": "Recrutement_Mensuel",
            "ABS": "Absentéisme_Global_Mois",
            "SOURCE": "KPI_Sourcing_Rendement",
            "PLAN": "Suivi_Plan_Action"
        }
        
        for key, sheet_name in expected.items():
            if sheet_name in all_sheets:
                df_raw = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                # APPLICATION DU FIX
                data[key] = clean_and_scale_data(df_raw)
            else:
                st.sidebar.warning(f"⚠️ Onglet manquant : {sheet_name}")
                
        st.sidebar.success("Données chargées et KPI normalisés (0.88 -> 88%)")
        
    except Exception as e:
        st.sidebar.error(f"Erreur critique : {e}")

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
                
                # Affichage formaté
                val_str = f"{val:.2f}%" if isinstance(val, (int, float)) else str(val)
                
                cols[index % 4].metric(label=indic, value=val_str)
        else:
            st.error(f"Colonne '{col_val}' introuvable.")

# --- 2. RECRUTEMENT ---
with tab2:
    st.header("Recrutement Mensuel")
    if "RECRUT" in data:
        df_rec = data["RECRUT"]
        # Création axe temps
        if 'Mois' in df_rec.columns:
            # Tri
            if 'Année' in df_rec.columns:
                df_rec = df_rec.sort_values(['Année', 'Mois'])
                df_rec['Période'] = df_rec['Mois'].astype(str) + "/" + df_rec['Année'].astype(str)
            else:
                df_rec['Période'] = df_rec['Mois'].astype(str)

            c1, c2 = st.columns(2)
            with c1:
                st.subheader("Taux de Service & Transformation")
                fig = px.line(df_rec, x='Période', y=['Taux Service', 'Taux Transfo'], markers=True)
                fig.update_layout(yaxis_ticksuffix="%") # Force le % sur l'axe
                st.plotly_chart(fig, use_container_width=True)

            with c2:
                st.subheader("Volume Commandes vs Hired")
                fig_bar = px.bar(df_rec, x='Période', y=['Nb Requisitions', 'Nb Hired'], barmode='group')
                st.plotly_chart(fig_bar, use_container_width=True)
                
            # Tableau de données formaté
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
        
        # Tableau
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
            # Aggrégation
            df_agg = df_src.groupby('Source', as_index=False)[['1. Appels Reçus', '3. Intégrés (Délégués)']].sum()
            df_agg = df_agg.sort_values('1. Appels Reçus', ascending=False)
            
            st.subheader("Volume vs Intégration par Source")
            fig_src = px.bar(df_agg, x='Source', y=['1. Appels Reçus', '3. Intégrés (Délégués)'], barmode='group')
            st.plotly_chart(fig_src, use_container_width=True)
            
            st.write("Détail Mensuel (Formaté)")
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
        
        # Jauge
        row_global = df_plan[df_plan['Catégorie / Section'].astype(str).str.contains('GLOBAL', case=False, na=False)]
        if not row_global.empty:
            val = row_global.iloc[0]['% Atteinte']
            # Normalement déjà converti en 0-100 par notre fonction clean_and_scale_data
            fig_gauge = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = val,
                number = {'suffix': "%"},
                title = {'text': "Avancement Global"},
                gauge = {'axis': {'range': [None, 100]}, 'bar': {'color': "green"}}
            ))
            st.plotly_chart(fig_gauge, use_container_width=True)
        
        # Tableau
        df_plan_show = df_plan.copy()
        if '% Atteinte' in df_plan_show.columns:
            df_plan_show['% Atteinte'] = df_plan_show['% Atteinte'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
        st.dataframe(df_plan_show)
