import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Configuration de la page
st.set_page_config(page_title="Dashboard Pilotage Randstad", layout="wide")

st.title("📊 Dashboard de Pilotage - Randstad / Merck")
st.markdown("---")

# --- FONCTION DE NETTOYAGE ---
def clean_data(df):
    """
    Nettoie les chiffres français (virgule, %) et les convertit en float.
    Ex: "85,50%" -> 85.50
    """
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                # On travaille sur une copie en string
                series = df[col].astype(str).str.replace('"', '').str.strip()
                
                # Cas 1 : Pourcentages (ex: "85,20%")
                if series.str.contains('%').any():
                    series = series.str.replace('%', '').str.replace(',', '.').replace(' ', '')
                    df[col] = pd.to_numeric(series, errors='coerce')
                
                # Cas 2 : Chiffres français (ex: "1 200,50")
                elif series.str.match(r'^-?\d+(?:[\s\u202f]?\d*)*(?:,\d+)?$').any():
                    series = series.str.replace(',', '.').replace(' ', '').replace('\u202f', '')
                    df[col] = pd.to_numeric(series, errors='coerce')
            except Exception:
                pass
    return df

# --- SIDEBAR : UPLOAD ---
st.sidebar.header("📂 Import des Données")
st.sidebar.info("Téléchargez le Google Sheet en format .xlsx et déposez-le ici.")

uploaded_file = st.sidebar.file_uploader("Fichier Dashboard Global (.xlsx)", type="xlsx")

data = {}

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_sheet_names = xls.sheet_names
        
        expected_sheets = {
            "YTD": "CONSOLIDATION_YTD",
            "RECRUT": "Recrutement_Mensuel",
            "ABS": "Absentéisme_Global_Mois",
            "SOURCE": "KPI_Sourcing_Rendement",
            "PLAN": "Suivi_Plan_Action"
        }
        
        for key, sheet_name in expected_sheets.items():
            if sheet_name in all_sheet_names:
                # On lit et on nettoie tout de suite
                df_loaded = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                data[key] = clean_data(df_loaded)
            else:
                st.sidebar.warning(f"⚠️ Onglet manquant : {sheet_name}")
                
        st.sidebar.success("Données chargées et formatées !")
        
    except Exception as e:
        st.sidebar.error(f"Erreur de lecture : {e}")

# --- DASHBOARD ---

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📈 Vue Globale (YTD)", "🤝 Recrutement", "🏥 Absentéisme", "🔍 Sourcing", "✅ Plan d'Action"])

# 1. VUE GLOBALE (YTD)
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
                
                # FORMATAGE EN %
                # Si c'est un chiffre (float), on ajoute %
                if isinstance(val, (int, float)):
                    val_str = f"{val:.2f}%"
                else:
                    val_str = str(val) # On laisse tel quel si c'est du texte
                
                col_idx = index % 4
                cols[col_idx].metric(label=indic, value=val_str)
        else:
            st.error("Colonne 'Valeur YTD' introuvable.")
    else:
        st.info("En attente du fichier...")

# 2. RECRUTEMENT
with tab2:
    st.header("Performance Recrutement Mensuel")
    if "RECRUT" in data:
        df_rec = data["RECRUT"]
        if 'Année' in df_rec.columns and 'Mois' in df_rec.columns:
            df_rec = df_rec.sort_values(['Année', 'Mois'])
            
            # Création colonne Période pour l'axe X (plus propre)
            df_rec['Période'] = df_rec['Mois'].astype(str) + "/" + df_rec['Année'].astype(str)

            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Taux de Service & Transformation")
                fig_taux = px.line(df_rec, x='Période', y=['Taux Service', 'Taux Transfo'], 
                                   markers=True, title="Evolution des Taux")
                # AJOUT DU FORMAT % SUR L'AXE Y
                fig_taux.update_layout(yaxis_ticksuffix="%")
                st.plotly_chart(fig_taux, use_container_width=True)

            with col2:
                st.subheader("Volumes")
                fig_vol = px.bar(df_rec, x='Période', y=['Nb Requisitions', 'Nb Hired'],
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
             df_abs['Période'] = df_abs['Mois'].astype(str) + "/" + df_abs['Année'].astype(str)

        fig_abs = px.area(df_abs, x='Période', y='Taux Absentéisme', 
                          title="Taux d'Absentéisme Global", markers=True,
                          color_discrete_sequence=['#FF5733'])
        # AJOUT DU FORMAT % SUR L'AXE Y
        fig_abs.update_layout(yaxis_ticksuffix="%")
        st.plotly_chart(fig_abs, use_container_width=True)
        
        kpi1, kpi2 = st.columns(2)
        if 'Taux Absentéisme' in df_abs.columns:
             val_mean = df_abs['Taux Absentéisme'].mean()
             # Formatage KPI
             kpi1.metric("Moyenne Annuelle", f"{val_mean:.2f}%")
        
        kpi1.dataframe(df_abs)
    else:
        st.info("En attente du fichier...")

# 4. SOURCING
with tab4:
    st.header("Entonnoir de Sourcing")
    if "SOURCE" in data:
        df_source = data["SOURCE"]
        
        if 'Source' in df_source.columns:
            # On aggrège
            df_agg = df_source.groupby('Source')[['1. Appels Reçus', '3. Intégrés (Délégués)']].sum().reset_index()
            
            # Calcul du rendement global pour le tri
            df_agg['Rendement'] = df_agg['3. Intégrés (Délégués)'] / df_agg['1. Appels Reçus']
            df_agg = df_agg.sort_values('1. Appels Reçus', ascending=False) # Tri par volume

            st.subheader("Efficacité par Canal (Volume vs Intégration)")
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
        
        row_global = df_plan[df_plan['Catégorie / Section'].astype(str).str.contains('GLOBAL', case=False, na=False)]
        
        if not row_global.empty:
            taux = row_global.iloc[0]['% Atteinte']
            # Formatage Jauge
            fig_gauge = go.Figure(go.Indicator(
                mode = "gauge+number",
                value = taux if isinstance(taux, (int, float)) else 0,
                number = {'suffix': "%"}, # AJOUT DU SUFFIXE %
                title = {'text': "Avancement Global"},
                gauge = {'axis': {'range': [None, 100]}, 'bar': {'color': "green"}}
            ))
            st.plotly_chart(fig_gauge, use_container_width=True)
        
        # Formatage du tableau détail pour afficher %
        st.subheader("Détail par chantier")
        # On peut reformater la colonne % Atteinte pour l'affichage tableau
        df_display = df_plan.copy()
        if '% Atteinte' in df_display.columns:
            df_display['% Atteinte'] = df_display['% Atteinte'].apply(lambda x: f"{x:.2f}%" if isinstance(x, (int, float)) else x)
            
        st.dataframe(df_display)
    else:
        st.info("En attente du fichier...")
