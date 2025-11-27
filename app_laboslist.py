import pandas as pd
import streamlit as st
import io
from datetime import datetime

# Configuration de la page
st.set_page_config(
    page_title="Extraction Laboratoires",
    page_icon="🔬",
    layout="wide"
)

def charger_donnees_exemple():
    """
    Fonction pour charger des données d'exemple
    Remplacez cette fonction par votre vraie méthode de chargement
    """
    # Création de données d'exemple - À REMPLACER PAR VOS VRAIES DONNÉES
    data = {
        'Nom du Laboratoire': ['Lab Central', 'Lab Ville', 'Lab Régional', 'Lab National'],
        'Pays': ['Pays A', 'Pays A', 'Pays B', 'Pays B'],
        'Ville /District (Département)': ['Ville 1', 'Ville 2', 'Ville 3', 'Ville 4'],
        'Adresse de la structure sanitaire': ['Adresse 1', 'Adresse 2', 'Adresse 3', 'Adresse 4'],
        'Sélectionnez le niveau de complexité du laboratoire': ['Level I', 'Level II', 'Level III', 'Level IV'],
        "Tests rapides d'anticorps contre le VIH": ['yes', 'no', 'yes', 'yes'],
        "Chaîne ELISA (Enzyme-Linked Immunosorbent Assay)": ['no', 'yes', 'yes', 'yes'],
        "Virus Ebola": ['notavailable', 'serology', 'notavailable', 'viralload'],
        "Test de Widal (typhoïde)": ['yes', 'yes', 'no', 'yes'],
        'Coordonnées GPS': ['GPS1', 'GPS2', 'GPS3', 'GPS4'],
        'Numéro de téléphone personnel du répondant': ['123', '456', '789', '012']
    }
    return pd.DataFrame(data)

def filtrer_laboratoires(df, niveau, variable, modalite):
    """
    Filtre les laboratoires selon les critères spécifiés
    """
    try:
        # Filtrer par niveau
        df_filtre = df[df['Sélectionnez le niveau de complexité du laboratoire'] == niveau]
        
        # Filtrer par variable et modalité
        if variable in df_filtre.columns:
            df_filtre = df_filtre[df_filtre[variable] == modalite]
        else:
            st.error(f"❌ La variable '{variable}' n'existe pas dans la base de données")
            return None
        
        return df_filtre
    except Exception as e:
        st.error(f"❌ Erreur lors du filtrage: {str(e)}")
        return None

def generer_excel(df_filtre, niveau, variable, modalite):
    """
    Génère un fichier Excel en mémoire
    """
    if df_filtre.empty:
        return None
    
    try:
        # Sélectionner les colonnes à exporter
        colonnes_base = [
            'Nom du Laboratoire', 
            'Pays',
            'Ville /District (Département)',
            'Adresse de la structure sanitaire',
            'Sélectionnez le niveau de complexité du laboratoire',
            variable
        ]
        
        # Colonnes supplémentaires optionnelles
        colonnes_optionnelles = [
            'Coordonnées GPS', 
            'Numéro de téléphone personnel du répondant',
            'Adresse électronique du répondant',
            'Fonction du répondant'
        ]
        
        # Garder seulement les colonnes disponibles
        colonnes_finales = [col for col in colonnes_base if col in df_filtre.columns]
        for col in colonnes_optionnelles:
            if col in df_filtre.columns:
                colonnes_finales.append(col)
        
        # Créer le fichier Excel en mémoire
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_filtre[colonnes_finales].to_excel(writer, sheet_name='Laboratoires', index=False)
            
            # Formater le fichier Excel
            workbook = writer.book
            worksheet = writer.sheets['Laboratoires']
            
            # Ajouter un en-tête
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Appliquer le format aux en-têtes
            for col_num, value in enumerate(df_filtre[colonnes_finales].columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Ajuster la largeur des colonnes
            for i, col in enumerate(df_filtre[colonnes_finales].columns):
                max_len = max(df_filtre[col].astype(str).str.len().max(), len(col)) + 2
                worksheet.set_column(i, i, min(max_len, 50))
        
        output.seek(0)
        return output
    
    except Exception as e:
        st.error(f"❌ Erreur lors de la génération du fichier Excel: {str(e)}")
        return None

def main():
    st.title("🔬 Extraction des Laboratoires par Critères")
    st.markdown("Sélectionnez les critères pour extraire la liste des laboratoires")
    
    # Sidebar pour les paramètres
    st.sidebar.header("📊 Paramètres de sélection")
    
    # Charger les données
    try:
        # REMPLACEZ CETTE LIGNE PAR LE CHARGEMENT DE VOS VRAIES DONNÉES
        #df = charger_donnees_exemple()
        df=pd.read_excel('Data_LabMab_2025_merge_final_LabMab_29_09_2025_for_R_.xlsx')
        
        # Option pour charger un fichier personnalisé
        st.sidebar.markdown("---")
        st.sidebar.subheader("📁 Charger vos données")
        fichier_upload = st.sidebar.file_uploader(
            "Téléchargez votre fichier Excel", 
            type=['xlsx', 'xls'],
            help="Chargez votre base de données des laboratoires"
        )
        
        if fichier_upload is not None:
            try:
                df = pd.read_excel(fichier_upload)
                st.sidebar.success("✅ Fichier chargé avec succès!")
            except Exception as e:
                st.sidebar.error(f"❌ Erreur lors du chargement: {str(e)}")
                st.sidebar.info("Utilisation des données d'exemple")
        
        # Sélection du niveau
        niveaux = sorted(df['Sélectionnez le niveau de complexité du laboratoire'].dropna().unique())
        niveau_selectionne = st.sidebar.selectbox(
            "🎯 Sélectionnez le niveau:",
            options=niveaux,
            index=0 if niveaux else None,
            help="Choisissez le niveau de complexité du laboratoire"
        )
        
        # Sélection de la variable
        variables = sorted(df.columns.tolist())
        variable_selectionnee = st.sidebar.selectbox(
            "📋 Sélectionnez la variable:",
            options=variables,
            index=variables.index("Tests rapides d'anticorps contre le VIH") if "Tests rapides d'anticorps contre le VIH" in variables else 0,
            help="Choisissez la variable à filtrer"
        )
        
        # Sélection de la modalité
        if variable_selectionnee:
            modalites = sorted(df[variable_selectionnee].dropna().unique())
            if modalites:
                modalite_selectionnee = st.sidebar.selectbox(
                    "🔍 Sélectionnez la modalité:",
                    options=modalites,
                    index=0,
                    help="Choisissez la valeur de la variable"
                )
            else:
                st.sidebar.warning("⚠️ Aucune modalité disponible pour cette variable")
                modalite_selectionnee = None
        
        # Bouton pour générer l'extraction
        if st.sidebar.button("🚀 Générer l'extraction", type="primary", use_container_width=True):
            if modalite_selectionnee is not None:
                with st.spinner("🔍 Extraction en cours..."):
                    # Filtrer les données
                    df_filtre = filtrer_laboratoires(df, niveau_selectionne, variable_selectionnee, modalite_selectionnee)
                    
                    if df_filtre is not None and not df_filtre.empty:
                        # Afficher les statistiques
                        st.success(f"✅ **{len(df_filtre)}** laboratoire(s) trouvé(s)")
                        
                        # Aperçu des données
                        st.subheader("👀 Aperçu des données")
                        
                        # Colonnes pour l'aperçu
                        colonnes_apercu = ['Nom du Laboratoire', 'Pays', 'Ville /District (Département)']
                        colonnes_disponibles = [col for col in colonnes_apercu if col in df_filtre.columns]
                        
                        st.dataframe(
                            df_filtre[colonnes_disponibles], 
                            use_container_width=True,
                            height=200
                        )
                        
                        # Générer le fichier Excel
                        excel_file = generer_excel(df_filtre, niveau_selectionne, variable_selectionnee, modalite_selectionnee)
                        
                        if excel_file:
                            # Nom du fichier
                            nom_fichier = f"Laboratoires_{niveau_selectionne.replace(' ', '_')}_{variable_selectionnee[:20]}_{modalite_selectionnee}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                            
                            # Bouton de téléchargement
                            st.download_button(
                                label="📥 Télécharger le fichier Excel",
                                data=excel_file,
                                file_name=nom_fichier,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True
                            )
                            
                            # Informations supplémentaires
                            with st.expander("📊 Détails de l'extraction"):
                                st.write(f"""
                                **Résumé de l'extraction:**
                                - **Niveau:** {niveau_selectionne}
                                - **Variable:** {variable_selectionnee}
                                - **Modalité:** {modalite_selectionnee}
                                - **Total laboratoires:** {len(df_filtre)}
                                - **Date:** {datetime.now().strftime('%d/%m/%Y %H:%M')}
                                """)
                        
                    else:
                        st.error("❌ Aucun laboratoire trouvé avec ces critères")
            else:
                st.sidebar.error("⚠️ Veuillez sélectionner une modalité valide")
        
        # Section pour les analyses rapides
        st.sidebar.markdown("---")
        st.sidebar.subheader("⚡ Analyses rapides")
        
        col1, col2 = st.sidebar.columns(2)
        
        with col1:
            if st.button("Tests VIH ✅", use_container_width=True):
                st.session_state.niveau = "Level I"
                st.session_state.variable = "Tests rapides d'anticorps contre le VIH"
                st.session_state.modalite = "yes"
                st.rerun()
        
        with col2:
            if st.button("Tests ELISA ✅", use_container_width=True):
                st.session_state.niveau = "Level I" 
                st.session_state.variable = "Chaîne ELISA (Enzyme-Linked Immunosorbent Assay)"
                st.session_state.modalite = "yes"
                st.rerun()
        
        # Informations sur les données
        st.sidebar.markdown("---")
        st.sidebar.subheader("ℹ️ Informations")
        st.sidebar.write(f"**Total laboratoires:** {len(df)}")
        st.sidebar.write(f"**Niveaux disponibles:** {', '.join(niveaux)}")
        
    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des données: {str(e)}")
        st.info("""
        **Mode d'utilisation de cette application:**
        
        1. **Chargez vos données** via le menu dans la sidebar, ou
        2. **Review**  charger le fichier directement:
           ```python
           df = pd.read_excel('Data_LabMab_2025_merge_final_LabMab_29_09_2025_for_R_.xlsx')
           ```
        3. **Vérification** Est-ce que la  DataFrame contient la colonne:
           - 'Sélectionnez le niveau de complexité du laboratoire'
        """)

if __name__ == "__main__":
    main()