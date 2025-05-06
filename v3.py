import streamlit as st
import pandas as pd
import numpy as np
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import tempfile
import os
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Configuration de la page
st.set_page_config(
    page_title="Évaluation Diversité et Inclusion V3",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Styles pour le PDF
styles = getSampleStyleSheet()
title_style = ParagraphStyle(
    'CustomTitle',
    parent=styles['Heading1'],
    fontSize=24,
    spaceAfter=30,
    textColor=colors.HexColor('#4472C4')
)
heading_style = ParagraphStyle(
    'CustomHeading',
    parent=styles['Heading2'],
    fontSize=16,
    spaceAfter=12,
    textColor=colors.HexColor('#4472C4')
)

# Définitions des indicateurs
INDICATEURS = {
    'taux_femmes': {
        'nom': 'Taux de féminisation',
        'description': 'Proportion de femmes dans l\'effectif total',
        'objectif': 0.4,
        'unite': '%',
        'formule': 'Nombre de femmes / Effectif total'
    },
    'taux_femmes_cadres': {
        'nom': 'Taux de femmes cadres',
        'description': 'Proportion de femmes parmi les cadres',
        'objectif': 0.4,
        'unite': '%',
        'formule': 'Nombre de femmes cadres / Nombre total de cadres'
    },
    'taux_handicap': {
        'nom': 'Taux de handicap',
        'description': 'Proportion de personnes en situation de handicap',
        'objectif': 0.06,
        'unite': '%',
        'formule': 'Nombre de personnes en situation de handicap / Effectif total'
    },
    'ecart_salaire': {
        'nom': 'Écart salarial',
        'description': 'Écart de rémunération entre hommes et femmes',
        'objectif': 0.05,
        'unite': '%',
        'formule': '(Salaire moyen hommes - Salaire moyen femmes) / Salaire moyen hommes'
    },
    'score_ages': {
        'nom': 'Score âges',
        'description': 'Indicateur de diversité des âges',
        'objectif': 0.7,
        'unite': '%',
        'formule': 'Score calculé sur la base de la répartition par âge'
    },
    'taux_cdi': {
        'nom': 'Taux de CDI',
        'description': 'Proportion de contrats à durée indéterminée',
        'objectif': 0.8,
        'unite': '%',
        'formule': 'Nombre de CDI / Effectif total'
    },
    'taux_formation': {
        'nom': 'Taux de formation',
        'description': 'Proportion d\'employés formés',
        'objectif': 0.05,
        'unite': '%',
        'formule': 'Nombre d\'employés formés / Effectif total'
    },
    'taux_recrutement_interne': {
        'nom': 'Taux de recrutement interne',
        'description': 'Proportion de recrutements internes',
        'objectif': 0.3,
        'unite': '%',
        'formule': 'Nombre de recrutements internes / Nombre total de recrutements'
    },
    'taux_temps_partiel': {
        'nom': 'Taux de temps partiel',
        'description': 'Proportion d\'employés à temps partiel',
        'objectif': 0.2,
        'unite': '%',
        'formule': 'Nombre d\'employés à temps partiel / Effectif total'
    },
    'taux_teletravail': {
        'nom': 'Taux de télétravail',
        'description': 'Proportion d\'employés en télétravail',
        'objectif': 0.2,
        'unite': '%',
        'formule': 'Nombre d\'employés en télétravail / Effectif total'
    },
    'taux_promotion_femmes': {
        'nom': 'Taux de promotion des femmes',
        'description': 'Proportion de femmes parmi les promotions',
        'objectif': 0.4,
        'unite': '%',
        'formule': 'Nombre de promotions femmes / Nombre total de promotions'
    }
}

def create_age_distribution_chart(df_age):
    fig = px.bar(
        df_age[:-1],  # Exclure la ligne Total
        x='Tranche d\'âge',
        y='Effectif',
        title='Répartition des effectifs par âge',
        color='Tranche d\'âge',
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    fig.update_layout(
        xaxis_title="Tranche d'âge",
        yaxis_title="Nombre d'employés",
        showlegend=False
    )
    return fig

def create_salary_gap_chart(df_remuneration):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Hommes',
        x=df_remuneration['Catégorie'],
        y=df_remuneration['Rémunération moyenne hommes'],
        marker_color='#4472C4'
    ))
    fig.add_trace(go.Bar(
        name='Femmes',
        x=df_remuneration['Catégorie'],
        y=df_remuneration['Rémunération moyenne femmes'],
        marker_color='#ED7D31'
    ))
    fig.update_layout(
        title='Écarts salariaux par catégorie',
        xaxis_title="Catégorie",
        yaxis_title="Rémunération moyenne (€)",
        barmode='group'
    )
    return fig

def create_indicators_radar(data):
    categories = ['Féminisation', 'Femmes cadres', 'Handicap', 'Âges', 'CDI', 'Formation', 'Recrutement interne']
    values = [
        data['taux_femmes'] * 100,
        data['taux_femmes_cadres'] * 100,
        data['taux_handicap'] * 100,
        data['score_ages'] * 100,
        data['taux_cdi'] * 100,
        data['taux_formation'] * 100,
        data['taux_recrutement_interne'] * 100
    ]
    
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=values,
        theta=categories,
        fill='toself',
        name='Valeurs actuelles'
    ))
    
    fig.update_layout(
        polar=dict(
            radialaxis=dict(
                visible=True,
                range=[0, 100]
            )
        ),
        showlegend=False,
        title='Indicateurs clés de diversité et inclusion'
    )
    return fig

def generate_pdf(data, company_name, year):
    # Création d'un fichier temporaire pour le PDF
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        doc = SimpleDocTemplate(tmp.name, pagesize=letter)
        elements = []
        
        # Titre
        elements.append(Paragraph(f"Rapport d'Évaluation Diversité et Inclusion<br/>{company_name} - {year}", title_style))
        elements.append(Spacer(1, 20))
        
        # Résumé exécutif
        elements.append(Paragraph("Résumé Exécutif", heading_style))
        elements.append(Paragraph(
            f"Ce rapport présente une analyse détaillée de la diversité et de l'inclusion au sein de {company_name} "
            f"pour l'année {year}. L'évaluation couvre plusieurs dimensions clés : la parité, l'inclusion des personnes "
            "en situation de handicap, l'équité salariale, la diversité des âges, et les pratiques de formation et de "
            "recrutement. Les résultats sont comparés aux objectifs recommandés et aux seuils légaux pour fournir une "
            "vue d'ensemble complète de la situation actuelle et des axes d'amélioration.",
            styles['Normal']
        ))
        elements.append(Spacer(1, 20))
        
        # Score global
        elements.append(Paragraph("Score Global", heading_style))
        score_data = [
            ['Indicateur', 'Valeur', 'Objectif', 'Écart', 'Explication'],
            ['Taux de féminisation', f"{data['taux_femmes']:.1%}", '40%', f"{data['taux_femmes']*100-40:+.1f}%", data['explications']['taux_femmes']],
            ['Taux de femmes cadres', f"{data['taux_femmes_cadres']:.1%}", '40%', f"{data['taux_femmes_cadres']*100-40:+.1f}%", data['explications']['taux_femmes_cadres']],
            ['Taux de handicap', f"{data['taux_handicap']:.1%}", '6%', f"{data['taux_handicap']*100-6:+.1f}%", data['explications']['taux_handicap']],
            ['Écart salarial', f"{data['ecart_salaire']:.1%}", '<5%', f"{data['ecart_salaire']*100-5:+.1f}%", data['explications']['ecart_salaire']],
            ['Score âges', f"{data['score_ages']:.1%}", '>70%', f"{data['score_ages']*100-70:+.1f}%", data['explications']['score_ages']],
            ['Taux de CDI', f"{data['taux_cdi']:.1%}", '>80%', f"{data['taux_cdi']*100-80:+.1f}%", data['explications']['taux_cdi']],
            ['Taux de formation', f"{data['taux_formation']:.1%}", '>5%', f"{data['taux_formation']*100-5:+.1f}%", data['explications']['taux_formation']],
            ['Taux de recrutement interne', f"{data['taux_recrutement_interne']:.1%}", '>30%', f"{data['taux_recrutement_interne']*100-30:+.1f}%", data['explications']['taux_recrutement_interne']],
            ['Taux de temps partiel', f"{data['taux_temps_partiel']:.1%}", '<20%', f"{data['taux_temps_partiel']*100-20:+.1f}%", data['explications']['taux_temps_partiel']],
            ['Taux de télétravail', f"{data['taux_teletravail']:.1%}", '>20%', f"{data['taux_teletravail']*100-20:+.1f}%", data['explications']['taux_teletravail']],
            ['Taux de promotion des femmes', f"{data['taux_promotion_femmes']:.1%}", '>40%', f"{data['taux_promotion_femmes']*100-40:+.1f}%", data['explications']['taux_promotion_femmes']]
        ]
        
        table = Table(score_data, colWidths=[2*inch, 1*inch, 1*inch, 1*inch, 3*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 20))
        
        # Points forts
        elements.append(Paragraph("Points Forts", heading_style))
        for point in data['points_forts']:
            elements.append(Paragraph(f"• {point}", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Axes d'amélioration
        elements.append(Paragraph("Axes d'Amélioration", heading_style))
        for point in data['axes_amelioration']:
            elements.append(Paragraph(f"• {point}", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Plan d'action
        elements.append(Paragraph("Plan d'Action Recommandé", heading_style))
        action_data = [
            ['Action', 'Priorité', 'Délai', 'Impact attendu'],
            ['Mettre en place un programme de mentorat', 'Haute', '3 mois', 'Amélioration de la progression des femmes'],
            ['Renforcer la politique de recrutement des personnes en situation de handicap', 'Haute', '6 mois', 'Atteinte du seuil légal de 6%'],
            ['Réaliser un audit des écarts salariaux', 'Haute', '3 mois', 'Réduction des écarts de rémunération'],
            ['Développer la formation continue', 'Moyenne', '6 mois', 'Amélioration des compétences et de la mobilité'],
            ['Mettre en place un programme de télétravail', 'Moyenne', '3 mois', 'Amélioration de la qualité de vie au travail'],
            ['Créer un comité diversité et inclusion', 'Moyenne', '1 mois', 'Suivi et pilotage des actions']
        ]
        
        table = Table(action_data, colWidths=[3*inch, 1*inch, 1*inch, 2*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(table)
        elements.append(Spacer(1, 20))
        
        # Conclusion
        elements.append(Paragraph("Conclusion", heading_style))
        elements.append(Paragraph(
            f"L'analyse des données de {company_name} révèle des opportunités significatives d'amélioration en matière "
            "de diversité et d'inclusion. Les actions recommandées, si elles sont mises en œuvre de manière cohérente "
            "et suivie, permettront de progresser vers une organisation plus inclusive et équitable. Il est recommandé "
            "de réaliser une nouvelle évaluation dans 12 mois pour mesurer les progrès accomplis.",
            styles['Normal']
        ))
        
        # Génération du PDF
        doc.build(elements)
        return tmp.name

# Interface Streamlit
st.title("📊 Évaluation Diversité et Inclusion V3")

# Sidebar
st.sidebar.title("Navigation")
page = st.sidebar.radio("Choisissez une section", ["Accueil", "Import des données", "Saisie manuelle", "Résultats", "Rapport", "Définitions"])

if page == "Accueil":
    st.header("Bienvenue dans l'application d'évaluation de la diversité et de l'inclusion")
    
    st.write("""
    Cette application vous permet de :
    1. Télécharger un modèle Excel pour saisir vos données
    2. Importer vos données complétées
    3. Visualiser les résultats en temps réel
    4. Générer un rapport PDF détaillé
    """)
    
    st.header("1. Téléchargement du modèle")
    st.write("Commencez par télécharger le modèle Excel pour y saisir vos données.")
    with open("modele_bilan_social_v3.xlsx", "rb") as file:
        st.download_button(
            label="📥 Télécharger le modèle Excel",
            data=file,
            file_name="modele_bilan_social_v3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

elif page == "Import des données":
    st.header("2. Import des données")
    st.write("Une fois le modèle complété, importez-le ici pour générer votre rapport.")
    
    uploaded_file = st.file_uploader("Choisissez votre fichier Excel", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # Lecture des données
            df_general = pd.read_excel(uploaded_file, sheet_name='Données générales')
            df_age = pd.read_excel(uploaded_file, sheet_name='Répartition par âge')
            df_remuneration = pd.read_excel(uploaded_file, sheet_name='Rémunérations')
            df_formation = pd.read_excel(uploaded_file, sheet_name='Formation et Recrutement')
            df_calculs = pd.read_excel(uploaded_file, sheet_name='Calculs automatiques')
            
            # Stockage des données dans la session
            st.session_state['df_general'] = df_general
            st.session_state['df_age'] = df_age
            st.session_state['df_remuneration'] = df_remuneration
            st.session_state['df_formation'] = df_formation
            st.session_state['df_calculs'] = df_calculs
            
            st.success("Fichier importé avec succès ! Vous pouvez maintenant consulter les résultats.")
            
        except Exception as e:
            st.error(f"Une erreur s'est produite lors de la lecture du fichier : {str(e)}")
            st.error("Veuillez vérifier que le fichier est bien au format attendu et qu'il contient toutes les feuilles requises.")

elif page == "Saisie manuelle":
    st.header("Saisie manuelle des données")
    
    # Informations générales
    st.subheader("Informations générales")
    company_name = st.text_input("Nom de l'entreprise")
    year = st.number_input("Année d'analyse", min_value=2000, max_value=2100, value=datetime.now().year)
    
    # Effectifs
    st.subheader("Effectifs")
    total_effectif = st.number_input("Effectif total", min_value=1, value=100)
    femmes_effectif = st.number_input("Nombre de femmes", min_value=0, max_value=total_effectif)
    femmes_cadres = st.number_input("Nombre de femmes cadres", min_value=0, max_value=femmes_effectif)
    personnes_handicap = st.number_input("Nombre de personnes en situation de handicap", min_value=0, max_value=total_effectif)
    
    # Rémunérations
    st.subheader("Rémunérations")
    salaire_moyen_hommes = st.number_input("Salaire moyen hommes (€)", min_value=0, value=30000)
    salaire_moyen_femmes = st.number_input("Salaire moyen femmes (€)", min_value=0, value=28000)
    
    # Contrats et formation
    st.subheader("Contrats et formation")
    nombre_cdi = st.number_input("Nombre de CDI", min_value=0, max_value=total_effectif)
    nombre_formation = st.number_input("Nombre d'employés formés", min_value=0, max_value=total_effectif)
    nombre_recrutement_interne = st.number_input("Nombre de recrutements internes", min_value=0)
    nombre_recrutement_total = st.number_input("Nombre total de recrutements", min_value=nombre_recrutement_interne)
    
    # Organisation du travail
    st.subheader("Organisation du travail")
    nombre_temps_partiel = st.number_input("Nombre d'employés à temps partiel", min_value=0, max_value=total_effectif)
    nombre_teletravail = st.number_input("Nombre d'employés en télétravail", min_value=0, max_value=total_effectif)
    nombre_promotion_femmes = st.number_input("Nombre de promotions femmes", min_value=0)
    nombre_promotion_total = st.number_input("Nombre total de promotions", min_value=nombre_promotion_femmes)
    
    if st.button("Calculer les indicateurs"):
        # Calcul des indicateurs
        data = {
            'taux_femmes': femmes_effectif / total_effectif,
            'taux_femmes_cadres': femmes_cadres / total_effectif if total_effectif > 0 else 0,
            'taux_handicap': personnes_handicap / total_effectif,
            'ecart_salaire': (salaire_moyen_hommes - salaire_moyen_femmes) / salaire_moyen_hommes if salaire_moyen_hommes > 0 else 0,
            'score_ages': 0.7,  # Valeur par défaut, à ajuster selon la répartition par âge
            'taux_cdi': nombre_cdi / total_effectif,
            'taux_formation': nombre_formation / total_effectif,
            'taux_recrutement_interne': nombre_recrutement_interne / nombre_recrutement_total if nombre_recrutement_total > 0 else 0,
            'taux_temps_partiel': nombre_temps_partiel / total_effectif,
            'taux_teletravail': nombre_teletravail / total_effectif,
            'taux_promotion_femmes': nombre_promotion_femmes / nombre_promotion_total if nombre_promotion_total > 0 else 0,
            'explications': {
                'taux_femmes': f"Le taux de féminisation est de {femmes_effectif} femmes sur {total_effectif} employés.",
                'taux_femmes_cadres': f"Le taux de femmes cadres est de {femmes_cadres} sur {total_effectif} employés.",
                'taux_handicap': f"Le taux de personnes en situation de handicap est de {personnes_handicap} sur {total_effectif} employés.",
                'ecart_salaire': f"L'écart salarial est de {(salaire_moyen_hommes - salaire_moyen_femmes) / salaire_moyen_hommes:.1%}.",
                'score_ages': "Score calculé sur la base de la répartition par âge.",
                'taux_cdi': f"Le taux de CDI est de {nombre_cdi} sur {total_effectif} employés.",
                'taux_formation': f"Le taux de formation est de {nombre_formation} sur {total_effectif} employés.",
                'taux_recrutement_interne': f"Le taux de recrutement interne est de {nombre_recrutement_interne} sur {nombre_recrutement_total} recrutements.",
                'taux_temps_partiel': f"Le taux de temps partiel est de {nombre_temps_partiel} sur {total_effectif} employés.",
                'taux_teletravail': f"Le taux de télétravail est de {nombre_teletravail} sur {total_effectif} employés.",
                'taux_promotion_femmes': f"Le taux de promotion des femmes est de {nombre_promotion_femmes} sur {nombre_promotion_total} promotions."
            }
        }
        
        # Stockage des données dans la session
        st.session_state['data'] = data
        st.session_state['company_name'] = company_name
        st.session_state['year'] = year
        
        st.success("Calculs effectués avec succès ! Vous pouvez maintenant consulter les résultats.")

elif page == "Définitions":
    st.header("Définitions des indicateurs")
    
    for code, indicateur in INDICATEURS.items():
        with st.expander(f"{indicateur['nom']} ({code})"):
            st.write(f"**Description :** {indicateur['description']}")
            st.write(f"**Objectif :** {indicateur['objectif']:.1%}")
            st.write(f"**Unité :** {indicateur['unite']}")
            st.write(f"**Formule :** {indicateur['formule']}")

elif page == "Résultats":
    if 'df_general' not in st.session_state:
        st.warning("Veuillez d'abord importer un fichier dans la section 'Import des données'.")
    else:
        st.header("3. Résultats de l'évaluation")
        
        # Extraction des données
        df_general = st.session_state['df_general']
        df_age = st.session_state['df_age']
        df_remuneration = st.session_state['df_remuneration']
        df_formation = st.session_state['df_formation']
        df_calculs = st.session_state['df_calculs']
        
        company_name = df_general.iloc[0, 1]
        year = df_general.iloc[1, 1]
        
        # Calcul des indicateurs
        data = {
            'taux_femmes': df_calculs.iloc[0, 2],
            'taux_femmes_cadres': df_calculs.iloc[1, 2],
            'taux_handicap': df_calculs.iloc[2, 2],
            'ecart_salaire': df_calculs.iloc[3, 2],
            'score_ages': df_calculs.iloc[4, 2],
            'taux_cdi': df_calculs.iloc[6, 2],
            'taux_formation': df_calculs.iloc[7, 2],
            'taux_recrutement_interne': df_calculs.iloc[8, 2],
            'taux_temps_partiel': df_calculs.iloc[9, 2],
            'taux_teletravail': df_calculs.iloc[11, 2],
            'taux_promotion_femmes': df_calculs.iloc[12, 2],
            'explications': {
                'taux_femmes': df_calculs.iloc[0, 6],
                'taux_femmes_cadres': df_calculs.iloc[1, 6],
                'taux_handicap': df_calculs.iloc[2, 6],
                'ecart_salaire': df_calculs.iloc[3, 6],
                'score_ages': df_calculs.iloc[4, 6],
                'taux_cdi': df_calculs.iloc[6, 6],
                'taux_formation': df_calculs.iloc[7, 6],
                'taux_recrutement_interne': df_calculs.iloc[8, 6],
                'taux_temps_partiel': df_calculs.iloc[9, 6],
                'taux_teletravail': df_calculs.iloc[11, 6],
                'taux_promotion_femmes': df_calculs.iloc[12, 6]
            }
        }
        
        # Détermination des points forts et axes d'amélioration
        data['points_forts'] = []
        data['axes_amelioration'] = []
        
        if data['taux_femmes'] >= 0.4:
            data['points_forts'].append("Bon taux de féminisation global")
        else:
            data['axes_amelioration'].append("Augmenter le taux de féminisation")
            
        if data['taux_femmes_cadres'] >= 0.4:
            data['points_forts'].append("Bonne représentation des femmes dans les postes cadres")
        else:
            data['axes_amelioration'].append("Améliorer la représentation des femmes dans les postes cadres")
            
        if data['taux_handicap'] >= 0.06:
            data['points_forts'].append("Seuil légal de 6% atteint pour l'emploi des personnes en situation de handicap")
        else:
            data['axes_amelioration'].append("Atteindre le seuil légal de 6% pour l'emploi des personnes en situation de handicap")
            
        if data['ecart_salaire'] <= 0.05:
            data['points_forts'].append("Faible écart salarial entre hommes et femmes")
        else:
            data['axes_amelioration'].append("Réduire l'écart salarial entre hommes et femmes")
            
        if data['score_ages'] >= 0.7:
            data['points_forts'].append("Bonne diversité des âges")
        else:
            data['axes_amelioration'].append("Améliorer la diversité des âges")
            
        if data['taux_cdi'] >= 0.8:
            data['points_forts'].append("Fort taux de CDI")
        else:
            data['axes_amelioration'].append("Augmenter le taux de CDI")
            
        if data['taux_formation'] >= 0.05:
            data['points_forts'].append("Bon taux de formation")
        else:
            data['axes_amelioration'].append("Développer la formation")
            
        if data['taux_recrutement_interne'] >= 0.3:
            data['points_forts'].append("Fort taux de recrutement interne")
        else:
            data['axes_amelioration'].append("Augmenter le taux de recrutement interne")
            
        if data['taux_temps_partiel'] <= 0.2:
            data['points_forts'].append("Taux de temps partiel maîtrisé")
        else:
            data['axes_amelioration'].append("Maîtriser le taux de temps partiel")
            
        if data['taux_teletravail'] >= 0.2:
            data['points_forts'].append("Bon taux de télétravail")
        else:
            data['axes_amelioration'].append("Développer le télétravail")
            
        if data['taux_promotion_femmes'] >= 0.4:
            data['points_forts'].append("Bon taux de promotion des femmes")
        else:
            data['axes_amelioration'].append("Augmenter le taux de promotion des femmes")
        
        # Stockage des données dans la session
        st.session_state['data'] = data
        st.session_state['company_name'] = company_name
        st.session_state['year'] = year
        
        # Indicateurs clés
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Taux de féminisation", f"{data['taux_femmes']:.1%}", f"{data['taux_femmes']*100-40:+.1f}%")
            st.metric("Taux de femmes cadres", f"{data['taux_femmes_cadres']:.1%}", f"{data['taux_femmes_cadres']*100-40:+.1f}%")
            st.metric("Taux de handicap", f"{data['taux_handicap']:.1%}", f"{data['taux_handicap']*100-6:+.1f}%")
            
        with col2:
            st.metric("Écart salarial", f"{data['ecart_salaire']:.1%}", f"{data['ecart_salaire']*100-5:+.1f}%")
            st.metric("Score âges", f"{data['score_ages']:.1%}", f"{data['score_ages']*100-70:+.1f}%")
            st.metric("Taux de CDI", f"{data['taux_cdi']:.1%}", f"{data['taux_cdi']*100-80:+.1f}%")
            
        with col3:
            st.metric("Taux de formation", f"{data['taux_formation']:.1%}", f"{data['taux_formation']*100-5:+.1f}%")
            st.metric("Taux de recrutement interne", f"{data['taux_recrutement_interne']:.1%}", f"{data['taux_recrutement_interne']*100-30:+.1f}%")
            st.metric("Taux de télétravail", f"{data['taux_teletravail']:.1%}", f"{data['taux_teletravail']*100-20:+.1f}%")
        
        # Graphiques
        st.subheader("Visualisations")
        col1, col2 = st.columns(2)
        
        with col1:
            st.plotly_chart(create_age_distribution_chart(df_age), use_container_width=True)
            
        with col2:
            st.plotly_chart(create_salary_gap_chart(df_remuneration), use_container_width=True)
            
        st.plotly_chart(create_indicators_radar(data), use_container_width=True)
        
        # Points forts et axes d'amélioration
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Points Forts")
            for point in data['points_forts']:
                st.write(f"✅ {point}")
                
        with col2:
            st.subheader("Axes d'Amélioration")
            for point in data['axes_amelioration']:
                st.write(f"⚠️ {point}")

elif page == "Rapport":
    if 'data' not in st.session_state:
        st.warning("Veuillez d'abord importer un fichier et consulter les résultats.")
    else:
        st.header("4. Génération du rapport")
        
        if st.button("📄 Générer le rapport PDF"):
            pdf_path = generate_pdf(
                st.session_state['data'],
                st.session_state['company_name'],
                st.session_state['year']
            )
            
            with open(pdf_path, "rb") as file:
                st.download_button(
                    label="📥 Télécharger le rapport PDF",
                    data=file,
                    file_name=f"rapport_diversite_inclusion_{st.session_state['company_name']}_{st.session_state['year']}.pdf",
                    mime="application/pdf"
                )
            os.unlink(pdf_path) 