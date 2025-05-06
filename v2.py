import streamlit as st
import pandas as pd
import numpy as np
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import tempfile
import os
from datetime import datetime

# Configuration de la page
st.set_page_config(
    page_title="Évaluation Diversité et Inclusion V2",
    page_icon="📊",
    layout="wide"
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

# Fonction pour générer le PDF
def generate_pdf(data, company_name, year):
    # Créer un fichier temporaire
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    
    # Configuration du document
    doc = SimpleDocTemplate(
        temp_file.name,
        pagesize=landscape(A4),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )
    
    # Liste des éléments du document
    elements = []
    
    # En-tête
    elements.append(Paragraph(f"Rapport d'Évaluation Diversité et Inclusion", title_style))
    elements.append(Paragraph(f"{company_name} - {year}", heading_style))
    elements.append(Paragraph(f"Date de génération : {datetime.now().strftime('%d/%m/%Y')}", styles['Normal']))
    elements.append(Spacer(1, 20))
    
    # Résumé exécutif
    elements.append(Paragraph("Résumé Exécutif", heading_style))
    elements.append(Paragraph(
        f"Ce rapport présente une analyse détaillée de la diversité et de l'inclusion dans votre entreprise. "
        f"Les résultats montrent des points forts et des axes d'amélioration qui sont détaillés dans les sections suivantes.",
        styles['Normal']
    ))
    elements.append(Spacer(1, 20))
    
    # Score global
    elements.append(Paragraph("Score Global", heading_style))
    score_data = [
        ['Indicateur', 'Valeur', 'Objectif', 'Statut', 'Explication'],
        ['Taux de féminisation', f"{data['feminization_rate']:.1f}%", '40%', '✅' if data['feminization_rate'] >= 40 else '⚠️',
         'Pourcentage de femmes dans l\'effectif total'],
        ['Taux de femmes cadres', f"{data['female_executives_rate']:.1f}%", '40%', '✅' if data['female_executives_rate'] >= 40 else '⚠️',
         'Pourcentage de femmes parmi les postes cadres'],
        ['Taux d\'emploi des personnes en situation de handicap', f"{data['disability_rate']:.1f}%", '6%', '✅' if data['disability_rate'] >= 6 else '⚠️',
         'Pourcentage de salariés en situation de handicap'],
        ['Écart de rémunération', f"{data['salary_gap']:.1f}%", '<5%', '✅' if data['salary_gap'] <= 5 else '⚠️',
         'Écart moyen de rémunération entre hommes et femmes'],
        ['Score d\'équilibre des âges', f"{data['age_balance']:.1f}%", '>70%', '✅' if data['age_balance'] >= 70 else '⚠️',
         'Mesure de la diversité des âges'],
        ['Taux d\'absentéisme', f"{data['absenteeism_rate']:.1f}%", '<4%', '✅' if data['absenteeism_rate'] <= 4 else '⚠️',
         'Taux d\'absence par rapport aux jours travaillés'],
        ['Taux de CDI', f"{data['cdi_rate']:.1f}%", '>80%', '✅' if data['cdi_rate'] >= 80 else '⚠️',
         'Pourcentage de contrats CDI dans l\'effectif'],
        ['Taux de formation', f"{data['training_rate']:.1f}%", '>5%', '✅' if data['training_rate'] >= 5 else '⚠️',
         'Taux de participation aux formations'],
        ['Taux de recrutement interne', f"{data['internal_recruitment_rate']:.1f}%", '>30%', '✅' if data['internal_recruitment_rate'] >= 30 else '⚠️',
         'Pourcentage de promotions internes']
    ]
    
    score_table = Table(score_data, colWidths=[2*inch, 1.5*inch, 1.5*inch, 1*inch, 3*inch])
    score_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (4, 0), (4, -1), 'LEFT'),
    ]))
    elements.append(score_table)
    elements.append(Spacer(1, 20))
    
    # Points forts
    elements.append(Paragraph("Points Forts", heading_style))
    strengths = []
    if data['feminization_rate'] >= 40:
        strengths.append("Un bon taux de féminisation global")
    if data['female_executives_rate'] >= 40:
        strengths.append("Une bonne représentation des femmes dans les postes cadres")
    if data['disability_rate'] >= 6:
        strengths.append("Un taux d'emploi des personnes en situation de handicap conforme à la loi")
    if data['salary_gap'] <= 5:
        strengths.append("Des écarts de rémunération maîtrisés")
    if data['age_balance'] >= 70:
        strengths.append("Une bonne répartition des âges")
    if data['absenteeism_rate'] <= 4:
        strengths.append("Un taux d'absentéisme maîtrisé")
    if data['cdi_rate'] >= 80:
        strengths.append("Une forte proportion de contrats CDI")
    if data['training_rate'] >= 5:
        strengths.append("Un bon taux de participation aux formations")
    if data['internal_recruitment_rate'] >= 30:
        strengths.append("Une bonne politique de promotion interne")
    
    if strengths:
        for strength in strengths:
            elements.append(Paragraph(f"• {strength}", styles['Normal']))
    else:
        elements.append(Paragraph("Aucun point fort significatif identifié", styles['Normal']))
    elements.append(Spacer(1, 20))
    
    # Axes d'amélioration
    elements.append(Paragraph("Axes d'Amélioration", heading_style))
    improvements = []
    if data['feminization_rate'] < 40:
        improvements.append("Augmenter le taux de féminisation global")
    if data['female_executives_rate'] < 40:
        improvements.append("Développer la présence des femmes dans les postes cadres")
    if data['disability_rate'] < 6:
        improvements.append("Atteindre le taux légal d'emploi des personnes en situation de handicap")
    if data['salary_gap'] > 5:
        improvements.append("Réduire les écarts de rémunération entre les hommes et les femmes")
    if data['age_balance'] < 70:
        improvements.append("Améliorer l'équilibre des âges dans l'entreprise")
    if data['absenteeism_rate'] > 4:
        improvements.append("Mettre en place des actions pour réduire l'absentéisme")
    if data['cdi_rate'] < 80:
        improvements.append("Augmenter la proportion de contrats CDI")
    if data['training_rate'] < 5:
        improvements.append("Développer la participation aux formations")
    if data['internal_recruitment_rate'] < 30:
        improvements.append("Renforcer la politique de promotion interne")
    
    for improvement in improvements:
        elements.append(Paragraph(f"• {improvement}", styles['Normal']))
    elements.append(Spacer(1, 20))
    
    # Plan d'action
    elements.append(Paragraph("Plan d'Action Recommandé", heading_style))
    action_plan = [
        ['Priorité', 'Action', 'Objectif', 'Délai'],
        ['Haute', 'Mise en place d\'un plan de recrutement ciblé', 'Augmenter la diversité', '6 mois'],
        ['Haute', 'Formation des managers à la diversité', 'Sensibilisation', '3 mois'],
        ['Moyenne', 'Audit des rémunérations', 'Réduire les écarts', '12 mois'],
        ['Moyenne', 'Programme de mentorat', 'Développement des talents', 'Ongoing'],
        ['Basse', 'Étude de l\'absentéisme', 'Comprendre les causes', '3 mois'],
        ['Moyenne', 'Développement de la formation continue', 'Augmenter la participation', '6 mois'],
        ['Moyenne', 'Plan de développement des carrières', 'Favoriser les promotions internes', '12 mois']
    ]
    
    action_table = Table(action_plan, colWidths=[1.5*inch, 3*inch, 2*inch, 1.5*inch])
    action_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4472C4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(action_table)
    elements.append(Spacer(1, 20))
    
    # Conclusion
    elements.append(Paragraph("Conclusion", heading_style))
    conclusion = (
        f"L'analyse de la diversité et de l'inclusion dans {company_name} révèle des opportunités "
        "d'amélioration significatives. En mettant en œuvre les actions recommandées, "
        "l'entreprise pourra renforcer sa performance sociale et sa compétitivité. "
        "Un suivi régulier des indicateurs permettra de mesurer les progrès accomplis."
    )
    elements.append(Paragraph(conclusion, styles['Normal']))
    
    # Pied de page
    elements.append(Spacer(1, 30))
    elements.append(Paragraph(
        "Ce rapport a été généré automatiquement par l'application d'évaluation de la Diversité et Inclusion",
        ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=8,
            textColor=colors.gray
        )
    ))
    
    # Génération du PDF
    doc.build(elements)
    return temp_file.name

# Interface Streamlit
st.title("📊 Évaluation de la Diversité et Inclusion V2")

# Section pour télécharger le modèle
st.header("1. Télécharger le modèle")
st.write("Commencez par télécharger notre modèle Excel complet pour saisir vos données.")
if st.button("📥 Télécharger le modèle Excel"):
    with open('modele_bilan_social_v2.xlsx', 'rb') as f:
        st.download_button(
            label="Cliquez pour télécharger",
            data=f,
            file_name="modele_bilan_social_v2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# Section pour importer les données
st.header("2. Importer vos données")
st.write("Une fois le modèle complété, importez-le ici pour générer votre rapport.")
uploaded_file = st.file_uploader("Choisissez votre fichier Excel", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Lecture des données
        df_general = pd.read_excel(uploaded_file, sheet_name='Données générales')
        df_age = pd.read_excel(uploaded_file, sheet_name='Répartition par âge')
        df_remuneration = pd.read_excel(uploaded_file, sheet_name='Rémunérations')
        df_formation = pd.read_excel(uploaded_file, sheet_name='Formation et Recrutement')
        
        # Extraction des données
        company_name = df_general.iloc[0, 1]
        year = df_general.iloc[1, 1]
        total_employees = df_general.iloc[2, 1]
        women_count = df_general.iloc[3, 1]
        men_count = df_general.iloc[4, 1]
        executives_count = df_general.iloc[5, 1]
        women_executives = df_general.iloc[6, 1]
        disabled_employees = df_general.iloc[7, 1]
        working_days = df_general.iloc[8, 1]
        absence_days = df_general.iloc[9, 1]
        cdi_count = df_general.iloc[10, 1]
        
        # Calcul des indicateurs
        data = {
            'feminization_rate': (women_count / total_employees) * 100,
            'female_executives_rate': (women_executives / executives_count) * 100,
            'disability_rate': (disabled_employees / total_employees) * 100,
            'salary_gap': df_remuneration['Écart (%)'].mean(),
            'age_balance': (df_age['Total'].std() / df_age['Total'].mean()) * 100,
            'absenteeism_rate': (absence_days / working_days) * 100,
            'cdi_rate': (cdi_count / total_employees) * 100,
            'training_rate': (df_formation.iloc[0, 1] / total_employees) * 100,
            'internal_recruitment_rate': (df_formation.iloc[4, 1] / df_formation.iloc[3, 1]) * 100
        }
        
        # Affichage des résultats
        st.header("3. Résultats de l'évaluation")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Indicateurs clés")
            st.metric("Taux de féminisation", f"{data['feminization_rate']:.1f}%")
            st.metric("Taux de femmes cadres", f"{data['female_executives_rate']:.1f}%")
            st.metric("Taux d'emploi des personnes en situation de handicap", f"{data['disability_rate']:.1f}%")
            st.metric("Écart de rémunération", f"{data['salary_gap']:.1f}%")
        
        with col2:
            st.subheader("Autres indicateurs")
            st.metric("Score d'équilibre des âges", f"{data['age_balance']:.1f}%")
            st.metric("Taux d'absentéisme", f"{data['absenteeism_rate']:.1f}%")
            st.metric("Taux de CDI", f"{data['cdi_rate']:.1f}%")
            st.metric("Taux de formation", f"{data['training_rate']:.1f}%")
            st.metric("Taux de recrutement interne", f"{data['internal_recruitment_rate']:.1f}%")
        
        # Génération du PDF
        if st.button("📄 Générer le rapport PDF"):
            pdf_path = generate_pdf(data, company_name, year)
            with open(pdf_path, 'rb') as f:
                st.download_button(
                    label="📥 Télécharger le rapport PDF",
                    data=f,
                    file_name=f"rapport_diversite_{company_name}_{year}.pdf",
                    mime="application/pdf"
                )
            os.unlink(pdf_path)  # Suppression du fichier temporaire
    
    except Exception as e:
        st.error(f"Une erreur est survenue lors du traitement du fichier : {str(e)}")
        st.write("Veuillez vérifier que le fichier correspond bien au modèle fourni.")
