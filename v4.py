import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import altair as alt
import io
import plotly.graph_objects as go
import plotly.express as px
from io import StringIO
import base64
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import tempfile
import os
from datetime import datetime

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Évaluateur D&I",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Titre et introduction de l'application
st.title("📊 Évaluateur de Diversité et Inclusion en Entreprise")
st.markdown("""
Cette application analyse les indicateurs sociaux d'une entreprise en matière de diversité et inclusion,
et attribue des notes de A à E sur 6 dimensions clés, basées sur des seuils adaptés au secteur énergie/industrie.
""")

# Fonction pour attribuer une note (A-E) selon les seuils définis
def attribuer_note(valeur, seuils, ordre_croissant=True):
    if ordre_croissant:
        if valeur >= seuils[0]:
            return "A"
        elif valeur >= seuils[1]:
            return "B"
        elif valeur >= seuils[2]:
            return "C"
        elif valeur >= seuils[3]:
            return "D"
        else:
            return "E"
    else:  # Ordre décroissant (plus petit = meilleur)
        if valeur <= seuils[0]:
            return "A"
        elif valeur <= seuils[1]:
            return "B"
        elif valeur <= seuils[2]:
            return "C"
        elif valeur <= seuils[3]:
            return "D"
        else:
            return "E"

# Fonction pour convertir une note en chiffre
def note_vers_chiffre(note):
    return {"A": 5, "B": 4, "C": 3, "D": 2, "E": 1}.get(note, 0)

# Fonction pour convertir un chiffre en note
def chiffre_vers_note(chiffre):
    if chiffre >= 4.5:
        return "A"
    elif chiffre >= 3.5:
        return "B"
    elif chiffre >= 2.5:
        return "C"
    elif chiffre >= 1.5:
        return "D"
    else:
        return "E"

# Fonction pour calculer l'équilibre des âges
def calculer_equilibre_age(moins_30, entre_30_50, plus_50):
    # Calcul de l'écart type des pourcentages
    valeurs = [moins_30, entre_30_50, plus_50]
    moyenne = sum(valeurs) / len(valeurs)
    ecart_type = np.sqrt(sum((x - moyenne) ** 2 for x in valeurs) / len(valeurs))
    # Score sur 100 (100 = parfait équilibre)
    return max(0, 100 - (ecart_type * 2))

# Définition des seuils pour chaque indicateur
seuils = {
    "taux_feminisation": [45, 40, 30, 20],  # %
    "taux_femmes_cadres": [40, 30, 20, 15],  # %
    "taux_handicap": [6, 5, 4, 3],  # % (le seuil légal est de 6%)
    "ecart_salaire": [3, 5, 10, 15],  # % (ordre décroissant: plus petit = meilleur)
    "equilibre_age": [80, 70, 60, 50],  # % (mesure d'équilibre)
    "taux_absenteisme": [3, 4, 5, 6]  # % (ordre décroissant: plus petit = meilleur)
}

# Définition des règles de notation
regles_notation = {
    "A": {
        "description": "Performance exemplaire",
        "seuils": {
            "taux_feminisation": "≥ 45%",
            "taux_femmes_cadres": "≥ 40%",
            "taux_handicap": "≥ 6%",
            "ecart_salaire": "≤ 3%",
            "equilibre_age": "≥ 80%",
            "taux_absenteisme": "≤ 3%"
        }
    },
    "B": {
        "description": "Performance satisfaisante",
        "seuils": {
            "taux_feminisation": "≥ 40%",
            "taux_femmes_cadres": "≥ 30%",
            "taux_handicap": "≥ 5%",
            "ecart_salaire": "≤ 5%",
            "equilibre_age": "≥ 70%",
            "taux_absenteisme": "≤ 4%"
        }
    },
    "C": {
        "description": "Performance moyenne",
        "seuils": {
            "taux_feminisation": "≥ 30%",
            "taux_femmes_cadres": "≥ 20%",
            "taux_handicap": "≥ 4%",
            "ecart_salaire": "≤ 10%",
            "equilibre_age": "≥ 60%",
            "taux_absenteisme": "≤ 5%"
        }
    },
    "D": {
        "description": "Performance insuffisante",
        "seuils": {
            "taux_feminisation": "≥ 20%",
            "taux_femmes_cadres": "≥ 15%",
            "taux_handicap": "≥ 3%",
            "ecart_salaire": "≤ 15%",
            "equilibre_age": "≥ 50%",
            "taux_absenteisme": "≤ 6%"
        }
    },
    "E": {
        "description": "Performance critique",
        "seuils": {
            "taux_feminisation": "< 20%",
            "taux_femmes_cadres": "< 15%",
            "taux_handicap": "< 3%",
            "ecart_salaire": "> 15%",
            "equilibre_age": "< 50%",
            "taux_absenteisme": "> 6%"
        }
    }
}

# Styles PDF
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

def generate_pdf(data, company_name, year):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        doc = SimpleDocTemplate(tmp.name, pagesize=letter)
        elements = []
        
        # Titre
        elements.append(Paragraph(f"Rapport d'Évaluation Diversité et Inclusion<br/>{company_name} - {year}", title_style))
        elements.append(Spacer(1, 20))
        
        # Note globale
        elements.append(Paragraph("Note Globale", heading_style))
        elements.append(Paragraph(f"Note : {data['note_globale']} (Score : {data['score_global']:.2f}/5)", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Résumé exécutif
        elements.append(Paragraph("Résumé Exécutif", heading_style))
        elements.append(Paragraph(
            f"Ce rapport présente une analyse détaillée de la diversité et de l'inclusion au sein de {company_name} "
            f"pour l'année {year}. L'évaluation couvre plusieurs dimensions clés : la parité, l'inclusion des personnes "
            "en situation de handicap, l'équité salariale, la diversité des âges, et les pratiques de formation et de "
            "recrutement.",
            styles['Normal']
        ))
        elements.append(Spacer(1, 20))
        
        # Notes et indicateurs
        elements.append(Paragraph("Notes par indicateur", heading_style))
        table_data = [["Indicateur", "Note", "Valeur", "Seuils"]]
        for k, v in data['resultats'].items():
            table_data.append([
                k,
                v['note'],
                f"{v['valeur']:.1f}%",
                f"A: {seuils[k.lower().replace(' ', '_')][0]}%, B: {seuils[k.lower().replace(' ', '_')][1]}%, C: {seuils[k.lower().replace(' ', '_')][2]}%, D: {seuils[k.lower().replace(' ', '_')][3]}%"
            ])
        
        table = Table(table_data, colWidths=[2*inch, 1*inch, 1*inch, 2*inch])
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
        if data['points_forts']:
            for point in data['points_forts']:
                elements.append(Paragraph(f"• {point}", styles['Normal']))
        else:
            elements.append(Paragraph("Aucun point fort identifié", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Axes d'amélioration
        elements.append(Paragraph("Axes d'Amélioration", heading_style))
        if data['axes_amelioration']:
            for point in data['axes_amelioration']:
                elements.append(Paragraph(f"• {point}", styles['Normal']))
        else:
            elements.append(Paragraph("Aucun axe d'amélioration critique identifié", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Recommandations
        elements.append(Paragraph("Recommandations", heading_style))
        for indicateur, details in data['resultats'].items():
            if details["note"] in ["D", "E"]:
                elements.append(Paragraph(f"Pour {indicateur} :", styles['Normal']))
                for rec in get_recommandations(indicateur, details['valeur'], details['note']).split('\n'):
                    if rec.strip():
                        elements.append(Paragraph(f"• {rec.strip()}", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Conclusion
        elements.append(Paragraph("Conclusion", heading_style))
        elements.append(Paragraph(
            f"L'analyse des données de {company_name} révèle des opportunités significatives d'amélioration en matière "
            "de diversité et d'inclusion. Les actions recommandées, si elles sont mises en œuvre de manière cohérente "
            "et suivie, permettront de progresser vers une organisation plus inclusive et équitable.",
            styles['Normal']
        ))
        
        # Date de génération
        elements.append(Spacer(1, 20))
        elements.append(Paragraph(
            f"Rapport généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
            styles['Normal']
        ))
        
        doc.build(elements)
        return tmp.name

# Interface Streamlit
st.markdown("## 📌 Explication des indicateurs")
st.write("""
1. **Taux de féminisation global** : Pourcentage de femmes dans l'effectif total
2. **Taux de femmes cadres** : Pourcentage de femmes parmi les postes d'encadrement
3. **Taux d'emploi des personnes en situation de handicap** : Pourcentage de salariés en situation de handicap (seuil légal = 6%)
4. **Écart de salaire hommes/femmes** : Écart moyen en % à poste équivalent (0% = parfaite égalité)
5. **Répartition des effectifs par âge** : Équilibre entre les tranches d'âge (<30 ans, 30-50 ans, >50 ans)
6. **Taux d'absentéisme** : Pourcentage de jours d'absence par rapport au nombre total de jours travaillés
""")

# Méthode d'entrée des données
st.markdown("## 📝 Entrée des données")
methode = st.radio("Choisissez la méthode d'entrée des données:", 
                  ["Saisie manuelle", "Téléchargement de fichier CSV"])

# Variables pour stocker les données saisies
indicateurs = {}

if methode == "Saisie manuelle":
    # Créer deux colonnes pour l'entrée des données
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Données générales")
        nom_entreprise = st.text_input("Nom de l'entreprise", "EDF SA")
        annee = st.number_input("Année", min_value=2000, max_value=2030, value=2022)
        
        st.subheader("Mixité et égalité professionnelle")
        indicateurs["taux_feminisation"] = st.number_input("Taux de féminisation global (%)", 
                                                         min_value=0.0, max_value=100.0, value=30.0)
        indicateurs["taux_femmes_cadres"] = st.number_input("Taux de femmes cadres (%)", 
                                                          min_value=0.0, max_value=100.0, value=28.0)
        indicateurs["ecart_salaire"] = st.number_input("Écart de salaire hommes/femmes (%)", 
                                                      min_value=0.0, max_value=50.0, value=5.0)
    
    with col2:
        st.subheader("Inclusion et diversité")
        indicateurs["taux_handicap"] = st.number_input("Taux d'emploi des personnes en situation de handicap (%)", 
                                                    min_value=0.0, max_value=20.0, value=5.5)
        
        st.subheader("Répartition par âge")
        moins_30 = st.number_input("Effectifs < 30 ans (%)", 
                                  min_value=0.0, max_value=100.0, value=15.0)
        entre_30_50 = st.number_input("Effectifs 30-50 ans (%)", 
                                     min_value=0.0, max_value=100.0, value=45.0)
        plus_50 = st.number_input("Effectifs > 50 ans (%)", 
                                min_value=0.0, max_value=100.0, value=40.0)
        
        # Vérifier que la somme des pourcentages est égale à 100%
        sum_age = moins_30 + entre_30_50 + plus_50
        if abs(sum_age - 100) > 0.01:  # Tolérance de 0.01%
            st.warning(f"La somme des pourcentages par âge doit être égale à 100% (actuellement {sum_age:.2f}%)")
        
        # Calculer l'équilibre des âges
        indicateurs["equilibre_age"] = calculer_equilibre_age(moins_30, entre_30_50, plus_50)
        
        st.subheader("Climat social")
        indicateurs["taux_absenteisme"] = st.number_input("Taux d'absentéisme (%)", 
                                                        min_value=0.0, max_value=20.0, value=4.2)

else:  # Import CSV
    uploaded_file = st.file_uploader("Importer un fichier CSV", type=["csv"])
    
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            st.success("Fichier importé avec succès !")
            st.dataframe(df)
            
            # Extraction des données
            nom_entreprise = df.get('Entreprise', ["Entreprise"])[0]
            annee = int(df.get('Année', [2022])[0])
            
            # Remplir les indicateurs
            indicateurs["taux_feminisation"] = float(df.get('Taux féminisation', [0])[0])
            indicateurs["taux_femmes_cadres"] = float(df.get('Taux femmes cadres', [0])[0])
            indicateurs["ecart_salaire"] = float(df.get('Écart salaire', [0])[0])
            indicateurs["taux_handicap"] = float(df.get('Taux handicap', [0])[0])
            
            # Calculer l'équilibre des âges
            moins_30 = float(df.get('Moins 30 ans', [0])[0])
            entre_30_50 = float(df.get('30-50 ans', [0])[0])
            plus_50 = float(df.get('Plus 50 ans', [0])[0])
            
            indicateurs["equilibre_age"] = calculer_equilibre_age(moins_30, entre_30_50, plus_50)
            indicateurs["taux_absenteisme"] = float(df.get('Taux absentéisme', [0])[0])
            
        except Exception as e:
            st.error(f"Erreur de lecture du CSV : {e}")
            st.stop()

# Calcul et affichage des résultats
if st.button("Calculer les résultats"):
    st.markdown("## 📊 Résultats de l'évaluation")
    
    # Calculer les notes pour chaque indicateur
    resultats = {
        "Taux de féminisation global": {
            "note": attribuer_note(indicateurs["taux_feminisation"], seuils["taux_feminisation"]),
            "valeur": indicateurs["taux_feminisation"]
        },
        "Taux de femmes cadres": {
            "note": attribuer_note(indicateurs["taux_femmes_cadres"], seuils["taux_femmes_cadres"]),
            "valeur": indicateurs["taux_femmes_cadres"]
        },
        "Taux d'emploi handicap": {
            "note": attribuer_note(indicateurs["taux_handicap"], seuils["taux_handicap"]),
            "valeur": indicateurs["taux_handicap"]
        },
        "Écart de salaire H/F": {
            "note": attribuer_note(indicateurs["ecart_salaire"], seuils["ecart_salaire"], False),
            "valeur": indicateurs["ecart_salaire"]
        },
        "Équilibre des âges": {
            "note": attribuer_note(indicateurs["equilibre_age"], seuils["equilibre_age"]),
            "valeur": indicateurs["equilibre_age"]
        },
        "Taux d'absentéisme": {
            "note": attribuer_note(indicateurs["taux_absenteisme"], seuils["taux_absenteisme"], False),
            "valeur": indicateurs["taux_absenteisme"]
        }
    }
    
    # Convertir les notes en valeurs numériques
    notes_numeriques = {cle: note_vers_chiffre(v["note"]) for cle, v in resultats.items()}
    
    # Calculer le score global
    score_global = sum(notes_numeriques.values()) / len(notes_numeriques)
    note_globale = chiffre_vers_note(score_global)
    
    # Préparation des données pour l'affichage
    df_resultats = pd.DataFrame({
        "Indicateur": list(resultats.keys()),
        "Note": [v["note"] for v in resultats.values()],
        "Score": list(notes_numeriques.values()),
        "Valeur": [f"{v['valeur']:.1f}%" for v in resultats.values()]
    })
    
    # Définir des couleurs pour chaque note
    couleurs_notes = {
        "A": "#4CAF50",  # Vert
        "B": "#8BC34A",  # Vert clair
        "C": "#FFC107",  # Jaune
        "D": "#FF9800",  # Orange
        "E": "#F44336"   # Rouge
    }
    
    # Affichage des résultats
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Note globale")
        st.markdown(f"<h1 style='text-align: center; color: {couleurs_notes[note_globale]}'>{note_globale}</h1>", unsafe_allow_html=True)
        st.markdown(f"<p style='text-align: center'>Score : {score_global:.2f}/5</p>", unsafe_allow_html=True)
        
        # Graphique radar des notes
        fig = go.Figure()
        fig.add_trace(go.Scatterpolar(
            r=[v["valeur"] for v in resultats.values()],
            theta=list(resultats.keys()),
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
            title='Indicateurs clés'
        )
        st.plotly_chart(fig)
    
    with col2:
        st.subheader("Détail des notes")
        for indicateur, details in resultats.items():
            st.markdown(f"""
            <div style='background-color: {couleurs_notes[details['note']]}; padding: 10px; margin: 5px 0; border-radius: 5px'>
                <strong>{indicateur}</strong><br>
                Note: {details['note']} | Valeur: {details['valeur']:.1f}%
            </div>
            """, unsafe_allow_html=True)
    
    # Points forts et axes d'amélioration
    points_forts = []
    axes_amelioration = []
    
    for indicateur, details in resultats.items():
        if details["note"] in ["A", "B"]:
            points_forts.append(f"{indicateur} : {details['valeur']:.1f}% (Note {details['note']})")
        elif details["note"] in ["D", "E"]:
            axes_amelioration.append(f"{indicateur} : {details['valeur']:.1f}% (Note {details['note']})")
    
    # Génération du PDF
    st.markdown("## 📄 Génération du rapport")
    
    # Bouton pour générer le PDF
    if st.button("Générer le rapport PDF"):
        try:
            # Préparation des données pour le PDF
            data = {
                "resultats": resultats,
                "points_forts": points_forts,
                "axes_amelioration": axes_amelioration,
                "note_globale": note_globale,
                "score_global": score_global
            }
            
            # Génération du PDF
            pdf_path = generate_pdf(data, nom_entreprise, annee)
            
            # Lecture et téléchargement du PDF
            with open(pdf_path, "rb") as file:
                pdf_data = file.read()
                st.download_button(
                    label="📥 Télécharger le rapport PDF",
                    data=pdf_data,
                    file_name=f"rapport_diversite_inclusion_{nom_entreprise}_{annee}.pdf",
                    mime="application/pdf"
                )
            
            # Nettoyage du fichier temporaire
            os.unlink(pdf_path)
            
            st.success("Le rapport PDF a été généré avec succès !")
            
        except Exception as e:
            st.error(f"Erreur lors de la génération du PDF : {str(e)}")
            st.error("Veuillez vérifier que toutes les données sont correctement saisies.")

# Fonction pour générer les recommandations
def get_recommandations(indicateur, valeur, note):
    recommandations = {
        "Taux de féminisation global": [
            "Mettre en place un plan de recrutement ciblé",
            "Développer des programmes de mentorat pour les femmes",
            "Renforcer la mixité dans les processus de recrutement"
        ],
        "Taux de femmes cadres": [
            "Créer un programme de développement de carrière spécifique",
            "Mettre en place un système de parrainage",
            "Former les managers à la détection des potentiels"
        ],
        "Taux d'emploi handicap": [
            "Renforcer les partenariats avec les structures d'insertion",
            "Adapter les postes de travail",
            "Former les équipes à l'accueil des personnes en situation de handicap"
        ],
        "Écart de salaire H/F": [
            "Réaliser un audit complet des rémunérations",
            "Mettre en place une grille salariale transparente",
            "Former les managers à la négociation équitable"
        ],
        "Équilibre des âges": [
            "Développer des programmes de transfert de compétences",
            "Mettre en place un plan de succession",
            "Adapter les conditions de travail aux différentes générations"
        ],
        "Taux d'absentéisme": [
            "Analyser les causes de l'absentéisme",
            "Mettre en place un suivi personnalisé",
            "Renforcer la prévention et le bien-être au travail"
        ]
    }
    
    return "\n".join([f"<li>{rec}</li>" for rec in recommandations.get(indicateur, [])])

# Pied de page
st.markdown("---")
st.markdown("""
**Diversité & Inclusion Analytics** | Développé avec Streamlit | Basé sur les référentiels du secteur énergie/industrie
""")

# Ajout d'informations sur la méthodologie dans la barre latérale
with st.sidebar:
    st.header("À propos")
    st.markdown("""
    Cet outil d'évaluation est basé sur les pratiques et standards du secteur énergie/industrie.
    Les seuils utilisés pour l'évaluation sont définis à partir de benchmarks sectoriels
    et des exigences légales en vigueur.
    
    **Sources**:
    - Bilans sociaux des entreprises du CAC40
    - Rapports de performance extra-financière
    - Études sectorielles
    - Cadre légal (notamment la loi Copé-Zimmermann, loi Avenir professionnel)
    """)
    
    # Affichage des références
    st.header("Références")
    st.markdown("""
    - Bruna, M. G. (2011). Diversité dans l'entreprise : d'impératif éthique à levier de créativité.
    - Arreola, F., & Sachet Milliat, A. (2022). Question(s) de diversité et inclusion dans l'emploi : nouvelles perspectives.
    - Thomas, D. A., & Ely, R. J. (1996). Making differences matter: A new paradigm for managing diversity.
    """) 