# -*- coding: utf-8 -*-
"""
Application de Gestion des Heures Supplémentaires des Enseignants Vacataires
Département d'Électrotechnique - Faculté de Génie Électrique - UDL-SBA
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import io
import re
from pathlib import Path

# --- Configuration de la Page Streamlit ---
st.set_page_config(
    page_title="Gestion H.SUP Vacataires - ELT UDL-SBA",
    page_icon="⚡",
    layout="wide"
)

TITRE_OFFICIEL = "Plateforme de Gestion des Heures Supplémentaires des Enseignants Vacataires (S1)"
st.title(f"⚡ {TITRE_OFFICIEL}")
st.markdown("**Département d'Électrotechnique - Faculté de génie électrique - UDL-SBA**")
st.divider()

BASE_DIR = Path(__file__).resolve().parent
FICHIER_PV = BASE_DIR / "Permanents-Vacataires-ELT2-2026-2027.xlsx"
FICHIER_EDT = BASE_DIR / "dataEDT-ELT-S1-2027.xlsx"
FICHIER_TEMPLATE = BASE_DIR / "Etats H.SUP VAC 2025-2026.xlsx"

# --- Fonctions de Nettoyage et Normalisation ---
def nettoyer_texte(valeur):
    if valeur is None or pd.isna(valeur):
        return ""
    txt = str(valeur).strip()
    return re.sub(r'\s+', ' ', txt)

def normaliser_nom(nom):
    txt = nettoyer_texte(nom).upper()
    # Suppression des accents et caractères spéciaux
    txt = re.sub(r'[ÉÈÊË]', 'E', txt)
    txt = re.sub(r'[ÀÂÄ]', 'A', txt)
    txt = re.sub(r'[ÎÏ]', 'I', txt)
    txt = re.sub(r'[ÔÖ]', 'O', txt)
    txt = re.sub(r'[ÙÛÜ]', 'U', txt)
    txt = re.sub(r'[^A-Z0-9 ]', '', txt)
    return re.sub(r'\s+', ' ', txt)

# --- Chargement des Données ---
@st.cache_data(show_spinner="Chargement des vacataires et de l'EDT...")
def charger_donnees():
    # 1. Charger Vacataires
    df_pv = pd.read_excel(FICHIER_PV, sheet_name="Permanents-Vacataires", dtype=str)
    df_pv.columns = [nettoyer_texte(c) for c in df_pv.columns]
    
    # Filtrer uniquement les vacataires
    df_vac = df_pv[df_pv["Qualité"].str.strip().str.lower() == "vacataire"].copy()
    df_vac["NOM_NORM"] = df_vac["NOM"].apply(normaliser_nom)
    
    # 2. Charger Emploi du temps (EDT)
    df_edt = pd.read_excel(FICHIER_EDT, sheet_name=0, dtype=str)
    df_edt.columns = [nettoyer_texte(c) for c in df_edt.columns]
    df_edt["ENSEIGNANT_NORM"] = df_edt["Enseignants"].apply(normaliser_nom)
    
    return df_vac, df_edt

# --- Extraction de la Charge Horaire par Vacataire ---
def extraire_charge_vacataire(nom_vacataire_norm, df_edt):
    rows = df_edt[df_edt["ENSEIGNANT_NORM"] == nom_vacataire_norm]
    if rows.empty:
        # Recherche partielle si correspondance exacte échouée
        rows = df_edt[df_edt["ENSEIGNANT_NORM"].str.contains(nom_vacataire_norm, na=False) | 
                      df_edt["ENSEIGNANT_NORM"].apply(lambda x: nom_vacataire_norm in x if x else False)]
    
    if rows.empty:
        return [], "", "Non affecté"

    modules = []
    types_cours = set()
    total_heures_semaine = 0.0

    for _, row in rows.iterrows():
        enseignement = nettoyer_texte(row["Enseignements"])
        code = nettoyer_texte(row["Code"]).upper()
        promo = nettoyer_texte(row["Promotion"])
        
        mod_label = f"{promo}-{enseignement}"
        modules.append(mod_label)
        
        if "COURS" in code or "COURS" in enseignement.upper():
            types_cours.add("Cours")
            total_heures_semaine += 1.5
        elif "TP" in code or "TP" in enseignement.upper():
            types_cours.add("TP")
            total_heures_semaine += 1.5
        elif "TD" in code or "TD" in enseignement.upper():
            types_cours.add("TD")
            total_heures_semaine += 1.5
        else:
            types_cours.add("TD/TP")
            total_heures_semaine += 1.5

    modules_str = " // ".join(list(dict.fromkeys(modules)))
    type_str = " + ".join(sorted(list(types_cours))) if types_cours else "C+TD+TP"
    
    return rows, modules_str, type_str

# --- Interface Streamlit ---
try:
    df_vacataires, df_edt = charger_donnees()
    st.sidebar.header("🔍 Sélection & Filtres")
    
    list_vacataires = sorted(df_vacataires["NOM"].dropna().unique())
    vacataire_select = st.sidebar.selectbox("Sélectionner un enseignant vacataire", list_vacataires)
    
    vac_info = df_vacataires[df_vacataires["NOM"] == vacataire_select].iloc[0]
    nom_norm = normaliser_nom(vacataire_select)
    
    edt_rows, modules_taught, type_teaching = extraire_charge_vacataire(nom_norm, df_edt)
    
    st.subheader(f"📌 Fiche de l'enseignant vacataire : {vacataire_select}")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Grade / Diplôme", vac_info.get("Grade", "Non spécifié"))
    col2.metric("Téléphone", vac_info.get("N°/TEL", "Non renseigné"))
    col3.metric("Email", vac_info.get("Email", "Non renseigné"))
    col4.metric("Séances EDT affectées", len(edt_rows))
    
    st.markdown("### 📚 Modules dispensés selon l'Emploi du Temps (S1)")
    if not edt_rows.empty:
        st.dataframe(edt_rows[["Enseignements", "Code", "Horaire", "Jours", "Lieu", "Promotion"]], use_container_width=True)
    else:
        st.warning("Aucune séance trouvée dans l'emploi du temps actuel pour cet enseignant.")

    st.divider()
    st.subheader("📝 Génération de l'état d'Heures Supplémentaires (Publipostage en Arabe)")
    
    # Formulaire de personnalisation des champs en arabe
    with st.form("form_hsup"):
        st.markdown("#### Informations complémentaires pour la fiche d'état :")
        c_ar1, c_ar2 = st.columns(2)
        with c_ar1:
            nom_arabe = st.text_input("اللقب (Nom en Arabe)", value="")
            prenom_arabe = st.text_input("الاسم (Prénom en Arabe)", value="")
        with c_ar2:
            grade_arabe = st.text_input("الرتبة (Grade)", value="أستاذ مؤقت")
            affiliation_arabe = st.text_input("الهيئة المستخدمة الأصلية (Affiliation)", value="الالكتروتقني")

        modules_arabe = st.text_area("المقياس المدرس (Modules enseignés)", value=modules_taught)
        type_enseignement_arabe = st.text_input("نوع التدريس (Type d'enseignement)", value=f"Cours+TD+TP" if type_teaching else "C+TD")
        
        nb_heures_mois = st.number_input("Nombre d'heures par mois (افتراضي)", min_value=0, max_value=200, value=48, step=6)
        
        bouton_generer = st.form_submit_button("⚡ Générer l'État Excel H.SUP (Arabe)")

    if bouton_generer:
        wb = openpyxl.load_workbook(FICHIER_TEMPLATE)
        ws = wb.active
        
        # Inscription des métadonnées dans les cellules du modèle Excel d'état H.SUP
        ws['A7'] = f"الاسم :{prenom_arabe if prenom_arabe else vacataire_select}"
        ws['C7'] = f"اللقب :{nom_arabe}"
        ws['A8'] = f"الرتبة :{grade_arabe}"
        ws['C8'] = f"الهيئة المستخدمةالاصلية  :{affiliation_arabe}"
        ws['A9'] = f"المقياس المدرس :{modules_arabe}"
        ws['D9'] = f"نوع التدريس :{type_enseignement_arabe}"
        
        # Sauvegarde en mémoire pour le téléchargement
        output_buffer = io.BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)
        
        file_name = f"Etat_HSUP_VAC_{normaliser_nom(vacataire_select).replace(' ', '_')}.xlsx"
        
        st.success("✅ Fichier d'état généré avec succès !")
        st.download_button(
            label="📥 Télécharger l'état d'Heures Supplémentaires (Excel)",
            data=output_buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

except Exception as e:
    st.error(f"Erreur d'initialisation ou de traitement des fichiers : {e}")