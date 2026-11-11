# -*- coding: utf-8 -*-
"""
Application de Gestion et Génération des Heures Supplémentaires des Vacataires
Département d'Électrotechnique - Faculté de Génie Électrique - UDL-SBA (Année 2026-2027)
"""

import streamlit as st
import pandas as pd
import openpyxl
import io
import re
from pathlib import Path

st.set_page_config(
    page_title="Gestion H.SUP Vacataires - 2026-2027",
    page_icon="⚡",
    layout="wide"
)

TITRE_OFFICIEL = "Plateforme de gestion des EDTs et H.SUP - Vacataires - 2026-2027"
st.title(f"⚡ {TITRE_OFFICIEL}")
st.markdown("**Département d'Électrotechnique - Faculté de génie électrique - UDL-SBA**")
st.divider()

BASE_DIR = Path(__file__).resolve().parent
FICHIER_PV = BASE_DIR / "Permanents-Vacataires-ELT2-2026-2027.xlsx"
FICHIER_EDT = BASE_DIR / "dataEDT-ELT-S1-2027.xlsx"
FICHIER_COORD = BASE_DIR / "Vos coordonnées en Arabe (réponses).xlsx"
FICHIER_TEMPLATE = BASE_DIR / "Etats H.SUP VAC 2025-2026.xlsx"

def nettoyer_texte(valeur):
    if valeur is None or pd.isna(valeur):
        return ""
    return re.sub(r'\s+', ' ', str(valeur)).strip()

def normaliser_nom(nom):
    txt = nettoyer_texte(nom).upper()
    for src, dst in [('É','E'), ('È','E'), ('Ê','E'), ('Ë','E'), ('À','A'), ('Â','A'), ('Ä','A'), ('Î','I'), ('Ï','I'), ('Ô','O'), ('Ö','O'), ('Ù','U'), ('Û','U'), ('Ü','U')]:
        txt = txt.replace(src, dst)
    txt = re.sub(r'[^A-Z0-9 ]', '', txt)
    return re.sub(r'\s+', ' ', txt).strip()

@st.cache_data(show_spinner="Chargement et filtrage des données...")
def charger_donnees():
    # 1. Vacataires
    df_pv = pd.read_excel(FICHIER_PV, sheet_name="Permanents-Vacataires", dtype=str)
    df_pv.columns = [nettoyer_texte(c) for c in df_pv.columns]
    df_vac = df_pv[df_pv["Qualité"].str.strip().str.lower() == "vacataire"].copy()
    df_vac["NOM_NORM"] = df_vac["NOM"].apply(normaliser_nom)

    # 2. EDT
    df_edt = pd.read_excel(FICHIER_EDT, sheet_name=0, dtype=str)
    df_edt.columns = [nettoyer_texte(c) for c in df_edt.columns]
    df_edt["ENSEIGNANT_NORM"] = df_edt["Enseignants"].apply(normaliser_nom)

    # 3. Coordonnées en arabe
    df_coord = pd.read_excel(FICHIER_COORD, dtype=str)
    df_coord.columns = [nettoyer_texte(c) for c in df_coord.columns]
    
    col_nom_fr = next((c for c in df_coord.columns if "francais" in c.lower() or "français" in c.lower()), df_coord.columns[2])
    col_nom_ar = next((c for c in df_coord.columns if "arabe" in c.lower()), df_coord.columns[3])

    dict_arabe = {}
    for _, row in df_coord.iterrows():
        nom_fr = nettoyer_texte(row.get(col_nom_fr, ""))
        nom_ar = nettoyer_texte(row.get(col_nom_ar, ""))
        if nom_fr:
            dict_arabe[normaliser_nom(nom_fr)] = nom_ar

    # Filtrer uniquement les vacataires qui ont une charge dans l'EDT
    vacataires_avec_charge = []
    for _, vac_row in df_vac.iterrows():
        nom_v = vac_row["NOM"]
        nom_n = vac_row["NOM_NORM"]
        seances = df_edt[df_edt["ENSEIGNANT_NORM"] == nom_n]
        if seances.empty:
            seances = df_edt[df_edt["ENSEIGNANT_NORM"].apply(lambda x: nom_n in x if x else False)]
        
        if not seances.empty:
            vacataires_avec_charge.append(nom_v)

    return df_vac, df_edt, dict_arabe, sorted(list(set(vacataires_avec_charge)))

try:
    df_vacataires, df_edt, dict_noms_arabe, liste_vacataires_actifs = charger_donnees()

    st.sidebar.header("🔍 Sélection des Vacataires (Avec Charge)")
    if not liste_vacataires_actifs:
        st.error("Aucun enseignant vacataire avec une charge horaire n'a été trouvé dans l'EDT.")
        st.stop()

    vacataire_select = st.sidebar.selectbox("Choisir un enseignant vacataire", liste_vacataires_actifs)

    # Informations de l'enseignant
    row_vac = df_vacataires[df_vacataires["NOM"] == vacataire_select].iloc[0]
    nom_norm = normaliser_nom(vacataire_select)
    nom_prenom_arabe = dict_noms_arabe.get(nom_norm, "")

    # Séparation du nom et prénom en arabe
parties_ar = nom_prenom_arabe.split()
if len(parties_ar) >= 2:
    prenom_ar = parties_ar[0]
    nom_ar = " ".join(parties_ar[1:])
else:
    prenom_ar = nom_prenom_arabe
    nom_ar = ""

    # Recherche des séances EDT
    seances_vac = df_edt[df_edt["ENSEIGNANT_NORM"] == nom_norm]
    if seances_vac.empty:
        seances_vac = df_edt[df_edt["ENSEIGNANT_NORM"].apply(lambda x: nom_norm in x if x else False)]

    nb_seances = len(seances_vac)
    total_heures_semaine = nb_seances * 1.5

    st.subheader(f"👨‍🏫 Fiche Enseignant : {vacataire_select} ({nom_prenom_arabe})")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Nom complet en Arabe", nom_prenom_arabe if nom_prenom_arabe else "Non renseigné")
    c2.metric("Grade", row_vac.get("Grade", "Vacataire"))
    c3.metric("Séances par semaine", nb_seances)
    c4.metric("Volume horaire hebdo", f"{total_heures_semaine} H")

    st.markdown("### 📅 Séances hebdomadaires de l'Emploi du Temps (S1)")
    st.dataframe(seances_vac[["Enseignements", "Code", "Horaire", "Jours", "Lieu", "Promotion"]], use_container_width=True)

    st.divider()
    st.subheader("⚙️ Génération et Actualisation de l'État H.SUP (2026-2027)")

    if st.button("🚀 Générer l'état Excel officiel actualisé", type="primary", use_container_width=True):
        wb = openpyxl.load_workbook(FICHIER_TEMPLATE)
        ws = wb.active

        # Inscription séparée du Nom et Prénom en Arabe selon le modèle requis (ex: A7, C7)
        ws['A7'] = f"الاسم : {prenom_ar}"
        ws['C7'] = f"اللقب : {nom_ar}"
        ws['A8'] = f"الرتبة : أستاذ مؤقت"
        ws['C8'] = f"الهيئة المستخدمة الاصلية : الالكتروتقني"
        
        # Modules enseignés
        modules_list = " // ".join(seances_vac["Enseignements"].dropna().unique())
        ws['A9'] = f"المقياس المدرس : {modules_list[:120]}"
        ws['D9'] = f"نوع التدريس : C+TD+TP"

        # Nettoyage des anciennes valeurs et mise en place des nouveaux volumes horaires calculés
        for r in range(11, 40):
            for col in ['B', 'C', 'D']:
                cell = ws[f'{col}{r}']
                if cell.value and "المجموع" in str(cell.value):
                    continue
                if col == 'B':
                    cell.value = f"{total_heures_semaine}سا"

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        nom_fichier_clean = normaliser_nom(vacataire_select).replace(" ", "_")
        file_name = f"Etat_HSUP_Vacataire_{nom_fichier_clean}_2026_2027.xlsx"

        st.success("✅ Fichier Excel mis à jour et nettoyé avec succès pour l'année 2026-2027 !")
        st.download_button(
            label="📥 Télécharger l'état Excel actualisé",
            data=buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

except Exception as e:
    st.error(f"Une erreur est survenue lors du traitement : {e}")
