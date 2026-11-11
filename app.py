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
from datetime import datetime, date, timedelta
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
FICHIER_TEMPLATE = BASE_DIR / "Etat_HSUP_Vacataire_2026_2027.xlsx"

def nettoyer_texte(valeur):
    if valeur is None or pd.isna(valeur):
        return ""
    return re.sub(r'\s+', ' ', str(valeur)).strip()

def normaliser_nom(nom):
    if nom is None or pd.isna(nom):
        return ""
    txt = str(nom).upper()
    for src, dst in [('É','E'), ('È','E'), ('Ê','E'), ('Ë','E'), ('À','A'), ('Â','A'), ('Ä','A'), ('Î','I'), ('Ï','I'), ('Ô','O'), ('Ö','O'), ('Ù','U'), ('Û','U'), ('Ü','U')]:
        txt = txt.replace(src, dst)
    txt = re.sub(r'[^A-Z0-9 ]', '', txt)
    return re.sub(r'\s+', ' ', txt).strip()

@st.cache_data(show_spinner="Chargement et synchronisation des sources...")
def charger_donnees():
    df_pv = pd.read_excel(FICHIER_PV, sheet_name="Permanents-Vacataires", dtype=str)
    df_pv.columns = [nettoyer_texte(c) for c in df_pv.columns]
    df_pv["NOM"] = df_pv["NOM"].fillna("").astype(str)
    df_pv["Qualité"] = df_pv["Qualité"].fillna("").astype(str)
    
    df_vac = df_pv[df_pv["Qualité"].str.strip().str.lower() == "vacataire"].copy()
    df_vac["NOM_NORM"] = df_vac["NOM"].apply(normaliser_nom)

    df_edt = pd.read_excel(FICHIER_EDT, sheet_name=0, dtype=str)
    df_edt.columns = [nettoyer_texte(c) for c in df_edt.columns]
    df_edt["Enseignants"] = df_edt["Enseignants"].fillna("").astype(str)
    df_edt["ENSEIGNANT_NORM"] = df_edt["Enseignants"].apply(normaliser_nom)

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

    vacataires_avec_charge = []
    for _, vac_row in df_vac.iterrows():
        nom_v = vac_row["NOM"]
        nom_n = vac_row["NOM_NORM"]
        if not nom_v or not nom_n:
            continue
        seances = df_edt[df_edt["ENSEIGNANT_NORM"] == nom_n]
        if seances.empty:
            seances = df_edt[df_edt["ENSEIGNANT_NORM"].apply(lambda x: nom_n in x if x else False)]
        
        if not seances.empty:
            vacataires_avec_charge.append(str(nom_v))

    liste_triee = sorted(list(set([str(x) for x in vacataires_avec_charge if pd.notna(x) and str(x).strip()])))
    return df_vac, df_edt, dict_arabe, liste_triee

try:
    df_vacataires, df_edt, dict_noms_arabe, liste_vacataires_actifs = charger_donnees()

    st.sidebar.header("🔍 Paramètres & Calendrier")
    date_debut_cours = st.sidebar.date_input("Date de début des enseignements", value=date(2026, 9, 23))

    if not liste_vacataires_actifs:
        st.error("Aucun enseignant vacataire avec une charge horaire n'a été trouvé dans l'EDT.")
        st.stop()

    vacataire_select = st.sidebar.selectbox("Choisir un enseignant vacataire", liste_vacataires_actifs)

    row_vac = df_vacataires[df_vacataires["NOM"] == vacataire_select].iloc[0]
    nom_norm = normaliser_nom(vacataire_select)
    
    nom_prenom_arabe = dict_noms_arabe.get(nom_norm, "")
    if not nom_prenom_arabe:
        mots_v = set(nom_norm.split())
        for fr_key, ar_val in dict_noms_arabe.items():
            mots_k = set(fr_key.split())
            if mots_v.intersection(mots_k):
                nom_prenom_arabe = ar_val
                break

    parties_ar = nom_prenom_arabe.split()
    if len(parties_ar) >= 2:
        prenom_ar = parties_ar[0]
        nom_ar = " ".join(parties_ar[1:])
    else:
        prenom_ar = nom_prenom_arabe
        nom_ar = ""

    seances_vac = df_edt[df_edt["ENSEIGNANT_NORM"] == nom_norm]
    if seances_vac.empty:
        seances_vac = df_edt[df_edt["ENSEIGNANT_NORM"].apply(lambda x: nom_n in x if x else False)]

    nb_seances = len(seances_vac)
    total_heures_semaine_ref = nb_seances * 1.5

    st.subheader(f"👨‍🏫 Fiche Enseignant : {vacataire_select} ({nom_prenom_arabe if nom_prenom_arabe else 'Nom arabe non trouvé'})")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Nom en Arabe", nom_prenom_arabe if nom_prenom_arabe else "Non renseigné")
    c2.metric("Grade", row_vac.get("Grade", "Vacataire"))
    c3.metric("Séances par semaine", nb_seances)
    c4.metric("Volume hebdo ref", f"{total_heures_semaine_ref} H")

    st.markdown("### 📅 Séances hebdomadaires de l'Emploi du Temps (S1)")
    st.dataframe(seances_vac[["Enseignements", "Code", "Horaire", "Jours", "Lieu", "Promotion"]], use_container_width=True)

    st.divider()
    st.subheader("⚙️ Génération intelligente de l'État H.SUP (2026-2027)")

    if st.button("🚀 Générer l'état Excel officiel actualisé", type="primary", use_container_width=True):
        fichier_a_ouvrir = FICHIER_TEMPLATE if FICHIER_TEMPLATE.exists() else (BASE_DIR / "Etats H.SUP VAC 2025-2026.xlsx")
        
        wb = openpyxl.load_workbook(fichier_a_ouvrir)
        ws = wb.active

        ws['A7'] = f"الاسم : {prenom_ar}"
        ws['C7'] = f"اللقب : {nom_ar}"
        ws['A8'] = f"الرتبة : أستاذ مؤقت  "
        ws['C8'] = f"الهيئة المستخدمةالاصلية  : الالكتروتقني  "
        
        modules_list = " // ".join(seances_vac["Enseignements"].dropna().unique())
        ws['A9'] = f"المقياس المدرس : {modules_list[:120]}"
        ws['D9'] = f"نوع التدريس : C+TD+TP"

        # Jours fériés officiels en Algérie
        jours_feries = [
            date(2026, 11, 1), # 1er Novembre
            date(2027, 5, 1),  # 1er Mai
            date(2027, 7, 5),  # Fête de l'Indépendance
        ]

        mois_blocks = [
            {"weeks_rows": [12, 13, 14, 15, 16], "total_row": 17, "annee": 2026, "mois_num": 9},   # Septembre
            {"weeks_rows": [19, 20, 21, 22, 23], "total_row": 24, "annee": 2026, "mois_num": 10},  # Octobre
            {"weeks_rows": [26, 27, 28, 29, 30], "total_row": 31, "annee": 2026, "mois_num": 11},  # Novembre
            {"weeks_rows": [33, 34, 35, 36, 37], "total_row": 38, "annee": 2026, "mois_num": 12},  # Décembre
        ]

        total_general = 0.0

        for block in mois_blocks:
            total_mois = 0.0
            annee_m = block["annee"]
            mois_m = block["mois_num"]

            for i, wr in enumerate(block["weeks_rows"]):
                cell_hrs = ws.cell(row=wr, column=2)
                
                # Génération des jours de cette semaine (découpage par blocs de 7 jours dans le mois)
                jour_debut_semaine = (i * 7) + 1
                jour_fin_semaine = min(jour_debut_semaine + 6, 31)

                jours_travailles_valides = 0
                jours_totaux_semaine = 0

                for j in range(jour_debut_semaine, jour_fin_semaine + 1):
                    try:
                        d_courante = date(annee_m, mois_m, j)
                        jours_totaux_semaine += 1
                        # Vérification effective : >= date début, pas vendredi (4) ou samedi (5), pas jour férié
                        if d_courante >= date_debut_cours:
                            if d_courante.weekday() not in [4, 5] and d_courante not in jours_feries:
                                jours_travailles_valides += 1
                    except ValueError:
                        break

                # Calcul proportionnel strict basé sur les jours effectifs de travail de la semaine
                if jours_travailles_valides > 0 and jours_totaux_semaine > 0:
                    heures_semaine = total_heures_semaine_ref * (jours_travailles_valides / jours_totaux_semaine)
                    cell_hrs.value = f"{round(heures_semaine, 1)}سا"
                    total_mois += heures_semaine
                else:
                    cell_hrs.value = None

            ws.cell(row=block["total_row"], column=2).value = f"{round(total_mois, 1)}سا" if total_mois > 0 else None
            total_general += total_mois

        ws.cell(row=39, column=2).value = f"{round(total_general, 1)}سا" if total_general > 0 else None

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        nom_fichier_clean = normaliser_nom(vacataire_select).replace(" ", "_")
        file_name = f"Etat_HSUP_Vacataire_{nom_fichier_clean}_2026_2027.xlsx"

        st.success(f"✅ Fichier d'état Excel généré avec succès ! (Total général : {round(total_general, 1)} H)")
        st.download_button(
            label="📥 Télécharger l'état Excel actualisé",
            data=buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

except Exception as e:
    st.error(f"Une erreur est survenue lors du traitement : {e}")
