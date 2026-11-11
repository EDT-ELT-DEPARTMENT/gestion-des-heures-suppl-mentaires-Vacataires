# -*- coding: utf-8 -*-
"""
Gestion et génération des états H.SUP des enseignants vacataires - ELT
Année universitaire 2026-2027.
Le calcul est fondé sur les séances récurrentes de l'EDT et les dates réelles
de chaque mois (pas sur une proratisation du volume hebdomadaire).
"""

import io
import re
import calendar
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(
    page_title="Gestion H.SUP Vacataires - 2026-2027",
    page_icon="⚡",
    layout="wide",
)

TITRE_OFFICIEL = "Plateforme de gestion des EDTs et H.SUP - Vacataires - 2026-2027"
st.title(f"⚡ {TITRE_OFFICIEL}")
st.markdown("**Département d'Électrotechnique - Faculté de génie électrique - UDL-SBA**")
st.divider()

BASE_DIR = Path(__file__).resolve().parent
FICHIER_PV = BASE_DIR / "Permanents-Vacataires-ELT2-2026-2027.xlsx"
FICHIER_EDT = BASE_DIR / "dataEDT-ELT-S1-2027.xlsx"
FICHIER_COORD = BASE_DIR / "Vos coordonnées en Arabe (réponses).xlsx"  # facultatif
FICHIER_TEMPLATE = BASE_DIR / "Etat_HSUP_Vacataire_2026_2027.xlsx"

# Les blocs et lignes correspondent à la feuille « Nom & prénom » du modèle fourni.
MOIS_BLOCKS = [
    {"nom": "Septembre 2026", "annee": 2026, "mois": 9,
     "lignes_semaines": [12, 13, 14, 15, 16], "ligne_total": 17, "ligne_entete": 11},
    {"nom": "Octobre 2026", "annee": 2026, "mois": 10,
     "lignes_semaines": [19, 20, 21, 22, 23], "ligne_total": 24, "ligne_entete": 18},
    {"nom": "Novembre 2026", "annee": 2026, "mois": 11,
     "lignes_semaines": [26, 27, 28, 29, 30], "ligne_total": 31, "ligne_entete": 25},
]

# Jours fériés fixes applicables à la période couverte par le modèle.
# Les fêtes religieuses mobiles doivent être ajoutées après publication officielle.
JOURS_FERIES = {
    date(2026, 11, 1): "1er Novembre – Révolution",
    date(2027, 5, 1): "Fête du Travail",
    date(2027, 7, 5): "Fête de l'Indépendance",
}

JOURS_FR = {
    "dimanche": 6, "lundi": 0, "mardi": 1, "mercredi": 2,
    "jeudi": 3, "vendredi": 4, "samedi": 5,
}


def nettoyer_texte(valeur):
    if valeur is None or pd.isna(valeur):
        return ""
    return re.sub(r"\s+", " ", str(valeur)).strip()


def normaliser_nom(nom):
    txt = nettoyer_texte(nom).upper()
    remplacements = {
        "É": "E", "È": "E", "Ê": "E", "Ë": "E", "À": "A",
        "Â": "A", "Ä": "A", "Î": "I", "Ï": "I", "Ô": "O",
        "Ö": "O", "Ù": "U", "Û": "U", "Ü": "U", "Ç": "C",
    }
    for src, dst in remplacements.items():
        txt = txt.replace(src, dst)
    txt = re.sub(r"[^A-Z0-9 ]", " ", txt)
    return re.sub(r"\s+", " ", txt).strip()


def lire_heure(h):
    """Convertit 8h, 8h30, 08:30 ou 8.30 en minutes après minuit."""
    s = nettoyer_texte(h).lower().replace(" ", "").replace("H", "h")
    m = re.fullmatch(r"(\d{1,2})(?:h|:|\.)(\d{1,2})?", s)
    if not m:
        raise ValueError(f"Format d'heure non reconnu : {h}")
    heure = int(m.group(1))
    minute = int(m.group(2) or 0)
    if heure > 23 or minute > 59:
        raise ValueError(f"Heure invalide : {h}")
    return heure * 60 + minute


def duree_seance_heures(horaire):
    """Durée réelle entre les deux bornes de l'horaire EDT."""
    s = nettoyer_texte(horaire).replace("–", "-").replace("—", "-")
    morceaux = re.split(r"\s*-\s*", s)
    if len(morceaux) != 2:
        raise ValueError(f"Horaire EDT non reconnu : {horaire}")
    debut = lire_heure(morceaux[0])
    fin = lire_heure(morceaux[1])
    if fin <= debut:
        raise ValueError(f"Fin antérieure ou égale au début : {horaire}")
    return (fin - debut) / 60.0


def correspond_enseignant(champ_enseignants, nom_normalise):
    """Correspondance du nom complet, sans confondre deux noms partiels."""
    champ = normaliser_nom(champ_enseignants)
    if not champ or not nom_normalise:
        return False
    if champ == nom_normalise:
        return True
    # L'EDT peut contenir plusieurs enseignants dans une même cellule.
    # Les limites de mots évitent, par exemple, de confondre BENALI et BENALIYA.
    return re.search(r"(?<![A-Z0-9])" + re.escape(nom_normalise) +
                     r"(?![A-Z0-9])", champ) is not None


def charger_donnees():
    for f in (FICHIER_PV, FICHIER_EDT, FICHIER_TEMPLATE):
        if not f.exists():
            raise FileNotFoundError(f"Fichier indispensable introuvable : {f.name}")

    df_pv = pd.read_excel(FICHIER_PV, sheet_name="Permanents-Vacataires", dtype=str)
    df_pv.columns = [nettoyer_texte(c) for c in df_pv.columns]
    colonnes_pv = {"NOM", "Qualité"}
    if not colonnes_pv.issubset(df_pv.columns):
        raise ValueError("Le fichier personnel doit contenir les colonnes NOM et Qualité.")
    df_pv["NOM"] = df_pv["NOM"].fillna("").astype(str).str.strip()
    df_pv["Qualité"] = df_pv["Qualité"].fillna("").astype(str).str.strip()
    df_vac = df_pv[df_pv["Qualité"].str.casefold().eq("vacataire")].copy()
    df_vac["NOM_NORM"] = df_vac["NOM"].apply(normaliser_nom)

    df_edt = pd.read_excel(FICHIER_EDT, sheet_name=0, dtype=str)
    df_edt.columns = [nettoyer_texte(c) for c in df_edt.columns]
    obligatoires = {"Enseignements", "Enseignants", "Horaire", "Jours"}
    if not obligatoires.issubset(df_edt.columns):
        raise ValueError("Colonnes manquantes dans l'EDT : " +
                         ", ".join(sorted(obligatoires - set(df_edt.columns))))
    for col in obligatoires:
        df_edt[col] = df_edt[col].fillna("").astype(str).str.strip()
    if "Département" in df_edt.columns:
        df_edt = df_edt[df_edt["Département"].str.strip().str.upper().eq("ELT")].copy()

    df_edt["JOUR_NORM"] = df_edt["Jours"].apply(normaliser_nom).str.lower()
    df_edt["DUREE_H"] = df_edt["Horaire"].apply(
        lambda x: duree_seance_heures(x) if x else 0.0
    )

    # Le fichier des coordonnées arabes n'est pas fourni avec tous les déploiements.
    dict_arabe = {}
    if FICHIER_COORD.exists():
        df_coord = pd.read_excel(FICHIER_COORD, dtype=str)
        df_coord.columns = [nettoyer_texte(c) for c in df_coord.columns]
        col_nom_fr = next(
            (c for c in df_coord.columns if "francais" in c.lower() or "français" in c.lower()),
            None,
        )
        col_nom_ar = next((c for c in df_coord.columns if "arabe" in c.lower()), None)
        if col_nom_fr and col_nom_ar:
            for _, row in df_coord.iterrows():
                nom_fr = nettoyer_texte(row.get(col_nom_fr, ""))
                nom_ar = nettoyer_texte(row.get(col_nom_ar, ""))
                if nom_fr and nom_ar:
                    dict_arabe[normaliser_nom(nom_fr)] = nom_ar

    actifs = []
    for _, vac in df_vac.iterrows():
        nom_n = vac["NOM_NORM"]
        if nom_n and df_edt["Enseignants"].apply(
            lambda x: correspond_enseignant(x, nom_n)
        ).any():
            actifs.append(vac["NOM"])
    return df_vac, df_edt, dict_arabe, sorted(set(actifs))


def dates_du_mois(annee, mois):
    nb = calendar.monthrange(annee, mois)[1]
    return [date(annee, mois, j) for j in range(1, nb + 1)]


def etiquettes_dates(annee, mois):
    labels = {}
    for d in dates_du_mois(annee, mois):
        jour = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"][d.weekday()]
        suffixe = f" — FÉRIÉ : {JOURS_FERIES[d]}" if d in JOURS_FERIES else ""
        labels[d.isoformat()] = f"{jour} {d.strftime('%d/%m/%Y')}{suffixe}"
    return labels


def calculer_heures_par_semaine(seances, annee, mois, date_debut, exclusions):
    """
    Le modèle fourni réserve cinq lignes par mois : les semaines sont donc
    les tranches 1–7, 8–14, 15–21, 22–28 et 29–fin du mois.
    Pour chaque date admissible, on additionne les séances EDT de son jour.
    Vendredi/samedi, jours fériés, dates antérieures au début et exclusions
    manuelles sont retirés.
    """
    resultats = []
    jours_invalides = {4, 5}  # vendredi et samedi en Algérie
    for numero in range(5):
        debut_jour = numero * 7 + 1
        fin_jour = min(debut_jour + 6, calendar.monthrange(annee, mois)[1])
        total = 0.0
        dates_comptees = []
        for jour_mois in range(debut_jour, fin_jour + 1):
            d = date(annee, mois, jour_mois)
            if d < date_debut or d.weekday() in jours_invalides:
                continue
            if d in JOURS_FERIES or d in exclusions:
                continue
            jour_fr = normaliser_nom(d.strftime("%A"))
            # strftime dépend de la langue système : on utilise weekday() pour
            # obtenir le libellé français de manière indépendante de Windows.
            nom_jour = {
                0: "lundi", 1: "mardi", 2: "mercredi", 3: "jeudi",
                4: "vendredi", 5: "samedi", 6: "dimanche",
            }[d.weekday()]
            lignes = seances[seances["JOUR_NORM"].eq(normaliser_nom(nom_jour).lower())]
            if not lignes.empty:
                total += float(lignes["DUREE_H"].sum())
                dates_comptees.append(d)
        resultats.append({
            "semaine": numero + 1,
            "debut": debut_jour,
            "fin": fin_jour,
            "heures": round(total, 2),
            "dates": dates_comptees,
        })
    return resultats


def format_heures_arabe(valeur):
    """Exemple : 4سا30د ; les fractions de quart d'heure restent exactes."""
    minutes = int(round(float(valeur) * 60))
    heures, reste = divmod(minutes, 60)
    if heures and reste:
        return f"{heures}سا{reste}د"
    if heures:
        return f"{heures}سا"
    if reste:
        return f"{reste}د"
    return None


def generer_excel(nom_vacataire, seances, dict_arabe, date_debut, exclusions_par_mois):
    wb = openpyxl.load_workbook(FICHIER_TEMPLATE)
    nom_feuille = "Nom & prénom" if "Nom & prénom" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[nom_feuille]
    # Le modèle contient aussi une ancienne fiche d'exemple (2023) :
    # on la retire pour ne pas joindre les données d'un autre enseignant.
    for autre_feuille in list(wb.sheetnames):
        if autre_feuille != nom_feuille:
            del wb[autre_feuille]

    nom_norm = normaliser_nom(nom_vacataire)
    nom_arabe = dict_arabe.get(nom_norm, "")
    prenom_ar = ""
    nom_ar = ""
    if nom_arabe:
        parties = nom_arabe.split()
        if len(parties) >= 2:
            prenom_ar, nom_ar = parties[0], " ".join(parties[1:])
        else:
            prenom_ar = nom_arabe

    # En-tête du formulaire fourni (feuille « Nom & prénom »).
    ws["A7"] = f"الاسم :{prenom_ar}" if prenom_ar else "الاسم :"
    ws["C7"] = f"اللقب :{nom_ar}" if nom_ar else f"اللقب :{nom_vacataire}"
    ws["A8"] = "الرتبة :أستاذ مؤقت"
    ws["C8"] = "الهيئة المستخدمة الأصلية : الالكتروتقني"

    matieres = sorted({
        nettoyer_texte(x) for x in seances["Enseignements"].tolist()
        if nettoyer_texte(x)
    })
    ws["A9"] = "المقياس المدرس : " + " // ".join(matieres)
    types = sorted({
        nettoyer_texte(x).split("-", 1)[0]
        for x in matieres if nettoyer_texte(x)
    })
    ws["D9"] = "نوع التدريس : " + "+".join(types)

    total_general = 0.0
    details = []
    for bloc in MOIS_BLOCKS:
        annee, mois = bloc["annee"], bloc["mois"]
        exclusions = exclusions_par_mois.get(mois, set())
        calculs = calculer_heures_par_semaine(
            seances, annee, mois, date_debut, exclusions
        )
        total_mois = 0.0
        for ligne, semaine in zip(bloc["lignes_semaines"], calculs):
            # Conserver les libellés arabes et ajouter la période exacte.
            ws.cell(ligne, 1).value = (
                f"الاسبوع {semaine['semaine']} : "
                f"{semaine['debut']:02d}-{semaine['fin']:02d}/{mois:02d}"
            )
            ws.cell(ligne, 2).value = format_heures_arabe(semaine["heures"])
            total_mois += semaine["heures"]
            details.append({
                "Mois": bloc["nom"],
                "Semaine": semaine["semaine"],
                "Période": f"{semaine['debut']:02d}–{semaine['fin']:02d}/{mois:02d}",
                "Heures calculées": semaine["heures"],
                "Dates exclues": ", ".join(
                    d.strftime("%d/%m/%Y") for d in sorted(exclusions)
                ),
            })
        ws.cell(bloc["ligne_total"], 2).value = format_heures_arabe(total_mois)
        total_general += total_mois

    ws["B32"] = format_heures_arabe(total_general)
    nom_onglet = re.sub(r"[\\\\/*?:\\[\\]]", " ", nom_vacataire).strip()[:31] or "Etat H.SUP"
    ws.title = nom_onglet
    ws["A33"] = (
        "ملاحظة : تم احتساب الساعات حسب الحصص الفعلية في جدول التوقيت، "
        "مع استثناء أيام العطل والتواريخ المحددة للاستبعاد."
    )
    # Une feuille de contrôle permet de vérifier le détail du calcul.
    if "Détail calcul" in wb.sheetnames:
        del wb["Détail calcul"]
    wd = wb.create_sheet("Détail calcul")
    wd.append(["Mois", "Semaine", "Période", "Heures calculées", "Dates exclues"])
    for ligne in details:
        wd.append(list(ligne.values()))
    wd.freeze_panes = "A2"
    wd.auto_filter.ref = wd.dimensions
    for col, largeur in {"A": 22, "B": 12, "C": 18, "D": 20, "E": 45}.items():
        wd.column_dimensions[col].width = largeur

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, total_general, details


try:
    df_vacataires, df_edt, dict_noms_arabe, liste_vacataires_actifs = charger_donnees()

    st.sidebar.header("🔍 Paramètres & calendrier")
    date_debut_cours = st.sidebar.date_input(
        "Date de début des enseignements", value=date(2026, 9, 23)
    )
    vacataire_select = st.sidebar.selectbox(
        "Choisir un enseignant vacataire", liste_vacataires_actifs
    ) if liste_vacataires_actifs else None

    if not vacataire_select:
        st.error("Aucun enseignant vacataire avec une charge horaire n'a été trouvé dans l'EDT.")
        st.stop()

    nom_norm = normaliser_nom(vacataire_select)
    seances_vac = df_edt[
        df_edt["Enseignants"].apply(lambda x: correspond_enseignant(x, nom_norm))
    ].copy()

    # Contrôles de calendrier : les jours fériés sont exclus automatiquement.
    st.subheader("📅 Calendrier et jours à exclure")
    st.info(
        "Les vendredis et samedis sont automatiquement exclus. "
        "Les jours fériés nationaux fixes connus sont exclus automatiquement. "
        "Sélectionne ci-dessous toute autre date sans enseignement à comptabiliser."
    )
    exclusions_par_mois = {}
    colonnes = st.columns(len(MOIS_BLOCKS))
    for col, bloc in zip(colonnes, MOIS_BLOCKS):
        with col:
            st.markdown(f"**{bloc['nom']}**")
            labels = etiquettes_dates(bloc["annee"], bloc["mois"])
            choix = st.multiselect(
                f"Dates à exclure – {bloc['nom']}",
                options=list(labels.keys()),
                format_func=lambda x, d=labels: d[x],
                default=[],
                key=f"exclusions_{bloc['annee']}_{bloc['mois']}",
            )
            exclusions_par_mois[bloc["mois"]] = {
                date.fromisoformat(x) for x in choix
            }

    st.divider()
    st.subheader(f"👨‍🏫 Fiche enseignant : {vacataire_select}")
    nom_arabe = dict_noms_arabe.get(nom_norm, "Non renseigné")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Nom en arabe", nom_arabe)
    ligne_pv = df_vacataires[df_vacataires["NOM_NORM"].eq(nom_norm)].iloc[0]
    c2.metric("Grade", nettoyer_texte(ligne_pv.get("Grade", "")) or "Vacataire")
    c3.metric("Séances EDT", len(seances_vac))
    c4.metric("Volume hebdomadaire EDT", f"{seances_vac['DUREE_H'].sum():g} h")

    colonnes_aff = [
        c for c in ["Enseignements", "Code", "Horaire", "Jours", "Lieu", "Promotion"]
        if c in seances_vac.columns
    ]
    st.markdown("### Séances hebdomadaires issues de l'EDT")
    st.dataframe(seances_vac[colonnes_aff], use_container_width=True, hide_index=True)

    erreurs_horaires = seances_vac[
        seances_vac["DUREE_H"].isna()
    ] if "DUREE_H" in seances_vac.columns else pd.DataFrame()
    if not erreurs_horaires.empty:
        st.error("Certaines séances ont un horaire non interprétable. Corrige la source EDT.")

    if st.button("🚀 Générer l'état Excel officiel actualisé", type="primary",
                 use_container_width=True):
        buffer, total_general, details = generer_excel(
            vacataire_select, seances_vac, dict_noms_arabe,
            date_debut_cours, exclusions_par_mois
        )
        nom_fichier_clean = normaliser_nom(vacataire_select).replace(" ", "_")
        file_name = f"Etat_HSUP_Vacataire_{nom_fichier_clean}_2026_2027.xlsx"
        st.success(
            f"État Excel généré. Total calculé sur les dates admissibles : "
            f"{format_heures_arabe(total_general) or '0سا'}."
        )
        st.download_button(
            label="📥 Télécharger l'état Excel actualisé",
            data=buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("#### Vérification du calcul hebdomadaire")
        st.dataframe(pd.DataFrame(details), use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Une erreur est survenue lors du traitement : {e}")
    st.exception(e)
