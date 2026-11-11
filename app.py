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


def normaliser_entete(entete):
    """Normalise un intitulé de colonne pour reconnaître les variantes."""
    txt = nettoyer_texte(entete).lower()
    remplacements = {
        "é": "e", "è": "e", "ê": "e", "ë": "e", "à": "a",
        "â": "a", "ä": "a", "î": "i", "ï": "i", "ô": "o",
        "ö": "o", "ù": "u", "û": "u", "ü": "u",
    }
    for src, dst in remplacements.items():
        txt = txt.replace(src, dst)
    return re.sub(r"[^a-z0-9]+", " ", txt).strip()


def construire_cle_francaise(ligne, colonnes):
    """Construit une clé française à partir d'une colonne complète ou NOM/PRÉNOM."""
    col_complete = next((c for c in colonnes if any(
        terme in normaliser_entete(c)
        for terme in ("nom et prenom", "nom complet", "nom prenom")
    ) and "arabe" not in normaliser_entete(c)), None)
    if col_complete:
        return nettoyer_texte(ligne.get(col_complete, ""))

    col_nom = next((c for c in colonnes if normaliser_entete(c) in ("nom", "nom famille", "nom de famille")), None)
    col_prenom = next((c for c in colonnes if normaliser_entete(c) in ("prenom", "prenom usuel")), None)
    morceaux = []
    # Dans le fichier du personnel, NOM contient parfois déjà NOM + PRÉNOM.
    if col_nom:
        morceaux.append(nettoyer_texte(ligne.get(col_nom, "")))
    if col_prenom:
        morceaux.append(nettoyer_texte(ligne.get(col_prenom, "")))
    return " ".join(x for x in morceaux if x).strip()


def ajouter_correspondances_arabes(df, dict_arabe):
    """Lit la colonne « Nom et prénom en arabe » et l'associe au nom français."""
    if df is None or df.empty:
        return
    colonnes = list(df.columns)
    col_ar = next((c for c in colonnes if (
        "nom et prenom en arabe" in normaliser_entete(c)
        or ("arabe" in normaliser_entete(c) and "nom" in normaliser_entete(c))
    )), None)
    if col_ar is None:
        col_ar = next((c for c in colonnes if "arabe" in normaliser_entete(c)), None)
    if col_ar is None:
        return
    for _, ligne in df.iterrows():
        nom_fr = construire_cle_francaise(ligne, colonnes)
        nom_ar = nettoyer_texte(ligne.get(col_ar, ""))
        if nom_fr and nom_ar:
            dict_arabe[normaliser_nom(nom_fr)] = nom_ar


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

    # Le nom arabe peut être présent dans le fichier du personnel ou dans
    # le fichier des coordonnées. La colonne attendue est notamment
    # « Nom et prénom en arabe ».
    dict_arabe = {}
    ajouter_correspondances_arabes(df_pv, dict_arabe)
    if FICHIER_COORD.exists():
        df_coord = pd.read_excel(FICHIER_COORD, dtype=str)
        df_coord.columns = [nettoyer_texte(c) for c in df_coord.columns]
        ajouter_correspondances_arabes(df_coord, dict_arabe)

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


def remplir_fiche(ws, nom_vacataire, seances, dict_arabe, date_debut,
                  exclusions_par_mois):
    """Remplit une feuille du modèle officiel pour un seul enseignant."""
    nom_norm = normaliser_nom(nom_vacataire)
    nom_arabe = dict_arabe.get(nom_norm, "")
    prenom_ar = ""
    nom_ar = ""
    if nom_arabe:
        parties = nom_arabe.split()
        if len(parties) >= 2:
            prenom_ar = parties[0]
            nom_ar = " ".join(parties[1:])
        elif parties:
            prenom_ar = parties[0]

    ws["A7"] = f"الاسم : {prenom_ar}" if prenom_ar else "الاسم :"
    ws["C7"] = f"اللقب : {nom_ar}" if nom_ar else "اللقب :"
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
    heures_par_semaine = {}
    for bloc in MOIS_BLOCKS:
        annee, mois = bloc["annee"], bloc["mois"]
        exclusions = exclusions_par_mois.get(mois, set())
        calculs = calculer_heures_par_semaine(
            seances, annee, mois, date_debut, exclusions
        )
        total_mois = 0.0
        for ligne, semaine in zip(bloc["lignes_semaines"], calculs):
            ws.cell(ligne, 1).value = (
                f"الاسبوع {semaine['semaine']} : "
                f"{semaine['debut']:02d}-{semaine['fin']:02d}/{mois:02d}"
            )
            ws.cell(ligne, 2).value = format_heures_arabe(semaine["heures"])
            total_mois += semaine["heures"]
            cle = f"{bloc['nom']} – S{semaine['semaine']}"
            heures_par_semaine[cle] = semaine["heures"]
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
    ws["A33"] = (
        "ملاحظة : تم احتساب الساعات حسب الحصص الفعلية في جدول التوقيت، "
        "مع استثناء أيام العطل والتواريخ المحددة للاستبعاد."
    )
    return total_general, details, heures_par_semaine


def generer_excel(nom_vacataire, seances, dict_arabe, date_debut, exclusions_par_mois):
    """Génère la fiche individuelle selon le modèle officiel."""
    wb = openpyxl.load_workbook(FICHIER_TEMPLATE)
    nom_feuille = "Nom & prénom" if "Nom & prénom" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[nom_feuille]
    for autre_feuille in list(wb.sheetnames):
        if autre_feuille != nom_feuille:
            del wb[autre_feuille]

    total_general, details, _ = remplir_fiche(
        ws, nom_vacataire, seances, dict_arabe, date_debut, exclusions_par_mois
    )
    ws.title = re.sub(r"[\\\\/*?:\\[\\]]", " ", nom_vacataire).strip()[:31] or "Etat H.SUP"

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


def type_enseignement(intitule):
    """Retourne Cours, TD ou TP selon le préfixe exact de la matière."""
    texte = nettoyer_texte(intitule)
    prefixe = texte.split("-", 1)[0].strip().casefold() if "-" in texte else ""
    return prefixe if prefixe in {"cours", "td", "tp"} else ""


def cle_matiere_sans_groupe(intitule):
    """Ignore les suffixes de groupes SG12, G2, etc. pour compter les matières."""
    texte = nettoyer_texte(intitule)
    texte = re.sub(r"-(?:SG|G|GR)\s*\d+\s*$", "", texte, flags=re.IGNORECASE)
    return texte


def nom_onglet_unique(nom, deja_utilises):
    base = re.sub(r"[\\\\/*?:\\[\\]]", " ", nettoyer_texte(nom)).strip()
    base = base[:31] or "Enseignant"
    candidat = base
    compteur = 2
    while candidat.casefold() in deja_utilises:
        suffixe = f" ({compteur})"
        candidat = base[:31-len(suffixe)] + suffixe
        compteur += 1
    deja_utilises.add(candidat.casefold())
    return candidat


def generer_excel_groupe(df_vacataires, df_edt, liste_enseignants, dict_arabe,
                         date_debut, exclusions_par_mois):
    """
    Crée un seul classeur : index cliquable, récapitulatif, puis une fiche
    officielle par enseignant (une feuille par enseignant).
    """
    wb = openpyxl.load_workbook(FICHIER_TEMPLATE)
    feuille_modele = "Nom & prénom" if "Nom & prénom" in wb.sheetnames else wb.sheetnames[0]
    modele = wb[feuille_modele]

    # Conserver une seule feuille modèle, puis la dupliquer pour chaque enseignant.
    for nom in list(wb.sheetnames):
        if nom != feuille_modele:
            del wb[nom]

    # Feuilles de navigation placées au début du classeur.
    ws_index = wb.create_sheet("Index enseignants", 0)
    ws_recap = wb.create_sheet("Récapitulatif", 1)
    ws_index.append(["INDEX DES FICHES H.SUP – 2026-2027"])
    ws_index.append(["Enseignant", "Nom en arabe", "Total H.SUP", "Accès direct"])
    ws_recap.append(["RÉCAPITULATIF DES CHARGES D'ENSEIGNEMENT – ELT"])
    recap_headers = [
        "Enseignant", "Nom et prénom en arabe",
        "Nombre Cours", "Nombre TD", "Nombre TP",
        "Heures Cours / semaine", "Heures TD / semaine",
        "Heures TP / semaine", "Total heures EDT / semaine",
        "Total H.SUP période",
    ]
    semaines_cols = []
    for bloc in MOIS_BLOCKS:
        for numero in range(1, 6):
            semaines_cols.append(f"{bloc['nom']} S{numero}")
    ws_recap.append(recap_headers + semaines_cols)

    deja_utilises = {"index enseignants", "récapitulatif"}
    # La feuille modèle devient la première fiche ; les autres sont des copies exactes.
    enseignants = []
    for nom in liste_enseignants:
        nom_norm = normaliser_nom(nom)
        seances = df_edt[
            df_edt["Enseignants"].apply(lambda x: correspond_enseignant(x, nom_norm))
        ].copy()
        if seances.empty:
            continue
        nom_onglet = nom_onglet_unique(nom, deja_utilises)
        enseignants.append((nom, seances, nom_onglet))

    if not enseignants:
        raise ValueError("Aucun enseignant avec séances EDT à inclure dans le classeur.")

    # Renommer le modèle pour le premier enseignant et dupliquer pour les suivants.
    premier_nom, premieres_seances, premier_onglet = enseignants[0]
    modele.title = premier_onglet
    fiches = [(premier_nom, premieres_seances, modele)]
    for nom, seances, nom_onglet in enseignants[1:]:
        copie = wb.copy_worksheet(modele)
        copie.title = nom_onglet
        fiches.append((nom, seances, copie))

    # Les feuilles d'index et de récap sont créées avant les fiches : déplacer
    # explicitement l'ordre pour garantir Index, Récapitulatif, puis fiches.
    for nom, seances, ws_fiche in fiches:
        total, details, heures_semaines = remplir_fiche(
            ws_fiche, nom, seances, dict_arabe, date_debut, exclusions_par_mois
        )
        nom_arabe = dict_arabe.get(normaliser_nom(nom), "")
        types = seances["Enseignements"].apply(type_enseignement)
        heures_type = {
            t: float(seances.loc[types.eq(t), "DUREE_H"].sum())
            for t in ("cours", "td", "tp")
        }
        nb_type = {}
        for t in ("cours", "td", "tp"):
            matieres = {
                cle_matiere_sans_groupe(x)
                for x in seances.loc[types.eq(t), "Enseignements"].tolist()
                if nettoyer_texte(x)
            }
            nb_type[t] = len(matieres)

        # Index : nom cliquable vers la fiche de l'enseignant.
        ligne_index = ws_index.max_row + 1
        ws_index.cell(ligne_index, 1, nom)
        ws_index.cell(ligne_index, 2, nom_arabe)
        ws_index.cell(ligne_index, 3, round(total, 2))
        lien = ws_index.cell(ligne_index, 4, "Ouvrir la fiche")
        lien.hyperlink = f"#'{ws_fiche.title}'!A1"
        lien.style = "Hyperlink"

        ligne_recap = [
            nom, nom_arabe,
            nb_type["cours"], nb_type["td"], nb_type["tp"],
            round(heures_type["cours"], 2), round(heures_type["td"], 2),
            round(heures_type["tp"], 2),
            round(float(seances["DUREE_H"].sum()), 2),
            round(total, 2),
        ]
        ligne_recap.extend(round(heures_semaines.get(cle, 0.0), 2)
                           for cle in semaines_cols)
        ws_recap.append(ligne_recap)

    # Mise en forme des deux feuilles de synthèse.
    for ws in (ws_index, ws_recap):
        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:{openpyxl.utils.get_column_letter(ws.max_column)}{ws.max_row}"
        ws.sheet_view.showGridLines = False
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True, size=14, color="FFFFFF")
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="17365D")
        for cell in ws[2]:
            cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="4472C4")
        for row in ws.iter_rows(min_row=3):
            for cell in row:
                if cell.row % 2 == 1:
                    cell.fill = openpyxl.styles.PatternFill("solid", fgColor="EAF1F8")

    ws_index.column_dimensions["A"].width = 32
    ws_index.column_dimensions["B"].width = 28
    ws_index.column_dimensions["C"].width = 18
    ws_index.column_dimensions["D"].width = 20
    ws_recap.column_dimensions["A"].width = 30
    ws_recap.column_dimensions["B"].width = 28
    for col in range(3, ws_recap.max_column + 1):
        ws_recap.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 17
    for row in ws_recap.iter_rows(min_row=3, min_col=3):
        for cell in row:
            cell.number_format = '0.00'

    # Hyperliens de retour vers l'index dans chaque fiche.
    for _, _, ws_fiche in fiches:
        ws_fiche["F1"] = "Retour à l'index"
        ws_fiche["F1"].hyperlink = "#'Index enseignants'!A1"
        ws_fiche["F1"].style = "Hyperlink"

    # L'ordre final des onglets est index, récapitulatif, fiches.
    wb._sheets = [ws_index, ws_recap] + [ws for _, _, ws in fiches]

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer, len(fiches)


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
    if nom_arabe == "Non renseigné":
        st.warning(
            "Nom arabe introuvable pour cet enseignant. Vérifie que le fichier source "
            "contient la colonne « Nom et prénom en arabe » et une colonne permettant "
            "d'identifier le nom français correspondant. Les champs arabes resteront vides."
        )
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

    st.divider()
    st.subheader("📚 Téléchargement groupé – Tous les enseignants vacataires")
    st.markdown(
        "Un seul classeur Excel multifeuille contenant un index cliquable, "
        "un récapitulatif des charges Cours/TD/TP et une fiche officielle par enseignant."
    )
    if st.button("📦 Générer le classeur Excel de tous les enseignants",
                 type="primary", use_container_width=True):
        with st.spinner("Génération des fiches de tous les enseignants..."):
            buffer_groupe, nb_fiches = generer_excel_groupe(
                df_vacataires, df_edt, liste_vacataires_actifs, dict_noms_arabe,
                date_debut_cours, exclusions_par_mois
            )
        st.success(f"Classeur groupé généré : {nb_fiches} fiches individuelles.")
        st.download_button(
            label="📥 Télécharger toutes les fiches H.SUP (Excel multifeuille)",
            data=buffer_groupe,
            file_name="Etats_HSUP_Tous_Vacataires_ELT_2026_2027.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="download_hsup_groupe",
        )

except Exception as e:
    st.error(f"Une erreur est survenue lors du traitement : {e}")
    st.exception(e)
