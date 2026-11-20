# -*- coding: utf-8 -*-
"""
Gestion et génération des états H.SUP des enseignants permanents et vacataires - ELT
Année universitaire 2026-2027.
Le calcul est fondé sur les séances récurrentes de l'EDT et les dates réelles
de chaque mois (pas sur une proratisation du volume hebdomadaire).
"""

import io
import re
import base64
import json
import urllib.request
import urllib.error
import urllib.parse
import calendar
import copy
import unicodedata
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

st.set_page_config(
    page_title="Gestion H.SUP Permanents et Vacataires - 2026-2027",
    page_icon="⚡",
    layout="wide",
)

TITRE_OFFICIEL = "Plateforme de gestion des EDTs et H.SUP - Permanents et Vacataires - 2026-2027"
st.title(f"⚡ {TITRE_OFFICIEL}")
st.markdown("**Département d'Électrotechnique - Faculté de génie électrique - UDL-SBA**")
st.divider()

BASE_DIR = Path(__file__).resolve().parent
def choisir_fichier(nom_standard, variantes=()):
    """Retrouve le fichier sous son nom normal ou avec suffixe de téléchargement."""
    candidats = [BASE_DIR / nom_standard] + [BASE_DIR / nom for nom in variantes]
    return next((f for f in candidats if f.exists()), candidats[0])


FICHIER_PV = choisir_fichier(
    "Permanents-Vacataires-ELT2-2026-2027.xlsx",
    (
        "Permanents-Vacataires-ELT2-2026-2027(3).xlsx",
        "Permanents-Vacataires-ELT2-2026-2027(2).xlsx",
        "Permanents-Vacataires-ELT2-2026-2027(1).xlsx",
    ),
)
FICHIER_EDT = choisir_fichier(
    "dataEDT-ELT-S1-2027.xlsx",
    (
        "dataEDT-ELT-S1-2027(6).xlsx",
        "dataEDT-ELT-S1-2027(5).xlsx",
        "dataEDT-ELT-S1-2027(3).xlsx",
        "dataEDT-ELT-S1-2027(1).xlsx",
    ),
)
# Le fichier récent contient séparément « الاسم : » et « اللقب : ».
FICHIER_COORD = choisir_fichier(
    "Vos coordonnées en Arabe (réponses)(1).xlsx",
    ("Vos coordonnées en Arabe (réponses).xlsx",),
)
FICHIER_TEMPLATE = choisir_fichier(
    "Etat_HSUP_Enseignant_2026_2027.xlsx",
    ("Etat_HSUP_Vacataire_ABBAS_MOHAMED_2026_2027.xlsx",),
)

# Les blocs et lignes correspondent à la feuille « Nom & prénom » du modèle fourni.
# Mois affichés dans la fiche : le tableau est reconstruit dynamiquement (une ligne par jour travaillé).
MOIS_BLOCKS = [
    {"nom": "Septembre 2026", "annee": 2026, "mois": 9},
    {"nom": "Octobre 2026", "annee": 2026, "mois": 10},
    {"nom": "Novembre 2026", "annee": 2026, "mois": 11},
    {"nom": "Décembre 2026", "annee": 2026, "mois": 12},
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
# Indexé par date.weekday() : 0 = lundi ... 6 = dimanche
NOMS_JOURS_FR = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
# Noms inscrits par défaut sous les cases de signature de l'état H.SUP (modifiables dans la barre latérale)
SIGNATAIRES = {
    "naib_chef": "ناصور كمال",    # نائبي رئيس القسم
    "chef": "ميلوى فريد",         # رئيس القسم
}
NOMS_JOURS_AR = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت", "الأحد"]


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
    """Associe les deux champs arabes à l'identité française de chaque vacataire.

    Le formulaire récent fournit explicitement « الاسم : » et « اللقب : ».
    Ces colonnes sont prioritaires et sont conservées séparément, sans les
    inverser ni dépendre de l'ancienne colonne arabe regroupée.
    """
    if df is None or df.empty:
        return

    colonnes = list(df.columns)
    colonnes_norm = {normaliser_entete(c): c for c in colonnes}

    # Identité française de référence : colonne complète du formulaire ou
    # combinaison NOM + PRÉNOM du fichier du personnel.
    col_fr = next((c for c in colonnes if normaliser_entete(c) in (
        "nom et prenom en francais", "nom et prenom francais",
        "nom prenom francais", "nom complet francais"
    )), None)
    col_ism = next((c for c in colonnes if nettoyer_texte(c) == "الاسم :"), None)
    col_laqab = next((c for c in colonnes if nettoyer_texte(c) == "اللقب :"), None)

    # Accepter les variantes d'espacement/ponctuation dans les en-têtes arabes.
    if col_ism is None:
        col_ism = next((c for c in colonnes if "الاسم" in nettoyer_texte(c)), None)
    if col_laqab is None:
        col_laqab = next((c for c in colonnes if "اللقب" in nettoyer_texte(c)), None)

    # Compatibilité avec l'ancien formulaire : nom arabe en une seule cellule.
    col_ar_complet = next((c for c in colonnes if (
        "nom et prenom en arabe" in normaliser_entete(c)
        or ("arabe" in normaliser_entete(c) and "nom" in normaliser_entete(c))
    )), None)

    for _, ligne in df.iterrows():
        nom_fr = nettoyer_texte(ligne.get(col_fr, "")) if col_fr else construire_cle_francaise(ligne, colonnes)
        if not nom_fr:
            continue

        ism = nettoyer_texte(ligne.get(col_ism, "")) if col_ism else ""
        laqab = nettoyer_texte(ligne.get(col_laqab, "")) if col_laqab else ""

        # Ne recourir à l'ancienne colonne regroupée que si les deux nouvelles
        # colonnes ne sont pas disponibles ou si elles sont vides sur cette ligne.
        if (not ism and not laqab) and col_ar_complet:
            complet = nettoyer_texte(ligne.get(col_ar_complet, ""))
            parties = complet.split()
            if len(parties) >= 2:
                ism, laqab = parties[0], " ".join(parties[1:])
            elif parties:
                ism = parties[0]

        if ism or laqab:
            dict_arabe[normaliser_nom(nom_fr)] = {
                "ism": ism,
                "laqab": laqab,
                "complet": " ".join(x for x in (ism, laqab) if x),
            }


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


def _tokens_nom(nom):
    return [t for t in normaliser_nom(nom).split() if t]


def _tokens_canoniques(nom):
    """Jetons sans lettres doublées : MOHAMMED = MOHAMED, ABDELLAH = ABDELAH."""
    return [re.sub(r"(.)\1+", r"\1", t) for t in _tokens_nom(nom)]


def _similarite_jetons(a, b):
    from difflib import SequenceMatcher
    return SequenceMatcher(None, a, b).ratio()


def trouver_identite_arabe(dict_arabe, nom_fr):
    """
    Retrouve l'identité arabe (الاسم / اللقب) d'un enseignant à partir de son nom français,
    même si l'orthographe diffère légèrement entre l'EDT et le formulaire.

    Niveaux de correspondance, du plus sûr au plus souple :
      1. nom normalisé identique ;
      2. mêmes mots dans un autre ordre (« MOHAMED ABBAS » = « ABBAS MOHAMED ») ;
      3. mêmes mots sans lettres doublées (« MOHAMMED » = « MOHAMED ») ;
      4. mots très proches (similarité >= 0,85), seulement si un seul candidat convient ;
      5. nom incomplet d'un côté (« FILALI INES » = « FILALI INES NAFISSA »), si un seul
         candidat convient.
    Retourne (identité, mode, clé_trouvée) ; mode = "exacte", "ordre", "orthographe",
    "approchee", "partielle" ou "aucune".
    """
    if not dict_arabe or not nom_fr:
        return {}, "aucune", ""
    cle = normaliser_nom(nom_fr)
    if cle in dict_arabe:
        return dict_arabe[cle], "exacte", cle
    jetons = frozenset(_tokens_nom(nom_fr))
    for k, v in dict_arabe.items():
        if frozenset(k.split()) == jetons:
            return v, "ordre", k
    canon = frozenset(_tokens_canoniques(nom_fr))
    for k, v in dict_arabe.items():
        if frozenset(_tokens_canoniques(k)) == canon:
            return v, "orthographe", k
    candidats = []
    ja = sorted(canon)
    for k, v in dict_arabe.items():
        jb = sorted(frozenset(_tokens_canoniques(k)))
        if len(ja) != len(jb):
            continue
        pires = [max(_similarite_jetons(x, y) for y in jb) for x in ja]
        pires += [max(_similarite_jetons(y, x) for x in ja) for y in jb]
        if min(pires) >= 0.85:
            candidats.append((sum(pires) / len(pires), k, v))
    candidats.sort(key=lambda t: -t[0])
    if candidats and (len(candidats) == 1 or candidats[0][0] - candidats[1][0] >= 0.03):
        return candidats[0][2], "approchee", candidats[0][1]
    # 5. nom partiel : tous les mots du plus court figurent dans le plus long (>= 2 mots)
    partiels = []
    for k, v in dict_arabe.items():
        jk = frozenset(_tokens_canoniques(k))
        court, long_ = (canon, jk) if len(canon) <= len(jk) else (jk, canon)
        if len(court) >= 2 and court <= long_ and len(long_) - len(court) <= 2:
            partiels.append((k, v))
    if len(partiels) == 1:
        return partiels[0][1], "partielle", partiels[0][0]
    return {}, "aucune", ""


def identite_arabe_de(dict_arabe, nom_fr):
    return trouver_identite_arabe(dict_arabe, nom_fr)[0]


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
    df_enseignants = df_pv[
        df_pv["Qualité"].str.casefold().str.contains(r"permanent|vacataire", na=False)
    ].copy()
    df_enseignants["NOM_NORM"] = df_enseignants["NOM"].apply(normaliser_nom)

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

    # Le fichier de réponses arabe est la source prioritaire pour les champs
    # « الاسم : » et « اللقب : ». Le fichier du personnel et l'EDT servent à
    # retrouver le même enseignant à partir de son nom français complet.
    dict_arabe = {}
    ajouter_correspondances_arabes(df_pv, dict_arabe)
    if FICHIER_COORD.exists():
        df_coord = pd.read_excel(FICHIER_COORD, dtype=str)
        df_coord.columns = [nettoyer_texte(c) for c in df_coord.columns]
        # Les réponses les plus récentes prévalent en cas de doublon.
        col_horodateur = next((c for c in df_coord.columns if "horodateur" in normaliser_entete(c)), None)
        if col_horodateur:
            df_coord = df_coord.sort_values(col_horodateur, kind="stable")
        ajouter_correspondances_arabes(df_coord, dict_arabe)

        # Deuxième niveau de correspondance : l'adresse e-mail commune aux
        # réponses et au fichier du personnel. Il corrige les variantes
        # d'orthographe/ordre des noms français (ex. Mohammed/Mohamed).
        def normaliser_email(valeur):
            return re.sub(r"\s+", "", nettoyer_texte(valeur)).casefold()

        col_email_coord = next((c for c in df_coord.columns
                                if normaliser_entete(c) in ("adresse e mail", "email", "e mail")), None)
        col_email_pv = next((c for c in df_pv.columns
                             if normaliser_entete(c) in ("email", "e mail", "adresse e mail")), None)
        if col_email_coord and col_email_pv:
            arabe_par_email = {}
            for _, ligne in df_coord.iterrows():
                email = normaliser_email(ligne.get(col_email_coord, ""))
                if not email:
                    continue
                ism = nettoyer_texte(ligne.get("الاسم :", ""))
                laqab = nettoyer_texte(ligne.get("اللقب :", ""))
                if ism or laqab:
                    arabe_par_email[email] = {
                        "ism": ism, "laqab": laqab,
                        "complet": " ".join(x for x in (ism, laqab) if x),
                    }
            for _, ligne in df_pv.iterrows():
                email = normaliser_email(ligne.get(col_email_pv, ""))
                nom_pv = nettoyer_texte(ligne.get("NOM", ""))
                if email in arabe_par_email and nom_pv:
                    # Ne remplace pas une association directe déjà fiable.
                    dict_arabe.setdefault(normaliser_nom(nom_pv), arabe_par_email[email])

    actifs = []
    for _, enseignant in df_enseignants.iterrows():
        nom_n = enseignant["NOM_NORM"]
        if nom_n and df_edt["Enseignants"].apply(
            lambda x: correspond_enseignant(x, nom_n)
        ).any():
            actifs.append(enseignant["NOM"])
    return df_enseignants, df_edt, dict_arabe, sorted(set(actifs))


def charger_edt_complet():
    """Charge l'EDT complet, sans filtrer par département."""
    df_edt_complet = pd.read_excel(FICHIER_EDT, sheet_name=0, dtype=str)
    df_edt_complet.columns = [nettoyer_texte(c) for c in df_edt_complet.columns]

    obligatoires = ["Enseignements", "Enseignants", "Horaire", "Jours"]
    for colonne in obligatoires:
        if colonne not in df_edt_complet.columns:
            raise ValueError(
                "Colonne manquante dans l'EDT complet : " + colonne
            )

    for colonne in obligatoires:
        df_edt_complet[colonne] = df_edt_complet[colonne].fillna("").astype(str).str.strip()

    if "Département" in df_edt_complet.columns:
        df_edt_complet["Département"] = (
            df_edt_complet["Département"].fillna("").astype(str).str.strip()
        )
    else:
        df_edt_complet["Département"] = ""

    if "Promotion" in df_edt_complet.columns:
        df_edt_complet["Promotion"] = (
            df_edt_complet["Promotion"].fillna("").astype(str).str.strip()
        )
    else:
        df_edt_complet["Promotion"] = ""

    if "Lieu" in df_edt_complet.columns:
        df_edt_complet["Lieu"] = (
            df_edt_complet["Lieu"].fillna("").astype(str).str.strip()
        )
    else:
        df_edt_complet["Lieu"] = ""

    df_edt_complet["JOUR_NORM"] = df_edt_complet["Jours"].apply(
        lambda x: normaliser_nom(x).lower()
    )
    df_edt_complet["DUREE_H"] = df_edt_complet["Horaire"].apply(
        lambda x: duree_seance_heures(x) if x else 0.0
    )

    return df_edt_complet


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


def cle_horaire_enseignement(ligne):
    """Clé d'une séance simultanée : jour + horaire + type (promotion ignorée)."""
    horaire = nettoyer_texte(ligne.get("Horaire", ""))
    jour = nettoyer_texte(ligne.get("JOUR_NORM", ligne.get("Jours", ""))).casefold()
    nature = type_enseignement(ligne.get("Enseignements", ""))
    return jour, re.sub(r"\s+", "", horaire).casefold(), nature


def dedoublonner_seances_simultanees(seances):
    """Une seule séance par jour/horaire/type pour un même enseignant.

    Les promotions et les intitulés peuvent différer : si le même enseignant
    assure au même moment plusieurs Cours, TD ou TP pour plusieurs promotions,
    le créneau est compté une seule fois. Les types différents restent distincts.
    """
    if seances.empty:
        return seances.copy()
    travail = seances.copy()
    if "JOUR_NORM" not in travail.columns:
        travail["JOUR_NORM"] = travail.get("Jours", "").apply(normaliser_nom).str.lower()
    travail["_TYPE_HSUP"] = travail["Enseignements"].apply(type_enseignement)
    travail["_HORAIRE_HSUP"] = travail["Horaire"].apply(
        lambda x: re.sub(r"\s+", "", nettoyer_texte(x)).casefold()
    )
    # Les lignes sans type ou horaire interprétable ne sont pas fusionnées.
    travail["_CLE_SIMULTANEE"] = travail.apply(
        lambda r: (r.get("JOUR_NORM", ""), r.get("_HORAIRE_HSUP", ""), r.get("_TYPE_HSUP", ""))
        if r.get("_HORAIRE_HSUP") and r.get("_TYPE_HSUP") else ("LIGNE", int(r.name)), axis=1
    )
    travail = travail.drop_duplicates(subset=["_CLE_SIMULTANEE"], keep="first")
    return travail.drop(columns=["_TYPE_HSUP", "_HORAIRE_HSUP", "_CLE_SIMULTANEE"], errors="ignore")


SEPARATEUR_CRENEAU_HSUP = " ▸ "


def _norm_creneau_hsup(texte):
    return re.sub(r"\s+", "", nettoyer_texte(texte)).casefold()


def libelle_creneau_hsup(enseignement, jour, horaire):
    """Libellé d'un enseignement distingué par son créneau (jour + horaire)."""
    creneau = " ".join(x for x in (nettoyer_texte(jour), nettoyer_texte(horaire)) if x)
    ens = nettoyer_texte(enseignement)
    return f"{ens}{SEPARATEUR_CRENEAU_HSUP}{creneau}" if creneau else ens


def ajouter_libelles_creneaux_hsup(seances):
    """Retourne une copie des séances avec la colonne _LIBELLE_HSUP (enseignement + créneau)."""
    travail = seances.copy()
    if travail.empty:
        travail["_LIBELLE_HSUP"] = []
        return travail
    jours = travail["Jours"] if "Jours" in travail.columns else [""] * len(travail)
    travail["_LIBELLE_HSUP"] = [
        libelle_creneau_hsup(e, j, h)
        for e, j, h in zip(travail["Enseignements"], jours, travail["Horaire"])
    ]
    return travail


def options_enseignements_hsup(seances):
    """Liste déroulante H.SUP : un choix par enseignement ET par créneau distinct.

    Un même enseignement placé sur deux créneaux différents apparaît donc deux fois.
    Les promotions différentes sur un même créneau restent un seul choix.
    """
    if seances is None or seances.empty:
        return []
    libelles = ajouter_libelles_creneaux_hsup(seances)["_LIBELLE_HSUP"]
    vus, options = set(), []
    for lib in libelles:
        cle = _norm_creneau_hsup(lib)
        if lib and cle not in vus:
            vus.add(cle)
            options.append(lib)
    return sorted(options, key=lambda x: x.casefold())


def selections_vers_options_hsup(selections, options):
    """Convertit des sélections enregistrées vers les options affichées.

    Compatible avec les anciennes sélections (nom d'enseignement seul) : elles
    sélectionnent alors tous les créneaux de cet enseignement.
    """
    selections_norm = {_norm_creneau_hsup(x) for x in (selections or []) if nettoyer_texte(x)}
    resultat = []
    for opt in options:
        nom_seul = opt.split(SEPARATEUR_CRENEAU_HSUP)[0]
        if _norm_creneau_hsup(opt) in selections_norm or _norm_creneau_hsup(nom_seul) in selections_norm:
            resultat.append(opt)
    return resultat


def filtrer_enseignements_hsup(seances, qualite_enseignant, enseignements_retenus=None):
    """Pour les permanents, ne conserve que les enseignements/créneaux désignés comme H.SUP.

    Une sélection peut être un enseignement seul (tous ses créneaux) ou un
    enseignement + créneau précis (« Enseignement ▸ Jour Horaire »).
    """
    est_permanent = "permanent" in nettoyer_texte(qualite_enseignant).casefold()
    if not est_permanent:
        return seances.copy()
    retenus = {_norm_creneau_hsup(x) for x in (enseignements_retenus or []) if nettoyer_texte(x)}
    if not retenus or seances.empty:
        return seances.iloc[0:0].copy()
    avec_libelles = ajouter_libelles_creneaux_hsup(seances)
    masque = [
        _norm_creneau_hsup(lib) in retenus or _norm_creneau_hsup(ens) in retenus
        for lib, ens in zip(avec_libelles["_LIBELLE_HSUP"], avec_libelles["Enseignements"])
    ]
    return seances[pd.Series(masque, index=seances.index)].copy()


def calculer_charge_equivalente_hebdomadaire(seances):
    """Calcule la charge hebdomadaire en équivalent-cours après déduplication.

    1 h Cours = 1 h équivalent-cours ; 1,5 h TD ou TP = 1 h équivalent-cours.
    Les promotions différentes au même jour/créneau/type ne sont comptées qu'une fois.
    """
    uniques = dedoublonner_seances_simultanees(seances)
    if uniques.empty:
        return {"cours": 0.0, "td": 0.0, "tp": 0.0, "total_eq": 0.0,
                "heures_physiques": 0.0, "seances_uniques": 0}
    types = uniques["Enseignements"].apply(type_enseignement)
    heures = {}
    for t in ("cours", "td", "tp"):
        heures[t] = float(pd.to_numeric(uniques.loc[types.eq(t), "DUREE_H"], errors="coerce").fillna(0).sum())
    total_eq = heures["cours"] + heures["td"] / 1.5 + heures["tp"] / 1.5
    return {"cours": round(heures["cours"], 2), "td": round(heures["td"], 2),
            "tp": round(heures["tp"], 2), "total_eq": round(total_eq, 2),
            "heures_physiques": round(float(pd.to_numeric(uniques["DUREE_H"], errors="coerce").fillna(0).sum()), 2),
            "seances_uniques": int(len(uniques))}


def limiter_enseignements_au_depassement(seances_toutes, enseignements_retenus, seuil=7.0):
    """Retient les séances choisies dans la limite du dépassement hebdomadaire.

    L'allocation suit l'ordre chronologique des séances dans la semaine. Si le
    dépassement restant ne couvre qu'une partie de la dernière séance, seule
    cette durée équivalente résiduelle est comptabilisée.
    """
    charge = calculer_charge_equivalente_hebdomadaire(seances_toutes)
    depassement = max(0.0, charge["total_eq"] - float(seuil))
    retenus = filtrer_enseignements_hsup(seances_toutes, "permanent", enseignements_retenus)
    retenus = dedoublonner_seances_simultanees(retenus)
    if retenus.empty or depassement <= 0:
        return retenus.iloc[0:0].copy(), charge, 0.0

    travail = retenus.copy()
    travail["_JOUR_ORDRE"] = travail.get("JOUR_NORM", travail.get("Jours", "")).apply(
        lambda x: JOURS_FR.get(normaliser_nom(x).lower(), 7)
    )
    def _debut(row):
        try:
            return lire_heure(re.split(r"\s*-\s*", nettoyer_texte(row.get("Horaire", "")).replace("–", "-").replace("—", "-"))[0])
        except Exception:
            return 9999
    travail["_DEBUT_ORDRE"] = travail.apply(_debut, axis=1)
    travail = travail.sort_values(["_JOUR_ORDRE", "_DEBUT_ORDRE"], kind="stable")

    restant = depassement
    lignes = []
    total_retenu_eq = 0.0
    for _, row in travail.iterrows():
        typ = type_enseignement(row.get("Enseignements", ""))
        facteur = 1.0 if typ == "cours" else (1.5 if typ in ("td", "tp") else None)
        if facteur is None or restant <= 1e-9:
            continue
        duree = max(0.0, float(row.get("DUREE_H", 0) or 0))
        duree_autorisee = min(duree, restant * facteur)
        if duree_autorisee <= 1e-9:
            continue
        copie = row.copy()
        copie["DUREE_H"] = round(duree_autorisee, 6)
        copie["DUREE_H_INITIALE"] = duree
        copie["HS_EQ_RETENUE"] = round(duree_autorisee / facteur, 6)
        lignes.append(copie)
        eq = duree_autorisee / facteur
        restant -= eq
        total_retenu_eq += eq
    if not lignes:
        return retenus.iloc[0:0].copy(), charge, 0.0
    resultat = pd.DataFrame(lignes).drop(columns=["_JOUR_ORDRE", "_DEBUT_ORDRE"], errors="ignore")
    return resultat, charge, round(total_retenu_eq, 2)


def calculer_heures_par_jour(seances, annee, mois, date_debut, date_fin, exclusions):
    """
    Charge horaire de l'enseignant JOUR PAR JOUR pour un mois donné.

    Pour chaque date réelle du mois on additionne les séances de l'EDT qui
    tombent ce jour de la semaine. Sont retirés : vendredis et samedis, jours
    fériés connus, dates hors de la période comptabilisée et exclusions
    manuelles. Seules les dates où l'enseignant a cours sont retournées.
    """
    resultats = []
    jours_invalides = {4, 5}  # vendredi et samedi en Algérie
    for d in dates_du_mois(annee, mois):
        if d < date_debut or d > date_fin or d.weekday() in jours_invalides:
            continue
        if d in JOURS_FERIES or d in exclusions:
            continue
        lignes = seances[seances["JOUR_NORM"].eq(NOMS_JOURS_FR[d.weekday()])]
        if lignes.empty:
            continue
        # Plusieurs promotions au même jour/créneau/type pour le même enseignant
        # représentent une seule séance (Cours avec Cours, TD avec TD, TP avec TP).
        lignes = dedoublonner_seances_simultanees(lignes)
        resultats.append({
            "date": d,
            "jour_fr": NOMS_JOURS_FR[d.weekday()].capitalize(),
            "jour_ar": NOMS_JOURS_AR[d.weekday()],
            "heures": round(float(lignes["DUREE_H"].sum()), 2),
            "nb_seances": int(len(lignes)),
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


def remplir_fiche(ws, nom_enseignant, seances, dict_arabe, date_debut,
                  date_fin, exclusions_par_mois, qualite_enseignant=""):
    """Remplit une feuille du modèle officiel pour un seul enseignant."""
    nom_norm = normaliser_nom(nom_enseignant)
    identite_arabe = identite_arabe_de(dict_arabe, nom_enseignant)
    ism_ar = nettoyer_texte(identite_arabe.get("ism", "")) if isinstance(identite_arabe, dict) else ""
    laqab_ar = nettoyer_texte(identite_arabe.get("laqab", "")) if isinstance(identite_arabe, dict) else ""

    # Respecter exactement les deux colonnes du formulaire, sans découpage
    # automatique qui pourrait inverser ou tronquer les noms composés.
    ws["A7"] = f"الاسم : {ism_ar}" if ism_ar else "الاسم :"
    ws["C7"] = f"اللقب : {laqab_ar}" if laqab_ar else "اللقب :"
    ws["A8"] = "الرتبة :" + (
        "أستاذ دائم" if "permanent" in nettoyer_texte(qualite_enseignant).casefold()
        else "أستاذ مؤقت"
    )
    ws["C8"] = "الهيئة المستخدمة الأصلية : الالكتروتقني"

    matieres = sorted({
        nettoyer_texte(x) for x in seances["Enseignements"].tolist()
        if nettoyer_texte(x)
    })
    ws["A9"] = "المقياس المدرس : " + " // ".join(matieres)
    # نوع التدريس : afficher uniquement les catégories réellement enseignées.
    # L'ordre est toujours Cours + TD + TP et aucun nom (français ou arabe)
    # ne doit être ajouté dans ce champ.
    types_presents = {
        type_enseignement(x)
        for x in seances["Enseignements"].tolist()
        if type_enseignement(x)
    }
    types_ordre = [
        libelle for libelle, cle in (
            ("Cours", "cours"),
            ("TD", "td"),
            ("TP", "tp"),
        )
        if cle in types_presents
    ]
    ws["D9"] = "نوع التدريس : " + "+".join(types_ordre)

    # ------------------------------------------------------------------
    # Tableau « اليوم و التاريخ / عدد الساعات » : UNE LIGNE PAR JOUR travaillé
    # (jour + date réelle), regroupée par mois, puis total du mois.
    # Les styles du modèle officiel sont capturés AVANT toute modification.
    # ------------------------------------------------------------------
    def _styles(ligne):
        return [copy.copy(ws.cell(ligne, c)._style) for c in (1, 2, 3)]

    st_mois, st_jour = _styles(11), _styles(12)
    st_total, st_gtotal = _styles(17), _styles(32)
    st_vide, st_dernier = _styles(33), _styles(39)
    st_d_vide = copy.copy(ws["D12"]._style)
    st_d_vide_fin, st_d_dernier = copy.copy(ws["D33"]._style), copy.copy(ws["D39"]._style)
    # (style, texte) des cases de signature de la colonne D du modèle
    cases_d = [
        (copy.copy(ws["D11"]._style), ws["D11"].value),
        (copy.copy(ws["D16"]._style), ws["D16"].value),
        (copy.copy(ws["D21"]._style), ws["D21"].value),
        (copy.copy(ws["D28"]._style), ws["D28"].value),
    ]
    style_dir, texte_dir = copy.copy(ws["D40"]._style), ws["D40"].value
    # style de la ligne « nom » placée sous chaque barre grise (D17 sous نائبي رئيس القسم, D22 sous رئيس القسم)
    styles_noms_d = [copy.copy(ws["D17"]._style), copy.copy(ws["D22"]._style)]

    # Effacer l'ancien contenu du modèle (lignes 11 à 60), colonnes A à F.
    for ligne in range(11, 61):
        ws.row_dimensions[ligne].height = None
        for colonne in range(1, 7):
            cell = ws.cell(ligne, colonne)
            cell.value = None
            cell.style = "Normal"

    mois_arabe = {9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر"}
    total_general = 0.0
    details = []
    heures_par_mois = {}
    ligne = 11
    premiere_ligne = ligne

    def _ecrire(ligne_ecr, styles, a=None, b=None, hauteur=None):
        for col, style in zip((1, 2, 3), styles):
            ws.cell(ligne_ecr, col)._style = copy.copy(style)
        ws.cell(ligne_ecr, 1).value = a
        ws.cell(ligne_ecr, 2).value = b
        if hauteur:
            ws.row_dimensions[ligne_ecr].height = hauteur

    for bloc in MOIS_BLOCKS:
        annee, mois = bloc["annee"], bloc["mois"]
        exclusions = exclusions_par_mois.get(mois, set())
        _ecrire(ligne, st_mois, f"شهر: {mois_arabe.get(mois, bloc['nom'])} {annee}", None, 18)
        ligne += 1
        jours = calculer_heures_par_jour(seances, annee, mois, date_debut, date_fin, exclusions)
        total_mois = 0.0
        if not jours:
            _ecrire(ligne, st_jour, "/", "/", 18.75)
            ligne += 1
        for jour in jours:
            _ecrire(ligne, st_jour,
                    f"{jour['jour_ar']} {jour['date'].strftime('%d/%m/%Y')}",
                    format_heures_arabe(jour["heures"]) or "/", 18.75)
            ligne += 1
            total_mois += jour["heures"]
            details.append({
                "Mois": bloc["nom"],
                "Jour": jour["jour_fr"],
                "Date": jour["date"].strftime("%d/%m/%Y"),
                "Séances": jour["nb_seances"],
                "Heures calculées": jour["heures"],
            })
        _ecrire(ligne, st_total, "            المجموع :", format_heures_arabe(total_mois) or "/", 21)
        ligne += 1
        total_general += total_mois
        heures_par_mois[bloc["nom"]] = round(total_mois, 2)

    # Total général
    _ecrire(ligne, st_gtotal, "المجموع الكلي    ", format_heures_arabe(total_general) or "/", 21.75)
    ligne_total_general = ligne
    ligne += 1
    # Lignes de fermeture du cadre (comme dans le modèle, sous le total général)
    for i in range(6):
        dernier = i == 5
        _ecrire(ligne, st_dernier if dernier else st_vide, None, None, 18)
        ws.cell(ligne, 4)._style = copy.copy(st_d_dernier if dernier else st_d_vide_fin)
        ligne += 1
    derniere_ligne_cadre = ligne - 1

    # Colonne D (signatures) : cases du modèle réparties sur toute la hauteur du tableau
    hauteur_tableau = ligne_total_general - premiere_ligne + 1
    positions = [premiere_ligne] + [
        premiere_ligne + max(1, int(round(hauteur_tableau * f))) for f in (0.23, 0.45, 0.77)
    ]
    for r in range(premiere_ligne, ligne_total_general + 1):
        ws.cell(r, 4)._style = copy.copy(st_d_vide)
    for pos, (style_case, texte_case) in zip(positions, cases_d):
        ws.cell(pos, 4)._style = copy.copy(style_case)
        ws.cell(pos, 4).value = texte_case
    # Noms par défaut sous « نائبي رئيس القسم » (barre n° 2) et « رئيس القسم » (barre n° 3)
    for pos, cle_nom, style_nom in ((positions[1], "naib_chef", styles_noms_d[0]),
                                    (positions[2], "chef", styles_noms_d[1])):
        nom_signataire = nettoyer_texte(SIGNATAIRES.get(cle_nom, ""))
        if nom_signataire and pos + 1 <= ligne_total_general:
            ws.cell(pos + 1, 4)._style = copy.copy(style_nom)
            ws.cell(pos + 1, 4).value = nom_signataire

    # Directeur de l'université sous le cadre, puis note de calcul
    ws.cell(derniere_ligne_cadre + 1, 4)._style = copy.copy(style_dir)
    ws.cell(derniere_ligne_cadre + 1, 4).value = texte_dir
    ws.row_dimensions[derniere_ligne_cadre + 1].height = 18.75
    ligne_note = derniere_ligne_cadre + 3
    ws.cell(ligne_note, 1).value = (
        f"ملاحظة : تم احتساب الساعات من {date_debut.strftime('%d/%m/%Y')} "
        f"إلى {date_fin.strftime('%d/%m/%Y')} حسب الحصص الفعلية في جدول التوقيت لكل يوم وتاريخ، "
        "مع استثناء أيام العطل والتواريخ المحددة للاستبعاد."
    )
    ws.merge_cells(start_row=ligne_note, start_column=1, end_row=ligne_note, end_column=4)
    ws.cell(ligne_note, 1).alignment = openpyxl.styles.Alignment(
        horizontal="right", vertical="top", wrap_text=True, readingOrder=2
    )
    ws.cell(ligne_note, 1).font = openpyxl.styles.Font(name="Arial", size=10)
    ws.row_dimensions[ligne_note].height = 30

    # Impression : une page de large, en-tête du tableau répété sur chaque page
    ws.print_area = f"A1:D{ligne_note}"
    ws.print_title_rows = "10:10"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    return total_general, details, heures_par_mois


def generer_excel(nom_enseignant, seances, dict_arabe, date_debut, date_fin,
                  exclusions_par_mois, qualite_enseignant=""):
    """Génère la fiche individuelle selon le modèle officiel."""
    wb = openpyxl.load_workbook(FICHIER_TEMPLATE)
    nom_feuille = "Nom & prénom" if "Nom & prénom" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[nom_feuille]
    # Conserver le sens droite-à-gauche du modèle officiel.
    ws.sheet_view.rightToLeft = True
    for autre_feuille in list(wb.sheetnames):
        if autre_feuille != nom_feuille:
            del wb[autre_feuille]

    total_general, details, _ = remplir_fiche(
        ws, nom_enseignant, seances, dict_arabe, date_debut, date_fin,
        exclusions_par_mois, qualite_enseignant
    )
    ws.title = re.sub(r"[\\\\/*?:\\[\\]]", " ", nom_enseignant).strip()[:31] or "Etat H.SUP"

    wd = wb.create_sheet("Détail calcul")
    wd.append(["Mois", "Jour", "Date", "Séances", "Heures calculées"])
    for ligne in details:
        wd.append(list(ligne.values()))
    wd.freeze_panes = "A2"
    wd.auto_filter.ref = wd.dimensions
    for col, largeur in {"A": 22, "B": 14, "C": 14, "D": 12, "E": 20}.items():
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


def generer_excel_groupe(df_enseignants, df_edt, liste_enseignants, dict_arabe,
                         date_debut, date_fin, exclusions_par_mois,
                         selections_hsup_par_enseignant=None,
                         exclusions_par_mois_par_enseignant=None,
                         df_edt_source=None):
    """
    Crée un seul classeur : index cliquable, récapitulatif, puis une fiche
    officielle par enseignant (une feuille par enseignant).
    """
    # Source EDT paramétrable : EDT complet pour la génération admin.
    edt_source = df_edt_source if df_edt_source is not None else df_edt

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
    ws_index.append(["Période de comptabilisation", f"Du {date_debut.strftime('%d/%m/%Y')} au {date_fin.strftime('%d/%m/%Y')}"])
    ws_index.append(["Enseignant", "Nom en arabe", "Total H.SUP", "Accès direct"])
    ws_recap.append(["RÉCAPITULATIF DES CHARGES D'ENSEIGNEMENT – ELT"])
    recap_headers = [
        "Enseignant", "Nom et prénom en arabe",
        "Nombre Cours", "Nombre TD", "Nombre TP",
        "Heures Cours retenues H.SUP / semaine", "Heures TD retenues H.SUP / semaine",
        "Heures TP retenues H.SUP / semaine", "Heures H.SUP physiques / semaine",
        "Charge totale équivalent-cours / semaine", "Dépassement au-dessus de 7 h éq-cours",
        "H.SUP retenues équivalent-cours / semaine", "Total H.SUP période",
    ]
    mois_cols = [bloc["nom"] for bloc in MOIS_BLOCKS]
    ws_recap.append(recap_headers + mois_cols)

    deja_utilises = {"index enseignants", "récapitulatif"}
    # La feuille modèle devient la première fiche ; les autres sont des copies exactes.
    enseignants = []
    for nom in liste_enseignants:
        nom_norm = normaliser_nom(nom)
        seances = edt_source[
            edt_source["Enseignants"].apply(lambda x: correspond_enseignant(x, nom_norm))
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

        # openpyxl ne recopie pas tous les paramètres de présentation d'une
        # feuille. En particulier, le sens droite-à-gauche du modèle officiel
        # peut être perdu lors de copy_worksheet(). On recopie donc explicitement
        # les paramètres d'affichage et de mise en page du modèle individuel.
        # sheet_view est une propriété sans setter : recopier ses attributs utiles.
        copie.sheet_view.rightToLeft = modele.sheet_view.rightToLeft
        copie.sheet_view.showGridLines = modele.sheet_view.showGridLines
        copie.sheet_view.zoomScale = modele.sheet_view.zoomScale
        copie.sheet_properties = copy.copy(modele.sheet_properties)
        copie.page_margins = copy.copy(modele.page_margins)
        copie.page_setup = copy.copy(modele.page_setup)
        copie.print_options = copy.copy(modele.print_options)
        copie.sheet_format = copy.copy(modele.sheet_format)
        copie.freeze_panes = modele.freeze_panes
        copie.sheet_view.rightToLeft = True
        copie.title = nom_onglet
        fiches.append((nom, seances, copie))

    # Garantir le même rendu RTL sur la première fiche et toutes les copies.
    modele.sheet_view.rightToLeft = True
    for _, _, feuille_fiche in fiches:
        feuille_fiche.sheet_view.rightToLeft = True

    # Les feuilles d'index et de récap sont créées avant les fiches : déplacer
    # explicitement l'ordre pour garantir Index, Récapitulatif, puis fiches.
    for nom, seances, ws_fiche in fiches:
        ligne_personnel = df_enseignants[df_enseignants["NOM_NORM"].eq(normaliser_nom(nom))]
        qualite_fiche = (
            nettoyer_texte(ligne_personnel.iloc[0].get("Qualité", ""))
            if not ligne_personnel.empty else ""
        )
        charge_complete = calculer_charge_equivalente_hebdomadaire(seances)
        hsup_eq_retenue = None
        if "permanent" in qualite_fiche.casefold():
            choix_hsup = (selections_hsup_par_enseignant or {}).get(normaliser_nom(nom), [])

            # Repli administratif : si aucun enseignement n'est retenu, utiliser
            # tous les enseignements EDT, dans la limite des 7 h équivalent-cours.
            if not choix_hsup:
                choix_hsup = sorted(
                    {
                        nettoyer_texte(x)
                        for x in seances["Enseignements"]
                        if nettoyer_texte(x)
                    }
                )

            seances, _, hsup_eq_retenue = limiter_enseignements_au_depassement(
                seances, choix_hsup, seuil=7.0
            )
        # Les vacataires gardent le calcul de leurs séances réelles.
        exclusions_fiche = exclusions_par_mois
        if exclusions_par_mois_par_enseignant:
            exclusions_fiche = exclusions_par_mois_par_enseignant.get(
                normaliser_nom(nom), exclusions_par_mois
            )
        total, details, heures_mois = remplir_fiche(
            ws_fiche, nom, seances, dict_arabe, date_debut, date_fin,
            exclusions_fiche, qualite_fiche
        )
        identite_arabe = identite_arabe_de(dict_arabe, nom)
        nom_arabe = (identite_arabe.get("complet", "")
                     if isinstance(identite_arabe, dict) else nettoyer_texte(identite_arabe))
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
            round(charge_complete["total_eq"], 2),
            round(max(0.0, charge_complete["total_eq"] - 7.0), 2) if "permanent" in qualite_fiche.casefold() else 0.0,
            round(float(hsup_eq_retenue or 0.0), 2) if "permanent" in qualite_fiche.casefold() else 0.0,
            round(total, 2),
        ]
        ligne_recap.extend(round(heures_mois.get(cle, 0.0), 2)
                           for cle in mois_cols)
        ws_recap.append(ligne_recap)

    # Mise en forme des deux feuilles de synthèse.
    ws_index.freeze_panes = "A4"
    ws_index.auto_filter.ref = f"A3:{openpyxl.utils.get_column_letter(ws_index.max_column)}{ws_index.max_row}"
    ws_recap.freeze_panes = "A3"
    ws_recap.auto_filter.ref = f"A2:{openpyxl.utils.get_column_letter(ws_recap.max_column)}{ws_recap.max_row}"
    for ws in (ws_index, ws_recap):
        ws.sheet_view.showGridLines = False
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True, size=14, color="FFFFFF")
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="17365D")
        entete_ligne = 3 if ws is ws_index else 2
        for cell in ws[entete_ligne]:
            cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
            cell.fill = openpyxl.styles.PatternFill("solid", fgColor="4472C4")
        for row in ws.iter_rows(min_row=entete_ligne + 1):
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


# ---------------------------------------------------------------------------
# Emploi du temps individuel en grille (PDF) - mêmes dimensions que le modèle
# « EMPLOI DU TEMPS » (PPER.03) : A4 paysage, marges 28,35 pt (1 cm).
# Toutes les cotes ci-dessous sont mesurées sur le fichier PDF de référence,
# en points, origine en haut à gauche de la page.
# ---------------------------------------------------------------------------
EDT_PAGE_W, EDT_PAGE_H = 841.89, 595.28
EDT_CODE, EDT_REVISION, EDT_DATE_DOC = "PPER.03", "00", "16/05/2026"
EDT_SEMESTRE = "Semestre 01 - departement d'Electrotechnique - FGE/UDL-SBA"
EDT_SANS_ACCENTS = True      # le modèle officiel est écrit sans accents
EDT_BLEU = (0.1176, 0.2275, 0.5412)
EDT_GRIS_TEXTE = (0.3922, 0.3922, 0.3922)
EDT_GRIS_BORD = (0.7059, 0.7059, 0.7059)
EDT_COULEURS_TYPE = {            # fond des cases selon le type de séance
    "cours": (0.8824, 0.9333, 1.0),
    "td": (0.9098, 0.9882, 0.9216),
    "tp": (1.0, 0.9216, 0.9216),
}
# Cadre d'en-tête (x0, haut, x1, bas) et séparations
EDT_ENTETE = (194.5, 28.35, 647.38, 103.23)
EDT_SEP_X = (280.19, 546.59)
EDT_SEP_Y = 71.55
EDT_LOGO = (200.2, 34.0, 274.5, 97.6)
# Grille
EDT_TAB_X0, EDT_TAB_X1 = 28.35, 813.54
EDT_COL_JOUR = 62.36
EDT_TAB_Y0, EDT_HAUT_ENTETE, EDT_HAUT_LIGNE = 174.09, 22.68, 49.89
EDT_TAILLE_CASE, EDT_INTERLIGNE, EDT_MARGE_CASE = 5.5, 9.075, 8.5
EDT_JOURS_ORDRE = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"]


# Logo de l'Université Djillali Liabès (extrait du modèle PDF officiel). Il sert de secours si
# logo.PNG est introuvable à côté de app.py : le PDF contient alors toujours le logo.
LOGO_UDL_BASE64 = (
    "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"
)


def trouver_logo():
    """Retrouve logo.PNG (la casse du nom varie entre Windows et GitHub)."""
    try:
        return next((p for p in BASE_DIR.iterdir() if p.name.lower() == "logo.png"), None)
    except OSError:
        return None


def _edt_texte(valeur):
    texte = nettoyer_texte(valeur).replace("’", "'").replace("–", "-").replace("—", "-")
    if EDT_SANS_ACCENTS:
        texte = "".join(c for c in unicodedata.normalize("NFKD", texte) if not unicodedata.combining(c))
    return texte.encode("latin-1", "replace").decode("latin-1")


def _edt_libelle_heure(minutes):
    h, m = divmod(int(minutes), 60)
    return f"{h}h{m:02d}" if m else f"{h}h"


def _edt_creneau(horaire):
    """'9h30 -11h' -> ((570, 660), '9h30 - 11h')."""
    morceaux = re.split(r"\s*-\s*", nettoyer_texte(horaire).replace("–", "-").replace("—", "-"))
    debut, fin = lire_heure(morceaux[0]), lire_heure(morceaux[1])
    return (debut, fin), f"{_edt_libelle_heure(debut)} - {_edt_libelle_heure(fin)}"


def _edt_lignes_case(seances_case, largeur_utile):
    """Lignes de texte d'une case : département, enseignement, lieu et promotion(s)."""
    from reportlab.lib.utils import simpleSplit
    groupes = {}
    for _, seance in seances_case:
        departement = _edt_texte(seance.get("Département", "")) or "Non précisé"
        cle = (
            departement,
            _edt_texte(seance.get("Enseignements", "")),
            _edt_texte(seance.get("Lieu", "")),
        )
        promo = _edt_texte(seance.get("Promotion", ""))
        groupes.setdefault(cle, [])
        if promo and promo not in groupes[cle]:
            groupes[cle].append(promo)

    lignes = []
    for (departement, intitule, lieu), promos in groupes.items():
        lignes += simpleSplit(
            f"Département : {departement}",
            "Helvetica",
            EDT_TAILLE_CASE,
            largeur_utile,
        ) or [""]
        lignes += simpleSplit(
            intitule, "Helvetica", EDT_TAILLE_CASE, largeur_utile
        ) or [""]
        if lieu:
            lignes.append(lieu)
        if promos:
            lignes += simpleSplit(
                ", ".join(promos), "Helvetica", EDT_TAILLE_CASE, largeur_utile
            )
    return lignes


def _edt_dessiner_page(c, nom, qualite, seances, logo, page=1, pages=1):
    H = EDT_PAGE_H

    def y(haut):                      # ordonnée PDF depuis une cote « depuis le haut »
        return H - haut

    def texte(x, haut, chaine, police, taille, couleur=(0, 0, 0), centre=False):
        c.setFillColorRGB(*couleur)
        c.setFont(police, taille)
        base = y(haut + 0.793 * taille)   # 'haut' = sommet du texte, comme dans le modèle
        (c.drawCentredString if centre else c.drawString)(x, base, chaine)

    # ---- cadre d'en-tête ----
    x0, h0, x1, h1 = EDT_ENTETE
    c.setStrokeColorRGB(0, 0, 0)
    c.setLineWidth(0.85)
    c.rect(x0, y(h1), x1 - x0, h1 - h0, stroke=1, fill=0)
    for sx in EDT_SEP_X:
        c.line(sx, y(h0), sx, y(h1))
    c.line(EDT_SEP_X[0], y(EDT_SEP_Y), EDT_SEP_X[1], y(EDT_SEP_Y))
    if logo is not None:
        lx0, lh0, lx1, lh1 = EDT_LOGO
        c.drawImage(logo, lx0, y(lh1), width=lx1 - lx0, height=lh1 - lh0,
                    preserveAspectRatio=True, anchor="c", mask="auto")
    centre_x = (EDT_SEP_X[0] + EDT_SEP_X[1]) / 2
    texte(centre_x, 35.0, "Universite Djillali Liabes", "Helvetica-Bold", 11, centre=True)
    texte(centre_x, 50.4, "Sidi Bel Abbes", "Helvetica", 10, centre=True)
    texte(centre_x, 81.5, "EMPLOI DU TEMPS", "Helvetica-Bold", 12, centre=True)
    for haut, chaine in ((34.7, f"Code : {EDT_CODE}"), (53.4, f"Revision : {EDT_REVISION}"),
                         (72.1, f"Date : {EDT_DATE_DOC}"), (90.8, f"Pages : {page}/{pages}")):
        texte(553.7, haut, chaine, "Helvetica", 9)

    # ---- titres ----
    milieu = EDT_PAGE_W / 2
    texte(milieu, 123.3, f"EMPLOI DU TEMPS INDIVIDUEL - {_edt_texte(nom).upper()}",
          "Helvetica-Bold", 11, EDT_BLEU, centre=True)
    texte(milieu, 143.2, EDT_SEMESTRE, "Helvetica-Oblique", 8, EDT_GRIS_TEXTE, centre=True)
    texte(milieu, 157.4, f"Qualite : {_edt_texte(qualite) or 'Vacataire'}", "Helvetica-Bold", 8,
          EDT_GRIS_TEXTE, centre=True)

    # ---- construction de la grille : colonnes = créneaux utilisés, lignes = jours ----
    creneaux = {}
    cases = {}
    for _, s in seances.iterrows():
        try:
            borne, libelle = _edt_creneau(s["Horaire"])
        except (ValueError, IndexError):
            continue
        creneaux[borne] = libelle
        jour = normaliser_nom(s["Jours"]).lower()
        cases.setdefault((jour, borne), []).append((s.name, s))
    colonnes = sorted(creneaux)
    jours = [j for j in EDT_JOURS_ORDRE if any(k[0] == j for k in cases)]
    jours += sorted({k[0] for k in cases} - set(jours))
    nb_col = max(1, len(colonnes))
    larg_col = (EDT_TAB_X1 - EDT_TAB_X0 - EDT_COL_JOUR) / nb_col
    xs = [EDT_TAB_X0, EDT_TAB_X0 + EDT_COL_JOUR] + [
        EDT_TAB_X0 + EDT_COL_JOUR + larg_col * (i + 1) for i in range(nb_col)
    ]
    xs[-1] = EDT_TAB_X1

    # en-tête du tableau
    c.setLineWidth(0.57)
    c.setStrokeColorRGB(0, 0, 0)
    c.setFillColorRGB(*EDT_BLEU)
    for i in range(len(xs) - 1):
        c.rect(xs[i], y(EDT_TAB_Y0 + EDT_HAUT_ENTETE), xs[i + 1] - xs[i], EDT_HAUT_ENTETE, stroke=1, fill=1)
    centre_entete = EDT_TAB_Y0 + EDT_HAUT_ENTETE / 2
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 7)
    c.drawCentredString((xs[0] + xs[1]) / 2, y(centre_entete) - 0.3 * 7, "JOUR")
    for i, borne in enumerate(colonnes):
        c.drawCentredString((xs[i + 1] + xs[i + 2]) / 2, y(centre_entete) - 0.3 * 7, creneaux[borne])

    # lignes de jours
    haut = EDT_TAB_Y0 + EDT_HAUT_ENTETE
    c.setStrokeColorRGB(*EDT_GRIS_BORD)
    for n_jour, jour in enumerate(jours):
        contenus = {}
        hauteur = EDT_HAUT_LIGNE
        for i, borne in enumerate(colonnes):
            lignes = _edt_lignes_case(cases.get((jour, borne), []), larg_col - 2 * EDT_MARGE_CASE)
            contenus[i] = lignes
            hauteur = max(hauteur, len(lignes) * EDT_INTERLIGNE + 12)
        # case « jour »
        c.setFillColorRGB(*((0.9725, 0.9804, 0.9882) if n_jour % 2 == 0 else (1, 1, 1)))
        c.rect(xs[0], y(haut + hauteur), xs[1] - xs[0], hauteur, stroke=1, fill=1)
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString((xs[0] + xs[1]) / 2, y(haut + hauteur / 2) - 0.3 * 7, "?" if not jour else jour.capitalize())
        centre = haut + hauteur / 2
        for i, borne in enumerate(colonnes):
            sc = cases.get((jour, borne), [])
            couleur = (1, 1, 1)
            if sc:
                couleur = EDT_COULEURS_TYPE.get(type_enseignement(sc[0][1].get("Enseignements", "")), (1, 1, 1))
            c.setFillColorRGB(*couleur)
            c.rect(xs[i + 1], y(haut + hauteur), xs[i + 2] - xs[i + 1], hauteur, stroke=1, fill=1)
            lignes = contenus[i]
            c.setFillColorRGB(0, 0, 0)
            c.setFont("Helvetica", EDT_TAILLE_CASE)
            for k, ligne in enumerate(lignes):
                base = centre + (k - (len(lignes) - 1) / 2) * EDT_INTERLIGNE + 0.3 * EDT_TAILLE_CASE
                c.drawString(xs[i + 1] + EDT_MARGE_CASE, y(base), ligne)
        haut += hauteur

    # pied de page
    texte(793.2, 572.0, f"{page}/{pages}", "Helvetica-Oblique", 7, (0.502, 0.502, 0.502))


def generer_pdf_edt(liste_edt):
    """liste_edt : [(nom, qualité, DataFrame des séances)] -> PDF (une page par enseignant)."""
    from reportlab.pdfgen import canvas
    from reportlab.lib.utils import ImageReader
    logo_path = trouver_logo()
    # logo.PNG s'il existe, sinon le logo intégré au code (jamais de PDF sans logo)
    logo = ImageReader(str(logo_path)) if logo_path else ImageReader(
        io.BytesIO(base64.b64decode(LOGO_UDL_BASE64)))
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(EDT_PAGE_W, EDT_PAGE_H))
    for nom, qualite, seances in liste_edt:
        c.setTitle(f"EDT {nom}")
        _edt_dessiner_page(c, nom, qualite, seances, logo)
        c.showPage()
    c.save()
    buffer.seek(0)
    return buffer


def nom_fichier_edt(nom):
    return f"EDT_{normaliser_nom(nom).replace(' ', '_')}_2027.pdf"


def generer_zip_edt(liste_edt):
    """Un PDF par enseignant, nommés comme le modèle (EDT_NOM_PRENOM_2027.pdf)."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
        for element in liste_edt:
            z.writestr(nom_fichier_edt(element[0]), generer_pdf_edt([element]).getvalue())
    buffer.seek(0)
    return buffer



def _supabase_config():
    """Lit les paramètres d'authentification dans les Secrets Streamlit."""
    try:
        url = str(st.secrets.get("SUPABASE_URL", "")).strip().rstrip("/")
        key = str(
            st.secrets.get("SUPABASE_ANON_KEY", "")
            or st.secrets.get("SUPABASE_KEY", "")
        ).strip()
        admins = str(st.secrets.get("HSUP_ADMIN_EMAILS", "")).strip()
        service_role = str(st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")).strip()
    except Exception:
        url, key, admins, service_role = "", "", "", ""
    return url, key, {x.strip().casefold() for x in admins.split(",") if x.strip()}, service_role


def _supabase_requete(url, anon_key, chemin, donnees, access_token=None):
    """Appel REST Supabase sans dépendance supplémentaire."""
    corps = json.dumps(donnees, ensure_ascii=False).encode("utf-8")
    headers = {"apikey": anon_key, "Content-Type": "application/json"}
    if access_token:
        headers["Authorization"] = f"Bearer {access_token}"
    req = urllib.request.Request(
        f"{url}/auth/v1/{chemin.lstrip('/')}",
        data=corps, headers=headers, method="POST" if chemin != "user" else "PUT"
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as response:
            return json.loads(response.read().decode("utf-8")), None
    except urllib.error.HTTPError as exc:
        try:
            detail = json.loads(exc.read().decode("utf-8"))
            message = detail.get("msg") or detail.get("message") or detail.get("error_description") or str(detail)
        except Exception:
            message = str(exc)
        return None, message
    except Exception as exc:
        return None, str(exc)


def _supabase_table_requete(table, methode="GET", donnees=None, params="", admin=False):
    """Accès PostgREST à la table hsup_preferences.

    Les enseignants utilisent leur JWT (RLS limité à auth.uid()). Les opérations
    globales d'administration utilisent uniquement la clé service_role côté serveur.
    """
    url, anon_key, _, service_role = _supabase_config()
    token = service_role if admin else st.session_state.get("hsup_access_token", "")
    cle = service_role if admin and service_role else anon_key
    if not url or not cle or not token:
        return None, "Configuration Supabase incomplète (clé API, session ou service role)."
    url_requete = f"{url}/rest/v1/{table}"
    if params:
        url_requete += "?" + params
    headers = {"apikey": cle, "Authorization": f"Bearer {token}",
               "Content-Type": "application/json", "Accept": "application/json"}
    if methode == "POST":
        headers["Prefer"] = "resolution=merge-duplicates,return=representation"
    corps = None if donnees is None else json.dumps(donnees, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url_requete, data=corps, headers=headers, method=methode)
    try:
        with urllib.request.urlopen(req, timeout=25) as response:
            contenu = response.read().decode("utf-8")
            return (json.loads(contenu) if contenu else []), None
    except urllib.error.HTTPError as exc:
        try:
            detail = json.loads(exc.read().decode("utf-8"))
            message = detail.get("message") or detail.get("hint") or str(detail)
        except Exception:
            message = str(exc)
        return None, message
    except Exception as exc:
        return None, str(exc)


def sauvegarder_preferences_hsup(
    user_id,
    nom_enseignant,
    selections=None,
    exclu=None,
    email=None,
    telephone=None,
    ism_arabe=None,
    laqab_arabe=None,
):
    """Crée ou met à jour la préférence H.SUP du compte enseignant."""
    if not user_id:
        return False, "Identifiant utilisateur Supabase absent."
    ligne = {"user_id": user_id, "teacher_name": nom_enseignant,
             "selected_teachings": list(selections or [])}
    if exclu is not None:
        ligne["excluded"] = bool(exclu)
    if email is not None:
        ligne["email"] = nettoyer_texte(email)
    if telephone is not None:
        ligne["telephone"] = nettoyer_texte(telephone)
    if ism_arabe is not None:
        ligne["ism_arabe"] = nettoyer_texte(ism_arabe)
    if laqab_arabe is not None:
        ligne["laqab_arabe"] = nettoyer_texte(laqab_arabe)
    resultat, erreur = _supabase_table_requete(
        "hsup_preferences", "POST", ligne, "on_conflict=user_id", admin=False
    )
    return (erreur is None), (erreur or "")


def charger_preferences_hsup_admin():
    """Charge les choix personnels et les exclusions globales avec la clé serveur."""
    resultat, erreur = _supabase_table_requete(
        "hsup_preferences", "GET", params="select=user_id,teacher_name,selected_teachings,excluded&order=teacher_name.asc",
        admin=True
    )
    if erreur:
        return {}, set(), erreur
    selections, exclus = {}, set()
    for ligne in resultat or []:
        nom = normaliser_nom(ligne.get("teacher_name", ""))
        if nom:
            selections[nom] = ligne.get("selected_teachings") or []
            if ligne.get("excluded"):
                exclus.add(nom)
    # Les exclusions sont stockées indépendamment des comptes enseignants :
    # l'administration peut donc exclure aussi un enseignant non encore inscrit.
    reglages, erreur_reglages = _supabase_table_requete(
        "hsup_admin_settings", "GET", params="select=excluded_teachers&id=eq.1", admin=True
    )
    if erreur_reglages:
        return {}, set(), erreur_reglages
    if reglages:
        exclus = {normaliser_nom(x) for x in (reglages[0].get("excluded_teachers") or [])}
    return selections, exclus, ""


def mettre_a_jour_exclusions_admin(noms_exclus, df_enseignants=None):
    """Enregistre la liste globale des enseignants exclus, indépendamment des comptes."""
    url, anon_key, _, service_role = _supabase_config()
    if not service_role:
        return False, "Ajoutez SUPABASE_SERVICE_ROLE_KEY aux Secrets Streamlit pour administrer les exclusions globales."
    liste = sorted({nettoyer_texte(n) for n in noms_exclus if nettoyer_texte(n)})
    payload = {"id": 1, "excluded_teachers": liste}
    _, erreur = _supabase_table_requete(
        "hsup_admin_settings", "POST", payload, "on_conflict=id", admin=True
    )
    return (erreur is None), (erreur or "")


def charger_coordonnees_enseignants_admin():
    """Charge les coordonnées saisies par les enseignants dans Supabase."""
    resultat, erreur = _supabase_table_requete(
        "hsup_preferences",
        "GET",
        params=(
            "select=user_id,teacher_name,email,telephone,"
            "ism_arabe,laqab_arabe&order=teacher_name.asc"
        ),
        admin=True,
    )

    coordonnees = {}
    for ligne in resultat or []:
        nom = normaliser_nom(ligne.get("teacher_name", ""))
        if not nom:
            continue
        coordonnees[nom] = {
            "email": nettoyer_texte(ligne.get("email", "")),
            "telephone": nettoyer_texte(ligne.get("telephone", "")),
            "nom_arabe": " ".join(
                x for x in (
                    nettoyer_texte(ligne.get("ism_arabe", "")),
                    nettoyer_texte(ligne.get("laqab_arabe", "")),
                ) if x
            ),
        }

    return coordonnees, erreur


def enregistrer_coordonnees_enseignant(
    user_id,
    nom_enseignant,
    email,
    telephone,
    ism_arabe,
    laqab_arabe,
):
    """Enregistre ou met à jour les coordonnées de l'enseignant dans Supabase."""
    if not user_id:
        return False, "Identifiant utilisateur Supabase absent."

    ligne = {
        "user_id": user_id,
        "teacher_name": nom_enseignant,
        "email": nettoyer_texte(email),
        "telephone": nettoyer_texte(telephone),
        "ism_arabe": nettoyer_texte(ism_arabe),
        "laqab_arabe": nettoyer_texte(laqab_arabe),
    }

    _, erreur = _supabase_table_requete(
        "hsup_preferences", "POST", ligne, "on_conflict=user_id", admin=False
    )
    return (erreur is None), (erreur or "")


def charger_reglages_admin_enseignants():
    """Charge les réglages administratifs des enseignants (H.SUP, jours et identité)."""
    resultat, erreur = _supabase_table_requete(
        "hsup_admin_teacher_settings",
        "GET",
        params=(
            "select=teacher_name,selected_teachings,excluded_dates,"
            "ism_arabe,laqab_arabe&order=teacher_name.asc"
        ),
        admin=True,
    )

    selections = {}
    dates_exclues = {}
    identites = {}

    for ligne in resultat or []:
        nom = normaliser_nom(ligne.get("teacher_name", ""))
        if not nom:
            continue
        selections[nom] = ligne.get("selected_teachings") or []
        dates_brutes = ligne.get("excluded_dates") or []
        dates_converties = []
        for d in dates_brutes:
            try:
                dates_converties.append(
                    d if isinstance(d, date) else date.fromisoformat(str(d))
                )
            except (TypeError, ValueError):
                continue
        dates_exclues[nom] = dates_converties
        identites[nom] = {
            "ism": nettoyer_texte(ligne.get("ism_arabe", "")),
            "laqab": nettoyer_texte(ligne.get("laqab_arabe", "")),
        }

    return selections, dates_exclues, identites, erreur


def enregistrer_reglages_admin_enseignant(
    nom_enseignant,
    selections,
    dates_exclues,
    ism_arabe=None,
    laqab_arabe=None,
):
    """Enregistre les réglages administratifs d'un enseignant dans Supabase."""
    nom_norm = normaliser_nom(nom_enseignant)
    if not nom_norm:
        return False, "Nom d'enseignant absent."

    dates_iso = sorted(
        d.isoformat() for d in (dates_exclues or []) if isinstance(d, date)
    )

    payload = {
        "teacher_name": nom_norm,
        "selected_teachings": list(selections or []),
        "excluded_dates": dates_iso,
    }
    if ism_arabe is not None:
        payload["ism_arabe"] = nettoyer_texte(ism_arabe)
    if laqab_arabe is not None:
        payload["laqab_arabe"] = nettoyer_texte(laqab_arabe)

    _, erreur = _supabase_table_requete(
        "hsup_admin_teacher_settings",
        "POST",
        payload,
        "on_conflict=teacher_name",
        admin=True,
    )
    return (erreur is None), (erreur or "")


def construire_exclusions_par_mois_admin(dates_exclues):
    """Convertit une liste de dates exclues en dictionnaire par mois."""
    exclusions = {}
    for d in dates_exclues or []:
        try:
            d_date = d if isinstance(d, date) else date.fromisoformat(str(d))
        except (TypeError, ValueError):
            continue
        exclusions.setdefault(d_date.month, set()).add(d_date)
    return exclusions


def extraire_noms_enseignants_edt(df_edt_source):
    """Extrait les noms d'enseignants depuis la colonne Enseignants de l'EDT."""
    noms = {}
    for valeur in df_edt_source["Enseignants"].fillna("").astype(str):
        for morceau in re.split(r"[,;]|\bet\b", valeur, flags=re.IGNORECASE):
            nom = nettoyer_texte(morceau)
            if not nom:
                continue
            cle = normaliser_nom(nom)
            if cle and cle not in noms:
                noms[cle] = nom
    return [noms[cle] for cle in sorted(noms)]


def construire_details_enseignants_historique(noms_enseignants, seances_par_enseignant=None):
    """Construit les détails par enseignant pour l'historique des générations."""
    seances_par_enseignant = seances_par_enseignant or {}
    details = []

    for nom in noms_enseignants:
        nom = nettoyer_texte(nom)
        nom_norm = normaliser_nom(nom)

        ligne = df_enseignants[df_enseignants["NOM_NORM"].eq(nom_norm)]
        qualite = (
            nettoyer_texte(ligne.iloc[0].get("Qualité", ""))
            if not ligne.empty
            else ""
        ) or "Enseignant"

        coordonnees = coordonnees_enseignants.get(nom_norm, {})

        nom_arabe = coordonnees.get("nom_arabe", "")
        email = coordonnees.get("email", "")
        telephone = coordonnees.get("telephone", "")

        if not nom_arabe:
            identite = identite_arabe_de(dict_noms_arabe, nom)
            if isinstance(identite, dict):
                ism = nettoyer_texte(identite.get("ism", ""))
                laqab = nettoyer_texte(identite.get("laqab", ""))
                nom_arabe = " ".join(x for x in (ism, laqab) if x)

        if not email or not telephone:
            if not ligne.empty:
                ligne_dict = ligne.iloc[0].to_dict()
                for colonne, valeur in ligne_dict.items():
                    entete = normaliser_entete(colonne)
                    valeur = nettoyer_texte(valeur)
                    if not valeur:
                        continue
                    if not email and any(
                        mot in entete for mot in ("email", "e mail", "adresse e mail")
                    ):
                        email = valeur
                    if not telephone and any(
                        mot in entete for mot in ("telephone", "téléphone", "phone", "mobile", "gsm")
                    ):
                        telephone = valeur

        if (not email or not telephone) and FICHIER_COORD.exists():
            try:
                df_coord = pd.read_excel(FICHIER_COORD, dtype=str)
                df_coord.columns = [nettoyer_texte(c) for c in df_coord.columns]

                for _, ligne_coord in df_coord.iterrows():
                    email_coord = ""
                    telephone_coord = ""
                    nom_coord = ""

                    for colonne, valeur in ligne_coord.items():
                        entete = normaliser_entete(colonne)
                        valeur = nettoyer_texte(valeur)
                        if not valeur:
                            continue
                        if any(mot in entete for mot in ("email", "e mail", "adresse e mail")):
                            email_coord = valeur
                        if any(mot in entete for mot in ("telephone", "téléphone", "phone", "mobile", "gsm")):
                            telephone_coord = valeur
                        if any(mot in entete for mot in ("nom et prenom", "nom complet", "nom prenom", "nom")):
                            nom_coord = valeur

                    correspondance = (
                        (email and email_coord and email.casefold() == email_coord.casefold())
                        or (nom_norm and nom_coord and normaliser_nom(nom_coord) == nom_norm)
                    )

                    if correspondance:
                        email = email or email_coord
                        telephone = telephone or telephone_coord
                        break
            except Exception:
                pass

        total_charge = ""
        if nom_norm in seances_par_enseignant:
            seances_enseignant = seances_par_enseignant[nom_norm]
            if seances_enseignant is not None and not seances_enseignant.empty:
                total_heures = float(
                    pd.to_numeric(
                        seances_enseignant["DUREE_H"], errors="coerce"
                    ).fillna(0).sum()
                )
                total_charge = f"{total_heures:.2f} h"

        details.append({
            "nom": nom,
            "qualite": qualite,
            "nom_arabe": nom_arabe,
            "email": email,
            "telephone": telephone,
            "total_charge": total_charge,
        })

    return details


def enregistrer_historique_generation(
    generation_type,
    enseignants,
    format_fichier,
    nom_fichier,
    statut="Succès",
    message="",
    total_hsup=None,
    enseignants_details=None,
):
    """Enregistre une génération d'état H.SUP dans l'historique Supabase."""
    profil = st.session_state.get("hsup_user") or {}
    email_admin = nettoyer_texte(profil.get("email", ""))

    if isinstance(enseignants, str):
        liste_enseignants = [enseignants]
    else:
        liste_enseignants = [
            nettoyer_texte(x) for x in (enseignants or []) if nettoyer_texte(x)
        ]

    if enseignants_details is None:
        enseignants_details = construire_details_enseignants_historique(
            liste_enseignants
        )

    payload = {
        "generated_by": email_admin,
        "generation_type": nettoyer_texte(generation_type),
        "teachers": liste_enseignants,
        "teacher_count": len(liste_enseignants),
        "teacher_details": enseignants_details,
        "period_start": date_debut_cours.isoformat(),
        "period_end": date_fin_comptabilisation.isoformat(),
        "file_format": nettoyer_texte(format_fichier),
        "file_name": nettoyer_texte(nom_fichier),
        "status": nettoyer_texte(statut) or "Succès",
        "message": nettoyer_texte(message),
        "total_hsup": None if total_hsup is None else str(total_hsup),
    }

    _, erreur = _supabase_table_requete(
        "hsup_generation_history", "POST", payload, admin=True
    )
    return (erreur is None), (erreur or "")


def charger_historique_generations(limite=500):
    """Charge l'historique des générations d'états H.SUP, du plus récent au plus ancien."""
    resultat, erreur = _supabase_table_requete(
        "hsup_generation_history",
        "GET",
        params=(
            "select=id,created_at,generated_by,generation_type,teachers,"
            "teacher_details,teacher_count,period_start,period_end,"
            "file_format,file_name,status,message,total_hsup"
            f"&order=created_at.desc&limit={int(limite)}"
        ),
        admin=True,
    )
    return (resultat or []), erreur


def generer_excel_historique(historique):
    """Génère un classeur Excel contenant l'historique des générations H.SUP."""
    colonnes = [
        "Date et heure",
        "Administrateur",
        "Type de génération",
        "Enseignants",
        "Nom et prénom en arabe",
        "E-mail",
        "Téléphone",
        "Qualité",
        "Total charge (vacataire)",
        "Début de période",
        "Fin de période",
        "Total H.SUP",
    ]

    lignes = []
    for ligne in historique:
        enseignants = ligne.get("teachers") or []
        if isinstance(enseignants, str):
            enseignants = [enseignants]

        details = ligne.get("teacher_details") or []
        if not isinstance(details, list):
            details = []

        details_par_nom = {}
        for detail in details:
            if isinstance(detail, dict):
                details_par_nom[normaliser_nom(detail.get("nom", ""))] = detail

        for nom in enseignants:
            detail = details_par_nom.get(normaliser_nom(nom), {})

            lignes.append({
                "Date et heure": nettoyer_texte(ligne.get("created_at", "")),
                "Administrateur": nettoyer_texte(ligne.get("generated_by", "")),
                "Type de génération": nettoyer_texte(ligne.get("generation_type", "")),
                "Enseignants": nettoyer_texte(nom),
                "Nom et prénom en arabe": nettoyer_texte(detail.get("nom_arabe", "")),
                "E-mail": nettoyer_texte(detail.get("email", "")),
                "Téléphone": nettoyer_texte(detail.get("telephone", "")),
                "Qualité": nettoyer_texte(detail.get("qualite", "")),
                "Total charge (vacataire)": nettoyer_texte(detail.get("total_charge", "")),
                "Début de période": nettoyer_texte(ligne.get("period_start", "")),
                "Fin de période": nettoyer_texte(ligne.get("period_end", "")),
                "Total H.SUP": nettoyer_texte(ligne.get("total_hsup", "")),
            })

    df_historique = pd.DataFrame(lignes, columns=colonnes)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_historique.to_excel(writer, sheet_name="Historique H.SUP", index=False)

        ws = writer.sheets["Historique H.SUP"]
        largeurs = {
            "A": 22, "B": 32, "C": 30, "D": 35, "E": 35, "F": 35, "G": 20,
            "H": 18, "I": 22, "J": 16, "K": 16, "L": 18,
        }
        for colonne, largeur in largeurs.items():
            ws.column_dimensions[colonne].width = largeur

        for cellule in ws[1]:
            cellule.font = openpyxl.styles.Font(bold=True)
            cellule.fill = openpyxl.styles.PatternFill(
                start_color="D9EAF7", end_color="D9EAF7", fill_type="solid"
            )
            cellule.alignment = openpyxl.styles.Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    buffer.seek(0)
    return buffer, df_historique


def authentifier_enseignant(liste_enseignants):
    """Inscription/connexion. L'identité française est liée au référentiel EDT."""
    url, anon_key, admins, service_role = _supabase_config()
    if not url or not anon_key:
        st.error(
            "L'inscription sécurisée n'est pas encore activée. "
            "Configure SUPABASE_URL et SUPABASE_ANON_KEY dans les Secrets Streamlit."
        )
        st.info(
            "La création de comptes nécessite un service d'authentification persistant. "
            "Les comptes ne doivent pas être enregistrés dans un fichier local de l'application."
        )
        st.stop()

    if st.session_state.get("hsup_access_token"):
        profil = st.session_state.get("hsup_user", {})
        email = str(profil.get("email", "")).casefold()
        nom = nettoyer_texte((profil.get("user_metadata") or {}).get("nom_enseignant", ""))
        if nom not in liste_enseignants:
            st.error("Le compte connecté n'est pas associé à un enseignant actif du fichier personnel/EDT.")
            if st.button("Se déconnecter", key="hsup_logout_invalid"):
                st.session_state.pop("hsup_access_token", None)
                st.session_state.pop("hsup_user", None)
                st.rerun()
            st.stop()
        st.sidebar.success(f"Compte connecté : {email}")
        if email in admins:
            st.sidebar.success("🛠️ Espace administration activé")
            st.sidebar.caption("Les outils administrateur se trouvent dans la section dédiée en bas de la page.")
        else:
            st.sidebar.info("Espace enseignant")
            st.sidebar.caption(
                "L'accès administration est réservé aux adresses déclarées dans "
                "HSUP_ADMIN_EMAILS dans .streamlit/secrets.toml."
            )
        if st.sidebar.button("Déconnexion", key="hsup_logout"):
            st.session_state.pop("hsup_access_token", None)
            st.session_state.pop("hsup_user", None)
            st.rerun()
        return nom, email in admins, profil

    st.subheader("🔐 Espace enseignant — Connexion et inscription")
    tab_connexion, tab_inscription = st.tabs(["Se connecter", "Créer mon compte"])
    with tab_connexion:
        with st.form("hsup_form_connexion"):
            email_login = st.text_input("Adresse e-mail", key="hsup_email_login")
            mdp_login = st.text_input("Mot de passe", type="password", key="hsup_mdp_login")
            envoyer_login = st.form_submit_button("Se connecter", type="primary", use_container_width=True)
        if envoyer_login:
            if not email_login.strip() or not mdp_login:
                st.error("Saisissez votre adresse e-mail et votre mot de passe.")
            else:
                resultat, erreur = _supabase_requete(
                    url, anon_key, "token?grant_type=password",
                    {"email": email_login.strip(), "password": mdp_login}
                )
                if erreur:
                    st.error(f"Connexion impossible : {erreur}")
                elif resultat and resultat.get("access_token"):
                    st.session_state["hsup_access_token"] = resultat["access_token"]
                    st.session_state["hsup_user"] = resultat.get("user", {})
                    st.rerun()

    with tab_inscription:
        with st.form("hsup_form_inscription"):
            nom_inscription = st.selectbox(
                "Choisir votre nom dans le fichier du personnel",
                options=liste_enseignants,
                key="hsup_nom_inscription"
            )
            email_inscription = st.text_input("Adresse e-mail professionnelle ou personnelle")
            telephone_inscription = st.text_input(
                "Numéro de téléphone — obligatoire",
                placeholder="Ex. : 0550123456 ou +213550123456",
                max_chars=20,
            )
            mdp_inscription = st.text_input("Créer un mot de passe", type="password")
            mdp_confirmation = st.text_input("Confirmer le mot de passe", type="password")
            st.markdown("**Renseignements obligatoires en arabe**")
            ism_inscription = st.text_input("الاسم — prénom en arabe (obligatoire)")
            laqab_inscription = st.text_input("اللقب — nom en arabe (obligatoire)")
            envoyer_inscription = st.form_submit_button(
                "Créer mon compte", type="primary", use_container_width=True
            )
        if envoyer_inscription:
            telephone_nettoye = re.sub(r"[\s().-]", "", telephone_inscription.strip())
            telephone_chiffres = re.sub(r"\D", "", telephone_nettoye)
            arabe_present = bool(re.search(r"[\u0600-\u06FF]", ism_inscription)) and bool(
                re.search(r"[\u0600-\u06FF]", laqab_inscription)
            )
            if not email_inscription.strip() or "@" not in email_inscription or len(mdp_inscription) < 8:
                st.error("Indiquez une adresse e-mail valide et un mot de passe d'au moins 8 caractères.")
            elif not ism_inscription.strip() or not laqab_inscription.strip() or not arabe_present:
                st.error("Le prénom (الاسم) et le nom (اللقب) doivent être renseignés en caractères arabes.")
            elif len(telephone_chiffres) < 8 or len(telephone_chiffres) > 15:
                st.error("Le numéro de téléphone est obligatoire et doit contenir entre 8 et 15 chiffres.")
            elif mdp_inscription != mdp_confirmation:
                st.error("Les deux mots de passe ne sont pas identiques.")
            else:
                metadata = {
                    "nom_enseignant": nom_inscription,
                    "ism_arabe": ism_inscription.strip(),
                    "laqab_arabe": laqab_inscription.strip(),
                    "telephone": telephone_nettoye,
                    "role": "enseignant"
                }
                resultat, erreur = _supabase_requete(
                    url, anon_key, "signup",
                    {"email": email_inscription.strip(), "password": mdp_inscription, "data": metadata}
                )
                if erreur:
                    st.error(f"Inscription impossible : {erreur}")
                elif resultat and resultat.get("access_token"):
                    st.session_state["hsup_access_token"] = resultat["access_token"]
                    st.session_state["hsup_user"] = resultat.get("user", {})

                    user_id_inscription = (
                        resultat.get("user", {}).get("id", "")
                    )
                    if user_id_inscription:
                        enregistrer_coordonnees_enseignant(
                            user_id_inscription,
                            nom_inscription,
                            email_inscription.strip(),
                            telephone_nettoye,
                            ism_inscription.strip(),
                            laqab_inscription.strip(),
                        )

                    st.success("Compte créé et connecté.")
                    st.rerun()
                else:
                    st.success(
                        "Demande de création envoyée. Consultez votre boîte e-mail pour confirmer "
                        "l'adresse, puis connectez-vous."
                    )
    st.stop()


def mettre_a_jour_metadonnees_compte(donnees):
    """Enregistre les préférences personnelles dans les métadonnées du compte Supabase."""
    url, anon_key, _, _service_role = _supabase_config()
    token = st.session_state.get("hsup_access_token")
    if not token:
        return False, "Aucun compte connecté."
    profil, erreur = _supabase_requete(
        url, anon_key, "user", {"data": donnees}, access_token=token
    )
    if erreur:
        return False, erreur
    if profil:
        st.session_state["hsup_user"] = profil
    return True, ""


def enregistrer_identite_enseignant_compte(ism, laqab, telephone):
    """Enregistre l'identité arabe et le téléphone dans le compte."""
    ok_meta, erreur_meta = mettre_a_jour_metadonnees_compte({
        "ism_arabe": ism,
        "laqab_arabe": laqab,
        "telephone": telephone,
    })

    if not ok_meta:
        return False, erreur_meta

    user_id = (
        st.session_state.get("hsup_user") or {}
    ).get("id", "")

    if user_id:
        ok_table, erreur_table = enregistrer_coordonnees_enseignant(
            user_id,
            enseignant_select,
            st.session_state.get("hsup_user", {}).get("email", ""),
            telephone,
            ism,
            laqab,
        )
        if not ok_table:
            return False, erreur_table

    return True, ""


def lire_date_compte(valeur, defaut):
    """Convertit une date ISO stockée dans le compte, avec une valeur de repli sûre."""
    try:
        return date.fromisoformat(str(valeur)) if valeur else defaut
    except (TypeError, ValueError):
        return defaut

try:
    df_enseignants, df_edt, dict_noms_arabe, liste_enseignants_actifs = charger_donnees()
    df_edt_complet = charger_edt_complet()
    coordonnees_enseignants, erreur_coordonnees_enseignants = (
        charger_coordonnees_enseignants_admin()
    )
    enseignant_connecte, est_administrateur, profil_compte = authentifier_enseignant(
        liste_enseignants_actifs
    )
    enseignant_select = enseignant_connecte

    st.sidebar.header("🔍 Paramètres & calendrier")
    metadata_compte_cal = profil_compte.get("user_metadata") or {}
    date_debut_defaut = lire_date_compte(
        metadata_compte_cal.get("hsup_date_debut"), date(2026, 9, 23)
    )
    date_fin_defaut = lire_date_compte(
        metadata_compte_cal.get("hsup_date_fin"), date(2026, 12, 17)
    )
    date_debut_cours = st.sidebar.date_input(
        "Date de début de comptabilisation des H.SUP",
        value=date_debut_defaut,
        min_value=date(2026, 1, 1),
        max_value=date(2026, 12, 31),
        format="DD/MM/YYYY",
        key="hsup_date_debut_widget",
    )
    date_fin_comptabilisation = st.sidebar.date_input(
        "Date de fin de comptabilisation des H.SUP",
        value=date_fin_defaut,
        min_value=date(2026, 1, 1),
        max_value=date(2026, 12, 31),
        format="DD/MM/YYYY",
        key="hsup_date_fin_widget",
    )
    if date_debut_cours > date_fin_comptabilisation:
        st.sidebar.error("La date de début doit être antérieure ou égale à la date de fin.")
        st.stop()
    st.sidebar.info(
        f"Période comptabilisée : du {date_debut_cours.strftime('%d/%m/%Y')} "
        f"au {date_fin_comptabilisation.strftime('%d/%m/%Y')} inclus."
    )
    st.sidebar.markdown("**✍️ Signataires de l'état H.SUP**")
    SIGNATAIRES["naib_chef"] = st.sidebar.text_input(
        "Nom sous « نائبي رئيس القسم »", value=SIGNATAIRES["naib_chef"], key="signataire_naib_chef")
    SIGNATAIRES["chef"] = st.sidebar.text_input(
        "Nom sous « رئيس القسم »", value=SIGNATAIRES["chef"], key="signataire_chef")
    st.sidebar.text_input(
        "Enseignant connecté", value=enseignant_select, disabled=True,
        help="Votre compte est lié à votre identité dans le fichier du personnel."
    )

    if not enseignant_select:
        st.error("Aucun enseignant permanent ou vacataire avec une charge horaire n'a été trouvé dans l'EDT.")
        st.stop()

    nom_norm = normaliser_nom(enseignant_select)
    seances_enseignant = df_edt[
        df_edt["Enseignants"].apply(lambda x: correspond_enseignant(x, nom_norm))
    ].copy()

    # Toutes les charges de l'enseignant, tous départements et promotions confondus.
    seances_enseignant_complet = df_edt_complet[
        df_edt_complet["Enseignants"].apply(lambda x: correspond_enseignant(x, nom_norm))
    ].copy()

    # Contrôles de calendrier : les jours fériés sont exclus automatiquement.
    st.subheader("📅 Calendrier et jours à exclure")
    st.info(
        "Les heures sont calculées JOUR PAR JOUR et date par date, uniquement entre les deux dates "
        "choisies (par défaut du 23 septembre au 17 décembre 2026 inclus). "
        "Les vendredis, samedis et jours fériés nationaux fixes connus sont exclus. "
        "Sélectionne ci-dessous toute autre date sans enseignement à comptabiliser."
    )
    exclusions_par_mois = {}
    exclusions_sauvees = metadata_compte_cal.get("hsup_dates_exclues", {}) or {}
    colonnes = st.columns(len(MOIS_BLOCKS))
    for col, bloc in zip(colonnes, MOIS_BLOCKS):
        with col:
            st.markdown(f"**{bloc['nom']}**")
            labels = etiquettes_dates(bloc["annee"], bloc["mois"])
            dates_dans_periode = [
                cle for cle in labels
                if date.fromisoformat(cle) >= date_debut_cours
                and date.fromisoformat(cle) <= date_fin_comptabilisation
            ]
            cle_widget = f"exclusions_{bloc['annee']}_{bloc['mois']}"
            deja_exclues = exclusions_sauvees.get(str(bloc["mois"]), [])
            deja_exclues = [x for x in deja_exclues if x in dates_dans_periode]
            selection_actuelle = st.session_state.get(cle_widget, deja_exclues)
            st.session_state[cle_widget] = [x for x in selection_actuelle if x in dates_dans_periode]
            choix = st.multiselect(
                f"Dates à exclure – {bloc['nom']}",
                options=dates_dans_periode,
                format_func=lambda x, d=labels: d[x],
                key=cle_widget,
            )
            exclusions_par_mois[bloc["mois"]] = {
                date.fromisoformat(x) for x in choix
            }

    if st.button("💾 Enregistrer ma période et mes jours à exclure", use_container_width=True,
                 key="hsup_save_calendar"):
        exclusions_a_sauver = {
            str(mois): sorted(d.isoformat() for d in dates)
            for mois, dates in exclusions_par_mois.items()
        }
        ok_cal, erreur_cal = mettre_a_jour_metadonnees_compte({
            "hsup_date_debut": date_debut_cours.isoformat(),
            "hsup_date_fin": date_fin_comptabilisation.isoformat(),
            "hsup_dates_exclues": exclusions_a_sauver,
        })
        if ok_cal:
            st.success("Votre période H.SUP et vos jours à exclure ont été enregistrés dans votre compte.")
            st.rerun()
        else:
            st.error(f"Impossible d'enregistrer vos paramètres : {erreur_cal}")

    st.divider()
    st.subheader(f"👨‍🏫 Fiche enseignant : {enseignant_select}")
    identite_arabe, mode_arabe, cle_arabe = trouver_identite_arabe(dict_noms_arabe, enseignant_select)
    metadata_compte = profil_compte.get("user_metadata") or {}
    ism_arabe = nettoyer_texte(metadata_compte.get("ism_arabe", "")) or nettoyer_texte(
        identite_arabe.get("ism", "") if isinstance(identite_arabe, dict) else ""
    )
    laqab_arabe = nettoyer_texte(metadata_compte.get("laqab_arabe", "")) or nettoyer_texte(
        identite_arabe.get("laqab", "") if isinstance(identite_arabe, dict) else ""
    )
    nom_arabe = " ".join(x for x in (ism_arabe, laqab_arabe) if x)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Nom en arabe", nom_arabe or "Non renseigné")
    ligne_pv = df_enseignants[df_enseignants["NOM_NORM"].eq(nom_norm)].iloc[0]
    c2.metric(
        "Qualité / Grade",
        " — ".join(x for x in (
            nettoyer_texte(ligne_pv.get("Qualité", "")),
            nettoyer_texte(ligne_pv.get("Grade", ""))
        ) if x) or "Non renseigné"
    )
    c3.metric("Séances EDT", len(seances_enseignant))
    c4.metric("Volume hebdomadaire EDT", f"{seances_enseignant['DUREE_H'].sum():g} h")
    charge_connecte = calculer_charge_equivalente_hebdomadaire(seances_enseignant)
    depassement_connecte = max(0.0, charge_connecte["total_eq"] - 7.0) if "permanent" in nettoyer_texte(ligne_pv.get("Qualité", "")).casefold() else 0.0
    if "permanent" in nettoyer_texte(ligne_pv.get("Qualité", "")).casefold():
        st.info(
            f"**Charge réglementaire permanente : 7 h équivalent-cours/semaine.** "
            f"Cours : {charge_connecte['cours']:.2f} h ; TD : {charge_connecte['td']:.2f} h ; "
            f"TP : {charge_connecte['tp']:.2f} h. Charge équivalente : "
            f"{charge_connecte['total_eq']:.2f} h ; dépassement : {depassement_connecte:.2f} h éq-cours. "
            "Les TD et TP sont convertis au taux de 1,5 h pour 1 h équivalent-cours."
        )
    if not nom_arabe:
        st.warning(
            "Le fichier source ne contient pas encore le nom et le prénom en arabe. "
            "Renseignez-les ci-dessous pour compléter votre fiche H.SUP."
        )
    telephone_compte = nettoyer_texte(metadata_compte.get("telephone", ""))
    with st.form("hsup_identite_arabe_form"):
        col_ism, col_laqab = st.columns(2)
        ism_saisi = col_ism.text_input("الاسم — prénom en arabe (obligatoire)", value=ism_arabe)
        laqab_saisi = col_laqab.text_input("اللقب — nom en arabe (obligatoire)", value=laqab_arabe)
        telephone_saisi = st.text_input(
            "Numéro de téléphone (obligatoire)", value=telephone_compte,
            placeholder="Ex. : 0550123456 ou +213550123456", max_chars=20
        )
        sauvegarder_arabe = st.form_submit_button("Enregistrer mes renseignements personnels")
    if sauvegarder_arabe:
        telephone_nettoye = re.sub(r"[\s().-]", "", telephone_saisi.strip())
        telephone_chiffres = re.sub(r"\D", "", telephone_nettoye)
        champs_arabes_valides = (
            bool(re.search(r"[\u0600-\u06FF]", ism_saisi))
            and bool(re.search(r"[\u0600-\u06FF]", laqab_saisi))
        )
        if not champs_arabes_valides:
            st.error("Le prénom et le nom doivent être renseignés en caractères arabes.")
        elif len(telephone_chiffres) < 8 or len(telephone_chiffres) > 15:
            st.error("Le numéro de téléphone est obligatoire et doit contenir entre 8 et 15 chiffres.")
        else:
            ok, erreur = enregistrer_identite_enseignant_compte(
                ism_saisi.strip(), laqab_saisi.strip(), telephone_nettoye
            )
            if ok:
                st.success("Vos renseignements personnels ont été enregistrés dans votre compte.")
                st.rerun()
            else:
                st.error(f"Enregistrement impossible : {erreur}")
    dict_noms_arabe[normaliser_nom(enseignant_select)] = {
        "ism": ism_saisi.strip() if 'ism_saisi' in locals() else ism_arabe,
        "laqab": laqab_saisi.strip() if 'laqab_saisi' in locals() else laqab_arabe,
        "complet": " ".join(x for x in (
            ism_saisi.strip() if 'ism_saisi' in locals() else ism_arabe,
            laqab_saisi.strip() if 'laqab_saisi' in locals() else laqab_arabe
        ) if x)
    }
    if mode_arabe != "exacte" and nom_arabe:
        libelles_mode = {
            "ordre": "mêmes mots dans un autre ordre",
            "orthographe": "orthographe légèrement différente (lettres doublées)",
            "approchee": "orthographe proche",
            "partielle": "nom incomplet d'un côté",
        }
        st.info(
            f"Correspondance du formulaire arabe : « {enseignant_select} » ↔ « {cle_arabe} » "
            f"({libelles_mode.get(mode_arabe, mode_arabe)}). Vérifiez que c'est bien le même enseignant."
        )

    colonnes_aff = [
        c for c in ["Enseignements", "Code", "Horaire", "Jours", "Lieu", "Promotion"]
        if c in seances_enseignant.columns
    ]
    st.markdown("### Séances hebdomadaires issues de l'EDT")
    st.dataframe(seances_enseignant[colonnes_aff], use_container_width=True, hide_index=True)

    # ---- Emploi du temps individuel en grille (PDF, dimensions du modèle PPER.03) ----
    st.markdown("### 🗓️ Emploi du temps individuel (grille PDF)")
    qualite_select = nettoyer_texte(ligne_pv.get("Qualité", "")) or "Enseignant"
    enseignements_hsup_selectionnes = None
    if "permanent" in qualite_select.casefold():
        st.markdown("### ⚙️ Déclaration des enseignements à comptabiliser comme heures supplémentaires")
        st.info(
            "Sélectionnez uniquement les enseignements que vous souhaitez faire apparaître "
            "dans votre état H.SUP. Les séances simultanées de même type (Cours, TD ou TP), "
            "même si elles concernent plusieurs promotions, seront comptées une seule fois. "
            "Cette sélection ne modifie pas l'emploi du temps affiché."
        )
        metadata_hsup = profil_compte.get("user_metadata") or {}
        options_hsup = options_enseignements_hsup(seances_enseignant)
        deja_hsup = metadata_hsup.get("hsup_enseignements_retenus", []) or []
        user_id_local = (profil_compte or {}).get("id")
        if user_id_local:
            prefs_local, err_local = _supabase_table_requete(
                "hsup_preferences", "GET",
                params="select=selected_teachings&user_id=eq." + urllib.parse.quote(str(user_id_local)),
                admin=False
            )
            if not err_local and prefs_local:
                deja_hsup = prefs_local[0].get("selected_teachings") or []
        deja_hsup = selections_vers_options_hsup(deja_hsup, options_hsup)
        enseignements_hsup_selectionnes = st.multiselect(
            "Enseignements retenus comme H.SUP (permanent uniquement)",
            options=options_hsup, default=deja_hsup, key="hsup_enseignements_selection"
        )
        if st.button("💾 Enregistrer mes enseignements H.SUP", key="hsup_save_teachings", use_container_width=True):
            ok_pref, erreur_pref = mettre_a_jour_metadonnees_compte({
                "hsup_enseignements_retenus": enseignements_hsup_selectionnes
            })
            if ok_pref:
                user_id_pref = (st.session_state.get("hsup_user") or {}).get("id")
                ok_table, erreur_table = sauvegarder_preferences_hsup(
                    user_id_pref, enseignant_select, enseignements_hsup_selectionnes
                )
                if ok_table:
                    st.success("Votre sélection H.SUP est enregistrée dans votre compte et dans la table Supabase partagée.")
                    st.rerun()
                else:
                    st.error("Métadonnées enregistrées, mais synchronisation de la table partagée impossible : " + erreur_table)
            else:
                st.error(f"Impossible d'enregistrer votre sélection : {erreur_pref}")
    if trouver_logo() is None:
        st.info(
            "logo.PNG introuvable à côté de app.py : le logo intégré de l'université est utilisé. "
            "Pour utiliser le vôtre, placez logo.PNG à la racine du dépôt GitHub."
        )
    st.download_button(
        label="📥 Télécharger l'emploi du temps individuel (PDF)",
        data=generer_pdf_edt([
            (enseignant_select, qualite_select, seances_enseignant_complet)
        ]),
        file_name=nom_fichier_edt(enseignant_select),
        mime="application/pdf",
        use_container_width=True,
        key="download_edt_individuel",
    )

    erreurs_horaires = seances_enseignant[
        seances_enseignant["DUREE_H"].isna()
    ] if "DUREE_H" in seances_enseignant.columns else pd.DataFrame()
    if not erreurs_horaires.empty:
        st.error("Certaines séances ont un horaire non interprétable. Corrige la source EDT.")

    if st.button("🚀 Générer l'état Excel officiel actualisé", type="primary",
                 use_container_width=True):
        if "permanent" in qualite_select.casefold():
            seances_pour_etat, charge_sem, hsup_eq = limiter_enseignements_au_depassement(
                seances_enseignant, enseignements_hsup_selectionnes, seuil=7.0
            )
            st.info(
                f"Charge réglementaire : 7 h éq-cours | Charge EDT : {charge_sem['total_eq']:.2f} h éq-cours "
                f"| Dépassement maximal : {max(0, charge_sem['total_eq']-7):.2f} h éq-cours "
                f"| H.SUP sélectionnées retenues : {hsup_eq:.2f} h éq-cours."
            )
            if seances_pour_etat.empty:
                st.warning("Aucune H.SUP à générer : sélection vide ou charge hebdomadaire ne dépassant pas 7 h équivalent-cours.")
                st.stop()
        else:
            seances_pour_etat = seances_enseignant.copy()
        buffer, total_general, details = generer_excel(
            enseignant_select, seances_pour_etat, dict_noms_arabe,
            date_debut_cours, date_fin_comptabilisation, exclusions_par_mois,
            qualite_select
        )
        nom_fichier_clean = normaliser_nom(enseignant_select).replace(" ", "_")
        file_name = f"Etat_HSUP_Enseignant_{nom_fichier_clean}_2026_2027.xlsx"

        details_individuel = construire_details_enseignants_historique(
            [enseignant_select],
            seances_par_enseignant={
                normaliser_nom(enseignant_select): seances_pour_etat
            },
        )

        ok_historique_individuel, erreur_historique_individuel = enregistrer_historique_generation(
            generation_type="État H.SUP individuel",
            enseignants=[enseignant_select],
            format_fichier="Excel individuel",
            nom_fichier=file_name,
            statut="Succès",
            message=(
                f"Total H.SUP calculé sur la période : "
                f"{format_heures_arabe(total_general) or '0سا'}."
            ),
            total_hsup=format_heures_arabe(total_general) or "0سا",
            enseignants_details=details_individuel,
        )
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
        st.markdown("#### Vérification du calcul jour par jour")
        st.dataframe(pd.DataFrame(details), use_container_width=True, hide_index=True)

    if est_administrateur:
        st.divider()
        st.header("🛠️ ESPACE ADMINISTRATION")
        st.success(
            "Accès administrateur confirmé. Vous pouvez préparer les emplois du temps "
            "et les états H.SUP de l'ensemble des enseignants."
        )
        st.subheader("📚 Téléchargement groupé de tous les enseignants")
        preferences_admin, exclus_admin, erreur_preferences = charger_preferences_hsup_admin()
        if erreur_preferences:
            st.error(
                "Impossible de lire les préférences partagées Supabase : " + erreur_preferences +
                " Configurez SUPABASE_SERVICE_ROLE_KEY et exécutez le script SQL fourni avec le code."
            )
        st.subheader("🚫 Exclure des enseignants des états H.SUP")
        noms_deja_exclus = [n for n in liste_enseignants_actifs if normaliser_nom(n) in exclus_admin]
        choix_exclusions = st.multiselect(
            "Enseignants à exclure de la génération des états H.SUP",
            options=liste_enseignants_actifs, default=noms_deja_exclus,
            key="hsup_admin_exclusions"
        )
        if st.button("💾 Enregistrer les exclusions administratives", key="hsup_admin_save_exclusions", use_container_width=True):
            ok_excl, err_excl = mettre_a_jour_exclusions_admin(choix_exclusions, df_enseignants)
            if ok_excl:
                st.success("Exclusions administratives enregistrées dans Supabase.")
                st.rerun()
            else:
                st.error("Enregistrement des exclusions impossible : " + err_excl)
        st.caption(
            "Les sélections d'enseignements de chaque permanent sont lues dans la table "
            "Supabase partagée. Un permanent sans sélection enregistrée n'aura aucune H.SUP. "
            "La table doit être créée avec le script SQL fourni et la clé service_role doit rester "
            "uniquement dans les Secrets Streamlit (jamais dans le dépôt GitHub)."
        )
        st.markdown(
            "Un seul classeur Excel multifeuille contenant un index cliquable, "
            "un récapitulatif des charges Cours/TD/TP et une fiche officielle par enseignant."
        )
        # ---- Emplois du temps PDF : un ou plusieurs enseignants, même ordre que l'Excel groupé ----
        st.markdown("#### 🗓️ Emplois du temps PDF (un ou plusieurs enseignants)")
        st.caption(
            "L'ordre est identique à celui du classeur Excel groupé : l'ordre alphabétique de la liste "
            "« Choisir un enseignant permanent ou vacataire »."
        )
        tous_edt = st.checkbox("Sélectionner tous les enseignants", key="edt_pdf_tous")
        choix_edt = st.multiselect(
            "Enseignants à inclure", options=liste_enseignants_actifs,
            default=[enseignant_select], disabled=tous_edt, key="edt_pdf_choix",
        )
        format_edt = st.radio(
            "Format du téléchargement",
            ["Un seul PDF (une page par enseignant)", "Un PDF par enseignant (fichier ZIP)"],
            horizontal=True, key="edt_pdf_format",
        )
        choisis = set(liste_enseignants_actifs) if tous_edt else set(choix_edt)
        # On parcourt la liste de référence : l'ordre est donc celui de l'Excel, quel que soit l'ordre de clic.
        selection_edt = [n for n in liste_enseignants_actifs if n in choisis]
        st.caption(f"{len(selection_edt)} enseignant(s) sélectionné(s).")
        if st.button("🗓️ Préparer le téléchargement des emplois du temps",
                     use_container_width=True, disabled=not selection_edt, key="edt_pdf_preparer"):
            with st.spinner("Génération des grilles d'emploi du temps..."):
                liste_edt = []
                for nom_edt in selection_edt:
                    n_norm = normaliser_nom(nom_edt)
                    seances_edt = df_edt_complet[
                        df_edt_complet["Enseignants"].apply(
                            lambda x: correspond_enseignant(x, n_norm)
                        )
                    ].copy()
                    if seances_edt.empty:
                        continue
                    ligne_edt = df_enseignants[df_enseignants["NOM_NORM"].eq(n_norm)]
                    qualite_edt = (nettoyer_texte(ligne_edt.iloc[0].get("Qualité", ""))
                                   if not ligne_edt.empty else "") or "Enseignant"
                    liste_edt.append((nom_edt, qualite_edt, seances_edt))
                if not liste_edt:
                    st.session_state.pop("edt_pdf_pret", None)
                    st.warning("Aucune séance EDT pour la sélection.")
                elif format_edt.startswith("Un PDF par"):
                    st.session_state["edt_pdf_pret"] = {
                        "data": generer_zip_edt(liste_edt).getvalue(),
                        "nom": "EDT_Tous_Enseignants_ELT_2027.zip", "mime": "application/zip",
                        "ordre": [e[0] for e in liste_edt],
                    }
                else:
                    st.session_state["edt_pdf_pret"] = {
                        "data": generer_pdf_edt(liste_edt).getvalue(),
                        "nom": (nom_fichier_edt(liste_edt[0][0]) if len(liste_edt) == 1
                                else "EDT_Tous_Enseignants_ELT_2027.pdf"),
                        "mime": "application/pdf",
                        "ordre": [e[0] for e in liste_edt],
                    }
        pret = st.session_state.get("edt_pdf_pret")
        if pret:
            st.success(f"{len(pret['ordre'])} emploi(s) du temps prêt(s) : " + " → ".join(pret["ordre"][:8])
                       + (" …" if len(pret["ordre"]) > 8 else ""))
            st.download_button(
                label=f"📥 Télécharger ({pret['nom']})",
                data=pret["data"], file_name=pret["nom"], mime=pret["mime"],
                use_container_width=True, key="download_edt_selection",
            )

        # ---- Etats des heures sups en Excel : un ou plusieurs enseignants ----
        st.markdown("#### 🗓️ Etats des heures sups en Excel (un ou plusieurs enseignants)")
        st.caption(
            "Générez les états H.SUP Excel pour n'importe quel enseignant "
            "présent dans l'EDT complet, inscrit ou non inscrit. Les enseignants "
            "exclus administrativement ne sont pas inclus. Pour les permanents, "
            "vous pouvez définir les enseignements retenus comme H.SUP et les "
            "jours à exclure. Avant de télécharger l'état, vous pouvez saisir "
            "le nom et le prénom en arabe de l'enseignant."
        )

        tous_hsup_excel = st.checkbox(
            "Sélectionner tous les enseignants éligibles",
            key="hsup_excel_tous",
        )

        # Tous les enseignants trouvés dans l'EDT complet, inscrits ou non.
        enseignants_edt_complet = extraire_noms_enseignants_edt(df_edt_complet)

        enseignants_hsup_eligibles = [
            n for n in enseignants_edt_complet
            if normaliser_nom(n) not in exclus_admin
        ]

        choix_hsup_excel = st.multiselect(
            "Enseignants à inclure dans les états H.SUP Excel",
            options=enseignants_hsup_eligibles,
            default=[enseignant_select] if enseignant_select in enseignants_hsup_eligibles else [],
            disabled=tous_hsup_excel,
            key="hsup_excel_choix",
        )

        format_hsup_excel = st.radio(
            "Format du téléchargement Excel",
            [
                "Un seul classeur Excel multifeuille",
                "Un fichier Excel par enseignant (fichier ZIP)",
            ],
            horizontal=True,
            key="hsup_excel_format",
        )

        choisis_hsup_excel = (
            set(enseignants_hsup_eligibles)
            if tous_hsup_excel
            else set(choix_hsup_excel)
        )

        selection_hsup_excel = [
            n for n in enseignants_edt_complet
            if n in choisis_hsup_excel and normaliser_nom(n) not in exclus_admin
        ]

        st.caption(
            f"{len(selection_hsup_excel)} enseignant(s) sélectionné(s) pour les états H.SUP Excel."
        )

        # Réglages administratifs : enseignements H.SUP, jours exclus et identité.
        reglages_admin_selections, reglages_admin_dates, reglages_admin_identites, erreur_reglages_admin = (
            charger_reglages_admin_enseignants()
        )

        if erreur_reglages_admin:
            st.error(
                "Impossible de charger les réglages administratifs : "
                + erreur_reglages_admin
            )

        dates_periode_admin = []
        if date_debut_cours and date_fin_comptabilisation:
            d = date_debut_cours
            while d <= date_fin_comptabilisation:
                dates_periode_admin.append(d)
                d += timedelta(days=1)

        for nom_reglage in selection_hsup_excel:
            nom_norm_reglage = normaliser_nom(nom_reglage)
            ligne_reglage = df_enseignants[
                df_enseignants["NOM_NORM"].eq(nom_norm_reglage)
            ]
            qualite_reglage = (
                nettoyer_texte(ligne_reglage.iloc[0].get("Qualité", ""))
                if not ligne_reglage.empty
                else ""
            ) or "Enseignant"

            identite_existante = reglages_admin_identites.get(
                nom_norm_reglage, {}
            )
            if not identite_existante.get("ism") and not identite_existante.get("laqab"):
                identite_existante = identite_arabe_de(
                    dict_noms_arabe, nom_reglage
                )
                if not isinstance(identite_existante, dict):
                    identite_existante = {}

            # Utiliser l'EDT complet : toutes les charges, tous départements.
            seances_reglage = df_edt_complet[
                df_edt_complet["Enseignants"].apply(
                    lambda x: correspond_enseignant(x, nom_norm_reglage)
                )
            ].copy()

            if seances_reglage.empty:
                continue

            # Un choix par enseignement ET par créneau (jour + horaire) distinct.
            options_enseignements_reglage = options_enseignements_hsup(seances_reglage)

            with st.expander(
                f"⚙️ Réglages H.SUP — {nom_reglage}",
                expanded=False,
            ):
                col_ism, col_laqab = st.columns(2)

                ism_reglage = col_ism.text_input(
                    "الاسم (Prénom en arabe)",
                    value=identite_existante.get("ism", ""),
                    key=f"hsup_admin_ism_{nom_norm_reglage}",
                )

                laqab_reglage = col_laqab.text_input(
                    "اللقب (Nom en arabe)",
                    value=identite_existante.get("laqab", ""),
                    key=f"hsup_admin_laqab_{nom_norm_reglage}",
                )

                if "permanent" in qualite_reglage.casefold():
                    selections_reglage = st.multiselect(
                        "Enseignements retenus comme H.SUP",
                        options=options_enseignements_reglage,
                        default=selections_vers_options_hsup(
                            reglages_admin_selections.get(nom_norm_reglage, []),
                            options_enseignements_reglage,
                        ),
                        help=(
                            "Un même enseignement placé sur deux créneaux différents "
                            "apparaît deux fois : choisissez le(s) créneau(x) à "
                            "comptabiliser en H.SUP."
                        ),
                        key=f"hsup_admin_enseignements_{nom_norm_reglage}",
                    )

                    dates_exclues_reglage = st.multiselect(
                        "Jours à exclure du calcul H.SUP",
                        options=dates_periode_admin,
                        default=reglages_admin_dates.get(
                            nom_norm_reglage, []
                        ),
                        format_func=lambda d: d.strftime("%d/%m/%Y"),
                        key=f"hsup_admin_dates_{nom_norm_reglage}",
                    )
                else:
                    selections_reglage = []
                    dates_exclues_reglage = []

                if st.button(
                    f"💾 Enregistrer les réglages de {nom_reglage}",
                    key=f"hsup_admin_save_{nom_norm_reglage}",
                    use_container_width=True,
                ):
                    ok_reglage, erreur_reglage = enregistrer_reglages_admin_enseignant(
                        nom_reglage,
                        selections_reglage,
                        dates_exclues_reglage,
                        ism_arabe=ism_reglage,
                        laqab_arabe=laqab_reglage,
                    )
                    if ok_reglage:
                        st.success(
                            f"Réglages enregistrés pour {nom_reglage}."
                        )
                        st.rerun()
                    else:
                        st.error(
                            "Enregistrement impossible : " + erreur_reglage
                        )

        if st.button(
            "🗓️ Préparer les états H.SUP Excel",
            use_container_width=True,
            disabled=not selection_hsup_excel,
            key="hsup_excel_preparer",
        ):
            with st.spinner("Génération des états H.SUP Excel..."):

                preferences_admin_hsup_excel, exclus_admin_hsup_excel, erreur_preferences_hsup_excel = (
                    charger_preferences_hsup_admin()
                )

                if erreur_preferences_hsup_excel:
                    st.session_state.pop("hsup_excel_pret", None)
                    st.error(
                        "Génération interrompue : préférences Supabase indisponibles. "
                        + erreur_preferences_hsup_excel
                    )
                else:
                    # Injecter les identités arabes saisies par l'administrateur.
                    dict_noms_arabe_generation = dict(dict_noms_arabe)
                    for nom_hsup_excel in selection_hsup_excel:
                        nom_norm_hsup_excel = normaliser_nom(nom_hsup_excel)
                        identite_admin = reglages_admin_identites.get(
                            nom_norm_hsup_excel, {}
                        )
                        if identite_admin.get("ism") or identite_admin.get("laqab"):
                            dict_noms_arabe_generation[nom_norm_hsup_excel] = {
                                "ism": identite_admin.get("ism", ""),
                                "laqab": identite_admin.get("laqab", ""),
                                "complet": " ".join(
                                    x for x in (
                                        identite_admin.get("ism", ""),
                                        identite_admin.get("laqab", ""),
                                    ) if x
                                ),
                            }

                    liste_hsup_excel = []
                    seances_par_enseignant_hsup_excel = {}

                    for nom_hsup_excel in selection_hsup_excel:
                        nom_norm_hsup_excel = normaliser_nom(nom_hsup_excel)

                        ligne_hsup_excel = df_enseignants[
                            df_enseignants["NOM_NORM"].eq(nom_norm_hsup_excel)
                        ]

                        # Un enseignant peut ne pas être encore inscrit ou ne pas
                        # figurer dans le fichier personnel : on génère quand même.
                        qualite_hsup_excel = (
                            nettoyer_texte(
                                ligne_hsup_excel.iloc[0].get("Qualité", "")
                            )
                            if not ligne_hsup_excel.empty
                            else ""
                        ) or "Enseignant"

                        # EDT complet : toutes les charges, tous départements.
                        seances_hsup_excel = df_edt_complet[
                            df_edt_complet["Enseignants"].apply(
                                lambda x: correspond_enseignant(x, nom_norm_hsup_excel)
                            )
                        ].copy()

                        if seances_hsup_excel.empty:
                            continue

                        if "permanent" in qualite_hsup_excel.casefold():
                            enseignements_retenus_hsup_excel = (
                                preferences_admin_hsup_excel.get(
                                    nom_norm_hsup_excel, []
                                )
                                or reglages_admin_selections.get(
                                    nom_norm_hsup_excel, []
                                )
                            )

                            # Si aucun enseignement n'est retenu, utiliser
                            # automatiquement tous les enseignements EDT.
                            if not enseignements_retenus_hsup_excel:
                                enseignements_retenus_hsup_excel = sorted(
                                    {
                                        nettoyer_texte(x)
                                        for x in seances_hsup_excel["Enseignements"]
                                        if nettoyer_texte(x)
                                    }
                                )

                            seances_pour_etat_hsup_excel, charge_semaine_hsup_excel, hsup_equivalent_hsup_excel = (
                                limiter_enseignements_au_depassement(
                                    seances_hsup_excel,
                                    enseignements_retenus_hsup_excel,
                                    seuil=7.0,
                                )
                            )
                        else:
                            seances_pour_etat_hsup_excel = seances_hsup_excel.copy()

                        if seances_pour_etat_hsup_excel.empty:
                            continue

                        seances_par_enseignant_hsup_excel[nom_norm_hsup_excel] = (
                            seances_pour_etat_hsup_excel
                        )

                        liste_hsup_excel.append(
                            (
                                nom_hsup_excel,
                                qualite_hsup_excel,
                                seances_pour_etat_hsup_excel,
                            )
                        )

                    if not liste_hsup_excel:
                        st.session_state.pop("hsup_excel_pret", None)
                        st.warning(
                            "Aucun état H.SUP Excel ne peut être généré pour la sélection. "
                            "Vérifiez que les enseignants sélectionnés ont des séances "
                            "dans l'EDT complet et que leurs noms correspondent."
                        )
                    elif format_hsup_excel.startswith("Un fichier Excel par enseignant"):

                        buffer_zip_hsup_excel = io.BytesIO()

                        with zipfile.ZipFile(
                            buffer_zip_hsup_excel,
                            mode="w",
                            compression=zipfile.ZIP_DEFLATED,
                        ) as archive_zip_hsup_excel:

                            for nom_fiche_hsup_excel, qualite_fiche_hsup_excel, seances_fiche_hsup_excel in liste_hsup_excel:

                                if "permanent" in qualite_fiche_hsup_excel.casefold():
                                    enseignements_retenus_fiche_hsup_excel = (
                                        preferences_admin_hsup_excel.get(
                                            normaliser_nom(nom_fiche_hsup_excel),
                                            [],
                                        )
                                        or reglages_admin_selections.get(
                                            normaliser_nom(nom_fiche_hsup_excel),
                                            [],
                                        )
                                    )

                                    if not enseignements_retenus_fiche_hsup_excel:
                                        enseignements_retenus_fiche_hsup_excel = sorted(
                                            {
                                                nettoyer_texte(x)
                                                for x in seances_fiche_hsup_excel["Enseignements"]
                                                if nettoyer_texte(x)
                                            }
                                        )

                                    seances_fiche_hsup_excel, _, _ = (
                                        limiter_enseignements_au_depassement(
                                            seances_fiche_hsup_excel,
                                            enseignements_retenus_fiche_hsup_excel,
                                            seuil=7.0,
                                        )
                                    )
                                else:
                                    seances_fiche_hsup_excel = seances_fiche_hsup_excel.copy()

                                if seances_fiche_hsup_excel.empty:
                                    continue

                                exclusions_fiche_hsup_excel = (
                                    construire_exclusions_par_mois_admin(
                                        reglages_admin_dates.get(
                                            normaliser_nom(nom_fiche_hsup_excel),
                                            [],
                                        )
                                    )
                                )

                                buffer_fiche_hsup_excel, total_fiche_hsup_excel, details_fiche_hsup_excel = (
                                    generer_excel(
                                        nom_fiche_hsup_excel,
                                        seances_fiche_hsup_excel,
                                        dict_noms_arabe_generation,
                                        date_debut_cours,
                                        date_fin_comptabilisation,
                                        exclusions_fiche_hsup_excel,
                                        qualite_fiche_hsup_excel,
                                    )
                                )

                                nom_fichier_hsup_excel = (
                                    "EtatHSUP_"
                                    + normaliser_nom(nom_fiche_hsup_excel).replace(" ", "_")
                                    + "_2026_2027.xlsx"
                                )

                                archive_zip_hsup_excel.writestr(
                                    nom_fichier_hsup_excel,
                                    buffer_fiche_hsup_excel.getvalue(),
                                )

                        buffer_zip_hsup_excel.seek(0)

                        details_hsup_excel = construire_details_enseignants_historique(
                            [e[0] for e in liste_hsup_excel],
                            seances_par_enseignant=seances_par_enseignant_hsup_excel,
                        )

                        st.session_state["hsup_excel_pret"] = {
                            "data": buffer_zip_hsup_excel.getvalue(),
                            "nom": "Etats_HSUP_Enseignants_ELT_2026_2027.zip",
                            "mime": "application/zip",
                            "ordre": [e[0] for e in liste_hsup_excel],
                        }

                        enregistrer_historique_generation(
                            generation_type="États H.SUP Excel - sélection multiple",
                            enseignants=[e[0] for e in liste_hsup_excel],
                            format_fichier="ZIP Excel",
                            nom_fichier="Etats_HSUP_Enseignants_ELT_2026_2027.zip",
                            statut="Succès",
                            message=f"{len(liste_hsup_excel)} fichier(s) Excel individuel(s) généré(s).",
                            enseignants_details=details_hsup_excel,
                        )

                    else:
                        details_hsup_excel = construire_details_enseignants_historique(
                            selection_hsup_excel,
                            seances_par_enseignant=seances_par_enseignant_hsup_excel,
                        )

                        exclusions_par_enseignant_hsup_excel = {
                            normaliser_nom(nom_hsup_excel): (
                                construire_exclusions_par_mois_admin(
                                    reglages_admin_dates.get(
                                        normaliser_nom(nom_hsup_excel),
                                        [],
                                    )
                                )
                            )
                            for nom_hsup_excel in selection_hsup_excel
                        }

                        # Sélections H.SUP effectives : préférences du compte
                        # ou réglages administratifs pour les non-inscrits.
                        selections_hsup_effectives_hsup_excel = {
                            normaliser_nom(nom_hsup_excel): (
                                preferences_admin_hsup_excel.get(
                                    normaliser_nom(nom_hsup_excel), []
                                )
                                or reglages_admin_selections.get(
                                    normaliser_nom(nom_hsup_excel), []
                                )
                            )
                            for nom_hsup_excel in selection_hsup_excel
                        }

                        buffer_hsup_excel, nombre_fiches_hsup_excel = generer_excel_groupe(
                            df_enseignants,
                            df_edt,
                            selection_hsup_excel,
                            dict_noms_arabe_generation,
                            date_debut_cours,
                            date_fin_comptabilisation,
                            exclusions_par_mois,
                            selections_hsup_par_enseignant=selections_hsup_effectives_hsup_excel,
                            exclusions_par_mois_par_enseignant=exclusions_par_enseignant_hsup_excel,
                            df_edt_source=df_edt_complet,
                        )

                        st.session_state["hsup_excel_pret"] = {
                            "data": buffer_hsup_excel.getvalue(),
                            "nom": "Etats_HSUP_Enseignants_Selectionnes_ELT_2026_2027.xlsx",
                            "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "ordre": selection_hsup_excel,
                            "nombre_fiches": nombre_fiches_hsup_excel,
                        }

                        enregistrer_historique_generation(
                            generation_type="États H.SUP Excel - sélection multiple",
                            enseignants=selection_hsup_excel,
                            format_fichier="Excel multifeuille",
                            nom_fichier="Etats_HSUP_Enseignants_Selectionnes_ELT_2026_2027.xlsx",
                            statut="Succès",
                            message=f"{nombre_fiches_hsup_excel} fiche(s) H.SUP générée(s).",
                            enseignants_details=details_hsup_excel,
                        )

        pret_hsup_excel = st.session_state.get("hsup_excel_pret")

        if pret_hsup_excel:
            st.success(
                f"{len(pret_hsup_excel['ordre'])} état(s) H.SUP Excel prêt(s) : "
                + " → ".join(pret_hsup_excel["ordre"][:8])
                + (" …" if len(pret_hsup_excel["ordre"]) > 8 else "")
            )

            st.download_button(
                label=f"📥 Télécharger les états H.SUP Excel ({pret_hsup_excel['nom']})",
                data=pret_hsup_excel["data"],
                file_name=pret_hsup_excel["nom"],
                mime=pret_hsup_excel["mime"],
                use_container_width=True,
                key="download_hsup_excel_selection",
            )

        if st.button("📦 Générer le classeur Excel de tous les enseignants",
                     type="primary", use_container_width=True):
            with st.spinner("Génération des fiches de tous les enseignants..."):
                preferences_admin, exclus_admin, erreur_preferences = charger_preferences_hsup_admin()
                if erreur_preferences:
                    st.error("Génération interrompue : préférences Supabase indisponibles. " + erreur_preferences)
                    st.stop()
                liste_a_generer = [n for n in liste_enseignants_actifs
                                   if normaliser_nom(n) not in exclus_admin]

                reglages_admin_selections_groupe, reglages_admin_dates_groupe, reglages_admin_identites_groupe, _ = (
                    charger_reglages_admin_enseignants()
                )

                dict_noms_arabe_groupe = dict(dict_noms_arabe)
                for nom_groupe in liste_a_generer:
                    nom_norm_groupe = normaliser_nom(nom_groupe)
                    identite_groupe = reglages_admin_identites_groupe.get(
                        nom_norm_groupe, {}
                    )
                    if identite_groupe.get("ism") or identite_groupe.get("laqab"):
                        dict_noms_arabe_groupe[nom_norm_groupe] = {
                            "ism": identite_groupe.get("ism", ""),
                            "laqab": identite_groupe.get("laqab", ""),
                            "complet": " ".join(
                                x for x in (
                                    identite_groupe.get("ism", ""),
                                    identite_groupe.get("laqab", ""),
                                ) if x
                            ),
                        }

                seances_par_enseignant_groupe = {}
                for nom_groupe in liste_a_generer:
                    nom_norm_groupe = normaliser_nom(nom_groupe)
                    seances_groupe = df_edt[
                        df_edt["Enseignants"].apply(
                            lambda x: correspond_enseignant(x, nom_norm_groupe)
                        )
                    ].copy()
                    if not seances_groupe.empty:
                        seances_par_enseignant_groupe[nom_norm_groupe] = seances_groupe

                details_groupe = construire_details_enseignants_historique(
                    liste_a_generer,
                    seances_par_enseignant=seances_par_enseignant_groupe,
                )

                selections_hsup_effectives = {}
                exclusions_par_enseignant_groupe = {}

                for nom_groupe in liste_a_generer:
                    nom_norm_groupe = normaliser_nom(nom_groupe)
                    selections_hsup_effectives[nom_norm_groupe] = (
                        preferences_admin.get(nom_norm_groupe, [])
                        or reglages_admin_selections_groupe.get(nom_norm_groupe, [])
                    )
                    exclusions_par_enseignant_groupe[nom_norm_groupe] = (
                        construire_exclusions_par_mois_admin(
                            reglages_admin_dates_groupe.get(nom_norm_groupe, [])
                        )
                    )

                buffer_groupe, nb_fiches = generer_excel_groupe(
                    df_enseignants, df_edt, liste_a_generer, dict_noms_arabe_groupe,
                    date_debut_cours, date_fin_comptabilisation, exclusions_par_mois,
                    selections_hsup_par_enseignant=selections_hsup_effectives,
                    exclusions_par_mois_par_enseignant=exclusions_par_enseignant_groupe,
                )

            st.success(f"Classeur groupé généré : {nb_fiches} fiches individuelles.")

            ok_historique_groupe, erreur_historique_groupe = enregistrer_historique_generation(
                generation_type="Classeur groupé - tous les enseignants",
                enseignants=liste_a_generer,
                format_fichier="Excel multifeuille",
                nom_fichier="Etats_HSUP_Tous_Enseignants_ELT_2026_2027.xlsx",
                statut="Succès",
                message=f"{nb_fiches} fiche(s) H.SUP individuelle(s) générée(s).",
                enseignants_details=details_groupe,
            )
            st.download_button(
                label="📥 Télécharger toutes les fiches H.SUP (Excel multifeuille)",
                data=buffer_groupe,
                file_name="Etats_HSUP_Tous_Enseignants_ELT_2026_2027.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_hsup_groupe",
            )

        # ---- Historique des générations des états H.SUP ----
        st.divider()
        st.markdown("#### 🕘 Historique des générations des états H.SUP")
        st.caption(
            "Cet historique enregistre chaque génération d'état H.SUP Excel : "
            "génération individuelle, classeur groupé de tous les enseignants ou "
            "génération Excel pour une sélection d'enseignants. "
            "Pour chaque enseignant, le nom en arabe, l'e-mail, le téléphone, "
            "la qualité, la charge du vacataire, la période et le total H.SUP "
            "sont enregistrés."
        )

        historique_generations, erreur_historique_generations = charger_historique_generations()

        if erreur_historique_generations:
            st.error(
                "Impossible de charger l'historique des générations : "
                + erreur_historique_generations
            )
        elif not historique_generations:
            st.info("Aucune génération d'état H.SUP n'a encore été enregistrée.")
        else:
            lignes_historique = []

            for ligne_historique in historique_generations:
                enseignants_historique = ligne_historique.get("teachers") or []
                if isinstance(enseignants_historique, str):
                    enseignants_historique = [enseignants_historique]

                details = ligne_historique.get("teacher_details") or []
                if not isinstance(details, list):
                    details = []

                details_par_nom = {}
                for detail in details:
                    if isinstance(detail, dict):
                        details_par_nom[
                            normaliser_nom(detail.get("nom", ""))
                        ] = detail

                for nom_historique in enseignants_historique:
                    detail = details_par_nom.get(
                        normaliser_nom(nom_historique), {}
                    )

                    lignes_historique.append({
                        "Date et heure": nettoyer_texte(ligne_historique.get("created_at", "")),
                        "Administrateur": nettoyer_texte(ligne_historique.get("generated_by", "")),
                        "Type de génération": nettoyer_texte(ligne_historique.get("generation_type", "")),
                        "Enseignants": nettoyer_texte(nom_historique),
                        "Nom et prénom en arabe": nettoyer_texte(detail.get("nom_arabe", "")),
                        "E-mail": nettoyer_texte(detail.get("email", "")),
                        "Téléphone": nettoyer_texte(detail.get("telephone", "")),
                        "Qualité": nettoyer_texte(detail.get("qualite", "")),
                        "Total charge (vacataire)": nettoyer_texte(detail.get("total_charge", "")),
                        "Début de période": nettoyer_texte(ligne_historique.get("period_start", "")),
                        "Fin de période": nettoyer_texte(ligne_historique.get("period_end", "")),
                        "Total H.SUP": nettoyer_texte(ligne_historique.get("total_hsup", "")),
                    })

            df_historique_affichage = pd.DataFrame(lignes_historique)

            st.dataframe(
                df_historique_affichage,
                use_container_width=True,
                hide_index=True,
            )

            buffer_historique_excel, df_historique_telechargeable = generer_excel_historique(
                historique_generations
            )

            st.download_button(
                label="📥 Télécharger l'historique des générations (Excel)",
                data=buffer_historique_excel,
                file_name="Historique_Generations_Etats_HSUP_ELT_2026_2027.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_historique_generations",
            )

    else:
        st.divider()
        st.info(
            "L'espace administration n'est pas activé pour ce compte. "
            "Pour donner l'accès administrateur, ajoutez l'adresse e-mail utilisée pour "
            "vous connecter dans HSUP_ADMIN_EMAILS dans le fichier .streamlit/secrets.toml, "
            "puis redémarrez Streamlit. Exemple : HSUP_ADMIN_EMAILS = \"admin@exemple.com\"."
        )
except Exception as e:
    st.error(f"Une erreur est survenue lors du traitement : {e}")
    st.exception(e)
