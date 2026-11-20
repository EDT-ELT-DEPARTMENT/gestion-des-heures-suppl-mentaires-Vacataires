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
FICHIER_TEMPLATE = BASE_DIR / "Etat_HSUP_Enseignant_2026_2027.xlsx"

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
                         date_debut, date_fin, exclusions_par_mois):
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
    ws_index.append(["Période de comptabilisation", f"Du {date_debut.strftime('%d/%m/%Y')} au {date_fin.strftime('%d/%m/%Y')}"])
    ws_index.append(["Enseignant", "Nom en arabe", "Total H.SUP", "Accès direct"])
    ws_recap.append(["RÉCAPITULATIF DES CHARGES D'ENSEIGNEMENT – ELT"])
    recap_headers = [
        "Enseignant", "Nom et prénom en arabe",
        "Nombre Cours", "Nombre TD", "Nombre TP",
        "Heures Cours / semaine", "Heures TD / semaine",
        "Heures TP / semaine", "Total heures EDT / semaine",
        "Total H.SUP période",
    ]
    mois_cols = [bloc["nom"] for bloc in MOIS_BLOCKS]
    ws_recap.append(recap_headers + mois_cols)

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
        total, details, heures_mois = remplir_fiche(
            ws_fiche, nom, seances, dict_arabe, date_debut, date_fin,
            exclusions_par_mois, qualite_fiche
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
    """Lignes de texte d'une case : enseignement (renvoyé à la ligne), lieu, promotion(s).

    Les séances identiques (même enseignement, même lieu) données à plusieurs
    promotions au même créneau sont regroupées : « M1CE, M1ER, M1ME ».
    """
    from reportlab.lib.utils import simpleSplit
    groupes = {}
    for _, seance in seances_case:
        cle = (_edt_texte(seance.get("Enseignements", "")), _edt_texte(seance.get("Lieu", "")))
        promo = _edt_texte(seance.get("Promotion", ""))
        groupes.setdefault(cle, [])
        if promo and promo not in groupes[cle]:
            groupes[cle].append(promo)
    lignes = []
    for (intitule, lieu), promos in groupes.items():
        lignes += simpleSplit(intitule, "Helvetica", EDT_TAILLE_CASE, largeur_utile) or [""]
        if lieu:
            lignes.append(lieu)
        if promos:
            lignes += simpleSplit(", ".join(promos), "Helvetica", EDT_TAILLE_CASE, largeur_utile)
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
    except Exception:
        url, key, admins = "", "", ""
    return url, key, {x.strip().casefold() for x in admins.split(",") if x.strip()}


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


def authentifier_enseignant(liste_enseignants):
    """Inscription/connexion. L'identité française est liée au référentiel EDT."""
    url, anon_key, admins = _supabase_config()
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
            mdp_inscription = st.text_input("Créer un mot de passe", type="password")
            mdp_confirmation = st.text_input("Confirmer le mot de passe", type="password")
            ism_inscription = st.text_input("الاسم (prénom en arabe) — facultatif")
            laqab_inscription = st.text_input("اللقب (nom en arabe) — facultatif")
            envoyer_inscription = st.form_submit_button(
                "Créer mon compte", type="primary", use_container_width=True
            )
        if envoyer_inscription:
            if not email_inscription.strip() or len(mdp_inscription) < 8:
                st.error("Indiquez un e-mail valide et un mot de passe d'au moins 8 caractères.")
            elif mdp_inscription != mdp_confirmation:
                st.error("Les deux mots de passe ne sont pas identiques.")
            else:
                metadata = {
                    "nom_enseignant": nom_inscription,
                    "ism_arabe": ism_inscription.strip(),
                    "laqab_arabe": laqab_inscription.strip(),
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
    url, anon_key, _ = _supabase_config()
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


def enregistrer_identite_arabe_compte(ism, laqab):
    """Enregistre les champs arabes dans les métadonnées du compte authentifié."""
    return mettre_a_jour_metadonnees_compte({"ism_arabe": ism, "laqab_arabe": laqab})


def lire_date_compte(valeur, defaut):
    """Convertit une date ISO stockée dans le compte, avec une valeur de repli sûre."""
    try:
        return date.fromisoformat(str(valeur)) if valeur else defaut
    except (TypeError, ValueError):
        return defaut

try:
    df_enseignants, df_edt, dict_noms_arabe, liste_enseignants_actifs = charger_donnees()
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
    if not nom_arabe:
        st.warning(
            "Le fichier source ne contient pas encore le nom et le prénom en arabe. "
            "Renseignez-les ci-dessous pour compléter votre fiche H.SUP."
        )
    with st.form("hsup_identite_arabe_form"):
        col_ism, col_laqab = st.columns(2)
        ism_saisi = col_ism.text_input("الاسم — prénom en arabe", value=ism_arabe)
        laqab_saisi = col_laqab.text_input("اللقب — nom en arabe", value=laqab_arabe)
        sauvegarder_arabe = st.form_submit_button("Enregistrer mes renseignements en arabe")
    if sauvegarder_arabe:
        ok, erreur = enregistrer_identite_arabe_compte(ism_saisi.strip(), laqab_saisi.strip())
        if ok:
            st.success("Vos renseignements en arabe ont été enregistrés dans votre compte.")
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
    if trouver_logo() is None:
        st.info(
            "logo.PNG introuvable à côté de app.py : le logo intégré de l'université est utilisé. "
            "Pour utiliser le vôtre, placez logo.PNG à la racine du dépôt GitHub."
        )
    st.download_button(
        label="📥 Télécharger l'emploi du temps individuel (PDF)",
        data=generer_pdf_edt([(enseignant_select, qualite_select, seances_enseignant)]),
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
        buffer, total_general, details = generer_excel(
            enseignant_select, seances_enseignant, dict_noms_arabe,
            date_debut_cours, date_fin_comptabilisation, exclusions_par_mois,
            qualite_select
        )
        nom_fichier_clean = normaliser_nom(enseignant_select).replace(" ", "_")
        file_name = f"Etat_HSUP_Enseignant_{nom_fichier_clean}_2026_2027.xlsx"
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
        st.subheader("📚 Administration — Téléchargement groupé de tous les enseignants")
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
                    seances_edt = df_edt[
                        df_edt["Enseignants"].apply(lambda x: correspond_enseignant(x, n_norm))
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

        if st.button("📦 Générer le classeur Excel de tous les enseignants",
                     type="primary", use_container_width=True):
            with st.spinner("Génération des fiches de tous les enseignants..."):
                buffer_groupe, nb_fiches = generer_excel_groupe(
                    df_enseignants, df_edt, liste_enseignants_actifs, dict_noms_arabe,
                    date_debut_cours, date_fin_comptabilisation, exclusions_par_mois
                )
            st.success(f"Classeur groupé généré : {nb_fiches} fiches individuelles.")
            st.download_button(
                label="📥 Télécharger toutes les fiches H.SUP (Excel multifeuille)",
                data=buffer_groupe,
                file_name="Etats_HSUP_Tous_Enseignants_ELT_2026_2027.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_hsup_groupe",
            )
except Exception as e:
    st.error(f"Une erreur est survenue lors du traitement : {e}")
    st.exception(e)
