# -*- coding: utf-8 -*-
"""
================================================================================
PROGRAMME : Préparation des fichiers de réinterrogation (Passage 2) — Interface
================================================================================
AUTEUR    : mg.kouame
DATE      : Fév 2026
VERSION   : 3.0 — Interface graphique à onglets

DÉPENDANCES (installer une seule fois) :
    pip install pandas numpy openpyxl pyreadstat

LANCEMENT :
    python reinterrogation_ENEM.py
================================================================================
⚠️  IMPORTANT : Cette application couvre les trimestres jusqu'au T4_2030.
    Au-delà, ouvrir le code et ajouter les entrées manquantes dans le
    dictionnaire NOMS_FICHIERS (section "CORRESPONDANCE TRIMESTRE → NOM FICHIER").
================================================================================
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import re
import threading
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════════════
#  CORRESPONDANCE TRIMESTRE → NOM FICHIER  (étendre ici si besoin après 2030)
# ═══════════════════════════════════════════════════════════════════════════════
NOMS_FICHIERS = {
    "T2_2024": "ENEM_2024T2.dta",
    "T3_2024": "ENEM_2024T3.dta",
    "T4_2024": "ENEM_2024T4.dta",
    "T1_2025": "ENEM_2025T1.dta",
    "T2_2025": "ENEM_2025T2.dta",
    "T3_2025": "ENEM_2025T3.dta",
    "T4_2025": "ENEM_2025T4.dta",
    "T1_2026": "ENEM_2026T1.dta",
    "T2_2026": "ENEM_2026T2.dta",
    "T3_2026": "ENEM_2026T3.dta",
    "T4_2026": "ENEM_2026T4.dta",
    "T1_2027": "ENEM_2027T1.dta",
    "T2_2027": "ENEM_2027T2.dta",
    "T3_2027": "ENEM_2027T3.dta",
    "T4_2027": "ENEM_2027T4.dta",
    "T1_2028": "ENEM_2028T1.dta",
    "T2_2028": "ENEM_2028T2.dta",
    "T3_2028": "ENEM_2028T3.dta",
    "T4_2028": "ENEM_2028T4.dta",
    "T1_2029": "ENEM_2029T1.dta",
    "T2_2029": "ENEM_2029T2.dta",
    "T3_2029": "ENEM_2029T3.dta",
    "T4_2029": "ENEM_2029T4.dta",
    "T1_2030": "ENEM_2030T1.dta",
    "T2_2030": "ENEM_2030T2.dta",
    "T3_2030": "ENEM_2030T3.dta",
    "T4_2030": "ENEM_2030T4.dta",
}

# Format valide d'un trimestre
REGEX_TRIMESTRE = re.compile(r'^T[1-4]_\d{4}$')

# Palette de couleurs
CLR_BLEU_MARINE  = '#172F6F'
CLR_BLEU_ROI     = '#3F7FE4'
CLR_BLEU_CLAIR   = '#E7EFF6'
CLR_BLANC        = '#FFFFFF'
CLR_FOND         = '#F0F4F8'
CLR_TEXTE        = '#1E1E2E'
CLR_VERT         = '#2E7D32'
CLR_ROUGE        = '#C62828'
CLR_GRIS         = '#888888'


# ═══════════════════════════════════════════════════════════════════════════════
#  LOGIQUE MÉTIER — TRAITEMENT DES DONNÉES
# ═══════════════════════════════════════════════════════════════════════════════

def run_traitement(params: dict, log, progress):
    """
    Exécute le traitement complet de préparation des fichiers de réinterrogation.
    params  : dictionnaire de tous les paramètres saisis dans l'interface
    log     : callable(str) pour écrire dans le journal
    progress: callable(int) pour mettre à jour la barre (0-100)
    """
    try:
        import pandas as pd
        import numpy as np
    except ImportError as e:
        log(f"❌ Module manquant : {e}")
        log("   Exécutez : pip install pandas numpy openpyxl pyreadstat")
        return False

    # ── Décompacter les paramètres ─────────────────────────────
    TRIMESTRE_ACTUEL        = params['trimestre_actuel']
    ANNEE_ACTUELLE          = params['annee_actuelle']
    TRIMESTRE_NUMERO        = params['trimestre_numero']
    MOIS_EN_COURS           = params['mois_en_cours']
    COHORTES_REINTERROGATION = params['cohortes']
    RANGS_PAR_COHORTE       = params['rangs']
    REPERTOIRE_BASE         = params['repertoire_base']
    DOSSIER_SORTIE          = params['dossier_sortie']
    FICHIER_SEMAINE_REF     = params['fichier_semaine_ref']

    log("=" * 60)
    log("PROGRAMME DE PRÉPARATION DES FICHIERS DE RÉINTERROGATION")
    log("=" * 60)
    log(f"📅 Trimestre actuel    : {TRIMESTRE_ACTUEL}")
    log(f"📅 Année               : {ANNEE_ACTUELLE} — Trimestre n°{TRIMESTRE_NUMERO}")
    log(f"📅 Mois de début       : {MOIS_EN_COURS}")
    log(f"🔄 Cohortes            : {', '.join(COHORTES_REINTERROGATION)}")
    progress(2)

    # ── Créer le dossier de sortie ────────────────────────────
    os.makedirs(DOSSIER_SORTIE, exist_ok=True)
    log(f"\n📁 Dossier de sortie   : {DOSSIER_SORTIE}")
    progress(4)

    # ── Fichier de référence des semaines ─────────────────────
    log("\n📅 Chargement du fichier de référence des semaines...")
    try:
        df_semaine_ref = pd.read_excel(FICHIER_SEMAINE_REF, sheet_name='Semaine_ref_trim')
        log(f"   ✓ {len(df_semaine_ref)} enregistrements chargés")
        for col in ['Trimestre', 'Numero_semaine', 'DateJ7', 'Date1', 'Date2']:
            if col not in df_semaine_ref.columns:
                log(f"   ❌ Colonne manquante : {col}")
                return False
    except Exception as e:
        log(f"   ❌ Erreur : {e}")
        return False
    progress(8)

    # ── Labels géographiques ──────────────────────────────────
    log("\n🏷️  Chargement des labels géographiques...")
    dict_labels = {}
    feuilles_labels = {
        'label_region'     : ('HH2', 'label_HH2'),
        'label_district'   : ('HH1', 'label_HH1'),
        'label_departement': ('HH3', 'label_HH3'),
        'label_sp'         : ('HH4', 'label_HH4'),
    }
    for nom_f, (col_code, col_label) in feuilles_labels.items():
        try:
            df_lbl = pd.read_excel(FICHIER_SEMAINE_REF, sheet_name=nom_f)
            if col_code in df_lbl.columns and col_label in df_lbl.columns:
                df_lbl = df_lbl.dropna(subset=[col_code]).drop_duplicates(subset=[col_code])
                dict_labels[col_code] = dict(zip(df_lbl[col_code], df_lbl[col_label]))
                log(f"   ✓ {nom_f} : {len(dict_labels[col_code])} correspondances")
            else:
                log(f"   ⚠️  {nom_f} : colonnes manquantes")
        except Exception as e:
            log(f"   ⚠️  {nom_f} : {e}")
    progress(12)

    # ── Chargement des cohortes ───────────────────────────────
    log("\n" + "=" * 60)
    log("CHARGEMENT DES COHORTES")
    log("=" * 60)

    liste_menages = []
    liste_membres = []
    nb_total = len(COHORTES_REINTERROGATION)

    for idx_c, cohorte in enumerate(COHORTES_REINTERROGATION):
        log(f"\n📥 Cohorte : {cohorte}")
        dossier        = os.path.join(REPERTOIRE_BASE, f"Base_brute_{cohorte}")
        fichier_menage = os.path.join(dossier, NOMS_FICHIERS[cohorte])
        fichier_membres = os.path.join(dossier, "membres.dta")

        if not os.path.exists(fichier_menage):
            log(f"   ❌ Introuvable : {fichier_menage}")
            continue
        if not os.path.exists(fichier_membres):
            log(f"   ❌ Introuvable : {fichier_membres}")
            continue

        try:
            menage  = pd.read_stata(fichier_menage,  convert_categoricals=False,
                                    convert_missing=False, preserve_dtypes=False)
            membres = pd.read_stata(fichier_membres, convert_categoricals=False,
                                    convert_missing=False, preserve_dtypes=False)
            menage['cohorte_origine']  = cohorte
            membres['cohorte_origine'] = cohorte
            liste_menages.append(menage)
            liste_membres.append(membres)
            log(f"   ✓ {len(menage)} ménages | {len(membres)} membres")
        except Exception as e:
            log(f"   ❌ Erreur lecture : {e}")
            continue

        progress(12 + int((idx_c + 1) / nb_total * 15))

    if not liste_menages:
        log("\n❌ Aucune cohorte chargée. Vérifiez les chemins et noms de sous-dossiers.")
        return False

    Menage  = pd.concat(liste_menages, ignore_index=True)
    Membres = pd.concat(liste_membres, ignore_index=True)
    log(f"\n✓ Total consolidé : {len(Menage)} ménages | {len(Membres)} membres")
    progress(28)

    # ── Rangs d'interrogation — ménages ──────────────────────
    log("\n🔢 Attribution des rangs — ménages...")
    Menage['rgmen'] = None
    Menage['rghab'] = None
    Menage['rang_last_trim'] = None

    for cohorte, rang_val in RANGS_PAR_COHORTE.items():
        mask = Menage['cohorte_origine'] == cohorte
        Menage.loc[mask, 'rgmen']          = rang_val
        Menage.loc[mask, 'rghab']          = rang_val
        Menage.loc[mask, 'rang_last_trim'] = rang_val - 1
        log(f"   ✓ {cohorte} : rang={rang_val} ({mask.sum()} ménages)")
    progress(32)

    # ── Métadonnées Survey Solutions ─────────────────────────
    Menage['_responsible'] = 'AgentReinterrogation_' + TRIMESTRE_ACTUEL
    Menage['_quantity']    = 1

    # ── Clé d'identification unique ───────────────────────────
    log("\n🔑 Création des clés d'identification...")
    Menage['V1interviewkey1er'] = (
        Menage['HH4'].astype(str) + "_" +
        Menage['HH8'].astype(str) + "_" +
        Menage['HH7'].astype(str) +
        Menage['HH7B'].astype(str) + 'T' +
        Menage['trimestreencours'].astype(str) +
        Menage['annee'].astype(str) +
        Menage['rghab'].astype(str) + "_" +
        Menage['HH9_1'].astype(str)
    )
    log(f"   ✓ {len(Menage)} clés créées")
    progress(36)

    # ── Fusion membres + ménage ───────────────────────────────
    log("\n🔗 Fusion membres ↔ ménages...")
    MembresVF = pd.merge(Membres, Menage, on='interview__key', how='left')
    log(f"   ✓ {len(MembresVF)} lignes après fusion")
    progress(40)

    # ── Variables temporelles ─────────────────────────────────
    log("\n📅 Mise à jour des variables temporelles...")
    Menage['trimestreencours']       = TRIMESTRE_NUMERO
    Menage['mois_en_cours']          = MOIS_EN_COURS
    Menage['annee']                  = ANNEE_ACTUELLE
    Menage['V1interviewkey']         = Menage['interview__key']
    Menage['V1interviewkey_nextTrim'] = Menage['interview__key']
    progress(42)

    # ── Préchargement variables Passage 1 — ménage ───────────
    log("\n💾 Préchargement variables Passage 1 — ménages...")
    vars_v1_menage = {
        'V1hha': 'hha', 'V1Q2': 'Q2', 'V1Q2_aut': 'Q2_aut',
        'V1GPS_longitude': 'GPS__Longitude', 'V1GPS_Lattitude': 'GPS__Latitude',
        'V1nom_prenom_cm': 'nom_prenom_cm',
        'V1HH10_1': 'HH10_1', 'V1HH10_2': 'HH10_2',
        'V1HH9_1': 'HH9_1', 'V1HH9': 'HH9', 'V1Q1_0': 'Q1_0',
        'V1HH13A': 'HH13A', 'V1HH10_1_1a': 'HH10_1_1a',
        'V1HH10_2_1': 'HH10_2_1', 'V1HH13B': 'HH13B',
    }
    for dest, src in vars_v1_menage.items():
        if src in Menage.columns:
            Menage[dest] = Menage[src]
        else:
            Menage[dest] = None
    progress(45)

    # ── Labels géographiques ──────────────────────────────────
    log("\n🏷️  Ajout des labels géographiques...")
    for var in ['HH1', 'HH2', 'HH3', 'HH4']:
        nom_label = f"{var}_label"
        if var in Menage.columns and var in dict_labels:
            Menage[nom_label] = Menage[var].map(dict_labels[var])
            log(f"   ✓ {nom_label} : {Menage[nom_label].notna().sum()}/{len(Menage)} trouvés")
        else:
            Menage[nom_label] = None
    progress(48)

    # ── Variable DateJ7 et semaines ───────────────────────────
    log("\n📅 Détermination des semaines de référence...")
    Menage['Semaine_ref'] = None
    Menage['Date1']       = None
    Menage['Date2']       = None

    if 'DateJ7' in Menage.columns:
        Menage['V1DateJ7'] = Menage['DateJ7']
        for idx, row in Menage.iterrows():
            corresp = df_semaine_ref[
                (df_semaine_ref['Trimestre'] == row['cohorte_origine']) &
                (df_semaine_ref['DateJ7']    == row['DateJ7'])
            ]
            if len(corresp) > 0:
                sem = corresp.iloc[0]['Numero_semaine']
                Menage.at[idx, 'Semaine_ref'] = sem
                dates = df_semaine_ref[
                    (df_semaine_ref['Trimestre']      == TRIMESTRE_ACTUEL) &
                    (df_semaine_ref['Numero_semaine'] == sem)
                ]
                if len(dates) > 0:
                    d1 = dates.iloc[0]['Date1']
                    d2 = dates.iloc[0]['Date2']
                    if pd.notna(d1):
                        Menage.at[idx, 'Date1'] = str(d1).replace('/', '-')
                    if pd.notna(d2):
                        Menage.at[idx, 'Date2'] = str(d2).replace('/', '-')
        n_sem = Menage['Semaine_ref'].notna().sum()
        log(f"   ✓ Semaines déterminées : {n_sem}/{len(Menage)} ménages")
    else:
        log("   ⚠️  Variable 'DateJ7' absente — semaines non calculées")
        Menage['V1DateJ7'] = None
    progress(60)

    # ── Variables ord_sem et HH01 ─────────────────────────────
    log("\n🔢 Création de ord_sem et HH01...")
    import numpy as np
    if Menage['Semaine_ref'].notna().any():
        np.random.seed(42)
        keys_uniques = Menage['interview__key'].unique()
        dict_aleat = {k: np.random.randint(10000000, 100000000) for k in keys_uniques}
        Menage['_var_aleat_tmp'] = Menage['interview__key'].map(dict_aleat)
        Menage['ord_sem'] = (
            "Tele_" + Menage['Semaine_ref'].astype(str) +
            f"_{TRIMESTRE_ACTUEL}_" + Menage['_var_aleat_tmp'].astype(str)
        )
        Menage['HH01'] = (
            Menage['HH8A'].astype(str) + "_" + Menage['HH8'].astype(str) + "_" +
            Menage['Semaine_ref'].astype(str) + f"_{TRIMESTRE_ACTUEL}_" +
            Menage['_var_aleat_tmp'].astype(str)
        )
        Menage.drop(columns=['_var_aleat_tmp'], inplace=True)
        log(f"   ✓ ord_sem et HH01 créées pour {len(Menage)} ménages")
    else:
        Menage['ord_sem'] = ""
        Menage['HH01']    = ""
        log("   ⚠️  Semaine_ref absente — ord_sem et HH01 vides")
    progress(65)

    # ── Fichier ménage final ──────────────────────────────────
    log("\n📋 Construction du fichier ménage final...")
    colonnes_menage = [
        'HH1_label','HH2_label','HH3_label','HH4_label','Semaine_ref',
        'interview__id','Cohorte','ord_sem','HH01','HH0','HH2A',
        'HH1','HH2','HH3','HH4','HH6','HH8','HH8A','HH7','HH7B','HH8B',
        'rghab','rgmen',
        'V1MODINTR','trimestreencours','mois_en_cours','annee',
        'Date1','Date2','Reference',
        'V1interviewkey','V1interviewkey_nextTrim','V1interviewkey1er','V1hha',
        'V1Q2','V1Q2_aut','V1GPS_longitude','V1GPS_Lattitude','V1nom_prenom_cm',
        'V1HH10_1','V1HH10_2','V1HH9_1','V1HH9','V1Q1_0',
        'V1HH13A','V1HH10_1_1a','V1HH10_2_1','V1HH13B',
    ] + [f'M0__{i}' for i in range(60)] + [
        '_responsible','_quantity','GPS__Longitude','GPS__Latitude',
        'interview__key','hh','hha','cohorte_origine',
    ]

    cols_ok = [c for c in colonnes_menage if c in Menage.columns]
    MenageVF = Menage[cols_ok].copy()
    MenageVF['Cohorte'] = MenageVF['cohorte_origine']

    f_xlsx = os.path.join(DOSSIER_SORTIE, "QX_EEC_VF.xlsx")
    f_csv  = os.path.join(DOSSIER_SORTIE, "QX_EEC_VF.csv")
    MenageVF.to_excel(f_xlsx, index=False)
    MenageVF.to_csv(f_csv,   index=False)
    log(f"   ✓ QX_EEC_VF.xlsx ({len(MenageVF)} ménages, {len(cols_ok)} colonnes)")
    progress(72)

    # ── Fichier membres ───────────────────────────────────────
    log("\n👥 Préparation du fichier membres...")

    # Nettoyer colonnes dupliquées après merge
    for col_base in ['interview__id','cohorte_origine','V1interviewkey1er',
                     'rgmen','rghab','rang_last_trim']:
        cx, cy = f"{col_base}_x", f"{col_base}_y"
        if cx in MembresVF.columns and cy in MembresVF.columns:
            MembresVF[col_base] = MembresVF[cy].fillna(MembresVF[cx])
            MembresVF.drop(columns=[cx, cy], inplace=True)
        elif cx in MembresVF.columns:
            MembresVF.rename(columns={cx: col_base}, inplace=True)
        elif cy in MembresVF.columns:
            MembresVF.rename(columns={cy: col_base}, inplace=True)

    # Clé individu
    if all(c in MembresVF.columns for c in ['V1interviewkey1er','rgmen','membres__id']):
        MembresVF['cle_individu'] = (
            MembresVF['V1interviewkey1er'].astype(str) +
            MembresVF['rgmen'].astype(str) + "1_" +
            MembresVF['membres__id'].astype(str)
        )

    # Rangs individus
    MembresVF['rang_ind'] = None
    for cohorte, rang_val in RANGS_PAR_COHORTE.items():
        mask = MembresVF['cohorte_origine'] == cohorte
        MembresVF.loc[mask, 'rang_ind'] = rang_val
    MembresVF['rang_last_trim'] = MembresVF['rang_ind'] - 1
    MembresVF['membre_id_v1']     = MembresVF['membres__id']
    MembresVF['rangind_1er']      = MembresVF['membres__id']
    MembresVF['membre_id_v1_IND'] = MembresVF['rang_ind'].astype(str) + "_ieme_interrogation"

    # Préchargement variables individuelles
    vars_v1_membres = {
        'V1M6_J':'M6_J','V1M6_M':'M6_M','V1M6_A':'M6_A','V1M7':'M7',
        'V1M4':'M4','V1M9':'M9','V1M12':'M12',
        'V1EF1':'EF1','V1FP1':'FP1','V1EP1a':'EP1a',
    }
    for dest, src in vars_v1_membres.items():
        MembresVF[dest] = MembresVF[src] if src in MembresVF.columns else None

    vars_q1 = ['Q1_01','Q1_1','Q1_4','Q1_7','Q1_9',
               'Q1_10__1','Q1_10__2','Q1_10__3','Q1_10__4',
               'Q1_12','Q1_13__1','Q1_13__2','Q1_13__3','Q1_13__4']
    for var in vars_q1:
        if var in MembresVF.columns:
            MembresVF[f'V1{var}'] = MembresVF[var]

    # Filtrage résidents
    keys_valides = set(MenageVF['interview__key'].dropna())
    MembresVF = MembresVF[MembresVF['interview__key'].isin(keys_valides)]
    if 'Statut_Res' in MembresVF.columns:
        MembresVF = MembresVF[MembresVF['Statut_Res'] == 1]
        log(f"   ✓ Résidents filtrés : {len(MembresVF)} individus")

    # Renommer cohorte_origine → Cohorte1
    if 'Cohorte1' in MembresVF.columns:
        MembresVF.drop(columns=['Cohorte1'], inplace=True)
    if 'cohorte_origine' in MembresVF.columns:
        MembresVF.rename(columns={'cohorte_origine': 'Cohorte1'}, inplace=True)

    colonnes_membres = [
        'membres__id','M0','Cohorte1',
        'membre_id_v1','rangind_1er','rang_last_trim','cle_individu','rang_ind',
        'V1M4','V1M9','V1M12','membre_id_v1_IND',
        'V1M6_J','V1M6_M','V1M6_A','V1M7',
        'V1Q1_01','V1Q1_1','V1Q1_4','V1Q1_7','V1Q1_9',
        'V1Q1_10__1','V1Q1_10__2','V1Q1_10__3','V1Q1_10__4',
        'V1Q1_12','V1Q1_13__1','V1Q1_13__2','V1Q1_13__3','V1Q1_13__4','V1EP1a',
        'interview__id',
        'AgeAnnee','hhb','hha_FT','hha_SE','hha_EMP','hha_ES','hha_PL',
        'hhavf_C','hha_P','M4Confirm','EN_EMP',
        'membre_id_v1A','membre_id_v1_INDA','statut_MO','cle_individuA',
        'V1interviewkey','V1interviewkey_nextTrim','V1interviewkey1er',
        'Statut_Res','hha_COMP',
    ]
    cols_m_ok = [c for c in colonnes_membres if c in MembresVF.columns]
    MembresVF = MembresVF[cols_m_ok].copy()

    if 'interview__id' in MembresVF.columns:
        MembresVF = MembresVF.sort_values(by='membres__id')
        MembresVF['numero'] = MembresVF.groupby('interview__id').cumcount() + 1

    fm_xlsx = os.path.join(DOSSIER_SORTIE, "membres.xlsx")
    fm_csv  = os.path.join(DOSSIER_SORTIE, "membres.csv")
    MembresVF.to_excel(fm_xlsx, index=False)
    MembresVF.to_csv(fm_csv,   index=False)
    log(f"   ✓ membres.xlsx ({len(MembresVF)} individus, {len(cols_m_ok)} colonnes)")
    progress(95)

    # ── Résumé final ──────────────────────────────────────────
    log("\n" + "=" * 60)
    log("✅ TRAITEMENT TERMINÉ AVEC SUCCÈS")
    log("=" * 60)
    log(f"\n📊 Résultats :")
    log(f"   • Ménages    : {len(MenageVF)}")
    log(f"   • Résidents  : {len(MembresVF)}")
    log(f"\n📁 Fichiers dans : {DOSSIER_SORTIE}")
    log(f"   ✓ QX_EEC_VF.xlsx / QX_EEC_VF.csv")
    log(f"   ✓ membres.xlsx   / membres.csv")
    progress(100)
    return True


# ═══════════════════════════════════════════════════════════════════════════════
#  INTERFACE GRAPHIQUE
# ═══════════════════════════════════════════════════════════════════════════════

class AppReinterrogation(tk.Tk):

    MAX_COHORTES = 4

    def __init__(self):
        super().__init__()
        self.title("ENEM — Préparation Réinterrogation Passage 2")
        self.configure(bg=CLR_FOND)
        self.resizable(True, True)
        self.minsize(820, 640)

        # ── Variables de saisie ────────────────────────────────
        self.v_trimestre_actuel  = tk.StringVar(value="T1_2026")
        self.v_annee             = tk.StringVar(value="2026")
        self.v_trim_numero       = tk.StringVar(value="1")
        self.v_mois_en_cours     = tk.StringVar(value="1")
        self.v_nb_cohortes       = tk.IntVar(value=3)
        self.v_repertoire_base   = tk.StringVar()
        self.v_dossier_parent    = tk.StringVar()
        self.v_dossier_sortie    = tk.StringVar()
        self.v_fichier_semaine   = tk.StringVar()

        # Listes dynamiques pour les cohortes et rangs
        self.cohorte_vars = [tk.StringVar() for _ in range(self.MAX_COHORTES)]
        self.rang_vars    = [tk.StringVar(value="1") for _ in range(self.MAX_COHORTES)]

        self._build_ui()
        self._on_trim_actuel_change()  # pré-remplir dossier sortie

    # ── En-tête ────────────────────────────────────────────────
    def _header(self):
        hdr = tk.Frame(self, bg=CLR_BLEU_MARINE)
        hdr.pack(fill='x')
        tk.Label(
            hdr,
            text="  📊  ENEM — Préparation des fichiers de réinterrogation (Passage 2)",
            font=('Calibri', 13, 'bold'), bg=CLR_BLEU_MARINE, fg=CLR_BLANC, anchor='w'
        ).pack(side='left', padx=16, pady=10)

        # Bandeau d'avertissement années
        warn = tk.Frame(self, bg='#FFF8E1')
        warn.pack(fill='x')
        tk.Label(
            warn,
            text=("⚠️  Cette application couvre les trimestres jusqu'au T4_2030.  "
                  "Au-delà, ouvrez le fichier .py et ajoutez les entrées dans le dictionnaire NOMS_FICHIERS."),
            font=('Calibri', 9), bg='#FFF8E1', fg='#5D4037', anchor='w'
        ).pack(side='left', padx=12, pady=4)

    # ── Corps principal avec onglets ───────────────────────────
    def _build_ui(self):
        self._header()

        # Notebook (onglets)
        style = ttk.Style()
        style.configure('TNotebook.Tab', font=('Calibri', 10, 'bold'), padding=[12, 6])
        style.configure('TNotebook', background=CLR_FOND)

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=12, pady=8)

        # Créer les 4 onglets
        self.tab_params   = self._scrollable_frame(nb)
        self.tab_cohortes = self._scrollable_frame(nb)
        self.tab_chemins  = self._scrollable_frame(nb)
        self.tab_aide     = self._scrollable_frame(nb)

        nb.add(self.tab_params['outer'],   text="⚙️  Paramètres généraux")
        nb.add(self.tab_cohortes['outer'], text="🔄  Cohortes & Rangs")
        nb.add(self.tab_chemins['outer'],  text="📁  Chemins")
        nb.add(self.tab_aide['outer'],     text="📖  Aide & Mode d'emploi")

        self._build_tab_params(self.tab_params['inner'])
        self._build_tab_cohortes(self.tab_cohortes['inner'])
        self._build_tab_chemins(self.tab_chemins['inner'])
        self._build_tab_aide(self.tab_aide['inner'])

        # Zone de lancement + log
        self._build_bottom()

    def _scrollable_frame(self, parent):
        """Crée un onglet avec scroll vertical."""
        outer = tk.Frame(parent, bg=CLR_FOND)
        canvas = tk.Canvas(outer, bg=CLR_FOND, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=CLR_FOND)
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        canvas.bind_all('<MouseWheel>',
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), 'units'))
        return {'outer': outer, 'inner': inner}

    # ══════════════════════════════════════════════════════════
    #  ONGLET 1 : PARAMÈTRES GÉNÉRAUX
    # ══════════════════════════════════════════════════════════
    def _build_tab_params(self, f):
        PAD = {'padx': 20, 'pady': 5}

        self._section_title(f, "Paramètres temporels du trimestre actuel")

        fields = [
            ("TRIMESTRE_ACTUEL",
             "Format exact : T1_2026  (T + numéro 1-4 + underscore + année)\n"
             "Exemple : T1_2026, T3_2025, T4_2027",
             self.v_trimestre_actuel),
            ("ANNEE_ACTUELLE",
             "Année en cours de collecte. Exemple : 2026",
             self.v_annee),
            ("TRIMESTRE_NUMERO",
             "Numéro du trimestre en cours : entier entre 1 et 4\n"
             "  T1 → 1   |   T2 → 2   |   T3 → 3   |   T4 → 4",
             self.v_trim_numero),
            ("MOIS_EN_COURS",
             "Mois du 1er jour du trimestre (entier 1-12)\n"
             "  T1 → 1 (janvier)   |   T2 → 4 (avril)\n"
             "  T3 → 7 (juillet)   |   T4 → 10 (octobre)",
             self.v_mois_en_cours),
        ]

        for label, desc, var in fields:
            self._field_with_desc(f, label, desc, var)
            self.v_trimestre_actuel.trace_add('write',
                lambda *a: self._on_trim_actuel_change())

    # ══════════════════════════════════════════════════════════
    #  ONGLET 2 : COHORTES & RANGS
    # ══════════════════════════════════════════════════════════
    def _build_tab_cohortes(self, f):
        self._section_title(f, "Nombre de cohortes à réinterroger")

        desc_nb = (
            "Indiquez combien de cohortes passées seront réinterrogées ce trimestre.\n"
            "En général 3, mais peut être 2 au démarrage du programme ou 4 en régime croisière."
        )
        tk.Label(f, text=desc_nb, font=('Calibri', 9, 'italic'),
                 bg=CLR_FOND, fg=CLR_GRIS, justify='left',
                 wraplength=680).pack(anchor='w', padx=20)

        nb_frame = tk.Frame(f, bg=CLR_FOND)
        nb_frame.pack(anchor='w', padx=20, pady=6)
        tk.Label(nb_frame, text="Nombre de cohortes :", font=('Calibri', 10, 'bold'),
                 bg=CLR_FOND).pack(side='left')
        for v in [2, 3, 4]:
            tk.Radiobutton(nb_frame, text=str(v), variable=self.v_nb_cohortes,
                           value=v, font=('Calibri', 10), bg=CLR_FOND,
                           command=self._refresh_cohorte_rows
                           ).pack(side='left', padx=8)

        # Conteneur des lignes cohorte/rang
        self._section_title(f, "Cohortes et rangs d'interrogation")

        desc_rangs = (
            "Pour chaque cohorte, indiquez :\n"
            "  • Le nom de la cohorte au format exact T1_2026 (trimestre de PASSAGE 1)\n"
            "  • Le rang d'interrogation = nombre de fois où ce ménage a déjà été interrogé + 1\n"
            "    Exemple : cohorte interrogée pour la 2ème fois → rang = 2\n"
            "  ⚠️  Les valeurs rgmen, rghab et rang_ind seront toutes égales au rang saisi."
        )
        tk.Label(f, text=desc_rangs, font=('Calibri', 9, 'italic'),
                 bg=CLR_FOND, fg=CLR_GRIS, justify='left',
                 wraplength=680).pack(anchor='w', padx=20)

        # En-tête de colonnes
        hdr = tk.Frame(f, bg=CLR_BLEU_CLAIR)
        hdr.pack(fill='x', padx=20, pady=(6, 0))
        tk.Label(hdr, text="  #", width=4,  font=('Calibri', 10, 'bold'),
                 bg=CLR_BLEU_CLAIR, fg=CLR_BLEU_MARINE).pack(side='left')
        tk.Label(hdr, text="Cohorte (format T1_2026)", width=30,
                 font=('Calibri', 10, 'bold'), bg=CLR_BLEU_CLAIR,
                 fg=CLR_BLEU_MARINE).pack(side='left', padx=4)
        tk.Label(hdr, text="Rang d'interrogation", width=22,
                 font=('Calibri', 10, 'bold'), bg=CLR_BLEU_CLAIR,
                 fg=CLR_BLEU_MARINE).pack(side='left', padx=4)

        self.cohorte_rows_frame = tk.Frame(f, bg=CLR_FOND)
        self.cohorte_rows_frame.pack(fill='x', padx=20)
        self.cohorte_row_widgets = []
        self._refresh_cohorte_rows()

    def _refresh_cohorte_rows(self):
        """Affiche ou masque les lignes de cohorte selon le nombre choisi."""
        for w in self.cohorte_row_widgets:
            w.pack_forget()
        self.cohorte_row_widgets.clear()

        nb = self.v_nb_cohortes.get()
        for i in range(nb):
            row = tk.Frame(self.cohorte_rows_frame,
                           bg=CLR_BLANC if i % 2 == 0 else CLR_BLEU_CLAIR,
                           relief='flat', bd=0)
            row.pack(fill='x', pady=1)
            tk.Label(row, text=f"  {i+1}", width=4,
                     font=('Calibri', 10), bg=row['bg']).pack(side='left')
            tk.Entry(row, textvariable=self.cohorte_vars[i],
                     font=('Calibri', 10), width=28,
                     relief='solid', bd=1).pack(side='left', padx=4, pady=4)
            tk.Entry(row, textvariable=self.rang_vars[i],
                     font=('Calibri', 10), width=10,
                     relief='solid', bd=1).pack(side='left', padx=4)
            tk.Label(row, text="(entier ≥ 1)",
                     font=('Calibri', 9, 'italic'), bg=row['bg'],
                     fg=CLR_GRIS).pack(side='left', padx=4)
            self.cohorte_row_widgets.append(row)

    # ══════════════════════════════════════════════════════════
    #  ONGLET 3 : CHEMINS
    # ══════════════════════════════════════════════════════════
    def _build_tab_chemins(self, f):
        self._section_title(f, "REPERTOIRE_BASE — Dossier racine des bases brutes")

        desc_rb = (
            "Sélectionnez le dossier racine qui contient TOUS les sous-dossiers\n"
            "des passages 1 précédents.\n\n"
            "⚠️  IMPORTANT — Structure obligatoire à créer dans ce dossier :\n"
            "   Chaque cohorte doit avoir son propre sous-dossier nommé exactement :\n"
            "      Base_brute_{trimestre}    (exemple : Base_brute_T4_2025)\n\n"
            "   Chaque sous-dossier doit contenir :\n"
            "      • Le fichier .dta du ménage (ex : ENEM_2025T4.dta)\n"
            "      • Le fichier membres.dta\n\n"
            "   ⚠️  Ces bases doivent contenir UNIQUEMENT les données terrain\n"
            "       (Passage 1), PAS les données des téléopérateurs."
        )
        tk.Label(f, text=desc_rb, font=('Calibri', 9, 'italic'),
                 bg=CLR_FOND, fg=CLR_GRIS, justify='left',
                 wraplength=680).pack(anchor='w', padx=20, pady=(0, 6))

        self._browse_row(f, "REPERTOIRE_BASE :", self.v_repertoire_base,
                         self._browse_repertoire_base)

        # ── DOSSIER_SORTIE ────────────────────────────────────
        self._section_title(f, "DOSSIER_SORTIE — Dossier de sortie des fichiers")

        desc_ds = (
            "Sélectionnez le dossier PARENT dans lequel le programme créera\n"
            "automatiquement un sous-dossier nommé  Reinterrogation_{TRIMESTRE_ACTUEL}\n"
            "Exemple : si vous choisissez  D:\\ENEM_Working\\Base_prechargement_ENEM\n"
            "          et que le trimestre est T1_2026,\n"
            "          le dossier créé sera : ...\\Reinterrogation_T1_2026\\"
        )
        tk.Label(f, text=desc_ds, font=('Calibri', 9, 'italic'),
                 bg=CLR_FOND, fg=CLR_GRIS, justify='left',
                 wraplength=680).pack(anchor='w', padx=20, pady=(0, 6))

        self._browse_row(f, "Dossier parent :", self.v_dossier_parent,
                         self._browse_dossier_parent)

        # Aperçu du chemin final
        self.lbl_sortie_apercu = tk.Label(
            f, textvariable=self.v_dossier_sortie,
            font=('Calibri', 9, 'bold'), bg=CLR_FOND, fg=CLR_BLEU_ROI,
            anchor='w', wraplength=660
        )
        self.lbl_sortie_apercu.pack(anchor='w', padx=20, pady=(2, 8))

        # ── FICHIER_SEMAINE_REF ───────────────────────────────
        self._section_title(f, "FICHIER_SEMAINE_REF — Fichier de référence des semaines")

        desc_sem = (
            "Sélectionnez le fichier Excel  Semaine_ref.xlsx\n"
            "Ce fichier doit contenir les feuilles suivantes (structure fixe) :\n"
            "   • Semaine_ref_trim  : calendrier des semaines par trimestre\n"
            "   • label_region      : correspondance codes → noms de régions\n"
            "   • label_district    : correspondance codes → noms de districts\n"
            "   • label_departement : correspondance codes → noms de départements\n"
            "   • label_sp          : correspondance codes → noms de sous-préfectures"
        )
        tk.Label(f, text=desc_sem, font=('Calibri', 9, 'italic'),
                 bg=CLR_FOND, fg=CLR_GRIS, justify='left',
                 wraplength=680).pack(anchor='w', padx=20, pady=(0, 6))

        self._browse_row(f, "Fichier Semaine_ref.xlsx :", self.v_fichier_semaine,
                         self._browse_semaine_ref, is_file=True)

    # ══════════════════════════════════════════════════════════
    #  ONGLET 4 : AIDE & MODE D'EMPLOI
    # ══════════════════════════════════════════════════════════
    def _build_tab_aide(self, f):
        self._section_title(f, "📚 Description de chaque paramètre")

        aide_params = """
TRIMESTRE_ACTUEL
   Le trimestre en cours de collecte (Passage 1 actuel).
   Format strict : T + numéro (1 à 4) + underscore + année à 4 chiffres.
   ✓ Exemples valides : T1_2026 · T3_2025 · T4_2027
   ✗ Exemples invalides : 2026T1 · T1-2026 · t1_2026

ANNEE_ACTUELLE
   L'année civile du trimestre actuel. Entier sur 4 chiffres.
   Exemple : 2026

TRIMESTRE_NUMERO
   Le numéro du trimestre actuel, entier entre 1 et 4.
   T1 → 1  |  T2 → 2  |  T3 → 3  |  T4 → 4

MOIS_EN_COURS
   Le mois de début du trimestre actuel, entier entre 1 et 12.
   T1 → 1 (janvier)  |  T2 → 4 (avril)
   T3 → 7 (juillet)  |  T4 → 10 (octobre)

COHORTES_REINTERROGATION
   Liste des trimestres de Passage 1 qui sont réinterrogés durant le trimestre actuel.
   Chaque trimestre doit être au format T1_2026.
   Exemple pour T1_2026 : on réinterroge généralement T1_2025, T4_2024, T4_2025.
   Le nombre de cohortes peut varier entre 2 et 4 selon l'avancement du programme.

RANGS_PAR_COHORTE
   Pour chaque cohorte, le rang = nombre total de fois où ce ménage sera interrogé à cette collecte.
   • Première réinterrogation après le Passage 1 → rang = 2
   • Deuxième réinterrogation → rang = 3
   • Etc.
   Les variables rgmen, rghab et rang_ind prendront toutes la même valeur.

REPERTOIRE_BASE
   Dossier racine contenant tous les sous-dossiers de bases brutes.
   Chaque sous-dossier = une cohorte, nommé : Base_brute_{trimestre}
   Exemple : Base_brute_T4_2025  |  Base_brute_T1_2025  |  Base_brute_T4_2024

DOSSIER_PARENT (pour DOSSIER_SORTIE)
   Dossier dans lequel le programme crée automatiquement :
   Reinterrogation_{TRIMESTRE_ACTUEL}
   Exemple : D:\\ENEM_Working\\Base_prechargement_ENEM → crée Reinterrogation_T1_2026

FICHIER_SEMAINE_REF
   Fichier Excel Semaine_ref.xlsx avec les feuilles de calendrier et de labels géographiques.
   Ce fichier doit être mis à jour avant chaque trimestre avec les nouvelles semaines.
"""
        txt_aide = scrolledtext.ScrolledText(
            f, font=('Consolas', 9), bg='#1E1E2E', fg='#A8D8A8',
            height=18, wrap='word', state='normal'
        )
        txt_aide.insert('end', aide_params)
        txt_aide.config(state='disabled')
        txt_aide.pack(fill='both', expand=True, padx=20, pady=(0, 10))

        # ── Mode d'emploi des fichiers résultats ──────────────
        self._section_title(f, "📋 Mode d'emploi des fichiers résultats")

        mode_emploi = """
📁 FICHIERS GÉNÉRÉS dans le dossier Reinterrogation_{TRIMESTRE_ACTUEL} :
   ✓ QX_EEC_VF.xlsx  /  QX_EEC_VF.csv   → Base ménages
   ✓ membres.xlsx    /  membres.csv      → Base individus (résidents)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📌 QX_EEC_VF.xlsx — INSTRUCTIONS AVANT ENVOI :

   1. Ne pas oublier de renseigner les bons comptes des agents téléopérateurs
      dans la variable  _responsible  (un compte par téléopérateur).

   2. Pour l'import dans Survey Solutions :
      → Conserver uniquement les variables allant de  interview__id  à  _quantity
      → Enregistrer sous le format : Texte (séparateur : tabulation) (*.txt)
      → Nommer le fichier exactement : ENEM_2026T1  (selon le trimestre actuel)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📌 membres.xlsx — INSTRUCTIONS AVANT ENVOI :

   1. Pour filtrer les individus correspondant aux ménages de QX_EEC_VF.xlsx :
      → Faire une RECHERCHEV sur la colonne  interview__id

   2. Pour l'import dans Survey Solutions :
      → Conserver uniquement les variables allant de  membres__id  à  interview__id
      → Enregistrer sous le format : Texte (séparateur : tabulation) (*.txt)
      → Nommer le fichier exactement : membres

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
        txt_mode = scrolledtext.ScrolledText(
            f, font=('Consolas', 9), bg='#1E2A1E', fg='#C8E6C9',
            height=18, wrap='word', state='normal'
        )
        txt_mode.insert('end', mode_emploi)
        txt_mode.config(state='disabled')
        txt_mode.pack(fill='both', expand=True, padx=20, pady=(0, 10))

    # ══════════════════════════════════════════════════════════
    #  ZONE BAS : BOUTON LANCER + LOG + PROGRESSION
    # ══════════════════════════════════════════════════════════
    def _build_bottom(self):
        sep = ttk.Separator(self, orient='horizontal')
        sep.pack(fill='x', padx=12)

        bottom = tk.Frame(self, bg=CLR_FOND)
        bottom.pack(fill='both', expand=True, padx=12, pady=8)

        # Boutons
        btn_frame = tk.Frame(bottom, bg=CLR_FOND)
        btn_frame.pack(fill='x')

        tk.Button(
            btn_frame, text="✖  Quitter",
            font=('Calibri', 11), bg='#E0E0E0', fg='#333333',
            relief='flat', cursor='hand2', width=12,
            command=self.destroy
        ).pack(side='right', padx=(6, 0))

        tk.Button(
            btn_frame, text="🚀  Lancer le traitement",
            font=('Calibri', 11, 'bold'), bg=CLR_BLEU_MARINE, fg=CLR_BLANC,
            relief='flat', cursor='hand2', width=22,
            command=self._launch
        ).pack(side='right')

        # Journal
        tk.Label(bottom, text="Journal d'exécution",
                 font=('Calibri', 10, 'bold'), bg=CLR_FOND,
                 fg=CLR_BLEU_MARINE).pack(anchor='w', pady=(8, 2))

        self.log_box = scrolledtext.ScrolledText(
            bottom, height=10, font=('Consolas', 9),
            bg='#1E1E2E', fg='#A8D8A8',
            relief='flat', state='disabled', wrap='word'
        )
        self.log_box.pack(fill='both', expand=True)

        # Barre de progression
        self.progress_var = tk.IntVar(value=0)
        self.progress_bar = ttk.Progressbar(
            bottom, variable=self.progress_var, maximum=100, length=600
        )
        self.progress_bar.pack(fill='x', pady=(4, 2))

        self.lbl_progress = tk.Label(
            bottom, text='', bg=CLR_FOND,
            font=('Calibri', 9), fg=CLR_GRIS
        )
        self.lbl_progress.pack()

    # ── Helpers UI ─────────────────────────────────────────────
    def _section_title(self, parent, title: str):
        tk.Label(
            parent, text=title,
            font=('Calibri', 11, 'bold'), bg=CLR_BLEU_CLAIR,
            fg=CLR_BLEU_MARINE, anchor='w', padx=12, pady=4
        ).pack(fill='x', padx=20, pady=(12, 4))

    def _field_with_desc(self, parent, label: str, desc: str, var: tk.StringVar):
        frame = tk.Frame(parent, bg=CLR_FOND)
        frame.pack(fill='x', padx=20, pady=4)
        tk.Label(frame, text=label, font=('Calibri', 10, 'bold'),
                 bg=CLR_FOND, fg=CLR_TEXTE, width=22, anchor='w').pack(side='left')
        right = tk.Frame(frame, bg=CLR_FOND)
        right.pack(side='left', fill='x', expand=True)
        tk.Entry(right, textvariable=var, font=('Calibri', 10),
                 width=22, relief='solid', bd=1).pack(anchor='w')
        tk.Label(right, text=desc, font=('Calibri', 9, 'italic'),
                 bg=CLR_FOND, fg=CLR_GRIS, justify='left',
                 wraplength=540).pack(anchor='w')

    def _browse_row(self, parent, label: str, var: tk.StringVar,
                    cmd, is_file: bool = False):
        frame = tk.Frame(parent, bg=CLR_FOND)
        frame.pack(fill='x', padx=20, pady=4)
        tk.Label(frame, text=label, font=('Calibri', 10, 'bold'),
                 bg=CLR_FOND, fg=CLR_TEXTE, width=26, anchor='w').pack(side='left')
        tk.Entry(frame, textvariable=var, font=('Calibri', 10),
                 width=48, relief='solid', bd=1,
                 state='readonly').pack(side='left', padx=(0, 6))
        tk.Button(frame, text="Parcourir…",
                  font=('Calibri', 10), bg=CLR_BLEU_ROI, fg=CLR_BLANC,
                  relief='flat', cursor='hand2',
                  command=cmd).pack(side='left')

    # ── Actions parcourir ──────────────────────────────────────
    def _browse_repertoire_base(self):
        d = filedialog.askdirectory(title="Sélectionner le dossier REPERTOIRE_BASE")
        if d:
            self.v_repertoire_base.set(d)

    def _browse_dossier_parent(self):
        d = filedialog.askdirectory(title="Sélectionner le dossier PARENT pour la sortie")
        if d:
            self.v_dossier_parent.set(d)
            self._on_trim_actuel_change()

    def _browse_semaine_ref(self):
        f = filedialog.askopenfilename(
            title="Sélectionner le fichier Semaine_ref.xlsx",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tous", "*.*")]
        )
        if f:
            self.v_fichier_semaine.set(f)

    def _on_trim_actuel_change(self, *_):
        parent = self.v_dossier_parent.get()
        trim   = self.v_trimestre_actuel.get().strip()
        if parent and trim:
            self.v_dossier_sortie.set(
                os.path.join(parent, f"Reinterrogation_{trim}")
            )
        elif parent:
            self.v_dossier_sortie.set(parent)

    # ── Validation ────────────────────────────────────────────
    def _valider(self) -> tuple:
        """Valide tous les champs. Retourne (True, params) ou (False, message)."""
        trim = self.v_trimestre_actuel.get().strip()
        if not REGEX_TRIMESTRE.match(trim):
            return False, "TRIMESTRE_ACTUEL invalide.\nFormat attendu : T1_2026 (T + 1-4 + _ + année)"

        try:
            annee = int(self.v_annee.get())
            if annee < 2024 or annee > 2030:
                return False, "ANNEE_ACTUELLE doit être entre 2024 et 2030."
        except ValueError:
            return False, "ANNEE_ACTUELLE doit être un entier (ex: 2026)."

        try:
            trim_num = int(self.v_trim_numero.get())
            if trim_num not in [1, 2, 3, 4]:
                return False, "TRIMESTRE_NUMERO doit être 1, 2, 3 ou 4."
        except ValueError:
            return False, "TRIMESTRE_NUMERO doit être un entier entre 1 et 4."

        try:
            mois = int(self.v_mois_en_cours.get())
            if mois < 1 or mois > 12:
                return False, "MOIS_EN_COURS doit être entre 1 et 12."
        except ValueError:
            return False, "MOIS_EN_COURS doit être un entier entre 1 et 12."

        # Cohortes
        nb = self.v_nb_cohortes.get()
        cohortes = []
        rangs = {}
        for i in range(nb):
            c = self.cohorte_vars[i].get().strip()
            if not c:
                return False, f"La cohorte n°{i+1} est vide."
            if not REGEX_TRIMESTRE.match(c):
                return False, (f"Cohorte n°{i+1} invalide : '{c}'\n"
                               f"Format attendu : T1_2026")
            if c not in NOMS_FICHIERS:
                return False, (f"Cohorte '{c}' non reconnue dans la liste des fichiers.\n"
                               f"Vérifiez le format ou mettez à jour NOMS_FICHIERS dans le code.")
            try:
                r = int(self.rang_vars[i].get())
                if r < 1:
                    return False, f"Le rang de la cohorte '{c}' doit être ≥ 1."
            except ValueError:
                return False, f"Le rang de la cohorte '{c}' doit être un entier."
            cohortes.append(c)
            rangs[c] = r

        if len(set(cohortes)) != len(cohortes):
            return False, "Deux cohortes identiques détectées. Chaque cohorte doit être unique."

        # Chemins
        if not self.v_repertoire_base.get():
            return False, "Veuillez sélectionner le REPERTOIRE_BASE."
        if not os.path.isdir(self.v_repertoire_base.get()):
            return False, "REPERTOIRE_BASE introuvable sur le disque."
        if not self.v_dossier_parent.get():
            return False, "Veuillez sélectionner le dossier parent de sortie."
        if not self.v_fichier_semaine.get():
            return False, "Veuillez sélectionner le fichier Semaine_ref.xlsx."
        if not os.path.isfile(self.v_fichier_semaine.get()):
            return False, "Fichier Semaine_ref.xlsx introuvable sur le disque."

        self._on_trim_actuel_change()

        params = {
            'trimestre_actuel' : trim,
            'annee_actuelle'   : annee,
            'trimestre_numero' : trim_num,
            'mois_en_cours'    : mois,
            'cohortes'         : cohortes,
            'rangs'            : rangs,
            'repertoire_base'  : self.v_repertoire_base.get(),
            'dossier_sortie'   : self.v_dossier_sortie.get(),
            'fichier_semaine_ref': self.v_fichier_semaine.get(),
        }
        return True, params

    # ── Lancement ─────────────────────────────────────────────
    def _launch(self):
        ok, result = self._valider()
        if not ok:
            messagebox.showerror("Paramètres invalides", result)
            return

        params = result
        self._log_clear()
        self.progress_var.set(0)
        self.lbl_progress.config(text="Traitement en cours…")

        def worker():
            success = run_traitement(params, self._log, self._update_progress)
            if success:
                self.lbl_progress.config(text="✅ Traitement terminé avec succès")
                messagebox.showinfo(
                    "✅ Terminé",
                    f"Traitement terminé avec succès !\n\n"
                    f"📁 Fichiers dans :\n{params['dossier_sortie']}"
                )
            else:
                self.lbl_progress.config(text="❌ Erreur — voir le journal")
                messagebox.showerror(
                    "Erreur",
                    "Le traitement a échoué.\nConsultez le journal pour plus de détails."
                )

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    # ── Helpers log / progression ──────────────────────────────
    def _log(self, msg: str):
        def _do():
            self.log_box.config(state='normal')
            self.log_box.insert('end', msg + '\n')
            self.log_box.see('end')
            self.log_box.config(state='disabled')
        self.after(0, _do)

    def _log_clear(self):
        self.log_box.config(state='normal')
        self.log_box.delete('1.0', 'end')
        self.log_box.config(state='disabled')

    def _update_progress(self, val: int):
        def _do():
            self.progress_var.set(val)
            self.lbl_progress.config(text=f"{val}%")
        self.after(0, _do)


# ═══════════════════════════════════════════════════════════════════════════════
#  POINT D'ENTRÉE
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    app = AppReinterrogation()
    app.mainloop()
