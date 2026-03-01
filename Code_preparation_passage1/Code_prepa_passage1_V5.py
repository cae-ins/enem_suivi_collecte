# -*- coding: utf-8 -*-
"""
Created on Tue Feb 24 21:41:34 2026

@author: mg.kouame
"""

"""
================================================================================
PROGRAMME : Préparation des fichiers de collecte terrain (Passage 1)
================================================================================
OBJECTIF  : Préparer les fichiers Excel pour la collecte de terrain du Passage 1
            en affectant les Zones de Dénombrement (ZD) aux agents de collecte
            et en générant les fichiers de Dénombrement, Ménage et Calendrier.

CONTEXTE  : Enquête trimestrielle sur l'emploi - Collecte sur 10 ans
            - Passage 1 : Collecte initiale sur le terrain (Dénombrement + Ménages)
            - Chaque trimestre visite de nouvelles ZD selon l'échantillonnage
            - Affectation automatique des agents selon leur région

PROCESSUS :
    1. Charger l'échantillon des ZD à visiter (filtré par sous-échantillon)
    2. Fusionner avec les données de géoréférencement
    3. Affecter automatiquement les agents de collecte par région
    4. Générer le fichier Dénombrement (1 ligne par ZD)
    5. Générer le fichier Ménage avec variable Reference (duplication hors-ABIDJAN)
    6. Générer le fichier Calendrier_menage avec 4 variables de dates de semaines
    7. Créer les fichiers séparés ABIDJAN / HORS-ABIDJAN

LOGIQUE CALENDRIER :
    - ABIDJAN (13 ZD) : 1 seule semaine par ZD
        * Ordre = numéro de semaine (1 à 13)
        * Date_debut_sem_1_ref / Date_fin_sem_1_ref remplies
        * Date_debut_sem_2_ref / Date_fin_sem_2_ref vides
    - HORS-ABIDJAN (32 régions, 7 ZD) : ZD1-6 sur 2 semaines, ZD7 sur 1 semaine
        * Ordre=1 → sem 1+2 | Ordre=2 → sem 3+4 | ... | Ordre=6 → sem 11+12
        * Ordre=7 → sem 13 (sem_1 = sem_2, valeur répétée)

MODIFICATIONS v3.0 :
    - Refonte complète du fichier Calendrier_menage
    - 4 nouvelles variables de dates : Date_debut_sem_1_ref, Date_fin_sem_1_ref,
      Date_debut_sem_2_ref, Date_fin_sem_2_ref
    - Déduplication du Calendrier via colonne Code1
    - Suppression des anciennes colonnes Date_debut_sem_ref / Date_fin_sem_ref

AUTEUR    : KOUAME KOUASSI GUY MARTIAL
DATE      : 26 décembre 2025
VERSION   : 3.0
================================================================================
"""

import pandas as pd
import numpy as np

# ============================================================================
# PARAMÈTRES DE CONFIGURATION
# ============================================================================

# --- PARAMÈTRES DE LA COLLECTE ---
TRIMESTRE_COLLECTE      = "1T2026"    # Ex: 1T2026, 2T2025, etc.
TRIMESTRE_COLLECTE_DATE = "T1_2026"   # Format pour le fichier Semaine_ref
NUMERO_TRIMESTRE        = 1           # Numéro du trimestre (1, 2, 3 ou 4)
MOIS_COLLECTE           = 1           # Mois de début (1=Janvier, 4=Avril, etc.)
ANNEE_COLLECTE          = 2026        # Année de collecte

# --- PARAMÈTRES D'ÉCHANTILLONNAGE ---
SOUS_ECHANTILLON        = 8           # Numéro du sous-échantillon (variable "sous_echant")

# --- PARAMÈTRES DE LA COLLECTE MÉNAGE ---
NOMBRE_MENAGES_PAR_ZD   = 6           # Nombre de ménages à interroger par ZD

# --- CHEMINS DES FICHIERS ---
DOSSIER_TRAVAIL          = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1"
DOSSIER_TRAVAIL_RESULTAT = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1\Resutat"

# Fichiers d'entrée
# Note : ZD 4091 enquêtée par anticipation en 2024 → remplacée par ZD 4085
# Note : ZD 6027 (GBOKLE) repositionnée en 7ème ZD du T1-2026 (semaine_ref=13, sous_echant=8)
FICHIER_ECHANTILLON  = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1\Echantillon_ZD_VF_ACTUALISEE.xlsx"
FEUILLE_ECHANTILLON  = "BASEGLO"

FICHIER_GEOREF       = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1\VF_BASE_ILOT_12012024_VF_work_Geovf.xlsx"

FICHIER_EQUIPES      = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1\EquipeParRegionVF.xlsx"
FEUILLE_EQUIPES      = "Equipe"

FICHIER_SEMAINES_REF = r"D:\ENEM_Working\Base_prechargement_ENEM\Semaine_reference\Semaine_ref.xlsx"
FEUILLE_SEMAINES_REF = "Semaine_ref_trim"

# Noms des fichiers de sortie
NOM_FICHIER_DENOMBREMENT           = f"Denombrement_{TRIMESTRE_COLLECTE}.xlsx"
NOM_FICHIER_MENAGE                 = f"Menage_{TRIMESTRE_COLLECTE}.xlsx"
NOM_FICHIER_CALENDRIER_MENAGE      = f"Calendrier_menage_{TRIMESTRE_COLLECTE}.xlsx"
NOM_FICHIER_MENAGE_ABIDJAN         = f"Menage_ABIDJAN_{TRIMESTRE_COLLECTE}.xlsx"
NOM_FICHIER_MENAGE_HORS_ABIDJAN    = f"Menage_HORS_ABIDJAN_{TRIMESTRE_COLLECTE}.xlsx"
NOM_FICHIER_DENOMBREMENT_ABIDJAN   = f"Denombrement_ABIDJAN_{TRIMESTRE_COLLECTE}.xlsx"
NOM_FICHIER_DENOMBREMENT_HORS_ABIDJAN = f"Denombrement_HORS_ABIDJAN_{TRIMESTRE_COLLECTE}.xlsx"

# ============================================================================
# AFFICHAGE DES PARAMÈTRES
# ============================================================================

print("=" * 80)
print("PRÉPARATION DES FICHIERS DE COLLECTE - PASSAGE 1")
print("=" * 80)
print(f"\n📅 PARAMÈTRES DE COLLECTE")
print(f"   Trimestre        : {TRIMESTRE_COLLECTE}")
print(f"   Année            : {ANNEE_COLLECTE}")
print(f"   Mois             : {MOIS_COLLECTE}")
print(f"   Sous-échantillon : {SOUS_ECHANTILLON}")
print(f"   Ménages/ZD       : {NOMBRE_MENAGES_PAR_ZD}")
print("\n" + "=" * 80 + "\n")

# ============================================================================
# CHARGEMENT DES DONNÉES
# ============================================================================

print("📂 CHARGEMENT DES DONNÉES...")

# Échantillon des ZD
if FEUILLE_ECHANTILLON:
    df_echantillon = pd.read_excel(FICHIER_ECHANTILLON, sheet_name=FEUILLE_ECHANTILLON)
else:
    df_echantillon = pd.read_excel(FICHIER_ECHANTILLON)
df_echantillon = df_echantillon[df_echantillon["sous_echant"] == SOUS_ECHANTILLON]

# Géoréférencement
df_georef = pd.read_excel(FICHIER_GEOREF)

# Équipes (agents de collecte uniquement)
df_equipes = pd.read_excel(FICHIER_EQUIPES, sheet_name=FEUILLE_EQUIPES)
df_equipes = df_equipes[df_equipes["Type de compte"] == "Agent de collecte"]

# Semaines de référence
df_semaines_ref = pd.read_excel(FICHIER_SEMAINES_REF, sheet_name=FEUILLE_SEMAINES_REF)

print(f"   ✓ Échantillon          : {len(df_echantillon)} ZD (sous-échantillon {SOUS_ECHANTILLON})")
print(f"   ✓ Géoréférencement     : {len(df_georef)} enregistrements")
print(f"   ✓ Agents de collecte   : {len(df_equipes)} agents")
print(f"   ✓ Semaines de référence: {len(df_semaines_ref)} semaines")

# ============================================================================
# PRÉPARATION ET FUSION DES DONNÉES
# ============================================================================

print("\n🔧 PRÉPARATION ET FUSION DES DONNÉES...")

# Formater NUM_ZD_Vf sur 4 chiffres et créer la clé de fusion
df_echantillon['NUM_ZD_Vf'] = df_echantillon['NUM_ZD_Vf'].astype(str).str.zfill(4)
df_echantillon["CLEZD"] = df_echantillon["NomSp"] + df_echantillon["NUM_ZD_Vf"]

# Fusion avec le géoréférencement
data_merged = pd.merge(df_echantillon, df_georef, on='CLEZD', how='left')

# Nettoyer les colonnes dupliquées (_x supprimées, _y renommées)
colonnes_x = [col for col in data_merged.columns if col.endswith('_x')]
data_merged = data_merged.drop(columns=colonnes_x)
data_merged.rename(columns=lambda x: x.rstrip('_y') if x.endswith('_y') else x, inplace=True)

print(f"   ✓ Fusion géoréférencement : {len(data_merged)} enregistrements")

# ============================================================================
# AFFECTATION DES AGENTS PAR RÉGION
# ============================================================================

print("\n👥 AFFECTATION DES AGENTS PAR RÉGION...")

# Produit cartésien ZD × Agents, puis filtrage par région
df_cross    = pd.merge(data_merged, df_equipes, how='cross')
df_resultat = df_cross[df_cross['Region'] == df_cross['NomReg']].copy()

# Reformater NUM_ZD_Vf
df_resultat['NUM_ZD_Vf'] = df_resultat['NUM_ZD_Vf'].astype(str).str.zfill(4)

print(f"   ✓ {len(df_resultat)} affectations ZD-Agent créées")
print("\n   📊 Répartition par région :")
repartition = df_resultat.groupby('Region').size().sort_values(ascending=False)
for region, count in list(repartition.items())[:10]:
    print(f"      • {region}: {count} affectations")
if len(repartition) > 10:
    print(f"      ... et {len(repartition) - 10} autres régions")

# ============================================================================
# CRÉATION DE LA VARIABLE CODE1
# ============================================================================

print("\n🔢 CRÉATION DE LA VARIABLE CODE1...")

df_resultat['cle_groupement'] = (
    df_resultat['NUM_ZD_Vf'].astype(str) + "_" + df_resultat['LibQtierCpt'].astype(str)
)

np.random.seed(42)
combinaisons_uniques = df_resultat['cle_groupement'].unique()
code_mapping = {cle: np.random.randint(10000000, 100000000) for cle in combinaisons_uniques}

df_resultat['Code1'] = "A" + df_resultat['cle_groupement'].map(code_mapping).astype(str)
df_resultat.drop(columns=['cle_groupement'], inplace=True)

print(f"   ✓ {len(combinaisons_uniques)} codes uniques générés")
print(f"   ✓ Exemple de Code1 : {df_resultat['Code1'].iloc[0]}")

# ============================================================================
# CRÉATION DE LA VARIABLE ORDRE
# ============================================================================

print("\n📋 CRÉATION DE LA VARIABLE ORDRE...")

# Abidjan : Ordre = semaine_ref (1 à 13)
# Hors-Abidjan : semaine_ref impaire (1,3,5,7,9,11,13) → Ordre (1 à 7)
correspondance_ordre = {1: 1, 3: 2, 5: 3, 7: 4, 9: 5, 11: 6, 13: 7}

def calculer_ordre(row):
    if row['NomReg'] == 'ABIDJAN':
        return row['semaine_ref']
    return correspondance_ordre.get(row['semaine_ref'], np.nan)

df_resultat['Ordre'] = df_resultat.apply(calculer_ordre, axis=1).fillna(0).astype(int)

nb_zero = (df_resultat['Ordre'] == 0).sum()
if nb_zero > 0:
    print(f"   ⚠️  {nb_zero} affectations avec Ordre=0 (semaine_ref invalide)")
else:
    print(f"   ✓ Toutes les affectations ont un Ordre valide")

# ============================================================================
# PRÉPARATION DES DATES DE RÉFÉRENCE (pour Dénombrement et Ménage)
# ============================================================================

print("\n📅 PRÉPARATION DES DATES DE RÉFÉRENCE...")
print(f"   🔍 Colonnes dans Semaine_ref.xlsx : {list(df_semaines_ref.columns)}")

df_semaines_trim = df_semaines_ref[df_semaines_ref['Trimestre'] == TRIMESTRE_COLLECTE_DATE].copy()

if len(df_semaines_trim) == 0:
    print(f"   ⚠️  ERREUR : Aucune ligne pour le trimestre '{TRIMESTRE_COLLECTE_DATE}'")
    print(f"   ⚠️  Valeurs disponibles : {df_semaines_ref['Trimestre'].unique()}")
    df_resultat['Date1_ref'] = ""
    df_resultat['Date2_ref'] = ""
else:
    print(f"   ✓ {len(df_semaines_trim)} semaines trouvées pour {TRIMESTRE_COLLECTE}")

    # Lookup rapide : N_semaine → (Date1, Date2)
    lookup_n_semaine = {
        row['N_semaine']: (row['Date1'], row['Date2'])
        for _, row in df_semaines_trim.iterrows()
    }

    df_resultat['Date1_ref'] = df_resultat['semaine_ref'].map(
        lambda s: lookup_n_semaine.get(s, (None, None))[0]
    )
    df_resultat['Date2_ref'] = df_resultat['semaine_ref'].map(
        lambda s: lookup_n_semaine.get(s, (None, None))[1]
    )

    # Formatage des dates
    for col in ['Date1_ref', 'Date2_ref']:
        if pd.api.types.is_datetime64_any_dtype(df_resultat[col]):
            df_resultat[col] = df_resultat[col].dt.strftime('%Y-%m-%d')
        df_resultat[col] = (
            df_resultat[col].fillna("").astype(str)
            .replace(['NaT', 'nan', 'None'], "")
        )

    nb_ok  = (df_resultat['Date1_ref'] != "").sum()
    nb_nok = (df_resultat['Date1_ref'] == "").sum()
    print(f"   ✅ Dates trouvées : {nb_ok} / {len(df_resultat)} lignes")
    if nb_nok > 0:
        print(f"   ⚠️  {nb_nok} lignes sans dates")

# ============================================================================
# GÉNÉRATION DU FICHIER DÉNOMBREMENT
# ============================================================================

print("\n📝 GÉNÉRATION DU FICHIER DÉNOMBREMENT...")

denombrement = pd.DataFrame({
    'Region'          : df_resultat['NomReg'],
    'sp'              : df_resultat['NomSp'],
    'ord_sem'         : "Sem_" + df_resultat['Ordre'].astype(str) + f"_{TRIMESTRE_COLLECTE}",
    'HH01'            : TRIMESTRE_COLLECTE,
    'HH0'             : f"1erPassage{TRIMESTRE_COLLECTE}-sp-" + df_resultat['NomSp'] + "-zd-" + df_resultat['NUM_ZD_Vf'],
    'HH2A'            : df_resultat['Dr'],
    'HH1'             : df_resultat['NumeroDistrict'],
    'HH2'             : df_resultat['NumeroRegion'],
    'HH3'             : df_resultat['NumeroDepart'],
    'HH4'             : df_resultat['NumeroSp'],
    'HH6'             : df_resultat['CodeMilieu'],
    'HH8'             : df_resultat['NUM_ZD_Vf'],
    'HH8A'            : np.where(df_resultat['Plusieurs Loc'] == 1, df_resultat['NomLoc'], 'Zd sur plusieurs localité'),
    'HH7'             : np.where(df_resultat['Zd campement'] == "Pas  campement", 1, 7),
    'HH8B'            : df_resultat['LibQtierCpt'],
    'trimestreencours': NUMERO_TRIMESTRE,
    'mois_en_cours'   : MOIS_COLLECTE,
    'annee'           : ANNEE_COLLECTE,
    'Date1'           : df_resultat['Date1_ref'],
    'Code1'           : df_resultat['Code1'],
    '_responsible'    : df_resultat['login'],
    '_quantity'       : 1,
    'Ordre'           : df_resultat['Ordre'],
    'cle'             : df_resultat['NumeroSp'].astype(str) + df_resultat['NUM_ZD_Vf'].astype(str),
})

denombrement_abidjan      = denombrement[denombrement['Region'] == 'ABIDJAN'].copy()
denombrement_hors_abidjan = denombrement[denombrement['Region'] != 'ABIDJAN'].copy()

denombrement.to_excel(DOSSIER_TRAVAIL_RESULTAT + "\\" + NOM_FICHIER_DENOMBREMENT, index=False)
denombrement_abidjan.to_excel(DOSSIER_TRAVAIL_RESULTAT + "\\" + NOM_FICHIER_DENOMBREMENT_ABIDJAN, index=False)
denombrement_hors_abidjan.to_excel(DOSSIER_TRAVAIL_RESULTAT + "\\" + NOM_FICHIER_DENOMBREMENT_HORS_ABIDJAN, index=False)

print(f"   ✓ {NOM_FICHIER_DENOMBREMENT} ({len(denombrement)} lignes)")
print(f"   ✓ {NOM_FICHIER_DENOMBREMENT_ABIDJAN} ({len(denombrement_abidjan)} lignes)")
print(f"   ✓ {NOM_FICHIER_DENOMBREMENT_HORS_ABIDJAN} ({len(denombrement_hors_abidjan)} lignes)")

# ============================================================================
# GÉNÉRATION DU FICHIER MÉNAGE
# ============================================================================

print(f"\n🏠 GÉNÉRATION DU FICHIER MÉNAGE ({NOMBRE_MENAGES_PAR_ZD} ménages par ZD)...")

menage = pd.DataFrame({
    'Region'          : df_resultat['NomReg'],
    'sp'              : df_resultat['NomSp'],
    'Cohorte'         : TRIMESTRE_COLLECTE,
    'ord_sem'         : "Sem_" + df_resultat['Ordre'].astype(str) + f"_{TRIMESTRE_COLLECTE}",
    'HH01'            : TRIMESTRE_COLLECTE,
    'HH0'             : f"1erPassage{TRIMESTRE_COLLECTE}-sp-" + df_resultat['NomSp'] + "-zd-" + df_resultat['NUM_ZD_Vf'],
    'HH2A'            : df_resultat['Dr'],
    'HH1'             : df_resultat['NumeroDistrict'],
    'HH2'             : df_resultat['NumeroRegion'],
    'HH3'             : df_resultat['NumeroDepart'],
    'HH4'             : df_resultat['NumeroSp'],
    'HH6'             : df_resultat['CodeMilieu'],
    'HH8'             : df_resultat['NUM_ZD_Vf'],
    'HH8A'            : np.where(df_resultat['Plusieurs Loc'] == 1, df_resultat['NomLoc'], 'Zd sur plusieurs localité'),
    'HH7'             : 1,
    'HH7B'            : 1,
    'HH8B'            : df_resultat['LibQtierCpt'],
    'rghab'           : 1,
    'rgmen'           : 1,
    'V1MODINTR'       : 1,
    'trimestreencours': NUMERO_TRIMESTRE,
    'mois_en_cours'   : MOIS_COLLECTE,
    'annee'           : ANNEE_COLLECTE,
    'Date1'           : df_resultat['Date1_ref'],
    'Date2'           : df_resultat['Date2_ref'],
    'Reference'       : 1,
    'Code1'           : df_resultat['Code1'],
    '_responsible'    : df_resultat['login'],
    '_quantity'       : NOMBRE_MENAGES_PAR_ZD,
    'Ordre'           : df_resultat['Ordre'],
    'cle'             : df_resultat['NumeroSp'].astype(str) + df_resultat['NUM_ZD_Vf'].astype(str),
})

# Duplication des lignes HORS-ABIDJAN avec Reference=2
menage_abidjan           = menage[menage['Region'] == 'ABIDJAN'].copy()
menage_non_abidjan_ref1  = menage[menage['Region'] != 'ABIDJAN'].copy()
menage_non_abidjan_ref2  = menage_non_abidjan_ref1.copy()
menage_non_abidjan_ref2['Reference'] = 2

# Tri pour regrouper les lignes Reference=1 et 2 par ZD
cle_tri = (
    menage_non_abidjan_ref1['Region'].astype(str) + "_" +
    menage_non_abidjan_ref1['sp'].astype(str) + "_" +
    menage_non_abidjan_ref1['HH8'].astype(str)
)
menage_non_abidjan_ref1 = menage_non_abidjan_ref1.assign(cle_tri=cle_tri)
menage_non_abidjan_ref2 = menage_non_abidjan_ref2.assign(cle_tri=cle_tri)

menage_non_abidjan_combine = (
    pd.concat([menage_non_abidjan_ref1, menage_non_abidjan_ref2], ignore_index=True)
    .sort_values(by=['cle_tri', 'Reference'])
    .drop(columns=['cle_tri'])
)

menage_final = pd.concat([menage_abidjan, menage_non_abidjan_combine], ignore_index=True)

# Sauvegarde
menage_final.to_excel(DOSSIER_TRAVAIL_RESULTAT + "\\" + NOM_FICHIER_MENAGE, index=False)
menage_abidjan.to_excel(DOSSIER_TRAVAIL_RESULTAT + "\\" + NOM_FICHIER_MENAGE_ABIDJAN, index=False)
menage_non_abidjan_combine.to_excel(DOSSIER_TRAVAIL_RESULTAT + "\\" + NOM_FICHIER_MENAGE_HORS_ABIDJAN, index=False)

print(f"   ✓ {NOM_FICHIER_MENAGE} ({len(menage_final)} lignes)")
print(f"   ✓ {NOM_FICHIER_MENAGE_ABIDJAN} ({len(menage_abidjan)} lignes)")
print(f"   ✓ {NOM_FICHIER_MENAGE_HORS_ABIDJAN} ({len(menage_non_abidjan_combine)} lignes)")

# Vérification de la duplication
nb_ref1 = (menage_non_abidjan_combine['Reference'] == 1).sum()
nb_ref2 = (menage_non_abidjan_combine['Reference'] == 2).sum()
if nb_ref1 == nb_ref2:
    print(f"   ✅ Duplication correcte : {nb_ref1} lignes Reference=1 et {nb_ref2} Reference=2")
else:
    print(f"   ⚠️  Déséquilibre : Reference=1 ({nb_ref1}) vs Reference=2 ({nb_ref2})")

# ============================================================================
# GÉNÉRATION DU FICHIER CALENDRIER_MENAGE (VERSION CORRIGÉE)
# ============================================================================

print("\n📅 GÉNÉRATION DU FICHIER CALENDRIER_MENAGE...")

# --- Lookup semaines depuis Semaine_ref.xlsx ---
# La colonne s'appelle 'Numero_semaine' avec valeurs texte : "Semaine 1", "Semaine 2"...
df_semaines_cal = df_semaines_ref[df_semaines_ref['Trimestre'] == TRIMESTRE_COLLECTE_DATE].copy()

if len(df_semaines_cal) == 0:
    print(f"   ⚠️  ERREUR : Aucune ligne pour le trimestre '{TRIMESTRE_COLLECTE_DATE}'")
else:
    print(f"   ✓ {len(df_semaines_cal)} semaines trouvées pour {TRIMESTRE_COLLECTE_DATE}")

# Extraire le numéro entier depuis "Semaine 1" → 1
df_semaines_cal['num_sem'] = (
    df_semaines_cal['Numero_semaine'].astype(str).str.extract(r'(\d+)').astype(int)
)

def formater_date(val):
    """Convertit une date en chaîne YYYY-MM-DD ou retourne chaîne vide."""
    if pd.isna(val) or val is None:
        return ""
    if hasattr(val, 'strftime'):
        return val.strftime('%Y-%m-%d')
    s = str(val).strip()
    return "" if s in ('NaT', 'nan', 'None', '') else s

# Dictionnaire : num_sem → (date_debut, date_fin)
lookup_semaines_cal = {
    row['num_sem']: (formater_date(row['Date_debut_sem_ref']), formater_date(row['Date_fin_sem_ref']))
    for _, row in df_semaines_cal.iterrows()
}
print(f"   ✓ Lookup semaines créé : {sorted(lookup_semaines_cal.keys())}")

# --- Table de correspondance Ordre → numéros de semaines ---
# Hors-Abidjan : Ordre 1-6 → 2 semaines, Ordre 7 → semaine 13 répétée
# Abidjan : Ordre = numéro de semaine
correspondance_cal = {
    1: (1,  2),
    2: (3,  4),
    3: (5,  6),
    4: (7,  8),
    5: (9,  10),
    6: (11, 12),
    7: (13, 13),  # Semaine 13 répétée pour sem_1 et sem_2
}

def get_dates_calendrier(row):
    """
    Retourne (Date_debut_sem_1, Date_fin_sem_1, Date_debut_sem_2, Date_fin_sem_2)

    - ABIDJAN : Ordre = numéro de semaine → sem_1 remplie, sem_2 vide
    - HORS-ABIDJAN : correspondance_cal[Ordre] → 2 semaines
        * Ordre 7 : semaine 13 répétée dans sem_1 et sem_2
    """
    ordre = int(row['Ordre']) if not pd.isna(row['Ordre']) else 0
    if ordre == 0:
        return ("", "", "", "")

    if row['Region'] == 'ABIDJAN':
        debut1, fin1 = lookup_semaines_cal.get(ordre, ("", ""))
        return (debut1, fin1, "", "")
    else:
        paire = correspondance_cal.get(ordre)
        if paire is None:
            return ("", "", "", "")
        sem1, sem2 = paire
        debut1, fin1 = lookup_semaines_cal.get(sem1, ("", ""))
        debut2, fin2 = lookup_semaines_cal.get(sem2, ("", ""))
        return (debut1, fin1, debut2, fin2)

# --- Copie de menage_final + déduplication via Code1 ---
calendrier_menage = menage_final.copy()

nb_avant = len(calendrier_menage)
calendrier_menage = calendrier_menage.drop_duplicates(subset=['Code1'], keep='first')
nb_apres = len(calendrier_menage)
print(f"   ✓ Déduplication Code1 : {nb_avant} → {nb_apres} lignes ({nb_avant - nb_apres} doublons supprimés)")

# --- Calcul des 4 variables de dates ---
resultats = calendrier_menage.apply(get_dates_calendrier, axis=1)
calendrier_menage['Date_debut_sem_1_ref'] = [r[0] for r in resultats]
calendrier_menage['Date_fin_sem_1_ref']   = [r[1] for r in resultats]
calendrier_menage['Date_debut_sem_2_ref'] = [r[2] for r in resultats]
calendrier_menage['Date_fin_sem_2_ref']   = [r[3] for r in resultats]

# --- Ordre des colonnes final ---
colonnes_calendrier = [
    'Region', 'sp', 'Cohorte', 'ord_sem', 'HH01', 'HH0', 'HH2A',
    'HH1', 'HH2', 'HH3', 'HH4', 'HH6', 'HH8', 'HH8A', 'HH7', 'HH7B', 'HH8B',
    'rghab', 'rgmen', 'V1MODINTR',
    'trimestreencours', 'mois_en_cours', 'annee', 'Date1', 'Date2',
    'Date_debut_sem_1_ref', 'Date_fin_sem_1_ref',
    'Date_debut_sem_2_ref', 'Date_fin_sem_2_ref',
    'Reference', 'Code1', '_responsible', '_quantity', 'Ordre', 'cle'
]

colonnes_presentes = [c for c in colonnes_calendrier if c in calendrier_menage.columns]
calendrier_menage = calendrier_menage[colonnes_presentes]

# --- Sauvegarde ---
calendrier_menage.to_excel(DOSSIER_TRAVAIL_RESULTAT + "\\" + NOM_FICHIER_CALENDRIER_MENAGE, index=False)
print(f"   ✓ {NOM_FICHIER_CALENDRIER_MENAGE} ({len(calendrier_menage)} lignes)")

# Vérification rapide
nb_sem1_ok = (calendrier_menage['Date_debut_sem_1_ref'] != "").sum()
nb_sem2_ok = (calendrier_menage['Date_debut_sem_2_ref'] != "").sum()
print(f"   ✓ Lignes avec sem_1 remplie : {nb_sem1_ok}")
print(f"   ✓ Lignes avec sem_2 remplie : {nb_sem2_ok} (Abidjan et Ordre=7 : valeurs vides ou répétées)")

# Échantillons de contrôle
print(f"\n   📋 Contrôle HORS-ABIDJAN (7 premières lignes) :")
ech = calendrier_menage[calendrier_menage['Region'] != 'ABIDJAN'][
    ['Region', 'HH8', 'Ordre', 'Date_debut_sem_1_ref', 'Date_fin_sem_1_ref',
     'Date_debut_sem_2_ref', 'Date_fin_sem_2_ref']
].head(7)
for _, r in ech.iterrows():
    print(f"      {r['Region']} | ZD {r['HH8']} | Ordre {int(r['Ordre'])} | "
          f"Sem1[{r['Date_debut_sem_1_ref']} → {r['Date_fin_sem_1_ref']}] | "
          f"Sem2[{r['Date_debut_sem_2_ref']} → {r['Date_fin_sem_2_ref']}]")

print(f"\n   📋 Contrôle ABIDJAN (5 premières lignes) :")
ech = calendrier_menage[calendrier_menage['Region'] == 'ABIDJAN'][
    ['HH8', 'Ordre', 'Date_debut_sem_1_ref', 'Date_fin_sem_1_ref',
     'Date_debut_sem_2_ref', 'Date_fin_sem_2_ref']
].head(5)
for _, r in ech.iterrows():
    print(f"      ZD {r['HH8']} | Ordre {int(r['Ordre'])} | "
          f"Sem1[{r['Date_debut_sem_1_ref']} → {r['Date_fin_sem_1_ref']}] | "
          f"Sem2['{r['Date_debut_sem_2_ref']}' → '{r['Date_fin_sem_2_ref']}'] (vide attendu)")

# ============================================================================
# RÉSUMÉ FINAL
# ============================================================================

print("\n" + "=" * 80)
print("✅ TRAITEMENT TERMINÉ AVEC SUCCÈS !")
print("=" * 80)
print(f"\n📊 RÉSUMÉ DE LA PRODUCTION")
print(f"   • ZD traitées         : {len(df_echantillon)}")
print(f"   • Affectations créées : {len(df_resultat)}")
print(f"   • Régions couvertes   : {df_resultat['Region'].nunique()}")
print(f"   • Agents mobilisés    : {df_resultat['login'].nunique()}")

print(f"\n📁 FICHIERS GÉNÉRÉS")
print(f"   COMPLETS :")
print(f"   • {NOM_FICHIER_DENOMBREMENT} ({len(denombrement)} lignes)")
print(f"   • {NOM_FICHIER_MENAGE} ({len(menage_final)} lignes)")
print(f"   • {NOM_FICHIER_CALENDRIER_MENAGE} ({len(calendrier_menage)} lignes)")
print(f"   ABIDJAN :")
print(f"   • {NOM_FICHIER_DENOMBREMENT_ABIDJAN} ({len(denombrement_abidjan)} lignes)")
print(f"   • {NOM_FICHIER_MENAGE_ABIDJAN} ({len(menage_abidjan)} lignes)")
print(f"   HORS-ABIDJAN :")
print(f"   • {NOM_FICHIER_DENOMBREMENT_HORS_ABIDJAN} ({len(denombrement_hors_abidjan)} lignes)")
print(f"   • {NOM_FICHIER_MENAGE_HORS_ABIDJAN} ({len(menage_non_abidjan_combine)} lignes)")

print(f"\n📌 À RETENIR POUR L'ENVOI :")
print(f"   ✓ menage.xlsx / _responsible contient déjà les comptes des agents")
print(f"   ✓ menage.xlsx / Variables à envoyer : de ord_sem à _quantity")
print(f"   ✓ menage.xlsx / Format d'export : Texte (séparateur : tabulation) (*.txt)")
print(f"   ✓ denombrement.xlsx / _responsible contient déjà les comptes des agents")
print(f"   ✓ denombrement.xlsx / Variables à envoyer : de ord_sem à _quantity")
print(f"   ✓ denombrement.xlsx / Format d'export : Texte (séparateur : tabulation) (*.txt)")

print(f"\n📂 Localisation : {DOSSIER_TRAVAIL_RESULTAT}")
print("=" * 80)