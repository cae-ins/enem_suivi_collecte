"""
================================================================================
PROGRAMME : Préparation des fichiers de collecte terrain (Passage 1)
================================================================================
OBJECTIF  : Préparer les fichiers Excel pour la collecte de terrain du Passage 1
            en affectant les Zones de Dénombrement (ZD) aux agents de collecte
            et en générant les fichiers de Dénombrement et Ménage.

CONTEXTE  : Enquête trimestrielle sur l'emploi - Collecte sur 10 ans
            - Passage 1 : Collecte initiale sur le terrain (Dénombrement + Ménages)
            - Chaque trimestre visite de nouvelles ZD selon l'échantillonnage
            - Affectation automatique des agents selon leur région
            
PROCESSUS :
    1. Charger l'échantillon des ZD à visiter (filtré par sous-échantillon)
    2. Fusionner avec les données de géoréférencement
    3. Affecter automatiquement les agents de collecte par région
    4. Générer le fichier Dénombrement (1 ligne par ZD)
    5. Générer le fichier Ménage (6 ménages par ZD)

AUTEUR    : KOUAME KOUASSI GUY MARTIAL 
DATE      : 26 décembre 2025
VERSION   : 1.0 - Version nettoyée et documentée
================================================================================
"""

import pandas as pd
import numpy as np

# ============================================================================
# PARAMÈTRES DE CONFIGURATION
# ============================================================================

# --- PARAMÈTRES DE LA COLLECTE ---
TRIMESTRE_COLLECTE = "1T2026"           # Trimestre de collecte (ex: 1T2026, 2T2025, etc.)
TRIMESTRE_COLLECTE_DATE = "T1_2026"     # Pour l'importation de la date : Trimestre de collecte (ex: 1T2026, 2T2025, etc.)
NUMERO_TRIMESTRE = 1                     # Numéro du trimestre (1, 2, 3 ou 4)
MOIS_COLLECTE = 1                        # Mois de début de collecte (1=Janvier, 4=Avril, etc.)
ANNEE_COLLECTE = 2026                    # Année de collecte

# --- PARAMÈTRES D'ÉCHANTILLONNAGE ---
SOUS_ECHANTILLON = 8                     # Numéro du sous-échantillon à traiter (ex: 7, 8, etc.)
                                         # Correspond à la variable "sous_echant" du fichier

# --- PARAMÈTRES DE LA COLLECTE MÉNAGE ---
NOMBRE_MENAGES_PAR_ZD = 6               # Nombre de ménages à interroger par ZD

# --- CHEMINS DES FICHIERS D'ENTRÉE ---
DOSSIER_TRAVAIL = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1"

DOSSIER_TRAVAIL_RESULTAT = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1\Resutat"

FICHIER_ECHANTILLON = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1\Echantillon_ZD_VF.xlsx"
FEUILLE_ECHANTILLON = "BASEGLO"          # Nom de la feuille Excel (ou None pour la feuille par défaut)

FICHIER_GEOREF = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1\VF_BASE_ILOT_12012024_VF_work_Geovf.xlsx"

FICHIER_EQUIPES = r"D:\ENEM_Working\Base_prechargement_ENEM\Code_prepa_passage1\EquipeParRegionVF.xlsx"
FEUILLE_EQUIPES = "Equipe"

FICHIER_SEMAINES_REF = r"D:\ENEM_Working\Base_prechargement_ENEM\Semaine_reference\Semaine_ref.xlsx"
FEUILLE_SEMAINES_REF = "Semaine_ref_trim"

# --- NOMS DES FICHIERS DE SORTIE ---
NOM_FICHIER_DENOMBREMENT = f"Denombrement_{TRIMESTRE_COLLECTE}.xlsx"
NOM_FICHIER_MENAGE = f"Menage_{TRIMESTRE_COLLECTE}.xlsx"

# ============================================================================
# AFFICHAGE DES PARAMÈTRES
# ============================================================================

print("="*80)
print("PRÉPARATION DES FICHIERS DE COLLECTE - PASSAGE 1")
print("="*80)
print(f"\n📅 PARAMÈTRES DE COLLECTE")
print(f"   Trimestre     : {TRIMESTRE_COLLECTE}")
print(f"   Année         : {ANNEE_COLLECTE}")
print(f"   Mois          : {MOIS_COLLECTE}")
print(f"   Sous-échantillon : {SOUS_ECHANTILLON}")
print(f"   Ménages/ZD    : {NOMBRE_MENAGES_PAR_ZD}")
print("\n" + "="*80 + "\n")

# ============================================================================
# CHARGEMENT DES DONNÉES
# ============================================================================

print("📂 CHARGEMENT DES DONNÉES...")

# Charger l'échantillon des ZD
if FEUILLE_ECHANTILLON:
    df_echantillon = pd.read_excel(FICHIER_ECHANTILLON, sheet_name=FEUILLE_ECHANTILLON)
else:
    df_echantillon = pd.read_excel(FICHIER_ECHANTILLON)

# Filtrer sur le sous-échantillon spécifié
df_echantillon = df_echantillon[df_echantillon["sous_echant"] == SOUS_ECHANTILLON]

# Charger le fichier de géoréférencement complet
df_georef = pd.read_excel(FICHIER_GEOREF)

# Charger la liste des équipes (feuille "Equipe", filtré sur Type="Agent de collecte")
df_equipes = pd.read_excel(FICHIER_EQUIPES, sheet_name=FEUILLE_EQUIPES)
df_equipes = df_equipes[df_equipes["Type de compte"] == "Agent de collecte"]

# Charger le fichier des semaines de référence
df_semaines_ref = pd.read_excel(FICHIER_SEMAINES_REF, sheet_name=FEUILLE_SEMAINES_REF)

print(f"   ✓ Échantillon          : {len(df_echantillon)} ZD (sous-échantillon {SOUS_ECHANTILLON})")
print(f"   ✓ Géoréférencement     : {len(df_georef)} enregistrements")
print(f"   ✓ Agents de collecte   : {len(df_equipes)} agents")
print(f"   ✓ Semaines de référence: {len(df_semaines_ref)} semaines")

# ============================================================================
# PRÉPARATION ET FORMATAGE DES DONNÉES
# ============================================================================

print("\n🔧 PRÉPARATION DES DONNÉES...")

# Formater le numéro de ZD sur 4 chiffres avec zéros devant (ex: 5 → 0005)
df_echantillon['NUM_ZD_Vf'] = df_echantillon['NUM_ZD_Vf'].astype(str).str.zfill(4)

# Créer une clé unique pour la fusion : NomSp (Sous-Préfecture) + NUM_ZD_Vf
df_echantillon["CLEZD"] = df_echantillon["NomSp"] + df_echantillon["NUM_ZD_Vf"]

# ============================================================================
# FUSION DES DONNÉES GÉOGRAPHIQUES
# ============================================================================

print("🔗 FUSION AVEC LE GÉORÉFÉRENCEMENT...")

# Fusionner l'échantillon avec le fichier de géoréférencement sur la clé CLEZD
data_merged = pd.merge(df_echantillon, df_georef, on='CLEZD', how='left')

# Nettoyer les colonnes dupliquées après la fusion
# - Supprimer les colonnes se terminant par "_x" (garder les valeurs de l'échantillon)
colonnes_x = [col for col in data_merged.columns if col.endswith('_x')]
data_merged = data_merged.drop(columns=colonnes_x)

# - Renommer les colonnes "_y" en supprimant le suffixe (garder les valeurs du géoréférencement)
data_merged.rename(columns=lambda x: x.rstrip('_y') if x.endswith('_y') else x, inplace=True)

print(f"   ✓ Fusion complétée : {len(data_merged)} enregistrements")

# ============================================================================
# AFFECTATION DES AGENTS PAR RÉGION
# ============================================================================

print("👥 AFFECTATION DES AGENTS PAR RÉGION...")

# Créer un produit cartésien entre les ZD et les agents
# (chaque ZD est associée à tous les agents)
df_cross = pd.merge(data_merged, df_equipes, how='cross')

# Filtrer pour ne garder que les agents de la même région que la ZD
# La colonne 'Region' vient de data_merged (région de la ZD)
# La colonne 'NomReg' vient de df_equipes (région de l'agent)
df_resultat = df_cross[df_cross['Region'] == df_cross['NomReg']]

print(f"   ✓ {len(df_resultat)} affectations ZD-Agent créées")

# Afficher un aperçu des affectations par région
print("\n   📊 Répartition par région :")
repartition = df_resultat.groupby('Region').size().sort_values(ascending=False)
for region, count in repartition.head(10).items():
    print(f"      • {region}: {count} affectations")
if len(repartition) > 10:
    print(f"      ... et {len(repartition) - 10} autres régions")

# ============================================================================
# FORMATAGE FINAL DES DONNÉES
# ============================================================================

# S'assurer que NUM_ZD_Vf est bien formaté sur 4 chiffres
df_resultat['NUM_ZD_Vf'] = df_resultat['NUM_ZD_Vf'].astype(str).str.zfill(4)

# ============================================================================
# CRÉATION DE LA VARIABLE CODE1
# ============================================================================

print("\n🔢 CRÉATION DE LA VARIABLE CODE1...")

# Créer une clé de groupement : NUM_ZD_Vf + NomQuartier (LibQtierCpt)
df_resultat['cle_groupement'] = df_resultat['NUM_ZD_Vf'].astype(str) + "_" + df_resultat['LibQtierCpt'].astype(str)

# Générer un nombre aléatoire de 8 chiffres pour chaque combinaison unique de NUM_ZD_Vf + LibQtierCpt
np.random.seed(42)  # Pour la reproductibilité (optionnel, retirer si on veut du vrai aléatoire)

# Obtenir les combinaisons uniques
combinaisons_uniques = df_resultat['cle_groupement'].unique()

# Créer un dictionnaire avec un code de 8 chiffres pour chaque combinaison
code_mapping = {}
for cle in combinaisons_uniques:
    # Générer un nombre aléatoire entre 10000000 et 99999999 (8 chiffres)
    code_aleatoire = np.random.randint(10000000, 100000000)
    code_mapping[cle] = code_aleatoire

# Appliquer le mapping pour créer Code1
df_resultat['code_8chiffres'] = df_resultat['cle_groupement'].map(code_mapping)

# Construire Code1 : "A" + 8 chiffres aléatoires
df_resultat['Code1'] = "A" + df_resultat['code_8chiffres'].astype(str)

print(f"   ✓ {len(combinaisons_uniques)} codes uniques générés (ZD + Quartier)")
print(f"   ✓ Exemple de Code1 : {df_resultat['Code1'].iloc[0]}")

# Nettoyer les colonnes temporaires
df_resultat.drop(columns=['cle_groupement', 'code_8chiffres'], inplace=True)

# ============================================================================
# CRÉATION DE LA VARIABLE ORDRE
# ============================================================================

print("\n📋 CRÉATION DE LA VARIABLE ORDRE...")

def calculer_ordre(row):
    """
    Calcule la valeur de Ordre selon la région et la semaine de référence
    
    Règles :
    - Pour ABIDJAN : Ordre = semaine_ref (identique)
    - Pour les autres régions : transformation selon table de correspondance
    """
    region = row['NomReg']
    semaine = row['semaine_ref']
    
    # Pour Abidjan : Ordre = semaine_ref
    if region == 'ABIDJAN':
        return semaine
    
    # Pour les autres régions : table de correspondance
    correspondance = {
        1: 1,
        3: 2,
        5: 3,
        7: 4,
        9: 5,
        11: 6,
        13: 7
    }
    
    # Retourner la correspondance, ou NaN si la semaine n'est pas dans la table
    return correspondance.get(semaine, np.nan)

# Appliquer la fonction pour créer la variable Ordre
df_resultat['Ordre'] = df_resultat.apply(calculer_ordre, axis=1)

# Convertir en entier (gérer les NaN éventuels)
df_resultat['Ordre'] = df_resultat['Ordre'].fillna(0).astype(int)

# Vérifier s'il y a des valeurs à 0 (cas problématiques)
nb_ordre_zero = (df_resultat['Ordre'] == 0).sum()
if nb_ordre_zero > 0:
    print(f"   ⚠️  {nb_ordre_zero} affectations avec Ordre=0 (semaine_ref invalide)")
else:
    print(f"   ✓ Toutes les affectations ont un Ordre valide")

# Afficher la répartition par région
print("\n   📊 Répartition des Ordre par région :")
repartition_ordre = df_resultat.groupby(['NomReg', 'Ordre']).size().reset_index(name='Count')
regions_principales = df_resultat['NomReg'].value_counts().head(3).index

for region in regions_principales:
    ordres_region = repartition_ordre[repartition_ordre['NomReg'] == region]
    print(f"      • {region}:")
    for _, row in ordres_region.head(3).iterrows():
        print(f"        - Ordre {int(row['Ordre'])}: {row['Count']} affectations")

# ============================================================================
# PRÉPARATION DES DATES DE RÉFÉRENCE
# ============================================================================

print("\n📅 PRÉPARATION DES DATES DE RÉFÉRENCE...")

# Vérifier que semaine_ref existe dans df_resultat
if 'semaine_ref' not in df_resultat.columns:
    print(f"   ⚠️  ERREUR : La colonne 'semaine_ref' n'existe pas dans {FICHIER_ECHANTILLON}")
    print(f"   ⚠️  Veuillez vérifier que cette colonne existe dans la feuille {FEUILLE_ECHANTILLON}")
    df_resultat['Date1_ref'] = ""
    df_resultat['Date2_ref'] = ""
else:
    print(f"   ✓ Colonne 'semaine_ref' trouvée dans l'échantillon")
    
    # Afficher les colonnes disponibles dans le fichier de référence
    print(f"   🔍 Colonnes dans Semaine_ref.xlsx : {list(df_semaines_ref.columns)}")
    
    # Filtrer STRICTEMENT sur le trimestre T1_2026
    df_semaines_trim = df_semaines_ref[df_semaines_ref['Trimestre'] == TRIMESTRE_COLLECTE_DATE].copy()
    
    if len(df_semaines_trim) == 0:
        print(f"   ⚠️  ERREUR : Aucune ligne trouvée pour le trimestre '{TRIMESTRE_COLLECTE}'")
        print(f"   ⚠️  Valeurs de Trimestre disponibles : {df_semaines_ref['Trimestre'].unique()}")
        df_resultat['Date1_ref'] = ""
        df_resultat['Date2_ref'] = ""
    else:
        print(f"   ✓ {len(df_semaines_trim)} semaines trouvées pour le trimestre {TRIMESTRE_COLLECTE}")
        
        # Afficher les valeurs disponibles
        print(f"   📊 Colonnes N_semaine disponibles : {sorted(df_semaines_trim['N_semaine'].unique())}")
        print(f"   📊 Valeurs semaine_ref dans échantillon : {sorted(df_resultat['semaine_ref'].unique())}")
        
        # Initialiser les colonnes Date1_ref et Date2_ref
        df_resultat['Date1_ref'] = None
        df_resultat['Date2_ref'] = None
        
        # MÉTHODE RECHERCHEV : Parcourir chaque ligne et chercher la correspondance
        nb_dates_trouvees = 0
        nb_dates_non_trouvees = 0
        
        print(f"\n   🔄 Application du RECHERCHEV sur {len(df_resultat)} lignes...")
        
        for idx, row in df_resultat.iterrows():
            semaine_ref_menage = row['semaine_ref']
            
            # Chercher la correspondance dans le fichier de référence
            # pour le TRIMESTRE_COLLECTE et la semaine_ref du ménage
            correspondance = df_semaines_trim[
                df_semaines_trim['N_semaine'] == semaine_ref_menage
            ]
            
            if len(correspondance) > 0:
                # Récupérer Date1 et Date2
                date1_ref = correspondance.iloc[0]['Date1']
                date2_ref = correspondance.iloc[0]['Date2']
                
                # Assigner dans df_resultat
                df_resultat.at[idx, 'Date1_ref'] = date1_ref
                df_resultat.at[idx, 'Date2_ref'] = date2_ref
                nb_dates_trouvees += 1
            else:
                nb_dates_non_trouvees += 1
        
        print(f"   ✅ Dates trouvées : {nb_dates_trouvees} / {len(df_resultat)} lignes")
        
        if nb_dates_non_trouvees > 0:
            print(f"   ⚠️  {nb_dates_non_trouvees} lignes sans dates")
            # Identifier les semaines problématiques
            semaines_sans_dates = df_resultat[df_resultat['Date1_ref'].isna()]['semaine_ref'].unique()
            print(f"   ⚠️  Semaines sans correspondance dans {TRIMESTRE_COLLECTE} : {sorted(semaines_sans_dates)}")
        
        # Convertir les dates au format YYYY-MM-DD si nécessaire
        if pd.api.types.is_datetime64_any_dtype(df_resultat['Date1_ref']):
            df_resultat['Date1_ref'] = df_resultat['Date1_ref'].dt.strftime('%Y-%m-%d')
        if pd.api.types.is_datetime64_any_dtype(df_resultat['Date2_ref']):
            df_resultat['Date2_ref'] = df_resultat['Date2_ref'].dt.strftime('%Y-%m-%d')
        
        # Remplacer les valeurs manquantes par des chaînes vides
        df_resultat['Date1_ref'] = df_resultat['Date1_ref'].fillna("").astype(str)
        df_resultat['Date2_ref'] = df_resultat['Date2_ref'].fillna("").astype(str)
        
        # Nettoyer les valeurs "NaT" ou "nan"
        df_resultat['Date1_ref'] = df_resultat['Date1_ref'].replace(['NaT', 'nan', 'None'], "")
        df_resultat['Date2_ref'] = df_resultat['Date2_ref'].replace(['NaT', 'nan', 'None'], "")
        
        # Afficher un échantillon des résultats
        if nb_dates_trouvees > 0:
            print(f"\n   📋 Échantillon des dates attribuées (5 premières lignes) :")
            echantillon = df_resultat[df_resultat['Date1_ref'] != ""][['NUM_ZD_Vf', 'NomSp', 'semaine_ref', 'Date1_ref', 'Date2_ref']].head()
            for idx, row in echantillon.iterrows():
                print(f"      ZD {row['NUM_ZD_Vf']} ({row['NomSp']}) | Semaine {row['semaine_ref']} | {row['Date1_ref']} → {row['Date2_ref']}")
        
        # Afficher la répartition par semaine
        print(f"\n   📊 Répartition des lignes par semaine de référence :")
        repartition = df_resultat[df_resultat['Date1_ref'] != ""].groupby('semaine_ref').size().sort_index()
        for semaine, count in repartition.head(7).items():
            print(f"      • Semaine {semaine}: {count} lignes")
        if len(repartition) > 7:
            print(f"      ... et {len(repartition) - 7} autres semaines")

# ============================================================================
# GÉNÉRATION DU FICHIER DÉNOMBREMENT (1 ligne par ZD)
# ============================================================================

print("\n📝 GÉNÉRATION DU FICHIER DÉNOMBREMENT...")

denombrement = pd.DataFrame()

# Informations administratives
denombrement['Region'] = df_resultat['NomReg']
denombrement['sp'] = df_resultat['NomSp']
denombrement['ord_sem'] = "Sem_" + df_resultat['Ordre'].astype(str) + f"_{TRIMESTRE_COLLECTE}" 
denombrement['HH01'] = TRIMESTRE_COLLECTE
denombrement['HH0'] = f"1erPassage{TRIMESTRE_COLLECTE}-sp-" + df_resultat['NomSp'] + "-zd-" + df_resultat['NUM_ZD_Vf']
denombrement['HH2A'] = df_resultat['Dr']  # Direction Régionale

# Codes géographiques
denombrement['HH1'] = df_resultat['NumeroDistrict']
denombrement['HH2'] = df_resultat['NumeroRegion']
denombrement['HH3'] = df_resultat['NumeroDepart']
denombrement['HH4'] = df_resultat['NumeroSp']
denombrement['HH6'] = df_resultat['CodeMilieu']
denombrement['HH8'] = df_resultat['NUM_ZD_Vf']

# Informations sur la localité
denombrement['HH8A'] = np.where(
    df_resultat['Plusieurs Loc'] == 1, 
    df_resultat['NomLoc'], 
    'Zd sur plusieurs localité'
)

# Type de zone (1=zone normale, 7=campement)
denombrement['HH7'] = np.where(
    df_resultat['Zd campement'] == "Pas  campement", 
    1, 
    7
)

denombrement['HH8B'] = df_resultat['LibQtierCpt']  # Libellé Quartier/Campement

# Informations temporelles
denombrement['trimestreencours'] = NUMERO_TRIMESTRE
denombrement['mois_en_cours'] = MOIS_COLLECTE
denombrement['annee'] = ANNEE_COLLECTE
denombrement['Date1'] = df_resultat['Date1_ref']  # Dates de début de semaine de référence
denombrement['Date2'] = df_resultat['Date2_ref']  # Dates de fin de semaine de référence

# Informations d'affectation
denombrement['Code1'] = df_resultat['Code1']
denombrement['_responsible'] = df_resultat['login']  # Agent responsable
denombrement['_quantity'] = 1  # 1 dénombrement par ZD
denombrement['Ordre'] = df_resultat['Ordre']
denombrement['cle'] = df_resultat['NumeroSp'].astype(str) + df_resultat['NUM_ZD_Vf'].astype(str)

# Sauvegarder le fichier
fichier_sortie_denom = DOSSIER_TRAVAIL_RESULTAT + "\\" + NOM_FICHIER_DENOMBREMENT
denombrement.to_excel(fichier_sortie_denom, index=False)
print(f"   ✓ Fichier créé : {NOM_FICHIER_DENOMBREMENT}")
print(f"   ✓ Nombre de lignes : {len(denombrement)}")

# ============================================================================
# GÉNÉRATION DU FICHIER MÉNAGE (N lignes par ZD)
# ============================================================================

print(f"\n🏠 GÉNÉRATION DU FICHIER MÉNAGE ({NOMBRE_MENAGES_PAR_ZD} ménages par ZD)...")

menage = pd.DataFrame()

# Informations administratives
menage['Region'] = df_resultat['NomReg']
menage['sp'] = df_resultat['NomSp']
menage['Cohorte'] = f"{TRIMESTRE_COLLECTE}"
menage['ord_sem'] = "Sem_" + df_resultat['Ordre'].astype(str) + f"_{TRIMESTRE_COLLECTE}"
menage['HH01'] = TRIMESTRE_COLLECTE
menage['HH0'] = f"1erPassage{TRIMESTRE_COLLECTE}-sp-" + df_resultat['NomSp'] + "-zd-" + df_resultat['NUM_ZD_Vf']
menage['HH2A'] = df_resultat['Dr']  # Direction Régionale

# Codes géographiques
menage['HH1'] = df_resultat['NumeroDistrict']
menage['HH2'] = df_resultat['NumeroRegion']
menage['HH3'] = df_resultat['NumeroDepart']
menage['HH4'] = df_resultat['NumeroSp']
menage['HH6'] = df_resultat['CodeMilieu']
menage['HH8'] = df_resultat['NUM_ZD_Vf']

# Informations sur la localité
menage['HH8A'] = np.where(
    df_resultat['Plusieurs Loc'] == 1, 
    df_resultat['NomLoc'], 
    'Zd sur plusieurs localité'
)

# Type de zone (toujours 1 pour le fichier Ménage)
menage['HH7'] = 1
menage['HH7B'] = 1

menage['HH8B'] = df_resultat['LibQtierCpt']  # Libellé Quartier/Campement

# Informations de collecte
menage['rghab'] = 1  # Rang habitation
menage['rgmen'] = 1  # Rang ménage
menage['V1MODINTR'] = 1  # Mode d'interview

# Informations temporelles
menage['trimestreencours'] = NUMERO_TRIMESTRE
menage['mois_en_cours'] = MOIS_COLLECTE
menage['annee'] = ANNEE_COLLECTE
menage['Date1'] = df_resultat['Date1_ref']  # Dates de début de semaine de référence
menage['Date2'] = df_resultat['Date2_ref']  # Dates de fin de semaine de référence
menage['Code1'] = df_resultat['Code1'] 

# Informations d'affectation
menage['_responsible'] = df_resultat['login']  # Agent responsable
menage['_quantity'] = NOMBRE_MENAGES_PAR_ZD  # Nombre de ménages à interroger par ZD
menage['Ordre'] = df_resultat['Ordre']
menage['cle'] = df_resultat['NumeroSp'].astype(str) + df_resultat['NUM_ZD_Vf'].astype(str) 

# Sauvegarder le fichier
fichier_sortie_menage = DOSSIER_TRAVAIL_RESULTAT + "\\" + NOM_FICHIER_MENAGE
menage.to_excel(fichier_sortie_menage, index=False)
print(f"   ✓ Fichier créé : {NOM_FICHIER_MENAGE}")
print(f"   ✓ Nombre de lignes : {len(menage)}")

# ============================================================================
# RÉSUMÉ FINAL
# ============================================================================

print("\n" + "="*80)
print("✅ TRAITEMENT TERMINÉ AVEC SUCCÈS !")
print("="*80)
print(f"\n📊 RÉSUMÉ DE LA PRODUCTION")
print(f"   • ZD traitées           : {len(df_echantillon)}")
print(f"   • Affectations créées   : {len(df_resultat)}")
print(f"   • Régions couvertes     : {df_resultat['Region'].nunique()}")
print(f"   • Agents mobilisés      : {df_resultat['login'].nunique()}")
print(f"\n📁 FICHIERS GÉNÉRÉS")
print(f"   • Dénombrement : {NOM_FICHIER_DENOMBREMENT} ({len(denombrement)} lignes)")
print(f"   • Ménage       : {NOM_FICHIER_MENAGE} ({len(menage)} lignes)")
print(f"\n📂 Localisation : {DOSSIER_TRAVAIL}")
print("="*80)
