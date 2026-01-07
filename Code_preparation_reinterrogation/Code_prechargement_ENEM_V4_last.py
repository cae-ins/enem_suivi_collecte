# -*- coding: utf-8 -*-
"""
Created on Fri Jan  2 15:13:11 2026

@author: mg.kouame
"""

# -*- coding: utf-8 -*-
"""
PROGRAMME : Préparation des fichiers de réinterrogation (Passage 2)
OBJECTIF  : Préparer les fichiers Excel pour la réinterrogation téléphonique
            des ménages enquêtés au Passage 1, avec préchargement de certaines
            réponses pour faciliter la collecte au Passage 2.

CONTEXTE  : Enquête trimestrielle sur l'emploi - Gestion multi-cohortes
            - Passage 1 : Collecte initiale sur le terrain
            - Passage 2 : Réinterrogation par téléopérateurs
            - Chaque trimestre réinterroge 3 cohortes précédentes
            
AUTEUR    : mg.kouame
DATE      : 22 décembre 2025
VERSION   : 2.1 - Gestion automatique des rangs d'interrogation par cohorte
"""

import pandas as pd
import numpy as np
import os
from datetime import datetime

# ==============================================================================
# 🔧 PARAMÈTRES À CONFIGURER PAR L'UTILISATEUR
# ==============================================================================

# ===== TRIMESTRE EN COURS DE COLLECTE =====
TRIMESTRE_ACTUEL = "T4_2025"  # Format: T1_2025, T2_2025, T3_2025, T4_2025, etc.
ANNEE_ACTUELLE = 2025
TRIMESTRE_NUMERO = 4  # 1, 2, 3 ou 4
MOIS_EN_COURS = 10    # Mois du début du trimestre (ex: T4 = octobre = 10)

# ===== COHORTES À RÉINTERROGER =====
# Liste des trimestres qui doivent être réinterrogés
# Exemple pour T4_2025 : on réinterroge T3_2024, T4_2024 et T3_2025
COHORTES_REINTERROGATION = [
    "T3_2024",  # Cohorte 1 (il y a 1 an)
    "T4_2024",  # Cohorte 2 (il y a 2 trimestres)
    "T3_2025"   # Cohorte 3 (trimestre précédent)
]

# ===== CONFIGURATION DES RANGS D'INTERROGATION PAR COHORTE =====
# Dictionnaire définissant les rangs pour chaque cohorte
# Format: {cohorte: {'rgmen': valeur, 'rghab': valeur, 'rang_ind': valeur}}
RANGS_PAR_COHORTE = {
    "T3_2024": {"rgmen": 4, "rghab": 4, "rang_ind": 4},  # 4ème interrogation
    "T4_2024": {"rgmen": 3, "rghab": 3, "rang_ind": 3},  # 3ème interrogation
    "T3_2025": {"rgmen": 2, "rghab": 2, "rang_ind": 2}   # 2ème interrogation
}


# ===== CHEMINS DES DOSSIERS =====
# Dossier racine contenant tous les sous-dossiers des trimestres
REPERTOIRE_BASE = r"D:\ENEM_Working\Base_prechargement_ENEM"

# Dossier de sortie pour les fichiers de réinterrogation
DOSSIER_SORTIE = r"D:\ENEM_Working\Base_prechargement_ENEM\Reinterrogation_" + TRIMESTRE_ACTUEL

# Fichier de référence des semaines
FICHIER_SEMAINE_REF = r"D:\ENEM_Working\Base_prechargement_ENEM\Semaine_reference\Semaine_ref.xlsx"

# ===== CORRESPONDANCE TRIMESTRE -> NOM FICHIER =====
# Dictionnaire pour mapper les trimestres aux noms de fichiers
NOMS_FICHIERS = {
    "T2_2024": "ENEM_2024T2.dta",
    "T3_2024": "ENEM_2024T3.dta",
    "T4_2024": "ENEM_2024T4.dta",
    "T1_2025": "ENEM_2025T1.dta",
    "T2_2025": "ENEM_2025T2.dta",
    "T3_2025": "ENEM_2025T3.dta",
    "T4_2025": "ENEM_2025T4.dta",
    "T1_2026": "ENEM_2026T1.dta"  # Pour extension future
}

# ==============================================================================
# 📊 AFFICHAGE DE LA CONFIGURATION
# ==============================================================================

print("=" * 70)
print("PROGRAMME DE PRÉPARATION DES FICHIERS DE RÉINTERROGATION")
print("=" * 70)
print(f"\n📅 Trimestre actuel : {TRIMESTRE_ACTUEL}")
print(f"📅 Année : {ANNEE_ACTUELLE} - Trimestre : {TRIMESTRE_NUMERO}")
print(f"📅 Mois de début : {MOIS_EN_COURS}")
print(f"\n🔄 Cohortes à réinterroger : {len(COHORTES_REINTERROGATION)}")
for i, cohorte in enumerate(COHORTES_REINTERROGATION, 1):
    print(f"   {i}. {cohorte}")

print(f"\n🔢 Configuration des rangs d'interrogation :")
for cohorte, rangs in RANGS_PAR_COHORTE.items():
    print(f"   • {cohorte} : rgmen={rangs['rgmen']}, rghab={rangs['rghab']}, rang_ind={rangs['rang_ind']}")

print(f"\n📁 Dossier de sortie : {DOSSIER_SORTIE}")
print(f"📁 Fichier de référence des semaines : {FICHIER_SEMAINE_REF}")
print("=" * 70)

# ==============================================================================
# 🔍 VALIDATION DE LA CONFIGURATION
# ==============================================================================

print("\n🔍 Validation de la configuration...")

# Vérifier que toutes les cohortes de réinterrogation ont des rangs définis
cohortes_sans_rangs = [c for c in COHORTES_REINTERROGATION if c not in RANGS_PAR_COHORTE]
if cohortes_sans_rangs:
    print(f"❌ ERREUR : Rangs non définis pour les cohortes : {cohortes_sans_rangs}")
    print(f"   Veuillez ajouter ces cohortes dans RANGS_PAR_COHORTE")
    exit(1)
else:
    print(f"✓ Configuration validée : tous les rangs sont définis")

# ==============================================================================
# 📅 CHARGEMENT DU FICHIER DE RÉFÉRENCE DES SEMAINES
# ==============================================================================

print("\n📅 Chargement du fichier de référence des semaines...")

try:
    # Charger le fichier Excel avec la feuille Semaine_ref_trim
    df_semaine_ref = pd.read_excel(
        FICHIER_SEMAINE_REF,
        sheet_name='Semaine_ref_trim'
    )
    
    print(f"✓ Fichier de référence chargé : {len(df_semaine_ref)} enregistrements")
    print(f"   Colonnes : {list(df_semaine_ref.columns)}")
    
    # Vérifier les colonnes nécessaires
    colonnes_requises = ['Trimestre', 'Numero_semaine', 'DateJ7', 'Date1', 'Date2']
    colonnes_manquantes = [col for col in colonnes_requises if col not in df_semaine_ref.columns]
    
    if colonnes_manquantes:
        print(f"❌ ERREUR : Colonnes manquantes dans le fichier : {colonnes_manquantes}")
        exit(1)
    
    # Afficher un aperçu
    print("\n   Aperçu des données de référence :")
    for trimestre in df_semaine_ref['Trimestre'].unique():
        nb_semaines = len(df_semaine_ref[df_semaine_ref['Trimestre'] == trimestre])
        print(f"      • {trimestre} : {nb_semaines} semaines")
    
except FileNotFoundError:
    print(f"❌ ERREUR : Fichier de référence introuvable : {FICHIER_SEMAINE_REF}")
    print("   Veuillez vérifier le chemin du fichier.")
    exit(1)
except Exception as e:
    print(f"❌ ERREUR lors du chargement du fichier de référence : {str(e)}")
    exit(1)

# ==============================================================================
# 🏷️ CHARGEMENT DES FICHIERS DE LABELS (RÉGION, DISTRICT, DÉPARTEMENT, SP)
# ==============================================================================

print("\n🏷️  Chargement des fichiers de labels géographiques...")

# Dictionnaire pour stocker les tables de correspondance
dict_labels = {}

# Liste des feuilles à charger
feuilles_labels = {
    'label_region': ('HH2', 'label_HH2'),
    'label_district': ('HH1', 'label_HH1'),
    'label_departement': ('HH3', 'label_HH3'),
    'label_sp': ('HH4', 'label_HH4')
}

nb_feuilles_chargees = 0

for nom_feuille, (col_code, col_label) in feuilles_labels.items():
    try:
        # Charger la feuille
        df_label = pd.read_excel(
            FICHIER_SEMAINE_REF,
            sheet_name=nom_feuille
        )
        
        # Vérifier que les colonnes existent
        if col_code in df_label.columns and col_label in df_label.columns:
            
            # 🔧 NETTOYAGE DES DONNÉES
            # Supprimer les lignes avec codes manquants
            df_label = df_label.dropna(subset=[col_code])
            
            # Supprimer les doublons en gardant la première occurrence
            nb_avant = len(df_label)
            df_label = df_label.drop_duplicates(subset=[col_code], keep='first')
            nb_doublons = nb_avant - len(df_label)
            
            # Créer un dictionnaire de correspondance code -> label
            dict_labels[col_code] = dict(zip(df_label[col_code], df_label[col_label]))
            
            nb_feuilles_chargees += 1
            print(f"   ✓ {nom_feuille} chargée : {len(df_label)} correspondances ({col_code} → {col_label})")
            
            if nb_doublons > 0:
                print(f"      ⚠️  {nb_doublons} doublons supprimés (première occurrence conservée)")
        else:
            print(f"   ⚠️  {nom_feuille} : colonnes manquantes ({col_code} ou {col_label})")
    
    except Exception as e:
        print(f"   ⚠️  Erreur lors du chargement de {nom_feuille} : {str(e)}")

if nb_feuilles_chargees == 0:
    print(f"   ⚠️  ATTENTION : Aucune feuille de labels n'a pu être chargée")
    print(f"   Les variables de labels ne seront pas créées")
else:
    print(f"\n   📊 Total : {nb_feuilles_chargees}/{len(feuilles_labels)} feuilles de labels chargées")

# ==============================================================================
# 🔧 CRÉATION DU DOSSIER DE SORTIE
# ==============================================================================

if not os.path.exists(DOSSIER_SORTIE):
    os.makedirs(DOSSIER_SORTIE)
    print(f"\n✓ Dossier de sortie créé : {DOSSIER_SORTIE}")
else:
    print(f"\n✓ Dossier de sortie existant : {DOSSIER_SORTIE}")

# ==============================================================================
# 📂 FONCTION : CHARGER LES DONNÉES D'UNE COHORTE
# ==============================================================================

def charger_cohorte(trimestre):
    """
    Charge les données (ménage + membres) d'un trimestre donné.
    
    Parameters:
        trimestre (str): Nom du trimestre (ex: "T3_2024")
    
    Returns:
        tuple: (DataFrame ménage, DataFrame membres) ou (None, None) si erreur
    """
    print(f"\n📥 Chargement de la cohorte : {trimestre}")
    
    # Construire les chemins
    dossier = os.path.join(REPERTOIRE_BASE, f"Base_brute_{trimestre}")
    fichier_menage = os.path.join(dossier, NOMS_FICHIERS[trimestre])
    fichier_membres = os.path.join(dossier, "membres.dta")
    
    # Vérifier l'existence des fichiers
    if not os.path.exists(fichier_menage):
        print(f"   ❌ ERREUR : Fichier ménage introuvable : {fichier_menage}")
        return None, None
    
    if not os.path.exists(fichier_membres):
        print(f"   ❌ ERREUR : Fichier membres introuvable : {fichier_membres}")
        return None, None
    
    try:
        # Charger la base ménage
        menage = pd.read_stata(
            fichier_menage,
            convert_categoricals=False,
            convert_missing=False,
            preserve_dtypes=False
        )
        print(f"   ✓ Ménages chargés : {len(menage)} observations")
        
        # Charger la base membres
        membres = pd.read_stata(
            fichier_membres,
            convert_categoricals=False,
            convert_missing=False,
            preserve_dtypes=False
        )
        print(f"   ✓ Membres chargés : {len(membres)} observations")
        
        # Ajouter une colonne pour identifier la cohorte d'origine
        menage['cohorte_origine'] = trimestre
        membres['cohorte_origine'] = trimestre
        
        return menage, membres
        
    except Exception as e:
        print(f"   ❌ ERREUR lors du chargement : {str(e)}")
        return None, None

# ==============================================================================
# 📊 CHARGEMENT ET CONSOLIDATION DE TOUTES LES COHORTES
# ==============================================================================

print("\n" + "=" * 70)
print("CHARGEMENT DES COHORTES")
print("=" * 70)

# Listes pour stocker les données de toutes les cohortes
liste_menages = []
liste_membres = []

# Compteurs
nb_cohortes_chargees = 0
nb_menages_total = 0
nb_membres_total = 0

# Charger chaque cohorte
for cohorte in COHORTES_REINTERROGATION:
    menage, membres = charger_cohorte(cohorte)
    
    if menage is not None and membres is not None:
        liste_menages.append(menage)
        liste_membres.append(membres)
        nb_cohortes_chargees += 1
        nb_menages_total += len(menage)
        nb_membres_total += len(membres)

# Vérifier qu'au moins une cohorte a été chargée
if nb_cohortes_chargees == 0:
    print("\n❌ ERREUR CRITIQUE : Aucune cohorte n'a pu être chargée.")
    print("   Vérifiez les chemins et les noms de fichiers.")
    exit(1)

# Consolider toutes les cohortes en un seul DataFrame
print(f"\n📊 CONSOLIDATION DES DONNÉES")
print(f"   Cohortes chargées : {nb_cohortes_chargees}/{len(COHORTES_REINTERROGATION)}")

Menage = pd.concat(liste_menages, ignore_index=True)
Membres = pd.concat(liste_membres, ignore_index=True)

print(f"   ✓ Total ménages : {len(Menage)}")
print(f"   ✓ Total membres : {len(Membres)}")

# ==============================================================================
# 🔢 ATTRIBUTION DES RANGS D'INTERROGATION POUR LES MÉNAGES
# ==============================================================================

print("\n" + "=" * 70)
print("ATTRIBUTION DES RANGS D'INTERROGATION - MÉNAGES")
print("=" * 70)

# Initialiser les colonnes de rangs
Menage['rgmen'] = None
Menage['rghab'] = None
Menage['rang_last_trim'] = None

# Compteurs pour les statistiques
nb_menages_avec_rangs = 0
nb_menages_sans_rangs = 0

# Attribuer les rangs en fonction de la cohorte d'origine
for idx, row in Menage.iterrows():
    cohorte = row['cohorte_origine']
    
    if cohorte in RANGS_PAR_COHORTE:
        # Récupérer les rangs pour cette cohorte
        rgmen_val = RANGS_PAR_COHORTE[cohorte]['rgmen']
        rghab_val = RANGS_PAR_COHORTE[cohorte]['rghab']
        
        # Attribuer les valeurs
        Menage.at[idx, 'rgmen'] = rgmen_val
        Menage.at[idx, 'rghab'] = rghab_val
        Menage.at[idx, 'rang_last_trim'] = rgmen_val - 1
        
        nb_menages_avec_rangs += 1
    else:
        nb_menages_sans_rangs += 1
        print(f"   ⚠️  Cohorte non configurée : {cohorte} (ménage {row['interview__key']})")

print(f"\n✓ Rangs attribués : {nb_menages_avec_rangs} ménages")

if nb_menages_sans_rangs > 0:
    print(f"⚠️  ATTENTION : {nb_menages_sans_rangs} ménages sans rangs")

# Afficher un résumé par cohorte
print(f"\n📊 Répartition des rangs par cohorte (ménages) :")
stats_rangs = Menage.groupby('cohorte_origine').agg({
    'rgmen': 'first',
    'rghab': 'first',
    'rang_last_trim': 'first',
    'interview__key': 'count'
}).rename(columns={'interview__key': 'nb_menages'})

for cohorte, row in stats_rangs.iterrows():
    print(f"   • {cohorte} : rgmen={int(row['rgmen'])}, rghab={int(row['rghab'])}, "
          f"rang_last_trim={int(row['rang_last_trim'])} | {int(row['nb_menages'])} ménages")

# ==============================================================================
# 🔧 PRÉPARATION DES MÉTADONNÉES SURVEY SOLUTIONS
# ==============================================================================

print("\n" + "=" * 70)
print("PRÉPARATION DES MÉTADONNÉES")
print("=" * 70)

# Affecter un agent responsable par défaut (à personnaliser selon l'affectation réelle)
Menage['_responsible'] = 'AgentReinterrogation_' + TRIMESTRE_ACTUEL

# Quantité = 1 signifie qu'il faut interroger ce ménage une fois
Menage['_quantity'] = 1

print(f"✓ Agent responsable : {Menage['_responsible'].iloc[0]}")

# ==============================================================================
# 🔑 CRÉATION DE LA CLÉ D'IDENTIFICATION UNIQUE
# ==============================================================================

print("\n📝 Création des clés d'identification...")

# Construire un identifiant unique pour retrouver le ménage au Passage 2
# Format : DISTRICT_SOUS-PREFECTURE_LOCALITE+QUARTIER+T+TRIMESTRE+ANNEE+RANG_MENAGE
Menage['V1interviewkey1er'] = (
    Menage['HH4'].astype(str) + "_" +           # District
    Menage['HH8'].astype(str) + "_" +           # Sous-préfecture
    Menage['HH7'].astype(str) +                 # Localité
    Menage['HH7B'].astype(str) + 'T' +          # Quartier
    Menage['trimestreencours'].astype(str) +     # Trimestre d'origine
    Menage['annee'].astype(str) +               # Année d'origine
    Menage['rghab'].astype(str) + "_" +         # Rang habitation
    Menage['HH9_1'].astype(str)                 # Numéro de porte
)

print(f"✓ Clés créées pour {len(Menage)} ménages")

# ==============================================================================
# 🔄 FUSION MEMBRES ET MÉNAGE
# ==============================================================================

print("\n🔗 Fusion des données membres et ménages...")

MembresVF = pd.merge(Membres, Menage, on='interview__key', how='left')

print(f"✓ Fusion complétée : {len(MembresVF)} lignes")

# ==============================================================================
# 📅 MISE À JOUR DES VARIABLES TEMPORELLES
# ==============================================================================

print("\n📅 Mise à jour des variables temporelles...")

# Variables de contexte temporel (trimestre actuel de réinterrogation)
Menage['trimestreencours'] = TRIMESTRE_NUMERO
Menage['mois_en_cours'] = MOIS_EN_COURS
Menage['annee'] = ANNEE_ACTUELLE

# Variables de traçabilité entre les passages
Menage['V1interviewkey'] = Menage['interview__key']              # Clé Passage 1
Menage['V1interviewkey_nextTrim'] = Menage['interview__key']     # Clé pour suivi

print(f"✓ Trimestre : {TRIMESTRE_NUMERO}, Année : {ANNEE_ACTUELLE}")

# ==============================================================================
# 💾 PRÉCHARGEMENT DES VARIABLES DU PASSAGE 1 (PRÉFIXE V1)
# ==============================================================================

print("\n💾 Préchargement des variables du Passage 1...")

# Ces variables commençant par "V1" stockent les réponses du Passage 1
# pour permettre la validation et la cohérence lors du Passage 2

# Variables temporelles et métadonnées
Menage['V1hha'] = Menage['hha']                     # Heure début interview P1
Menage['V1Q2'] = Menage['Q2']                       # Question 2
Menage['V1Q2_aut'] = Menage['Q2_aut']              # Question 2 (autre)

# Coordonnées GPS du ménage
Menage['V1GPS_longitude'] = Menage['GPS__Longitude']
Menage['V1GPS_Lattitude'] = Menage['GPS__Latitude']

# Informations sur le chef de ménage
Menage['V1nom_prenom_cm'] = Menage['nom_prenom_cm']

# Variables d'identification du logement
Menage['V1HH10_1'] = Menage['HH10_1']              # Type de logement
Menage['V1HH10_2'] = Menage['HH10_2']              # Statut d'occupation

# Informations de contact
Menage['V1HH9_1'] = Menage['HH9_1']                # Numéro de téléphone
Menage['V1HH9'] = Menage['HH9']                    # Téléphone disponible (oui/non)
Menage['V1Q1_0'] = Menage['Q1_0']                  # Contact alternatif

# Variables complémentaires
Menage['V1HH13A'] = Menage['HH13A']                # Agent enquêteur
Menage['V1HH10_1_1a'] = Menage['HH10_1_1a']        # Précision type logement
Menage['V1HH10_2_1'] = Menage['HH10_2_1']          # Précision statut occupation
Menage['V1HH13B'] = Menage['HH13B']                # Superviseur

print(f"✓ Variables préchargées")

# ==============================================================================
# 🏷️ AJOUT DES LABELS GÉOGRAPHIQUES (HH1_label, HH2_label, HH3_label, HH4_label)
# ==============================================================================

print("\n🏷️  Ajout des labels géographiques...")

# Liste des variables à labelliser
variables_a_labelliser = ['HH1', 'HH2', 'HH3', 'HH4']

nb_labels_ajoutes = 0

for var in variables_a_labelliser:
    nom_label = f"{var}_label"
    
    # Vérifier si la variable existe dans Menage
    if var in Menage.columns:
        # Vérifier si on a le dictionnaire de correspondance
        if var in dict_labels:
            # Créer la variable label en mappant les codes
            Menage[nom_label] = Menage[var].map(dict_labels[var])
            
            # Compter les valeurs non trouvées
            nb_non_trouves = Menage[nom_label].isna().sum()
            nb_trouves = len(Menage) - nb_non_trouves
            
            print(f"   ✓ {nom_label} créée : {nb_trouves}/{len(Menage)} correspondances trouvées")
            
            if nb_non_trouves > 0:
                print(f"      ⚠️  {nb_non_trouves} codes sans correspondance dans le fichier de labels")
            
            nb_labels_ajoutes += 1
        else:
            print(f"   ⚠️  {nom_label} : dictionnaire de correspondance non disponible")
            Menage[nom_label] = None
    else:
        print(f"   ⚠️  {var} : variable non trouvée dans les données ménage")
        Menage[nom_label] = None

if nb_labels_ajoutes > 0:
    print(f"\n   📊 Total : {nb_labels_ajoutes}/{len(variables_a_labelliser)} variables de labels créées")
    
    # Afficher un échantillon
    print(f"\n   Échantillon des labels (1 premier ménage) :")
    echantillon_labels = Menage[['interview__key', 'HH1', 'HH1_label', 'HH2', 'HH2_label', 
                                   'HH3', 'HH3_label', 'HH4', 'HH4_label']].head(1)
    
    for idx, row in echantillon_labels.iterrows():
        print(f"      Ménage {row['interview__key'][:15]}...")
        if pd.notna(row['HH1']):
            print(f"         District (HH1): {row['HH1']} → {row['HH1_label']}")
        if pd.notna(row['HH2']):
            print(f"         Région (HH2): {row['HH2']} → {row['HH2_label']}")
        if pd.notna(row['HH3']):
            print(f"         Département (HH3): {row['HH3']} → {row['HH3_label']}")
        if pd.notna(row['HH4']):
            print(f"         Sous-préf. (HH4): {row['HH4']} → {row['HH4_label']}")
else:
    print(f"   ⚠️  Aucune variable de label n'a pu être créée")

# ==============================================================================
# 📅 AJOUT DE LA VARIABLE DateJ7 ET DÉTERMINATION DE LA SEMAINE + DATES
# ==============================================================================

print("\n📅 Détermination de la semaine de référence et mise à jour des dates...")

# Vérifier si la variable DateJ7 existe dans la base ménage
if 'DateJ7' in Menage.columns:
    # Précharger DateJ7 du Passage 1
    Menage['V1DateJ7'] = Menage['DateJ7']
    print(f"✓ Variable V1DateJ7 créée (DateJ7 du Passage 1)")
    
    # Initialiser les colonnes
    Menage['Semaine_ref'] = None
    
    # ÉTAPE 1 : Déterminer la semaine de référence pour chaque ménage
    # (basé sur la cohorte d'origine et DateJ7)
    nb_semaines_trouvees = 0
    nb_semaines_non_trouvees = 0
    
    print(f"\n   Étape 1 : Détermination des semaines de référence...")
    
    for idx, row in Menage.iterrows():
        cohorte_origine = row['cohorte_origine']
        datej7_menage = row['DateJ7']
        
        # Chercher la correspondance dans le fichier de référence
        correspondance = df_semaine_ref[
            (df_semaine_ref['Trimestre'] == cohorte_origine) &
            (df_semaine_ref['DateJ7'] == datej7_menage)
        ]
        
        if len(correspondance) > 0:
            semaine_ref = correspondance.iloc[0]['Numero_semaine']
            Menage.at[idx, 'Semaine_ref'] = semaine_ref
            nb_semaines_trouvees += 1
        else:
            nb_semaines_non_trouvees += 1
    
    print(f"   ✓ Semaines déterminées : {nb_semaines_trouvees} / {len(Menage)} ménages")
    
    if nb_semaines_non_trouvees > 0:
        print(f"   ⚠️  ATTENTION : {nb_semaines_non_trouvees} ménages sans correspondance")
    
    # Afficher la répartition par semaine
    print(f"\n   Répartition des ménages par semaine :")
    repartition_semaines = Menage['Semaine_ref'].value_counts().sort_index()
    for semaine, nb in repartition_semaines.items():
        if pd.notna(semaine):
            print(f"      • {semaine} : {nb} ménages")
    
    # ÉTAPE 2 : Attribuer les dates Date1 et Date2 du TRIMESTRE ACTUEL
    # (basé sur le trimestre de réinterrogation et la semaine de référence)
    print(f"\n   Étape 2 : Attribution des dates du trimestre actuel ({TRIMESTRE_ACTUEL})...")
    
    Menage['Date1'] = None
    Menage['Date2'] = None
    
    nb_dates_mises_a_jour = 0
    nb_dates_non_trouvees = 0
    
    for idx, row in Menage.iterrows():
        semaine_ref = row['Semaine_ref']
        
        if pd.notna(semaine_ref):
            # Chercher les dates dans le fichier de référence pour le TRIMESTRE ACTUEL
            correspondance_dates = df_semaine_ref[
                (df_semaine_ref['Trimestre'] == TRIMESTRE_ACTUEL) &
                (df_semaine_ref['Numero_semaine'] == semaine_ref)
            ]
            
            if len(correspondance_dates) > 0:
                date1_ref = correspondance_dates.iloc[0]['Date1']
                date2_ref = correspondance_dates.iloc[0]['Date2']
                
                Menage.at[idx, 'Date1'] = date1_ref
                Menage.at[idx, 'Date2'] = date2_ref
                nb_dates_mises_a_jour += 1
            else:
                nb_dates_non_trouvees += 1
    
    print(f"   ✓ Dates mises à jour (Date1, Date2) : {nb_dates_mises_a_jour} ménages")
    
    if nb_dates_non_trouvees > 0:
        print(f"   ⚠️  ATTENTION : {nb_dates_non_trouvees} ménages sans dates")
        print(f"      Vérifiez que le fichier Semaine_ref.xlsx contient bien toutes les semaines pour {TRIMESTRE_ACTUEL}")
    
    # Afficher un échantillon des dates mises à jour
    print(f"\n   Échantillon des dates attribuées (2 premiers ménages) :")
    echantillon = Menage[['interview__key', 'cohorte_origine', 'Semaine_ref', 'Date1', 'Date2']].head(2)
    for idx, row in echantillon.iterrows():
        if pd.notna(row['Semaine_ref']):
            print(f"      {row['interview__key'][:15]}... | Cohorte: {row['cohorte_origine']} | {row['Semaine_ref']} | Dates {TRIMESTRE_ACTUEL}: {row['Date1']} → {row['Date2']}")
    
else:
    print(f"⚠️  ATTENTION : Variable 'DateJ7' non trouvée dans les données ménage")
    print(f"   Les variables 'Semaine_ref', 'Date1' et 'Date2' ne pourront pas être créées")
    Menage['V1DateJ7'] = None
    Menage['Semaine_ref'] = None
    Menage['Date1'] = None
    Menage['Date2'] = None


# ==============================================================================
# 🔢 CRÉATION DES VARIABLES ord_sem ET HH01
# ==============================================================================

print("\n🔢 Création des variables ord_sem et HH01...")

# Vérifier que Semaine_ref existe avant de créer ord_sem et HH01
if 'Semaine_ref' not in Menage.columns or Menage['Semaine_ref'].isna().all():
    print(f"   ⚠️  ATTENTION : Semaine_ref non disponible")
    print(f"   ⚠️  Les variables ord_sem et HH01 ne pourront pas être créées correctement")
    Menage['ord_sem'] = ""
    Menage['HH01'] = ""
else:
    # Générer une variable aléatoire de 8 chiffres UNIQUE par interview__key
    np.random.seed(42)  # Pour la reproductibilité (retirer pour du vrai aléatoire)
    
    # Obtenir les interview__key uniques
    interview_keys_uniques = Menage['interview__key'].unique()
    
    # Créer un dictionnaire de correspondance : interview__key → code aléatoire 8 chiffres
    dict_code_aleatoire = {}
    for key in interview_keys_uniques:
        # Générer un nombre aléatoire entre 10000000 et 99999999 (8 chiffres)
        code_aleatoire = np.random.randint(10000000, 100000000)
        dict_code_aleatoire[key] = code_aleatoire
    
    # Appliquer le mapping pour créer la variable aléatoire
    Menage['Variable_aleatoire'] = Menage['interview__key'].map(dict_code_aleatoire)
    
    print(f"   ✓ Variable aléatoire de 8 chiffres créée pour {len(dict_code_aleatoire)} ménages uniques")
    if len(dict_code_aleatoire) > 0:
        print(f"   ✓ Exemple : interview__key {list(dict_code_aleatoire.keys())[0][:15]}... → {list(dict_code_aleatoire.values())[0]}")
    
    # 1. CONSTRUCTION DE ord_sem
    # Format : "Tele_" + Semaine_ref + "_" + TRIMESTRE_ACTUEL + "_" + Variable_aleatoire
    Menage['ord_sem'] = (
        "Tele_" + 
        Menage['Semaine_ref'].astype(str) + 
        f"_{TRIMESTRE_ACTUEL}_" + 
        Menage['Variable_aleatoire'].astype(str)
    )
    
    print(f"   ✓ Variable ord_sem créée")
    if len(Menage) > 0 and pd.notna(Menage['ord_sem'].iloc[0]):
        print(f"   ✓ Exemple : {Menage['ord_sem'].iloc[0]}")
    
    # 2. CONSTRUCTION DE HH01
    # Format : HH8A + HH8 + "_" + Semaine_ref + "_" + TRIMESTRE_ACTUEL + "_" + Variable_aleatoire
    Menage['HH01'] = (
        Menage['HH8A'].astype(str) + 
        "_" +
        Menage['HH8'].astype(str) + 
        "_" + 
        Menage['Semaine_ref'].astype(str) + 
        f"_{TRIMESTRE_ACTUEL}_" + 
        Menage['Variable_aleatoire'].astype(str)
    )
    
    print(f"   ✓ Variable HH01 créée")
    if len(Menage) > 0 and pd.notna(Menage['HH01'].iloc[0]):
        print(f"   ✓ Exemple : {Menage['HH01'].iloc[0]}")
    
    # Afficher un échantillon des résultats
    print(f"\n   📋 Échantillon des variables créées (2 premiers ménages) :")
    colonnes_echantillon = ['interview__key', 'Semaine_ref', 'Variable_aleatoire', 'ord_sem', 'HH01']
    # Vérifier que toutes les colonnes existent
    colonnes_disponibles = [col for col in colonnes_echantillon if col in Menage.columns]
    if len(colonnes_disponibles) > 0:
        echantillon = Menage[colonnes_disponibles].head(2)
        for idx, row in echantillon.iterrows():
            print(f"      Ménage {row['interview__key'][:15]}...")
            if 'Semaine_ref' in row:
                print(f"         Semaine_ref      : {row['Semaine_ref']}")
            if 'Variable_aleatoire' in row:
                print(f"         Code aléatoire   : {row['Variable_aleatoire']}")
            if 'ord_sem' in row:
                print(f"         ord_sem          : {row['ord_sem']}")
            if 'HH01' in row:
                print(f"         HH01             : {row['HH01']}")
            print()
    
    # Supprimer la variable temporaire Variable_aleatoire (optionnel)
    Menage.drop(columns=['Variable_aleatoire'], inplace=True)
    
    print(f"✓ Variables ord_sem et HH01 créées avec succès !")


# ==============================================================================
# 📋 CRÉATION DU FICHIER MÉNAGE FINAL
# ==============================================================================

print("\n📋 Création du fichier ménage...")

# Sélectionner les colonnes nécessaires pour le fichier ménage
colonnes_menage = [
    
    # Variables de labels géographiques
    'HH1_label', 'HH2_label', 'HH3_label', 'HH4_label', 'Semaine_ref', 'Reference',
    
    # Identifiants et métadonnées Survey Solutions 
    'interview__id','Cohorte','ord_sem','HH01','HH0','HH2A','HH1','HH2','HH3','HH4','HH6','HH8',
    
    'HH8A','HH7','HH7B','HH8B',
    
    # ✨ RANGS D'INTERROGATION
    'rghab', 'rgmen',
    
    # Contexte temporel
    'V1MODINTR','trimestreencours','mois_en_cours','annee',
    
    # Variables préchargées du Passage 1 et Clés de liaison entre passages et DateJ7
    'Date1','Date2','Reference','V1interviewkey','V1interviewkey_nextTrim','V1interviewkey1er','V1hha',
    
    'V1Q2','V1Q2_aut','V1GPS_longitude','V1GPS_Lattitude','V1nom_prenom_cm','V1HH10_1','V1HH10_2','V1HH9_1',
    
    'V1HH9','V1Q1_0','V1HH13A','V1HH10_1_1a','V1HH10_2_1','V1HH13B',
]

# Ajouter les colonnes M0__0 à M0__59 (composition du ménage)
colonnes_m0 = [f'M0__{i}' for i in range(60)]

# Ajouter le reste des colonnes (composition du ménage)
colonnes_m1 = [
    '_responsible','_quantity','GPS__Longitude','GPS__Latitude','interview__key','hh','hha','cohorte_origine',
]

colonnes_menage.extend(colonnes_m0)
colonnes_menage.extend(colonnes_m1)

# Filtrer pour ne garder que les colonnes existantes
colonnes_menage_existantes = [col for col in colonnes_menage if col in Menage.columns]

# Créer le dataframe final
MenageVF = Menage[colonnes_menage_existantes]

# Variable Cohorte mise à jour avec variable cohorte_origine
MenageVF['Cohorte'] = MenageVF['cohorte_origine']

# Exporter vers Excel et CSV
fichier_menage_xlsx = os.path.join(DOSSIER_SORTIE, "QX_EEC_VF.xlsx")
fichier_menage_csv = os.path.join(DOSSIER_SORTIE, "QX_EEC_VF.csv")

MenageVF.to_excel(fichier_menage_xlsx, index=False)
MenageVF.to_csv(fichier_menage_csv, index=False)

print(f"✓ Fichier ménage créé : {len(MenageVF)} ménages")
print(f"   Excel : {fichier_menage_xlsx}")
print(f"   CSV   : {fichier_menage_csv}")

# ==============================================================================
# 👥 PRÉPARATION DU FICHIER MEMBRES
# ==============================================================================

print("\n👥 Préparation du fichier membres...")

# 🔍 DIAGNOSTIC : Vérifier les colonnes dupliquées après la fusion
print(f"\n   Colonnes dans MembresVF : {len(MembresVF.columns)}")
colonnes_dupliquees = [col for col in MembresVF.columns if col.endswith('_x') or col.endswith('_y')]
if colonnes_dupliquees:
    print(f"   ⚠️  Colonnes dupliquées détectées : {colonnes_dupliquees[:10]}...")

# 🔧 RÉSOLUTION : Nettoyer les colonnes dupliquées
# Si interview__id existe en doublon, on garde la version du ménage (_y)
if 'interview__id_x' in MembresVF.columns and 'interview__id_y' in MembresVF.columns:
    print(f"   🔧 Résolution des doublons interview__id...")
    MembresVF['interview__id'] = MembresVF['interview__id_y']
    MembresVF = MembresVF.drop(columns=['interview__id_x', 'interview__id_y'])
    print(f"   ✓ interview__id nettoyée")
elif 'interview__id_x' in MembresVF.columns:
    MembresVF['interview__id'] = MembresVF['interview__id_x']
    MembresVF = MembresVF.drop(columns=['interview__id_x'])

# Nettoyer les autres colonnes dupliquées
for col_base in ['cohorte_origine', 'V1interviewkey1er', 'rgmen', 'rghab', 'rang_last_trim']:
    col_x = f"{col_base}_x"
    col_y = f"{col_base}_y"
    
    if col_x in MembresVF.columns and col_y in MembresVF.columns:
        # Privilégier la version du ménage (_y) si elle existe
        MembresVF[col_base] = MembresVF[col_y].fillna(MembresVF[col_x])
        MembresVF = MembresVF.drop(columns=[col_x, col_y])
        print(f"   ✓ {col_base} nettoyée (fusion _x et _y)")
    elif col_x in MembresVF.columns:
        MembresVF[col_base] = MembresVF[col_x]
        MembresVF = MembresVF.drop(columns=[col_x])
    elif col_y in MembresVF.columns:
        MembresVF[col_base] = MembresVF[col_y]
        MembresVF = MembresVF.drop(columns=[col_y])

# Vérifier que cohorte_origine existe maintenant
if 'cohorte_origine' not in MembresVF.columns:
    print(f"   ⚠️  ATTENTION : cohorte_origine toujours absente après nettoyage")
else:
    print(f"   ✓ cohorte_origine présente : {MembresVF['cohorte_origine'].nunique()} cohortes")

# Vérifier que les colonnes nécessaires existent
colonnes_critiques = ['interview__id', 'membres__id', 'rgmen', 'V1interviewkey1er']
for col in colonnes_critiques:
    if col not in MembresVF.columns:
        print(f"   ⚠️  ATTENTION : Colonne manquante : {col}")

# Créer une clé unique pour chaque individu
if all(col in MembresVF.columns for col in ['V1interviewkey1er', 'rgmen', 'membres__id']):
    MembresVF['cle_individu'] = (
        MembresVF['V1interviewkey1er'].astype(str) +
        MembresVF['rgmen'].astype(str) + "1_" +
        MembresVF['membres__id'].astype(str)
    )
    print(f"   ✓ cle_individu créée")
else:
    print(f"   ⚠️  Impossible de créer cle_individu : colonnes manquantes")
    MembresVF['cle_individu'] = None

# ==============================================================================
# 🔢 ATTRIBUTION DES RANGS D'INTERROGATION POUR LES INDIVIDUS
# ==============================================================================

print("\n" + "=" * 70)
print("ATTRIBUTION DES RANGS D'INTERROGATION - INDIVIDUS")
print("=" * 70)

# Initialiser la colonne rang_ind
MembresVF['rang_ind'] = None

# Compteurs
nb_individus_avec_rangs = 0
nb_individus_sans_rangs = 0

# Attribuer rang_ind en fonction de la cohorte d'origine
for idx, row in MembresVF.iterrows():
    cohorte = row['cohorte_origine']
    
    if cohorte in RANGS_PAR_COHORTE:
        rang_ind_val = RANGS_PAR_COHORTE[cohorte]['rang_ind']
        MembresVF.at[idx, 'rang_ind'] = rang_ind_val
        nb_individus_avec_rangs += 1
    else:
        nb_individus_sans_rangs += 1

print(f"\n✓ Rangs attribués : {nb_individus_avec_rangs} individus")

if nb_individus_sans_rangs > 0:
    print(f"⚠️  ATTENTION : {nb_individus_sans_rangs} individus sans rangs")

# Afficher un résumé par cohorte
print(f"\n📊 Répartition des rangs par cohorte (individus) :")
stats_rangs_ind = MembresVF.groupby('cohorte_origine').agg({
    'rang_ind': 'first',
    'membres__id': 'count'
}).rename(columns={'membres__id': 'nb_individus'})

for cohorte, row in stats_rangs_ind.iterrows():
    print(f"   • {cohorte} : rang_ind={int(row['rang_ind'])} | {int(row['nb_individus'])} individus")

# Calculer rang_last_trim pour les individus (basé sur rang_ind)
print(f"\n🔢 Calcul de rang_last_trim pour les individus...")
MembresVF['rang_last_trim'] = MembresVF['rang_ind'] - 1
print(f"✓ rang_last_trim calculé")

# Afficher un échantillon
print(f"\n   Échantillon (2 premiers) :")
echantillon = MembresVF[['membres__id', 'cohorte_origine', 'rang_ind', 'rang_last_trim']].head(2)
for idx, row in echantillon.iterrows():
    print(f"      Membre {row['membres__id']} | {row['cohorte_origine']} | "
          f"rang_ind={int(row['rang_ind'])}, rang_last_trim={int(row['rang_last_trim'])}")

# Variables de suivi longitudinal des individus
MembresVF['membre_id_v1'] = MembresVF['membres__id']
MembresVF['rangind_1er'] = MembresVF['membres__id']
MembresVF['membre_id_v1_IND'] = MembresVF['membre_id_v1_IND']

# Préchargement des variables individuelles du Passage 1
variables_precharge = {
    'V1M4': 'M4',
    'V1M9': 'M9',
    'V1M12': 'M12',
    'V1EF1': 'EF1',
    'V1FP1': 'FP1',
    'V1EP1a': 'EP1a'
}

for var_dest, var_source in variables_precharge.items():
    if var_source in MembresVF.columns:
        MembresVF[var_dest] = MembresVF[var_source]
    else:
        MembresVF[var_dest] = None
        print(f"   ⚠️  Variable {var_source} non trouvée pour créer {var_dest}")

# Variables de contact et localisation (module Q1)
variables_q1 = ['Q1_01', 'Q1_1', 'Q1_4', 'Q1_7', 'Q1_9',
                'Q1_10__1', 'Q1_10__2', 'Q1_10__3', 'Q1_10__4',
                'Q1_12',
                'Q1_13__1', 'Q1_13__2', 'Q1_13__3', 'Q1_13__4']

for var in variables_q1:
    if var in MembresVF.columns:
        MembresVF[f'V1{var}'] = MembresVF[var]

print(f"\n✓ Variables préchargées pour {len(MembresVF)} individus")

# ==============================================================================
# 🔍 FILTRAGE : CONSERVER UNIQUEMENT LES RÉSIDENTS ET LES MÉNAGES VALIDES
# ==============================================================================

print("\n🔍 Filtrage des données membres...")

avant_filtrage = len(MembresVF)

# 1. Filtrer pour ne garder que les interview__key présents dans MenageVF
interview_keys_valides = set(MenageVF['interview__key'].dropna())
print(f"\n   Nombre de ménages valides : {len(interview_keys_valides)}")

MembresVF = MembresVF[MembresVF['interview__key'].isin(interview_keys_valides)]
print(f"   ✓ Après filtrage par interview__key : {len(MembresVF)} / {avant_filtrage} individus")

# 2. Filtrer pour ne garder que les résidents (Statut_Res = 1)
if 'Statut_Res' in MembresVF.columns:
    avant_filtrage_residents = len(MembresVF)
    MembresVF = MembresVF[MembresVF['Statut_Res'] == 1]
    print(f"   ✓ Résidents conservés : {len(MembresVF)} / {avant_filtrage_residents} individus")
else:
    print(f"   ⚠️  ATTENTION : Variable 'Statut_Res' non trouvée")
    print(f"      Tous les individus sont conservés (pas de filtrage par statut de résidence)")

print(f"\n   📊 Total final : {len(MembresVF)} individus retenus")

"""
# ==============================================================================
# 🔄 RENOMMAGE DE cohorte_origine EN Cohorte1 DANS MembresVF
# ==============================================================================

print("\n🔄 Renommage de cohorte_origine en Cohorte1 dans MembresVF...")

if 'cohorte_origine' in MembresVF.columns:
    MembresVF.rename(columns={'cohorte_origine': 'Cohorte1'}, inplace=True)
    print(f"   ✓ Variable cohorte_origine renommée en Cohorte1")
    print(f"   ✓ Valeurs : {MembresVF['Cohorte1'].unique()}")
else:
    print(f"   ⚠️  ATTENTION : Variable cohorte_origine non trouvée dans MembresVF")
    print(f"   ⚠️  Impossible de renommer en Cohorte1")
"""

# ==============================================================================
# 🔄 RENOMMAGE DE cohorte_origine EN Cohorte1 DANS MembresVF
# ==============================================================================

print("\n🔄 Renommage de cohorte_origine en Cohorte1 dans MembresVF...")

# 1. Supprimer d'abord toute colonne Cohorte1 existante (vide)
if 'Cohorte1' in MembresVF.columns:
    MembresVF = MembresVF.drop(columns=['Cohorte1'])
    print(f"   ✓ Ancienne colonne Cohorte1 (vide) supprimée")

# 2. Renommer cohorte_origine en Cohorte1
if 'cohorte_origine' in MembresVF.columns:
    MembresVF.rename(columns={'cohorte_origine': 'Cohorte1'}, inplace=True)
    print(f"   ✓ Variable cohorte_origine renommée en Cohorte1")
    print(f"   ✓ Nombre de valeurs non-nulles : {MembresVF['Cohorte1'].notna().sum()}")
else:
    print(f"   ⚠️  ATTENTION : Variable cohorte_origine non trouvée dans MembresVF")
    print(f"   ⚠️  Impossible de renommer en Cohorte1")
    
# ==============================================================================
# 📊 SÉLECTION DES COLONNES FINALES
# ==============================================================================

print("\n📊 Sélection des colonnes finales...")

colonnes_membres = [
    # Identifiants
    'membres__id', 'M0', 'Cohorte1',
    
    # ✨ Variables de suivi longitudinal (AVEC RANGS)
    'membre_id_v1', 'rangind_1er', 'rang_last_trim', 'cle_individu', 'rang_ind',
    
    # Variables préchargées du Passage 1
    'V1M4', 'V1M9', 'V1M12', 'membre_id_v1_IND', 
    'V1Q1_01', 'V1Q1_1', 'V1Q1_4', 'V1Q1_7', 'V1Q1_9',
    'V1Q1_10__1', 'V1Q1_10__2', 'V1Q1_10__3', 'V1Q1_10__4',
    'V1Q1_12',
    'V1Q1_13__1', 'V1Q1_13__2', 'V1Q1_13__3', 'V1Q1_13__4','V1EP1a',
    'interview__id',
    
    # Variables de contexte
    'AgeAnnee', 'hhb',
    'hha_FT', 'hha_SE', 'hha_EMP', 'hha_ES', 'hha_PL',
    'hhavf_C', 'hha_P',
    'M4Confirm', 'EN_EMP',
    
    # Variables supplémentaires
    'membre_id_v1A','membre_id_v1_INDA',
    'statut_MO', 'cle_individuA','V1interviewkey', 
    'V1interviewkey_nextTrim', 'V1interviewkey1er',
    'Statut_Res', 'hha_COMP'
]

# Filtrer pour ne garder que les colonnes existantes
colonnes_membres_existantes = [col for col in colonnes_membres if col in MembresVF.columns]
colonnes_manquantes = [col for col in colonnes_membres if col not in MembresVF.columns]

print(f"   ✓ Colonnes trouvées : {len(colonnes_membres_existantes)}/{len(colonnes_membres)}")
if colonnes_manquantes:
    print(f"   ⚠️  Colonnes manquantes : {colonnes_manquantes[:10]}")

MembresVF = MembresVF[colonnes_membres_existantes].copy()

# ==============================================================================
# 🔢 TRI ET NUMÉROTATION
# ==============================================================================

print("\n🔢 Tri et numérotation...")

# Vérifier que interview__id existe et est unique
if 'interview__id' not in MembresVF.columns:
    print(f"   ❌ ERREUR : interview__id manquante dans MembresVF")
else:
    if MembresVF['interview__id'].ndim != 1:
        print(f"   ❌ ERREUR : interview__id a {MembresVF['interview__id'].ndim} dimensions au lieu de 1")
    else:
        MembresVF = MembresVF.sort_values(by='membres__id')
        MembresVF['numero'] = MembresVF.groupby('interview__id').cumcount() + 1
        print(f"   ✓ Numérotation créée : {MembresVF['numero'].max()} membres maximum par ménage")

# ==============================================================================
# 💾 EXPORT DES FICHIERS MEMBRES
# ==============================================================================

print("\n💾 Export des fichiers membres...")

fichier_membres_xlsx = os.path.join(DOSSIER_SORTIE, "membres.xlsx")
fichier_membres_csv = os.path.join(DOSSIER_SORTIE, "membres.csv")

MembresVF.to_excel(fichier_membres_xlsx, index=False)
MembresVF.to_csv(fichier_membres_csv, index=False)

print(f"✓ Fichiers membres créés")
print(f"   Excel : {fichier_membres_xlsx}")
print(f"   CSV   : {fichier_membres_csv}")

# ==============================================================================
# 📈 STATISTIQUES FINALES PAR COHORTE
# ==============================================================================

nombre_menages = MenageVF['interview__id'].nunique()

print("\n" + "=" * 70)
print("📊 RÉSUMÉ FINAL")
print("=" * 70)
print(f"\n✅ Traitement terminé avec succès !")
print(f"\n📅 Trimestre de réinterrogation : {TRIMESTRE_ACTUEL}")
print(f"\n📊 Statistiques globales :")
print(f"   • Nombre total de ménages : {nombre_menages}")
print(f"   • Nombre total de résidents : {len(MembresVF)}")

print(f"\n📊 Répartition par cohorte d'origine :")
stats_cohortes_menage = MenageVF['cohorte_origine'].value_counts().sort_index()
for cohorte, nb in stats_cohortes_menage.items():
    rangs = RANGS_PAR_COHORTE.get(cohorte, {})
    print(f"   • {cohorte} : {nb} ménages | rgmen={rangs.get('rgmen','N/A')}, rang_ind={rangs.get('rang_ind','N/A')}")

print(f"\n📁 Fichiers générés dans : {DOSSIER_SORTIE}")
print(f"   ✓ QX_EEC_VF.xlsx / .csv (ménages)")
print(f"   ✓ membres.xlsx / .csv (individus)")

print("\n" + "=" * 70)
print("✅ PROGRAMME TERMINÉ")
print("=" * 70)
