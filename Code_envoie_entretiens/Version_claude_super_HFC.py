# -*- coding: utf-8 -*-
"""
REJET INTELLIGENT VERS RESPONSABLE ORIGINAL - VERSION CORRIGÉE
Étape 1 : Enrichit le fichier Excel avec les infos du responsable actuel
Étape 2 : Rejette uniquement les entretiens non rejetés vers leur responsable original
Auteur: mg.kouame
Date: Déc 2025
"""

import os
import re
import time
import csv
import traceback
from datetime import datetime
import pandas as pd
import ssaw
from ssaw import InterviewsApi, UsersApi

import os
#Affichage du repertoire de travail par defaut de Python
print(os.getcwd())

#Definition du nouvel envirionnement de travail
pwd=r"D:\Ancien_HFC\HFC\Base\Base_rejet\Base_globale\Agent_enquêteur\Erreur_DR\Code_python_rejet"
os.chdir(pwd)

# ----------------------------
# CONFIGURATION
# ----------------------------
SURVEY_URL = "http://154.68.47.214:9700"
WORKSPACE = "enemenage"

# Credentials
API_USER = "UserAPIEEC"
API_PASSWORD = "UserAPIEEC1"
HQ_USERNAME = "SupDaouda"
HQ_PASSWORD = "SupDaouda01"

# Chemins fichiers
PWD = r"D:\Ancien_HFC\HFC\Base\Base_rejet\Base_globale\Agent_enquêteur\Erreur_DR\Code_python_rejet"
os.chdir(PWD)

# NOUVEAU FICHIER
EXCEL_INPUT = "Base_rejet_Agent_enqueteur_V2.xlsx"
SHEET_INPUT = "Fichier_rejet_py"
EXCEL_OUTPUT = "Base_rejet_Agent_enqueteur_V2_ENRICHI.xlsx"

# Colonnes
COL_KEY = "INTERVIEW_KEY"
COL_ID = "INTERVIEW_ID"
COL_RESP_NAME = "Numero_agent_terrain"       # À créer (ResponsibleName)
COL_RESP_ID = "Numero_agent_terrain_ID"      # À créer (ResponsibleId)
COL_STATUS = "STATUS_ACTUEL"                  # À créer (statut de l'entretien)

# Logs
LOG_ENRICHMENT = "log_enrichment.csv"
LOG_REJECT = "log_reject.csv"
DEBUG_DIR = "debug_smart_reject"
os.makedirs(DEBUG_DIR, exist_ok=True)

# Paramètres
MAX_ROWS = None  # None pour tout traiter
RETRY_MAX = 3
RETRY_INITIAL_WAIT = 1.0

# ----------------------------
# HELPERS
# ----------------------------
def timestamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def save_debug(filename, text):
    path = os.path.join(DEBUG_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path

def normalize_guid(s):
    if not s:
        return ""
    return re.sub(r"[{}\-\s]", "", str(s).strip())

def looks_like_guid(s):
    if not s:
        return False
    s = str(s).strip()
    return bool(re.fullmatch(r"[A-Fa-f0-9]{32}", s) or re.fullmatch(r"[A-Fa-f0-9\-]{36}", s))

def retry_operation(fn, *args, max_retries=RETRY_MAX, **kwargs):
    wait = RETRY_INITIAL_WAIT
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_exc = e
            if attempt < max_retries:
                time.sleep(wait)
                wait *= 2
            else:
                raise last_exc

def make_client():
    """Crée le client ssaw avec fallback"""
    try:
        return ssaw.Client(SURVEY_URL, api_user=API_USER, api_password=API_PASSWORD, workspace=WORKSPACE), API_USER
    except:
        try:
            return ssaw.Client(SURVEY_URL, api_user=HQ_USERNAME, api_password=HQ_PASSWORD, workspace=WORKSPACE), HQ_USERNAME
        except Exception as e:
            raise RuntimeError(f"Impossible de créer client: {e}")

# ----------------------------
# ÉTAPE 1 : ENRICHISSEMENT - VERSION CORRIGÉE
# ----------------------------
def get_interview_details_by_key(interviews_api, interview_key):
    """
    MÉTHODE CORRIGÉE : Utilise get_list() avec filtre sur 'key'
    pour récupérer les métadonnées de l'entretien
    """
    try:
        print(f"      → Recherche via get_list avec key={interview_key}")
        
        # Utiliser get_list avec filtre sur la clé
        interviews = list(interviews_api.get_list(
            fields=['id', 'status', 'responsible_id', 'responsible_name', 'key'],
            key=interview_key,
            take=1
        ))
        
        if interviews and len(interviews) > 0:
            interview = interviews[0]
            print(f"      ✓ Entretien trouvé via get_list")
            return {
                'ResponsibleName': interview.responsible_name,
                'ResponsibleId': str(interview.responsible_id),
                'Status': interview.status
            }
        else:
            print(f"      ⚠ Aucun résultat avec get_list")
            return None
            
    except Exception as e:
        print(f"      ⚠ get_list échoué: {str(e)[:80]}")
        # Sauvegarder erreur complète
        save_debug(
            f"error_get_list_{interview_key.replace('-', '')}_{int(time.time())}.txt",
            f"Key: {interview_key}\nError: {traceback.format_exc()}"
        )
        return None

def get_interview_details_by_id(interviews_api, interview_id):
    """
    Récupère les détails via l'ID (GUID normalisé)
    """
    try:
        # Reformater avec tirets pour GraphQL
        if len(interview_id) == 32:
            interview_id_formatted = f"{interview_id[0:8]}-{interview_id[8:12]}-{interview_id[12:16]}-{interview_id[16:20]}-{interview_id[20:32]}"
        else:
            interview_id_formatted = interview_id
        
        print(f"      → Recherche via get_list avec id={interview_id_formatted[:12]}...")
        
        # Chercher par ID
        interviews = list(interviews_api.get_list(
            fields=['id', 'status', 'responsible_id', 'responsible_name', 'key'],
            id=interview_id_formatted,
            take=1
        ))
        
        if interviews and len(interviews) > 0:
            interview = interviews[0]
            print(f"      ✓ Entretien trouvé via get_list (ID)")
            return {
                'ResponsibleName': interview.responsible_name,
                'ResponsibleId': str(interview.responsible_id),
                'Status': interview.status
            }
        else:
            print(f"      ⚠ Aucun résultat avec l'ID")
            return None
            
    except Exception as e:
        print(f"      ⚠ Recherche par ID échouée: {str(e)[:80]}")
        save_debug(
            f"error_by_id_{interview_id[:8]}_{int(time.time())}.txt",
            f"ID: {interview_id}\nError: {traceback.format_exc()}"
        )
        return None

def enrich_excel():
    """
    ÉTAPE 1 : Enrichit le fichier Excel avec les informations des responsables
    """
    print("=" * 70)
    print("ÉTAPE 1 : ENRICHISSEMENT DU FICHIER EXCEL")
    print("=" * 70)
    
    # 1. Charger Excel
    print(f"\n📂 Chargement de {EXCEL_INPUT}...")
    try:
        df = pd.read_excel(EXCEL_INPUT, sheet_name=SHEET_INPUT, dtype=str)
        print(f"   ✅ {len(df)} lignes chargées")
    except Exception as e:
        print(f"   ❌ Erreur: {e}")
        return None
    
    # Vérifier colonnes requises
    if COL_KEY not in df.columns and COL_ID not in df.columns:
        print(f"   ❌ Aucune colonne {COL_KEY} ou {COL_ID} trouvée!")
        return None
    
    # 2. Initialiser colonnes
    df[COL_RESP_NAME] = ""
    df[COL_RESP_ID] = ""
    df[COL_STATUS] = ""
    
    # 3. Client
    print("\n🔌 Connexion à Survey Solutions...")
    try:
        client, user = make_client()
        print(f"   ✅ Connecté avec: {user}")
    except Exception as e:
        print(f"   ❌ {e}")
        return None
    
    interviews_api = InterviewsApi(client)
    
    # 4. Log
    with open(LOG_ENRICHMENT, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow([
            COL_KEY, COL_ID, COL_RESP_NAME, COL_RESP_ID, COL_STATUS, 
            "Resultat", "Message", "Timestamp"
        ])
    
    # 5. Traitement
    print("\n🔄 Enrichissement des données...\n")
    
    success_count = 0
    error_count = 0
    
    total = len(df) if MAX_ROWS is None else min(MAX_ROWS, len(df))
    
    for idx, row in df.iterrows():
        if MAX_ROWS and idx >= MAX_ROWS:
            break
        
        key = row.get(COL_KEY, "").strip() if COL_KEY in df.columns else ""
        iid = row.get(COL_ID, "").strip() if COL_ID in df.columns else ""
        
        print(f"[{idx+1}/{total}] Key: {key[:20] if key else 'N/A'}...")
        
        details = None
        
        # MÉTHODE 1 : Recherche par KEY (prioritaire)
        if key:
            print(f"    🔍 Méthode 1: Recherche par KEY...")
            details = retry_operation(get_interview_details_by_key, interviews_api, key, max_retries=2)
        
        # MÉTHODE 2 : Recherche par ID (fallback)
        if not details and iid and looks_like_guid(iid):
            print(f"    🔍 Méthode 2: Recherche par ID...")
            interview_id_normalized = normalize_guid(iid)
            details = retry_operation(get_interview_details_by_id, interviews_api, interview_id_normalized, max_retries=2)
        
        # Résultat
        if details:
            resp_name = details['ResponsibleName']
            resp_id = details['ResponsibleId']
            status = details['Status']
            
            # Mettre à jour DataFrame
            df.at[idx, COL_RESP_NAME] = resp_name
            df.at[idx, COL_RESP_ID] = resp_id
            df.at[idx, COL_STATUS] = status
            
            print(f"    ✅ Responsable: {resp_name}")
            print(f"       ID: {resp_id[:12] if resp_id else 'N/A'}...")
            print(f"       Statut: {status}\n")
            
            with open(LOG_ENRICHMENT, "a", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow([
                    key, iid, resp_name, resp_id, status,
                    "OK", "Succès", timestamp()
                ])
            success_count += 1
        else:
            print(f"    ❌ Impossible de récupérer les détails\n")
            with open(LOG_ENRICHMENT, "a", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow([
                    key, iid, "", "", "", "ERROR", "Détails introuvables", timestamp()
                ])
            error_count += 1
        
        time.sleep(0.3)
    
    # 6. Sauvegarder Excel enrichi
    print(f"\n💾 Sauvegarde du fichier enrichi...")
    try:
        df.to_excel(EXCEL_OUTPUT, sheet_name=SHEET_INPUT, index=False)
        print(f"   ✅ Fichier sauvegardé: {EXCEL_OUTPUT}")
    except Exception as e:
        print(f"   ❌ Erreur sauvegarde: {e}")
        return None
    
    # 7. Résumé
    print("\n" + "=" * 70)
    print("RÉSUMÉ ENRICHISSEMENT")
    print("=" * 70)
    print(f"Total traité    : {success_count + error_count}")
    print(f"✅ Succès       : {success_count}")
    print(f"❌ Erreurs      : {error_count}")
    print(f"📄 Log          : {LOG_ENRICHMENT}")
    print(f"📁 Fichier Excel: {EXCEL_OUTPUT}")
    print("=" * 70)
    
    return df

# ----------------------------
# ÉTAPE 2 : REJET INTELLIGENT (INCHANGÉE)
# ----------------------------
def smart_reject(interviews_api, interview_id, responsible_name, responsible_id, current_status):
    """
    Rejette intelligemment :
    - Si déjà rejeté → Skip
    - Sinon → Rejeter vers le responsable original
    """
    steps = []
    
    # Vérifier si déjà rejeté
    if current_status in ['RejectedByHeadquarters', 'RejectedBySupervisor']:
        msg = f"Déjà rejeté (statut: {current_status})"
        print(f"    ℹ️  {msg}")
        steps.append(("check_status", "SKIP", msg))
        return True, steps, msg
    
    print(f"    🔄 Rejet vers responsable original: {responsible_name}")
    
    # Normaliser l'ID du responsable
    resp_id_normalized = normalize_guid(responsible_id)
    
    # Tentative 1 : hqreject
    try:
        retry_operation(
            interviews_api.hqreject,
            interview_id,
            resp_id_normalized,
            "Rejet automatique vers responsable original"
        )
        msg = f"Rejeté vers {responsible_name} (hqreject)"
        steps.append(("hqreject", "OK", msg))
        print(f"    ✅ {msg}")
        return True, steps, msg
    except Exception as e:
        print(f"    ⚠️  hqreject échoué: {str(e)[:60]}")
        steps.append(("hqreject", "ERROR", str(e)[:100]))
    
    # Tentative 2 : reject
    try:
        retry_operation(
            interviews_api.reject,
            interview_id,
            resp_id_normalized,
            "Rejet automatique vers responsable original"
        )
        msg = f"Rejeté vers {responsible_name} (reject)"
        steps.append(("reject", "OK", msg))
        print(f"    ✅ {msg}")
        return True, steps, msg
    except Exception as e:
        print(f"    ❌ reject échoué: {str(e)[:60]}")
        steps.append(("reject", "ERROR", str(e)[:100]))
        msg = f"Échec rejet: {str(e)[:80]}"
        return False, steps, msg

def reject_interviews():
    """
    ÉTAPE 2 : Rejette les entretiens vers leur responsable original
    """
    print("\n\n" + "=" * 70)
    print("ÉTAPE 2 : REJET VERS RESPONSABLES ORIGINAUX")
    print("=" * 70)
    
    # 1. Charger fichier enrichi
    print(f"\n📂 Chargement de {EXCEL_OUTPUT}...")
    try:
        df = pd.read_excel(EXCEL_OUTPUT, sheet_name=SHEET_INPUT, dtype=str)
        print(f"   ✅ {len(df)} lignes chargées")
    except Exception as e:
        print(f"   ❌ Erreur: {e}")
        return
    
    # Vérifier colonnes
    required = [COL_KEY, COL_ID, COL_RESP_NAME, COL_RESP_ID, COL_STATUS]
    missing = [col for col in required if col not in df.columns]
    if missing:
        print(f"   ❌ Colonnes manquantes: {missing}")
        print(f"   ℹ️  Exécutez d'abord l'étape 1 (enrichissement)")
        return
    
    # Filtrer lignes valides
    df_valid = df[
        (df[COL_RESP_NAME].notna()) & 
        (df[COL_RESP_NAME] != "") &
        (df[COL_RESP_ID].notna()) &
        (df[COL_RESP_ID] != "")
    ].copy()
    
    print(f"   ✅ {len(df_valid)} lignes valides à traiter")
    
    if len(df_valid) == 0:
        print("   ℹ️  Aucune ligne à traiter")
        return
    
    # 2. Client
    print("\n🔌 Connexion à Survey Solutions...")
    try:
        client, user = make_client()
        print(f"   ✅ Connecté avec: {user}")
    except Exception as e:
        print(f"   ❌ {e}")
        return
    
    interviews_api = InterviewsApi(client)
    
    # 3. Log
    with open(LOG_REJECT, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow([
            COL_KEY, COL_ID, COL_RESP_NAME, COL_STATUS,
            "Steps", "Resultat", "Message", "Timestamp"
        ])
    
    # 4. Traitement
    print("\n🔄 Rejet des entretiens...\n")
    
    success_count = 0
    skip_count = 0
    error_count = 0
    
    total = len(df_valid) if MAX_ROWS is None else min(MAX_ROWS, len(df_valid))
    
    for idx, row in df_valid.iterrows():
        if MAX_ROWS and idx >= MAX_ROWS:
            break
        
        key = row[COL_KEY]
        iid = row[COL_ID]
        resp_name = row[COL_RESP_NAME]
        resp_id = row[COL_RESP_ID]
        status = row[COL_STATUS]
        
        print(f"[{idx+1}/{total}] {key[:20]}... → {resp_name}")
        print(f"    Status actuel: {status}")
        
        # Normaliser ID
        interview_id = normalize_guid(iid) if looks_like_guid(iid) else iid
        
        # Rejeter
        success, steps, message = smart_reject(
            interviews_api, interview_id, resp_name, resp_id, status
        )
        
        # Détecter skip
        is_skip = any(step[0] == "check_status" and step[1] == "SKIP" for step in steps)
        
        # Compter
        if success:
            if is_skip:
                skip_count += 1
            else:
                success_count += 1
        else:
            error_count += 1
        
        # Log
        steps_str = " → ".join([f"{s[0]}:{s[1]}" for s in steps])
        result = "SKIP" if is_skip else ("OK" if success else "ERROR")
        
        with open(LOG_REJECT, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                key, interview_id, resp_name, status,
                steps_str, result, message, timestamp()
            ])
        
        print(f"    {'✅' if success else '❌'} {message}\n")
        
        time.sleep(0.5)
    
    # 5. Résumé
    print("=" * 70)
    print("RÉSUMÉ REJET")
    print("=" * 70)
    print(f"Total traité    : {success_count + skip_count + error_count}")
    print(f"✅ Rejetés      : {success_count}")
    print(f"⏭️  Déjà rejetés : {skip_count}")
    print(f"❌ Erreurs      : {error_count}")
    print(f"📄 Log          : {LOG_REJECT}")
    print("=" * 70)

# ----------------------------
# MAIN
# ----------------------------
def main():
    print("""
╔══════════════════════════════════════════════════════════════════════╗
║  REJET INTELLIGENT VERS RESPONSABLES ORIGINAUX                      ║
║  Étape 1 : Enrichissement du fichier Excel                          ║
║  Étape 2 : Rejet des entretiens non rejetés                         ║
╚══════════════════════════════════════════════════════════════════════╝
    """)
    
    # Demander quelle étape exécuter
    print("Quelle étape souhaitez-vous exécuter ?")
    print("  1 - Enrichissement seulement")
    print("  2 - Rejet seulement (nécessite fichier enrichi)")
    print("  3 - Les deux (enrichissement puis rejet)")
    
    choice = input("\nVotre choix (1/2/3) [3]: ").strip() or "3"
    
    if choice == "1":
        enrich_excel()
    elif choice == "2":
        reject_interviews()
    elif choice == "3":
        df_enrichi = enrich_excel()
        if df_enrichi is not None:
            input("\n⏸️  Appuyez sur Entrée pour continuer vers l'étape 2...")
            reject_interviews()
    else:
        print("❌ Choix invalide")

if __name__ == "__main__":
    main()