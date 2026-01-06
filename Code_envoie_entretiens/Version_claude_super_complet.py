# -*- coding: utf-8 -*-
"""
Created on Tue Dec  9 08:33:06 2025

@author: mg.kouame
"""

# -*- coding: utf-8 -*-
"""
Réaffectation INTELLIGENTE d'entretiens SurveySolutions
- Vérifie le statut actuel de l'entretien
- Saute si déjà affecté au bon responsable
- Gère les cas où le rejet n'est pas possible (affectation directe)
Auteur: adapté pour mg.kouame
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

#Affichage du repertoire de travail par defaut de 
#Python
print(os.getcwd())

#Definition du nouvel envirionnement de travail
#pwd=r"C:\Users\mg.kouame\Documents\Point_de_collecte_ENEM\Sous_Echantillons_Itineraires\Trimestre4_2025\Fichier teleoperateur\Fusion"
pwd=r"D:\ENEM_Working\Activite_quotidienne_Trimestre_T4_2025\Resultat\Point_suivi_collecte"


os.chdir(pwd)


# ----------------------------
# CONFIG
# ----------------------------
SURVEY_URL = "http://154.68.47.214:9700"
WORKSPACE = "enemenage"

API_USER = "UserAPIEEC"
API_PASSWORD = "UserAPIEEC1"
HQ_USERNAME = "SupDaouda"
HQ_PASSWORD = "SupDaouda01"

#pwd = r"C:\Users\mg.kouame\Documents\Point_de_collecte_ENEM\Sous_Echantillons_Itineraires\Trimestre4_2025\Fichier teleoperateur\Fusion"
os.chdir(pwd)

EXCEL_PATH = r"Base_rejet_Agent_enqueteur_V2_ENRICHI.xlsx"
SHEET_NAME = "Fichier_rejet_py"
COL_KEY = "INTERVIEW_KEY"
COL_RESP = "Numero_agent_terrain"
COL_ID = "INTERVIEW_ID"

LOG_CSV = "reassign_api_log.csv"
DEBUG_DIR = "debug_api_errors"
os.makedirs(DEBUG_DIR, exist_ok=True)

MAX_ROWS = None   # test, None pour tout
RETRY_MAX = 3
RETRY_INITIAL_WAIT = 1.0

USER_UUID_CACHE = {}

# ----------------------------
# Helpers
# ----------------------------
def timestamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def save_debug_text(filename, text):
    path = os.path.join(DEBUG_DIR, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path

def looks_like_guid(s):
    if not s:
        return False
    s = str(s).strip()
    return bool(re.fullmatch(r"[A-Fa-f0-9]{32}", s) or re.fullmatch(r"[A-Fa-f0-9\-]{36}", s))

def normalize_guid(s):
    if not s:
        return ""
    return re.sub(r"[{}\-\s]", "", str(s).strip())

def retry_operation(fn, *args, max_retries=RETRY_MAX, initial_wait=RETRY_INITIAL_WAIT, **kwargs):
    wait = initial_wait
    last_exc = None
    for attempt in range(1, max_retries + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_exc = e
            if attempt < max_retries:
                time.sleep(wait)
                wait *= 2
                continue
            else:
                raise last_exc

# ----------------------------
# Client creation
def make_client_try(api_user=API_USER, api_password=API_PASSWORD, fallback_hq=True):
    try:
        client = ssaw.Client(SURVEY_URL, api_user=api_user, api_password=api_password, workspace=WORKSPACE)
        return client, api_user
    except Exception as e:
        if fallback_hq:
            try:
                client = ssaw.Client(SURVEY_URL, api_user=HQ_USERNAME, api_password=HQ_PASSWORD, workspace=WORKSPACE)
                return client, HQ_USERNAME
            except Exception as e2:
                raise RuntimeError(f"Impossible de créer client: {e}, {e2}")
        raise RuntimeError(f"Impossible de créer client: {e}")

# ----------------------------
# Recherche UUID
def get_user_uuid_by_username(client, username):
    if username in USER_UUID_CACHE:
        return USER_UUID_CACHE[username]
    
    try:
        users_api = UsersApi(client)
        
        # Chercher dans superviseurs
        supervisors = list(users_api.list_supervisors())
        for sup in supervisors:
            if sup.get('UserName', '').lower() == username.lower():
                user_uuid = sup.get('UserId')
                USER_UUID_CACHE[username] = user_uuid
                return user_uuid
        
        # Chercher dans interviewers
        for sup in supervisors:
            sup_id = sup.get('UserId')
            if sup_id:
                try:
                    interviewers = list(users_api.list_interviewers(sup_id))
                    for interviewer in interviewers:
                        if interviewer.get('UserName', '').lower() == username.lower():
                            user_uuid = interviewer.get('UserId')
                            USER_UUID_CACHE[username] = user_uuid
                            return user_uuid
                except Exception:
                    continue
        
        return None
        
    except Exception as e:
        save_debug_text(
            f"user_search_error_{username}_{int(time.time())}.txt",
            f"Erreur recherche {username}\n{traceback.format_exc()}"
        )
        return None

# ----------------------------
# Obtenir infos de l'entretien
def get_interview_info(interviews_api, interview_id):
    """
    Récupère les informations de l'entretien :
    - statut actuel
    - responsable actuel (ID et nom)
    """
    try:
        # Utiliser get_list avec filtre sur l'ID
        interviews = list(interviews_api.get_list(
            fields=['id', 'status', 'responsible_id', 'responsible_name'],
            key=interview_id,
            take=1
        ))
        
        if interviews:
            interview = interviews[0]
            return {
                'status': interview.status,
                'responsible_id': interview.responsible_id,
                'responsible_name': interview.responsible_name
            }
        return None
        
    except Exception as e:
        print(f"    ⚠ Erreur get_interview_info: {e}")
        # Fallback: essayer via l'API REST directe
        try:
            info = interviews_api.get_info(interview_id)
            if info:
                # Parser la réponse
                return {
                    'status': info.get('Status'),
                    'responsible_id': info.get('ResponsibleId'),
                    'responsible_name': info.get('ResponsibleName', 'Inconnu')
                }
        except Exception as e2:
            print(f"    ⚠ Erreur fallback get_info: {e2}")
        
        return None

# ----------------------------
# Résolution interview ID par key
def get_interview_id_by_key(interviews_api, interview_key):
    try:
        interview_obj = interviews_api._get_interview_by_key(fields=["id"], key=interview_key)
        if interview_obj is None:
            return None
        for attr in ("id", "Id", "InterviewId"):
            if hasattr(interview_obj, attr):
                return getattr(interview_obj, attr)
        if isinstance(interview_obj, dict):
            return interview_obj.get("id") or interview_obj.get("InterviewId")
    except Exception:
        pass
    return None

# ----------------------------
# RÉAFFECTATION INTELLIGENTE
def smart_reassign_interview(interviews_api, client, interview_id, responsible_username):
    """
    Réaffectation intelligente :
    1. Vérifie l'état actuel de l'entretien
    2. Saute si déjà affecté au bon responsable
    3. Essaie rejet puis affectation (ou affectation directe si rejet impossible)
    """
    steps_log = []
    
    # Étape 0: Vérifier l'état actuel de l'entretien
    print(f"  [0/4] Vérification de l'état actuel...")
    current_info = get_interview_info(interviews_api, interview_id)
    
    if current_info:
        current_status = current_info.get('status', 'Inconnu')
        current_responsible = current_info.get('responsible_name', 'Inconnu')
        current_resp_id = current_info.get('responsible_id', '')
        
        print(f"    → Statut actuel: {current_status}")
        print(f"    → Responsable actuel: {current_responsible}")
        
        steps_log.append(("check_status", "INFO", f"Statut={current_status}, Resp={current_responsible}"))
    else:
        print(f"    ⚠ Impossible de récupérer l'info de l'entretien")
        current_status = "Inconnu"
        current_responsible = "Inconnu"
        current_resp_id = ""
    
    # Étape 1: Obtenir UUID du responsable cible
    print(f"  [1/4] Recherche UUID de '{responsible_username}'...")
    responsible_uuid = retry_operation(
        get_user_uuid_by_username, 
        client, 
        responsible_username, 
        max_retries=2
    )
    
    if not responsible_uuid:
        msg = f"UUID introuvable pour '{responsible_username}'"
        steps_log.append(("get_uuid", "ERROR", msg))
        return False, steps_log, msg
    
    print(f"    ✓ UUID trouvé: {responsible_uuid[:8]}...")
    steps_log.append(("get_uuid", "OK", f"UUID: {responsible_uuid[:8]}..."))
    
    # VÉRIFICATION: Est-ce déjà le bon responsable ?
    if current_resp_id and normalize_guid(str(current_resp_id)) == normalize_guid(str(responsible_uuid)):
        print(f"    ℹ Entretien déjà affecté au bon responsable ({responsible_username})")
        
        # Vérifier si déjà rejeté
        if current_status in ['RejectedByHeadquarters', 'RejectedBySupervisor']:
            msg = f"Déjà rejeté vers {responsible_username} - rien à faire"
            print(f"    ✓ {msg}")
            steps_log.append(("skip", "OK", msg))
            return True, steps_log, msg
        else:
            msg = f"Déjà affecté à {responsible_username} (statut: {current_status})"
            print(f"    ✓ {msg}")
            steps_log.append(("skip", "OK", msg))
            return True, steps_log, msg
    
    # Étape 2: Tentative de rejet
    print(f"  [2/4] Tentative de rejet...")
    reject_success = False
    reject_not_allowed = False
    
    try:
        retry_operation(
            interviews_api.hqreject, 
            interview_id, 
            "Réaffectation automatique", 
            max_retries=2
        )
        steps_log.append(("hqreject", "OK", "Rejeté"))
        reject_success = True
        print("    ✓ hqreject OK")
        
    except Exception as e:
        error_msg = str(e).lower()
        
        # Détecter si le rejet n'est pas autorisé
        if any(keyword in error_msg for keyword in ['not allowed', 'invalid status', 'cannot be rejected', 'already rejected']):
            reject_not_allowed = True
            print(f"    ℹ Rejet non autorisé (statut: {current_status}) - tentative affectation directe")
            steps_log.append(("hqreject", "SKIPPED", f"Rejet non autorisé pour statut {current_status}"))
        else:
            print(f"    ⚠ hqreject échoué: {str(e)[:60]}...")
            
            # Fallback: reject
            try:
                retry_operation(
                    interviews_api.reject, 
                    interview_id, 
                    "Réaffectation automatique", 
                    max_retries=2
                )
                steps_log.append(("reject", "OK", "Rejeté"))
                reject_success = True
                print("    ✓ reject OK")
                
            except Exception as e2:
                error_msg2 = str(e2).lower()
                if any(keyword in error_msg2 for keyword in ['not allowed', 'invalid status', 'cannot be rejected']):
                    reject_not_allowed = True
                    print(f"    ℹ Rejet non autorisé - tentative affectation directe")
                    steps_log.append(("reject", "SKIPPED", "Rejet non autorisé"))
                else:
                    msg = f"Impossible de rejeter: {str(e2)[:80]}"
                    steps_log.append(("reject", "ERROR", msg))
                    print(f"    ✗ {msg}")
    
    # Pause si rejet réussi
    if reject_success:
        time.sleep(2)
    
    # Étape 3: Affectation
    print(f"  [3/4] Affectation à '{responsible_username}' (UUID: {responsible_uuid[:8]}...)...")
    
    try:
        retry_operation(
            interviews_api.assign,
            interview_id,
            responsible_uuid,
            responsiblename=responsible_username,
            max_retries=2
        )
        
        if reject_not_allowed:
            msg = f"Affectation directe réussie (rejet non autorisé)"
            steps_log.append(("assign_direct", "OK", msg))
        else:
            msg = f"Affecté à {responsible_username}"
            steps_log.append(("assign", "OK", msg))
        
        print(f"    ✓ Affectation réussie")
        return True, steps_log, msg
        
    except Exception as e:
        msg = f"Erreur affectation: {str(e)[:100]}"
        steps_log.append(("assign", "ERROR", msg))
        print(f"    ✗ {msg}")
        return False, steps_log, msg

# ----------------------------
# MAIN
def main():
    print("="*60)
    print("RÉAFFECTATION INTELLIGENTE (avec vérification statut)")
    print("="*60)
    
    # Lecture Excel
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, dtype=str)
    except Exception as e:
        print(f"Erreur lecture Excel: {e}")
        return

    if COL_RESP not in df.columns:
        print(f"Colonne requise manquante: {COL_RESP}")
        return

    df[COL_KEY] = df.get(COL_KEY, "").fillna("").astype(str).str.strip()
    df[COL_ID] = df.get(COL_ID, "").fillna("").astype(str).str.strip()
    df[COL_RESP] = df[COL_RESP].fillna("").astype(str).str.strip()

    df_valid = df[(df[COL_ID] != "") | (df[COL_KEY] != "")]
    df_valid = df_valid[df_valid[COL_RESP] != ""]
    
    print(f"Lignes valides: {len(df_valid)}")
    if len(df_valid) == 0:
        print("Aucune ligne à traiter.")
        return

    if MAX_ROWS:
        df_valid = df_valid.head(MAX_ROWS)
        print(f"Mode TEST: {MAX_ROWS} premières lignes")

    # Client
    try:
        client, used_cred = make_client_try()
        print(f"Client créé: {used_cred}\n")
    except Exception as e:
        print(f"Impossible de créer client: {e}")
        return

    interviews_api = InterviewsApi(client)

    # Log CSV
    with open(LOG_CSV, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow([
            "INTERVIEW_KEY", "INTERVIEW_ID", "Numero_agent_terrain",
            "steps", "status", "message", "timestamp", "used_cred"
        ])

    # Boucle
    count = 0
    success_count = 0
    error_count = 0
    skipped_count = 0
    
    for idx, row in df_valid.iterrows():
        count += 1
        ik = row.get(COL_KEY, "") or ""
        iid_raw = row.get(COL_ID, "") or ""
        resp = row.get(COL_RESP, "").strip()
        
        print(f"\n{'='*60}")
        print(f"[{count}/{len(df_valid)}] {ik[:20]}... -> {resp}")
        print('='*60)

        interview_id = None

        # Utiliser ID si présent
        if iid_raw and looks_like_guid(iid_raw):
            interview_id = normalize_guid(iid_raw)
            print(f"  ✓ ID fourni: {interview_id[:8]}...")
        
        # Sinon résoudre par key
        elif ik:
            print("  → Résolution par KEY...")
            try:
                interview_id = retry_operation(
                    get_interview_id_by_key, 
                    interviews_api, 
                    ik, 
                    max_retries=2
                )
                if interview_id:
                    interview_id = normalize_guid(interview_id)
                    print(f"  ✓ ID trouvé: {interview_id[:8]}...")
                else:
                    print("  ✗ ID introuvable")
            except Exception as e:
                print(f"  ✗ Erreur: {e}")

        # Si pas d'ID
        if not interview_id:
            msg = "INTERVIEW_ID introuvable"
            print(f"  ✗ {msg}\n")
            with open(LOG_CSV, "a", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow([
                    ik, iid_raw, resp, "", "ERROR", msg, timestamp(), used_cred
                ])
            error_count += 1
            continue

        # RÉAFFECTATION INTELLIGENTE
        success, steps_log, message = smart_reassign_interview(
            interviews_api, client, interview_id, resp
        )
        
        # Log
        status = "OK" if success else "ERROR"
        steps_str = " -> ".join([f"{s[0]}:{s[1]}" for s in steps_log])
        
        # Détecter si c'était un skip
        is_skip = any(step[0] == "skip" for step in steps_log)
        
        with open(LOG_CSV, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                ik, interview_id, resp, steps_str, status, message, timestamp(), used_cred
            ])

        if success:
            if is_skip:
                skipped_count += 1
            else:
                success_count += 1
        else:
            error_count += 1
            dbg_name = f"error_{count}_{ik[:10] if ik else 'noKey'}_{int(time.time())}.txt"
            dbg_text = f"Key: {ik}\nID: {interview_id}\nResp: {resp}\nSteps:\n"
            for step in steps_log:
                dbg_text += f"  - {step[0]}: {step[1]} - {step[2]}\n"
            save_debug_text(dbg_name, dbg_text)

        print(f"\n  RÉSULTAT: {status} - {message}")
        time.sleep(0.5)

    # Résumé
    print("\n" + "="*60)
    print("RÉSUMÉ")
    print("="*60)
    print(f"Total traité       : {count}")
    print(f"Réaffectations OK  : {success_count}")
    print(f"Déjà affectés      : {skipped_count}")
    print(f"Erreurs            : {error_count}")
    print(f"Log détaillé       : {LOG_CSV}")
    print(f"Dossier debug      : {DEBUG_DIR}")
    print("="*60)

if __name__ == "__main__":
    main()