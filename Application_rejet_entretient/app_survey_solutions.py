# -*- coding: utf-8 -*-
"""
================================================================================
APPLICATION : Gestion des entretiens Survey Solutions
================================================================================
  Projet 1 : Enrichissement du fichier Excel (récupération responsable actuel)
  Projet 2 : Réaffectation intelligente des entretiens
================================================================================
AUTEUR  : mg.kouame
VERSION : 1.0 - Interface Tkinter
================================================================================
"""

import os
import re
import sys
import time
import csv
import socket
import threading
import traceback
import subprocess
import importlib
import importlib.util
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# pandas sera importé après auto-install si nécessaire
try:
    import pandas as pd
except ImportError:
    pd = None

# ==============================================================================
# AUTO-INSTALLATION DES DÉPENDANCES
# ==============================================================================

import sys
import subprocess
import importlib

# Liste des modules requis : (nom_import, nom_pip)
MODULES_REQUIS = [
    ("pandas",   "pandas"),
    ("openpyxl", "openpyxl"),
    ("ssaw",     "ssaw"),
]

def verifier_connexion_internet():
    """Vérifie si une connexion internet est disponible."""
    import socket
    try:
        socket.setdefaulttimeout(5)
        socket.create_connection(("8.8.8.8", 53))
        return True
    except OSError:
        return False

def installer_module(nom_pip, fenetre_log=None):
    """
    Installe un module via pip.
    Retourne (succès: bool, message: str).
    """
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", nom_pip,
             "--quiet", "--no-warn-script-location"],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0:
            return True, f"✅ '{nom_pip}' installé avec succès."
        else:
            return False, f"❌ Erreur pip pour '{nom_pip}' :\n{result.stderr[:300]}"
    except subprocess.TimeoutExpired:
        return False, f"❌ Timeout lors de l'installation de '{nom_pip}'."
    except Exception as e:
        return False, f"❌ Erreur inattendue : {e}"

def verifier_et_installer_modules():
    """
    Vérifie tous les modules requis.
    Installe ceux qui manquent.
    Retourne la liste des modules toujours manquants après tentative.
    """
    # Déterminer lesquels manquent
    manquants = []
    for nom_import, nom_pip in MODULES_REQUIS:
        spec = importlib.util.find_spec(nom_import)
        if spec is None:
            manquants.append((nom_import, nom_pip))

    if not manquants:
        return []  # Tout est déjà installé

    # Il manque des modules → fenêtre de progression
    root_splash = tk.Tk()
    root_splash.title("Installation des dépendances")
    root_splash.geometry("520x340")
    root_splash.resizable(False, False)
    root_splash.configure(bg="#0D1117")
    root_splash.eval('tk::PlaceWindow . center')

    tk.Label(root_splash,
             text="🔧  Initialisation de l'application",
             font=("Segoe UI", 13, "bold"), fg="#79C0FF",
             bg="#0D1117").pack(pady=(20, 4))
    tk.Label(root_splash,
             text="Des modules Python nécessaires sont manquants.\n"
                  "Installation automatique en cours…",
             font=("Segoe UI", 9), fg="#C9D1D9",
             bg="#0D1117", justify="center").pack(pady=(0, 12))

    zone = scrolledtext.ScrolledText(
        root_splash, font=("Consolas", 9),
        bg="#161B22", fg="#C9D1D9",
        height=8, relief="flat", state="normal"
    )
    zone.pack(fill="x", padx=16)
    zone.tag_config("ok",   foreground="#56D364")
    zone.tag_config("err",  foreground="#F85149")
    zone.tag_config("warn", foreground="#E3B341")
    zone.tag_config("info", foreground="#C9D1D9")

    barre = ttk.Progressbar(root_splash, mode="determinate",
                            maximum=len(manquants), value=0)
    barre.pack(fill="x", padx=16, pady=8)

    lbl_statut = tk.Label(root_splash, text="",
                          font=("Segoe UI", 8), fg="#8B949E",
                          bg="#0D1117")
    lbl_statut.pack()

    encore_manquants = []

    def log_splash(msg, tag="info"):
        zone.insert("end", msg + "\n", tag)
        zone.see("end")
        root_splash.update()

    def lancer_installations():
        nonlocal encore_manquants

        # Vérifier internet
        log_splash("🌐 Vérification de la connexion internet…")
        root_splash.update()

        if not verifier_connexion_internet():
            log_splash(
                "❌ Aucune connexion internet détectée !\n"
                "   Veuillez vous connecter à internet pour que\n"
                "   l'installation des modules puisse se faire,\n"
                "   puis relancez l'application.",
                "err"
            )
            lbl_statut.config(
                text="⚠️  Connexion internet requise — relancez après connexion.",
                fg="#E3B341"
            )

            tk.Button(
                root_splash,
                text="Fermer l'application",
                font=("Segoe UI", 9, "bold"),
                bg="#8B0000", fg="white", relief="flat", cursor="hand2",
                command=lambda: (root_splash.destroy(), sys.exit(1))
            ).pack(pady=10)
            return

        log_splash("✅ Connexion internet disponible.", "ok")
        log_splash(f"📦 {len(manquants)} module(s) à installer : "
                   f"{', '.join(p for _, p in manquants)}\n")

        for i, (nom_import, nom_pip) in enumerate(manquants, 1):
            lbl_statut.config(text=f"Installation de '{nom_pip}'… ({i}/{len(manquants)})")
            log_splash(f"⬇️  Installation de '{nom_pip}'…")
            root_splash.update()

            ok, msg = installer_module(nom_pip)
            tag = "ok" if ok else "err"
            log_splash(msg, tag)

            if ok:
                # Recharger le module dans sys.modules
                try:
                    importlib.invalidate_caches()
                    importlib.import_module(nom_import)
                except Exception:
                    encore_manquants.append((nom_import, nom_pip))
            else:
                encore_manquants.append((nom_import, nom_pip))

            barre['value'] = i
            root_splash.update()

        if encore_manquants:
            noms = ", ".join(p for _, p in encore_manquants)
            log_splash(
                f"\n⚠️  Modules toujours manquants : {noms}\n"
                "   Certaines fonctionnalités peuvent ne pas fonctionner.",
                "warn"
            )
            lbl_statut.config(
                text="⚠️  Installation partielle — voir les détails ci-dessus.",
                fg="#E3B341"
            )
            tk.Button(
                root_splash,
                text="Continuer quand même →",
                font=("Segoe UI", 9, "bold"),
                bg="#C0782A", fg="white", relief="flat", cursor="hand2",
                command=root_splash.destroy
            ).pack(pady=10)
        else:
            log_splash("\n✅ Tous les modules sont installés. Lancement…", "ok")
            lbl_statut.config(text="✅ Prêt !", fg="#56D364")
            root_splash.update()
            root_splash.after(1200, root_splash.destroy)

    # Lancer dans un thread pour ne pas bloquer l'interface splash
    threading.Thread(target=lancer_installations, daemon=True).start()
    root_splash.mainloop()

    return encore_manquants

# ==============================================================================
# CONSTANTES GLOBALES
# ==============================================================================
SSAW_OK = False  # sera mis à True après auto-install si ssaw est disponible
COULEUR_FOND      = "#F0F4F8"
COULEUR_TITRE     = "#1A3A5C"
COULEUR_SECTION   = "#2E6DA4"
COULEUR_BTN       = "#2E6DA4"
COULEUR_BTN_OK    = "#1E8C45"
COULEUR_BTN_QUIT  = "#8B0000"
COULEUR_BTN_WARN  = "#C0782A"
POLICE_TITRE      = ("Segoe UI", 13, "bold")
POLICE_SECTION    = ("Segoe UI", 10, "bold")
POLICE_NORMALE    = ("Segoe UI", 9)
POLICE_MONO       = ("Consolas", 9)

DEFAULT_SURVEY_URL   = "http://154.68.47.214:9700"
DEFAULT_WORKSPACE    = "enemenage"
DEFAULT_API_USER     = "UserAPIEEC"
DEFAULT_API_PASSWORD = "UserAPIEEC1"
DEFAULT_HQ_USERNAME  = "SupDaouda"
DEFAULT_HQ_PASSWORD  = "SupDaouda01"

# ==============================================================================
# HELPERS COMMUNS
# ==============================================================================

def timestamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def normalize_guid(s):
    if not s:
        return ""
    return re.sub(r"[{}\-\s]", "", str(s).strip())

def looks_like_guid(s):
    if not s:
        return False
    s = str(s).strip()
    return bool(
        re.fullmatch(r"[A-Fa-f0-9]{32}", s) or
        re.fullmatch(r"[A-Fa-f0-9\-]{36}", s)
    )

def retry_operation(fn, *args, max_retries=3, initial_wait=1.0, **kwargs):
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
            else:
                raise last_exc

def make_client(survey_url, workspace, api_user, api_password, hq_username, hq_password):
    if not SSAW_OK:
        raise RuntimeError("Le module 'ssaw' n'est pas installé. Lancez : pip install ssaw")
    try:
        client = ssaw.Client(survey_url, api_user=api_user,
                             api_password=api_password, workspace=workspace)
        return client, api_user
    except Exception:
        try:
            client = ssaw.Client(survey_url, api_user=hq_username,
                                 api_password=hq_password, workspace=workspace)
            return client, hq_username
        except Exception as e:
            raise RuntimeError(f"Impossible de créer le client Survey Solutions : {e}")

def get_sheets(filepath):
    """Retourne la liste des feuilles d'un fichier Excel."""
    try:
        return pd.ExcelFile(filepath).sheet_names
    except Exception as e:
        raise RuntimeError(f"Impossible de lire le fichier : {e}")

def get_columns(filepath, sheet):
    """Retourne la liste des colonnes d'une feuille Excel."""
    try:
        df = pd.read_excel(filepath, sheet_name=sheet, nrows=0)
        return list(df.columns)
    except Exception as e:
        raise RuntimeError(f"Impossible de lire les colonnes : {e}")

# ==============================================================================
# WIDGETS RÉUTILISABLES
# ==============================================================================

def make_scrolled_frame(parent):
    """Crée un frame scrollable vertical."""
    canvas = tk.Canvas(parent, bg=COULEUR_FOND, highlightthickness=0)
    scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    frame = tk.Frame(canvas, bg=COULEUR_FOND)
    frame.bind("<Configure>",
               lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    return frame

def make_section(parent, titre):
    f = tk.LabelFrame(
        parent, text=f"  {titre}  ",
        font=POLICE_SECTION, fg=COULEUR_SECTION,
        bg=COULEUR_FOND, padx=12, pady=8, relief="groove"
    )
    f.pack(fill="x", padx=14, pady=6)
    return f

def make_file_row(parent, label, var, types, aide, row):
    """Ligne fichier avec bouton Parcourir."""
    tk.Label(parent, text=label, font=POLICE_SECTION,
             bg=COULEUR_FOND, fg=COULEUR_TITRE, anchor="w"
             ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(6, 1))
    if aide:
        tk.Label(parent, text=aide, font=("Segoe UI", 8), fg="#555",
                 bg=COULEUR_FOND, wraplength=520, justify="left"
                 ).grid(row=row+1, column=0, columnspan=3, sticky="w")
    e = tk.Entry(parent, textvariable=var, font=POLICE_MONO, width=58)
    e.grid(row=row+2, column=0, columnspan=2, sticky="ew", pady=2)
    tk.Button(
        parent, text="📁 Parcourir", font=POLICE_NORMALE,
        bg="#E8EFF8", relief="flat", cursor="hand2",
        command=lambda v=var, t=types: _browse_file(v, t)
    ).grid(row=row+2, column=2, sticky="w", padx=4)
    parent.columnconfigure(0, weight=1)
    return row + 3

def make_dir_row(parent, label, var, aide, row):
    """Ligne dossier avec bouton Parcourir."""
    tk.Label(parent, text=label, font=POLICE_SECTION,
             bg=COULEUR_FOND, fg=COULEUR_TITRE, anchor="w"
             ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(6, 1))
    if aide:
        tk.Label(parent, text=aide, font=("Segoe UI", 8), fg="#555",
                 bg=COULEUR_FOND, wraplength=520, justify="left"
                 ).grid(row=row+1, column=0, columnspan=3, sticky="w")
    e = tk.Entry(parent, textvariable=var, font=POLICE_MONO, width=58)
    e.grid(row=row+2, column=0, columnspan=2, sticky="ew", pady=2)
    tk.Button(
        parent, text="📁 Parcourir", font=POLICE_NORMALE,
        bg="#E8EFF8", relief="flat", cursor="hand2",
        command=lambda v=var: _browse_dir(v)
    ).grid(row=row+2, column=2, sticky="w", padx=4)
    parent.columnconfigure(0, weight=1)
    return row + 3

def _browse_file(var, types):
    f = filedialog.askopenfilename(filetypes=types)
    if f:
        var.set(f)

def _browse_dir(var):
    d = filedialog.askdirectory()
    if d:
        var.set(d)

def make_combo_row(parent, label, var, values, aide, row, width=36):
    tk.Label(parent, text=label, font=POLICE_NORMALE,
             bg=COULEUR_FOND, anchor="w", width=26
             ).grid(row=row, column=0, sticky="w", pady=2)
    combo = ttk.Combobox(parent, textvariable=var, values=values,
                         width=width, font=POLICE_NORMALE, state="readonly")
    combo.grid(row=row, column=1, sticky="ew", padx=6)
    if aide:
        tk.Label(parent, text=aide, font=("Segoe UI", 8), fg="#666",
                 bg=COULEUR_FOND, wraplength=300, justify="left"
                 ).grid(row=row+1, column=0, columnspan=3, sticky="w", padx=4)
    parent.columnconfigure(1, weight=1)
    return combo, row + (2 if aide else 1)

def make_entry_row(parent, label, var, aide, row, width=36):
    tk.Label(parent, text=label, font=POLICE_NORMALE,
             bg=COULEUR_FOND, anchor="w", width=26
             ).grid(row=row, column=0, sticky="w", pady=2)
    e = tk.Entry(parent, textvariable=var, font=POLICE_NORMALE, width=width)
    e.grid(row=row, column=1, sticky="ew", padx=6)
    if aide:
        tk.Label(parent, text=aide, font=("Segoe UI", 8), fg="#666",
                 bg=COULEUR_FOND, wraplength=340, justify="left"
                 ).grid(row=row+1, column=0, columnspan=3, sticky="w", padx=4)
    parent.columnconfigure(1, weight=1)
    return e, row + (2 if aide else 1)

def make_logs_zone(parent):
    """Zone de logs colorée style terminal."""
    z = scrolledtext.ScrolledText(
        parent, font=POLICE_MONO, bg="#0D1117", fg="#C9D1D9",
        insertbackground="white", relief="flat", wrap="none"
    )
    z.pack(fill="both", expand=True, padx=10, pady=(0, 8))
    z.tag_config("ok",    foreground="#56D364")
    z.tag_config("err",   foreground="#F85149")
    z.tag_config("warn",  foreground="#E3B341")
    z.tag_config("titre", foreground="#79C0FF", font=("Consolas", 9, "bold"))
    z.tag_config("info",  foreground="#C9D1D9")
    return z

def log_write(zone, msg, master):
    """Écrit une ligne dans la zone de logs (thread-safe)."""
    def _w():
        if "✅" in msg or "✓" in msg:
            tag = "ok"
        elif "❌" in msg or "ERREUR" in msg or "ERROR" in msg:
            tag = "err"
        elif "⚠" in msg or "WARN" in msg:
            tag = "warn"
        elif msg.strip().startswith("=") or "RÉSUMÉ" in msg or "ÉTAPE" in msg:
            tag = "titre"
        else:
            tag = "info"
        zone.insert("end", msg + "\n", tag)
        zone.see("end")
    master.after(0, _w)

# ==============================================================================
# LOGIQUE MÉTIER — PROJET 1 : ENRICHISSEMENT
# ==============================================================================

def _p1_get_by_key(interviews_api, key):
    try:
        items = list(interviews_api.get_list(
            fields=['id', 'status', 'responsible_id', 'responsible_name', 'key'],
            key=key, take=1
        ))
        if items:
            it = items[0]
            return {'ResponsibleName': it.responsible_name,
                    'ResponsibleId':   str(it.responsible_id),
                    'Status':          it.status}
    except Exception:
        pass
    return None

def _p1_get_by_id(interviews_api, iid):
    try:
        n = normalize_guid(iid)
        fmt = f"{n[0:8]}-{n[8:12]}-{n[12:16]}-{n[16:20]}-{n[20:32]}" if len(n) == 32 else n
        items = list(interviews_api.get_list(
            fields=['id', 'status', 'responsible_id', 'responsible_name', 'key'],
            id=fmt, take=1
        ))
        if items:
            it = items[0]
            return {'ResponsibleName': it.responsible_name,
                    'ResponsibleId':   str(it.responsible_id),
                    'Status':          it.status}
    except Exception:
        pass
    return None

def run_enrichissement(params, log, fin):
    """Thread cible pour le Projet 1."""
    try:
        fichier_input   = params['fichier_input']
        sheet_input     = params['sheet_input']
        dossier_output  = params['dossier_output']
        nom_output      = params['nom_output']
        col_key         = params['col_key']
        col_id          = params['col_id']
        col_resp_name   = params['col_resp_name']
        col_resp_id     = params['col_resp_id']
        col_status      = params['col_status']
        survey_url      = params['survey_url']
        workspace       = params['workspace']
        api_user        = params['api_user']
        api_password    = params['api_password']
        hq_username     = params['hq_username']
        hq_password     = params['hq_password']

        sep = "=" * 65
        log(sep)
        log("  PROJET 1 — ENRICHISSEMENT DU FICHIER EXCEL")
        log(sep)

        # Charger Excel
        log(f"\n📂 Chargement : {os.path.basename(fichier_input)}  [feuille: {sheet_input}]")
        df = pd.read_excel(fichier_input, sheet_name=sheet_input, dtype=str)
        df = df.fillna("")
        log(f"   ✓ {len(df)} lignes chargées")

        # Vérifier colonnes clé
        if col_key not in df.columns and col_id not in df.columns:
            raise ValueError(f"Aucune colonne '{col_key}' ou '{col_id}' trouvée dans la feuille.")

        # Initialiser colonnes résultat
        df[col_resp_name] = ""
        df[col_resp_id]   = ""
        df[col_status]    = ""

        # Connexion
        log("\n🔌 Connexion à Survey Solutions...")
        client, used = make_client(survey_url, workspace, api_user, api_password,
                                   hq_username, hq_password)
        log(f"   ✓ Connecté avec : {used}")
        interviews_api = InterviewsApi(client)

        # Fichier log CSV
        os.makedirs(dossier_output, exist_ok=True)
        log_csv = os.path.join(dossier_output, "log_enrichissement.csv")
        with open(log_csv, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([col_key, col_id, col_resp_name, col_resp_id,
                                    col_status, "Resultat", "Message", "Timestamp"])

        log("\n🔄 Enrichissement en cours...\n")
        nb_ok = nb_err = 0
        total = len(df)

        for idx, row in df.iterrows():
            key = str(row.get(col_key, "")).strip()
            iid = str(row.get(col_id,  "")).strip()
            log(f"[{idx+1}/{total}] KEY: {key[:22] if key else 'N/A'}")

            details = None
            if key:
                details = retry_operation(_p1_get_by_key, interviews_api, key, max_retries=2)
            if not details and iid and looks_like_guid(iid):
                details = retry_operation(_p1_get_by_id, interviews_api,
                                          normalize_guid(iid), max_retries=2)

            if details:
                df.at[idx, col_resp_name] = details['ResponsibleName']
                df.at[idx, col_resp_id]   = details['ResponsibleId']
                df.at[idx, col_status]    = details['Status']
                log(f"   ✅ Responsable : {details['ResponsibleName']} | Statut : {details['Status']}")
                with open(log_csv, "a", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow([key, iid, details['ResponsibleName'],
                                            details['ResponsibleId'], details['Status'],
                                            "OK", "Succès", timestamp()])
                nb_ok += 1
            else:
                log(f"   ❌ Impossible de récupérer les détails")
                with open(log_csv, "a", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow([key, iid, "", "", "",
                                            "ERROR", "Introuvable", timestamp()])
                nb_err += 1
            time.sleep(0.3)

        # Sauvegarder
        output_path = os.path.join(dossier_output, nom_output)
        df.to_excel(output_path, sheet_name=sheet_input, index=False)
        log(f"\n💾 Fichier enrichi sauvegardé : {output_path}")

        log(f"\n{sep}")
        log("RÉSUMÉ ENRICHISSEMENT")
        log(sep)
        log(f"  Total traité : {total}")
        log(f"  ✅ Succès    : {nb_ok}")
        log(f"  ❌ Erreurs   : {nb_err}")
        log(f"  📄 Log CSV   : {log_csv}")
        log(sep)
        fin(True)

    except Exception as e:
        log(f"\n❌ ERREUR FATALE : {e}")
        log(traceback.format_exc())
        fin(False)

# ==============================================================================
# LOGIQUE MÉTIER — PROJET 2 : RÉAFFECTATION
# ==============================================================================

USER_UUID_CACHE = {}

def _p2_get_user_uuid(client, username):
    if username in USER_UUID_CACHE:
        return USER_UUID_CACHE[username]
    try:
        users_api = UsersApi(client)
        for sup in list(users_api.list_supervisors()):
            if sup.get('UserName', '').lower() == username.lower():
                uid = sup.get('UserId')
                USER_UUID_CACHE[username] = uid
                return uid
        for sup in list(users_api.list_supervisors()):
            sup_id = sup.get('UserId')
            if sup_id:
                try:
                    for iv in list(users_api.list_interviewers(sup_id)):
                        if iv.get('UserName', '').lower() == username.lower():
                            uid = iv.get('UserId')
                            USER_UUID_CACHE[username] = uid
                            return uid
                except Exception:
                    continue
    except Exception:
        pass
    return None

def _p2_get_interview_info(interviews_api, interview_id):
    try:
        items = list(interviews_api.get_list(
            fields=['id', 'status', 'responsible_id', 'responsible_name'],
            key=interview_id, take=1
        ))
        if items:
            it = items[0]
            return {'status': it.status,
                    'responsible_id':   it.responsible_id,
                    'responsible_name': it.responsible_name}
    except Exception:
        pass
    try:
        info = interviews_api.get_info(interview_id)
        if info:
            return {'status':           info.get('Status'),
                    'responsible_id':   info.get('ResponsibleId'),
                    'responsible_name': info.get('ResponsibleName', '?')}
    except Exception:
        pass
    return None

def _p2_get_id_by_key(interviews_api, key):
    try:
        obj = interviews_api._get_interview_by_key(fields=["id"], key=key)
        if obj is None:
            return None
        for attr in ("id", "Id", "InterviewId"):
            if hasattr(obj, attr):
                return getattr(obj, attr)
        if isinstance(obj, dict):
            return obj.get("id") or obj.get("InterviewId")
    except Exception:
        pass
    return None

def _p2_smart_reassign(interviews_api, client, interview_id, resp_username, log):
    steps = []

    # Info actuelle
    info = _p2_get_interview_info(interviews_api, interview_id)
    current_status = info['status']        if info else "Inconnu"
    current_resp   = info['responsible_name'] if info else "Inconnu"
    current_rid    = info['responsible_id']   if info else ""
    log(f"   Statut actuel : {current_status} | Responsable : {current_resp}")

    # UUID cible
    resp_uuid = retry_operation(_p2_get_user_uuid, client, resp_username, max_retries=2)
    if not resp_uuid:
        msg = f"UUID introuvable pour '{resp_username}'"
        steps.append(("get_uuid", "ERROR", msg))
        return False, steps, msg
    log(f"   UUID trouvé   : {str(resp_uuid)[:12]}...")
    steps.append(("get_uuid", "OK", f"UUID:{str(resp_uuid)[:8]}"))

    # Déjà le bon responsable ?
    if current_rid and normalize_guid(str(current_rid)) == normalize_guid(str(resp_uuid)):
        if current_status in ['RejectedByHeadquarters', 'RejectedBySupervisor']:
            msg = f"Déjà rejeté vers {resp_username}"
        else:
            msg = f"Déjà affecté à {resp_username} (statut: {current_status})"
        log(f"   ℹ️  {msg}")
        steps.append(("skip", "OK", msg))
        return True, steps, msg

    # Tentative rejet
    reject_ok = False
    reject_not_allowed = False
    for fn_name, fn in [("hqreject", interviews_api.hqreject),
                        ("reject",   interviews_api.reject)]:
        if reject_ok or reject_not_allowed:
            break
        try:
            retry_operation(fn, interview_id, "Réaffectation automatique", max_retries=2)
            steps.append((fn_name, "OK", "Rejeté"))
            reject_ok = True
            log(f"   ✅ {fn_name} OK")
        except Exception as e:
            em = str(e).lower()
            if any(k in em for k in ['not allowed', 'invalid status',
                                     'cannot be rejected', 'already rejected']):
                reject_not_allowed = True
                steps.append((fn_name, "SKIPPED", "Rejet non autorisé"))
                log(f"   ℹ️  Rejet non autorisé ({fn_name}) — affectation directe")
            else:
                steps.append((fn_name, "ERROR", str(e)[:80]))
                log(f"   ⚠️  {fn_name} échoué : {str(e)[:60]}")

    if reject_ok:
        time.sleep(2)

    # Affectation
    try:
        retry_operation(
            interviews_api.assign,
            interview_id,
            resp_uuid,
            responsiblename=resp_username,
            max_retries=2
        )
        msg = (f"Affectation directe réussie (rejet non autorisé)"
               if reject_not_allowed else f"Affecté à {resp_username}")
        steps.append(("assign", "OK", msg))
        log(f"   ✅ {msg}")
        return True, steps, msg
    except Exception as e:
        msg = f"Erreur affectation : {str(e)[:100]}"
        steps.append(("assign", "ERROR", msg))
        log(f"   ❌ {msg}")
        return False, steps, msg

def run_reaffectation(params, log, fin):
    """Thread cible pour le Projet 2."""
    try:
        fichier_input   = params['fichier_input']
        sheet_name      = params['sheet_name']
        dossier_output  = params['dossier_output']
        col_key         = params['col_key']
        col_id          = params['col_id']
        col_resp        = params['col_resp']
        survey_url      = params['survey_url']
        workspace       = params['workspace']
        api_user        = params['api_user']
        api_password    = params['api_password']
        hq_username     = params['hq_username']
        hq_password     = params['hq_password']

        sep = "=" * 65
        log(sep)
        log("  PROJET 2 — RÉAFFECTATION INTELLIGENTE DES ENTRETIENS")
        log(sep)

        # Charger Excel
        log(f"\n📂 Chargement : {os.path.basename(fichier_input)}  [feuille: {sheet_name}]")
        df = pd.read_excel(fichier_input, sheet_name=sheet_name, dtype=str)
        df = df.fillna("")
        log(f"   ✓ {len(df)} lignes chargées")

        # Filtrer lignes valides
        df[col_key]  = df[col_key].astype(str).str.strip()
        df[col_id]   = df[col_id].astype(str).str.strip()
        df[col_resp] = df[col_resp].astype(str).str.strip()
        df_valid = df[
            ((df[col_id] != "") | (df[col_key] != "")) &
            (df[col_resp] != "")
        ].copy()
        log(f"   ✓ {len(df_valid)} lignes valides à traiter")

        if len(df_valid) == 0:
            log("   ℹ️  Aucune ligne à traiter.")
            fin(True)
            return

        # Connexion
        log("\n🔌 Connexion à Survey Solutions...")
        client, used = make_client(survey_url, workspace, api_user, api_password,
                                   hq_username, hq_password)
        log(f"   ✓ Connecté avec : {used}")
        interviews_api = InterviewsApi(client)

        # Log CSV
        os.makedirs(dossier_output, exist_ok=True)
        log_csv = os.path.join(dossier_output, "log_reaffectation.csv")
        with open(log_csv, "w", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([col_key, col_id, col_resp,
                                    "Steps", "Resultat", "Message", "Timestamp"])

        log("\n🔄 Réaffectation en cours...\n")
        nb_ok = nb_skip = nb_err = 0
        total = len(df_valid)

        for i, (idx, row) in enumerate(df_valid.iterrows(), 1):
            key      = row.get(col_key, "")
            iid_raw  = row.get(col_id,  "")
            resp     = row.get(col_resp, "").strip()
            log(f"\n{'─'*50}")
            log(f"[{i}/{total}]  KEY: {key[:22]}  →  {resp}")

            interview_id = None
            if iid_raw and looks_like_guid(iid_raw):
                interview_id = normalize_guid(iid_raw)
                log(f"   ID fourni : {interview_id[:12]}...")
            elif key:
                log("   Résolution par KEY...")
                try:
                    raw = retry_operation(_p2_get_id_by_key, interviews_api,
                                          key, max_retries=2)
                    if raw:
                        interview_id = normalize_guid(raw)
                        log(f"   ID trouvé  : {interview_id[:12]}...")
                    else:
                        log("   ❌ ID introuvable")
                except Exception as e:
                    log(f"   ❌ Erreur résolution KEY : {e}")

            if not interview_id:
                msg = "INTERVIEW_ID introuvable"
                log(f"   ❌ {msg}")
                with open(log_csv, "a", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow([key, iid_raw, resp, "",
                                            "ERROR", msg, timestamp()])
                nb_err += 1
                continue

            success, steps, message = _p2_smart_reassign(
                interviews_api, client, interview_id, resp, log)

            is_skip = any(s[0] == "skip" for s in steps)
            result  = "SKIP" if is_skip else ("OK" if success else "ERROR")
            steps_str = " → ".join(f"{s[0]}:{s[1]}" for s in steps)

            with open(log_csv, "a", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow([key, interview_id, resp,
                                        steps_str, result, message, timestamp()])

            if success:
                if is_skip:
                    nb_skip += 1
                else:
                    nb_ok += 1
            else:
                nb_err += 1
            time.sleep(0.5)

        log(f"\n{sep}")
        log("RÉSUMÉ RÉAFFECTATION")
        log(sep)
        log(f"  Total traité        : {total}")
        log(f"  ✅ Réaffectés       : {nb_ok}")
        log(f"  ⏭️  Déjà affectés    : {nb_skip}")
        log(f"  ❌ Erreurs          : {nb_err}")
        log(f"  📄 Log CSV          : {log_csv}")
        log(sep)
        fin(True)

    except Exception as e:
        log(f"\n❌ ERREUR FATALE : {e}")
        log(traceback.format_exc())
        fin(False)

# ==============================================================================
# FENÊTRE PROJET 1 — ENRICHISSEMENT
# ==============================================================================

class FenetreEnrichissement(tk.Toplevel):

    def __init__(self, master):
        super().__init__(master)
        self.title("Projet 1 — Enrichissement du fichier Excel")
        self.geometry("860x780")
        self.minsize(760, 640)
        self.configure(bg=COULEUR_FOND)

        # Variables
        self.v_fichier      = tk.StringVar()
        self.v_feuille      = tk.StringVar()
        self.v_dossier_out  = tk.StringVar()
        self.v_nom_output   = tk.StringVar(value="Base_rejet_Agent_enqueteur_V2_ENRICHI.xlsx")
        self.v_col_key      = tk.StringVar()
        self.v_col_id       = tk.StringVar()
        self.v_col_resp_name= tk.StringVar(value="Numero_agent_terrain")
        self.v_col_resp_id  = tk.StringVar(value="Numero_agent_terrain_ID")
        self.v_col_status   = tk.StringVar(value="STATUS_ACTUEL")
        self.v_survey_url   = tk.StringVar(value=DEFAULT_SURVEY_URL)
        self.v_workspace    = tk.StringVar(value=DEFAULT_WORKSPACE)
        self.v_api_user     = tk.StringVar(value=DEFAULT_API_USER)
        self.v_api_pwd      = tk.StringVar(value=DEFAULT_API_PASSWORD)
        self.v_hq_user      = tk.StringVar(value=DEFAULT_HQ_USERNAME)
        self.v_hq_pwd       = tk.StringVar(value=DEFAULT_HQ_PASSWORD)

        self._build()

    def _build(self):
        # En-tête
        hdr = tk.Frame(self, bg=COULEUR_TITRE, pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="PROJET 1 — ENRICHISSEMENT DU FICHIER EXCEL",
                 font=POLICE_TITRE, fg="white", bg=COULEUR_TITRE).pack()
        tk.Label(hdr,
                 text="Récupère le responsable actuel et le statut de chaque entretien "
                      "depuis Survey Solutions et enrichit le fichier Excel.",
                 font=("Segoe UI", 8), fg="#cde", bg=COULEUR_TITRE).pack()

        # Notebook
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook.Tab", font=POLICE_SECTION, padding=[10, 4],
                        background="#D0DFF0")
        style.map("TNotebook.Tab",
                  background=[("selected", COULEUR_SECTION)],
                  foreground=[("selected", "white")])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=6)

        t_params = tk.Frame(nb, bg=COULEUR_FOND)
        t_logs   = tk.Frame(nb, bg=COULEUR_FOND)
        nb.add(t_params, text="  ① Paramètres  ")
        nb.add(t_logs,   text="  ② Exécution & Logs  ")

        self._build_params(t_params)
        self._build_logs(t_logs)

    # ── Onglet Paramètres ────────────────────────────────────────────────────

    def _build_params(self, parent):
        frame = make_scrolled_frame(parent)
        XL = [("Fichiers Excel", "*.xlsx *.xls")]

        # Fichier input
        s1 = make_section(frame, "Fichier Input")
        r = make_file_row(s1, "Fichier Excel Input",
                          self.v_fichier, XL,
                          "Fichier contenant les entretiens à enrichir.", 0)
        # Feuille
        tk.Label(s1, text="Feuille", font=POLICE_NORMALE,
                 bg=COULEUR_FOND, anchor="w", width=26
                 ).grid(row=r, column=0, sticky="w", pady=2)
        self.combo_feuille = ttk.Combobox(s1, textvariable=self.v_feuille,
                                          width=34, font=POLICE_NORMALE, state="readonly")
        self.combo_feuille.grid(row=r, column=1, sticky="ew", padx=6)
        tk.Button(s1, text="🔄 Charger", font=POLICE_NORMALE, bg="#E8EFF8",
                  relief="flat", cursor="hand2",
                  command=self._charger_feuilles
                  ).grid(row=r, column=2, sticky="w", padx=4)
        s1.columnconfigure(1, weight=1)
        r += 1

        # Colonnes
        s2 = make_section(frame, "Colonnes du fichier Input")
        tk.Label(s2, text="▶ Après avoir sélectionné le fichier et la feuille, "
                          "cliquez sur 🔄 Charger colonnes pour remplir les listes.",
                 font=("Segoe UI", 8), fg="#555", bg=COULEUR_FOND, wraplength=540
                 ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        tk.Button(s2, text="🔄 Charger colonnes", font=POLICE_NORMALE,
                  bg="#E8EFF8", relief="flat", cursor="hand2",
                  command=self._charger_colonnes
                  ).grid(row=1, column=0, sticky="w", pady=4)
        s2.columnconfigure(1, weight=1)

        self.combo_col_key, r2 = make_combo_row(
            s2, "COL_KEY  (INTERVIEW_KEY) *", self.v_col_key, [],
            "Colonne contenant les clés d'entretien (format XX-XX-XX-XX).", 2)
        self.combo_col_id, r2 = make_combo_row(
            s2, "COL_ID   (INTERVIEW_ID) *", self.v_col_id, [],
            "Colonne contenant les ID GUID (ex: 9d5496fafd9549f8ba026a2ef006c588).", r2)

        # Colonnes créées par le programme
        s3 = make_section(frame, "Colonnes créées par le programme (noms personnalisables)")
        make_entry_row(s3, "COL_RESP_NAME", self.v_col_resp_name,
                       "Nom de la colonne qui recevra le nom du responsable.", 0)
        make_entry_row(s3, "COL_RESP_ID", self.v_col_resp_id,
                       "Nom de la colonne qui recevra l'ID du responsable.", 2)
        make_entry_row(s3, "COL_STATUS", self.v_col_status,
                       "Nom de la colonne qui recevra le statut de l'entretien.", 4)

        # Fichier output
        s4 = make_section(frame, "Fichier Output")
        make_dir_row(s4, "Dossier de sortie", self.v_dossier_out,
                     "Dossier où sera enregistré le fichier enrichi et le log CSV.", 0)
        make_entry_row(s4, "Nom du fichier de sortie", self.v_nom_output,
                       "Nom du fichier Excel enrichi (défaut proposé ci-dessus).", 3)

        # Config Survey Solutions
        s5 = make_section(frame, "Configuration Survey Solutions")
        tk.Label(s5, text="⚙️  Modifiez ces valeurs si nécessaire. "
                          "Les valeurs par défaut correspondent à l'environnement de production.",
                 font=("Segoe UI", 8), fg="#555", bg=COULEUR_FOND, wraplength=540
                 ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        make_entry_row(s5, "SURVEY_URL",    self.v_survey_url,    None, 1)
        make_entry_row(s5, "WORKSPACE",     self.v_workspace,     None, 2)
        make_entry_row(s5, "API_USER",      self.v_api_user,      None, 3)
        make_entry_row(s5, "API_PASSWORD",  self.v_api_pwd,       None, 4)
        make_entry_row(s5, "HQ_USERNAME",   self.v_hq_user,       None, 5)
        make_entry_row(s5, "HQ_PASSWORD",   self.v_hq_pwd,        None, 6)

    # ── Onglet Logs ──────────────────────────────────────────────────────────

    def _build_logs(self, parent):
        bar = tk.Frame(parent, bg=COULEUR_FOND, pady=6)
        bar.pack(fill="x", padx=10)

        self.btn_lancer = tk.Button(
            bar, text="▶  LANCER L'ENRICHISSEMENT",
            font=("Segoe UI", 10, "bold"),
            bg=COULEUR_BTN_OK, fg="white",
            padx=16, pady=7, relief="flat", cursor="hand2",
            command=self._lancer
        )
        self.btn_lancer.pack(side="left", padx=(0, 10))

        tk.Button(bar, text="🗑 Effacer", font=POLICE_NORMALE,
                  bg="#D5D5D5", relief="flat", cursor="hand2",
                  command=lambda: self.zone_logs.delete("1.0", "end")
                  ).pack(side="left")

        self.lbl_statut = tk.Label(bar, text="", font=("Segoe UI", 9, "bold"),
                                   bg=COULEUR_FOND)
        self.lbl_statut.pack(side="left", padx=14)

        self.zone_logs = make_logs_zone(parent)

    # ── Actions ──────────────────────────────────────────────────────────────

    def _charger_feuilles(self):
        f = self.v_fichier.get()
        if not f or not os.path.exists(f):
            messagebox.showwarning("Fichier manquant",
                                   "Veuillez d'abord sélectionner un fichier Excel valide.",
                                   parent=self)
            return
        try:
            sheets = get_sheets(f)
            self.combo_feuille['values'] = sheets
            if sheets:
                self.combo_feuille.current(0)
        except Exception as e:
            messagebox.showerror("Erreur", str(e), parent=self)

    def _charger_colonnes(self):
        f  = self.v_fichier.get()
        sh = self.v_feuille.get()
        if not f or not sh:
            messagebox.showwarning("Information manquante",
                                   "Sélectionnez d'abord un fichier et une feuille.",
                                   parent=self)
            return
        try:
            cols = get_columns(f, sh)
            for combo in [self.combo_col_key, self.combo_col_id]:
                combo['values'] = cols
            messagebox.showinfo("Colonnes chargées",
                                f"{len(cols)} colonne(s) trouvée(s).", parent=self)
        except Exception as e:
            messagebox.showerror("Erreur", str(e), parent=self)

    def _lancer(self):
        # Validation
        erreurs = []
        if not self.v_fichier.get()     : erreurs.append("Fichier Input manquant.")
        if not self.v_feuille.get()     : erreurs.append("Feuille non sélectionnée.")
        if not self.v_dossier_out.get() : erreurs.append("Dossier de sortie manquant.")
        if not self.v_nom_output.get()  : erreurs.append("Nom du fichier de sortie manquant.")
        if not self.v_col_key.get() and not self.v_col_id.get():
            erreurs.append("Sélectionnez au moins COL_KEY ou COL_ID.")
        if erreurs:
            messagebox.showerror("Paramètres invalides",
                                 "\n".join(f"• {e}" for e in erreurs), parent=self)
            return

        self.btn_lancer.config(state="disabled", text="⏳ En cours…")
        self.lbl_statut.config(text="⏳ En cours…", fg="#E3B341")

        params = {
            'fichier_input':  self.v_fichier.get(),
            'sheet_input':    self.v_feuille.get(),
            'dossier_output': self.v_dossier_out.get(),
            'nom_output':     self.v_nom_output.get(),
            'col_key':        self.v_col_key.get(),
            'col_id':         self.v_col_id.get(),
            'col_resp_name':  self.v_col_resp_name.get(),
            'col_resp_id':    self.v_col_resp_id.get(),
            'col_status':     self.v_col_status.get(),
            'survey_url':     self.v_survey_url.get(),
            'workspace':      self.v_workspace.get(),
            'api_user':       self.v_api_user.get(),
            'api_password':   self.v_api_pwd.get(),
            'hq_username':    self.v_hq_user.get(),
            'hq_password':    self.v_hq_pwd.get(),
        }

        def log(msg):
            log_write(self.zone_logs, msg, self)

        def fin(ok):
            def _ui():
                self.btn_lancer.config(state="normal",
                                       text="▶  LANCER L'ENRICHISSEMENT")
                if ok:
                    self.lbl_statut.config(text="✅ Terminé !", fg="#56D364")
                    messagebox.showinfo("Succès",
                                        "Enrichissement terminé avec succès !\n"
                                        f"Fichier : {self.v_nom_output.get()}\n"
                                        f"Dossier : {self.v_dossier_out.get()}",
                                        parent=self)
                else:
                    self.lbl_statut.config(text="❌ Erreur", fg="#F85149")
                    messagebox.showerror("Erreur",
                                         "Une erreur est survenue. Voir les logs.",
                                         parent=self)
            self.after(0, _ui)

        threading.Thread(target=run_enrichissement,
                         args=(params, log, fin), daemon=True).start()

# ==============================================================================
# FENÊTRE PROJET 2 — RÉAFFECTATION
# ==============================================================================

class FenetreReaffectation(tk.Toplevel):

    def __init__(self, master):
        super().__init__(master)
        self.title("Projet 2 — Réaffectation intelligente des entretiens")
        self.geometry("860x780")
        self.minsize(760, 640)
        self.configure(bg=COULEUR_FOND)

        self.v_fichier     = tk.StringVar()
        self.v_feuille     = tk.StringVar()
        self.v_dossier_out = tk.StringVar()
        self.v_col_key     = tk.StringVar()
        self.v_col_id      = tk.StringVar()
        self.v_col_resp    = tk.StringVar()
        self.v_survey_url  = tk.StringVar(value=DEFAULT_SURVEY_URL)
        self.v_workspace   = tk.StringVar(value=DEFAULT_WORKSPACE)
        self.v_api_user    = tk.StringVar(value=DEFAULT_API_USER)
        self.v_api_pwd     = tk.StringVar(value=DEFAULT_API_PASSWORD)
        self.v_hq_user     = tk.StringVar(value=DEFAULT_HQ_USERNAME)
        self.v_hq_pwd      = tk.StringVar(value=DEFAULT_HQ_PASSWORD)

        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg="#1A5C3A", pady=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="PROJET 2 — RÉAFFECTATION INTELLIGENTE DES ENTRETIENS",
                 font=POLICE_TITRE, fg="white", bg="#1A5C3A").pack()
        tk.Label(hdr,
                 text="Vérifie le statut de chaque entretien et le réaffecte "
                      "à l'agent enquêteur/téléopérateur désigné.",
                 font=("Segoe UI", 8), fg="#cec", bg="#1A5C3A").pack()

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook.Tab", font=POLICE_SECTION, padding=[10, 4],
                        background="#D0DFF0")
        style.map("TNotebook.Tab",
                  background=[("selected", "#1A5C3A")],
                  foreground=[("selected", "white")])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=6)

        t_params = tk.Frame(nb, bg=COULEUR_FOND)
        t_logs   = tk.Frame(nb, bg=COULEUR_FOND)
        nb.add(t_params, text="  ① Paramètres  ")
        nb.add(t_logs,   text="  ② Exécution & Logs  ")

        self._build_params(t_params)
        self._build_logs(t_logs)

    def _build_params(self, parent):
        frame = make_scrolled_frame(parent)
        XL = [("Fichiers Excel", "*.xlsx *.xls")]

        s1 = make_section(frame, "Fichier Input")
        r = make_file_row(s1, "Fichier Excel Input",
                          self.v_fichier, XL,
                          "Fichier contenant les entretiens à réaffecter.\n"
                          "Peut être le fichier enrichi produit par le Projet 1, "
                          "ou tout autre fichier Excel valide.", 0)
        tk.Label(s1, text="Feuille", font=POLICE_NORMALE,
                 bg=COULEUR_FOND, anchor="w", width=26
                 ).grid(row=r, column=0, sticky="w", pady=2)
        self.combo_feuille = ttk.Combobox(s1, textvariable=self.v_feuille,
                                          width=34, font=POLICE_NORMALE, state="readonly")
        self.combo_feuille.grid(row=r, column=1, sticky="ew", padx=6)
        tk.Button(s1, text="🔄 Charger", font=POLICE_NORMALE, bg="#E8EFF8",
                  relief="flat", cursor="hand2",
                  command=self._charger_feuilles
                  ).grid(row=r, column=2, sticky="w", padx=4)
        s1.columnconfigure(1, weight=1)

        s2 = make_section(frame, "Colonnes du fichier Input")
        tk.Label(s2,
                 text="▶ Après avoir sélectionné le fichier et la feuille, "
                      "cliquez sur 🔄 Charger colonnes.",
                 font=("Segoe UI", 8), fg="#555", bg=COULEUR_FOND, wraplength=540
                 ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        tk.Button(s2, text="🔄 Charger colonnes", font=POLICE_NORMALE,
                  bg="#E8EFF8", relief="flat", cursor="hand2",
                  command=self._charger_colonnes
                  ).grid(row=1, column=0, sticky="w", pady=4)
        s2.columnconfigure(1, weight=1)

        self.combo_col_key, r3 = make_combo_row(
            s2, "COL_KEY  (INTERVIEW_KEY) *", self.v_col_key, [],
            "Colonne contenant les clés d'entretien (format XX-XX-XX-XX).", 2)
        self.combo_col_id, r3 = make_combo_row(
            s2, "COL_ID   (INTERVIEW_ID) *", self.v_col_id, [],
            "Colonne contenant les ID GUID (ex: 9d5496fafd9549f8ba026a2ef006c588).", r3)
        self.combo_col_resp, r3 = make_combo_row(
            s2, "COL_RESP  (Agent cible) *", self.v_col_resp, [],
            "Colonne contenant le numéro de compte de l'agent destinataire "
            "(ex: Eqpe29Agent1).", r3)

        s3 = make_section(frame, "Dossier de sortie (logs)")
        make_dir_row(s3, "Dossier de sortie", self.v_dossier_out,
                     "Dossier où sera enregistré le log CSV de réaffectation.", 0)

        s4 = make_section(frame, "Configuration Survey Solutions")
        tk.Label(s4,
                 text="⚙️  Modifiez ces valeurs si nécessaire.",
                 font=("Segoe UI", 8), fg="#555", bg=COULEUR_FOND
                 ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        make_entry_row(s4, "SURVEY_URL",   self.v_survey_url,  None, 1)
        make_entry_row(s4, "WORKSPACE",    self.v_workspace,   None, 2)
        make_entry_row(s4, "API_USER",     self.v_api_user,    None, 3)
        make_entry_row(s4, "API_PASSWORD", self.v_api_pwd,     None, 4)
        make_entry_row(s4, "HQ_USERNAME",  self.v_hq_user,     None, 5)
        make_entry_row(s4, "HQ_PASSWORD",  self.v_hq_pwd,      None, 6)

    def _build_logs(self, parent):
        bar = tk.Frame(parent, bg=COULEUR_FOND, pady=6)
        bar.pack(fill="x", padx=10)

        self.btn_lancer = tk.Button(
            bar, text="▶  LANCER LA RÉAFFECTATION",
            font=("Segoe UI", 10, "bold"),
            bg="#1A5C3A", fg="white",
            padx=16, pady=7, relief="flat", cursor="hand2",
            command=self._lancer
        )
        self.btn_lancer.pack(side="left", padx=(0, 10))

        tk.Button(bar, text="🗑 Effacer", font=POLICE_NORMALE,
                  bg="#D5D5D5", relief="flat", cursor="hand2",
                  command=lambda: self.zone_logs.delete("1.0", "end")
                  ).pack(side="left")

        self.lbl_statut = tk.Label(bar, text="", font=("Segoe UI", 9, "bold"),
                                   bg=COULEUR_FOND)
        self.lbl_statut.pack(side="left", padx=14)

        self.zone_logs = make_logs_zone(parent)

    def _charger_feuilles(self):
        f = self.v_fichier.get()
        if not f or not os.path.exists(f):
            messagebox.showwarning("Fichier manquant",
                                   "Sélectionnez d'abord un fichier Excel valide.",
                                   parent=self)
            return
        try:
            sheets = get_sheets(f)
            self.combo_feuille['values'] = sheets
            if sheets:
                self.combo_feuille.current(0)
        except Exception as e:
            messagebox.showerror("Erreur", str(e), parent=self)

    def _charger_colonnes(self):
        f  = self.v_fichier.get()
        sh = self.v_feuille.get()
        if not f or not sh:
            messagebox.showwarning("Information manquante",
                                   "Sélectionnez d'abord un fichier et une feuille.",
                                   parent=self)
            return
        try:
            cols = get_columns(f, sh)
            for combo in [self.combo_col_key, self.combo_col_id, self.combo_col_resp]:
                combo['values'] = cols
            messagebox.showinfo("Colonnes chargées",
                                f"{len(cols)} colonne(s) trouvée(s).", parent=self)
        except Exception as e:
            messagebox.showerror("Erreur", str(e), parent=self)

    def _lancer(self):
        erreurs = []
        if not self.v_fichier.get()     : erreurs.append("Fichier Input manquant.")
        if not self.v_feuille.get()     : erreurs.append("Feuille non sélectionnée.")
        if not self.v_dossier_out.get() : erreurs.append("Dossier de sortie manquant.")
        if not self.v_col_key.get() and not self.v_col_id.get():
            erreurs.append("Sélectionnez au moins COL_KEY ou COL_ID.")
        if not self.v_col_resp.get()    : erreurs.append("COL_RESP non sélectionnée.")
        if erreurs:
            messagebox.showerror("Paramètres invalides",
                                 "\n".join(f"• {e}" for e in erreurs), parent=self)
            return

        self.btn_lancer.config(state="disabled", text="⏳ En cours…")
        self.lbl_statut.config(text="⏳ En cours…", fg="#E3B341")

        params = {
            'fichier_input':  self.v_fichier.get(),
            'sheet_name':     self.v_feuille.get(),
            'dossier_output': self.v_dossier_out.get(),
            'col_key':        self.v_col_key.get(),
            'col_id':         self.v_col_id.get(),
            'col_resp':       self.v_col_resp.get(),
            'survey_url':     self.v_survey_url.get(),
            'workspace':      self.v_workspace.get(),
            'api_user':       self.v_api_user.get(),
            'api_password':   self.v_api_pwd.get(),
            'hq_username':    self.v_hq_user.get(),
            'hq_password':    self.v_hq_pwd.get(),
        }

        def log(msg):
            log_write(self.zone_logs, msg, self)

        def fin(ok):
            def _ui():
                self.btn_lancer.config(state="normal",
                                       text="▶  LANCER LA RÉAFFECTATION")
                if ok:
                    self.lbl_statut.config(text="✅ Terminé !", fg="#56D364")
                    messagebox.showinfo("Succès",
                                        "Réaffectation terminée avec succès !\n"
                                        f"Logs : {self.v_dossier_out.get()}",
                                        parent=self)
                else:
                    self.lbl_statut.config(text="❌ Erreur", fg="#F85149")
                    messagebox.showerror("Erreur",
                                         "Une erreur est survenue. Voir les logs.",
                                         parent=self)
            self.after(0, _ui)

        threading.Thread(target=run_reaffectation,
                         args=(params, log, fin), daemon=True).start()

# ==============================================================================
# FENÊTRE PRINCIPALE — MENU D'ACCUEIL
# ==============================================================================

class MenuAccueil(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Gestion des entretiens Survey Solutions — ENEM")
        self.geometry("640x480")
        self.resizable(False, False)
        self.configure(bg=COULEUR_FOND)
        self._build()

    def _build(self):
        # ── En-tête ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=COULEUR_TITRE, pady=18)
        hdr.pack(fill="x")
        tk.Label(hdr,
                 text="GESTION DES ENTRETIENS\nSURVEY SOLUTIONS — ENEM",
                 font=("Segoe UI", 15, "bold"), fg="white",
                 bg=COULEUR_TITRE, justify="center").pack()
        tk.Label(hdr,
                 text="Application de gestion des entretiens ménages",
                 font=("Segoe UI", 9), fg="#AAC8E8",
                 bg=COULEUR_TITRE).pack(pady=(4, 0))

        # ── Avertissement ssaw ────────────────────────────────────────────────
        if not SSAW_OK:
            warn = tk.Frame(self, bg="#7B2D00", pady=6)
            warn.pack(fill="x")
            tk.Label(warn,
                     text="⚠️  Le module 'ssaw' n'est pas installé. "
                          "Lancez : pip install ssaw",
                     font=("Segoe UI", 9, "bold"), fg="#FFD580",
                     bg="#7B2D00").pack()

        # ── Corps ─────────────────────────────────────────────────────────────
        corps = tk.Frame(self, bg=COULEUR_FOND)
        corps.pack(fill="both", expand=True, padx=40, pady=20)

        tk.Label(corps,
                 text="Choisissez un module à ouvrir :",
                 font=("Segoe UI", 11), fg=COULEUR_TITRE,
                 bg=COULEUR_FOND).pack(pady=(0, 20))

        # Bouton Projet 1
        f1 = tk.Frame(corps, bg="#E8F0FB", relief="groove", bd=1)
        f1.pack(fill="x", pady=8, ipady=10)
        tk.Label(f1,
                 text="📋  PROJET 1 — Enrichissement",
                 font=("Segoe UI", 11, "bold"), fg=COULEUR_TITRE,
                 bg="#E8F0FB").pack()
        tk.Label(f1,
                 text="Récupère le responsable actuel et le statut de chaque entretien\n"
                      "depuis Survey Solutions et enrichit le fichier Excel.",
                 font=("Segoe UI", 8), fg="#444",
                 bg="#E8F0FB", justify="center").pack(pady=(2, 6))
        tk.Button(f1,
                  text="▶  Ouvrir le module Enrichissement",
                  font=("Segoe UI", 10, "bold"),
                  bg=COULEUR_BTN, fg="white",
                  padx=14, pady=6, relief="flat", cursor="hand2",
                  command=self._ouvrir_enrichissement
                  ).pack()

        # Bouton Projet 2
        f2 = tk.Frame(corps, bg="#E8F5EE", relief="groove", bd=1)
        f2.pack(fill="x", pady=8, ipady=10)
        tk.Label(f2,
                 text="🔄  PROJET 2 — Réaffectation intelligente",
                 font=("Segoe UI", 11, "bold"), fg="#1A5C3A",
                 bg="#E8F5EE").pack()
        tk.Label(f2,
                 text="Vérifie le statut de chaque entretien et le réaffecte\n"
                      "à l'agent enquêteur/téléopérateur désigné.",
                 font=("Segoe UI", 8), fg="#444",
                 bg="#E8F5EE", justify="center").pack(pady=(2, 6))
        tk.Button(f2,
                  text="▶  Ouvrir le module Réaffectation",
                  font=("Segoe UI", 10, "bold"),
                  bg="#1A5C3A", fg="white",
                  padx=14, pady=6, relief="flat", cursor="hand2",
                  command=self._ouvrir_reaffectation
                  ).pack()

        # ── Bouton Quitter ────────────────────────────────────────────────────
        tk.Button(
            self,
            text="✖   QUITTER L'APPLICATION",
            font=("Segoe UI", 10, "bold"),
            bg=COULEUR_BTN_QUIT, fg="white",
            padx=20, pady=8, relief="flat", cursor="hand2",
            command=self._quitter
        ).pack(pady=12)

        # ── Pied de page ──────────────────────────────────────────────────────
        tk.Label(self,
                 text="© mg.kouame — ENEM Survey Solutions Manager",
                 font=("Segoe UI", 7), fg="#999",
                 bg=COULEUR_FOND).pack(side="bottom", pady=4)

    def _ouvrir_enrichissement(self):
        win = FenetreEnrichissement(self)
        win.grab_set()

    def _ouvrir_reaffectation(self):
        win = FenetreReaffectation(self)
        win.grab_set()

    def _quitter(self):
        if messagebox.askyesno("Quitter",
                               "Voulez-vous vraiment quitter l'application ?",
                               parent=self):
            self.destroy()

# ==============================================================================
# POINT D'ENTRÉE
# ==============================================================================

if __name__ == "__main__":
    # ── Étape 1 : Auto-installation des dépendances ──────────────────────────
    modules_manquants = verifier_et_installer_modules()

    # Recharger pandas et ssaw après installation éventuelle
    try:
        import pandas as pd
    except ImportError:
        pd = None

    try:
        import ssaw
        from ssaw import InterviewsApi, UsersApi
        SSAW_OK = True
    except ImportError:
        SSAW_OK = False

    # ── Étape 2 : Lancer l'application principale ────────────────────────────
    app = MenuAccueil()
    app.mainloop()
