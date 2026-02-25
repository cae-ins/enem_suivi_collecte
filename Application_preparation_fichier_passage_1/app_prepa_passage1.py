# -*- coding: utf-8 -*-
"""
================================================================================
APPLICATION : Interface de préparation des fichiers de collecte terrain (Passage 1)
================================================================================
AUTEUR  : KOUAME KOUASSI GUY MARTIAL
VERSION : 3.0 - Interface Tkinter
================================================================================
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import pandas as pd
import numpy as np

# ==============================================================================
# LOGIQUE MÉTIER (traitement des données)
# ==============================================================================

def formater_date(val):
    if pd.isna(val) or val is None:
        return ""
    if hasattr(val, 'strftime'):
        return val.strftime('%Y-%m-%d')
    s = str(val).strip()
    return "" if s in ('NaT', 'nan', 'None', '') else s

def lancer_traitement(params, log_func, fin_func):
    """Exécute tout le pipeline de traitement. Appelé dans un thread séparé."""
    try:
        TRIMESTRE_COLLECTE      = params['TRIMESTRE_COLLECTE']
        TRIMESTRE_COLLECTE_DATE = params['TRIMESTRE_COLLECTE_DATE']
        NUMERO_TRIMESTRE        = params['NUMERO_TRIMESTRE']
        MOIS_COLLECTE           = params['MOIS_COLLECTE']
        ANNEE_COLLECTE          = params['ANNEE_COLLECTE']
        SOUS_ECHANTILLON        = params['SOUS_ECHANTILLON']
        NOMBRE_MENAGES_PAR_ZD   = params['NOMBRE_MENAGES_PAR_ZD']
        DOSSIER_TRAVAIL_RESULTAT= params['DOSSIER_TRAVAIL_RESULTAT']
        FICHIER_ECHANTILLON     = params['FICHIER_ECHANTILLON']
        FEUILLE_ECHANTILLON     = params['FEUILLE_ECHANTILLON']
        FICHIER_GEOREF          = params['FICHIER_GEOREF']
        FICHIER_EQUIPES         = params['FICHIER_EQUIPES']
        FEUILLE_EQUIPES         = params['FEUILLE_EQUIPES']
        FICHIER_SEMAINES_REF    = params['FICHIER_SEMAINES_REF']
        FEUILLE_SEMAINES_REF    = "Semaine_ref_trim"

        # Noms fichiers de sortie
        NOM_DENOMBREMENT        = f"Denombrement_{TRIMESTRE_COLLECTE}.xlsx"
        NOM_MENAGE              = f"Menage_{TRIMESTRE_COLLECTE}.xlsx"
        NOM_CALENDRIER          = f"Calendrier_menage_{TRIMESTRE_COLLECTE}.xlsx"
        NOM_MENAGE_ABI          = f"Menage_ABIDJAN_{TRIMESTRE_COLLECTE}.xlsx"
        NOM_MENAGE_HORS         = f"Menage_HORS_ABIDJAN_{TRIMESTRE_COLLECTE}.xlsx"
        NOM_DENOM_ABI           = f"Denombrement_ABIDJAN_{TRIMESTRE_COLLECTE}.xlsx"
        NOM_DENOM_HORS          = f"Denombrement_HORS_ABIDJAN_{TRIMESTRE_COLLECTE}.xlsx"

        sep = "=" * 70

        log_func(sep)
        log_func("  PRÉPARATION DES FICHIERS DE COLLECTE - PASSAGE 1")
        log_func(sep)
        log_func(f"  Trimestre        : {TRIMESTRE_COLLECTE}")
        log_func(f"  Année            : {ANNEE_COLLECTE}")
        log_func(f"  Mois             : {MOIS_COLLECTE}")
        log_func(f"  Sous-échantillon : {SOUS_ECHANTILLON}")
        log_func(f"  Ménages/ZD       : {NOMBRE_MENAGES_PAR_ZD}")
        log_func(sep)

        # ── CHARGEMENT ─────────────────────────────────────────────────────────
        log_func("\n📂 CHARGEMENT DES DONNÉES...")

        df_echantillon = pd.read_excel(FICHIER_ECHANTILLON, sheet_name=FEUILLE_ECHANTILLON)
        df_echantillon = df_echantillon[df_echantillon["sous_echant"] == SOUS_ECHANTILLON]

        df_georef      = pd.read_excel(FICHIER_GEOREF)
        df_equipes     = pd.read_excel(FICHIER_EQUIPES, sheet_name=FEUILLE_EQUIPES)
        df_equipes     = df_equipes[df_equipes["Type de compte"] == "Agent de collecte"]
        df_semaines_ref= pd.read_excel(FICHIER_SEMAINES_REF, sheet_name=FEUILLE_SEMAINES_REF)

        log_func(f"   ✓ Échantillon          : {len(df_echantillon)} ZD")
        log_func(f"   ✓ Géoréférencement     : {len(df_georef)} enregistrements")
        log_func(f"   ✓ Agents de collecte   : {len(df_equipes)} agents")
        log_func(f"   ✓ Semaines de référence: {len(df_semaines_ref)} lignes")

        # ── PRÉPARATION & FUSION ───────────────────────────────────────────────
        log_func("\n🔧 PRÉPARATION ET FUSION DES DONNÉES...")

        df_echantillon['NUM_ZD_Vf'] = df_echantillon['NUM_ZD_Vf'].astype(str).str.zfill(4)
        df_echantillon["CLEZD"]     = df_echantillon["NomSp"] + df_echantillon["NUM_ZD_Vf"]

        data_merged = pd.merge(df_echantillon, df_georef, on='CLEZD', how='left')
        cols_x = [c for c in data_merged.columns if c.endswith('_x')]
        data_merged.drop(columns=cols_x, inplace=True)
        data_merged.rename(columns=lambda c: c.rstrip('_y') if c.endswith('_y') else c, inplace=True)

        log_func(f"   ✓ Fusion géoréférencement : {len(data_merged)} enregistrements")

        # ── AFFECTATION AGENTS ─────────────────────────────────────────────────
        log_func("\n👥 AFFECTATION DES AGENTS PAR RÉGION...")

        df_cross    = pd.merge(data_merged, df_equipes, how='cross')
        df_resultat = df_cross[df_cross['Region'] == df_cross['NomReg']].copy()
        df_resultat['NUM_ZD_Vf'] = df_resultat['NUM_ZD_Vf'].astype(str).str.zfill(4)

        log_func(f"   ✓ {len(df_resultat)} affectations ZD-Agent créées")
        repartition = df_resultat.groupby('Region').size().sort_values(ascending=False)
        for region, count in list(repartition.items())[:10]:
            log_func(f"      • {region}: {count} affectations")
        if len(repartition) > 10:
            log_func(f"      ... et {len(repartition)-10} autres régions")

        # ── CODE1 ──────────────────────────────────────────────────────────────
        log_func("\n🔢 CRÉATION DE LA VARIABLE CODE1...")

        df_resultat['cle_grp'] = (
            df_resultat['NUM_ZD_Vf'].astype(str) + "_" + df_resultat['LibQtierCpt'].astype(str)
        )
        np.random.seed(42)
        combs = df_resultat['cle_grp'].unique()
        mapping = {c: np.random.randint(10000000, 100000000) for c in combs}
        df_resultat['Code1'] = "A" + df_resultat['cle_grp'].map(mapping).astype(str)
        df_resultat.drop(columns=['cle_grp'], inplace=True)
        log_func(f"   ✓ {len(combs)} codes uniques générés")

        # ── ORDRE ──────────────────────────────────────────────────────────────
        log_func("\n📋 CRÉATION DE LA VARIABLE ORDRE...")

        corresp_ordre = {1:1, 3:2, 5:3, 7:4, 9:5, 11:6, 13:7}

        def calculer_ordre(row):
            if row['NomReg'] == 'ABIDJAN':
                return row['semaine_ref']
            return corresp_ordre.get(row['semaine_ref'], np.nan)

        df_resultat['Ordre'] = df_resultat.apply(calculer_ordre, axis=1).fillna(0).astype(int)
        nb_zero = (df_resultat['Ordre'] == 0).sum()
        if nb_zero > 0:
            log_func(f"   ⚠️  {nb_zero} affectations avec Ordre=0 (semaine_ref invalide)")
        else:
            log_func("   ✓ Toutes les affectations ont un Ordre valide")

        # ── DATES DE RÉFÉRENCE (Dénombrement & Ménage) ─────────────────────────
        log_func("\n📅 PRÉPARATION DES DATES DE RÉFÉRENCE...")

        df_sem_trim = df_semaines_ref[df_semaines_ref['Trimestre'] == TRIMESTRE_COLLECTE_DATE].copy()

        if len(df_sem_trim) == 0:
            log_func(f"   ⚠️  ERREUR : Aucune ligne pour '{TRIMESTRE_COLLECTE_DATE}'")
            log_func(f"   ⚠️  Valeurs disponibles : {df_semaines_ref['Trimestre'].unique()}")
            df_resultat['Date1_ref'] = ""
            df_resultat['Date2_ref'] = ""
        else:
            log_func(f"   ✓ {len(df_sem_trim)} semaines trouvées pour {TRIMESTRE_COLLECTE}")
            lookup_n = {
                row['N_semaine']: (row['Date1'], row['Date2'])
                for _, row in df_sem_trim.iterrows()
            }
            df_resultat['Date1_ref'] = df_resultat['semaine_ref'].map(
                lambda s: lookup_n.get(s, (None, None))[0])
            df_resultat['Date2_ref'] = df_resultat['semaine_ref'].map(
                lambda s: lookup_n.get(s, (None, None))[1])

            for col in ['Date1_ref', 'Date2_ref']:
                if pd.api.types.is_datetime64_any_dtype(df_resultat[col]):
                    df_resultat[col] = df_resultat[col].dt.strftime('%Y-%m-%d')
                df_resultat[col] = (
                    df_resultat[col].fillna("").astype(str)
                    .replace(['NaT', 'nan', 'None'], ""))

            nb_ok = (df_resultat['Date1_ref'] != "").sum()
            log_func(f"   ✅ Dates trouvées : {nb_ok} / {len(df_resultat)} lignes")

        # ── DÉNOMBREMENT ───────────────────────────────────────────────────────
        log_func("\n📝 GÉNÉRATION DU FICHIER DÉNOMBREMENT...")

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
            'HH8A'            : np.where(df_resultat['Plusieurs Loc']==1, df_resultat['NomLoc'], 'Zd sur plusieurs localité'),
            'HH7'             : np.where(df_resultat['Zd campement']=="Pas  campement", 1, 7),
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

        denom_abi  = denombrement[denombrement['Region'] == 'ABIDJAN'].copy()
        denom_hors = denombrement[denombrement['Region'] != 'ABIDJAN'].copy()

        denombrement.to_excel(os.path.join(DOSSIER_TRAVAIL_RESULTAT, NOM_DENOMBREMENT),  index=False)
        denom_abi.to_excel(  os.path.join(DOSSIER_TRAVAIL_RESULTAT, NOM_DENOM_ABI),      index=False)
        denom_hors.to_excel( os.path.join(DOSSIER_TRAVAIL_RESULTAT, NOM_DENOM_HORS),     index=False)

        log_func(f"   ✓ {NOM_DENOMBREMENT} ({len(denombrement)} lignes)")
        log_func(f"   ✓ {NOM_DENOM_ABI} ({len(denom_abi)} lignes)")
        log_func(f"   ✓ {NOM_DENOM_HORS} ({len(denom_hors)} lignes)")

        # ── MÉNAGE ─────────────────────────────────────────────────────────────
        log_func(f"\n🏠 GÉNÉRATION DU FICHIER MÉNAGE ({NOMBRE_MENAGES_PAR_ZD} ménages/ZD)...")

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
            'HH8A'            : np.where(df_resultat['Plusieurs Loc']==1, df_resultat['NomLoc'], 'Zd sur plusieurs localité'),
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

        menage_abi  = menage[menage['Region'] == 'ABIDJAN'].copy()
        men_ref1    = menage[menage['Region'] != 'ABIDJAN'].copy()
        men_ref2    = men_ref1.copy()
        men_ref2['Reference'] = 2

        cle_tri = men_ref1['Region'].astype(str)+"_"+men_ref1['sp'].astype(str)+"_"+men_ref1['HH8'].astype(str)
        men_ref1 = men_ref1.assign(cle_tri=cle_tri)
        men_ref2 = men_ref2.assign(cle_tri=cle_tri)

        men_combine = (
            pd.concat([men_ref1, men_ref2], ignore_index=True)
            .sort_values(['cle_tri','Reference'])
            .drop(columns=['cle_tri'])
        )
        menage_final = pd.concat([menage_abi, men_combine], ignore_index=True)

        menage_final.to_excel( os.path.join(DOSSIER_TRAVAIL_RESULTAT, NOM_MENAGE),     index=False)
        menage_abi.to_excel(   os.path.join(DOSSIER_TRAVAIL_RESULTAT, NOM_MENAGE_ABI), index=False)
        men_combine.to_excel(  os.path.join(DOSSIER_TRAVAIL_RESULTAT, NOM_MENAGE_HORS),index=False)

        log_func(f"   ✓ {NOM_MENAGE} ({len(menage_final)} lignes)")
        log_func(f"   ✓ {NOM_MENAGE_ABI} ({len(menage_abi)} lignes)")
        log_func(f"   ✓ {NOM_MENAGE_HORS} ({len(men_combine)} lignes)")

        nb1 = (men_combine['Reference']==1).sum()
        nb2 = (men_combine['Reference']==2).sum()
        if nb1 == nb2:
            log_func(f"   ✅ Duplication correcte : {nb1} lignes Ref=1 et {nb2} Ref=2")
        else:
            log_func(f"   ⚠️  Déséquilibre Ref=1({nb1}) vs Ref=2({nb2})")

        # ── CALENDRIER ─────────────────────────────────────────────────────────
        log_func("\n📅 GÉNÉRATION DU FICHIER CALENDRIER_MENAGE...")

        df_sem_cal = df_semaines_ref[df_semaines_ref['Trimestre'] == TRIMESTRE_COLLECTE_DATE].copy()
        if len(df_sem_cal) == 0:
            log_func(f"   ⚠️  ERREUR : Aucune ligne pour '{TRIMESTRE_COLLECTE_DATE}'")
        else:
            log_func(f"   ✓ {len(df_sem_cal)} semaines trouvées")

        df_sem_cal['num_sem'] = (
            df_sem_cal['Numero_semaine'].astype(str).str.extract(r'(\d+)').astype(int)
        )

        lookup_cal = {
            row['num_sem']: (formater_date(row['Date_debut_sem_ref']), formater_date(row['Date_fin_sem_ref']))
            for _, row in df_sem_cal.iterrows()
        }

        corresp_cal = {1:(1,2), 2:(3,4), 3:(5,6), 4:(7,8), 5:(9,10), 6:(11,12), 7:(13,13)}

        def get_dates_cal(row):
            ordre = int(row['Ordre']) if not pd.isna(row['Ordre']) else 0
            if ordre == 0:
                return ("","","","")
            if row['Region'] == 'ABIDJAN':
                d1, f1 = lookup_cal.get(ordre, ("",""))
                return (d1, f1, "", "")
            paire = corresp_cal.get(ordre)
            if paire is None:
                return ("","","","")
            s1, s2 = paire
            d1, f1 = lookup_cal.get(s1, ("",""))
            d2, f2 = lookup_cal.get(s2, ("",""))
            return (d1, f1, d2, f2)

        cal = menage_final.copy()
        nb_av = len(cal)
        cal = cal.drop_duplicates(subset=['Code1'], keep='first')
        log_func(f"   ✓ Déduplication Code1 : {nb_av} → {len(cal)} lignes ({nb_av-len(cal)} doublons)")

        res = cal.apply(get_dates_cal, axis=1)
        cal['Date_debut_sem_1_ref'] = [r[0] for r in res]
        cal['Date_fin_sem_1_ref']   = [r[1] for r in res]
        cal['Date_debut_sem_2_ref'] = [r[2] for r in res]
        cal['Date_fin_sem_2_ref']   = [r[3] for r in res]

        cols_cal = [
            'Region','sp','Cohorte','ord_sem','HH01','HH0','HH2A',
            'HH1','HH2','HH3','HH4','HH6','HH8','HH8A','HH7','HH7B','HH8B',
            'rghab','rgmen','V1MODINTR',
            'trimestreencours','mois_en_cours','annee','Date1','Date2',
            'Date_debut_sem_1_ref','Date_fin_sem_1_ref',
            'Date_debut_sem_2_ref','Date_fin_sem_2_ref',
            'Reference','Code1','_responsible','_quantity','Ordre','cle'
        ]
        cal = cal[[c for c in cols_cal if c in cal.columns]]
        cal.to_excel(os.path.join(DOSSIER_TRAVAIL_RESULTAT, NOM_CALENDRIER), index=False)
        log_func(f"   ✓ {NOM_CALENDRIER} ({len(cal)} lignes)")

        # ── RÉSUMÉ ─────────────────────────────────────────────────────────────
        log_func("\n" + sep)
        log_func("✅ TRAITEMENT TERMINÉ AVEC SUCCÈS !")
        log_func(sep)
        log_func(f"\n📊 RÉSUMÉ")
        log_func(f"   • ZD traitées         : {len(df_echantillon)}")
        log_func(f"   • Affectations créées : {len(df_resultat)}")
        log_func(f"   • Régions couvertes   : {df_resultat['Region'].nunique()}")
        log_func(f"   • Agents mobilisés    : {df_resultat['login'].nunique()}")
        log_func(f"\n📁 FICHIERS GÉNÉRÉS dans : {DOSSIER_TRAVAIL_RESULTAT}")
        log_func(f"   • {NOM_DENOMBREMENT}  ({len(denombrement)} lignes)")
        log_func(f"   • {NOM_MENAGE}        ({len(menage_final)} lignes)")
        log_func(f"   • {NOM_CALENDRIER}    ({len(cal)} lignes)")
        log_func(f"   • {NOM_DENOM_ABI}     ({len(denom_abi)} lignes)")
        log_func(f"   • {NOM_MENAGE_ABI}    ({len(menage_abi)} lignes)")
        log_func(f"   • {NOM_DENOM_HORS}    ({len(denom_hors)} lignes)")
        log_func(f"   • {NOM_MENAGE_HORS}   ({len(men_combine)} lignes)")

        fin_func(True)

    except Exception as e:
        import traceback
        log_func(f"\n❌ ERREUR : {str(e)}")
        log_func(traceback.format_exc())
        fin_func(False)


# ==============================================================================
# INTERFACE GRAPHIQUE TKINTER
# ==============================================================================

class App(tk.Tk):

    COULEUR_FOND    = "#F0F4F8"
    COULEUR_TITRE   = "#1A3A5C"
    COULEUR_SECTION = "#2E6DA4"
    COULEUR_BTN     = "#2E6DA4"
    COULEUR_BTN_OK  = "#1E8C45"
    COULEUR_ALERTE  = "#C0392B"
    POLICE_TITRE    = ("Segoe UI", 13, "bold")
    POLICE_SECTION  = ("Segoe UI", 10, "bold")
    POLICE_NORMALE  = ("Segoe UI", 9)
    POLICE_MONO     = ("Consolas", 9)

    def __init__(self):
        super().__init__()
        self.title("Préparation des fichiers de collecte terrain — Passage 1")
        self.geometry("1050x820")
        self.minsize(900, 700)
        self.configure(bg=self.COULEUR_FOND)
        self.resizable(True, True)

        # Variables Tkinter
        self.var_trimestre       = tk.StringVar(value="1T2026")
        self.var_trimestre_date  = tk.StringVar(value="T1_2026")
        self.var_num_trim        = tk.IntVar(value=1)
        self.var_mois            = tk.IntVar(value=1)
        self.var_annee           = tk.IntVar(value=2026)
        self.var_sous_echant     = tk.IntVar(value=8)
        self.var_nb_menages      = tk.IntVar(value=6)
        self.var_dossier_result  = tk.StringVar()
        self.var_fichier_ech     = tk.StringVar()
        self.var_feuille_ech     = tk.StringVar(value="BASEGLO")
        self.var_fichier_georef  = tk.StringVar()
        self.var_fichier_equipes = tk.StringVar()
        self.var_feuille_equipes = tk.StringVar(value="Equipe")
        self.var_fichier_sem_ref = tk.StringVar()

        self._build_ui()

    # ── Construction de l'interface ──────────────────────────────────────────

    def _build_ui(self):
        # ---- En-tête --------------------------------------------------------
        header = tk.Frame(self, bg=self.COULEUR_TITRE, pady=10)
        header.pack(fill="x")
        tk.Label(header, text="PRÉPARATION DES FICHIERS DE COLLECTE — PASSAGE 1",
                 font=("Segoe UI", 14, "bold"), fg="white", bg=self.COULEUR_TITRE
                 ).pack()
        tk.Label(header,
                 text="⚠️  Cette application prend en charge la collecte jusqu'au T4-2030.\n"
                      "Au-delà, veuillez mettre à jour le fichier FICHIER_SEMAINES_REF "
                      "pour prolonger les trimestres.",
                 font=("Segoe UI", 8), fg="#FFD700", bg=self.COULEUR_TITRE, justify="center"
                 ).pack(pady=(2, 0))

        # ---- Notebook (onglets) ---------------------------------------------
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook.Tab", font=self.POLICE_SECTION,
                        padding=[12, 5], background="#D0DFF0")
        style.map("TNotebook.Tab", background=[("selected", self.COULEUR_SECTION)],
                  foreground=[("selected", "white")])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=8)

        self.tab_params   = tk.Frame(nb, bg=self.COULEUR_FOND)
        self.tab_fichiers = tk.Frame(nb, bg=self.COULEUR_FOND)
        self.tab_aide     = tk.Frame(nb, bg=self.COULEUR_FOND)
        self.tab_logs     = tk.Frame(nb, bg=self.COULEUR_FOND)

        nb.add(self.tab_params,   text="  ① Paramètres  ")
        nb.add(self.tab_fichiers, text="  ② Fichiers  ")
        nb.add(self.tab_aide,     text="  ③ Aide & Mode d'emploi  ")
        nb.add(self.tab_logs,     text="  ④ Exécution & Logs  ")

        self._build_tab_params()
        self._build_tab_fichiers()
        self._build_tab_aide()
        self._build_tab_logs()

    # ── Onglet 1 : Paramètres ────────────────────────────────────────────────

    def _build_tab_params(self):
        p = self.tab_params
        canvas = tk.Canvas(p, bg=self.COULEUR_FOND, highlightthickness=0)
        scroll = ttk.Scrollbar(p, orient="vertical", command=canvas.yview)
        frame  = tk.Frame(canvas, bg=self.COULEUR_FOND)

        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        def section(parent, titre):
            f = tk.LabelFrame(parent, text=f"  {titre}  ",
                              font=self.POLICE_SECTION, fg=self.COULEUR_SECTION,
                              bg=self.COULEUR_FOND, padx=12, pady=8, relief="groove")
            f.pack(fill="x", padx=16, pady=8)
            return f

        def ligne(parent, label, widget_func, aide="", row=0):
            tk.Label(parent, text=label, font=self.POLICE_NORMALE,
                     bg=self.COULEUR_FOND, anchor="w", width=28).grid(
                row=row, column=0, sticky="w", pady=3)
            w = widget_func(parent)
            w.grid(row=row, column=1, sticky="ew", padx=8)
            if aide:
                tk.Label(parent, text=aide, font=("Segoe UI", 8),
                         fg="#666", bg=self.COULEUR_FOND, wraplength=380, justify="left"
                         ).grid(row=row, column=2, sticky="w", padx=4)
            parent.columnconfigure(1, weight=1)
            return w

        # ── Section Trimestre ────────────────────────────────
        s1 = section(frame, "Paramètres du trimestre de collecte")
        ligne(s1, "TRIMESTRE_COLLECTE *",
              lambda p: tk.Entry(p, textvariable=self.var_trimestre, width=14, font=self.POLICE_NORMALE),
              "Format exact : 1T2026  (numéro du trimestre + T + année)", row=0)
        ligne(s1, "TRIMESTRE_COLLECTE_DATE *",
              lambda p: tk.Entry(p, textvariable=self.var_trimestre_date, width=14, font=self.POLICE_NORMALE),
              "Format exact : T1_2026  (T + numéro + _ + année)", row=1)
        ligne(s1, "NUMERO_TRIMESTRE *",
              lambda p: ttk.Spinbox(p, from_=1, to=4, textvariable=self.var_num_trim, width=6),
              "Entier entre 1 et 4  (1=T1, 2=T2, 3=T3, 4=T4)", row=2)
        ligne(s1, "MOIS_COLLECTE *",
              lambda p: ttk.Spinbox(p, from_=1, to=12, textvariable=self.var_mois, width=6),
              "Mois de début du trimestre  (ex: T1=1, T2=4, T3=7, T4=10)", row=3)
        ligne(s1, "ANNEE_COLLECTE *",
              lambda p: ttk.Spinbox(p, from_=2024, to=2035, textvariable=self.var_annee, width=8),
              "Année de la collecte  (ex: 2026)", row=4)

        # ── Section Échantillonnage ──────────────────────────
        s2 = section(frame, "Paramètres d'échantillonnage")

        # Spinbox SOUS_ECHANTILLON
        fr_se = tk.Frame(s2, bg=self.COULEUR_FOND)
        tk.Label(s2, text="SOUS_ECHANTILLON *", font=self.POLICE_NORMALE,
                 bg=self.COULEUR_FOND, anchor="w", width=28).grid(row=0, column=0, sticky="w", pady=3)
        ttk.Spinbox(s2, from_=1, to=8, textvariable=self.var_sous_echant, width=6).grid(
            row=0, column=1, sticky="w", padx=8)
        aide_se = (
            "Entier entre 1 et 8.\n"
            "Correspond au numéro du lot de ZD (segment) à enquêter ce trimestre.\n"
            "• T1-2026 = sous-échantillon 8 (dernier lot des segments 1)\n"
            "• T2-2026 = sous-échantillon 1 (début des segments 2)\n"
            "• T3-2026 = sous-échantillon 2  …et ainsi de suite.\n"
            "HORS-ABIDJAN : 56 ZD/région × 7 ZD/trimestre = 8 trimestres pour épuiser chaque segment.\n"
            "ABIDJAN       : 104 ZD         × 13 ZD/trimestre = 8 trimestres par segment."
        )
        tk.Label(s2, text=aide_se, font=("Segoe UI", 12), fg="#444",
                 bg=self.COULEUR_FOND, wraplength=420, justify="left"
                 ).grid(row=1, column=0, columnspan=3, sticky="w", padx=4, pady=(0,4))
        s2.columnconfigure(1, weight=1)

        # ── Section Ménage ───────────────────────────────────
        s3 = section(frame, "Paramètres des ménages")
        ligne(s3, "NOMBRE_MENAGES_PAR_ZD",
              lambda p: ttk.Spinbox(p, from_=1, to=20, textvariable=self.var_nb_menages, width=6),
              "Nombre de ménages à interroger par ZD (défaut : 6)", row=0)

    # ── Onglet 2 : Fichiers ──────────────────────────────────────────────────

    def _build_tab_fichiers(self):
        p = self.tab_fichiers
        canvas = tk.Canvas(p, bg=self.COULEUR_FOND, highlightthickness=0)
        scroll = ttk.Scrollbar(p, orient="vertical", command=canvas.yview)
        frame  = tk.Frame(canvas, bg=self.COULEUR_FOND)
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        def section(titre):
            f = tk.LabelFrame(frame, text=f"  {titre}  ",
                              font=self.POLICE_SECTION, fg=self.COULEUR_SECTION,
                              bg=self.COULEUR_FOND, padx=12, pady=8, relief="groove")
            f.pack(fill="x", padx=16, pady=8)
            return f

        def row_fichier(parent, label, var, types, aide, row):
            tk.Label(parent, text=label, font=self.POLICE_SECTION,
                     bg=self.COULEUR_FOND, fg=self.COULEUR_TITRE, anchor="w"
                     ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(6,1))
            tk.Label(parent, text=aide, font=("Segoe UI", 8), fg="#555",
                     bg=self.COULEUR_FOND, wraplength=550, justify="left"
                     ).grid(row=row+1, column=0, columnspan=3, sticky="w")
            e = tk.Entry(parent, textvariable=var, font=self.POLICE_MONO, width=62)
            e.grid(row=row+2, column=0, columnspan=2, sticky="ew", pady=2)
            tk.Button(parent, text="📁 Parcourir", font=self.POLICE_NORMALE,
                      bg="#E8EFF8", relief="flat", cursor="hand2",
                      command=lambda v=var, t=types: self._parcourir_fichier(v, t)
                      ).grid(row=row+2, column=2, sticky="w", padx=4)
            parent.columnconfigure(0, weight=1)
            return row+3

        def row_dossier(parent, label, var, aide, row):
            tk.Label(parent, text=label, font=self.POLICE_SECTION,
                     bg=self.COULEUR_FOND, fg=self.COULEUR_TITRE, anchor="w"
                     ).grid(row=row, column=0, columnspan=3, sticky="w", pady=(6,1))
            tk.Label(parent, text=aide, font=("Segoe UI", 8), fg="#555",
                     bg=self.COULEUR_FOND, wraplength=550, justify="left"
                     ).grid(row=row+1, column=0, columnspan=3, sticky="w")
            e = tk.Entry(parent, textvariable=var, font=self.POLICE_MONO, width=62)
            e.grid(row=row+2, column=0, columnspan=2, sticky="ew", pady=2)
            tk.Button(parent, text="📁 Parcourir", font=self.POLICE_NORMALE,
                      bg="#E8EFF8", relief="flat", cursor="hand2",
                      command=lambda v=var: self._parcourir_dossier(v)
                      ).grid(row=row+2, column=2, sticky="w", padx=4)
            parent.columnconfigure(0, weight=1)
            return row+3

        def row_feuille(parent, label, var, var_fichier, aide, row):
            tk.Label(parent, text=label, font=self.POLICE_NORMALE,
                     bg=self.COULEUR_FOND, anchor="w"
                     ).grid(row=row, column=0, sticky="w", pady=(4,1))
            combo = ttk.Combobox(parent, textvariable=var, width=30, font=self.POLICE_NORMALE)
            combo.grid(row=row, column=1, sticky="w", padx=6)
            tk.Button(parent, text="🔄 Charger les feuilles", font=("Segoe UI", 8),
                      bg="#E8EFF8", relief="flat", cursor="hand2",
                      command=lambda c=combo, f=var_fichier: self._charger_feuilles(c, f)
                      ).grid(row=row, column=2, sticky="w", padx=2)
            tk.Label(parent, text=aide, font=("Segoe UI", 8), fg="#666",
                     bg=self.COULEUR_FOND).grid(row=row+1, column=0, columnspan=3, sticky="w")
            return row+2

        XL = [("Fichiers Excel", "*.xlsx *.xls")]

        # Dossier résultat
        s0 = section("Dossier de sortie")
        row_dossier(s0, "DOSSIER_TRAVAIL_RESULTAT",
                    self.var_dossier_result,
                    "Dossier où seront enregistrés les 7 fichiers Excel générés.", 0)

        # Échantillon
        s1 = section("Fichier Échantillon des ZD  (FICHIER_ECHANTILLON)")
        r  = row_fichier(s1, "FICHIER_ECHANTILLON", self.var_fichier_ech, XL,
                         "Fichier Excel contenant l'échantillon complet des ZD à enquêter "
                         "(ex: Echantillon_ZD_VF_ACTUALISEE.xlsx).", 0)
        row_feuille(s1, "FEUILLE_ECHANTILLON", self.var_feuille_ech, self.var_fichier_ech,
                    "Défaut : BASEGLO — Cliquez sur 🔄 après avoir sélectionné le fichier.", r)

        # Géoréférencement
        s2 = section("Fichier Géoréférencement  (FICHIER_GEOREF)")
        row_fichier(s2, "FICHIER_GEOREF", self.var_fichier_georef, XL,
                    "Base géo-référencée des ZD (ex: VF_BASE_ILOT_12012024_VF_work_Geovf.xlsx).", 0)

        # Équipes
        s3 = section("Fichier Équipes  (FICHIER_EQUIPES)")
        r  = row_fichier(s3, "FICHIER_EQUIPES", self.var_fichier_equipes, XL,
                         "Fichier contenant les comptes, mots de passe et régions des agents "
                         "(ex: EquipeParRegionVF.xlsx).", 0)
        row_feuille(s3, "FEUILLE_EQUIPES", self.var_feuille_equipes, self.var_fichier_equipes,
                    "Défaut : Equipe — Cliquez sur 🔄 après avoir sélectionné le fichier.", r)

        # Semaines de référence
        s4 = section("Fichier Semaines de référence  (FICHIER_SEMAINES_REF)")
        row_fichier(s4, "FICHIER_SEMAINES_REF", self.var_fichier_sem_ref, XL,
                    "Fichier des semaines de référence (Semaine_ref.xlsx). "
                    "La feuille utilisée est toujours 'Semaine_ref_trim'.", 0)

    # ── Onglet 3 : Aide ──────────────────────────────────────────────────────

    def _build_tab_aide(self):
        p = self.tab_aide
        txt = scrolledtext.ScrolledText(p, font=("Segoe UI", 12), bg="#FAFCFF",
                                        wrap="word", state="normal", relief="flat")
        txt.pack(fill="both", expand=True, padx=10, pady=8)

        contenu = """
╔══════════════════════════════════════════════════════════════════════════════╗
║          GUIDE D'UTILISATION — PRÉPARATION FICHIERS PASSAGE 1              ║
╚══════════════════════════════════════════════════════════════════════════════╝

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 DESCRIPTION DES PARAMÈTRES (Onglet ① Paramètres)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

▶ TRIMESTRE_COLLECTE
    Format OBLIGATOIRE : {numéro}T{année}   ex: 1T2026, 2T2026, 3T2025
    C'est le code du trimestre en cours de collecte.

▶ TRIMESTRE_COLLECTE_DATE
    Format OBLIGATOIRE : T{numéro}_{année}   ex: T1_2026, T2_2026, T3_2025
    Même information que TRIMESTRE_COLLECTE mais dans un format différent,
    utilisé pour interroger le fichier Semaine_ref.xlsx.
    ⚠️  Les deux paramètres doivent correspondre au même trimestre.

▶ NUMERO_TRIMESTRE
    Entier entre 1 et 4.
    T1 → 1 | T2 → 2 | T3 → 3 | T4 → 4

▶ MOIS_COLLECTE
    Mois de début du trimestre :
    T1 → 1 (Janvier) | T2 → 4 (Avril) | T3 → 7 (Juillet) | T4 → 10 (Octobre)

▶ ANNEE_COLLECTE
    Année de la collecte, ex: 2026.

▶ SOUS_ECHANTILLON  ⚠️  PARAMÈTRE CRITIQUE
    Entier entre 1 et 8. Identifie le lot de ZD enquêté ce trimestre.

    LOGIQUE D'ENCHAINEMENT :
    Chaque ZD possède 6 segments. On visite un segment par lot de 8 trimestres.
    • Trimestres 1–8   → segments 1  des ZD (sous-échantillons 1 à 8)
    • Trimestres 9–16  → segments 2  des ZD (sous-échantillons 1 à 8 de nouveau)
    • Trimestres 17–24 → segments 3  des ZD  … et ainsi de suite.

    EXEMPLE concret autour du T1-2026 :
    ┌─────────────┬──────────────────────┬─────────────────┐
    │  Trimestre  │  SOUS_ECHANTILLON    │  Segment visité │
    ├─────────────┼──────────────────────┼─────────────────┤
    │   T4-2025   │         7            │  Segment 1      │
    │   T1-2026   │         8            │  Segment 1 (fin)│
    │   T2-2026   │         1            │  Segment 2 (deb)│
    │   T3-2026   │         2            │  Segment 2      │
    └─────────────┴──────────────────────┴─────────────────┘

    HORS-ABIDJAN  : 56 ZD/région × 7 ZD/trimestre = 8 trimestres/segment
    ABIDJAN       : 104 ZD       × 13 ZD/trimestre = 8 trimestres/segment

▶ NOMBRE_MENAGES_PAR_ZD
    Nombre de ménages à enquêter dans chaque ZD. Valeur par défaut : 6.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 DESCRIPTION DES FICHIERS (Onglet ② Fichiers)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

▶ DOSSIER_TRAVAIL_RESULTAT
    Dossier où seront enregistrés les 7 fichiers Excel produits.
    Assurez-vous d'avoir les droits d'écriture dans ce dossier.

▶ FICHIER_ECHANTILLON  +  FEUILLE_ECHANTILLON
    Fichier Excel de l'échantillon tiré pour la collecte.
    Contient notamment les colonnes : sous_echant, NUM_ZD_Vf, NomSp, semaine_ref…
    Après avoir sélectionné le fichier, cliquez sur 🔄 pour charger les feuilles
    disponibles et choisir la bonne (défaut : BASEGLO).

▶ FICHIER_GEOREF
    Base géo-référencée des ZD. Contient les coordonnées et codes géographiques
    de chaque Zone de Dénombrement. Pas de sélection de feuille (1ère feuille).

▶ FICHIER_EQUIPES  +  FEUILLE_EQUIPES
    Fichier des comptes agents : logins, mots de passe, régions, équipes.
    Après avoir sélectionné le fichier, cliquez sur 🔄 pour charger les feuilles
    (défaut : Equipe). Seuls les comptes de type "Agent de collecte" sont utilisés.

▶ FICHIER_SEMAINES_REF
    Calendrier des semaines de référence. La feuille utilisée est toujours
    'Semaine_ref_trim'. Ce fichier couvre les trimestres jusqu'au T4-2030.
    ⚠️  Au-delà de 2030, ouvrir ce fichier et ajouter les nouvelles lignes.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 FICHIERS PRODUITS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Le programme génère 7 fichiers Excel dans le dossier de sortie :

  1. Denombrement_{trim}.xlsx            — toutes régions
  2. Denombrement_ABIDJAN_{trim}.xlsx    — Abidjan uniquement
  3. Denombrement_HORS_ABIDJAN_{trim}.xlsx
  4. Menage_{trim}.xlsx                  — toutes régions
  5. Menage_ABIDJAN_{trim}.xlsx
  6. Menage_HORS_ABIDJAN_{trim}.xlsx
  7. Calendrier_menage_{trim}.xlsx       — avec les 4 dates de semaines de référence

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 MODE D'UTILISATION DES FICHIERS RÉSULTATS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📌 À RETENIR POUR L'ENVOI AUX TÉLÉOPÉRATEURS :

  ✓ menage.xlsx
      → La variable _responsible contient déjà les comptes des agents
      → Variables à envoyer : de ord_sem jusqu'à _quantity (inclus)
      → Format d'export    : Texte (séparateur : tabulation) (*.txt)

  ✓ denombrement.xlsx
      → La variable _responsible contient déjà les comptes des agents
      → Variables à envoyer : de ord_sem jusqu'à _quantity (inclus)
      → Format d'export    : Texte (séparateur : tabulation) (*.txt)

  ✓ Calendrier_menage.xlsx
      → Contient les 4 colonnes de dates de semaines de référence :
          • Date_debut_sem_1_ref / Date_fin_sem_1_ref
          • Date_debut_sem_2_ref / Date_fin_sem_2_ref  (vide pour Abidjan)
      → Pour Abidjan  : 1 semaine par ZD (Ordre 1→sem 1, Ordre 13→sem 13)
      → Pour autres   : 2 semaines par ZD (Ordre 1→sem 1+2, …, Ordre 7→sem 13)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 PROCÉDURE PAS À PAS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1. Renseignez tous les paramètres dans l'onglet ① Paramètres
  2. Sélectionnez tous les fichiers d'entrée dans l'onglet ② Fichiers
     (utilisez les boutons 📁 Parcourir et 🔄 Charger les feuilles)
  3. Vérifiez le dossier de sortie
  4. Allez dans l'onglet ④ Exécution & Logs
  5. Cliquez sur ▶ LANCER LE TRAITEMENT
  6. Suivez les logs en temps réel
  7. En cas de succès ✅, récupérez les fichiers dans le dossier de sortie
"""
        txt.insert("1.0", contenu)
        txt.configure(state="disabled")

    # ── Onglet 4 : Logs & Exécution ──────────────────────────────────────────

    def _build_tab_logs(self):
        p = self.tab_logs

        # Barre de boutons
        bar = tk.Frame(p, bg=self.COULEUR_FOND, pady=6)
        bar.pack(fill="x", padx=10)

        self.btn_lancer = tk.Button(
            bar, text="▶  LANCER LE TRAITEMENT",
            font=("Segoe UI", 11, "bold"),
            bg=self.COULEUR_BTN_OK, fg="white",
            padx=20, pady=8, relief="flat", cursor="hand2",
            command=self._lancer
        )
        self.btn_lancer.pack(side="left", padx=(0, 10))

        tk.Button(
            bar, text="🗑  Effacer les logs",
            font=self.POLICE_NORMALE, bg="#D5D5D5", relief="flat", cursor="hand2",
            command=self._effacer_logs
        ).pack(side="left")

        self.lbl_statut = tk.Label(bar, text="", font=("Segoe UI", 9, "bold"),
                                   bg=self.COULEUR_FOND)
        self.lbl_statut.pack(side="left", padx=16)

        # Zone de logs
        self.zone_logs = scrolledtext.ScrolledText(
            p, font=self.POLICE_MONO, bg="#0D1117", fg="#C9D1D9",
            insertbackground="white", relief="flat", wrap="none"
        )
        self.zone_logs.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        self.zone_logs.tag_config("ok",    foreground="#56D364")
        self.zone_logs.tag_config("err",   foreground="#F85149")
        self.zone_logs.tag_config("warn",  foreground="#E3B341")
        self.zone_logs.tag_config("titre", foreground="#79C0FF", font=("Consolas", 9, "bold"))
        self.zone_logs.tag_config("info",  foreground="#C9D1D9")

    # ── Helpers ──────────────────────────────────────────────────────────────

    def _parcourir_fichier(self, var, types):
        f = filedialog.askopenfilename(filetypes=types)
        if f:
            var.set(f)

    def _parcourir_dossier(self, var):
        d = filedialog.askdirectory()
        if d:
            var.set(d)

    def _charger_feuilles(self, combo, var_fichier):
        chemin = var_fichier.get()
        if not chemin or not os.path.exists(chemin):
            messagebox.showwarning("Fichier introuvable",
                                   "Veuillez d'abord sélectionner un fichier Excel valide.")
            return
        try:
            xl = pd.ExcelFile(chemin)
            combo['values'] = xl.sheet_names
            messagebox.showinfo("Feuilles chargées",
                                f"{len(xl.sheet_names)} feuille(s) trouvée(s) :\n" +
                                "\n".join(xl.sheet_names))
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire le fichier :\n{e}")

    def _log(self, msg):
        """Ajoute une ligne dans la zone de logs (thread-safe)."""
        def _write():
            tag = "info"
            if "✅" in msg or "✓" in msg:
                tag = "ok"
            elif "❌" in msg or "⚠️" in msg or "ERREUR" in msg:
                tag = ("err" if "❌" in msg or "ERREUR" in msg else "warn")
            elif msg.startswith("=") or "TERMINÉ" in msg or "COLLECTE" in msg:
                tag = "titre"
            self.zone_logs.insert("end", msg + "\n", tag)
            self.zone_logs.see("end")
        self.after(0, _write)

    def _effacer_logs(self):
        self.zone_logs.delete("1.0", "end")
        self.lbl_statut.config(text="")

    def _valider_params(self):
        """Retourne un dict de paramètres ou None si une erreur est détectée."""
        erreurs = []

        # Vérifications de base
        if not self.var_trimestre.get().strip():
            erreurs.append("TRIMESTRE_COLLECTE est vide.")
        if not self.var_trimestre_date.get().strip():
            erreurs.append("TRIMESTRE_COLLECTE_DATE est vide.")
        if not self.var_dossier_result.get().strip():
            erreurs.append("DOSSIER_TRAVAIL_RESULTAT est vide.")
        if not self.var_fichier_ech.get().strip():
            erreurs.append("FICHIER_ECHANTILLON est vide.")
        if not self.var_fichier_georef.get().strip():
            erreurs.append("FICHIER_GEOREF est vide.")
        if not self.var_fichier_equipes.get().strip():
            erreurs.append("FICHIER_EQUIPES est vide.")
        if not self.var_fichier_sem_ref.get().strip():
            erreurs.append("FICHIER_SEMAINES_REF est vide.")

        # Vérifier existence des fichiers
        for label, var in [
            ("FICHIER_ECHANTILLON",  self.var_fichier_ech),
            ("FICHIER_GEOREF",       self.var_fichier_georef),
            ("FICHIER_EQUIPES",      self.var_fichier_equipes),
            ("FICHIER_SEMAINES_REF", self.var_fichier_sem_ref),
        ]:
            if var.get() and not os.path.exists(var.get()):
                erreurs.append(f"{label} : fichier introuvable.\n  → {var.get()}")

        # Vérifier/créer dossier résultat
        doss = self.var_dossier_result.get().strip()
        if doss and not os.path.exists(doss):
            try:
                os.makedirs(doss, exist_ok=True)
            except Exception as e:
                erreurs.append(f"Impossible de créer DOSSIER_TRAVAIL_RESULTAT : {e}")

        if erreurs:
            messagebox.showerror("Paramètres invalides",
                                 "Veuillez corriger les erreurs suivantes :\n\n" +
                                 "\n\n".join(f"• {e}" for e in erreurs))
            return None

        return {
            'TRIMESTRE_COLLECTE'     : self.var_trimestre.get().strip(),
            'TRIMESTRE_COLLECTE_DATE': self.var_trimestre_date.get().strip(),
            'NUMERO_TRIMESTRE'       : self.var_num_trim.get(),
            'MOIS_COLLECTE'          : self.var_mois.get(),
            'ANNEE_COLLECTE'         : self.var_annee.get(),
            'SOUS_ECHANTILLON'       : self.var_sous_echant.get(),
            'NOMBRE_MENAGES_PAR_ZD'  : self.var_nb_menages.get(),
            'DOSSIER_TRAVAIL_RESULTAT': doss,
            'FICHIER_ECHANTILLON'    : self.var_fichier_ech.get().strip(),
            'FEUILLE_ECHANTILLON'    : self.var_feuille_ech.get().strip() or "BASEGLO",
            'FICHIER_GEOREF'         : self.var_fichier_georef.get().strip(),
            'FICHIER_EQUIPES'        : self.var_fichier_equipes.get().strip(),
            'FEUILLE_EQUIPES'        : self.var_feuille_equipes.get().strip() or "Equipe",
            'FICHIER_SEMAINES_REF'   : self.var_fichier_sem_ref.get().strip(),
        }

    def _lancer(self):
        params = self._valider_params()
        if params is None:
            return

        self.btn_lancer.config(state="disabled", text="⏳ Traitement en cours…")
        self.lbl_statut.config(text="⏳ En cours…", fg="#E3B341")
        self._effacer_logs()

        def fin(succes):
            def _ui():
                self.btn_lancer.config(state="normal", text="▶  LANCER LE TRAITEMENT")
                if succes:
                    self.lbl_statut.config(text="✅ Terminé avec succès !", fg="#56D364")
                    messagebox.showinfo("Succès",
                                        "Le traitement s'est terminé avec succès.\n"
                                        f"Les fichiers sont disponibles dans :\n{params['DOSSIER_TRAVAIL_RESULTAT']}")
                else:
                    self.lbl_statut.config(text="❌ Erreur — voir les logs", fg="#F85149")
                    messagebox.showerror("Erreur",
                                         "Une erreur s'est produite.\nConsultez les logs pour plus de détails.")
            self.after(0, _ui)

        t = threading.Thread(target=lancer_traitement,
                             args=(params, self._log, fin), daemon=True)
        t.start()


# ==============================================================================
# POINT D'ENTRÉE
# ==============================================================================

if __name__ == "__main__":
    app = App()
    app.mainloop()
