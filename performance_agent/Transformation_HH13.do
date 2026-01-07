*******************************************************
* process_trimestres.do
* Parcourt la liste des trimestres, applique les
* transformations sur HH13 -> HH13_es et sauvegarde.
* ATTENTION : Vérifier les globals de chemins avant exécution.
*******************************************************

clear all
set more off

* -----------------------------
* 1) Définir les dossiers sources
* -----------------------------
global wd_T3_2024 "D:\ENEM_Working\Apurement salaire\Base_brute_T3_2024\"
global wd_T4_2024 "D:\ENEM_Working\Apurement salaire\Base_brute_T4_2024\"
global wd_T1_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T1_2025\"
global wd_T2_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T2_2025\"
global wd_T3_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T3_2025\"
global wd_T4_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T4_2025\"

* Définir dossier de sortie unique
global out_main "D:\ENEM_Working\Apurement salaire\Output\"

* Créer le dossier de sortie s'il n'existe pas
capture mkdir "$out_main"

* -----------------------------
* 2) Liste des trimestres et fichiers (TEST avec T3_2025 uniquement)
* -----------------------------
*local trimestres "T3_2025"
*local fichiers   "ENEM_2025T3.dta"

* Pour traiter tous les trimestres, décommenter ci-dessous :
 local trimestres "T3_2024 T4_2024 T1_2025 T2_2025 T3_2025 T4_2025"
 local fichiers   "ENEM_2024T3.dta ENEM_2024T4.dta ENEM_2025T1.dta ENEM_2025T2.dta ENEM_2025T3.dta ENEM_2025T4.dta"

* -----------------------------
* 3) Boucle de traitement
* -----------------------------
local n = wordcount("`trimestres'")
di as txt "=============================================="
di as txt "Nombre de trimestres à traiter : `n'"
di as txt "=============================================="

forvalues i = 1/`n' {

    * CORRECTION : Ajouter les backticks
    local tr : word `i' of `trimestres'
    local file : word `i' of `fichiers'

    di as txt ""
    di as txt "--------------------------------------------------"
    di as txt "Traitement [`i'/`n'] : Trimestre = `tr'"
    di as txt "Fichier : `file'"
    
    * CORRECTION : Récupérer le chemin directement depuis le global
    local wd "${wd_`tr'}"
    
    di as txt "Dossier source : `wd'"
    di as txt "Dossier sortie : $out_main"

    * Vérifier que le fichier source existe
    capture confirm file "`wd'\`file'"
    if _rc {
        di as error ">>> ERREUR : Fichier introuvable : `wd'\`file'"
        di as error ">>> Passage au trimestre suivant..."
        continue
    }
    
    di as txt ">>> Fichier trouvé, chargement..."

    * Charger la base
    use "`wd'\`file'", clear
    
    di as txt ">>> Base chargée : " _N " observations"

    * -----------------------------
    * Transformations sur HH13
    * -----------------------------
    
    * Vérifier que la variable HH13 existe
    capture confirm variable HH13
    if _rc {
        di as error ">>> ERREUR : Variable HH13 introuvable dans `file'"
        continue
    }
    
    di as txt ">>> Application des transformations..."

    * Cloner la variable (si existe déjà, on ignore)
    capture drop HH13_es
    clonevar HH13_es = HH13

    * Remplacer les modalités ciblées par 99999
    replace HH13_es = 999 if inlist(HH13, ///
        109, 111, 112, 113, 114, 119, 120, ///
		116, 117, 118, 126, 127, 134,      ///
        121, 122, 123, 124, 125, 128, 129, 130, ///
        131, 132, 133, 135, 136, 137, 138, ///
        139, 140, 141, 142, 143, 144, 145, ///
        146, 147, 148)

    * Coder les valeurs manquantes en 99998
    replace HH13_es = 9998 if missing(HH13)

    * Définir le label de valeur
*    capture label drop HH13_es_lbl
*    label define HH13_es_lbl 99999 "Bad selection" 99998 "Non rempli", add
*	 label define HH13  99998 "Non rempli", add
    label values HH13_es HH13
*    label values HH13_es HH13_es_lbl
	
    * Contrôles succincts
    quietly count if HH13_es == 99999
    local nbad = r(N)
    quietly count if HH13_es == 99998
    local nmiss = r(N)
    
    di as txt ">>> `nbad' observations recodées en 'Bad selection' (99999)"
    di as txt ">>> `nmiss' observations codées en 'Non rempli' (99998)"

    * Tabulation de contrôle
    di as txt ">>> Distribution de HH13_es :"
    tabulate HH13_es, missing

    * -----------------------------
    * Creation variable Type agent a 2 modalités : 1.Agent terrain 2.Agent teleoperateur
    * -----------------------------
	
	cap drop type_agent
	gen type_agent = 1 if inlist(HH12, 1, 2, 3 , 4, 5, 6,7) | inlist(HH12, 8, 9, 10, 11, 12, 13, 14) | inlist(HH12, 15, 16, 17, 18, 19, 20, 21) | inlist(HH12, 22, 23, 24, 25, 26, 27, 28) | inlist(HH12, 29, 30, 31, 32, 33) 
	replace type_agent=2 if inlist(HH12, 34,35)

    capture label drop type_agent	
    label define type_agent 1 "Agent terrain" 2 "Agent téléopérateur"
    label values type_agent type_agent	
	
    * -----------------------------
    * Sauvegarde
    * -----------------------------
    local outfile = subinstr("`file'", ".dta", "_mod.dta", .)
*    local fullpath "${out_main}\`outfile'" 
    local fullpath "`wd'\`file'"  /// Je decide replacer dans la base definitive
    
    di as txt ">>> Sauvegarde vers : `fullpath'"
    
    save "`fullpath'", replace
    
    * Vérifier que le fichier a bien été créé
    capture confirm file "`fullpath'"
    if _rc == 0 {
        di as result ">>> SUCCESS : Fichier sauvegardé avec succès !"
        di as result ">>> Chemin complet : `fullpath'"
    }
    else {
        di as error ">>> ERREUR : La sauvegarde a échoué !"
    }

    * Nettoyer pour le prochain tour
    clear
}

di as txt ""
di as txt "=============================================="
di as txt "Traitement terminé pour tous les trimestres."
di as txt "=============================================="
di as txt "Fichiers de sortie dans : $out_main"