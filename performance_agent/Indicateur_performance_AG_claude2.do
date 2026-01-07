
********************************************************************************
* PROGRAMME : Calcul des durées d'administration des questionnaires ENEM
* OBJECTIF  : Calculer les temps moyens d'administration par agent et trimestre
* VERSION   : 3.1 - Correction gestion variables manquantes
********************************************************************************

clear all
set more off
set maxvar 10000
capture log close

********************************************************************************
* 1. CONFIGURATION DES CHEMINS D'ACCÈS
********************************************************************************

* Dossier principal des résultats
global cd_resu "D:\ENEM_Working\Apurement salaire\Dossier_travail_Toure"

* Création des dossiers
capture mkdir "$cd_resu"
capture mkdir "$cd_resu\results"
capture mkdir "$cd_resu\results\par_agent"
capture mkdir "$cd_resu\logs"

* Ouverture du fichier log
log using "$cd_resu\logs\calcul_durees_`c(current_date)'.log", replace

* Dossiers sources par trimestre
global wd_T3_2024 "D:\ENEM_Working\Apurement salaire\Base_brute_T3_2024\"
global wd_T4_2024 "D:\ENEM_Working\Apurement salaire\Base_brute_T4_2024\"
global wd_T1_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T1_2025\"
global wd_T2_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T2_2025\"
global wd_T3_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T3_2025\"
global wd_T4_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T4_2025\"

* Affichage configuration
di as text _newline "=========================================="
di as text "CONFIGURATION DES CHEMINS"
di as text "=========================================="
di as text "Dossier résultats : " as result "$cd_resu"

********************************************************************************
* 2. DÉFINITION DU PROGRAMME PRINCIPAL
********************************************************************************

program define calc_duree_section
    args section
    
    local vstart "hha_`section'"
    local vend "hhavf_`section'"
    local vduree "duree_`section'"
    
    * Vérifier l'existence des variables
    capture confirm variable `vstart'
    local exist_start = (_rc == 0)
    
    capture confirm variable `vend'
    local exist_end = (_rc == 0)
    
    if `exist_start' & `exist_end' {
        * Extraction des heures
        capture drop `vstart'_debut_heure
        gen `vstart'_debut_heure = ""
        quietly replace `vstart'_debut_heure = substr(`vstart', 12, 8) if !missing(`vstart') & `vstart'!="##N/A##"
        
        capture drop `vend'_fin_heure
        gen `vend'_fin_heure = ""
        quietly replace `vend'_fin_heure = substr(`vend', 12, 8) if !missing(`vend') & `vend'!="##N/A##"
        
        * Calcul de la durée
        capture drop `vduree'
        gen `vduree' = .
        quietly replace `vduree' = ///
            abs((clock(`vend'_fin_heure, "hms") - clock(`vstart'_debut_heure, "hms")) / (1000*60)) ///
            if !missing(`vstart'_debut_heure) & !missing(`vend'_fin_heure) & ///
               `vstart'!="##N/A##" & `vend'!="##N/A##" & ///
               `vstart'_debut_heure != "" & `vend'_fin_heure != ""
        
        quietly replace `vduree' = 0 if missing(`vduree')
        
        quietly count if `vduree' > 0
        di as text "  Section `section' : " as result %8.0fc r(N) " observations avec durée > 0"
    }
    else {
        * Variables absentes - créer durée à 0
        capture drop `vduree'
        gen `vduree' = 0
        di as text "  Section `section' : " as error "Variables manquantes - durée = 0"
    }
end

********************************************************************************
* 3. TRAITEMENT PAR TRIMESTRE
********************************************************************************

* Listes
local trimestres "T3_2024 T4_2024 T1_2025 T2_2025 T3_2025 T4_2025"
local fichiers "ENEM_2024T3.dta ENEM_2024T4.dta ENEM_2025T1.dta ENEM_2025T2.dta ENEM_2025T3.dta ENEM_2025T4.dta"
local sections_menage "COMP ED FP IM EM HAND"
local sections_individuel "FT SE EMP CQ PL WKT WKI P RHE TP DI SRH R C"

local compteur = 0

foreach trimestre of local trimestres {
    local compteur = `compteur' + 1
    local fichier : word `compteur' of `fichiers'
    
    di as text _newline(2) "=========================================="
    di as text "TRAITEMENT : `trimestre'"
    di as text "=========================================="
    
    local chemin_base "${wd_`trimestre'}\`fichier'"
    local chemin_membres "${wd_`trimestre'}\membres.dta"
    
    * Vérifications
    capture confirm file "`chemin_base'"
    if _rc {
        di as error "Fichier base introuvable - passage au suivant"
        continue
    }
    
    capture confirm file "`chemin_membres'"
    if _rc {
        di as error "Fichier membres introuvable - passage au suivant"
        continue
    }
    
    * Chargement
    use "`chemin_base'", clear
    di as text "Ménages chargés : " as result _N
    
    * Merge
    merge 1:m interview__key using "`chemin_membres'", ///
        keepusing(membres__id HH14a EP1a M4Confirm ///
                  hha_COMP hhavf_COMP hha_ED hhavf_ED hha_FP hhavf_FP ///
                  hha_IM hhavf_IM hha_EM hhavf_EM hha_HAND hhavf_HAND ///
                  hha_FT hhavf_FT hha_SE hhavf_SE hha_EMP hhavf_EMP ///
                  hha_CQ hhavf_CQ hha_PL hhavf_PL hha_WKT hhavf_WKT ///
                  hha_WKI hhavf_WKI hha_P hhavf_P hha_RHE hhavf_RHE ///
                  hha_TP hhavf_TP hha_DI hhavf_DI hha_SRH hhavf_SRH ///
                  hha_R hhavf_R hha_C hhavf_C HH15A hhaa )
    
    keep if _merge == 3
    drop _merge
    di as text "Individus après merge : " as result _N
    
    * Vérifier HH13_es
    capture confirm variable type_agent	
    capture confirm variable HH13_es
    if _rc {
        di as error "HH13_es ou type_agent absente - création variable par défaut"
        gen HH13_es = .
    }
    
    * ========== QUESTIONNAIRE MÉNAGE ==========
    di as text _newline "Calcul durées - QUESTIONNAIRE MÉNAGE..."
    foreach section of local sections_menage {
        calc_duree_section `section'
    }
    
    * Durée totale ménage
    capture drop duree_QUEST_MEN
    gen duree_QUEST_MEN = 0
    foreach section of local sections_menage {
        * CORRECTION : Utiliser capture pour éviter erreurs si variable manquante
        capture confirm variable duree_`section'
        if _rc == 0 {
            quietly replace duree_QUEST_MEN = duree_QUEST_MEN + duree_`section'
        }
    }
    quietly replace duree_QUEST_MEN = 0 if missing(duree_QUEST_MEN)
    
    * Stats ménage
    qui sum duree_QUEST_MEN if duree_QUEST_MEN > 0
    if r(N) > 0 {
        di as text "  Durée moy. ménage : " as result %6.2f r(mean) " min"
    }
    
    * ========== QUESTIONNAIRE INDIVIDUEL ==========
    di as text _newline "Calcul durées - QUESTIONNAIRE INDIVIDUEL..."
    foreach section of local sections_individuel {
        calc_duree_section `section'
    }
    
    * Durée totale individuel
    capture drop duree_QUEST_IND
    gen duree_QUEST_IND = 0
    foreach section of local sections_individuel {
        * CORRECTION : Utiliser capture
        capture confirm variable duree_`section'
        if _rc == 0 {
            quietly replace duree_QUEST_IND = duree_QUEST_IND + duree_`section'
        }
    }
    quietly replace duree_QUEST_IND = 0 if missing(duree_QUEST_IND)
    
    * Stats individuel
    qui sum duree_QUEST_IND if duree_QUEST_IND > 0
    if r(N) > 0 {
        di as text "  Durée moy. individuel : " as result %6.2f r(mean) " min"
    }
    
    * ========== DURÉE TOTALE ==========
    capture drop duree_QUEST_TOTAL
    gen duree_QUEST_TOTAL = duree_QUEST_MEN + duree_QUEST_IND
    label variable duree_QUEST_TOTAL "Durée totale (ménage + individuel) min"
    
    * ========== DURÉE GLOBALE MENAGE : AUTRE MANIERE DE CALCULER==========
    di as text "Calcul durée globale..."
    capture drop hha_debut_heure HH15A_fin_heure duree_GL_MEN
    
    capture confirm variable hha
    if _rc == 0 {
        gen hha_debut_heure = substr(hha, 12, 8)
    }
    else {
        gen hha_debut_heure = ""
    }
    
    capture confirm variable HH15A
    if _rc == 0 {
        gen HH15A_fin_heure = substr(HH15A, 12, 8)
    }
    else {
        gen HH15A_fin_heure = ""
    }
    
    gen duree_GL_MEN = .
    quietly replace duree_GL_MEN = ///
        abs((clock(HH15A_fin_heure, "hms") - clock(hha_debut_heure, "hms")) / (1000*60)) ///
        if !missing(hha_debut_heure) & !missing(HH15A_fin_heure) & ///
           hha_debut_heure != "" & HH15A_fin_heure != ""
    quietly replace duree_GL_MEN = 0 if missing(duree_GL_MEN)

    * ========== DURÉE GLOBALE INDIVIDUEL : AUTRE MANIERE DE CALCULER==========
    di as text "Calcul durée globale..."
    capture drop hhaa date_fin duree_GL_IND
    
    capture confirm variable hhaa
    if _rc == 0 {
        gen hhaa_debut_heure = substr(hhaa, 12, 8)
    }
    else {
        gen hhaa_debut_heure = ""
    }
    
    capture confirm variable date_fin
    if _rc == 0 {
        gen date_fin_fin_heure = substr(date_fin, 12, 8)
    }
    else {
        gen date_fin_fin_heure = ""
    }
    
    gen duree_GL_IND = .
    quietly replace duree_GL_IND = ///
        abs((clock(date_fin_fin_heure, "hms") - clock(hhaa_debut_heure, "hms")) / (1000*60)) ///
        if !missing(hhaa_debut_heure) & !missing(date_fin_fin_heure) & ///
           hhaa_debut_heure != "" & date_fin_fin_heure != ""
    quietly replace duree_GL_IND = 0 if missing(duree_GL_IND)

    * ========== DURÉE TOTALE AUTRE MANIERE DE CALCULER ==========
    capture drop duree_QUEST_GL_TOTAL
    gen duree_QUEST_GL_TOTAL = duree_GL_MEN + duree_GL_IND
    label variable duree_QUEST_GL_TOTAL "Durée global totale autre forme (ménage + individuel) min"

    
    * Variable trimestre
    gen str10 trimestre = "`trimestre'"
    label variable HH13_es "Agent terrain"
    label variable trimestre "Trimestre"
    
    order trimestre HH13_es interview__key membres__id
    
    * ========== PERFORMANCES PAR AGENT ==========
    di as text _newline "Calcul performances par agent..."
    
    * CORRECTION : Créer un identifiant unique pour le comptage
    * (interview__key est string, donc on ne peut pas faire count dessus)
    * Note : On compte les INDIVIDUS traités par agent (pas les ménages)
    * Car base = niveau individuel (interview__key + membres__id)
    gen id_unique = 1
	
    * Note : On compte le nombre de ménages traités par agent 	
	gen id_men_unique = 1 if membres__id==1
	
	* Calcul Entretien entierement rempli
	cap drop entretient_end
	gen entretient_end =.
	replace entretient_end=1 if HH14a==1
	replace entretient_end=1 if HH14a==. & resultat_enqb==1
	
	
	* Calcul Individu injoignable 
	cap drop entretient_nojoin
	gen entretient_nojoin=.
	replace entretient_nojoin=1 if inlist(HH14a,3,4,5,7,8,10) | inlist(HH14a,11,12,13,14,16,17) 

	replace entretient_nojoin=1 if ( HH14a==9 & HH14_Aut!="" & HH14_Aut!="##N/A##")

	replace entretient_nojoin=1 if HH14a==. & (inlist(resultat_enqb, 3, 4, 7, 8, 11,12) | inlist(resultat_enqb, 13, 14, 15, 16, 17, 18))

	replace entretient_nojoin=1 if (HH14a==. & inlist(resultat_enqb, 9, 19) & resultat_enqb_aut!="" & resultat_enqb_aut!="##N/A##")

	replace entretient_nojoin=1 if (HH14a==9 & inlist(resultat_enqb, 9, 19) & resultat_enqb_aut!="" & resultat_enqb_aut!="##N/A##")
	
	
	* Calcul Individu rempli partiellement
	cap drop entretient_partialend
	gen entretient_partialend =.
	replace entretient_partialend = 1 if HH14a==2
	replace entretient_partialend =1 if HH14a==. & resultat_enqb==2
		
    
    preserve
/*
    collapse (count) nb_questionnaires=id_unique /// Compte nombre individus
					 nb_menages= id_men_unique /// Compte nombre menages
					 interview_end = entretient_end ///
					 interview_nojoin = entretient_nojoin ///
					 interview_partialend = entretient_partialend ///
             (mean) moy_duree_MEN=duree_QUEST_MEN ///
                    moy_duree_IND=duree_QUEST_IND ///
                    moy_duree_TOTAL=duree_QUEST_TOTAL ///
                    moy_duree_GL_MEN=duree_GL_MEN ///
					moy_duree_GL_IND=duree_GL_IND ///
					moy_duree_GL_TOTAL=duree_QUEST_GL_TOTAL ///
             (sd) sd_duree_MEN=duree_GL_MEN ///
                  sd_duree_IND=duree_GL_IND ///
             (min) min_duree_MEN=duree_GL_MEN ///
                   min_duree_IND=duree_GL_IND ///
             (p50) med_duree_MEN=duree_GL_MEN ///
                   med_duree_IND=duree_GL_IND ///
             (max) max_duree_MEN=duree_GL_MEN ///
                   max_duree_IND=duree_GL_IND, ///
             by(HH13_es type_agent)
*/
			 
    collapse (count) nb_questionnaires=id_unique /// Compte nombre individus
					 nb_menages= id_men_unique /// Compte nombre menages
					 interview_end = entretient_end ///
					 interview_nojoin = entretient_nojoin ///
					 interview_partialend = entretient_partialend ///
             (mean) moy_duree_GL_MEN=duree_GL_MEN ///
					moy_duree_GL_IND=duree_GL_IND ///
					moy_duree_GL_TOTAL=duree_QUEST_GL_TOTAL ///
             (sd) sd_duree_MEN=duree_GL_MEN ///
                  sd_duree_IND=duree_GL_IND ///
             (min) min_duree_MEN=duree_GL_MEN ///
                   min_duree_IND=duree_GL_IND ///
             (p50) med_duree_MEN=duree_GL_MEN ///
                   med_duree_IND=duree_GL_IND ///
             (max) max_duree_MEN=duree_GL_MEN ///
                   max_duree_IND=duree_GL_IND, ///
             by(HH13_es type_agent)
			 
			 
    gen str10 trimestre = "`trimestre'"
    order trimestre HH13_es
    
    * Labels
    label variable nb_questionnaires "Nb individus traités"
    label variable nb_menages "Nb menages traités"	
    label variable interview_end "Nb interview achevés"
    label variable interview_nojoin "Nb interview injoignable"
    label variable interview_partialend "Nb interview achevés partiellement"
*    label variable moy_duree_MEN "Durée moy. ménage (min)"
*    label variable moy_duree_IND "Durée moy. individuel (min)"
*    label variable moy_duree_TOTAL "Durée moy. totale (min)"

    label variable moy_duree_GL_MEN "Autre forme: Durée moy. ménage (min)"
    label variable moy_duree_GL_IND "Autre forme: Durée moy. individuel (min)"
    label variable moy_duree_GL_TOTAL "Autre forme: Durée moy. totale (min)"
    
	
    format moy_* sd_* min_* med_* max_* %9.2f
    
    * Affichage
    di as text _newline "Résumé par agent :"
*    list HH13_es nb_questionnaires moy_duree_MEN moy_duree_IND moy_duree_TOTAL, ///

    list HH13_es type_agent  nb_questionnaires nb_menages interview_end interview_nojoin interview_partialend moy_duree_GL_MEN moy_duree_GL_IND moy_duree_GL_TOTAL, ///
        separator(0) abbreviate(15)
    
    * Sauvegarde
    save "$cd_resu\results\par_agent\performance_agents_`trimestre'.dta", replace
    export excel using "$cd_resu\results\par_agent\performance_agents_`trimestre'.xlsx", ///
        firstrow(varlabels) replace
    
    restore

    * ========== PERFORMANCES PAR REGION ==========
    di as text _newline "Calcul performances par région..."
	
	
    preserve
			 
    collapse (count) nb_questionnaires=id_unique /// Compte nombre individus
					 nb_menages= id_men_unique /// Compte nombre menages
					 interview_end = entretient_end ///
					 interview_nojoin = entretient_nojoin ///
					 interview_partialend = entretient_partialend ///
             (mean) moy_duree_GL_MEN=duree_GL_MEN ///
					moy_duree_GL_IND=duree_GL_IND ///
					moy_duree_GL_TOTAL=duree_QUEST_GL_TOTAL ///
             (sd) sd_duree_MEN=duree_GL_MEN ///
                  sd_duree_IND=duree_GL_IND ///
             (min) min_duree_MEN=duree_GL_MEN ///
                   min_duree_IND=duree_GL_IND ///
             (p50) med_duree_MEN=duree_GL_MEN ///
                   med_duree_IND=duree_GL_IND ///
             (max) max_duree_MEN=duree_GL_MEN ///
                   max_duree_IND=duree_GL_IND, ///
             by(HH2 type_agent)
			 
			 
    gen str10 trimestre = "`trimestre'"
    order trimestre HH2
    
    * Labels
    label variable nb_questionnaires "Nb individus traités"
    label variable nb_menages "Nb menages traités"	
    label variable interview_end "Nb interview achevés"
    label variable interview_nojoin "Nb interview injoignable"
    label variable interview_partialend "Nb interview achevés partiellement"
*    label variable moy_duree_MEN "Durée moy. ménage (min)"
*    label variable moy_duree_IND "Durée moy. individuel (min)"
*    label variable moy_duree_TOTAL "Durée moy. totale (min)"

    label variable moy_duree_GL_MEN "Autre forme: Durée moy. ménage (min)"
    label variable moy_duree_GL_IND "Autre forme: Durée moy. individuel (min)"
    label variable moy_duree_GL_TOTAL "Autre forme: Durée moy. totale (min)"
    
	
    format moy_* sd_* min_* med_* max_* %9.2f
    
    * Affichage
    di as text _newline "Résumé par région :"
*    list HH13_es nb_questionnaires moy_duree_MEN moy_duree_IND moy_duree_TOTAL, ///

    list HH2 type_agent  nb_questionnaires nb_menages interview_end interview_nojoin interview_partialend moy_duree_GL_MEN moy_duree_GL_IND moy_duree_GL_TOTAL, ///
        separator(0) abbreviate(15)
    
    * Sauvegarde
    save "$cd_resu\results\par_agent\performance_regions_`trimestre'.dta", replace
    export excel using "$cd_resu\results\par_agent\performance_regions_`trimestre'.xlsx", ///
        firstrow(varlabels) replace
    
    restore
	
	
    
    * ========== SAUVEGARDE RÉSULTATS INDIVIDUELS ==========
    local fichier_sortie "$cd_resu\results\individus_`trimestre'"
    
    save "`fichier_sortie'.dta", replace
    di as text _newline "  Fichier sauvegardé : " as result "`fichier_sortie'.dta"
    
    export excel using "`fichier_sortie'.xlsx", firstrow(variables) replace
    export delimited using "`fichier_sortie'.csv", replace delimiter(",")
    
    * ========== BASE CONSOLIDÉE ==========
    if `compteur' == 1 {
        save "$cd_resu\results\base_consolidee.dta", replace
    }
    else {
        append using "$cd_resu\results\base_consolidee.dta"
        save "$cd_resu\results\base_consolidee.dta", replace
    }
}

********************************************************************************
* 4. CONSOLIDATION PERFORMANCES PAR AGENT
********************************************************************************

di as text _newline(2) "=========================================="
di as text "CONSOLIDATION PERFORMANCES"
di as text "=========================================="

clear
local first = 1
foreach trimestre of local trimestres {
    capture confirm file "$cd_resu\results\par_agent\performance_agents_`trimestre'.dta"
    if _rc == 0 {
        if `first' {
            use "$cd_resu\results\par_agent\performance_agents_`trimestre'.dta", clear
            local first = 0
        }
        else {
            append using "$cd_resu\results\par_agent\performance_agents_`trimestre'.dta"
        }
    }
}

capture confirm variable HH13_es
if _rc == 0 {
    save "$cd_resu\results\par_agent\performance_agents_tous_trimestres.dta", replace
    export excel using "$cd_resu\results\par_agent\performance_agents_tous_trimestres.xlsx", ///
        firstrow(varlabels) replace
    
    di as text "Performance globale par agent :"
    
    preserve
/*	
    collapse (sum) nb_total=nb_questionnaires ///
             (mean) moy_globale_MEN=moy_duree_MEN ///
                    moy_globale_IND=moy_duree_IND ///
                    moy_globale_TOTAL=moy_duree_TOTAL, ///
             by(HH13_es)
*/
    collapse (sum) nb_total=nb_questionnaires ///
             (mean) moy_globale_MEN=moy_duree_GL_MEN ///
                    moy_globale_IND=moy_duree_GL_IND ///
                    moy_globale_TOTAL=moy_duree_GL_TOTAL, ///
             by(HH13_es)
			 
			 
    format moy_* %9.2f
    gsort -nb_total
    
    list HH13_es nb_total moy_globale_MEN moy_globale_IND moy_globale_TOTAL, separator(0)

    
    export excel using "$cd_resu\results\par_agent\synthese_globale_agents.xlsx", ///
        firstrow(varlabels) replace
    restore
}

********************************************************************************
* 5. RÉSULTATS FINAUX (BASE INDIVIDUELLE)
********************************************************************************

di as text _newline(2) "=========================================="
di as text "PRODUCTION RÉSULTATS FINAUX"
di as text "=========================================="

use "$cd_resu\results\base_consolidee.dta", clear

* CORRECTION : Conserver seulement les variables qui existent
*local vars_to_keep "interview__key membres__id trimestre HH13_es duree_QUEST_MEN duree_QUEST_IND duree_QUEST_TOTAL duree_GL_MEN duree_GL_IND duree_QUEST_GL_TOTAL"

local vars_to_keep "interview__key membres__id trimestre HH13_es type_agent id_unique id_men_unique duree_QUEST_MEN duree_QUEST_IND duree_QUEST_TOTAL duree_GL_MEN duree_GL_IND duree_QUEST_GL_TOTAL entretient_end entretient_nojoin entretient_partialend"

* Ajouter les durées de section si elles existent
foreach section in COMP ED FP IM EM HAND FT SE EMP CQ PL WKT WKI P RHE TP DI SRH R C {
    capture confirm variable duree_`section'
    if _rc == 0 {
        local vars_to_keep "`vars_to_keep' duree_`section'"
    }
}

keep `vars_to_keep'

* Labels
label variable trimestre "Trimestre de collecte"
label variable HH13_es "Agent terrain"
label variable duree_QUEST_MEN "Durée totale ménage (min)"
label variable duree_QUEST_IND "Durée totale individuel (min)"
label variable duree_QUEST_TOTAL "Durée totale (min)"
label variable duree_GL_MEN "Durée globale MEN (min)"
label variable duree_GL_IND "Durée globale IND (min)"
label variable entretient_end "Interview achevés"
label variable entretient_nojoin "Interview injoignable"
label variable entretient_partialend "Interview achevés partiellement"

* Stats globales
di as text _newline "Statistiques par trimestre :"
/*
tabstat duree_QUEST_MEN duree_QUEST_IND duree_QUEST_TOTAL, ///
    by(trimestre) stat(n mean sd p50) format(%9.2f)
*/

tabstat duree_GL_MEN duree_GL_IND duree_QUEST_GL_TOTAL, ///
    by(trimestre) stat(n mean sd p50) format(%9.2f)
	
	
* Sauvegarde finale
save "$cd_resu\results\base_finale_durees.dta", replace
export excel using "$cd_resu\results\base_finale_durees.xlsx", firstrow(varlabels) replace
export delimited using "$cd_resu\results\base_finale_durees.csv", replace delimiter(",")

di as text _newline "Fichiers finaux créés :"
di as result "  - base_finale_durees.dta/.xlsx/.csv"

********************************************************************************
* 6. RAPPORTS SYNTHÉTIQUES
********************************************************************************

di as text _newline "=========================================="
di as text "RAPPORTS SYNTHÉTIQUES"
di as text "=========================================="

* Rapport par trimestre
preserve
/*
collapse (count) n=duree_QUEST_MEN ///
         (mean) moy_quest_men=duree_QUEST_MEN ///
                moy_quest_ind=duree_QUEST_IND ///
                moy_quest_total=duree_QUEST_TOTAL ///
         (sd) sd_quest_men=duree_QUEST_MEN ///
              sd_quest_ind=duree_QUEST_IND ///
         (p50) med_quest_men=duree_QUEST_MEN ///
               med_quest_ind=duree_QUEST_IND, ///
         by(trimestre)
*/
		 
	collapse (count) n=id_unique /// Compte nombre d'entretiens individuels
					 n_menage=id_men_unique /// Compte nombre d'entretiens menages
         (mean) moy_quest_men=duree_GL_MEN ///
                moy_quest_ind=duree_GL_IND ///
                moy_quest_total=duree_QUEST_GL_TOTAL ///
         (sd) sd_quest_men=duree_GL_MEN ///
              sd_quest_ind=duree_GL_IND ///
         (p50) med_quest_men=duree_GL_MEN ///
               med_quest_ind=duree_GL_IND, ///
         by(trimestre)
	 
		 
format n* moy_* sd_* med_* %9.2f

di as text "Synthèse par trimestre :"
list, separator(0)

export excel using "$cd_resu\results\rapport_synthese_trimestre.xlsx", ///
    firstrow(variables) replace
restore

* Rapport comparaison
preserve
*collapse (mean) duree_QUEST_MEN duree_QUEST_IND duree_QUEST_TOTAL, by(trimestre)

collapse (mean) duree_GL_MEN duree_GL_IND duree_QUEST_GL_TOTAL, by(trimestre)


format duree_* %9.2f

di as text "Comparaison ménage vs individuel :"
list, separator(0)

export excel using "$cd_resu\results\rapport_comparaison_MEN_IND.xlsx", ///
    firstrow(variables) replace
restore

di as text _newline "=========================================="
di as text "TRAITEMENT TERMINÉ"
di as text "=========================================="
di as result "Résultats dans : $cd_resu\results"

log close