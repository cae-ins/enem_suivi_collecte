********************************************************************************
* SCRIPT DE MISE À JOUR : Compatibilité Base_Travail_BT_vf.dta
* Date : 05/02/2026
* Objectif : Renommer automatiquement toutes les variables de minuscules → MAJUSCULES
* Résultat : Création de Base_Travail_BT_vf_RENAMED.dta compatible avec votre code
********************************************************************************

clear all
set more off
set maxvar 10000

********************************************************************************
* CONFIGURATION DES CHEMINS
********************************************************************************

global donnees = "D:\ENEM_Working\Article\Analyses_Thematiques_ENE\Analyses_Thematiques_ENE"

di as text _newline(2) "=========================================="
di as text "SCRIPT DE RENOMMAGE AUTOMATIQUE"
di as text "=========================================="
di as text "Base source : Base_Travail_BT_vf.dta"
di as text "Base cible  : Base_Travail_BT_vf_RENAMED.dta"
di as text "=========================================="

********************************************************************************
* ÉTAPE 1 : Charger la nouvelle base
********************************************************************************

use "$donnees\Base_Travail_BT_vf.dta", clear

di as text _newline "✓ Base chargée : " as result _N " observations"

********************************************************************************
* ÉTAPE 2 : Liste exhaustive des variables extraites de votre code
********************************************************************************

* Variables extraites automatiquement de votre code original
* Classées par catégorie pour meilleure lisibilité

* A. Variables principales d'emploi et transport
* Je retire MO dans la liste emploi_transport
local emploi_transport "EP2g1 EP2h EP2i EP2e EP2i_unit EP2g"

* B. Variables de classification socio-économique
*local classification "CISE_18_informel_Emp CISE_18_new branche1"
local classification "CISE_18_informel_Emp CISE_18_new"

* C. Variables géographiques et administratives
local geo_admin "HH2 HH6"

* D. Variables temporelles
*local temporel "trimestre"

* E. Variables de pondération
*local ponderation "pmencor_ind_annuel"

* F. Variables socio-démographiques
*local sociodem "age sexe milieu_resid2 groupe_age4 Niv_inst_AG3"
local sociodem "Niv_inst_AG3"

* G. Variables de résultats d'enquête
*local resultats "resultat_enqb HH14a HH14 HH14_Aut resultat_enqb_aut rgmen"
*local resultats "HH14a HH14 HH14_Aut"


* H. Concaténer toutes les variables
*local all_vars_to_rename "`emploi_transport' `classification' `geo_admin' `temporel' `ponderation' `sociodem' `resultats'"
local all_vars_to_rename "`emploi_transport' `classification' `geo_admin'  `sociodem' `resultats'"


********************************************************************************
* ÉTAPE 3 : Renommage automatique avec rapport détaillé
********************************************************************************

di as text _newline(2) "=========================================="
di as text "RENOMMAGE EN COURS..."
di as text "=========================================="

local count_renamed 0
local count_already_ok 0
local count_missing 0

foreach var of local all_vars_to_rename {
    * Créer le nom en minuscule
    local var_lower = lower("`var'")
    
    * Vérifier si la variable en minuscule existe
    capture confirm variable `var_lower'
    if _rc == 0 {
        * Si oui, la renommer en majuscule
        rename `var_lower' `var'
        local count_renamed = `count_renamed' + 1
        di as result "  ✓ Renommé : " as text "`var_lower'" as result " → " as text "`var'"
    }
    else {
        * Vérifier si la variable existe déjà en majuscule
        capture confirm variable `var'
        if _rc == 0 {
            local count_already_ok = `count_already_ok' + 1
            di as text "  ○ Déjà OK : `var'"
        }
        else {
            local count_missing = `count_missing' + 1
            di as error "  ✗ Absente : `var' / `var_lower'"
        }
    }
}

********************************************************************************
* ÉTAPE 4 : Traitement des variables restantes (tout en MAJUSCULE par sécurité)
********************************************************************************

di as text _newline(2) "=========================================="
di as text "TRAITEMENT DES AUTRES VARIABLES"
di as text "=========================================="

* Lister toutes les variables actuelles
quietly describe, varlist
local remaining_vars `r(varlist)'

local count_other 0

foreach v of local remaining_vars {
    * Ne traiter que les variables encore en minuscules
    if regexm("`v'", "^[a-z]") {
        local v_upper = upper("`v'")
        
        * Vérifier que le nom en majuscule n'existe pas déjà
        capture confirm variable `v_upper'
        if _rc != 0 {
            rename `v' `v_upper'
            local count_other = `count_other' + 1
            di as text "  → Renommé (autre) : `v' → `v_upper'"
        }
    }
}

********************************************************************************
* ÉTAPE 5 : Sauvegarde de la base corrigée
********************************************************************************

di as text _newline(2) "=========================================="
di as text "SAUVEGARDE DE LA BASE"
di as text "=========================================="

save "$donnees\Base_Travail_BT_vf_RENAMED.dta", replace

di as result "✓ Base sauvegardée avec succès !"
di as text "  Fichier : " as result "Base_Travail_BT_vf_RENAMED.dta"

********************************************************************************
* ÉTAPE 6 : Rapport final de compatibilité
********************************************************************************

di as text _newline(2) "=========================================="
di as text "RAPPORT FINAL DE COMPATIBILITÉ"
di as text "=========================================="

di as text _newline "Variables de votre code :"
di as result "  ✓ Renommées      : " as text %3.0f `count_renamed'
di as text    "  ○ Déjà correctes : " %3.0f `count_already_ok'
di as error   "  ✗ Absentes       : " as text %3.0f `count_missing'

di as text _newline "Autres variables de la base :"
di as result "  → Renommées      : " as text %3.0f `count_other'

* Statistiques globales
quietly describe, varlist
local total_vars : word count `r(varlist)'

local count_upper 0
local count_lower 0

foreach var of varlist _all {
    if regexm("`var'", "[A-Z]") {
        local count_upper = `count_upper' + 1
    }
    else {
        local count_lower = `count_lower' + 1
    }
}

di as text _newline "État final de la base :"
di as result "  Variables MAJUSCULES : " as text %4.0f `count_upper'
di as text    "  Variables minuscules : " %4.0f `count_lower'
di as text    "  Total des variables  : " %4.0f `total_vars'

********************************************************************************
* ÉTAPE 7 : Vérification des variables critiques
********************************************************************************

di as text _newline(2) "=========================================="
di as text "VÉRIFICATION DES VARIABLES CRITIQUES"
di as text "=========================================="

* Liste des variables absolument nécessaires pour votre analyse
local critical_vars "EP2g1 EP2h EP2i MO age sexe trimestre pmencor_ind_annuel"


local all_critical_ok = 1

foreach var of local critical_vars {
    capture confirm variable `var'
    if _rc == 0 {
        di as result "  ✓ `var' : OK"
    }
    else {
        di as error "  ✗ `var' : MANQUANTE !"
        local all_critical_ok = 0
    }
}

********************************************************************************
* ÉTAPE 8 : Message final et instructions
********************************************************************************

di as text _newline(2) "=========================================="
di as text "RÉSULTAT FINAL"
di as text "=========================================="

if `all_critical_ok' == 1 & `count_missing' == 0 {
    di as result _newline "✓✓✓ SUCCÈS COMPLET ✓✓✓"
    di as text _newline "Toutes les variables ont été renommées avec succès."
    di as text "Votre code d'analyse fonctionnera parfaitement."
}
else if `all_critical_ok' == 1 {
    di as result _newline "✓ SUCCÈS (avec avertissements)"
    di as text _newline "Les variables critiques sont présentes."
    di as error "Quelques variables secondaires sont absentes (voir ci-dessus)."
    di as text "Votre code devrait fonctionner, mais vérifiez les parties optionnelles."
}
else {
    di as error _newline "✗ ATTENTION : Variables critiques manquantes !"
    di as text "Vérifiez la base source avant de continuer."
}

di as text _newline(2) "=========================================="
di as text "PROCHAINE ÉTAPE"
di as text "=========================================="
di as text "Modifiez UNE SEULE ligne dans votre code d'analyse :"
di as text _newline "  AVANT :"
*di as error "  use " as text quote as error `"$donnees\master_data_indiv_men.dta"' as text quote as error ", replace"
di as text _newline "  AVANT :"
di as text `"  use "$donnees\master_data_indiv_men.dta", replace"'

di as text _newline "  APRÈS :"
*di as result "  use " as text quote as result `"$donnees\Base_Travail_BT_vf_RENAMED.dta"' as text quote as result ", replace"

di as result `"  use "$donnees\Base_Travail_BT_vf_RENAMED.dta", replace"'

di as text _newline "Ensuite, lancez votre code normalement !"
di as text "=========================================="