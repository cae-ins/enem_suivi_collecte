********************************************************************************
* PROGRAMME : Automatisation du suivi de la collecte ENEM
* OBJECTIF  : Générer automatiquement les fichiers de contrôle qualité
*             pour le suivi de la collecte trimestrielle
*
* VÉRIFICATIONS EFFECTUÉES :
*   1. Âges manquants (terrain et téléopérateurs)
*   2. Sexes manquants (terrain et téléopérateurs)
*   3. Taille des ménages par ZD (> 12 ménages)
*   4. Positions GPS des ménages
*   5. Individus non classables (emploi et chômage)
*   6. Extraction codification emploi (principal, secondaire, pluriactivité)
*
* AUTEUR    : KOUAME KOUASSI GUY MARTIAL
* DATE      : 29 décembre 2025
* VERSION   : 2.1 - Ajout extraction codification emploi
********************************************************************************

clear all
set more off
capture log close

********************************************************************************
* 🔧 SECTION 1 : CONFIGURATION DES PARAMÈTRES
********************************************************************************

* ===== TRIMESTRE À TRAITER =====
* Modifier uniquement cette ligne pour changer le trimestre analysé
global TRIMESTRE_ACTUEL "T3_2025"

* ===== CHEMINS DES DOSSIERS SOURCES =====
global wd_T2_2024 "D:\ENEM_Working\Apurement salaire\Base_brute_T2_2024\"
global wd_T3_2024 "D:\ENEM_Working\Apurement salaire\Base_brute_T3_2024\"
global wd_T4_2024 "D:\ENEM_Working\Apurement salaire\Base_brute_T4_2024\"
global wd_T1_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T1_2025\"
global wd_T2_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T2_2025\"
global wd_T3_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T3_2025\"
global wd_T4_2025 "D:\ENEM_Working\Apurement salaire\Base_brute_T4_2025\"

* ===== DOSSIER DE SORTIE =====
global DOSSIER_SORTIE "D:\ENEM_Working\Activite_quotidienne_Trimestre_${TRIMESTRE_ACTUEL}"
global Point_suivi_collecte "${DOSSIER_SORTIE}\Resultat\Point_suivi_collecte\"

* ===== CORRESPONDANCE TRIMESTRE -> NOM DE FICHIER MÉNAGE =====
if "$TRIMESTRE_ACTUEL" == "T2_2024" global NOM_FICHIER "ENEM_2024T2.dta"
if "$TRIMESTRE_ACTUEL" == "T3_2024" global NOM_FICHIER "ENEM_2024T3.dta"
if "$TRIMESTRE_ACTUEL" == "T4_2024" global NOM_FICHIER "ENEM_2024T4.dta"
if "$TRIMESTRE_ACTUEL" == "T1_2025" global NOM_FICHIER "ENEM_2025T1.dta"
if "$TRIMESTRE_ACTUEL" == "T2_2025" global NOM_FICHIER "ENEM_2025T2.dta"
if "$TRIMESTRE_ACTUEL" == "T3_2025" global NOM_FICHIER "ENEM_2025T3.dta"
if "$TRIMESTRE_ACTUEL" == "T4_2025" global NOM_FICHIER "ENEM_2025T4.dta"

* ===== CORRESPONDANCE TRIMESTRE -> NOM DE FICHIER PLURIACTIVITÉ =====
if "$TRIMESTRE_ACTUEL" == "T2_2024" global NOM_FICHIER_PLURIACTIVITE "r_activite_s.dta"
if "$TRIMESTRE_ACTUEL" == "T3_2024" global NOM_FICHIER_PLURIACTIVITE "r_activite_s.dta"
if "$TRIMESTRE_ACTUEL" == "T4_2024" global NOM_FICHIER_PLURIACTIVITE "r_activite_s.dta"
if "$TRIMESTRE_ACTUEL" == "T1_2025" global NOM_FICHIER_PLURIACTIVITE "r_activite_s.dta"
if "$TRIMESTRE_ACTUEL" == "T2_2025" global NOM_FICHIER_PLURIACTIVITE "r_activite_s.dta"
if "$TRIMESTRE_ACTUEL" == "T3_2025" global NOM_FICHIER_PLURIACTIVITE "r_activite_s.dta"
if "$TRIMESTRE_ACTUEL" == "T4_2025" global NOM_FICHIER_PLURIACTIVITE "r_activite_s.dta"

********************************************************************************
* 📊 SECTION 2 : AFFICHAGE DE LA CONFIGURATION
********************************************************************************

di as text _newline(2) "=========================================="
di as text "SUIVI DE LA COLLECTE ENEM"
di as text "=========================================="
di as text "Trimestre analysé : " as result "$TRIMESTRE_ACTUEL"
di as text "Fichier ménage    : " as result "$NOM_FICHIER"
di as text "Fichier pluriact. : " as result "$NOM_FICHIER_PLURIACTIVITE"
di as text "Dossier source    : " as result "${wd_${TRIMESTRE_ACTUEL}}"
di as text "Dossier sortie    : " as result "$Point_suivi_collecte"
di as text "=========================================="

********************************************************************************
* 📂 SECTION 3 : CRÉATION DES DOSSIERS DE SORTIE
********************************************************************************

di as text _newline "📂 Création des dossiers de sortie..."

capture mkdir "$DOSSIER_SORTIE"
capture mkdir "${DOSSIER_SORTIE}\Resultat"
capture mkdir "$Point_suivi_collecte"

di as result "   ✓ Dossiers créés"

********************************************************************************
* 📥 SECTION 4 : CHARGEMENT ET FUSION DES DONNÉES
********************************************************************************

di as text _newline "📥 Chargement des données du trimestre ${TRIMESTRE_ACTUEL}..."

* Vérifier l'existence du fichier ménage
local chemin_menage "${wd_${TRIMESTRE_ACTUEL}}\$NOM_FICHIER"
capture confirm file "`chemin_menage'"
if _rc {
    di as error "❌ ERREUR : Fichier ménage introuvable"
    di as error "   Chemin : `chemin_menage'"
    exit 601
}

* Vérifier l'existence du fichier membres
local chemin_membres "${wd_${TRIMESTRE_ACTUEL}}\membres.dta"
capture confirm file "`chemin_membres'"
if _rc {
    di as error "❌ ERREUR : Fichier membres introuvable"
    di as error "   Chemin : `chemin_membres'"
    exit 601
}

* Charger la base ménage
use "`chemin_menage'", clear
local nb_menages = _N
di as text "   ✓ Base ménage chargée : " as result `nb_menages' " ménages"

* Fusionner avec la base membres (1:m)
merge 1:m interview__key using "`chemin_membres'"

* Conserver uniquement les observations appariées
keep if _merge == 3
drop _merge

local nb_individus = _N
di as text "   ✓ Fusion complétée : " as result `nb_individus' " individus"

* Vérifier les variables essentielles
local vars_requises "interview__key AgeAnnee M5 Statut_Res rgmen HH2 HH12 HH13"
foreach var of local vars_requises {
    capture confirm variable `var'
    if _rc {
        di as error "   ⚠️  ATTENTION : Variable `var' non trouvée"
    }
}

********************************************************************************
* 🔍 SECTION 5 : VÉRIFICATION DES ÂGES MANQUANTS
********************************************************************************

di as text _newline(2) "=========================================="
di as text "🔍 VÉRIFICATION 1 : ÂGES MANQUANTS"
di as text "=========================================="

* ----- AGENTS TERRAIN (rgmen == 1) -----
di as text _newline "📍 Agents terrain (rgmen == 1)..."

quietly count if missing(AgeAnnee) & Statut_Res == 1 & rgmen == 1
local nb_age_missing_terrain = r(N)

di as text "   Individus sans âge : " as result `nb_age_missing_terrain'

if `nb_age_missing_terrain' > 0 {
    export excel interview__key membres__id M0__0 M6_A M6_J M6_M M7 AgeAnnee HH2 HH12 HH13 M5 rgmen ///
        using "$Point_suivi_collecte\AGE_missing_terrain_${TRIMESTRE_ACTUEL}.xlsx" ///
        if missing(AgeAnnee) & Statut_Res == 1 & rgmen == 1, ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : AGE_missing_terrain_${TRIMESTRE_ACTUEL}.xlsx"
}
else {
    di as result "   ✓ Aucun âge manquant"
}

* ----- TÉLÉOPÉRATEURS (rgmen != 1) -----
di as text _newline "📞 Téléopérateurs (rgmen != 1)..."

quietly count if missing(AgeAnnee) & Statut_Res == 1 & rgmen != 1
local nb_age_missing_tele = r(N)

di as text "   Individus sans âge : " as result `nb_age_missing_tele'

if `nb_age_missing_tele' > 0 {
    export excel interview__key membres__id M0__0 M6_A M6_J M6_M M7 AgeAnnee HH2 HH12 HH13 M5 rgmen ///
        using "$Point_suivi_collecte\AGE_missing_tele_${TRIMESTRE_ACTUEL}.xlsx" ///
        if missing(AgeAnnee) & Statut_Res == 1 & rgmen != 1, ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : AGE_missing_tele_${TRIMESTRE_ACTUEL}.xlsx"
}
else {
    di as result "   ✓ Aucun âge manquant"
}

********************************************************************************
* 🔍 SECTION 6 : VÉRIFICATION DES SEXES MANQUANTS
********************************************************************************

di as text _newline(2) "=========================================="
di as text "🔍 VÉRIFICATION 2 : SEXES MANQUANTS"
di as text "=========================================="

* ----- AGENTS TERRAIN (rgmen == 1) -----
di as text _newline "📍 Agents terrain (rgmen == 1)..."

quietly count if missing(M5) & Statut_Res == 1 & rgmen == 1
local nb_sexe_missing_terrain = r(N)

di as text "   Individus sans sexe : " as result `nb_sexe_missing_terrain'

if `nb_sexe_missing_terrain' > 0 {
    export excel interview__key HH2 HH12 HH13 M5 rgmen ///
        using "$Point_suivi_collecte\SEXE_missing_terrain_${TRIMESTRE_ACTUEL}.xlsx" ///
        if missing(M5) & Statut_Res == 1 & rgmen == 1, ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : SEXE_missing_terrain_${TRIMESTRE_ACTUEL}.xlsx"
}
else {
    di as result "   ✓ Aucun sexe manquant"
}

* ----- TÉLÉOPÉRATEURS (rgmen != 1) -----
di as text _newline "📞 Téléopérateurs (rgmen != 1)..."

quietly count if missing(M5) & Statut_Res == 1 & rgmen != 1
local nb_sexe_missing_tele = r(N)

di as text "   Individus sans sexe : " as result `nb_sexe_missing_tele'

if `nb_sexe_missing_tele' > 0 {
    export excel interview__key HH2 HH12 HH13 M5 rgmen ///
        using "$Point_suivi_collecte\SEXE_missing_tele_${TRIMESTRE_ACTUEL}.xlsx" ///
        if missing(M5) & Statut_Res == 1 & rgmen != 1, ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : SEXE_missing_tele_${TRIMESTRE_ACTUEL}.xlsx"
}
else {
    di as result "   ✓ Aucun sexe manquant"
}

********************************************************************************
* 🔍 SECTION 7 : VÉRIFICATION DE LA TAILLE DES MÉNAGES PAR ZD
********************************************************************************

di as text _newline(2) "=========================================="
di as text "🔍 VÉRIFICATION 3 : TAILLE DES ZD (> 12 MÉNAGES)"
di as text "=========================================="

preserve

* Charger la base ménage
use "`chemin_menage'", clear
local nb_menages = _N
di as text "   ✓ Base ménage chargée : " as result `nb_menages' " ménages"

* Filtrer pour ne garder que les ménages terrain
keep if rgmen == 1

di as text "   Nombre individus en passage 1 : " as result _N

di as text _newline "📊 Calcul du nombre de ménages par ZD..."

* Créer un indicateur par ménage
gen zd = 1

* Agréger par région et ZD
collapse (sum) nb_menages=zd if rgmen==1, by(HH2 HH8A HH8B HH8)

* Trier par nombre de ménages décroissant
gsort -nb_menages

di as text "   Total de ZD : " as result _N

* Filtrer les ZD avec 12 ménages ou plus
quietly count if nb_menages >= 12 & !missing(nb_menages)
local nb_zd_probleme = r(N)

di as text "   ZD avec ≥ 12 ménages : " as result `nb_zd_probleme'

if `nb_zd_probleme' > 0 {
    * Afficher les ZD problématiques
    di as text _newline "   ⚠️  ZD problématiques :"
    list HH2 HH8 HH8A HH8B nb_menages if nb_menages >= 12, ///
        separator(0) abbreviate(20)
    
    * Exporter vers Excel
    export excel HH2 HH8 HH8A HH8B nb_menages ///
        using "$Point_suivi_collecte\ZD_taille_superieure_12_${TRIMESTRE_ACTUEL}.xlsx" ///
        if nb_menages >= 12 & !missing(nb_menages), ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : ZD_taille_superieure_12_${TRIMESTRE_ACTUEL}.xlsx"
}
else {
    di as result "   ✓ Aucune ZD avec plus de 12 ménages"
}

restore

********************************************************************************
* 🔍 SECTION 8 : VÉRIFICATION DES POSITIONS GPS
********************************************************************************

di as text _newline(2) "=========================================="
di as text "🔍 VÉRIFICATION 4 : POSITIONS GPS DES MÉNAGES"
di as text "=========================================="

preserve

keep if rgmen==1
di as text _newline "🌍 Préparation des données GPS..."

* Vérifier l'existence des variables GPS
capture confirm variable GPS__Latitude GPS__Longitude
if _rc {
    di as error "   ⚠️  Variables GPS non trouvées, export ignoré"
    restore
}
else {
    * Convertir les coordonnées GPS en string
    tostring GPS__Latitude, generate(gps__Latitude_str) force
    tostring GPS__Longitude, generate(gps__Longitude_str) force
    tostring GPS__Accuracy, generate(gps__Accuracy_str) force
    tostring GPS__Altitude, generate(gps__Altitude_str) force
    
    * Compter le nombre de ménages avec GPS
    quietly count if !missing(GPS__Latitude) & !missing(GPS__Longitude)
    local nb_gps = r(N)
    
    di as text "   Ménages avec coordonnées GPS : " as result `nb_gps'
    
    * Exporter les données GPS
    export delimited interview__key ///
        gps__Latitude_str gps__Longitude_str gps__Accuracy_str gps__Altitude_str ///
        GPS__Timestamp ///
        HH01 HH0 HH2A HH1 HH2 HH3 HH4 HH6 HH8 HH8A HH7 HH7B HH8B ///
        rghab rgmen V1MODINTR trimestreencours mois_en_cours annee ///
        Date1 Date2 Reference HH12 nom_CE HH13 nom_agent HH9 HH9_1 ///
        using "$Point_suivi_collecte\GPS_menage_ZD_${TRIMESTRE_ACTUEL}.csv", ///
        delimiter(";") datafmt replace
    
    di as result "   ✓ Fichier exporté : GPS_menage_ZD_${TRIMESTRE_ACTUEL}.csv"
    di as text "      (Format CSV délimité par ';' pour géomaticiens)"
    
    restore
}

********************************************************************************
* 🔍 SECTION 9 : VÉRIFICATION DES INDIVIDUS NON CLASSABLES
********************************************************************************

di as text _newline(2) "=========================================="
di as text "🔍 VÉRIFICATION 5 : INDIVIDUS NON CLASSABLES"
di as text "=========================================="

preserve

* ----- Critère : Personnes en Âge de Travailler (PAT) -----
di as text _newline "📋 Identification des PAT (≥ 15 ans)..."

capture drop PAT
gen PAT = (M4Confirm >= 15 & !missing(M4Confirm))

quietly count if PAT == 1
di as text "   PAT identifiées : " as result r(N)

* ===== VÉRIFICATION 5A : CLASSEMENT EMPLOI =====
di as text _newline "💼 Classement EMPLOI..."

capture drop classement_emploi
gen classement_emploi = 1

* Un individu est non classable si TOUTES les variables emploi sont vides
replace classement_emploi = 0 if ///
    missing(SE1) & missing(SE2) & missing(SE3) & missing(SE4) & ///
    missing(SE5) & missing(SE7) & missing(SE8) & missing(SE9) & ///
    missing(SE9A) & missing(SE9B) & missing(SE10) & missing(SE11)

label define classement_emploi_lbl 1 "Classable" 0 "Inclassable"
label values classement_emploi classement_emploi_lbl

* Compter les non classables
quietly count if classement_emploi == 0 & Statut_Res == 1
local nb_inclassable_emploi = r(N)

di as text "   Individus non classables (emploi) : " as result `nb_inclassable_emploi'

if `nb_inclassable_emploi' > 0 {
    * Afficher un échantillon
    di as text _newline "   Échantillon (10 premiers) :"
    list interview__key membres__id M0__0 HH2 HH12 HH13 rgmen classement_emploi ///
        if classement_emploi == 0 & Statut_Res == 1 ///
        in 1/10, separator(0) abbreviate(15)
    
    * Exporter vers Excel
    export excel interview__key membres__id M0__0 HH2 HH12 HH13 classement_emploi rgmen ///
        using "$Point_suivi_collecte\classement_emploi_${TRIMESTRE_ACTUEL}.xlsx" ///
        if classement_emploi == 0 & Statut_Res == 1, ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : classement_emploi_${TRIMESTRE_ACTUEL}.xlsx"
}
else {
    di as result "   ✓ Tous les individus sont classables (emploi)"
}

* ===== VÉRIFICATION 5B : CLASSEMENT CHÔMAGE =====
di as text _newline "📉 Classement CHÔMAGE..."

capture drop classement_chomage
gen classement_chomage = 1

* Un individu est non classable si TOUTES les variables chômage sont vides
replace classement_chomage = 0 if ///
    missing(SRH1) & missing(SRH2) & missing(SRH2A) & missing(SRH11)

label define classement_chomage_lbl 1 "Classable" 0 "Inclassable"
label values classement_chomage classement_chomage_lbl

* Compter les non classables
quietly count if classement_chomage == 0 & Statut_Res == 1
local nb_inclassable_chomage = r(N)

di as text "   Individus non classables (chômage) : " as result `nb_inclassable_chomage'

if `nb_inclassable_chomage' > 0 {
    * Afficher un échantillon
    di as text _newline "   Échantillon (10 premiers) :"
    list interview__key membres__id M0__0 HH2 HH12 HH13 rgmen classement_chomage ///
        if classement_chomage == 0 & Statut_Res == 1 ///
        in 1/10, separator(0) abbreviate(15)
    
    * Exporter vers Excel
    export excel interview__key membres__id M0__0 HH2 HH12 HH13 classement_chomage rgmen ///
        using "$Point_suivi_collecte\classement_chomage_${TRIMESTRE_ACTUEL}.xlsx" ///
        if classement_chomage == 0 & Statut_Res == 1, ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : classement_chomage_${TRIMESTRE_ACTUEL}.xlsx"
}
else {
    di as result "   ✓ Tous les individus sont classables (chômage)"
}

restore

********************************************************************************
* 📋 SECTION 10 : EXTRACTION CODIFICATION EMPLOI
********************************************************************************

di as text _newline(2) "=========================================="
di as text "📋 EXTRACTION CODIFICATION EMPLOI"
di as text "=========================================="

* ===== EXTRACTION 6A : AUTRE EMPLOI (PLURIACTIVITÉ) =====
di as text _newline "🔄 Extraction autre emploi (pluriactivité)..."

* Vérifier l'existence du fichier pluriactivité
local chemin_pluriactivite "${wd_${TRIMESTRE_ACTUEL}}\$NOM_FICHIER_PLURIACTIVITE"
capture confirm file "`chemin_pluriactivite'"

if _rc {
    di as error "   ⚠️  Fichier pluriactivité non trouvé, extraction ignorée"
    di as error "   Chemin : `chemin_pluriactivite'"
}
else {
    preserve
    
    use "`chemin_pluriactivite'", clear
    local nb_obs_pluri = _N
    
    di as text "   Observations chargées : " as result `nb_obs_pluri'
    
    export excel interview__key interview__id membres__id r_activite_s__id ///
        PL3_B PL3_C PL3_D PL3_D1a PL3_E PL3_F PL3_F1a PL3_G PL3_G_Aut PL3_H ///
        using "$Point_suivi_collecte\codification_autre_emploi_${TRIMESTRE_ACTUEL}.xlsx", ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : codification_autre_emploi_${TRIMESTRE_ACTUEL}.xlsx"
    
    restore
}

* ===== EXTRACTION 6B : EMPLOI PRINCIPAL =====
di as text _newline "💼 Extraction emploi principal..."

preserve

* Charger la base ménage
use "`chemin_menage'", clear

* Fusionner avec la base membres
merge 1:m interview__key using "`chemin_membres'"
keep if _merge == 3
drop _merge

* Compter les individus en emploi
quietly count if EN_EMP == 1 & Statut_Res == 1
local nb_emp_principal = r(N)

di as text "   Individus en emploi : " as result `nb_emp_principal'

if `nb_emp_principal' > 0 {
    export excel interview__key interview__id membres__id V1interviewkey V1interviewkey1er ///
        rgmen rghab Statut_Res EF3 EF4 EF4_3 M4Confirm ///
        EP1a EP1b EP1b1 EP2a1 EP2b EP2b1 EP2c EP2d EP2e EP3 EP13 EP13a EP2jb ///
        HH01 HH0 HH2A HH1 HH2 HH3 HH4 HH6 HH8 HH8A HH7 HH7B HH8B HH12 HH13 ///
        using "$Point_suivi_collecte\codification_emploi_principal_${TRIMESTRE_ACTUEL}.xlsx" ///
        if EN_EMP == 1 & Statut_Res == 1, ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : codification_emploi_principal_${TRIMESTRE_ACTUEL}.xlsx"
}
else {
    di as text "   ℹ️  Aucun individu en emploi trouvé"
}

restore

* ===== EXTRACTION 6C : EMPLOI SECONDAIRE =====
di as text _newline "🔧 Extraction emploi secondaire..."

preserve

* Charger la base ménage
use "`chemin_menage'", clear

* Fusionner avec la base membres
merge 1:m interview__key using "`chemin_membres'"
keep if _merge == 3
drop _merge

* Compter les individus en emploi (pour emploi secondaire)
quietly count if EN_EMP == 1 & Statut_Res == 1
local nb_emp_secondaire = r(N)

di as text "   Individus en emploi : " as result `nb_emp_secondaire'

if `nb_emp_secondaire' > 0 {
    export excel interview__key interview__id membres__id V1interviewkey V1interviewkey1er ///
        rgmen rghab Statut_Res EF3 EF4 EF4_3 M4Confirm ///
        EP1a EP1b EP1b1 EP2a1 EP2b EP2b1 EP2c EP2d EP2e EP3 EP13 EP13a EP2jb ///
        ES1a ES1b ES1b1 ES2a ES2b ES2b1 ES2c ES2d ES2e ES2g ES3 ES4 ///
        HH01 HH0 HH2A HH1 HH2 HH3 HH4 HH6 HH8 HH8A HH7 HH7B HH8B HH12 HH13 ///
        using "$Point_suivi_collecte\codification_emploi_secondaire_${TRIMESTRE_ACTUEL}.xlsx" ///
        if EN_EMP == 1 & Statut_Res == 1, ///
        firstrow(varl) replace
    
    di as result "   ✓ Fichier exporté : codification_emploi_secondaire_${TRIMESTRE_ACTUEL}.xlsx"
}
else {
    di as text "   ℹ️  Aucun individu en emploi trouvé"
}

restore

********************************************************************************
* 📊 SECTION 11 : RAPPORT FINAL
********************************************************************************

di as text _newline(2) "=========================================="
di as text "📊 RÉSUMÉ DU SUIVI DE COLLECTE"
di as text "=========================================="
di as text _newline "Trimestre : " as result "$TRIMESTRE_ACTUEL"
di as text _newline "📁 Fichiers générés dans :"
di as result "   $Point_suivi_collecte"
di as text _newline "📋 Vérifications effectuées :"
di as text "   1. Âges manquants (terrain et téléopérateurs)"
di as text "   2. Sexes manquants (terrain et téléopérateurs)"
di as text "   3. Taille des ZD (> 12 ménages)"
di as text "   4. Positions GPS (export pour géomaticiens)"
di as text "   5. Individus non classables (emploi et chômage)"
di as text "   6. Codification emploi (principal, secondaire, pluriactivité)"

di as text _newline "=========================================="
di