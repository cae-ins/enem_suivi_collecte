
use "$donnees_BT\Base_Travail_BT_vf_RENAMED_plus_salaire_EP.dta", replace

*===============================================================================
* CALCUL DE LA PRESSION TEMPORELLE DE LA MOBILITÉ
* Calcul l’indice de Pression Temporelle de la Mobilité (IPTM)
* Pression temporelle = Somme(EP2h) / Somme(temps_tra_EP_semaine)
* Pour les personnes en emploi uniquement (EN_EMP==1)
*===============================================================================


* b. Répartition des travailleurs selon la durée du trajet

cap drop cat_workers
gen cat_workers =1 if 15>EP2h & MO==1

replace cat_workers =2  if EP2h>=15 & 30>EP2h  & MO==1

replace cat_workers =3  if EP2h>=30 & 60>EP2h  & MO==1

replace cat_workers =4  if EP2h>=60 & .>EP2h  & MO==1

lab define cat_workers_lbl 1 "<15 min" 2 " 15–30 min" 3 "30–59 min" 4 "≥60 min"
lab values cat_workers cat_workers_lbl 


cap drop cat_workers_dist
gen cat_workers_dist =1 if 5>EP2i & MO==1

replace cat_workers_dist =2  if EP2i>=5 & 10>EP2i  & MO==1

replace cat_workers_dist =3  if EP2i>=10 & .>EP2i  & MO==1

*replace cat_workers_dist =4  if EP2i>=10 & .>EP2i  & MO==1

lab define cat_workers_dist_lbl 1 "<5 Km" 2 " 5–10 Km" 3 "≥10 Km"
lab values cat_workers_dist cat_workers_dist_lbl 



*clear matrix : Pression temporelle de la mobilité
mat define RESU1 = (.)

* Conditions de base pour toutes les analyses
local condition "EN_EMP==1 & EP2h<9998 & temps_tra_EP_semaine!=. & temps_tra_EP_semaine>0"

***** Le temps de trajet EP2h est en minutes et l'heures travaillées temps_tra_EP_semaine est en heures. PAr consequent il faut ramener le Le temps de trajet EP2h en heure en divisant par 60.

*Ensuite le EP2h est le temps de trajet jounalier, donc faut le mutiplier par 2 pour estimer le temsp (aller-retour) et mutiplier  par 5 pour estimer le temps de trajet de la semaine (alller-retour). Ce qui serait conforme à temps_tra_EP_semaine

*** Donc le coeficient multiplicatif de EP2h est : 2*5/60 = 10/60 = 1/6
*-------------------------------------------------------------------------------
* Milieu de résidence
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/3 {
    qui summ EP2h [aw=pmencor_ind_annuel] if `condition' & milieu_resid2==`i'
    local somme_ep2h = r(sum)*(1/6)
    
    qui summ temps_tra_EP_semaine [aw=pmencor_ind_annuel] if `condition' & milieu_resid2==`i'
    local somme_temps = r(sum)
    
    local pression = `somme_ep2h' / `somme_temps'
    mat RESU1 = RESU1 \ (`pression')*100
}

*-------------------------------------------------------------------------------
* Sexe
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/2 {
    qui summ EP2h [aw=pmencor_ind_annuel] if `condition' & sexe==`i'
    local somme_ep2h = r(sum)*(1/6)
    
    qui summ temps_tra_EP_semaine [aw=pmencor_ind_annuel] if `condition'    &  sexe==`i'
    local somme_temps = r(sum)
    
    local pression = `somme_ep2h' / `somme_temps'
    mat RESU1 = RESU1 \ (`pression')*100
}

*-------------------------------------------------------------------------------
* Groupe d'Âge
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/4 {
    qui summ EP2h [aw=pmencor_ind_annuel] if `condition' & groupe_age4==`i'
    local somme_ep2h = r(sum)*(1/6)
    
    qui summ temps_tra_EP_semaine [aw=pmencor_ind_annuel] if `condition' & groupe_age4==`i'
    local somme_temps = r(sum)
    
    local pression = `somme_ep2h' / `somme_temps'
    mat RESU1 = RESU1 \ (`pression')*100
}

*-------------------------------------------------------------------------------
* Niveau d'Instruction
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/5 {
    qui summ EP2h [aw=pmencor_ind_annuel] if `condition' & Niv_inst_AG3==`i'
    local somme_ep2h = r(sum)*(1/6)
    
    qui summ temps_tra_EP_semaine [aw=pmencor_ind_annuel] if `condition' & Niv_inst_AG3==`i'
    local somme_temps = r(sum)
    
    local pression = `somme_ep2h' / `somme_temps'
    mat RESU1 = RESU1 \ (`pression')*100
}

*-------------------------------------------------------------------------------
* Statut de l'emploi
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/5 {
    qui summ EP2h [aw=pmencor_ind_annuel] if `condition' & CISE_18_new==`i'
    local somme_ep2h = r(sum)*(1/6)
    
    qui summ temps_tra_EP_semaine [aw=pmencor_ind_annuel] if `condition' & CISE_18_new==`i'
    local somme_temps = r(sum)
    
    local pression = `somme_ep2h' / `somme_temps'
    mat RESU1 = RESU1 \ (`pression')*100
}

*-------------------------------------------------------------------------------
* Intervalle de temps de parcour
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/4 {
    qui summ EP2h [aw=pmencor_ind_annuel] if `condition' & cat_workers==`i'
    local somme_ep2h = r(sum)*(1/6)
    
    qui summ temps_tra_EP_semaine [aw=pmencor_ind_annuel] if `condition' & cat_workers==`i'
    local somme_temps = r(sum)
    
    local pression = `somme_ep2h' / `somme_temps'
    mat RESU1 = RESU1 \ (`pression')*100
}


*-------------------------------------------------------------------------------
* Secteur d'activité
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/4 {
    qui summ EP2h [aw=pmencor_ind_annuel] if `condition' & branche1==`i'
    local somme_ep2h = r(sum)*(1/6)
    
    qui summ temps_tra_EP_semaine [aw=pmencor_ind_annuel] if `condition' & branche1==`i'
    local somme_temps = r(sum)
    
    local pression = `somme_ep2h' / `somme_temps'
    mat RESU1 = RESU1 \ (`pression')*100
}


*-------------------------------------------------------------------------------
* Ensemble
*-------------------------------------------------------------------------------
qui summ EP2h [aw=pmencor_ind_annuel] if `condition'
local somme_ep2h = r(sum)*(1/6)

qui summ temps_tra_EP_semaine [aw=pmencor_ind_annuel] if `condition'
local somme_temps = r(sum)

local pression = `somme_ep2h' / `somme_temps'
mat RESU1 = RESU1 \ (`pression')*100

*-------------------------------------------------------------------------------
* A.2 Mise en forme du Tableau
*-------------------------------------------------------------------------------

/* Définition des entêtes de lignes et colonnes */
*Lignes 
matrix rownames RESU1 = "Milieu de Residence" "." "Abidjan" "Autre Urbain" "Rural" ///
    "Sexe" "Masculin" "Feminin" ///
    "Groupe d'Age" "16-24 ans" "25-35 ans" "36-64 ans" "65 ans et plus" ///
    "Niveau d'Instruction" "Aucun" "Primaire" "Secondaire 1er Cycle" ///
    "Secondaire 2nd Cycle" "Superieure" ///
    "Statut de l'emploi" "Employeur" "Worker indépendants unemployees" ///
    "Non-salariés (Entrep) dépendants" "Employés" "Travailleurs familiaux" ///
	"Temps de trajet" "Moins de 15 minutes" "15–30 minutes" "30–60 minutes" "Plus de 60 minutes" ///
	"Secteur d'activité" "Agriculture"  "Industrie"  "Commerce" "Autres services"  ///
    "Ensemble"

*Colonnes 
matrix colnames RESU1 = "PRESSION TEMPORELLE"

/* Exportation sur Excel dans le dossier Resultats_Tab*/
putexcel set "${Resultats_Tab2}\Them9_implication_mobilite_travail_surEmploi.xlsx", ///
    sheet("Pression_temporelle") modify

/* Mise en forme */
putexcel B4 = matrix(RESU1), colnames nformat(number_d2)
putexcel A5 = matrix(RESU1), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Pression temporelle de la mobilité selon les caractéristiques socio-démographiques"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Pression temporelle (Temps de trajet / Heures travaillées)"
putexcel B3, bold

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Démographiques"
putexcel (A3:A4), merge bold

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close

di as result "Tableau de pression temporelle créé avec succès!"

*===============================================================================
* CALCUL DU Revenu horaire
* Revenu horaire = Somme(revenu_principal_corrige) / Somme(temps_tra_EP_mensuel)
* Pour les personnes en emploi uniquement (EN_EMP==1)
*===============================================================================

cap drop temps_tra_EP_mensuel
gen temps_tra_EP_mensuel = temps_tra_EP_semaine*4 if EN_EMP==1


*clear matrix : Revenu horaire de la mobilité
mat define RESU1 = (.)

* Conditions de base pour toutes les analyses
local condition "EN_EMP==1 & temps_tra_EP_mensuel!=. & temps_tra_EP_mensuel>0"


*-------------------------------------------------------------------------------
* Milieu de résidence
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/3 {
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & milieu_resid2==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition' & milieu_resid2==`i'
    local somme_temps_tra_EP_mensuel = r(sum)
    
    local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
    mat RESU1 = RESU1 \ (`revenu_horaire')
}

*-------------------------------------------------------------------------------
* Sexe
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/2 {
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & sexe==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition' & sexe==`i'
    local somme_temps_tra_EP_mensuel = r(sum)
    
    local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
    mat RESU1 = RESU1 \ (`revenu_horaire')
}

*-------------------------------------------------------------------------------
* Groupe d'Âge
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/4 {
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & groupe_age4==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition' & groupe_age4==`i'
    local somme_temps_tra_EP_mensuel = r(sum)
    
    local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
    mat RESU1 = RESU1 \ (`revenu_horaire')
}

*-------------------------------------------------------------------------------
* Niveau d'Instruction
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/5 {
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & Niv_inst_AG3==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition' & Niv_inst_AG3==`i'
    local somme_temps_tra_EP_mensuel = r(sum)
    
    local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
    mat RESU1 = RESU1 \ (`revenu_horaire')
}

*-------------------------------------------------------------------------------
* Statut de l'emploi
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/5 {
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & CISE_18_new==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition' & CISE_18_new==`i'
    local somme_temps_tra_EP_mensuel = r(sum)
    
    local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
    mat RESU1 = RESU1 \ (`revenu_horaire')
}

*-------------------------------------------------------------------------------
* Intervalle de temps de parcour
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/4 {
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & cat_workers==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition' & cat_workers==`i'
    local somme_temps_tra_EP_mensuel = r(sum)
    
    local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
    mat RESU1 = RESU1 \ (`revenu_horaire')
}


*-------------------------------------------------------------------------------
* Intervalle de temps de parcour
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/3 {
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & cat_workers_dist==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition' & cat_workers_dist==`i'
    local somme_temps_tra_EP_mensuel = r(sum)
    
    local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
    mat RESU1 = RESU1 \ (`revenu_horaire')
}


*-------------------------------------------------------------------------------
* Secteur d'activité
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/4 {
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & branche1==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition' & branche1==`i'
    local somme_temps_tra_EP_mensuel = r(sum)
    
    local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
    mat RESU1 = RESU1 \ (`revenu_horaire')
}


*-------------------------------------------------------------------------------
* Ensemble
*-------------------------------------------------------------------------------
qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition'
local somme_revenu_principal_corrige = r(sum)

qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition'
local somme_temps_tra_EP_mensuel = r(sum)

local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
mat RESU1 = RESU1 \ (`revenu_horaire')

*-------------------------------------------------------------------------------
* A.2 Mise en forme du Tableau
*-------------------------------------------------------------------------------

/* Définition des entêtes de lignes et colonnes */
*Lignes 
matrix rownames RESU1 = "Milieu de Residence" "." "Abidjan" "Autre Urbain" "Rural" ///
    "Sexe" "Masculin" "Feminin" ///
    "Groupe d'Age" "16-24 ans" "25-35 ans" "36-64 ans" "65 ans et plus" ///
    "Niveau d'Instruction" "Aucun" "Primaire" "Secondaire 1er Cycle" ///
    "Secondaire 2nd Cycle" "Superieure" ///
    "Statut de l'emploi" "Employeur" "Worker indépendants unemployees" ///
    "Non-salariés (Entrep) dépendants" "Employés" "Travailleurs familiaux" ///
	"Temps de trajet" "Moins de 15 minutes" "15–30 minutes" "30–60 minutes" "Plus de 60 minutes" ///
	"Distance parcourru" "Moins de 5 Km" "5–10 Km" "10 Km et plus"  ///
	"Secteur d'activité" "Agriculture"  "Industrie"  "Commerce" "Autres services"  ///		
    "Ensemble"

*Colonnes 
matrix colnames RESU1 = "Revenu horaire et mobilité"

/* Exportation sur Excel dans le dossier Resultats_Tab*/
putexcel set "${Resultats_Tab2}\Them9_implication_mobilite_travail_surEmploi.xlsx", ///
    sheet("Revenu_horaire") modify

/* Mise en forme */
putexcel B4 = matrix(RESU1), colnames nformat(number_d2)
putexcel A5 = matrix(RESU1), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Revenu horaire selon les caractéristiques socio-démographiques"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Pression temporelle (Revenu horaire / Heures travaillées mensuel)"
putexcel B3, bold

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Démographiques"
putexcel (A3:A4), merge bold

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close


di as result "Tableau du Revenu horaire créé avec succès!"


*===============================================================================
* CALCUL DE LA Part du revenu absorbee par la mobilite
*Calcul de l’indice de Pression Économique de la Mobilité (IPEM)

*On calcul d'abord : Cout implicite
* Cout implicite=Temps trajet (heures)×Revenu horaire

*Puis 
* Part du revenu absorbee = Somme(Cout implicite) / Somme(Revenu mensuel)
* Pour les personnes en emploi uniquement (EN_EMP==1)
*===============================================================================

cap drop revenu_horaire 
gen revenu_horaire = revenu_principal_corrige/temps_tra_EP_mensuel if EN_EMP ==1 & !inlist(EP2g1,5,7,8)

cap drop cout_implicite
gen cout_implicite = 22*2*(EP2h/60)* revenu_horaire if EN_EMP==1 & !inlist(EP2g1,5,7,8) & EP2h<9998


*clear matrix : Part du revenu absorbee par la mobilite
mat define RESU1 = (.)

* Conditions de base pour toutes les analyses
local condition "EN_EMP==1 & !inlist(EP2g1,5,7,8) & MO==1 & age >= 16 & EP2h<9998"
*

*-------------------------------------------------------------------------------
* Milieu de résidence
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/3 {
    qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition' & milieu_resid2==`i'
    local somme_cout_implicite = r(sum)
    
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & milieu_resid2==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
    mat RESU1 = RESU1 \ (`part_rev_absorbe')*100
}

*-------------------------------------------------------------------------------
* Sexe
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/2 {
    qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition' & sexe==`i'
    local somme_cout_implicite = r(sum)
    
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & sexe==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
    mat RESU1 = RESU1 \ (`part_rev_absorbe')*100
}

*-------------------------------------------------------------------------------
* Groupe d'Âge
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/4 {
    qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition' & groupe_age4==`i'
    local somme_cout_implicite = r(sum)
    
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & groupe_age4==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
    mat RESU1 = RESU1 \ (`part_rev_absorbe')*100
}

*-------------------------------------------------------------------------------
* Niveau d'Instruction
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/5 {
    qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition' & Niv_inst_AG3==`i'
    local somme_cout_implicite = r(sum)
    
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & Niv_inst_AG3==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
    mat RESU1 = RESU1 \ (`part_rev_absorbe')*100
}

*-------------------------------------------------------------------------------
* Statut de l'emploi
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/5 {
    qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition' & CISE_18_new==`i'
    local somme_cout_implicite = r(sum)
    
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & CISE_18_new==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
    mat RESU1 = RESU1 \ (`part_rev_absorbe')*100
}

*-------------------------------------------------------------------------------
* Intervalle de temps de parcour
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/4 {
    qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition' & cat_workers==`i'
    local somme_cout_implicite = r(sum)
    
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & cat_workers==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
    mat RESU1 = RESU1 \ (`part_rev_absorbe')*100
}


*-------------------------------------------------------------------------------
* Intervalle de temps de parcour
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/3 {
    qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition' & cat_workers_dist==`i'
    local somme_cout_implicite = r(sum)
    
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & cat_workers_dist==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
    mat RESU1 = RESU1 \ (`part_rev_absorbe')*100
}


*-------------------------------------------------------------------------------
* Secteur d'activité
*-------------------------------------------------------------------------------
mat RESU1 = RESU1 \ (.) 

forvalues i = 1/4 {
    qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition' & branche1==`i'
    local somme_cout_implicite = r(sum)
    
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & branche1==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
    mat RESU1 = RESU1 \ (`part_rev_absorbe')*100
}



*-------------------------------------------------------------------------------
* Ensemble
*-------------------------------------------------------------------------------
qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition'
local somme_cout_implicite = r(sum)

qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition'
local somme_revenu_principal_corrige = r(sum)

local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
mat RESU1 = RESU1 \ (`part_rev_absorbe')*100

*-------------------------------------------------------------------------------
* A.2 Mise en forme du Tableau
*-------------------------------------------------------------------------------

/* Définition des entêtes de lignes et colonnes */
*Lignes 
matrix rownames RESU1 = "Milieu de Residence" "." "Abidjan" "Autre Urbain" "Rural" ///
    "Sexe" "Masculin" "Feminin" ///
    "Groupe d'Age" "16-24 ans" "25-35 ans" "36-64 ans" "65 ans et plus" ///
    "Niveau d'Instruction" "Aucun" "Primaire" "Secondaire 1er Cycle" ///
    "Secondaire 2nd Cycle" "Superieure" ///
    "Statut de l'emploi" "Employeur" "Worker indépendants unemployees" ///
    "Non-salariés (Entrep) dépendants" "Employés" "Travailleurs familiaux" ///
	"Temps de trajet" "Moins de 15 minutes" "15–30 minutes" "30–60 minutes" "Plus de 60 minutes" ///
		"Distance parcourru" "Moins de 5 Km" "5–10 Km" "10 Km et plus"  ///
	"Secteur d'activité" "Agriculture"  "Industrie"  "Commerce" "Autres services"  ///				
    "Ensemble"

*Colonnes 
matrix colnames RESU1 = "Part du revenu absorbee"

/* Exportation sur Excel dans le dossier Resultats_Tab*/
putexcel set "${Resultats_Tab2}\Them9_implication_mobilite_travail_surEmploi.xlsx", ///
    sheet("Part_revenu_absorbee") modify

/* Mise en forme */
putexcel B4 = matrix(RESU1), colnames nformat(number_d2)
putexcel A5 = matrix(RESU1), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Part du revenu absorbee selon les caractéristiques socio-démographiques"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Part du revenu absorbee (Cout implicite) / (Revenu mensuel)"
putexcel B3, bold

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Démographiques"
putexcel (A3:A4), merge bold

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close












**** RELABELISATION DE LA VARIABLE REGION HH2
cap drop region
gen region =1  if HH2 ==10101
replace region = 2  if HH2 ==10916
replace region = 3  if HH2 ==11319
replace region = 4  if HH2 ==11120
replace region = 5  if HH2 ==10821
replace region = 6  if HH2 ==11322
replace region = 7  if HH2 ==11423
replace region = 8  if HH2 ==11018
replace region = 9  if HH2 ==10524
replace region = 10  if HH2 ==11204
replace region = 11  if HH2 ==10325
replace region = 12  if HH2 ==10617
replace region = 13  if HH2 ==11408
replace region = 14  if HH2 ==10926
replace region = 15  if HH2 ==11027
replace region = 16  if HH2 ==11228
replace region = 17  if HH2 ==10702
replace region = 18  if HH2 ==10829
replace region = 19  if HH2 ==10405
replace region = 20  if HH2 ==10510
replace region = 21  if HH2 ==10930
replace region = 22  if HH2 ==10615
replace region = 23  if HH2 ==10712
replace region = 24  if HH2 ==10833
replace region = 25  if HH2 ==10331
replace region = 26  if HH2 ==10811
replace region = 27  if HH2 ==11103
replace region = 28  if HH2 ==10309
replace region = 29  if HH2 ==10413
replace region = 30  if HH2 ==11132
replace region = 31  if HH2 ==11006
replace region = 32  if HH2 ==11314
replace region = 33  if HH2 ==10207


*===============================================================================
* CALCUL DE LA Part du revenu absorbee par la mobilite

*On calcul d'abord : Cout implicite
* Cout implicite=Temps trajet (heures)×Revenu horaire

*Puis 
* Part du revenu absorbee = Somme(Cout implicite) / Somme(Revenu mensuel)
* Pour les personnes en emploi uniquement (EN_EMP==1)
*===============================================================================

cap drop revenu_horaire 
gen revenu_horaire = revenu_principal_corrige/temps_tra_EP_mensuel if EN_EMP ==1 & !inlist(EP2g1,5,7,8)

cap drop cout_implicite
gen cout_implicite = 22*2*(EP2h/60)* revenu_horaire if EN_EMP==1 & !inlist(EP2g1,5,7,8) & EP2h<9998


* Conditions de base pour toutes les analyses
local condition "EN_EMP==1 & !inlist(EP2g1,5,7,8) & MO==1 & age >= 16 & EP2h<9998"

*clear matrix : Part du revenu absorbee par la mobilite par region
mat define RESU1 = (.)

mat RESU1 = RESU1 \ (.) 

forvalues i = 1/33 {
    qui summ cout_implicite [aw=pmencor_ind_annuel] if `condition' & region==`i'
    local somme_cout_implicite = r(sum)
    
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & region==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    local part_rev_absorbe = `somme_cout_implicite' / `somme_revenu_principal_corrige'
    mat RESU1 = RESU1 \ (`part_rev_absorbe')*100
}

*===============================================================================
* CALCUL DU Revenu horaire
* Revenu horaire = Somme(revenu_principal_corrige) / Somme(temps_tra_EP_mensuel)
* Pour les personnes en emploi uniquement (EN_EMP==1)
*===============================================================================

cap drop temps_tra_EP_mensuel
gen temps_tra_EP_mensuel = temps_tra_EP_semaine*4 if EN_EMP==1


*clear matrix : Pression temporelle de la mobilité
mat define RESU2 = (.)

* Conditions de base pour toutes les analyses
local condition "EN_EMP==1 & temps_tra_EP_mensuel!=. & temps_tra_EP_mensuel>0"


*-------------------------------------------------------------------------------
* Milieu de résidence
*-------------------------------------------------------------------------------
mat RESU2 = RESU2 \ (.) 

forvalues i = 1/33 {
    qui summ revenu_principal_corrige [aw=pmencor_ind_annuel] if `condition' & region==`i'
    local somme_revenu_principal_corrige = r(sum)
    
    qui summ temps_tra_EP_mensuel [aw=pmencor_ind_annuel] if `condition' & region==`i'
    local somme_temps_tra_EP_mensuel = r(sum)
    
    local revenu_horaire = `somme_revenu_principal_corrige' / `somme_temps_tra_EP_mensuel'
    mat RESU2 = RESU2 \ (`revenu_horaire')
}


*===============================================================================
* CALCUL DE LA PRESSION TEMPORELLE DE LA MOBILITÉ
* Pression temporelle = Somme(EP2h) / Somme(temps_tra_EP_semaine)
* Pour les personnes en emploi uniquement (EN_EMP==1)
*===============================================================================


*clear matrix : Pression temporelle de la mobilité
mat define RESU3 = (.)

* Conditions de base pour toutes les analyses
local condition "EN_EMP==1 & EP2h<9998 & temps_tra_EP_semaine!=. & temps_tra_EP_semaine>0"

***** Le temps de trajet EP2h est en minutes et l'heures travaillées temps_tra_EP_semaine est en heures. PAr consequent il faut ramener le Le temps de trajet EP2h en heure en divisant par 60.

*Ensuite le EP2h est le temps de trajet jounalier, donc faut le mutiplier par 2 pour estimer le temsp (aller-retour) et mutiplier  par 5 pour estimer le temps de trajet de la semaine (alller-retour). Ce qui serait conforme à temps_tra_EP_semaine

*** Donc le coeficient multiplicatif de EP2h est : 2*5/60 = 10/60 = 1/6
*-------------------------------------------------------------------------------
* Milieu de résidence
*-------------------------------------------------------------------------------
mat RESU3 = RESU3 \ (.) 

forvalues i = 1/33 {
    qui summ EP2h [aw=pmencor_ind_annuel] if `condition' & region==`i'
    local somme_ep2h = r(sum)*(1/6)
    
    qui summ temps_tra_EP_semaine [aw=pmencor_ind_annuel] if `condition' & region==`i'
    local somme_temps = r(sum)
    
    local pression = `somme_ep2h' / `somme_temps'
    mat RESU3 = RESU3 \ (`pression')*100
}


mat RESU = RESU1[1..35,1], RESU2[1..35,1], RESU3[1..35,1]


*A.2 Mise en forme du Tableau 
		/*---------------------------*/
		
/* Définition des entête de lignes et colonnes */

*Lignes 

matrix rownames RESU = "REGION" "." "ABIDJAN" "AGNEBY-TIASSA"	"BAFING" "BAGOUE" "BELIER" "BERE" "BOUNKANI" "CAVALLY" "FOLON"	"GBEKE"	"GBOKLE" "GÔH" "GONTOUGO" "GRAND-PONTS"	"GUEMON"	"HAMBOL" "HAUT-SASSANDRA" "IFFOU" "INDENIE-DJUABLIN"	"KABADOUGOU" "LA ME" "LÔH-DJIBOUA" "MARAHOUE" "MORONOU"	"NAWA"	"N'ZI" "PORO" "SAN-PEDRO" "SUD-COMOE" "TCHOLOGO" "TONKPI"	"WORODOUGOU" "YAMOUSSOUKRO"


*Colonnes 

matrix colnames RESU = "Part du revenu absorbee" "Revenu horaire" "Pression temporelle"


/* Exportation sur Excel dans le dossier Resultats_Tab*/

putexcel set "${Resultats_Tab2}\Them9_implication_mobilite_travail_surEmploi.xlsx", sheet("Impact_mobilite_region") modify


/* Mise en forme */
putexcel B4 = matrix(RESU), colnames  
putexcel A5 = matrix(RESU), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Impact de la mobilité domicile-travail sur l'emploi et le revenu selon la region"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Impact de la mobilité domicile-travail sur l'emploi "
putexcel (B3:E3), merge

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Demographiques"
putexcel (A3:A4), merge

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close















di as result "Tableau de pression temporelle créé avec succès!"