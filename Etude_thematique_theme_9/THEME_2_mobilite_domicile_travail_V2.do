
clear
**************************************************************************
******* THEME 2 : Mobilité domicile travail (MARTIAL)  ********
*******															  ********
**************************************************************************

******* I. Caractéristiques générales de l'emploi principal

******* (Module : Emploi principal)
******* Objectif : Identifier les grandes catégories d'emplois et leur statut socio-économique

global dofile_theme = "C:\Users\mg.kouame\OneDrive - GOUVCI\CAE_INS\EMPLOYMENT_ISSUES\Analyses_Thematiques_ENE\Analyses Thematiques ENE\Analyses_Thematiques_ENE\dofiles"

do "$dofile_theme\Theme9_prelancement.do"

*** RENOMMAGE DE CERTAINEME VARIABLE


do "$dofile_theme\Duree_de_travail_EP.do"



*Utilisation de la nouvelle base pour le travail
use "$donnees_BT\Base_Travail_BT_vf_RENAMED_plus_salaire_EP.dta", replace


global Pilote_Analyse2    = "D:\ENEM_Working\Article\Analyses_Thematiques_ENE\Analyses_Thematiques_ENE"

*global donnees   = "D:\ENEM_Working\Article\Analyses_Thematiques_ENE\Analyses_Thematiques_ENE\donnees\master_data_indiv_men_3"

/* Définition du dossier de travail du bulletin trimestriel */
global Analyses_Thematiques_ENE = "${Pilote_Analyse2}"


global Resultats_Tab2 = "${Pilote_Analyse2}\docs"


*use "$donnees\Base_Travail_BT_vf_RENAMED_plus_salaire_EP.dta", replace



*replace trimestre="T2" if trimestre=="T2-2024"
*replace trimestre="T3" if trimestre=="T3-2024"
*replace trimestre="T4" if trimestre=="T4-2024"

replace trimestre="T2-2024" if trimestre=="24T2"
replace trimestre="T3-2024" if trimestre=="24T3"
replace trimestre="T4-2024" if trimestre=="24T4"


*replace pmencor_ind_annuel =1

/* 3.x. Répartition de l'emploi par nature selon les caractéristiques */

clonevar form_empEP = CISE_18_informel_Emp 
recode  form_empEP (2=0)
lab define form_empEP 1 "Emploi formel" 0 " Emploi informel"
lab values form_empEP form_empEP 



/*

III. Caractéristiques générales des trajets domicile-travail

*/

***********************************************************
******													***
****** Temps moyen par moyen de tranport (en minutes)	***
******													***
***********************************************************


*clear matrix : Transport assuré par l'employeur
mat define RESU1 = (.)
/* Milieu de résidence x Transport en commun public */
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==1 & EP2h<9998, over(milieu_resid2)
mat list r(table)
mat RESU1 = RESU1 \ ((r(table)[rownumb(r(table),"b"), 1..3])')

/* Sexe */
mat RESU1 = RESU1 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==1 & EP2h<9998, over(sexe)
mat list r(table)
mat RESU1 = RESU1 \ ((r(table)[rownumb(r(table),"b"), 1..2])')

/*Groupe d'Age */
mat RESU1 = RESU1 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==1 & EP2h<9998, over(groupe_age4)
mat list r(table)
mat RESU1 = RESU1 \ ((r(table)[rownumb(r(table),"b"), 1..4])')

/* Niveau d'Instruction */
mat RESU1 = RESU1 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==1 & EP2h<9998, over(Niv_inst_AG3)
mat list r(table)
mat RESU1 = RESU1 \ ((r(table)[rownumb(r(table),"b"), 1..5])')


/* Statut de l'emploi */
mat RESU1 = RESU1 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==1 & EP2h<9998, over(CISE_18_new)
mat list r(table)
mat RESU1 = RESU1 \ ((r(table)[rownumb(r(table),"b"), 1..5])')

/* Secteur d'activité */
mat RESU1 = RESU1 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==1 & EP2h<9998, over(branche1)
mat list r(table)
mat RESU1 = RESU1 \ ((r(table)[rownumb(r(table),"b"), 1..4])')



/* Ensemble */
mean EP2h [pw=pmencor_ind_annuel] if   MO==1 & age >= 16 & EP2g1==1 & EP2h<9998
mat list r(table)
mat RESU1 = RESU1 \ ((r(table)[rownumb(r(table),"b"), 1])')




*Annee
*clear matrix : Transport en commun public
mat define RESU2 = (.)
/* Milieu de résidence x Transport en commun public */
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==2 & EP2h<9998, over(milieu_resid2)
mat list r(table)
mat RESU2 = RESU2 \ ((r(table)[rownumb(r(table),"b"), 1..3])')

/* Sexe */
mat RESU2 = RESU2 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==2 & EP2h<9998, over(sexe)
mat list r(table)
mat RESU2 = RESU2 \ ((r(table)[rownumb(r(table),"b"), 1..2])')

/*Groupe d'Age */
mat RESU2 = RESU2 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==2 & EP2h<9998, over(groupe_age4)
mat list r(table)
mat RESU2 = RESU2 \ ((r(table)[rownumb(r(table),"b"), 1..4])')

/* Niveau d'Instruction */
mat RESU2 = RESU2 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==2 & EP2h<9998, over(Niv_inst_AG3)
mat list r(table)
mat RESU2 = RESU2 \ ((r(table)[rownumb(r(table),"b"), 1..5])')


/* Statut de l'emploi */
mat RESU2 = RESU2 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==2 & EP2h<9998, over(CISE_18_new)
mat list r(table)
mat RESU2 = RESU2 \ ((r(table)[rownumb(r(table),"b"), 1..5])')


/* Secteur d'activité */
mat RESU2 = RESU2 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==2 & EP2h<9998, over(branche1)
mat list r(table)
mat RESU2 = RESU2 \ ((r(table)[rownumb(r(table),"b"), 1..4])')


/* Ensemble */
mean EP2h [pw=pmencor_ind_annuel] if   MO==1 & age >= 16 & EP2g1==2 & EP2h<9998
mat list r(table)
mat RESU2 = RESU2 \ ((r(table)[rownumb(r(table),"b"), 1])')


*clear matrix : Transport en commun privée
mat define RESU3 = (.)
/* Milieu de résidence x Transport en commun privée */
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==3 & EP2h<9998, over(milieu_resid2)
mat list r(table)
mat RESU3 = RESU3 \ ((r(table)[rownumb(r(table),"b"), 1..3])')

/* Sexe */
mat RESU3 = RESU3 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==3 & EP2h<9998, over(sexe)
mat list r(table)
mat RESU3 = RESU3 \ ((r(table)[rownumb(r(table),"b"), 1..2])')

/*Groupe d'Age */
mat RESU3 = RESU3 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==3 & EP2h<9998, over(groupe_age4)
mat list r(table)
mat RESU3 = RESU3 \ ((r(table)[rownumb(r(table),"b"), 1..4])')

/* Niveau d'Instruction */
mat RESU3 = RESU3 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==3 & EP2h<9998, over(Niv_inst_AG3)
mat list r(table)
mat RESU3 = RESU3 \ ((r(table)[rownumb(r(table),"b"), 1..5])')


/* Statut de l'emploi */
mat RESU3 = RESU3 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==3 & EP2h<9998, over(CISE_18_new)
mat list r(table)
mat RESU3 = RESU3 \ ((r(table)[rownumb(r(table),"b"), 1..5])')

/* Secteur d'activité */
mat RESU3 = RESU3 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==3 & EP2h<9998, over(branche1)
mat list r(table)
mat RESU3 = RESU3 \ ((r(table)[rownumb(r(table),"b"), 1..4])')


/* Ensemble */
mean EP2h [pw=pmencor_ind_annuel] if  MO==1 & age >= 16 & EP2g1==3 & EP2h<9998
mat list r(table)
mat RESU3 = RESU3 \ ((r(table)[rownumb(r(table),"b"), 1])')


*clear matrix : Transport personnel
mat define RESU4 = (.)
/* Milieu de résidence x Transport personnel */
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==4 & EP2h<9998, over(milieu_resid2)
mat list r(table)
mat RESU4 = RESU4 \ ((r(table)[rownumb(r(table),"b"), 1..3])')

/* Sexe */
mat RESU4 = RESU4 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==4 & EP2h<9998, over(sexe)
mat list r(table)
mat RESU4 = RESU4 \ ((r(table)[rownumb(r(table),"b"), 1..2])')

/*Groupe d'Age */
mat RESU4 = RESU4 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16 &  EP2g1==4 & EP2h<9998, over(groupe_age4)
mat list r(table)
mat RESU4 = RESU4 \ ((r(table)[rownumb(r(table),"b"), 1..4])')

/* Niveau d'Instruction */
mat RESU4 = RESU4 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==4 & EP2h<9998, over(Niv_inst_AG3)
mat list r(table)
mat RESU4 = RESU4 \ ((r(table)[rownumb(r(table),"b"), 1..5])')


/* Statut de l'emploi */
mat RESU4 = RESU4 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16 &  EP2g1==4 & EP2h<9998, over(CISE_18_new)
mat list r(table)
mat RESU4 = RESU4 \ ((r(table)[rownumb(r(table),"b"), 1..5])')

/* Secteur d'activité */
mat RESU4 = RESU4 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==4 & EP2h<9998, over(branche1)
mat list r(table)
mat RESU4 = RESU4 \ ((r(table)[rownumb(r(table),"b"), 1..4])')


/* Ensemble */
mean EP2h [pw=pmencor_ind_annuel] if  MO==1 & age >= 16 & EP2g1==4 & EP2h<9998
mat list r(table)
mat RESU4 = RESU4 \ ((r(table)[rownumb(r(table),"b"), 1])')





*"T3"
*clear matrix : A pieds
mat define RESU5 = (.)
/* Milieu de résidence x A pieds */
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==5 & EP2h<9998, over(milieu_resid2)
mat list r(table)
mat RESU5 = RESU5 \ ((r(table)[rownumb(r(table),"b"), 1..3])')

/* Sexe */
mat RESU5 = RESU5 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==5 & EP2h<9998, over(sexe)
mat list r(table)
mat RESU5 = RESU5 \ ((r(table)[rownumb(r(table),"b"), 1..2])')

/*Groupe d'Age */
mat RESU5 = RESU5 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==5 & EP2h<9998, over(groupe_age4)
mat list r(table)
mat RESU5 = RESU5 \ ((r(table)[rownumb(r(table),"b"), 1..4])')

/* Niveau d'Instruction */
mat RESU5 = RESU5 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==5 & EP2h<9998, over(Niv_inst_AG3)
mat list r(table)
mat RESU5 = RESU5 \ ((r(table)[rownumb(r(table),"b"), 1..5])')


/* Statut de l'emploi */
mat RESU5 = RESU5 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==5 & EP2h<9998, over(CISE_18_new)
mat list r(table)
mat RESU5 = RESU5 \ ((r(table)[rownumb(r(table),"b"), 1..5])')

/* Secteur d'activité */
mat RESU5 = RESU5 \ (.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2g1==5 & EP2h<9998, over(branche1)
mat list r(table)
mat RESU5 = RESU5 \ ((r(table)[rownumb(r(table),"b"), 1..4])')


/* Ensemble */
mean EP2h [pw=pmencor_ind_annuel] if  MO==1 & age >= 16 & EP2g1==5 & EP2h<9998
mat list r(table)
mat RESU5 = RESU5 \ ((r(table)[rownumb(r(table),"b"), 1])')







mat RESU = RESU1[1..30,1], RESU2[1..30,1], RESU3[1..30,1], RESU4[1..30,1], RESU5[1..30,1]



*A.2 Mise en forme du Tableau 
		/*---------------------------*/
		
/* Définition des entête de lignes et colonnes */

*Lignes 

matrix rownames RESU = "Milieu de Residence" "Abidjan" "Autre Urbain" "Rural" "Sexe" "Masculin"  "Feminin" "Groupe d'Age" "16-24 ans" "25-35 ans" "36-64 ans" "65 ans et plus"  "Niveau d'Instruction"  "Aucun" "Primaire" "Secondaire  1er Cycle" "Secondaire  2nd Cycle" "Superieure"   "Statut de l'emploi" "Employeur"  "worker indépendants unemployees" "Non-salariés (Entrep) dépendants"  "Employés"  "Travailleurs familiaux" "Secteur d'activité" "Agriculture"  "Industrie"  "Commerce" "Autres services" "Ensemble"

*Colonnes 

matrix colnames RESU = "Transport assuré par l'employeur" "Transport en commun public" "Transport en commun privée" "Transport personnel" "A pieds" 


/* Exportation sur Excel dans le dossier Resultats_Tab*/

putexcel set "${Resultats_Tab2}\Tableau_Emploi_Ensemble_theme9_V2.xlsx", sheet("Temps par moyen transport") modify


/* Mise en forme */
putexcel B4 = matrix(RESU), colnames  nformat(number_d2)
putexcel A5 = matrix(RESU), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Répartition du Temps moyen par moyen de Transport  (en minutes) selon les caractéristiques"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Temps moyen de trajet par moyen de transport (en minutes)"
putexcel (B3:E3), merge

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Demographiques"
putexcel (A3:A4), merge

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close





* b. Répartition des travailleurs selon la durée du trajet

cap drop cat_workers
gen cat_workers =1 if 15>EP2h & MO==1

replace cat_workers =2  if EP2h>=15 & 30>EP2h  & MO==1

replace cat_workers =3  if EP2h>=30 & 60>EP2h  & MO==1

replace cat_workers =4  if EP2h>=60 & .>EP2h  & MO==1

lab define cat_workers_lbl 1 "<15 min" 2 " 15–30 min" 3 "30–59 min" 4 "≥60 min"
lab values cat_workers cat_workers_lbl 



*"Annee"
*clear matrix
mat define RESU = (.,.,.,.)
/* Milieu de résidence */
proportion cat_workers [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(milieu_resid2)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..3])',(r(table)[rownumb(r(table),"b"),4..6])',(r(table)[rownumb(r(table),"b"),7..9])',(r(table)[rownumb(r(table),"b"),10..12])')*100

/* Sexe */
mat RESU = RESU \ (.,.,.,.) 
proportion cat_workers [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(sexe)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..2])',(r(table)[rownumb(r(table),"b"),3..4])',(r(table)[rownumb(r(table),"b"),5..6])',(r(table)[rownumb(r(table),"b"),7..8])')*100

/*Groupe d'Age */
mat RESU = RESU \ (.,.,.,.)
proportion cat_workers [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(groupe_age4)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"b"),5..8])',(r(table)[rownumb(r(table),"b"),9..12])',(r(table)[rownumb(r(table),"b"),13..16])')*100

/* Niveau d'Instruction */
mat RESU = RESU \ (.,.,.,.) 
proportion cat_workers [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(Niv_inst_AG3)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"b"),6..10])',(r(table)[rownumb(r(table),"b"),11..15])',(r(table)[rownumb(r(table),"b"),16..20])')*100


/* Statut de l'emploi */
mat RESU = RESU \ (.,.,.,.) 
proportion cat_workers [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(CISE_18_new)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"b"),6..10])',(r(table)[rownumb(r(table),"b"),11..15])',(r(table)[rownumb(r(table),"b"),16..20])')*100

/* Secteur d'activité */
mat RESU = RESU \ (.,.,.,.) 
proportion cat_workers [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(branche1)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"b"),5..8])',(r(table)[rownumb(r(table),"b"),9..12])',(r(table)[rownumb(r(table),"b"),13..16])')*100


/* Ensemble */
proportion cat_workers [pw=pmencor_ind_annuel] if  MO==1 & age >= 16
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1])',(r(table)[rownumb(r(table),"b"),2])',(r(table)[rownumb(r(table),"b"),3])',(r(table)[rownumb(r(table),"b"),4])')*100




*mat RESU =  RESU[1..25,1], RESU[1..25,2], RESU[1..25,3] RESU[1..25,4], RESU_[1..25,4]



*A.2 Mise en forme du Tableau 
		/*---------------------------*/
		
/* Définition des entête de lignes et colonnes */

*Lignes 

matrix rownames RESU = "Milieu de Residence" "Abidjan" "Autre Urbain" "Rural" "Sexe" "Masculin"  "Feminin" "Groupe d'Age" "16-24 ans" "25-35 ans" "36-64 ans" "65 ans et plus"  "Niveau d'Instruction"  "Aucun" "Primaire" "Secondaire  1er Cycle" "Secondaire  2nd Cycle" "Superieure"   "Statut de l'emploi" "Employeur"  "worker indépendants unemployees" "Non-salariés (Entrep) dépendants"  "Employés"  "Travailleurs familiaux" "Secteur d'activité" "Agriculture"  "Industrie"  "Commerce" "Autres services"  "Ensemble"

*Colonnes 

matrix colnames RESU = "15 min et moins"  "15–29 min" "30–59 min" "60 min et plus"


/* Exportation sur Excel dans le dossier Resultats_Tab*/

putexcel set "${Resultats_Tab2}\Tableau_Emploi_Ensemble_theme9_V2.xlsx", sheet("Travailleurs_par_duree_trajet") modify


/* Mise en forme */
putexcel B4 = matrix(RESU), colnames  nformat(number_d2)
putexcel A5 = matrix(RESU), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Répartition des travailleurs par durée du trajet (en minutes) selon les caractéristiques"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Répartition des travailleurs selon la durée du trajet (en minutes)"
putexcel (B3:E3), merge

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Demographiques"
putexcel (A3:A4), merge

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close



* c. Répartition des travailleurs selon la distance estimée du trajet
/*
moins de 5 KM
entre 5 et 10KM 
plus de 10Km
*/
cap drop cat_workers_dist
gen cat_workers_dist =1 if 5>EP2i & MO==1

replace cat_workers_dist =2  if EP2i>=5 & 10>EP2i  & MO==1

replace cat_workers_dist =3  if EP2i>=10 & .>EP2i  & MO==1

*replace cat_workers_dist =4  if EP2i>=10 & .>EP2i  & MO==1

lab define cat_workers_dist_lbl 1 "<5 Km" 2 " 5–10 Km" 3 "≥10 Km"
lab values cat_workers_dist cat_workers_dist_lbl 


*"Annee"
*clear matrix
mat define RESU = (.,.,.)
/* Milieu de résidence */
proportion cat_workers_dist [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(milieu_resid2)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..3])',(r(table)[rownumb(r(table),"b"),4..6])',(r(table)[rownumb(r(table),"b"),7..9])')*100

/* Sexe */
mat RESU = RESU \ (.,.,.) 
proportion cat_workers_dist [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(sexe)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..2])',(r(table)[rownumb(r(table),"b"),3..4])',(r(table)[rownumb(r(table),"b"),5..6])')*100

/*Groupe d'Age */
mat RESU = RESU \ (.,.,.)
proportion cat_workers_dist [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(groupe_age4)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"b"),5..8])',(r(table)[rownumb(r(table),"b"),9..12])')*100

/* Niveau d'Instruction */
mat RESU = RESU \ (.,.,.) 
proportion cat_workers_dist [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(Niv_inst_AG3)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"b"),6..10])',(r(table)[rownumb(r(table),"b"),11..15])')*100


/* Statut de l'emploi */
mat RESU = RESU \ (.,.,.) 
proportion cat_workers_dist [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(CISE_18_new)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"b"),6..10])',(r(table)[rownumb(r(table),"b"),11..15])')*100

/* Moyen de déplacement */
mat RESU = RESU \ (.,.,.) 
proportion cat_workers_dist [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(EP2g1)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"b"),9..13])',(r(table)[rownumb(r(table),"b"),17..21])')*100

/* Secteur d'activité */
mat RESU = RESU \ (.,.,.) 
proportion cat_workers_dist [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(branche1)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"b"),5..8])',(r(table)[rownumb(r(table),"b"),9..12])')*100



/* Ensemble */

proportion cat_workers_dist [pw=pmencor_ind_annuel] if  MO==1 & age >= 16
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1]),(r(table)[rownumb(r(table),"b"),2]),(r(table)[rownumb(r(table),"b"),3]))*100




*mat RESU = RESU[1..25,1], RESU[1..25,2], RESU[1..25,3],  RESU[1..25,4]



*A.2 Mise en forme du Tableau 
		/*---------------------------*/
		
/* Définition des entête de lignes et colonnes */

*Lignes 

matrix rownames RESU = "Milieu de Residence" "Abidjan" "Autre Urbain" "Rural" "Sexe" "Masculin"  "Feminin" "Groupe d'Age" "16-24 ans" "25-35 ans" "36-64 ans" "65 ans et plus"  "Niveau d'Instruction"  "Aucun" "Primaire" "Secondaire  1er Cycle" "Secondaire  2nd Cycle" "Superieure"   "Statut de l'emploi" "Employeur"  "worker indépendants unemployees" "Non-salariés (Entrep) dépendants"  "Employés"  "Travailleurs familiaux" "Moyen de déplacement" "Transport assuré par l'employeur" "Transport en commun public" "Transport en commun privée" "Transport personnel" "A pieds" "Secteur d'activité" "Agriculture"  "Industrie"  "Commerce" "Autres services" "Ensemble"

*Colonnes 

matrix colnames RESU = "Moins de 5 Km" "5–10 Km"  "10 Km et plus"


/* Exportation sur Excel dans le dossier Resultats_Tab*/

putexcel set "${Resultats_Tab2}\Tableau_Emploi_Ensemble_theme9_V2.xlsx", sheet("Travailleurs_par_longeur_trajet") modify


/* Mise en forme */
putexcel B4 = matrix(RESU), colnames  nformat(number_d2)
putexcel A5 = matrix(RESU), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Répartition des travailleurs par la distance estimée du trajet (en Km) selon les caractéristiques"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Répartition des travailleurs selon la distance estimée du trajet (en Km)"
putexcel (B3:E3), merge

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Demographiques"
putexcel (A3:A4), merge

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close



****	E. Répartition selon la variable EP2e (Lieu de travail dans la localité de résidence)



*"Annee"
*clear matrix
mat define RESU = (.,.)
/* Milieu de résidence */
proportion EP2e [pw=pmencor_ind_annuel] if MO==1 & age >= 16  , over(milieu_resid2)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..3])',(r(table)[rownumb(r(table),"se"),1..3])')*100

/* Sexe */
mat RESU = RESU \ (.,.) 
proportion EP2e [pw=pmencor_ind_annuel] if MO==1 & age >= 16  , over(sexe)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..2])',(r(table)[rownumb(r(table),"se"),1..2])')*100

/*Groupe d'Age */
mat RESU = RESU \ (.,.) 
proportion EP2e [pw=pmencor_ind_annuel] if MO==1 & age >= 16  , over(groupe_age4)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"se"),1..4])')*100

/* Niveau d'Instruction */
mat RESU = RESU \ (.,.) 
proportion EP2e [pw=pmencor_ind_annuel] if MO==1 & age >= 16  , over(Niv_inst_AG3)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"se"),1..5])')*100


/* Statut de l'emploi */
mat RESU = RESU \ (.,.) 
proportion EP2e [pw=pmencor_ind_annuel] if MO==1 & age >= 16  , over(CISE_18_new)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"se"),1..5])')*100

/*Secteur d'activité */
mat RESU = RESU \ (.,.) 
proportion EP2e [pw=pmencor_ind_annuel] if MO==1 & age >= 16  , over(branche1)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"se"),1..4])')*100


/* Ensemble */
proportion EP2e [pw=pmencor_ind_annuel] if  MO==1 & age >= 16 
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1])',(r(table)[rownumb(r(table),"se"),1])')*100




*A.2 Mise en forme du Tableau 
		/*---------------------------*/
		
/* Définition des entête de lignes et colonnes */

*Lignes 

matrix rownames RESU = "Milieu de Residence" "Abidjan" "Autre Urbain" "Rural" "Sexe" "Masculin"  "Feminin" "Groupe d'Age" "16-24 ans" "25-35 ans" "36-64 ans" "65 ans et plus"  "Niveau d'Instruction"  "Aucun" "Primaire" "Secondaire  1er Cycle" "Secondaire  2nd Cycle" "Superieure"   "Statut de l'emploi" "Employeur"  "worker indépendants unemployees" "Non-salariés (Entrep) dépendants"  "Employés"  "Travailleurs familiaux"  "Secteur d'activité" "Agriculture"  "Industrie"  "Commerce" "Autres services" "Ensemble"

*Colonnes 

matrix colnames RESU = "Dans la localité de residence(%)"  "Ecart type de la proportion" 


/* Exportation sur Excel dans le dossier Resultats_Tab*/

putexcel set "${Resultats_Tab2}\Tableau_Emploi_Ensemble_theme9_V2.xlsx", sheet("Lieu de travail") modify


/* Mise en forme */
putexcel B4 = matrix(RESU), colnames  nformat(number_d2)
putexcel A5 = matrix(RESU), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : 2.	Répartition selon la variable EP2e (Lieu de travail dans la localité de résidence) selon les caractéristiques"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Lieu de travail dans la localité de résidence"
putexcel (B3:E3), merge

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Demographiques"
putexcel (A3:A4), merge

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close





/*

IV. Modes de transport et accessibilité

Tableau 1.1 – Répartition des modes de transport principaux utilisés selon les caractéristiques


*/

*** d. 	Répartition des modes de transport principaux (EP2g1) utilisés


*"Annne"
*clear matrix
mat define RESU = (.,.,.,.,.,.,.)
/* Milieu de résidence */
proportion EP2g1 [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & inlist(EP2g1,1,2,3,4,5,7,8), over(milieu_resid2)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..3])',(r(table)[rownumb(r(table),"b"),4..6])',(r(table)[rownumb(r(table),"b"),7..9])',(r(table)[rownumb(r(table),"b"),10..12])',(r(table)[rownumb(r(table),"b"),13..15])',(r(table)[rownumb(r(table),"b"),16..18])',(r(table)[rownumb(r(table),"b"),19..21])')*100

/* Sexe */
mat RESU = RESU \ (.,.,.,.,.,.,.) 
proportion EP2g1 [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & inlist(EP2g1,1,2,3,4,5,7,8) , over(sexe)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..2])',(r(table)[rownumb(r(table),"b"),3..4])',(r(table)[rownumb(r(table),"b"),5..6])',(r(table)[rownumb(r(table),"b"),7..8])',(r(table)[rownumb(r(table),"b"),9..10])',(r(table)[rownumb(r(table),"b"),11..12])',(r(table)[rownumb(r(table),"b"),13..14])')*100

/*Groupe d'Age */
mat RESU = RESU \ (.,.,.,.,.,.,.)
proportion EP2g1 [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & inlist(EP2g1,1,2,3,4,5,7,8), over(groupe_age4)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"b"),5..8])',(r(table)[rownumb(r(table),"b"),9..12])',(r(table)[rownumb(r(table),"b"),13..16])',(r(table)[rownumb(r(table),"b"),17..20])',(r(table)[rownumb(r(table),"b"),21..24])',(r(table)[rownumb(r(table),"b"),25..28])')*100

/* Niveau d'Instruction */
mat RESU = RESU \ (.,.,.,.,.,.,.) 
proportion EP2g1 [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & inlist(EP2g1,1,2,3,4,5,7,8), over(Niv_inst_AG3)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"b"),6..10])',(r(table)[rownumb(r(table),"b"),11..15])',(r(table)[rownumb(r(table),"b"),16..20])',(r(table)[rownumb(r(table),"b"),21..25])',(r(table)[rownumb(r(table),"b"),26..30])',(r(table)[rownumb(r(table),"b"),31..35])')*100


/* Statut de l'emploi */
mat RESU = RESU \ (.,.,.,.,.,.,.) 
proportion EP2g1 [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & inlist(EP2g1,1,2,3,4,5,7,8), over(CISE_18_new)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"b"),6..10])',(r(table)[rownumb(r(table),"b"),11..15])',(r(table)[rownumb(r(table),"b"),16..20])',(r(table)[rownumb(r(table),"b"),21..25])',(r(table)[rownumb(r(table),"b"),26..30])',(r(table)[rownumb(r(table),"b"),31..35])')*100


/*Secteur d'activité*/
mat RESU = RESU \ (.,.,.,.,.,.,.)
proportion EP2g1 [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & inlist(EP2g1,1,2,3,4,5,7,8), over(branche1)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"b"),5..8])',(r(table)[rownumb(r(table),"b"),9..12])',(r(table)[rownumb(r(table),"b"),13..16])',(r(table)[rownumb(r(table),"b"),17..20])',(r(table)[rownumb(r(table),"b"),21..24])',(r(table)[rownumb(r(table),"b"),25..28])')*100


/* Ensemble */
proportion EP2g1 [pw=pmencor_ind_annuel] if  MO==1 & age >= 16 & inlist(EP2g1,1,2,3,4,5,7,8)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1])',(r(table)[rownumb(r(table),"b"),2])',(r(table)[rownumb(r(table),"b"),3])',(r(table)[rownumb(r(table),"b"),4])',(r(table)[rownumb(r(table),"b"),5])',(r(table)[rownumb(r(table),"b"),6])',(r(table)[rownumb(r(table),"b"),7])')*100




mat RESU =  RESU[1..30,1], RESU[1..30,2], RESU[1..30,3], RESU[1..30,4],  RESU[1..30,5], RESU[1..30,6],  RESU[1..30,7]



*A.2 Mise en forme du Tableau 
		/*---------------------------*/
		
/* Définition des entête de lignes et colonnes */

*Lignes 

matrix rownames RESU = "Milieu de Residence" "Abidjan" "Autre Urbain" "Rural" "Sexe" "Masculin"  "Feminin" "Groupe d'Age" "16-24 ans" "25-35 ans" "36-64 ans" "65 ans et plus"  "Niveau d'Instruction"  "Aucun" "Primaire" "Secondaire  1er Cycle" "Secondaire  2nd Cycle" "Superieure"   "Statut de l'emploi" "Employeur"  "worker indépendants unemployees" "Non-salariés (Entrep) dépendants"  "Employés"  "Travailleurs familiaux"   "Secteur d'activité" "Agriculture"  "Industrie"  "Commerce" "Autres services"  "Ensemble"

*Colonnes 

matrix colnames RESU = "Transport par l'employeur"  "Transport en commun public"  "Transport en commun privée"  "Transport personnel"  "A pieds"  "Travailleur à domicile"  "Vendeur ambulant"


/* Exportation sur Excel dans le dossier Resultats_Tab*/

putexcel set "${Resultats_Tab2}\Tableau_Emploi_Ensemble_theme9_V2.xlsx", sheet("modes de transport principaux") modify


/* Mise en forme */
putexcel B4 = matrix(RESU), colnames  nformat(number_d2)
putexcel A5 = matrix(RESU), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Répartition des modes de transport principaux utilisés selon les caractéristiques"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "6.	Répartition des modes de transport principaux utilisés"
putexcel (B3:E3), merge

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Demographiques"
putexcel (A3:A4), merge

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close





























*****  1. Caractéristiques générales des trajets domicile-travail

*** Temps moyen de trajet domicile-travail (en minutes)

/*

Nettoyer les valeurs aberrantes de EP2h et EP2i, en tenant compte du contexte logique :

Si la personne travaille dans sa localité (EP2e == 1), on s'attend à des temps et distances faibles.

Si la personne travaille hors localité (EP2e == 2), les valeurs peuvent être plus élevées.

Le moyen de transport (EP2g) permet de fixer des seuils plausibles par mode de déplacement.


*/
* 1. Conversion de la distance en kilomètres
gen EP2i_km = EP2i if trimestre=="T2"

replace EP2i_km = EP2i if inlist(trimestre,"T3","T4")

replace EP2i_km = EP2i / 1000 if EP2i_unit == 2 & trimestre=="T2" & EP2i!=9998

replace EP2i_km = EP2i / 1000 if EP2i_unit == 2 & inlist(trimestre,"T3","T4")

summarize EP2i_km if trimestre=="T2" & EP2i_km!=9998, detail

/*
* 1. Créer variable seuil_temps vide pour le trimestre T2
gen seuil_temps = . if trimestre == "T2"

* 2. Appliquer les moyennes conditionnelles
foreach g in ///
    "EP2e==1 & inlist(EP2g, 1, 11, 12, 13, 17) " ///
    "EP2e==2 & inlist(EP2g, 1, 11, 12, 13, 17)" ///
    "EP2e==1 & inlist(EP2g, 2, 3)" ///
    "EP2e==2 & inlist(EP2g, 2, 3)" ///
    "EP2e==1 & inlist(EP2g, 4, 5, 6, 14)" ///
    "EP2e==2 & inlist(EP2g, 4, 5, 6, 14)" ///
    "EP2e==1 & inlist(EP2g, 15)" ///
    "EP2e==2 & inlist(EP2g, 15)" ///
    "EP2e==1 & inlist(EP2g, 8, 9, 10)" ///
    "EP2e==2 & inlist(EP2g, 8, 9, 10)" {
    
    * Créer une variable temporaire pour stocker la moyenne
    gen double moy = .
    quietly {
        * Calculer la moyenne pour la sous-population
        su EP2h if `g' & trimestre == "T2", meanonly
        replace moy = r(mean) if `g' & trimestre == "T2"
    }
    
    * Appliquer la valeur à seuil_temps
    replace seuil_temps = moy if `g' & trimestre == "T2"
    
    * Supprimer la variable temporaire
    drop moy
}
*/

/*

* Initialiser la variable seuil_km pour le trimestre T2
gen seuil_km = . if trimestre == "T2"

* Appliquer les moyennes conditionnelles de EP2i à seuil_km
foreach cond in ///
    "EP2e==1 & inlist(EP2g, 1, 11, 12, 13, 17)" ///
    "EP2e==2 & inlist(EP2g, 1, 11, 12, 13, 17)" ///
    "EP2e==1 & inlist(EP2g, 2, 3)" ///
    "EP2e==2 & inlist(EP2g, 2, 3)" ///
    "EP2e==1 & inlist(EP2g, 4, 5, 6, 14)" ///
    "EP2e==2 & inlist(EP2g, 4, 5, 6, 14)" ///
    "EP2e==1 & inlist(EP2g, 15)" ///
    "EP2e==2 & inlist(EP2g, 15)" ///
    "EP2e==1 & inlist(EP2g, 8, 9, 10)" ///
    "EP2e==2 & inlist(EP2g, 8, 9, 10)" {

    * Créer une variable temporaire pour la moyenne
    gen double moy_km = .

    quietly {
        * Calculer la moyenne de EP2i pour la condition et le trimestre T2
        su EP2i if `cond' & trimestre == "T2", meanonly
        replace moy_km = r(mean) if `cond' & trimestre == "T2"
    }

    * Appliquer cette moyenne à seuil_km
    replace seuil_km = moy_km if `cond' & trimestre == "T2"

    * Supprimer la variable temporaire
    drop moy_km
}

*/




* 5. Création des moyennes de remplacement pour imputation


gen EP2h_filtre = EP2h if trimestre=="T2" 
*& EP2h<90 & EP2h!=9998 & EP2h>10 
bysort branche1 EP2e HH6 EP2g : egen moy_EP2h = mean(EP2h_filtre) if (trimestre=="T2" & EP2h<180 & EP2h!=9998 & EP2h<. )
*& EP2h>3 )

* 6. Imputation des valeurs extrêmes
replace EP2h = moy_EP2h if (trimestre=="T2" & 180<=EP2h & EP2h!=9998 & EP2h<.  ) 
*| (10>EP2h)


* 6. Imputation des valeurs extrêmes

* 7. Reconversion de la distance imputée (si unité = mètre)
*replace EP2i = EP2i_km * 1000 if EP2i_unit == 2 & extrem_dist == 1 & trimestre=="T2"
*replace EP2i = EP2i_km if EP2i_unit == 1 & extrem_dist == 1 & trimestre=="T2"

* 8. Création d'indicateurs d'imputation
*gen imput_EP2h = extrem_temps if  trimestre=="T2"
*gen imput_EP2i = extrem_dist if  trimestre=="T2"

/*
* 5. Création des valeurs de référence (médianes) pour imputation
egen ref_EP2h = median(EP2h) if extrem_temps == 0, by(EP2e EP2g)
egen ref_EP2i_km = median(EP2i_km) if extrem_dist == 0, by(EP2e EP2g)

* 6. Imputation des valeurs extrêmes
replace EP2h = ref_EP2h if extrem_temps == 1
replace EP2i_km = ref_EP2i_km if extrem_dist == 1

* 7. Reconversion de la distance imputée (si unité = mètre)
*replace EP2i = EP2i_km * 1000 if EP2i_unit == 2 & extrem_dist == 1
*replace EP2i = EP2i_km if EP2i_unit == 1 & extrem_dist == 1

* 8. Création d'indicateurs d'imputation
gen imput_EP2h = extrem_temps
gen imput_EP2i = extrem_dist
*/


/*
Thème 9 : Mobilité domicile-travail (MARTIAL)
1. Caractéristiques générales des trajets domicile-travail

*Temps moyen de trajet domicile-travail (en minutes)
*/

*"T3"
*clear matrix
mat define RESU = (.,.)
/* Milieu de résidence */
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2h<9998, over(milieu_resid2)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..3])',(r(table)[rownumb(r(table),"se"),1..3])')

/* Sexe */
mat RESU = RESU \ (.,.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2h<9998, over(sexe)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..2])',(r(table)[rownumb(r(table),"se"),1..2])')

/*Groupe d'Age */
mat RESU = RESU \ (.,.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2h<9998, over(groupe_age4)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"se"),1..4])')

/* Niveau d'Instruction */
mat RESU = RESU \ (.,.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2h<9998, over(Niv_inst_AG3)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"se"),1..5])')


/* Statut de l'emploi */
mat RESU = RESU \ (.,.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16   & EP2h<9998, over(CISE_18_new)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..5])',(r(table)[rownumb(r(table),"se"),1..5])')

/*Secteur d'activité */
mat RESU = RESU \ (.,.) 
mean EP2h [pw=pmencor_ind_annuel] if MO==1 & age >= 16  & EP2h<9998, over(branche1)
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1..4])',(r(table)[rownumb(r(table),"se"),1..4])')


/* Ensemble */
mean EP2h [pw=pmencor_ind_annuel] if  MO==1 & age >= 16 & EP2h<9998
mat list r(table)
mat RESU = RESU \ ((r(table)[rownumb(r(table),"b"), 1])',(r(table)[rownumb(r(table),"se"),1])')



mat RESU = RESU[1..30,1], RESU[1..30,2]



*A.2 Mise en forme du Tableau 
		/*---------------------------*/
		
/* Définition des entête de lignes et colonnes */

*Lignes 

matrix rownames RESU = "Milieu de Residence" "Abidjan" "Autre Urbain" "Rural" "Sexe" "Masculin"  "Feminin" "Groupe d'Age" "16-24 ans" "25-35 ans" "36-64 ans" "65 ans et plus"  "Niveau d'Instruction"  "Aucun" "Primaire" "Secondaire  1er Cycle" "Secondaire  2nd Cycle" "Superieure"   "Statut de l'emploi" "Employeur"  "worker indépendants unemployees" "Non-salariés (Entrep) dépendants"  "Employés"  "Travailleurs familiaux" "Secteur d'activité" "Agriculture"  "Industrie"  "Commerce" "Autres services"  "Ensemble"
*Colonnes 

matrix colnames RESU = "Temps moyen"  "Std temps moyen"


/* Exportation sur Excel dans le dossier Resultats_Tab*/

putexcel set "${Resultats_Tab2}\Tableau_Emploi_Ensemble_theme9_V2.xlsx", sheet("Temps moyen de trajet_") modify


/* Mise en forme */
putexcel B4 = matrix(RESU), colnames  nformat(number_d2)
putexcel A5 = matrix(RESU), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Répartition du Temps moyen de trajet domicile-travail (en minutes) selon les caractéristiques"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Temps moyen de trajet domicile-travail (en minutes)"
putexcel (B3:E3), merge

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Demographiques"
putexcel (A3:A4), merge

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close



* proportion Mode de transport dominant selon la zone géographique

*profil colonnes
*clear matrix
mat define RESU = J(33,8,.)

proportion HH2 [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(EP2g1)
mat list r(table)
mat RESU = (r(table)[rownumb(r(table),"b"),1..8] \ r(table)[rownumb(r(table),"b"),9..16] \ r(table)[rownumb(r(table),"b"),17..24] \ r(table)[rownumb(r(table),"b"),25..32] \ r(table)[rownumb(r(table),"b"),33..40] \ r(table)[rownumb(r(table),"b"),41..48] \ r(table)[rownumb(r(table),"b"),49..56] \ r(table)[rownumb(r(table),"b"),57..64] \ r(table)[rownumb(r(table),"b"),65..72] \ r(table)[rownumb(r(table),"b"),73..80] \ r(table)[rownumb(r(table),"b"),81..88] \ r(table)[rownumb(r(table),"b"),89..96] \ r(table)[rownumb(r(table),"b"),97..104] \ r(table)[rownumb(r(table),"b"),105..112] \ r(table)[rownumb(r(table),"b"),113..120] \ r(table)[rownumb(r(table),"b"),121..128] \ r(table)[rownumb(r(table),"b"),129..136] \ r(table)[rownumb(r(table),"b"),137..144] \ r(table)[rownumb(r(table),"b"),145..152] \ r(table)[rownumb(r(table),"b"),153..160] \ r(table)[rownumb(r(table),"b"),161..168] \ r(table)[rownumb(r(table),"b"),169..176] \ r(table)[rownumb(r(table),"b"),177..184] \ r(table)[rownumb(r(table),"b"),185..192] \ r(table)[rownumb(r(table),"b"),193..200] \ r(table)[rownumb(r(table),"b"),201..208] \ r(table)[rownumb(r(table),"b"),209..216] \ r(table)[rownumb(r(table),"b"),217..224] \ r(table)[rownumb(r(table),"b"),225..232] \ r(table)[rownumb(r(table),"b"),233..240] \ r(table)[rownumb(r(table),"b"),241..248] \ r(table)[rownumb(r(table),"b"),249..256] \ r(table)[rownumb(r(table),"b"),257..264]) *100


*A.2 Mise en forme du Tableau 
		/*---------------------------*/
		
/* Définition des entête de lignes et colonnes */

*Lignes 

matrix rownames RESU = "ABIDJAN" "YAMOUSSOUKRO" "SAN-PEDRO" "GBOKLE " "NAWA " "INDENIE-DJUABLIN "  "SUD-COMOE" "KABADOUGOU " "FOLON" "LÔH-DJIBOUA " "GÔH " "HAUT-SASSANDRA "  "MARAHOUE"  "N'ZI" "BELIER" "IFFOU" "MORONOU " "AGNEBY-TIASSA "   "GRAND-PONTS" "LA ME"  "TONKPI" "CAVALLY"  "GUEMON"  "PORO"  "BAGOUE" "TCHOLOGO" "GBEKE" "HAMBOL" "WORODOUGOU " "BAFING" "BERE" "GONTOUGO" "BOUNKANI"

*Colonnes 

matrix colnames RESU = "Transport par l'employeur"  "Transport en commun public"  "Transport en commun privée"  "Transport personnel"  "A pieds"  "Travailleur à domicile"  "Vendeur ambulant"


/* Exportation sur Excel dans le dossier Resultats_Tab*/

putexcel set "${Resultats_Tab2}\Tableau_Emploi_Ensemble_theme9_V2.xlsx", sheet("Transport dominant Region_ligne") modify


/* Mise en forme */
putexcel B4 = matrix(RESU), colnames  
putexcel A5 = matrix(RESU), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Repartition du transport par Region"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Repartition du transport "
putexcel (B3:E3), merge

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Demographiques"
putexcel (A3:A4), merge

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close

*pour Region en colonnes

mat define RESU = J(33,8,.)

proportion EP2g1 [pw=pmencor_ind_annuel] if MO==1 & age >= 16 , over(HH2)
mat list r(table)
mat RESU = (r(table)[rownumb(r(table),"b"),1..33] \ r(table)[rownumb(r(table),"b"),34..66] \ r(table)[rownumb(r(table),"b"),67..99] \ r(table)[rownumb(r(table),"b"),100..132] \ r(table)[rownumb(r(table),"b"),133..165] \ r(table)[rownumb(r(table),"b"),166..198] \ r(table)[rownumb(r(table),"b"),199..231] \ r(table)[rownumb(r(table),"b"),232..264]) * 100 


*A.2 Mise en forme du Tableau 
		/*---------------------------*/
		
/* Définition des entête de lignes et colonnes */

*Lignes 

matrix colnames RESU = "ABIDJAN" "YAMOUSSOUKRO " "SAN-PEDRO" "GBOKLE " "NAWA " "INDENIE-DJUABLIN "  "SUD-COMOE" "KABADOUGOU " "FOLON" "LÔH-DJIBOUA " "GÔH " "HAUT-SASSANDRA "  "MARAHOUE"  "N'ZI" "BELIER" "IFFOU" "MORONOU " "AGNEBY-TIASSA "   "GRAND-PONTS" "LA ME"  "TONKPI" "CAVALLY"  "GUEMON"  "PORO"  "BAGOUE" "TCHOLOGO" "GBEKE" "HAMBOL" "WORODOUGOU " "BAFING" "BERE" "GONTOUGO" "BOUNKANI"

*Colonnes 

matrix rownames RESU = "Transport par l'employeur"  "Transport en commun public"  "Transport en commun privée"  "Transport personnel"  "A pieds"  "Travailleur à domicile"  "Vendeur ambulant"


/* Exportation sur Excel dans le dossier Resultats_Tab*/

putexcel set "${Resultats_Tab2}\Tableau_Emploi_Ensemble_theme9_V2.xlsx", sheet("Transport dominant Region_col") modify


/* Mise en forme */
putexcel B4 = matrix(RESU), colnames  
putexcel A5 = matrix(RESU), rownames

/* Titre du tableau */
putexcel B1 = "Tableau : Repartition du transport par Region"
putexcel B1, bold border(bottom)

*En tête colonne du Tableau
putexcel B3 = "Repartition du transport "
putexcel (B3:E3), merge

*En tête ligne du Tableau
putexcel A3 = "Caractéristiques Socio Demographiques"
putexcel (A3:A4), merge

*Sauvegarde définitive du Tableau
putexcel save

*Fermeture du fichier
putexcel close





* Pression temporelle de la mobilité selon les caractéristique socio-démographique et Temps de trajet

***** La pression moyenne est calculer pour les individus en emploi
cap drop pression_moyenne
gen pression_moyenne = EP2h / temps_tra_EP_semaine if EN_EMP==1





do "$dofile_theme\Implication_mobilite_travail_Emploi.do"






















