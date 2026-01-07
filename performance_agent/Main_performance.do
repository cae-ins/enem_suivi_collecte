* -----------------------------
* 1) Définir les dossiers sources
* -----------------------------

global wd "D:\ENEM_Working\Apurement salaire\Dossier_travail_Toure\"

*******************************************************
* process_trimestres.do
* Parcourt la liste des trimestres, applique les
* transformations sur HH13 -> HH13_es et sauvegarde.
* ATTENTION : Vérifier les globals de chemins avant exécution.
*******************************************************

do "$wd\Transformation_HH13.do"


********************************************************************************
* PROGRAMME : Calcul des durées d'administration des questionnaires ENEM
* OBJECTIF  : Calculer les temps moyens d'administration par agent et trimestre
* VERSION   : 3.1 - Correction gestion variables manquantes
********************************************************************************

do "$wd\Indicateur_performance_AG_claude2.do"



