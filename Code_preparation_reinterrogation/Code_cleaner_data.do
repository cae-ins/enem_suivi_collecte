
global wd_T2_2024 "D:\ENEM_Working\Base_prechargement_ENEM\Base_brute_T2_2024\"

global wd_T3_2024 "D:\ENEM_Working\Base_prechargement_ENEM\Base_brute_T3_2024\"

global wd_T4_2024 "D:\ENEM_Working\Base_prechargement_ENEM\Base_brute_T4_2024\"

global wd_T1_2025 "D:\ENEM_Working\Base_prechargement_ENEM\Base_brute_T1_2025\"

global wd_T2_2025 "D:\ENEM_Working\Base_prechargement_ENEM\Base_brute_T2_2025\"

global wd_T3_2025 "D:\ENEM_Working\Base_prechargement_ENEM\Base_brute_T3_2025\"

global wd_T4_2025 "D:\ENEM_Working\Base_prechargement_ENEM\Base_brute_T4_2025\"



***** T2-2024


use "$wd_T2_2024\ENEM_2024T2.dta", clear

drop if interview__key == "22-01-05-49"
drop if interview__key == "35-53-67-01"
drop if interview__key == "68-67-94-64"
drop if interview__key == "72-37-40-16"
drop if interview__key == "77-98-79-25"
drop if interview__key == "04-99-87-31"
drop if interview__key == "22-77-86-77"
drop if interview__key == "06-22-03-87"
drop if interview__key == "90-18-87-78"
drop if interview__key == "93-27-08-63"
drop if interview__key == "64-26-43-89"
drop if interview__key == "83-97-01-40"
drop if interview__key == "34-00-83-22"
drop if interview__key == "07-68-43-81"
drop if interview__key == "13-47-97-94"
drop if interview__key == "21-40-72-82"
drop if interview__key == "31-15-47-16"
drop if interview__key == "87-43-35-93"
drop if interview__key == "31-35-48-19"
drop if interview__key == "49-09-93-51"
drop if interview__key == "89-89-22-16"
drop if interview__key == "86-33-67-81"
drop if interview__key == "56-61-52-99"
drop if interview__key == "79-93-49-22"
drop if interview__key == "05-51-15-90"

count if !inlist(HH14,3,4,6,7,8)

keep if !inlist(HH14,3,4,6,7,8)

**** Entretient entirement vide du passage 1
cap drop vide
gen vide=1 if  inlist(M0__0,"") & inlist(HH14,.)
replace vide =1 if inlist(M0__0,"","##N/A##")
keep if !inlist(vide,1)


save "$wd_T2_2024\ENEM_2024T2.dta", replace


*gen zd=1
* Calculer le nombre de ménages par région et par zone de dénombrement
*collapse (sum) zd, by(HH2 HH8A HH8B HH8)
 
 
/*

HH14

  6 Refus
  7 logement du ménage détruit
  8 Logement vacant

*/

***** T3-2024

*** Au T2-2024, pour la region INDENIE-DJUABLIN,  ZD=4002  est non visité
*** Elle est visité au T3-2024

*** Au T2-2024, pour la region KABADOUGOU,  ZD=6006, quartier (HH8B)=NIENESSO NAFANA, est non visité
*** Elle est visité au T3-2024;
*** Aussi au T3-2024 pour la region KABADOUGOU,  ZD=6006, quartier(HH8B)=MASSADOUGOU, est visité au T3-2024

*** Au T2-2024, pour la region FOLON,  ZD=6006, quartier (HH8B)=KARALA, est non visité. Elle est visité au T3-2024;
*** Au T2-2024, pour la region FOLON,  ZD=6006, quartier (HH8B)=MAHANDIANA-SOBALA, est non visité. Elle est visité au T3-2024;

*** Au T2-2024, pour la region TONKPI,  ZD=0006, quartier (HH8B)=ZOKOUAVILLE, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region TONKPI,  ZD=0056, quartier (HH8B)=LIBREVILLE, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region TONKPI,  ZD=6029, quartier (HH8B)=BOUNTA, est non visité.*** Elle est visité au T3-2024;

*** Au T2-2024, pour la region CAVALLY,  ZD=6019, quartier (HH8B)=SOUBRE, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region CAVALLY,  ZD=0057, quartier (HH8B)=HOUPHOUET BOIGNY 2, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region TONKPI,  ZD=0035, quartier (HH8B)=CIB, est non visité.*** Elle est visité au T3-2024;

*** Au T2-2024, pour la region GUEMON,  ZD=0011, quartier (HH8B)=RESIDENTIEL EXTENSION 2, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region GUEMON,  ZD=6009, quartier (HH8B)=BAMBOIBLY, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region GUEMON,  ZD=6049, quartier (HH8B)=REMIKRO, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region GUEMON,  ZD=0063, quartier (HH8B)=RESIDENTIEL  EXTENSION  1, est non visité.*** Elle est visité au T3-2024;


*** Au T2-2024, pour la region PORO,  ZD=0176, quartier (HH8B)=DELAFOSSE, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region PORO,  ZD=6003, quartier (HH8B)=LOUHOUA, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region PORO,  ZD=0255, quartier (HH8B)=LATONON, est non visité.*** Elle est visité au T3-2024;

*** Au T2-2024, pour la region BAGOUE,  ZD=0016, quartier (HH8B)=BELE, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region BAGOUE,  ZD=0044, quartier (HH8B)=TIOGONA, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region BAGOUE,  ZD=0012, quartier (HH8B)=GBON 5, est non visité.*** Elle est visité au T3-2024;


*** Au T2-2024, pour la region TCHOLOGO,  ZD=6044, quartier (HH8B)=TCHASSANAKAHA, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region TCHOLOGO,  ZD=0018, quartier (HH8B)=MIGARBAVOGO, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region TCHOLOGO,  ZD=0079, quartier (HH8B)=BROMAKOTE, est non visité.*** Elle est visité au T3-2024;


*** Au T2-2024, pour la region GBEKE ,  ZD=0599, quartier (HH8B)=OLLIENOU, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region GBEKE ,  ZD=0109, quartier (HH8B)=AHOUGNANSOU, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region GBEKE ,  ZD=0079, quartier (HH8B)=BROMAKOTE, est non visité.*** Elle est visité au T3-2024;


*** Au T2-2024, pour la region HAMBOL ,  ZD=0020, quartier (HH8B)=SOUROUKAHA, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region HAMBOL ,  ZD=0024, quartier (HH8B)=DABAKALAKRO, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region HAMBOL ,  ZD=6007, quartier (HH8B)=SEPIKAHA, est non visité.*** Elle est visité au T3-2024;

*** Au T2-2024, pour la region WORODOUGOU ,  ZD=6002, quartier (HH8B)=KOUMBARA, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region WORODOUGOU ,  ZD=6026, quartier (HH8B)=GBEMAZO, est non visité.*** Elle est visité au T3-2024;
*** Au T2-2024, pour la region WORODOUGOU ,  ZD=6014, quartier (HH8B)=ANTOINEKRO, est non visité.*** Elle est visité au T3-2024;

*11319 - je n'ai pas encore tout lister


use "$wd_T3_2024\ENEM_2024T3.dta", clear

keep if rgmen==1

drop if interview__key == "75-88-05-97"
drop if interview__key == "47-65-63-23"
drop if interview__key == "75-10-93-53"
drop if interview__key == "97-32-98-98"
drop if interview__key == "66-72-62-33"
drop if interview__key == "53-49-14-58"
drop if interview__key == "46-41-06-69"
drop if interview__key == "64-26-43-89"
drop if interview__key == "19-37-71-08"
drop if interview__key == "53-72-60-41"
drop if interview__key == "32-60-74-42"
drop if interview__key == "20-41-58-21"
drop if interview__key == "14-19-82-10"
drop if interview__key == "83-97-01-40"
drop if interview__key == "14-74-91-35"
drop if interview__key == "70-28-54-68"
drop if interview__key == "90-18-87-78"
drop if interview__key == "75-85-61-73"
drop if interview__key == "19-22-38-04"
drop if interview__key == "34-88-03-65"
drop if interview__key == "65-46-97-23"
drop if interview__key == "01-87-11-19"
drop if interview__key == "25-12-47-81"
drop if interview__key == "54-00-78-12"
drop if interview__key == "18-06-14-19"
drop if interview__key == "66-20-52-42"
drop if interview__key == "39-30-11-85"
drop if interview__key == "40-48-28-05"
drop if interview__key == "70-52-47-50"
drop if interview__key == "05-56-25-26"
drop if interview__key == "13-90-70-89"
drop if interview__key == "21-30-44-25"
drop if interview__key == "40-00-43-07"
drop if interview__key == "48-51-68-52"
drop if interview__key == "91-27-28-52"
drop if interview__key == "52-32-16-79"


count if !inlist(HH14,3,4,6,7,8)

keep if !inlist(HH14,3,4,6,7,8)

**** Entretient entirement vide du passage 1
cap drop vide
gen vide=1 if  inlist(M0__0,"") & inlist(HH14,.)
replace vide =1 if inlist(M0__0,"","##N/A##")

keep if !inlist(vide,1)



save  "$wd_T3_2024\ENEM_2024T3.dta", replace
*gen zd=1
* Calculer le nombre de ménages par région et par zone de dénombrement
*collapse (sum) zd, by(HH2 HH8A HH8B HH8)

***** T4-2024
use "$wd_T4_2024\ENEM_2024T4.dta", clear

keep if rgmen==1

drop if interview__key == "22-19-50-04"
drop if interview__key == "71-23-05-43"
drop if interview__key == "72-27-67-82"
drop if interview__key == "67-99-64-56"
drop if interview__key == "81-33-30-18"
drop if interview__key == "63-24-96-44"
drop if interview__key == "12-93-14-47"
drop if interview__key == "17-26-57-50"
drop if interview__key == "87-75-07-58"
drop if interview__key == "12-61-59-30"
drop if interview__key == "98-99-29-90"
drop if interview__key == "12-83-32-35"
drop if interview__key == "69-30-53-46"
drop if interview__key == "53-75-03-40"
drop if interview__key == "48-29-91-67"
drop if interview__key == "14-14-95-20"
drop if interview__key == "98-56-21-21"
drop if interview__key == "00-75-73-50"
drop if interview__key == "72-37-75-88"
drop if interview__key == "89-80-42-60"
drop if interview__key == "96-37-49-58"
drop if interview__key == "27-86-38-53"
drop if interview__key == "93-32-36-58"
drop if interview__key == "40-60-68-14"
drop if interview__key == "01-44-82-41"
drop if interview__key == "90-21-62-71"
drop if interview__key == "29-87-22-58"


count if !inlist(HH14,3,4,6,7,8)

keep if !inlist(HH14,3,4,6,7,8)

**** Entretient entirement vide du passage 1
cap drop vide
gen vide=1 if  inlist(M0__0,"") & inlist(HH14,.)
replace vide =1 if inlist(M0__0,"","##N/A##")

keep if !inlist(vide,1)


save "$wd_T4_2024\ENEM_2024T4.dta", replace

*gen zd=1
* Calculer le nombre de ménages par région et par zone de dénombrement
*collapse (sum) zd, by(HH2 HH8A HH8B HH8)


***** T1-2025

use "$wd_T1_2025\ENEM_2025T1.dta", clear

keep if rgmen==1

drop if interview__key == "46-21-18-35"
drop if interview__key == "14-93-48-48"
drop if interview__key == "20-41-48-31"
drop if interview__key == "07-35-83-69"
drop if interview__key == "50-09-81-76"
drop if interview__key == "11-39-00-66"
drop if interview__key == "16-32-20-24"
drop if interview__key == "38-68-65-72"
drop if interview__key == "45-75-38-94"
drop if interview__key == "50-08-89-16"
drop if interview__key == "57-27-38-76"
drop if interview__key == "80-61-75-69"
drop if interview__key == "90-28-88-43"
drop if interview__key == "04-25-36-82"
drop if interview__key == "07-88-35-49"
drop if interview__key == "08-74-28-44"
drop if interview__key == "09-64-25-80"
drop if interview__key == "10-48-23-98"
drop if interview__key == "16-92-94-41"
drop if interview__key == "36-69-62-54"
drop if interview__key == "43-67-91-91"
drop if interview__key == "47-17-94-11"
drop if interview__key == "51-71-17-14"
drop if interview__key == "52-66-31-42"
drop if interview__key == "58-11-91-61"
drop if interview__key == "61-72-49-50"
drop if interview__key == "66-45-79-47"
drop if interview__key == "71-50-89-95"
drop if interview__key == "98-90-17-71"


count if !inlist(HH14,3,4,6,7,8)

keep if !inlist(HH14,3,4,6,7,8)

**** Entretient entirement vide du passage 1
cap drop vide
gen vide=1 if  inlist(M0__0,"") & inlist(HH14,.)
replace vide =1 if inlist(M0__0,"","##N/A##")

keep if !inlist(vide,1)

save "$wd_T1_2025\ENEM_2025T1.dta", replace

***** T2-2025

use "$wd_T2_2025\ENEM_2025T2.dta", clear

keep if rgmen==1

drop if interview__key == "17-15-06-43"
drop if interview__key == "76-83-99-28"
drop if interview__key == "60-88-86-77"
drop if interview__key == "15-08-07-34"
drop if interview__key == "00-60-72-30"
drop if interview__key == "96-20-02-94"

count if !inlist(HH14,3,4,6,7,8)

keep if !inlist(HH14,3,4,6,7,8)

**** Entretient entirement vide du passage 1
cap drop vide
gen vide=1 if  inlist(M0__0,"") & inlist(HH14,.)
replace vide =1 if inlist(M0__0,"","##N/A##")

keep if !inlist(vide,1)


*gen zd=1
* Calculer le nombre de ménages par région et par zone de dénombrement
*collapse (sum) zd, by(HH2 HH8A HH8B HH8)

save "$wd_T2_2025\ENEM_2025T2.dta", replace

***** T3-2025

use "$wd_T3_2025\ENEM_2025T3.dta", clear

keep if rgmen==1

drop if interview__key == "75-97-53-34"
drop if interview__key == "59-53-46-55"
drop if interview__key == "50-73-50-30"
drop if interview__key == "19-66-95-48"
drop if interview__key == "23-27-58-03"
drop if interview__key == "37-35-99-90"
drop if interview__key == "70-17-05-68"
drop if interview__key == "71-89-25-61"
drop if interview__key == "96-69-31-38"


count if !inlist(HH14,3,4,6,7,8)

keep if !inlist(HH14,3,4,6,7,8)

**** Entretient entirement vide du passage 1
cap drop vide
gen vide=1 if  inlist(M0__0,"") & inlist(HH14,.)
replace vide =1 if inlist(M0__0,"","##N/A##")

keep if !inlist(vide,1)

save "$wd_T3_2025\ENEM_2025T3.dta", replace


***** T4-2025

use "$wd_T4_2025\ENEM_2025T4.dta", clear

keep if rgmen==1



drop if interview__key == "67-88-26-57"
drop if interview__key == "40-91-52-57"

count if !inlist(HH14,3,4,6,7,8)

keep if !inlist(HH14,3,4,6,7,8)

cap drop vide
gen vide=1 if  inlist(M0__0,"") & inlist(HH14,.)
replace vide =1 if inlist(M0__0,"","##N/A##")

keep if !inlist(vide,1)


save "$wd_T4_2025\ENEM_2025T4.dta", replace





















