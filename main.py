#Pour demarrer je charge l'excel

import pandas as pd

# Demander le nom du fichier
nom_fichier = input("Entrez le nom du fichier  : ")
print("Nom du fichier d'entrée =", nom_fichier)

#charger les donnees feuille  dans un dataframe Pandas
#chargement premier onglet = Egogramme`
df_excel = pd.read_excel(nom_fichier, sheet_name='Egogramme', usecols='C', skiprows=6, nrows=61)

#conversion du DataFrame excel en liste de tuple

reponse_data = [(reponse,) for reponse in df_excel['Note']]

#affichage de la liste de tuple

print("reponse =",reponse_data)
ligne_0 = reponse_data [0][0]
ligne_1 = reponse_data [1][0]
ligne_2 = reponse_data [2][0]
ligne_3 = reponse_data [3][0]
ligne_4 = reponse_data [4][0]
ligne_5 = reponse_data [5][0]
ligne_6 = reponse_data [6][0]
ligne_7 = reponse_data [7][0]
ligne_8 = reponse_data [8][0]
ligne_9 = reponse_data [9][0]
ligne_10 = reponse_data [10][0]
ligne_11 = reponse_data [11][0]
ligne_12 = reponse_data [12][0]
ligne_13 = reponse_data [13][0]
ligne_14 = reponse_data [14][0]
ligne_15 = reponse_data [15][0]
ligne_16 = reponse_data [16][0]
ligne_17 = reponse_data [17][0]
ligne_18 = reponse_data [18][0]
ligne_19 = reponse_data [19][0]
ligne_20 = reponse_data [20][0]
ligne_21 = reponse_data [21][0]
ligne_22 = reponse_data [22][0]
ligne_23 = reponse_data [23][0]
ligne_24 = reponse_data [24][0]
ligne_25 = reponse_data [25][0]
ligne_26 = reponse_data [26][0]
ligne_27 = reponse_data [27][0]
ligne_28 = reponse_data [28][0]
ligne_29 = reponse_data [29][0]
ligne_30 = reponse_data [30][0]
ligne_31 = reponse_data [31][0]
ligne_32 = reponse_data [32][0]
ligne_33 = reponse_data [33][0]
ligne_34 = reponse_data [34][0]
ligne_35 = reponse_data [35][0]
ligne_36 = reponse_data [36][0]
ligne_37 = reponse_data [37][0]
ligne_38 = reponse_data [38][0]
ligne_39 = reponse_data [39][0]
ligne_40 = reponse_data [40][0]
ligne_41 = reponse_data [41][0]
ligne_42 = reponse_data [42][0]
ligne_43 = reponse_data [43][0]
ligne_44 = reponse_data [44][0]
ligne_45 = reponse_data [45][0]
ligne_46 = reponse_data [46][0]
ligne_47 = reponse_data [47][0]
ligne_48 = reponse_data [48][0]
ligne_49 = reponse_data [49][0]
ligne_50 = reponse_data [50][0]
ligne_51 = reponse_data [51][0]
ligne_52 = reponse_data [52][0]
ligne_53 = reponse_data [53][0]
ligne_54 = reponse_data [54][0]
ligne_55 = reponse_data [55][0]
ligne_56 = reponse_data [56][0]
ligne_57 = reponse_data [57][0]
ligne_58 = reponse_data [58][0]
ligne_59 = reponse_data [59][0]

#calcul des Egogramme

parent_nouricier  =  ligne_3 + ligne_7 + ligne_13 + ligne_21 + ligne_27 + ligne_32 + ligne_35 + ligne_49 + ligne_56 + ligne_59
print("parent nouricier =",parent_nouricier)

parent_normatif  =  ligne_5 + ligne_11 + ligne_15 + ligne_23 + ligne_25 + ligne_31 + ligne_36 + ligne_47 + ligne_51 + ligne_55
print("parent normatif =",parent_normatif)

adulte  =  ligne_0 + ligne_10 + ligne_16 + ligne_18 + ligne_26 + ligne_28 + ligne_37 + ligne_41 + ligne_43 + ligne_53
print("adulte =",adulte)

enfant_libre  =  ligne_4 + ligne_6 + ligne_14 + ligne_22 + ligne_24 + ligne_40 + ligne_42 + ligne_46 + ligne_52 + ligne_58
print("enfant libre =",enfant_libre)

enfant_adapte_soumis  =  ligne_2 + ligne_8 + ligne_12 + ligne_20 + ligne_30 + ligne_33 + ligne_39 + ligne_45 + ligne_50 + ligne_57
print("enfant adapte soumis =",enfant_adapte_soumis)

enfant_adapte_rebelle  =  ligne_1 + ligne_9 + ligne_17 + ligne_19 + ligne_29 + ligne_34 + ligne_38 + ligne_44 + ligne_48 + ligne_54
print("enfant adapte rebelle =",enfant_adapte_rebelle)




