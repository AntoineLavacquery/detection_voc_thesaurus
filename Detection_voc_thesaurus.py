import os
import os.path
from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import *
from tkinter import scrolledtext
from tkinter import messagebox
import pandas as pd
from rapidfuzz import fuzz
import operator
import sys
import re
import xml.etree.ElementTree as ET

# Fonction de récupération des XML du thésaurus et transformation en tableau
def thesaurus_to_df():
    global path
    global df_thes
    df_thes = pd.DataFrame() # dataframe du thésaurus

    # Sélection du répertoire
    path = askdirectory(title="Sélection du répertoire")
    if len(path) == 0:
        dir_button.configure(text="...")
    else:
        dir_button.configure(text=path)
    file_list = os.listdir(path)

    # Pour chaque nom de fichier dans le répertoire
    # exemple : merimee_coll.mcc.xml puis merimee_couv.mcc.xml puis etc.
    for file_name in file_list:
        # Ouverture du fichier à partir de son nom (en tenant compte d'éventuelles erreurs)
        with open(f"{path}/{file_name}") as f:
            file = f.read()
            # Impossible d'explorer l'arborescence si le fichier n'est pas un XML
            try:
                root = ET.fromstring(file)
            except:
                dir_button.configure(text="Erreur")
                messagebox.showerror('Erreur importation thésaurus', 'Merci de vérifier que le répertoire ne contient que les fichiers XML du thésaurus')
                return

        # Récupération du nom du champ/fichier XML qui servira de nom de colonne à partir d'un champ de l'XML
        # exemple : DENO, REPR, JDAT, etc.
        title = root.find('.//{http://purl.org/dc/elements/1.1}title').text
        # Nom du champ toujours constitué de 3 ou 4 lettres en MAJUSCULE
        try:
            col_name = re.search('[A-Z][A-Z][A-Z][A-Z]', title).group()
        except:
            col_name = re.search('[A-Z][A-Z][A-Z]', title).group()
        
        term_list = [] # list des termes récupérés à partir du XML
        termEM_list = [] # liste des dictionnaires (key: terme à employer, values: termes employé pour -> terme à employer (key))

        # Pour chaque élément dans le XML
        for mark in root.iter('{http://www.culture.gouv.fr/ns/mcc/lex/1.0}term'):
            forbid = False
            typeEM = False
            dictEM = {} # dictionnaire de constitution des termes avec relation EP et EM
            
            # On exclu les noeuds interdits "forbid"
            if "forbid" in mark.attrib:
                forbid = True
            
            # On exclu les noeuds enfants avec relation "employer" ("EM")
            for child in mark:
                if "EM" in child.attrib.values():
                    typeEM = True

            # Si aucun des 2 cas vus plus haut, on accepte le terme
            if not (forbid or typeEM):
                # Constitution de la liste de termes
                term = mark.attrib["lib"]
                term_list.append(term)

                # Parmi les noeuds enfants, on créé/agrandit un dictionnaire de valeurs des termes ayant la relation "employé pour" ("EP")
                # La key est le terme à préférer, les values sont les termes employés pour, qui devraient être remplacés
                for child in mark:
                    if "EP" in child.attrib.values():
                        dictEM.setdefault(term, []).append(child.attrib["lib"])

                # Constitution de la liste de dictionnaires des termes à partir des relations "EP" "EM" s'ils existent
                if dictEM:
                    termEM_list.append(dictEM)

        # Constitution du dataframe thésaurus en ajoutant notre colonne ayant pour valeurs les termes du thésaurus
        term_list = list(dict.fromkeys(term_list)) # au cas où, suppression des doublons
        if len(term_list) > len(df_thes): # agrandissement du dataframe au fur et à mesure des ajouts
            df_thes = df_thes.reindex(range(len(term_list)))
        df_thes[col_name] = pd.Series(term_list)

        # Ajout de la colonne des dictionnaires des relations "EP" "EM" si elle existe sous le nom "*nomcolonne*$"
        # exemple : COLL -> COLL$ ; DENO -> DENO$ ; TOIT -> TOIT$ ; etc. Permet de facilement dicriminer ou retrouver la colonne à partir du nom d'origine
        if termEM_list:
            df_thes[f"{col_name}$"] = pd.Series(termEM_list)

# Fonction de sauvegarde du thésaurus
def save_df_thes():
    global df_thes

    # Si df_thes n'existe pas (utilisateur sauvegarde sans avoir chargé le répertoire)
    try:
        # Si df_thes existe mais est vide (l'utilisateur a annulé la sélection du répertoire)
        if not df_thes.empty:
            # On exporte le dataframe vers un fichier CSV dans le chemin choisi par l'utilisateur
            name = asksaveasfile(mode="w", filetypes=[("CSV", ".csv")], defaultextension=".csv")
            df_thes.to_csv(name.name, sep=";", index=True, encoding="ANSI")
            save_thes_button.configure(text=name.name)
        # Sinon on affiche un message d'erreur
        else:
            messagebox.showerror('Erreur de sauvegarde thésaurus', "Merci de sélectionner le répertoire du thésaurus.")
    # On affiche un message d'erreur
    except:
        # save_thes_button.configure(text="")
        messagebox.showerror('Erreur de sauvegarde thésaurus', "Merci de sélectionner le répertoire du thésaurus.")

# Fonction de calcul du score de similitude
def is_prox(word1, word2):
    replaced = False
    digits = [
        ["1", "un"],
        ["2", "deux"],
        ["3", "trois"],
        ["4", "quatre"],
        ["5", "cinq"],
        ["6", "six"],
        ["7", "sept"],
        ["8", "huit"],
        ["9", "neuf"]
        ]
    # Calcul du ratio entre les deux termes
    ratio = fuzz.partial_ratio(word1, word2)

    # Boucle qui remplace les éventuels chiffres et les réinscrit en toute lettres
    for digit in digits:
        if digit[0] in word1:
            word1 = word1.replace(digit[0], digit[1])
            replaced = True

    # S'il y a effectivement eu un remplacement
    if replaced:
        ratio_repl = fuzz.partial_ratio(word1, word2) # calcul d'un nouveau ratio
        if ratio_repl > ratio: # on ne garde que le cas ayant fait le meilleur score
            return ratio_repl
        else:
            return ratio
    else:
        return ratio

# Fonction de conversion d'une chaine de caractères vers une liste
def splitString(string, char):
    newList = (string).split(char)
    newList = [el.strip(" ") for el in newList]
    newList = list(dict.fromkeys(newList))
    return newList

# Fonction de vérification, importante, contient l'essentiel du code
def openTSV():
    # count = 0 # initialisation du compteur de résultat en début de fonction
    txt_box.delete(1.0, END) # on s'assure que la boite de texte est bien vide
    global df # dataframe contenant l'export Cindoc (TSV)
    # global df_thes # voir plus haut, dataframe contenant le thésaurus
    # Ouverture et importation des données dans le dataframe, try:except: au cas où quelque chose se passe mal qui permet de renvoyer l'info à l'utilisateur
    file = askopenfile(mode="rb", filetypes=[("TSV", ".txt"), ("All files", ".*")])
    global file_name
    file_name = os.path.basename(file.name) # name.txt
    try:
        df = pd.read_csv(file.name, sep="\t", encoding="ANSI")
        # txt_box.insert(INSERT, "Modèle\n\n[COLONNE/CHAMP][ligne](REF) : \"terme non trouvé\"\nPROCHE ---> \"terme proche\" (score de proximité) | etc.\n\n")
        file_button.configure(text=f"{file_name}")
    except:
        file_button.configure(text=f"Sélection du fichier TSV à vérifier (*.txt)")
        messagebox.showerror('Erreur d\'importation', sys.exc_info())

# Fonction de comparaison des termes
def checker():
    global df
    global file_name
    count = 0 # initialisation du compteur de résultat en début de fonction
    txt_box.delete(1.0, END) # on s'assure que la boite de texte est bien vide
    txt_box.insert(INSERT, "Modèle\n\n[COLONNE/CHAMP][ligne](REF) : \"terme non trouvé\"\nPROCHE ---> \"terme proche\" (score de proximité) | etc.\n\n")
    # Pour chaque colonne du thésaurus
    for col in df_thes:
        # Constitution de la liste des termes de référence
        # On ignore les colonnes "$", copie des autre colonnes contenant des dict qui contiennent les termes "EP" (key) et "EM" (value)
        if "$" not in col:
            ref_list = df_thes[col].dropna().unique().tolist()
            ref_list = list(dict.fromkeys(ref_list))
        # Initialisation de la liste globale de résultats 
        col_list_results = []
        
        # Si la colonne est dans l'export Cindoc
        if col in df:
            # Inscription l'entête de la colonne concernée dans le log SI SOUHAIT D'INSCRIRE MÊME LES COLONNES N'AYANT PAS D'ERREURS
            # col_list_results.append(f"\nColonne/champ : {col}\n")
            # Pour chaque cellule de la colonne sélectionnée, on constitue une liste de termes à vérifier
            for i in range(len(df)):
                # Vérification qu'il s'agit bien d'une chaîne de caractères
                if not isinstance(df.at[i, col], str):
                    pass
                else:
                    terms_to_verify = splitString(str(df.at[i, col]), ";")

                    # Pour chaque terme de la liste à vérifier
                    for term in terms_to_verify:
                        scores_dict = {} # initialisation dictionnaire de tous les scores -> toutes les combinaisons possibles en les termes
                        other_col = "" # initialisation nom de l'autre colonne qui pourrait contenir le terme testé
                        employ = "" # initialisation nom du terme à employer en lieu de celui détecté (puisé dans les colonnes "$")
                        term_org = ""

                        # Conditions particulières à certaines colonnes
                        # Si colonne = COLL, suppression des chiffres devant le terme (les élements sont dénombrés, vérification uniquement du terme qui suit)
                        # exemple : "12 repérées" -> "repérées"
                        if col == "COLL":
                            term_org = term
                            term = ''.join([i for i in term if not i.isdigit()])
                            term = term.strip()

                        # Si col = DPRO, suppression de la date si elle est au bon format (ainsi les dates au mauvais format apparaitront pour correction)
                        # exemple : "1993/03/25 : classé MH" -> "classé MH"
                        if col == "DPRO":
                            if re.search(r"[0-9]{4}\/[0-9]{2}\/[0-9]{2}", term):
                                term_org = term
                                term = re.sub(r"[0-9]{4}\/[0-9]{2}\/[0-9]{2}", "", term)
                                term = re.sub(r"\:", "", term)
                                term = term.strip()
                        
                        # Si col = REMP, EXEC ou ORIG, suppression de l'élement particulier à la notice
                        # exemple : "remploi provenant de : 67, Bischeim" -> "remploi provenant de"
                        if col == "REMP" or col == "EXEC" or col == "ORIG":
                            term_org = term
                            term = re.sub(r" :.*", "", term)
                            term = term.strip()

                    # Si col = AFIG, ATEL, AUTR ou INSC, suppression de l'élement particulier à la notice
                    # exemple : "Helmstetter (sculpteur)" -> "sculpteur"
                        if col == "AFIG" or col == "ATEL" or col == "AUTR" or col == "INSC":
                            if "(" in term or ")" in term: # si entre parenthèse trouvé
                                term_org = term
                                term = term.replace(")", "") # suppression de la dernière parenthèse
                                term = term.split("(") # liste autour de la parenthèse ouvrante
                                term = term[1] # on prend le dernier élement de la liste
                                # restera à resciender éventuellement si découverte de plusieurs élements et de virgules
                        
                        # Si col = EDIF / PART on vérifie les colonnes REFE et REFP ?

                        # Si le terme n'est pas dans la liste de référence
                        if term not in ref_list:
                            # Boucle exploration des autres colonnes
                            for col2 in df_thes:
                                # Recherche du terme dans les autres colonnes, sauf celles possédant les termes "EM"
                                if col2 != col and "$" not in col2:
                                    ref_list2 = df_thes[col2].dropna().unique().tolist()
                                    ref_list2 = list(dict.fromkeys(ref_list2))
                                    if term in ref_list2:
                                        other_col = col2
                                
                                # Recherche si le terme n'est pas employé pour en désigner un autre à partir des colonnes $ qui contiennent les termes "EP" du thésaurus
                                if col2 == f"{col}$":
                                    for cell in df_thes[col2]:
                                        if isinstance(cell, dict):
                                            for dictvalue in cell.values():
                                                if term in dictvalue:
                                                    employ = next(iter(cell))
                            
                            # Boucle de constitution de tous les scores de similitude entre le terme et tous ceux du thésaurus
                            for ref in ref_list:
                                score = is_prox(term, ref) # Appel de la fonction
                                scores_dict[ref] = score # Inscription de la paire clé (terme du thésaurus), valeur (score)
                            
                            # Ne sont retenus que les 5 plus proches termes
                            five_best_dict = dict(sorted(scores_dict.items(), key=operator.itemgetter(1), reverse=True)[:5])
                            best_guesses = ""
                            for key in five_best_dict: # inscription des termes et de leur score dans une chaine de caractères
                                best_guesses += f"\"{key}\" (score = {str(int(five_best_dict[key]))}) | "
                            
                            if term_org:
                                term_org = f"(dans \"{term_org}\") "

                            # Si le terme est employé pour désigner un autre terme et dans colonne COLL, on passe (cas particulier colonne COLL)
                            if employ != "" and col == "COLL":
                                pass

                            # Si le terme est employé pour désigner un autre terme et dans colonne COLL, on passe (cas particulier colonne COLL)
                            elif col == "MATR" or col == "STRU" or col == "STAD" or col == "ETAT" and ((list(five_best_dict.values()))[0] == 100 or (list(five_best_dict.values()))[1] == 100):
                                pass

                            # Si le terme est un terme incorrect car employé pour désigner un autre terme, on indique directement l'autre terme
                            elif employ != "":
                                result_sentence = f"[{col}][{i+2}]({df['REF'][i]}) : \"{term}\" {term_org}/!\ employer \"{employ}\" (cf. thésaurus)\nChercheur⸱se⸱s : {df['NOMS'][i]}\n\n"
                                count += 1

                            # Si le terme a été trouvé dans un autre colonne, on l'inscrit dans le log
                            elif other_col != "":
                                result_sentence = f"[{col}][{i+2}]({df['REF'][i]}) : \"{term}\" {term_org}/!\ présent dans le thésaurus, colonne [{other_col}]\nPROCHE ---> {best_guesses}\nChercheur⸱se⸱s : {df['NOMS'][i]}\n\n"
                                count += 1
                            # Sinon on inscrit juste les résultats par similitudes
                            else:
                                result_sentence = f"[{col}][{i+2}]({df['REF'][i]}) : \"{term}\" {term_org}\nPROCHE ---> {best_guesses}\nChercheur⸱se⸱s : {df['NOMS'][i]}\n\n"
                                count += 1
                            # Condition pour s'assurer que deux lignes de résultats ne soient pas identique (A VOIR SI PERTINENT/UTILISE)
                            try:
                                if result_sentence not in col_list_results:
                                    col_list_results.append(result_sentence)
                            except:
                                pass
                
            # Retour au niveau de la boucle "pour chaque colonne", affichage des résultats accumulés dans col_list_results
            if col_list_results:
                col_list_results.insert(0, f"Colonne/champ : {col}\n\n")
                for line in col_list_results:
                    txt_box.insert(INSERT, line)
    
    txt_box.insert('1.0', f'Fichier/base : {file_name}\nErreurs : {count}\n\n')
    file_button.configure(text=f"{file_name}")

# -------------------------------------------------- INTERFACE
valeury = 5

# Fenetre principale
window = Tk()
window.configure(bg="grey85")
window.geometry('1000x600')
window.title("Détection de vocabulaire thésaurus")
window.columnconfigure(0, weight=1)
window.columnconfigure(1, weight=1)

# ----------------------------------------- Frame 1
frame1 = Frame(window)
frame1.grid(row=0, column=0, columnspan= 2, padx = 8, pady = (8, 0), ipadx= 200, ipady = 0, sticky=NSEW)
frame1.columnconfigure(0, weight=1)
frame1.columnconfigure(1, weight=1)

# Texte explicatif
doc_txt = """Ce programme permet de comparer les termes des colonnes/champs des exports Cindoc avec un thésaurus de référence contenu dans un répertoire.
Attention : la fenêtre pourrait ne pas répondre lors de la recherche. Il faut attendre, le programme n'a pas planté (barre de chargement à implémenter)."""
exp_txt = Label(frame1, text=doc_txt, anchor=W)
exp_txt.grid(row=0, column=0, sticky=E, pady=10, padx=30)

# ----------------------------------------- Frame 2
frame2 = Frame(window)
frame2.grid(row=1, column=0, padx = (8, 0), pady = (8, 0), ipadx= 380, ipady = 0, sticky=NSEW)
frame2.columnconfigure(0, weight=1)
frame2.columnconfigure(1, weight=1)
# Texte thésaurus
thes_text = Label(frame2, text="Sélection du répertoire contenant les fichiers XML du thésaurus.")
thes_text.grid(row=0, column=0, pady=valeury, columnspan=3)
# Bouton selection répertoire thésaurus
dir_button = Button(frame2, text="Sélection du répertoire", command=thesaurus_to_df)
dir_button.grid(row=1, column=0, pady=valeury, sticky=S)
# Bouton sauvegarde thésaurus
save_thes_button = Button(frame2, text="Sauvegarde/contrôle du thésaurus généré", command=save_df_thes)
save_thes_button.grid(row=1, column=1, pady=valeury, sticky=S)

# ----------------------------------------- Frame 3
frame3 = Frame(window)
frame3.grid(row=1, column=1, padx = 8, pady = (8, 0), ipadx= 380, ipady = 0, sticky=NSEW)
frame3.columnconfigure(0, weight=1)
frame3.columnconfigure(1, weight=1)
# Texte TSV
tsv_text = Label(frame3, text="Sélection du fichier d'export Cindoc au format TSV (*.txt).")
tsv_text.grid(row=0, column=0, pady=valeury, columnspan=3)
# Bouton selection répertoire thésaurus
file_button = Button(frame3, text="Sélection du fichier", command=openTSV)
file_button.grid(row=1, column=0, columnspan=3, pady=valeury, sticky=S)

# ----------------------------------------- Frame 4
frame4 = Frame(window)
frame4.grid(row=3, column=0, columnspan= 2, padx = 8, pady = 8, ipadx= 555, ipady = 100, sticky=NSEW)
frame4.columnconfigure(0, weight=1)
frame4.columnconfigure(1, weight=1)
frame4.rowconfigure(2, weight=1)
# Bouton de sauvegarde
search_button = Button(frame4, text="Rechercher", command=checker)
search_button.grid(row=0, column=0, pady=(10, 5), padx=valeury, sticky=E)

# Champ de résultat
txt_box = scrolledtext.ScrolledText(frame4)
txt_box.grid(row=2, column=0, columnspan=2, pady=valeury, sticky=NSEW)

window.mainloop()