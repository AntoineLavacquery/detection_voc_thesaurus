import re

def in_parenthesis(string):
    try:
        part = re.search(r"\((.*?)\)", string)[1].strip()
        return(part)
    except:
        return(string)

def before_comma(string):
    try:
        part = re.search(r"^(.*?):", string)[1].strip()
        return(part)
    except:
        return(string)

def after_comma(string):
    try:
        part = re.search(r":\s?(.*)$", string)[1].strip()
        return(part)
    except:
        return(string)
    
def before_parenthesis(string):
    try:
        part = re.search(r"^(.*?)\(", string)[1].strip()
        return(part)
    except:
        return(string)

def make_list_by_ponctuation(term, ponctuation):
    for ponct in ponctuation:
        term = term.replace(ponct, "$")
    terms = term.split("$")
    terms = [term.strip() for term in terms if len(term) != 0]
    return(terms)

def split_all(term):
    term = term.replace(")", "")
    terms = re.split(r":|\(|,", term)
    terms = [term.strip() for term in terms]
    return(terms)

# Fonction qui se sert des fonctions déclarées plus haut : permet de ne considérer qu'une partie des termes à vérifier en fonction de la morphologie des champs
def adapt_to_col(term, col_name):
    match col_name:
        case "ADRS":
            return(in_parenthesis(term))
        case "PERS":
            return(in_parenthesis(term))
        case "AUTR":
            term = in_parenthesis(term)
            return(make_list_by_ponctuation(term, ","))
        case "DPRO":
            return(after_comma(term))
        case "REMP":
            return(before_comma(term))
        case "DEPL":
            return(before_comma(term))
        case "TECH":
            return(before_parenthesis(term))
        case "SCLE":
            return(before_parenthesis(term))
        case "COPY":
            return(term)
        case _:
            return(split_all(term)) 

