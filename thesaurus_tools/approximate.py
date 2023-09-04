from rapidfuzz import fuzz
import operator

def make_calculation(word1, word2, alg):
    match alg:
        case "Simple Ratio":
            # Calcul du ratio entre les deux termes
            ratio = fuzz.ratio(word1, word2)
            return(ratio)
        case "Partial Ratio":
            # Calcul du ratio entre les deux termes
            ratio = fuzz.partial_ratio(word1, word2)
            return(ratio)
        case "Token Sort Ratio":
            # Calcul du ratio entre les deux termes
            ratio = fuzz.token_sort_ratio(word1, word2)
            return(ratio)
        case _:
            pass


def calculate_prox(word1, word2, alg):
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
    
    ratio = make_calculation(word1, word2, alg)

    # Boucle qui remplace les éventuels chiffres et les réinscrit en toute lettres
    for digit in digits:
        if digit[0] in word1:
            word1 = word1.replace(digit[0], digit[1])
            replaced = True

    # S'il y a effectivement eu un remplacement
    if replaced:
        ratio_repl = make_calculation(word1, word2, alg)
        if ratio_repl > ratio: # on ne garde que le cas ayant fait le meilleur score
            return ratio_repl
        else:
            return ratio
    else:
        return ratio


def calculate_proximity(term, refs, alg, num):
    scores = {}

    for ref in refs:
        score = calculate_prox(term, ref, alg)
        scores[ref] = score

    # num plus proche termes (num est passé en paramêtre : permet de faire choisir le nombre de propositions)
    scores = dict(sorted(scores.items(), key=operator.itemgetter(1), reverse=True)[:num])
    scores["$"] = 0 # cette clé du dictionnaire sera utilisée pour afficher l'input si l'utilisateur souhaite rentrer sa propre valeur de remplacement ("$" est utilisé arbitrairement pour reconnaitre ce champ ailleurs dans le code)
    return(scores)


