# Recherche
1. Ouvrir le fichier du thésaurus (souvent ThesaurusEntry.csv).
2. Sélectionner le type entre Mérimée ou Palissy.
3. Ouvrir l'export à vérifier (format Excel).
4. Optionnel : ajouter des colonnes à ignorer (par exemple si trop de faux positifs ou pour travailler en plusieurs fois). Si plusieurs colonnes, simplement les séparer par un espace.
5. Lancer la recherche.

# Correction
[Voir provenance] donne la provenance du terme sur l'ensemble de la colonne/champ (en gras). Autrement dit, si un terme est sélectionné, il remplacera la partie en gras du/des cellules concernée.
[Remplacer par :] permet de sélectionner le terme candidat au remplacement (le bouton devient vert).
⚠️ Si erreur : désélectionner l'ensemble des éventuels boutons pressés avant de sélectionner le nouveau.
⚠️ Pour le dernier cas [Remplacer par :] [*champ libre*] : le bouton est à presser uniquement une fois que le champ libre est bien complet. Si modification, désélectionner le bouton, puis le resélectionner.

# Sauvegarde
Colonnes à compiler dans GESTION : pré-rempli mais possibilité de modifier. Si GESTION était déjà présent dans le fichier d'origine, il sera compilé dans la nouvelle colonne GESTION sous le nom de "GESTION_old". exemple : {GESTION_old=Ipsum lorem}.

Le nom du fichier proposé à la sauvegarde contient toujours la date et l'heure du traitement : permet de travailler sur de multiples versions facilement. Si l'excel en entrée avait déjà une date, elle est mise à jour.