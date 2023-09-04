# Deux classes, une pour chaque type de thésaurus
# Chaque clAsse contient sa table de correspondance
# les clés "sep" NE FINALEMENT SONT PAS UTILISEES NE PAS RECHERCHER LEUR UTILISATION AILLEURS DANS LE CODE

class ColCorrespMerimee:
    def __init__(self, col_cindoc, thesaurus):
        self.col_cindoc = col_cindoc
        self.thesaurus = thesaurus
        self.correspondances = {
            "COPY": {
                "col": ["DossierCopyright"],
                "sep": ";"
            },
            "DBOR": {
                "regex": "xxxxx"
            },
            "ETUD": {
                "col": ["CadreEtude", "CadreEtudeAffixe"],
                "sep": ";:(,"
            },
            "DOSS": {
                "col": ["TypeDossier", "Nature"],
                "sep": ";"
            },
            "DENO": {
                "col": ["DenominationArchi"],
                "sep": ";(,"
            },
            "GENR": {
                "col": ["Genre"],
                "sep": ";"
            },
            "VOCA": {
                "col": ["Vocable"],
                "sep": ";"
            },
            "APPL": {
                "col": ["AppellationArchi"],
                "sep": ";(,"
            },
            "PARN": {
                "col": ["DenominationArchi"],
                "sep": ";("
            },
            "AIRE": {
                "col": ["AireEtude"],
                "sep": ";"
            },
            "PLOC": {
                "col": ["PrecisionLocalisation"],
                "sep": ";:,"
            },
            "CANT": {
                "col": ["Canton"],
                "sep": ";,"
            },
            "LIEU": {
                "col": ["LieuDitNom", "LieuDitPrecision"],
                "sep": ";(,"
            },
            "ADRS": {
                "col": ["AdresseTypeVoie"],
                "sep": ";(,"
            },
            "IMPL": {
                "col": ["MilieuImplantation"],
                "sep": ";"
            },
            "HYDR": {
                "col": ["HydrographieNom", "HydrographiePrecision"],
                "sep": ";(,"
                },
            "SCLE": {
                "col": ["DatationPeriode"],
                "sep": ";(,"
            },
            "SCLD": {
                "col": ["DatationPeriode"],
                "sep": ";(,"
            },
            "JDAT": {
                "col": ["DatationJustification"],
                "sep": ";"
            },
            "AUTR": {
                "col": ["ProfessionAuteurOeuvreArchitecture"],
                "sep": ";:(,"
            },
            "JATT": {
                "col": ["JustificationProfession"],
                "sep": ";"
            },
            "PERS": {
                "col": ["ProfessionPersonnalite"],
                "sep": ";:(,"
            },
            "REMP": {
                "col": ["RemploiIntro"],
                "sep": ";:,"
            },
            "DEPL": {
                "col": ["DeplacementIntro"],
                "sep": ";:,"
            },
            "MURS": {
                "col": ["MurMateriauGrosOeuvre", "MurMiseEnOeuvre", "MurRevetement"],
                "sep": ";("
            },
            "TOIT": {
                "col": ["Toit"],
                "sep": ";("
            },
            "PLAN": {
                "col": ["Plan"],
                "sep": ";"
            },
            "ETAG": {
                "col": ["Etage"],
                "sep": ";"
            },
            "VOUT": {
                "col": ["CouvrementValeur", "CouvrementAffixe"],
                "sep": ";"
            },
            "ELEV": {
                "col": ["Elevation"],
                "sep": ";"
            },
            "COUV": {
                "col": ["CouvertureType", "CouvertureForme", "CouverturePartieToit"],
                "sep": ";"
            },
            "ESCA": {
                "col": ["EscalierEmplacement", "EscalierForme", "EscalierStructure", "AutreOrganeCirculation"],
                "sep": ";:,"
            },
            "ENER": {
                "col": ["EnergieNature", "EnergieOrigine", "EnergieMachine"],
                "sep": ";"
            },
            "VERT": {
                "col": ["Jardin"],
                "sep": ";"
            },
            "TECH": {
                "col": ["Technique"],
                "sep": ";("
            },
            "REPR": {
                "col": ["RepresentationValeur", "RepresentationSymbole"],
                "sep": ";:,"
            },
            "DIMS": {
                "col": ["TypeMesure", "DimensionUnite"],
                "sep": ";"
            },
            "TYPO": {
                "col": ["Typologie"],
                "sep": ";"
            },
            "ETAT": {
                "col": ["EtatConservationArchi"],
                "sep": ";"
            },
            "DPRO": {
                "col": ["ProtectionOeuvreArchi"],
                "sep": ";"
            },
            "SITE": {
                "col": ["SiteProtection"],
                "sep": ";"
            },
            "INTE": {
                "col": ["InteretOeuvre"],
                "sep": ";"
            },
            "STAT": {
                "col": ["Statut"],
                "sep": ";("
            }
        }
        if self.col_cindoc in self.correspondances and "col" in self.correspondances[self.col_cindoc]:
            self.cols = self.correspondances[self.col_cindoc]['col']
        if self.col_cindoc in self.correspondances and "sep" in self.correspondances[self.col_cindoc]:
            self.sep = self.correspondances[self.col_cindoc]['sep']
        else:
            self.sep = ""

    def make_voc_list(self):
        try:
            voc = []
            voc.append("?")
            for column in self.cols:
                    for word in self.thesaurus[column]:
                        if word == "(c) Région Grand-Est - Inventaire général":
                            word = "(c) Région Grand Est - Inventaire général"
                        voc.append(word)
            return voc
        except:
            #  return "Colonne Cindoc non concernée par la correspondance"
            pass

class ColCorrespPalissy:
    def __init__(self, col_cindoc, thesaurus):
        self.col_cindoc = col_cindoc
        self.thesaurus = thesaurus
        self.correspondances = {
            "COPY": {
                "col": ["DossierCopyright"],
                "sep": ";"
            },
            "ETUD": {
                "col": ["CadreEtude", "CadreEtudeAffixe"],
                "sep": ";:(,"
            },
            "DOSS": {
                "col": ["Nature"],
                "sep": ";"
            },
            "DENO": {
                "col": ["DenominationObjet"],
                "sep": ";(,"
            },
            "TITR": {
                "col": ["Titre"],
                "sep": ";(,"
            },
            "APPL": {
                "col": ["AppellationObjet"],
                "sep": ";(,"
            },
            "PART": {
                "col": ["DenominationObjet"],
                "sep": ";("
            },
            "PARN": {
                "col": ["DenominationObjet"],
                "sep": ";("
            },
            "INSEE": {
                "col": ["Commune"],
                "sep": ";"
            },
            "AIRE": {
                "col": ["AireEtude"],
                "sep": ";"
            },
            "PLOC": {
                "col": ["PrecisionLocalisation"],
                "sep": ";:,"
            },
            "CANT": {
                "col": ["Canton"],
                "sep": ";,"
            },
            "LIEU": {
                "col": ["LieuDitNom", "LieuDitPrecision"],
                "sep": ";(,"
            },
            "ADRS": {
                "col": ["AdresseTypeVoie"],
                "sep": ";(,"
            },
            "IMPL": {
                "col": ["MilieuImplantation"],
                "sep": ";"
            },
            "DEPL": {
                "col": ["DeplacementOeuvreObjet"],
                "sep": ";:(,"
            },
            "CATE": {
                "col": ["Categorie"],
                "sep": ";"
                },
            "STRU": {
                "col": ["StructureLibelle", "StructureAffixe"],
                "sep": ";(,"
            },
            "MATR": {
                "col": ["MateriauPrincipal", "MateriauTechnique", "MateriauAffixe"],
                "sep": ";(,"
            },
            "REPR": {
                "col": ["RepresentationCaractereGeneral", "RepresentationTheme", "RepresentationSujet"],
                "sep": ";:(,"
            },
            "DIMS": {
                "col": ["TypeMesure", "DimensionUnite"],
                "sep": ";"
            },
            "ETAT": {
                "col": ["EtatConservationObjet"],
                "sep": ";"
            },
            "INSC": {
                "col": ["InscriptionValeur", "InscriptionAffixe"],
                "sep": ";,("
            },
            "AUTR": {
                "col": ["ProfessionAuteurOeuvreObjet"],
                "sep": ";:(,"
            },
            "AFIG": {
                "col": ["ProfessionAfig"],
                "sep": ";:,"
            },
            "ATEL": {
                "col": ["ProfessionAtelierEcole"],
                "sep": ";("
            },
            "PERS": {
                "col": ["ProfessionPersonnalite"],
                "sep": ";:(,"
            },
            "EXEC": {
                "col": ["LieuExecution"],
                "sep": ";:(,"
            },
            "ORIG": {
                "col": ["Origine"],
                "sep": ";:(,"
            },
            "STAD": {
                "col": ["StadeCreation", "StadeCreationAffixe"],
                "sep": ";(,"
            },
            "SCLE": {
                "col": ["DatationPeriode"],
                "sep": ";"
            },
            "SCLD": {
                "col": ["DatationPeriode"],
                "sep": ";"
            },
            "STAT": {
                "col": ["Statut"],
                "sep": ";("
            },
            "PROT": {
                "col": ["ProtectionOeuvreObjet"],
                "sep": ";"
            },
            "INTE": {
                "col": ["InteretOeuvre"],
                "sep": ";"
            }
        }
        if self.col_cindoc in self.correspondances and "col" in self.correspondances[self.col_cindoc]:
            self.cols = self.correspondances[self.col_cindoc]['col']
        if self.col_cindoc in self.correspondances and "sep" in self.correspondances[self.col_cindoc]:
            self.sep = self.correspondances[self.col_cindoc]['sep']
        else:
            self.sep = ""
        # if "sep" in self.correspondances[self.col]:
        #     self.sep = self.correspondances[self.col]['sep']
        # if "regex" in self.correspondances[self.col]:
        #     self.regex = self.correspondances[self.col]['regex']

    def make_voc_list(self):
        try:
            voc = []
            voc.append("?")
            for column in self.cols:
                    for word in self.thesaurus[column]:
                        voc.append(word)
            return voc
        except:
            #  return "Colonne Cindoc non concernée par la correspondance"
            pass