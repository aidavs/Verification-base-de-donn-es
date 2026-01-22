import pandas as pd  # permet de lire et traiter les fichiers exel
from dataclasses import dataclass  # pour créer des structures de données simples sans écrire beaucoup de code come ___ini____
from typing import Any  # pour pouvoir utiliser "any" comme type


# Chargement du fichier
chemin_fichier = input(
    "Rentrer le chemin du fichier absolu (ex: C:/chemin/fichier.xlsx) : "
)


df = None  # on met none par défaut pour éviter les erreurs si le fichier n'est pas trouvé

# vérifie si le fichier excel peut être chargé
try:
    if chemin_fichier.endswith((".xls", ".xlsx")):
        df = pd.read_excel(chemin_fichier)
    else:
        raise ValueError("Format de fichier invalide (xls ou xlsx attendu)")
except FileNotFoundError:
    print("Fichier introuvable, le chemin est incorrect")
except ValueError as e:
    print(e)


@dataclass
# préciser  le type de données des warnings
class WarningDonnees:
    ligne: int
    colonne: str
    valeur: Any
    regle: str


def verifier_regles(ligne: dict, num_ligne: int):
    warnings = []

    # 1️ Colonnes obligatoires
    for col in ['A', 'B', 'D', 'F', 'G', 'I']:
        if col not in ligne or pd.isna(ligne[col]):
            warnings.append(
                WarningDonnees(
                    num_ligne,
                    col,
                    None,
                    "Il manque des données dans ces colonnes obligatoires vides"
                )
            )

    # 2️ Colonne A : départements français ou 999 pour l'étranger
    valeurs_A = (
        list(range(1, 96)) +
        list(range(971, 979)) +
        [984] +
        list(range(986, 990)) +
        [999]
    )

    # vérifie si le département est valide
    if not pd.isna(ligne.get('A')) and ligne.get('A') not in valeurs_A:
        warnings.append(
            WarningDonnees(
                num_ligne,
                "A",
                ligne.get('A'),
                "Numéro de département non valide "
            )
        )

    # Précisez quand le test a été réalisé pour les marqueurs sériques maternels
    if ligne.get('B') in [3, 6, 8]:
        if ligne.get('D') in [0, 99]:
            warnings.append(
                WarningDonnees(
                    num_ligne,
                    "D",
                    ligne.get('D'),
                    "Précisez quand le test a été réalisé pour les marqueurs sériques maternels"
                )
            )
        
        # si il y a rien
        if (isinstance(ligne.get('E'), (int))):
             warnings.append(
                WarningDonnees(
                    num_ligne,
                    "E",
                    ligne.get('E'),
                    "Il manque une donné"
                )
            )
        # vérifie si'il y a un nombre
        if not (pd.isna(ligne.get('E'))):
            warnings.append(
                WarningDonnees(
                    num_ligne,
                    "E",
                    ligne.get('E'),
                    "Il faut entrer un nombre"
                )
            )

    # 4️ B = 99
    if ligne.get('B') == 99 and not isinstance(ligne.get('C'), str):
        warnings.append(
            WarningDonnees(
                num_ligne,
                "C",
                ligne.get('C'),
                "faites une déscription du marqueur atypique"
            )
        )

    # Cette colonne doit rester vide si aucun test n'a été fait ou si la date du test est inconnu
    if ligne.get('D') in [0, 99] and not pd.isna(ligne.get('E')):
        warnings.append(
            WarningDonnees(
                num_ligne,
                "E",
                ligne.get('E'),
                "Cette colonne doit rester vide si aucun test n'a été fait ou si la date du test est inconnu"
            )
        )

    # Précisez la valeur du risque connu découvert grâce aux tests faits durant la grossesse
    if ligne.get('D') in [1, 2, 3] and pd.isna(ligne.get('E')):
        warnings.append(
            WarningDonnees(
                num_ligne,
                "E",
                None,
                "Précisez la valeur du risque connu découvert grâce aux tests faits durant la grossesse "
            )
        )

    # Vous donnez le résultats pour un test autre que celui de la trisomie 21
    if ligne.get('F') == 1 and ligne.get('G') not in [0, 1, 99]:
        warnings.append(
            WarningDonnees(
                num_ligne,
                "G",
                ligne.get('G'),
                "Vous donnez le résultats pour un test autre que celui de la trisomie 21, pourtant vous avez fait un test pour la trisomie 21 seule"
            )
        )

    # Quand il faut précisé le type d'anomalie génétique
    if ligne.get('G') in [4, 5] and not isinstance(ligne.get('H'), str):
        warnings.append(
            WarningDonnees(
                num_ligne,
                "H",
                ligne.get('H'),
                "Veuillez préciser le type de maladie"
            )
        )

    if ligne.get('F') == 3 and ligne.get('G') not in [0, 1, 2, 3, 4, 5, 99]:
        warnings.append(
            WarningDonnees(
                num_ligne,
                "G",
                ligne.get('G'),
                "Veuillez préciser le type de test faits ou si aucun test n'a été fait"
            )
        )
    
    if ligne.get('F') == 2 and ligne.get('G') not in [0, 1, 2, 3, 99]:
        warnings.append(
            WarningDonnees(
                num_ligne,
                "G",
                ligne.get('G'),
                "Veuillez préciser le type de test faits ou si aucun test n'a été fait"
            )
        )

    # Si le résultat n'est pas négatif et est exploitable il faut le type de test
    if ligne.get('G') in [1, 2, 3, 4, 5] and ligne.get('I') == 0:
        warnings.append(
            WarningDonnees(
                num_ligne,
                "I",
                0,
                "Si le résultat n'est pas négatif et est exploitable il faut le type de test. Veuillez entrer le test utilisé"
            )
        )

    #  Règles de progression (warnings informatifs)
    progression = {1: 2, 2: 3, 3: 4, 4: 5, 5: 6}

    if ligne.get('G') in progression:
        attendu = progression[ligne.get('G')]
        if ligne.get('I') != attendu:
            warnings.append(
                WarningDonnees(
                    num_ligne,
                    "I",
                    ligne.get('I'),
                    f"Progression attendue : G={ligne.get('G')} → I={attendu}"
                )
            )

    return warnings


# Vérification
toutes_les_alertes = []

if df is not None:
    for index, row in df.iterrows():
        alertes = verifier_regles(row.to_dict(), index + 2)
        toutes_les_alertes.extend(alertes)
else:
    print("La vérification des données n'a pas pu être effectuée en raison d'une erreur lors du chargement du fichier.")
