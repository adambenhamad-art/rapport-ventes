"""Crée ventes.xlsx : une année de ventes réalistes d'une épicerie fine, pour tester le rapport.

Les ventes suivent une saisonnalité (pic des dattes pendant le Ramadan, hausse en fin
d'année…) et le fichier contient volontairement quelques lignes erronées pour montrer
que le rapport les détecte.
"""

import random
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# (produit, catégorie, prix unitaire, popularité)
PRODUITS = [
    ("Café en grains 1kg", "Boissons", 18.50, 5),
    ("Thé vert à la menthe", "Boissons", 6.90, 8),
    ("Dattes Deglet Nour 500g", "Épicerie", 7.80, 6),
    ("Huile d'olive vierge 1L", "Épicerie", 12.40, 5),
    ("Miel de fleurs 500g", "Épicerie", 11.00, 4),
    ("Amandes grillées 250g", "Épicerie", 5.60, 5),
    ("Savon d'Alep", "Hygiène", 4.50, 4),
    ("Coffret cadeau", "Cadeaux", 34.90, 1),
]

# Activité relative de chaque mois (1 = normal)
SAISON = {1: 0.8, 2: 0.85, 3: 1.35, 4: 1.0, 5: 0.95, 6: 0.9,
          7: 0.75, 8: 0.7, 9: 1.0, 10: 1.05, 11: 1.2, 12: 1.6}

RAMADAN_2025 = (date(2025, 3, 1), date(2025, 3, 30))

LIGNES_ERRONEES = [
    [None, "Thé vert à la menthe", 2, 6.90, "Boissons"],           # date manquante
    [date(2025, 6, 14), "Miel de fleurs 500g", "deux", 11.00, "Épicerie"],  # quantité en lettres
    [date(2025, 9, 3), "Savon d'Alep", -3, 4.50, "Hygiène"],        # quantité négative
]


def generer(chemin="ventes.xlsx", annee=2025, ventes_par_jour=4):
    random.seed(2025)  # mêmes données à chaque lancement
    lignes = []

    jour = date(annee, 1, 1)
    while jour.year == annee:
        nombre = round(random.gauss(ventes_par_jour * SAISON[jour.month], 1.2))
        for _ in range(max(nombre, 0)):
            poids = []
            for nom, _categorie, _prix, popularite in PRODUITS:
                bonus = 4 if nom.startswith("Dattes") and RAMADAN_2025[0] <= jour <= RAMADAN_2025[1] else 1
                bonus *= 4 if nom == "Coffret cadeau" and jour.month == 12 else 1
                poids.append(popularite * bonus)
            nom, categorie, prix, _ = random.choices(PRODUITS, weights=poids)[0]
            quantite = random.choices([1, 2, 3, 4, 5, 6, 10], weights=[30, 25, 15, 10, 8, 7, 5])[0]
            lignes.append([jour, nom, quantite, prix, categorie])
        jour += timedelta(days=1)

    for ligne in LIGNES_ERRONEES:
        lignes.insert(random.randint(0, len(lignes)), ligne)

    classeur = Workbook()
    feuille = classeur.active
    feuille.title = "Ventes"
    feuille.append(["Date", "Produit", "Quantité", "Prix unitaire", "Catégorie"])
    for cellule in feuille[1]:
        cellule.font = Font(bold=True)
        cellule.fill = PatternFill("solid", fgColor="D9E1F2")
    for ligne in lignes:
        feuille.append(ligne)

    for cellule in feuille["A"][1:]:
        cellule.number_format = "DD/MM/YYYY"
    for cellule in feuille["D"][1:]:
        cellule.number_format = '#,##0.00 "€"'
    for colonne, largeur in {"A": 12, "B": 28, "C": 10, "D": 14, "E": 14}.items():
        feuille.column_dimensions[colonne].width = largeur
    feuille.freeze_panes = "A2"

    classeur.save(chemin)
    print(f"Fichier d'exemple créé : {chemin} ({len(lignes)} lignes, dont "
          f"{len(LIGNES_ERRONEES)} volontairement erronées)")


if __name__ == "__main__":
    generer()
