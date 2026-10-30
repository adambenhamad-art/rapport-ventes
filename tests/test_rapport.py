from datetime import date, datetime

import pytest
from openpyxl import Workbook, load_workbook

from rapport import (Vente, analyser, creer_rapport, lire_date, lire_nombre, lire_ventes,
                     main, trouver_colonnes)


# --- Lecture des valeurs ---

@pytest.mark.parametrize("valeur, attendu", [
    (18.5, 18.5),
    (3, 3.0),
    ("18,50", 18.5),
    ("18,50 €", 18.5),
    ("1 234,50", 1234.5),
    ("1.234,50", 1234.5),
    ("12.40", 12.4),
])
def test_lire_nombre(valeur, attendu):
    assert lire_nombre(valeur) == pytest.approx(attendu)


@pytest.mark.parametrize("valeur", ["deux", "", None, True])
def test_lire_nombre_invalide(valeur):
    with pytest.raises((ValueError, TypeError)):
        lire_nombre(valeur)


@pytest.mark.parametrize("valeur", [
    date(2025, 3, 14), datetime(2025, 3, 14, 10, 30), "14/03/2025", "2025-03-14", "14-03-2025",
])
def test_lire_date(valeur):
    assert lire_date(valeur) == date(2025, 3, 14)


def test_trouver_colonnes_ignore_accents_majuscules_et_ordre():
    colonnes = trouver_colonnes(["PRIX UNITAIRE", " Désignation ", "Qté", "Date", "Famille"])
    assert colonnes == {"prix": 0, "produit": 1, "quantite": 2, "date": 3, "categorie": 4}


# --- Lecture de fichiers ---

def fichier_excel(chemin, lignes):
    classeur = Workbook()
    for ligne in lignes:
        classeur.active.append(ligne)
    classeur.save(chemin)
    return chemin


def test_lire_excel_avec_titre_au_dessus_et_anomalies(tmp_path):
    chemin = fichier_excel(tmp_path / "ventes.xlsx", [
        ["Ventes de la boutique"],
        [],
        ["Date", "Produit", "Quantité", "Prix unitaire"],
        [date(2025, 1, 5), "Thé", 2, 6.9],
        [None, "Thé", 2, 6.9],
        [date(2025, 1, 6), "Miel", "deux", 11],
        [date(2025, 1, 7), "Savon", -1, 4.5],
        [date(2025, 1, 8), "", 1, 4.5],
        [],
        [date(2025, 2, 1), "Miel", 1, 11],
    ])
    ventes, anomalies = lire_ventes(chemin)

    assert [v.produit for v in ventes] == ["Thé", "Miel"]
    assert [a.ligne for a in anomalies] == [5, 6, 7, 8]


def test_lire_csv_format_francais(tmp_path):
    chemin = tmp_path / "ventes.csv"
    chemin.write_text("Date;Article;Qté;Prix\n05/01/2025;Thé;2;6,90\n06/01/2025;Miel;1;11,00\n",
                      encoding="utf-8")
    ventes, anomalies = lire_ventes(chemin)

    assert anomalies == []
    assert ventes[0] == Vente(date(2025, 1, 5), "Thé", 2, 6.9)
    assert sum(v.total for v in ventes) == pytest.approx(24.8)


def test_colonnes_manquantes(tmp_path):
    chemin = fichier_excel(tmp_path / "mauvais.xlsx", [["Nom", "Adresse"], ["Ali", "Paris"]])
    with pytest.raises(ValueError, match="date"):
        lire_ventes(chemin)


# --- Analyse ---

VENTES = [
    Vente(date(2025, 1, 10), "Thé", 2, 5.0, "Boissons"),     # 10 €
    Vente(date(2025, 1, 20), "Miel", 1, 30.0, "Épicerie"),   # 30 €
    Vente(date(2025, 3, 5), "Thé", 4, 5.0, "Boissons"),      # 20 €
]


def test_analyser_totaux():
    analyse = analyser(VENTES)
    assert analyse.chiffre_affaires == 60
    assert analyse.nb_ventes == 3
    assert analyse.panier_moyen == 20
    assert analyse.quantite_totale == 7


def test_analyser_classement_produits():
    produits = analyser(VENTES).produits
    # À égalité de chiffre d'affaires, ordre alphabétique
    assert [(p.nom, p.total, p.quantite) for p in produits] == [("Miel", 30, 1), ("Thé", 30, 6)]
    assert sum(p.part for p in produits) == pytest.approx(1)


def test_analyser_inclut_les_mois_sans_vente():
    mois = analyser(VENTES).mois
    assert [(m.libelle, m.total) for m in mois] == [("Janv. 2025", 40), ("Févr. 2025", 0),
                                                     ("Mars 2025", 20)]
    assert mois[0].variation is None
    assert mois[1].variation == -1
    assert mois[2].variation is None  # pas de calcul possible après un mois à 0


def test_analyser_sur_deux_annees():
    analyse = analyser([Vente(date(2024, 12, 1), "Thé", 1, 10.0),
                        Vente(date(2025, 1, 1), "Thé", 1, 15.0)])
    assert [m.libelle for m in analyse.mois] == ["Déc. 2024", "Janv. 2025"]
    assert analyse.mois[1].variation == pytest.approx(0.5)


# --- Rapport complet ---

def test_creer_rapport(tmp_path):
    chemin = tmp_path / "rapport.xlsx"
    creer_rapport(analyser(VENTES), [], chemin, titre="Ma boutique")

    classeur = load_workbook(chemin)
    assert classeur.sheetnames == ["Tableau de bord", "Par produit", "Par mois", "Données"]
    tableau_de_bord = classeur["Tableau de bord"]
    assert "Ma boutique" in tableau_de_bord["B2"].value
    assert tableau_de_bord["B6"].value == 60
    assert len(tableau_de_bord._charts) == 3


def test_main_de_bout_en_bout(tmp_path, capsys):
    source = fichier_excel(tmp_path / "ventes.xlsx", [
        ["Date", "Produit", "Quantité", "Prix unitaire"],
        [date(2025, 1, 5), "Thé", 2, 6.9],
        ["pas une date", "Thé", 1, 6.9],
    ])
    sortie = tmp_path / "rapport.xlsx"
    main([str(source), "-o", str(sortie)])

    assert "Anomalies" in load_workbook(sortie).sheetnames
    assert "1 ligne(s) en erreur" in capsys.readouterr().out


def test_main_fichier_introuvable(tmp_path):
    with pytest.raises(SystemExit, match="introuvable"):
        main([str(tmp_path / "absent.xlsx")])
