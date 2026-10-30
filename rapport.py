"""Rapport de ventes automatique.

Transforme un fichier de ventes (Excel ou CSV) en rapport Excel professionnel :
tableau de bord, analyse par produit et par mois, graphiques, et liste des
lignes en erreur.

Utilisation :
    python rapport.py ventes.xlsx
    python rapport.py ventes.csv -o rapport_2025.xlsx --titre "Épicerie Al Baraka"

Colonnes reconnues automatiquement (majuscules et accents ignorés) :
    Date        : date, jour
    Produit     : produit, article, désignation, product
    Quantité    : quantité, qté, qty, quantity
    Prix        : prix unitaire, prix, pu, price, unit price
    Catégorie   : catégorie, famille, category (facultative)
"""

import argparse
import csv
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------------------
# Réglages
# ---------------------------------------------------------------------------

SYNONYMES = {
    "date": ["date", "jour", "date de vente"],
    "produit": ["produit", "article", "designation", "product", "libelle"],
    "quantite": ["quantite", "qte", "qty", "quantity", "nombre"],
    "prix": ["prix unitaire", "prix", "pu", "price", "unit price", "prix unit"],
    "categorie": ["categorie", "famille", "category", "rayon"],
}
OBLIGATOIRES = ("date", "produit", "quantite", "prix")

MOIS_COURTS = ["Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin",
               "Juil.", "Août", "Sept.", "Oct.", "Nov.", "Déc."]

BLEU = "1F3864"
BLEU_CLAIR = "D9E1F2"
GRIS = "595959"
FOND_CARTE = "F2F5FB"
FORMAT_EURO = '#,##0.00 "€"'
FORMAT_POURCENT = "0.0%"


# ---------------------------------------------------------------------------
# 1. Modèle de données
# ---------------------------------------------------------------------------

@dataclass
class Vente:
    date: date
    produit: str
    quantite: int
    prix: float
    categorie: str = ""

    @property
    def total(self):
        return self.quantite * self.prix


@dataclass
class Anomalie:
    ligne: int
    contenu: str
    raison: str


@dataclass
class LigneProduit:
    nom: str
    categorie: str
    quantite: int
    total: float
    part: float


@dataclass
class LigneMois:
    annee: int
    mois: int
    total: float
    nb_ventes: int
    variation: float | None  # évolution par rapport au mois précédent

    @property
    def libelle(self):
        return f"{MOIS_COURTS[self.mois - 1]} {self.annee}"


@dataclass
class Analyse:
    ventes: list
    chiffre_affaires: float
    quantite_totale: int
    debut: date
    fin: date
    produits: list = field(default_factory=list)
    mois: list = field(default_factory=list)

    @property
    def nb_ventes(self):
        return len(self.ventes)

    @property
    def panier_moyen(self):
        return self.chiffre_affaires / self.nb_ventes

    @property
    def meilleur_mois(self):
        return max(self.mois, key=lambda m: m.total)

    @property
    def mois_le_plus_faible(self):
        return min(self.mois, key=lambda m: m.total)

    @property
    def a_des_categories(self):
        return any(v.categorie for v in self.ventes)


# ---------------------------------------------------------------------------
# 2. Lecture du fichier (Excel ou CSV)
# ---------------------------------------------------------------------------

def normaliser(texte):
    """'Quantité ' -> 'quantite' : minuscules, sans accents ni espaces superflus."""
    texte = unicodedata.normalize("NFKD", str(texte)).encode("ascii", "ignore").decode()
    return " ".join(texte.lower().replace("_", " ").replace(".", " ").split())


def trouver_colonnes(entete):
    """Associe chaque champ (date, produit…) au numéro de sa colonne."""
    noms = [normaliser(nom) if nom is not None else "" for nom in entete]
    colonnes = {}
    for champ, synonymes in SYNONYMES.items():
        for index, nom in enumerate(noms):
            if nom in synonymes and index not in colonnes.values():
                colonnes[champ] = index
                break
    return colonnes


def lire_nombre(valeur):
    """Accepte 18.5, '18,50', '18,50 €', '1 234,50'…"""
    if isinstance(valeur, bool):
        raise ValueError
    if isinstance(valeur, (int, float)):
        return float(valeur)
    texte = str(valeur).replace("€", "").replace(" ", "").replace(" ", "").replace(" ", "")
    if "," in texte:
        texte = texte.replace(".", "").replace(",", ".")
    return float(texte)


def lire_date(valeur):
    if isinstance(valeur, datetime):
        return valeur.date()
    if isinstance(valeur, date):
        return valeur
    texte = str(valeur).strip()
    for format_date in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(texte, format_date).date()
        except ValueError:
            pass
    raise ValueError


def afficher(valeur):
    """Valeur brute -> texte lisible pour l'onglet Anomalies."""
    if valeur is None:
        return "(vide)"
    if isinstance(valeur, (datetime, date)):
        return f"{valeur:%d/%m/%Y}"
    if isinstance(valeur, float):
        return f"{valeur:g}".replace(".", ",")
    return str(valeur)


def lignes_excel(chemin):
    feuille = load_workbook(chemin, data_only=True, read_only=True).active
    yield from feuille.iter_rows(values_only=True)


def lignes_csv(chemin):
    for encodage in ("utf-8-sig", "cp1252"):
        try:
            texte = Path(chemin).read_text(encoding=encodage)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialecte = csv.Sniffer().sniff(texte[:4096], delimiters=";,\t")
    except csv.Error:
        dialecte = csv.excel
    yield from csv.reader(texte.splitlines(), dialecte)


def lire_ventes(chemin):
    """Renvoie (ventes valides, anomalies). Lève ValueError si le fichier est inexploitable."""
    lignes = lignes_csv(chemin) if str(chemin).lower().endswith(".csv") else lignes_excel(chemin)

    # L'en-tête peut ne pas être sur la première ligne (titre, logo…) : on le cherche.
    colonnes = None
    for numero, ligne in enumerate(lignes, start=1):
        colonnes = trouver_colonnes(ligne)
        if all(champ in colonnes for champ in OBLIGATOIRES):
            break
        if numero >= 15:
            colonnes = None
            break
    if not colonnes or not all(champ in colonnes for champ in OBLIGATOIRES):
        manquantes = ", ".join(c for c in OBLIGATOIRES if c not in (colonnes or {}))
        raise ValueError(f"colonnes introuvables : {manquantes}")

    ventes, anomalies = [], []
    for numero, ligne in enumerate(lignes, start=numero + 1):
        if not ligne or all(v is None or str(v).strip() == "" for v in ligne):
            continue
        valeur = lambda champ: ligne[colonnes[champ]] if colonnes[champ] < len(ligne) else None
        contenu = " | ".join(afficher(v) for v in ligne)

        try:
            jour = lire_date(valeur("date"))
        except (TypeError, ValueError):
            anomalies.append(Anomalie(numero, contenu, "Date manquante ou invalide"))
            continue
        produit = str(valeur("produit") or "").strip()
        if not produit:
            anomalies.append(Anomalie(numero, contenu, "Produit manquant"))
            continue
        try:
            quantite = lire_nombre(valeur("quantite"))
            prix = lire_nombre(valeur("prix"))
        except (TypeError, ValueError):
            anomalies.append(Anomalie(numero, contenu, "Quantité ou prix non numérique"))
            continue
        if quantite <= 0 or quantite != int(quantite):
            anomalies.append(Anomalie(numero, contenu, "Quantité négative, nulle ou non entière"))
            continue
        if prix < 0:
            anomalies.append(Anomalie(numero, contenu, "Prix négatif"))
            continue

        categorie = str(valeur("categorie") or "").strip() if "categorie" in colonnes else ""
        ventes.append(Vente(jour, produit, int(quantite), prix, categorie))

    return ventes, anomalies


# ---------------------------------------------------------------------------
# 3. Analyse
# ---------------------------------------------------------------------------

def analyser(ventes):
    ventes = sorted(ventes, key=lambda v: v.date)
    chiffre_affaires = sum(v.total for v in ventes)

    par_produit = defaultdict(lambda: {"quantite": 0, "total": 0.0, "categorie": ""})
    par_mois = defaultdict(lambda: {"total": 0.0, "nb": 0})
    for vente in ventes:
        produit = par_produit[vente.produit]
        produit["quantite"] += vente.quantite
        produit["total"] += vente.total
        produit["categorie"] = produit["categorie"] or vente.categorie
        mois = par_mois[(vente.date.year, vente.date.month)]
        mois["total"] += vente.total
        mois["nb"] += 1

    produits = [
        LigneProduit(nom, p["categorie"], p["quantite"], p["total"],
                     p["total"] / chiffre_affaires if chiffre_affaires else 0)
        for nom, p in par_produit.items()
    ]
    produits.sort(key=lambda p: (-p.total, p.nom))  # à égalité : ordre alphabétique

    # Tous les mois de la période, même ceux sans vente
    debut, fin = ventes[0].date, ventes[-1].date
    lignes_mois, precedent = [], None
    annee, mois = debut.year, debut.month
    while (annee, mois) <= (fin.year, fin.month):
        donnees = par_mois.get((annee, mois), {"total": 0.0, "nb": 0})
        variation = (donnees["total"] - precedent) / precedent if precedent else None
        lignes_mois.append(LigneMois(annee, mois, donnees["total"], donnees["nb"], variation))
        precedent = donnees["total"]
        annee, mois = (annee + 1, 1) if mois == 12 else (annee, mois + 1)

    return Analyse(ventes, chiffre_affaires, sum(v.quantite for v in ventes),
                   debut, fin, produits, lignes_mois)


# ---------------------------------------------------------------------------
# 4. Création du rapport Excel
# ---------------------------------------------------------------------------

def euros(montant):
    return f"{montant:,.2f} €".replace(",", " ").replace(".", ",")


def bandeau(feuille, cellules, texte, taille=16):
    feuille.merge_cells(cellules)
    cellule = feuille[cellules.split(":")[0]]
    cellule.value = texte
    cellule.font = Font(bold=True, size=taille, color="FFFFFF")
    cellule.fill = PatternFill("solid", fgColor=BLEU)
    cellule.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def tableau(feuille, nom, entetes, lignes, ligne_debut, formats):
    """Écrit un tableau Excel (filtres, lignes alternées) et renvoie sa dernière ligne."""
    for colonne, entete in enumerate(entetes, start=1):
        feuille.cell(row=ligne_debut, column=colonne, value=entete)
    for i, ligne in enumerate(lignes, start=ligne_debut + 1):
        for colonne, valeur in enumerate(ligne, start=1):
            cellule = feuille.cell(row=i, column=colonne, value=valeur)
            if formats.get(colonne):
                cellule.number_format = formats[colonne]
    fin = ligne_debut + max(len(lignes), 1)
    zone = f"A{ligne_debut}:{get_column_letter(len(entetes))}{fin}"
    table = Table(displayName=nom, ref=zone)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    feuille.add_table(table)
    feuille.freeze_panes = f"A{ligne_debut + 1}"
    return fin


def largeurs(feuille, valeurs):
    for colonne, largeur in valeurs.items():
        feuille.column_dimensions[colonne].width = largeur


def feuille_produits(classeur, analyse):
    feuille = classeur.create_sheet("Par produit")
    avec_categorie = analyse.a_des_categories
    entetes = ["Rang", "Produit"] + (["Catégorie"] if avec_categorie else []) + \
              ["Quantité vendue", "Chiffre d'affaires", "Part du CA"]
    lignes = [
        [rang, p.nom] + ([p.categorie] if avec_categorie else []) + [p.quantite, p.total, p.part]
        for rang, p in enumerate(analyse.produits, start=1)
    ]
    col_ca, col_part = len(entetes) - 1, len(entetes)
    fin = tableau(feuille, "Produits", entetes, lignes, 1,
                  {col_ca: FORMAT_EURO, col_part: FORMAT_POURCENT})

    lettre_ca = get_column_letter(col_ca)
    feuille.conditional_formatting.add(
        f"{lettre_ca}2:{lettre_ca}{fin}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color="5B9BD5"))
    largeurs(feuille, {"A": 7, "B": 28, "C": 16, "D": 18, "E": 18, "F": 12})
    return feuille, col_ca, fin


def feuille_mois(classeur, analyse):
    feuille = classeur.create_sheet("Par mois")
    lignes = [[m.libelle, m.nb_ventes, m.total, m.variation] for m in analyse.mois]
    fin = tableau(feuille, "Mois", ["Mois", "Nombre de ventes", "Chiffre d'affaires", "Évolution"],
                  lignes, 1, {3: FORMAT_EURO, 4: "+0.0%;-0.0%;0.0%"})
    zone = f"D2:D{fin}"
    feuille.conditional_formatting.add(zone, CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color="00703C", bold=True)))
    feuille.conditional_formatting.add(zone, CellIsRule(
        operator="lessThan", formula=["0"], font=Font(color="C00000", bold=True)))
    largeurs(feuille, {"A": 14, "B": 18, "C": 20, "D": 12})
    return feuille, fin


def feuille_donnees(classeur, analyse):
    feuille = classeur.create_sheet("Données")
    lignes = [[v.date, v.produit, v.categorie, v.quantite, v.prix, v.total] for v in analyse.ventes]
    tableau(feuille, "Ventes", ["Date", "Produit", "Catégorie", "Quantité", "Prix unitaire", "Total"],
            lignes, 1, {1: "DD/MM/YYYY", 5: FORMAT_EURO, 6: FORMAT_EURO})
    largeurs(feuille, {"A": 12, "B": 28, "C": 16, "D": 11, "E": 14, "F": 14})


def feuille_anomalies(classeur, anomalies):
    feuille = classeur.create_sheet("Anomalies")
    feuille.sheet_properties.tabColor = "C00000"
    lignes = [[a.ligne, a.raison, a.contenu] for a in anomalies]
    tableau(feuille, "Anomalies", ["Ligne du fichier", "Problème", "Contenu de la ligne"], lignes, 1, {})
    largeurs(feuille, {"A": 16, "B": 40, "C": 60})
    feuille.page_setup.orientation = "landscape"
    feuille.sheet_properties.pageSetUpPr.fitToPage = True
    feuille.page_setup.fitToHeight = 0


def carte(feuille, colonne, libelle, valeur, format_nombre=None):
    """Carte d'indicateur sur 3 colonnes : petit libellé gris + grande valeur."""
    fin = get_column_letter(colonne + 2)
    debut = get_column_letter(colonne)
    fond = PatternFill("solid", fgColor=FOND_CARTE)
    bordure = Side(style="thick", color=BLEU)
    for ligne in (5, 6):
        feuille.merge_cells(f"{debut}{ligne}:{fin}{ligne}")
        for col in range(colonne, colonne + 3):
            feuille.cell(row=ligne, column=col).fill = fond
        feuille.cell(row=ligne, column=colonne).border = Border(left=bordure)

    titre = feuille[f"{debut}5"]
    titre.value = libelle.upper()
    titre.font = Font(size=9, bold=True, color=GRIS)
    titre.alignment = Alignment(indent=1, vertical="bottom")

    cellule = feuille[f"{debut}6"]
    cellule.value = valeur
    cellule.font = Font(size=20, bold=True, color=BLEU)
    cellule.alignment = Alignment(indent=1, vertical="center", horizontal="left")
    if format_nombre:
        cellule.number_format = format_nombre


def feuille_tableau_de_bord(classeur, analyse, anomalies, titre, produits, col_ca, fin_produits,
                            mois, fin_mois):
    feuille = classeur.active
    feuille.title = "Tableau de bord"
    feuille.sheet_view.showGridLines = False
    feuille.sheet_properties.tabColor = BLEU
    largeurs(feuille, {"A": 2, **{get_column_letter(c): 11 for c in range(2, 14)}})

    texte_titre = "Tableau de bord des ventes" + (f" — {titre}" if titre else "")
    bandeau(feuille, "B2:M2", texte_titre)
    feuille.row_dimensions[2].height = 34
    feuille.merge_cells("B3:M3")
    sous_titre = feuille["B3"]
    sous_titre.value = (f"Période du {analyse.debut:%d/%m/%Y} au {analyse.fin:%d/%m/%Y}"
                        f"  ·  {analyse.nb_ventes} ventes analysées"
                        f"  ·  généré le {datetime.now():%d/%m/%Y à %H:%M}")
    sous_titre.font = Font(italic=True, color=GRIS)
    sous_titre.alignment = Alignment(indent=1)

    feuille.row_dimensions[5].height = 20
    feuille.row_dimensions[6].height = 36
    carte(feuille, 2, "Chiffre d'affaires", analyse.chiffre_affaires, FORMAT_EURO)
    carte(feuille, 5, "Nombre de ventes", analyse.nb_ventes, "#,##0")
    carte(feuille, 8, "Panier moyen", analyse.panier_moyen, FORMAT_EURO)
    carte(feuille, 11, "Articles vendus", analyse.quantite_totale, "#,##0")

    # Faits marquants : libellé en gras + phrase
    feuille["B8"].value = "Faits marquants"
    feuille["B8"].font = Font(bold=True, size=12, color=BLEU)
    premier = analyse.produits[0]
    faits = [
        ("Produit n°1 : ", f"{premier.nom} — {euros(premier.total)} "
                          f"({premier.part:.0%} du chiffre d'affaires)", BLEU),
        ("Meilleur mois : ", f"{analyse.meilleur_mois.libelle} — "
                            f"{euros(analyse.meilleur_mois.total)}", "00703C"),
        ("Mois le plus faible : ", f"{analyse.mois_le_plus_faible.libelle} — "
                                  f"{euros(analyse.mois_le_plus_faible.total)}", "C00000"),
    ]
    if anomalies:
        faits.append(("Attention : ", f"{len(anomalies)} ligne(s) du fichier contiennent des erreurs, "
                                      f"détaillées dans l'onglet « Anomalies »", "C55A11"))
    else:
        faits.append(("Qualité des données : ", "aucune erreur détectée dans le fichier", "00703C"))
    for i, (libelle, texte, couleur) in enumerate(faits, start=9):
        feuille.merge_cells(f"B{i}:M{i}")
        feuille[f"B{i}"].value = CellRichText(
            TextBlock(InlineFont(b=True, color=couleur), "■  " + libelle), texte)
        feuille[f"B{i}"].alignment = Alignment(indent=1)

    # Graphique 1 : évolution mensuelle
    courbe = LineChart()
    courbe.title = "Évolution du chiffre d'affaires (€)"
    courbe.title.overlay = False
    courbe.y_axis.numFmt = "#,##0"
    courbe.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="D9D9D9"))
    courbe.legend = None
    courbe.add_data(Reference(mois, min_col=3, min_row=1, max_row=fin_mois), titles_from_data=True)
    courbe.set_categories(Reference(mois, min_col=1, min_row=2, max_row=fin_mois))
    courbe.varyColors = False
    serie = courbe.series[0]
    serie.smooth = False
    serie.graphicalProperties.line.width = 28000
    serie.graphicalProperties.line.solidFill = "2F5597"
    serie.marker.symbol = "circle"
    serie.marker.size = 6
    serie.marker.graphicalProperties = GraphicalProperties(
        solidFill="2F5597", ln=LineProperties(solidFill="2F5597"))
    courbe.y_axis.delete = False
    courbe.x_axis.delete = False
    courbe.width, courbe.height = 12.8, 8
    feuille.add_chart(courbe, "B15")

    # Graphique 2 : répartition par produit (pourcentages seulement, légende à droite)
    camembert = PieChart()
    camembert.title = "Répartition du chiffre d'affaires"
    camembert.title.overlay = False
    camembert.add_data(Reference(produits, min_col=col_ca, min_row=1, max_row=fin_produits),
                       titles_from_data=True)
    camembert.set_categories(Reference(produits, min_col=2, min_row=2, max_row=fin_produits))
    camembert.dataLabels = DataLabelList(showPercent=True, showVal=False, showCatName=False,
                                         showSerName=False, showLegendKey=False)
    camembert.legend.position = "r"
    # Camembert à gauche, légende à droite, sans chevauchement
    camembert.layout = Layout(manualLayout=ManualLayout(
        layoutTarget="inner", xMode="edge", yMode="edge", x=0.06, y=0.17, w=0.44, h=0.76))
    camembert.legend.layout = Layout(manualLayout=ManualLayout(
        xMode="edge", yMode="edge", x=0.58, y=0.2, w=0.4, h=0.72))
    camembert.width, camembert.height = 13, 8
    feuille.add_chart(camembert, "H15")

    # Graphique 3 : classement des produits
    barres = BarChart()
    barres.type = "bar"
    barres.title = "Chiffre d'affaires par produit (€)"
    barres.title.overlay = False
    barres.legend = None
    barres.add_data(Reference(produits, min_col=col_ca, min_row=1, max_row=fin_produits),
                    titles_from_data=True)
    barres.set_categories(Reference(produits, min_col=2, min_row=2, max_row=fin_produits))
    barres.x_axis.scaling.orientation = "maxMin"  # n°1 en haut
    barres.y_axis.crosses = "max"                 # axe des montants en bas
    barres.y_axis.numFmt = "#,##0"
    barres.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="D9D9D9"))
    barres.y_axis.delete = False
    barres.x_axis.delete = False
    barres.gapWidth = 60
    barres.series[0].graphicalProperties.solidFill = "2F5597"
    barres.dataLabels = DataLabelList(showVal=True, showPercent=False, showCatName=False,
                                      showSerName=False, showLegendKey=False)
    barres.dataLabels.numFmt = '#,##0 "€"'
    barres.width, barres.height = 26, 9
    feuille.add_chart(barres, "B32")

    # Impression : une page en paysage
    feuille.page_setup.orientation = "landscape"
    feuille.page_setup.fitToWidth = 1
    feuille.page_setup.fitToHeight = 1
    feuille.sheet_properties.pageSetUpPr.fitToPage = True


def creer_rapport(analyse, anomalies, chemin, titre=""):
    classeur = Workbook()
    produits, col_ca, fin_produits = feuille_produits(classeur, analyse)
    mois, fin_mois = feuille_mois(classeur, analyse)
    feuille_donnees(classeur, analyse)
    if anomalies:
        feuille_anomalies(classeur, anomalies)
    feuille_tableau_de_bord(classeur, analyse, anomalies, titre, produits, col_ca, fin_produits,
                            mois, fin_mois)
    classeur.save(chemin)


# ---------------------------------------------------------------------------
# Programme principal
# ---------------------------------------------------------------------------

def main(arguments=None):
    parseur = argparse.ArgumentParser(
        description="Génère un rapport Excel professionnel à partir d'un fichier de ventes.")
    parseur.add_argument("fichier", help="fichier des ventes (.xlsx ou .csv)")
    parseur.add_argument("-o", "--sortie", default="rapport.xlsx",
                         help="nom du rapport (défaut : rapport.xlsx)")
    parseur.add_argument("-t", "--titre", default="", help="nom de l'entreprise, affiché en en-tête")
    arguments = parseur.parse_args(arguments)

    print(f"Lecture de {arguments.fichier}…")
    try:
        ventes, anomalies = lire_ventes(arguments.fichier)
    except FileNotFoundError:
        sys.exit(f"Erreur : le fichier « {arguments.fichier} » est introuvable.")
    except ValueError as erreur:
        sys.exit(f"Erreur : fichier non reconnu ({erreur}).")

    if not ventes:
        sys.exit("Erreur : aucune vente valide dans le fichier.")

    analyse = analyser(ventes)
    try:
        creer_rapport(analyse, anomalies, arguments.sortie, arguments.titre)
    except PermissionError:
        sys.exit(f"Erreur : impossible d'écrire « {arguments.sortie} ». Est-il ouvert dans Excel ?")

    print(f"\nRapport créé : {arguments.sortie}")
    print(f"  Période            : {analyse.debut:%d/%m/%Y} → {analyse.fin:%d/%m/%Y}")
    print(f"  Chiffre d'affaires : {euros(analyse.chiffre_affaires)}")
    print(f"  Ventes             : {analyse.nb_ventes}  (panier moyen {euros(analyse.panier_moyen)})")
    print(f"  Produit n°1        : {analyse.produits[0].nom}")
    print(f"  Meilleur mois      : {analyse.meilleur_mois.libelle}")
    if anomalies:
        print(f"  ⚠ {len(anomalies)} ligne(s) en erreur, détaillées dans l'onglet « Anomalies »")


if __name__ == "__main__":
    main()
