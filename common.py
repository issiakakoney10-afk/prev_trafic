"""Constantes et petits utilitaires partagés par tout le moteur.

Tout ce qui est une coordonnée de cellule ici a été vérifié à la main en
lisant les formules réelles du fichier maître (pas une supposition) :
voir schema/input_cells.json et l'analyse menée en amont.
"""
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
MASTER = BASE / "schema" / "master_reference.xlsx"
TEMPLATES_DIR = BASE / "templates"
OUTPUT_DIR = BASE / "output"

LEGACY_SHEETS = {"ANNUEL_DCAQ_AVEC_TC2_Option2"}  # copiées telles quelles, jamais modifiées

MONTHS_FR = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
             "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]

# Colonnes du "détail mensuel" (grille C:N) et du bloc "Annoncé DOMSE" (U:AF),
# dans l'ordre janvier -> décembre.
DIRECT_MONTH_COLS = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
ANNOUNCED_MONTH_COLS = ["U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF"]

# Les 2 ancres mensuelles "chiffre annoncé" et la ligne où elles vivent dans
# 'Trafic mensXXXX' (vérifié : T20/T73 = en-têtes "Annoncé DOMSE",
# la ligne juste en dessous porte les valeurs U:AF).
ANCHOR_ROWS = {
    "escales": 21,          # 'Escale' - nombre total d'escales de navires
    "tonnage_global": 74,   # 'TG' - trafic global en tonnes
}

# CALC. PREV : ligne "Réalisé <année>" de chacun des 7 indicateurs (espacées
# régulièrement de 12 lignes), et décalage de ligne pour chacune des 3
# méthodes de projection utilisées pour les mois non réalisés.
CALC_PREV_REALISE_ROWS = {
    "escales": 6,
    "tonnage_global": 18,
    "national": 30,
    "transit": 42,
    "transbordement": 54,
    "conteneurs_teu": 66,
    "conteneurs_transbordes_teu": 78,
}
METHOD_OFFSETS = {
    "sur_la_base_du_realise": 4,
    "correction_des_previsions": 5,
    "coefficients_saisonniers": 6,
}
METHOD_LABELS = {
    "sur_la_base_du_realise": "Sur la base du réalisé",
    "correction_des_previsions": "Correction des prévisions",
    "coefficients_saisonniers": "Prévision avec les coefficients saisonniers",
}
# Colonnes janvier -> décembre dans CALC. PREV (décalées d'une lettre par
# rapport à 'Trafic mensXXXX' qui commence à C).
CALC_PREV_MONTH_COLS = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]

# Pour chaque indicateur, la ligne de 'Trafic mensXXXX' utilisée par
# CALC. PREV pour (a) un mois où le détail complet est saisi (colonnes C:N)
# et (b) un mois où seul le chiffre annoncé est disponible (colonnes U:AF).
# Vérifié cellule par cellule sur le fichier maître.
CALC_PREV_TRAFIC_MENS_ROWS = {
    "escales": {"direct": 21, "announced": 21},
    "tonnage_global": {"direct": 73, "announced": 74},
    "national": {"direct": 1069, "announced": 1069},
    "transit": {"direct": 1077, "announced": 1077},
    "transbordement": {"direct": 1075, "announced": 1075},
    "conteneurs_teu": {"direct": 1201, "announced": 1204},
    "conteneurs_transbordes_teu": {"direct": 1205, "announced": 1205},
}
YEAR_SHEET = "Trafic mens2026"  # feuille de l'exercice en cours

# ANNUEL a sa PROPRE numérotation de lignes pour la ligne "Réalisé <année>"
# de chaque indicateur — CE N'EST PAS la même que CALC_PREV_REALISE_ROWS.
# Les 2 feuilles partagent leurs 12 premières lignes de bloc (escales,
# tonnage_global : blocs de 12 lignes dans les 2 feuilles, donc coïncidence
# pour ces 2-là), mais à partir de "national" les blocs d'ANNUEL ne font
# plus que 10 lignes (au lieu de 12 dans CALC. PREV) : un vrai écart du
# fichier d'origine, pas une erreur de saisie. Vérifié le 2026-09-17 en
# comparant, pour chacun des 7 indicateurs, la ligne "Réalisé <année>" de
# ANNUEL et la formule qu'elle contient déjà dans le fichier de référence
# (ex. ANNUEL!B28 "Réalisé 2026" national -> formule '=CALC. PREV'!B30',
# donc ANNUEL national réalisé = ligne 28, alors que CALC_PREV_REALISE_ROWS
# donne 30 pour CALC. PREV). Avant cette correction, _regenerate_annuel()
# écrivait par erreur dans la ligne CALC_PREV_REALISE_ROWS (donc la ligne
# "Prévision <année>" pour 'national', et des lignes vides et jamais lues
# pour 'transit'/'transbordement'/'conteneurs_teu'/'conteneurs_transbordes_teu')
# au lieu de la vraie ligne "Réalisé <année>" — la frontière réalisé/projeté
# n'avait donc jamais d'effet réel sur ANNUEL pour ces 5 indicateurs.
ANNUEL_REALISE_ROWS = {
    "escales": 6,
    "tonnage_global": 18,
    "national": 28,
    "transit": 38,
    "transbordement": 50,
    "conteneurs_teu": 62,
    "conteneurs_transbordes_teu": 74,
}

# COEFFICIENTS_SAISONNIERS_2026 : objectifs annuels indépendants (les 2 autres
# catégories, marchandises générales et transbordement, sont CALCULÉES par
# différence dans le fichier -> ne pas les exposer comme saisie).
OBJECTIFS_ANNUELS = [
    ("J4", "Escales de navires (nombre)"),
    ("J24", "Trafic global (tonnes) - objectif DG"),
    ("J44", "Produits pétroliers, national + transit (tonnes)"),
    ("J84", "Produits de la pêche hors transbordement (tonnes)"),
    ("J104", "Trafic national (tonnes) - objectif DG"),
    ("J144", "Trafic transit (tonnes) - objectif DG"),
    ("J164", "Conteneurs, total (EVP)"),
    ("J184", "Conteneurs transbordés (EVP)"),
]
OBJECTIFS_SHEET = "COEFFICIENTS_SAISONNIERS_2026"

# NON-RESPECT_ENGA_TRANSBO_TC2 : engagement contractuel de trafic de
# conteneurs transbordés (business plan TC2), ligne 35, une colonne par année.
TC2_SHEET = "NON-RESPECT_ENGA_TRANSBO_TC2"
TC2_ENGAGEMENT_ROW = 35
TC2_YEARS_COLS = {2025: "G", 2026: "H", 2027: "I", 2028: "J", 2029: "K", 2030: "L"}


def is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")
