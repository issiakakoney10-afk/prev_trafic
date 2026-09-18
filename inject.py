"""
Le moteur : prend le fichier maître + tout ce que l'utilisateur a fourni
(modèles remplis, chiffres annoncés, choix de méthode...) et produit un
nouveau classeur, au même format, avec les valeurs à jour et tout recalculé.

Règle d'or : on n'écrit JAMAIS dans une cellule qui est une formule dans le
fichier de référence. Le fichier de référence (schema/master_reference.xlsx)
sert justement à vérifier ça avant chaque écriture.
"""
import datetime as dt
import shutil
import subprocess
import tempfile
import uuid
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl

from common import (
    MASTER, OUTPUT_DIR, is_formula,
    ANCHOR_ROWS, ANNOUNCED_MONTH_COLS, ANNUEL_REALISE_ROWS, DIRECT_MONTH_COLS,
    CALC_PREV_MONTH_COLS, CALC_PREV_REALISE_ROWS, CALC_PREV_TRAFIC_MENS_ROWS,
    METHOD_OFFSETS, TC2_SHEET, TC2_ENGAGEMENT_ROW, TC2_YEARS_COLS, YEAR_SHEET,
)

CALC_PREV_SHEET = "CALC. PREV"
ANNUEL_SHEET = "ANNUEL"


@dataclass
class GenerationReport:
    messages: list = field(default_factory=list)
    n_cells_written: int = 0

    def add(self, msg):
        self.messages.append(msg)


def _copy_sheet_values(report, ref_ws, src_ws, dst_ws, sheet_label):
    """Copie src_ws -> dst_ws cellule par cellule, uniquement là où ref_ws
    dit que ce n'est PAS une formule (donc une cellule de saisie légitime),
    et seulement quand src_ws propose une valeur non vide."""
    max_row = max(ref_ws.max_row, src_ws.max_row)
    max_col = max(ref_ws.max_column, src_ws.max_column)
    n = 0
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            ref_val = ref_ws.cell(row=row, column=col).value
            if is_formula(ref_val):
                continue
            new_val = src_ws.cell(row=row, column=col).value
            if new_val is None:
                continue
            dst_ws.cell(row=row, column=col).value = new_val
            n += 1
    report.add(f"{sheet_label} : {n} cellules mises à jour")
    report.n_cells_written += n


def _apply_trafic_mens_template(report, wb, ref_wb, path: Path):
    tpl = openpyxl.load_workbook(path, data_only=False)
    if YEAR_SHEET not in tpl.sheetnames:
        raise ValueError(
            f"Le fichier '{path.name}' ne contient pas l'onglet '{YEAR_SHEET}' "
            "attendu pour le détail mensuel."
        )
    _copy_sheet_values(report, ref_wb[YEAR_SHEET], tpl[YEAR_SHEET], wb[YEAR_SHEET],
                        "Détail mensuel (Trafic mens2026)")


def _apply_parametres_template(report, wb, ref_wb, path: Path):
    tpl = openpyxl.load_workbook(path, data_only=False)
    for sheet in ("COEFFICIENTS_SAISONNIERS_2026", "Stats exo 2025brute"):
        if sheet in tpl.sheetnames:
            _copy_sheet_values(report, ref_wb[sheet], tpl[sheet], wb[sheet], sheet)
    tc2_sheet_name = "TC2 - Engagement contractuel"
    if tc2_sheet_name in tpl.sheetnames:
        ws_tc2 = tpl[tc2_sheet_name]
        n = 0
        for row in range(4, 4 + len(TC2_YEARS_COLS)):
            year = ws_tc2.cell(row=row, column=1).value
            val = ws_tc2.cell(row=row, column=2).value
            if year in TC2_YEARS_COLS and val is not None:
                col_letter = TC2_YEARS_COLS[year]
                wb[TC2_SHEET][f"{col_letter}{TC2_ENGAGEMENT_ROW}"] = val
                n += 1
        report.add(f"Engagement contractuel TC2 : {n} années mises à jour")
        report.n_cells_written += n


def _apply_monthly_anchor(report, wb, month: int, escales, tonnage_global):
    col = ANNOUNCED_MONTH_COLS[month - 1]
    n = 0
    if escales is not None:
        wb[YEAR_SHEET][f"{col}{ANCHOR_ROWS['escales']}"] = escales
        n += 1
    if tonnage_global is not None:
        wb[YEAR_SHEET][f"{col}{ANCHOR_ROWS['tonnage_global']}"] = tonnage_global
        n += 1
    report.add(f"Chiffres annoncés du mois {month:02d} : {n} valeur(s) enregistrée(s)")
    report.n_cells_written += n


def _regenerate_calc_prev_boundaries(report, wb, last_full_detail_month: int):
    """Pour chacun des 7 indicateurs, pointe chaque mois <= last_full_detail_month
    vers la ligne 'détail complet' de Trafic mens2026, et chaque mois suivant
    vers la ligne 'chiffre annoncé' (qui vaut 0 tant qu'elle n'est pas
    renseignée). Remplace ce que l'analyste devait faire à la main chaque
    mois."""
    ws = wb[CALC_PREV_SHEET]
    n = 0
    for kpi, realise_row in CALC_PREV_REALISE_ROWS.items():
        rows = CALC_PREV_TRAFIC_MENS_ROWS[kpi]
        for month_idx in range(1, 13):
            col = CALC_PREV_MONTH_COLS[month_idx - 1]
            if month_idx <= last_full_detail_month:
                src_col = DIRECT_MONTH_COLS[month_idx - 1]
                src_row = rows["direct"]
            else:
                src_col = ANNOUNCED_MONTH_COLS[month_idx - 1]
                src_row = rows["announced"]
            formula = f"='{YEAR_SHEET}'!{src_col}{src_row}"
            cell = ws[f"{col}{realise_row}"]
            if cell.value != formula:
                cell.value = formula
                n += 1
    report.add(f"CALC. PREV : {n} références mensuelles réalignées "
               f"(détail complet jusqu'au mois {last_full_detail_month:02d})")
    report.n_cells_written += n


def _regenerate_annuel(report, wb, last_known_month: int, method_key: str):
    """Dans ANNUEL, pour chacun des 7 indicateurs : les mois <= last_known_month
    (au moins un chiffre connu, détaillé ou annoncé) pointent directement
    vers la ligne 'Réalisé' de CALC. PREV ; les mois suivants pointent vers
    la ligne de la méthode de projection choisie. ANNUEL a sa PROPRE
    frontière réalisé/projeté (généralement plus tardive que celle de
    CALC. PREV, qui elle ne distingue que détail complet / chiffre annoncé).

    Important : la ligne à ÉCRIRE dans ANNUEL (ANNUEL_REALISE_ROWS) n'est PAS
    la même chose que la ligne à LIRE dans CALC. PREV (CALC_PREV_REALISE_ROWS)
    — les deux feuilles ne partagent la même numérotation que pour les 2
    premiers indicateurs (escales, tonnage_global) ; à partir de 'national'
    les blocs d'ANNUEL sont plus courts de 2 lignes dans le fichier
    d'origine. Confondre les deux (comme le faisait une version antérieure)
    écrit dans la mauvaise ligne d'ANNUEL et laisse la vraie ligne
    'Réalisé <année>' figée sur son ancienne formule."""
    offset = METHOD_OFFSETS[method_key]
    ws = wb[ANNUEL_SHEET]
    n = 0
    for kpi, calc_prev_realise_row in CALC_PREV_REALISE_ROWS.items():
        annuel_row = ANNUEL_REALISE_ROWS[kpi]
        for month_idx in range(1, 13):
            col = CALC_PREV_MONTH_COLS[month_idx - 1]
            target_row = (calc_prev_realise_row if month_idx <= last_known_month
                          else calc_prev_realise_row + offset)
            formula = f"='CALC. PREV'!{col}{target_row}"
            cell = ws[f"{col}{annuel_row}"]
            if cell.value != formula:
                cell.value = formula
                n += 1
    report.add(f"ANNUEL : {n} cellules réalignées (connu jusqu'au mois "
               f"{last_known_month:02d}, méthode « {method_key} » au-delà)")
    report.n_cells_written += n


def _recalculate(src_path: Path) -> Path:
    """Convertit src_path avec LibreOffice pour forcer le recalcul de toutes
    les formules. Dossier de sortie ET profil utilisateur LibreOffice
    (-env:UserInstallation) uniques à cet appel : par défaut, LibreOffice
    verrouille un profil unique partagé et une 2e conversion simultanée
    échoue silencieusement (constaté en test) — utile pour la version
    hébergée, où plusieurs personnes peuvent cliquer sur 'Générer' au même
    moment."""
    out_dir = Path(tempfile.mkdtemp(prefix="recalc_", dir=src_path.parent))
    profile_dir = Path(tempfile.mkdtemp(prefix="lo_profile_", dir=src_path.parent))
    try:
        result = subprocess.run(
            [
                "soffice", "--headless", "--norestore",
                f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to", "xlsx", "--outdir", str(out_dir), str(src_path),
            ],
            capture_output=True, text=True, timeout=180,
        )
        recalced = out_dir / src_path.name
        if result.returncode != 0 or not recalced.exists():
            raise RuntimeError(
                "Échec du recalcul LibreOffice.\n"
                f"stdout: {result.stdout}\nstderr: {result.stderr}"
            )
        moved = src_path.parent / f"_recalced_{out_dir.name}.xlsx"
        shutil.move(str(recalced), str(moved))
        return moved
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)
        shutil.rmtree(profile_dir, ignore_errors=True)


def generate_output(
    trafic_mens_template_path: Path | None = None,
    parametres_template_path: Path | None = None,
    monthly_anchor: dict | None = None,      # {"month": int, "escales": float|None, "tonnage_global": float|None}
    last_full_detail_month: int | None = None,   # pilote CALC. PREV (direct vs annoncé)
    last_known_month: int | None = None,         # pilote ANNUEL (réalisé vs méthode de projection)
    annuel_method: str | None = None,
    output_name: str | None = None,
) -> tuple[Path, GenerationReport]:
    report = GenerationReport()
    wb = openpyxl.load_workbook(MASTER, data_only=False)
    ref_wb = openpyxl.load_workbook(MASTER, data_only=False)  # référence, jamais modifiée

    if trafic_mens_template_path:
        _apply_trafic_mens_template(report, wb, ref_wb, Path(trafic_mens_template_path))
    if parametres_template_path:
        _apply_parametres_template(report, wb, ref_wb, Path(parametres_template_path))
    if monthly_anchor:
        _apply_monthly_anchor(report, wb, monthly_anchor["month"],
                               monthly_anchor.get("escales"), monthly_anchor.get("tonnage_global"))
    if last_full_detail_month:
        _regenerate_calc_prev_boundaries(report, wb, last_full_detail_month)
    if last_known_month and annuel_method:
        _regenerate_annuel(report, wb, last_known_month, annuel_method)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    uniq = uuid.uuid4().hex[:8]  # évite toute collision entre générations simultanées
    name = output_name or f"Previsions_de_trafics_{stamp}.xlsx"
    pre_recalc_path = OUTPUT_DIR / f"_pre_recalc_{stamp}_{uniq}.xlsx"
    wb.save(pre_recalc_path)

    report.add("Recalcul des formules (LibreOffice)…")
    recalced_path = _recalculate(pre_recalc_path)
    final_path = OUTPUT_DIR / name
    if final_path.exists():
        # 2 générations avec le même nom (auto ou choisi) : on ne réécrit jamais
        # un fichier existant, on suffixe plutôt.
        final_path = OUTPUT_DIR / f"{final_path.stem}_{uniq}{final_path.suffix}"
    shutil.move(str(recalced_path), str(final_path))
    pre_recalc_path.unlink(missing_ok=True)

    report.add(f"Fichier généré : {final_path.name}")
    return final_path, report
