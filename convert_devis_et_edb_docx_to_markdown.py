#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversion DOC/DOCX -> Markdown (dataset pour training LLM)
- NDC (colonne G) + EDB (colonne F)
- Filtre Excel : B=1, C=1, D=OUI, E=NON
- Utilise Mammoth pour conversion propre (ignore headers/footers automatiquement)
- Ignore page de garde, synthèse, table des matières
- Préserve titres, paragraphes, listes, tableaux
- Format Markdown homogène pour training

Dépendances:
  pip install pandas openpyxl mammoth html2text
Optionnel (pour .doc):
  pip install pywin32 ou avoir LibreOffice
"""

from __future__ import annotations

import re
import os
import shutil
import subprocess
import traceback
import logging
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import mammoth
import html2text

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ==============================================================================
# CONFIGURATION - MODIFIEZ ICI
# ==============================================================================

# Fichier Excel source
EXCEL_NAME = "couverture_EDB_NDC_par_RITM.xlsx"

# Colonnes contenant les chemins des fichiers (index 0-based: A=0, B=1, etc.)
COL_EDB = 5  # Colonne F = index 5
COL_NDC = 6  # Colonne G = index 6

# ------------------------------
# FILTRES EXCEL
# ------------------------------
# Chaque filtre est un tuple: (colonne_index, valeur_attendue)
# - colonne_index: 0=A, 1=B, 2=C, 3=D, 4=E, etc.
# - valeur_attendue: la valeur que doit avoir la cellule (ou None pour ignorer)
#
# Pour DÉSACTIVER un filtre: mettez None comme valeur
# Pour RETIRER un filtre: supprimez la ligne ou commentez-la
#
# Exemple: (1, 1) signifie "colonne B doit égaler 1"
# Exemple: (3, "OUI") signifie "colonne D doit égaler 'OUI'"
# Exemple: (1, None) signifie "pas de filtre sur colonne B"

EXCEL_FILTERS = [
    (1, 1),       # Colonne B = 1        (mettre None pour désactiver)
    (2, 1),       # Colonne C = 1        (mettre None pour désactiver)
    (3, "OUI"),   # Colonne D = "OUI"    (mettre None pour désactiver)
    (4, "NON"),   # Colonne E = "NON"    (mettre None pour désactiver)
]

# Pour désactiver TOUS les filtres, décommentez la ligne suivante:
# EXCEL_FILTERS = []

# ------------------------------
# Dossiers de sortie
# ------------------------------
OUTPUT_DIRNAME = "dataset_markdown"
LOG_DIRNAME = "_logs"
SUBDIR_NDC = "ndc"
SUBDIR_EDB = "edb"

# Style mapping Mammoth : ignorer les styles de TOC et mapper les autres
MAMMOTH_STYLE_MAP = """
p[style-name='toc 1'] => !
p[style-name='toc 2'] => !
p[style-name='toc 3'] => !
p[style-name='toc 4'] => !
p[style-name='TM1'] => !
p[style-name='TM2'] => !
p[style-name='TM3'] => !
p[style-name='TM4'] => !
p[style-name='TOC Heading'] => !
p[style-name='TOC 1'] => !
p[style-name='TOC 2'] => !
p[style-name='TOC 3'] => !
p[style-name='Title'] => h1
p[style-name='Titre'] => h1
p[style-name='Heading 1'] => h1
p[style-name='Heading 2'] => h2
p[style-name='Heading 3'] => h3
p[style-name='Heading 4'] => h4
p[style-name='Titre 1'] => h1
p[style-name='Titre 2'] => h2
p[style-name='Titre 3'] => h3
p[style-name='Titre 4'] => h4
p[style-name='List Paragraph'] => p
p[style-name='Paragraphedeliste'] => p
p[style-name='No Spacing'] => p
p[style-name='Body Text'] => p
"""

# Patterns pour détecter le début du vrai contenu
# On cherche un titre de chapitre principal (I., II., 1., 2., ou connu)
CHAPTER_START_PATTERNS = [
    # Titres avec numérotation romaine
    r'^#{1,2}\s+[IVXLCDM]+\.\s+',
    r'^#{1,2}\s+[IVXLCDM]+\s+',
    # Titres avec numérotation décimale
    r'^#{1,2}\s+\d+\.\s+',
    r'^#{1,2}\s+\d+\s+',
    # Titres connus sans numérotation
    r'^#\s+Description\s+du\s+projet',
    r'^#\s+Introduction',
    r'^#\s+Contexte',
    r'^#\s+Périmètre',
    r'^#\s+Perimetre',
]

# Patterns pour détecter la fin du préambule/TOC
TOC_END_MARKERS = [
    "Table des matières",
    "Table des matieres",
    "Sommaire",
    "TABLE DES MATIÈRES",
    "SOMMAIRE",
]

# Patterns pour nettoyer le contenu indésirable
CLEANUP_PATTERNS = [
    # Lignes de TOC avec numéros de page (ex: "I.1. Contexte 4")
    (r'^[IVXLCDM]+(?:\.\d+)*\.?\s+.+?\s+\d+\s*$', '', re.MULTILINE),
    (r'^\d+(?:\.\d+)*\.?\s+.+?\s+\d+\s*$', '', re.MULTILINE),
    # Liens de TOC avec ancres
    (r'\[([^\]]+)\]\(#[^)]+\)', r'\1', 0),
    # Images en base64
    (r'!\[.*?\]\(data:image[^)]+\)', '', 0),
    # Lignes vides multiples
    (r'\n{4,}', '\n\n\n', 0),
]


# ------------------------------
# Conversion DOC -> DOCX
# ------------------------------
def convert_doc_to_docx(input_doc: Path, workdir: Path) -> Optional[Path]:
    """Convertit un fichier .doc en .docx."""
    workdir.mkdir(parents=True, exist_ok=True)
    out_docx = workdir / (input_doc.stem + ".docx")

    if out_docx.exists():
        return out_docx

    if os.name == "nt":
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(input_doc.resolve()))
            doc.SaveAs(str(out_docx.resolve()), FileFormat=16)
            doc.Close(False)
            word.Quit()
            if out_docx.exists():
                return out_docx
        except Exception:
            pass

    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if soffice:
        try:
            cmd = [soffice, "--headless", "--nologo", "--convert-to", "docx",
                   "--outdir", str(workdir), str(input_doc)]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if out_docx.exists():
                return out_docx
        except Exception:
            pass

    return None


# ------------------------------
# Conversion DOCX -> Markdown
# ------------------------------
def clean_html_toc(html: str) -> str:
    """
    Supprime les éléments de TOC du HTML.
    Les entrées de TOC sont des liens vers #_Toc...
    """
    # Supprimer les paragraphes contenant des liens TOC
    # Pattern: <p>...<a href="#_Toc...">...</a>...</p>
    html = re.sub(
        r'<p[^>]*>\s*<a\s+href="#_Toc[^"]*"[^>]*>.*?</a>\s*</p>',
        '',
        html,
        flags=re.DOTALL | re.IGNORECASE
    )

    # Supprimer les liens TOC restants (qui pourraient être dans d'autres éléments)
    html = re.sub(
        r'<a\s+href="#_Toc[^"]*"[^>]*>(.*?)</a>',
        r'\1',
        html,
        flags=re.DOTALL | re.IGNORECASE
    )

    # Nettoyer les H1 qui contiennent "Table des matières" fusionné avec autre chose
    # Pattern: <h1>Table des matières...<a id="...">Vrai titre</h1>
    # On garde seulement le vrai titre
    def fix_toc_h1(match):
        content = match.group(1)
        # Chercher si "Table des matières" ou "Sommaire" est présent
        toc_patterns = [
            r'Table\s+des\s+mati[eè]res',
            r'Sommaire',
            r'TABLE\s+DES\s+MATI[EÈ]RES',
            r'SOMMAIRE'
        ]
        for pattern in toc_patterns:
            if re.search(pattern, content, re.IGNORECASE):
                # Supprimer la partie TOC et garder le reste après le dernier anchor
                # On cherche le dernier <a id="..."></a> et on garde ce qui suit
                parts = re.split(r'<a\s+id="[^"]*"[^>]*>\s*</a>', content)
                if len(parts) > 1 and parts[-1].strip():
                    return f'<h1>{parts[-1].strip()}</h1>'
                # Sinon supprimer tout le H1
                return ''
        return match.group(0)

    html = re.sub(r'<h1[^>]*>(.*?)</h1>', fix_toc_h1, html, flags=re.DOTALL | re.IGNORECASE)

    # Supprimer les ancres id="_Toc..." qui restent
    html = re.sub(r'<a\s+id="_Toc[^"]*"[^>]*>\s*</a>', '', html, flags=re.IGNORECASE)

    return html


def docx_to_markdown(docx_path: Path) -> str:
    """
    Convertit un fichier DOCX en Markdown propre.
    Utilise Mammoth pour la conversion HTML puis html2text.
    """
    # Convertir avec Mammoth
    with open(docx_path, 'rb') as f:
        result = mammoth.convert_to_html(
            f,
            style_map=MAMMOTH_STYLE_MAP,
            include_embedded_style_map=False,
        )
        html_content = result.value

    # Nettoyer le HTML (supprimer TOC)
    html_content = clean_html_toc(html_content)

    # Convertir HTML en Markdown
    h2t = html2text.HTML2Text()
    h2t.ignore_links = False
    h2t.ignore_images = True  # Ignorer les images
    h2t.ignore_emphasis = False
    h2t.body_width = 0  # Pas de wrap automatique
    h2t.unicode_snob = True
    h2t.skip_internal_links = True  # Ignorer les liens internes (#...)

    markdown = h2t.handle(html_content)

    # Post-traitement
    markdown = post_process_markdown(markdown)

    return markdown


def post_process_markdown(content: str) -> str:
    """
    Post-traite le Markdown pour :
    - Supprimer le préambule (page de garde, synthèse, etc.)
    - Supprimer la table des matières
    - Nettoyer le formatage
    """
    lines = content.split('\n')

    # Étape 1: Trouver le début du vrai contenu
    start_index = find_content_start(lines)

    # Garder uniquement le contenu principal
    if start_index > 0:
        lines = lines[start_index:]

    content = '\n'.join(lines)

    # Étape 2: Appliquer les patterns de nettoyage
    for pattern, replacement, flags in CLEANUP_PATTERNS:
        if flags:
            content = re.sub(pattern, replacement, content, flags=flags)
        else:
            content = re.sub(pattern, replacement, content)

    # Étape 3: Nettoyer les tableaux
    content = clean_tables(content)

    # Étape 4: Normaliser les titres
    content = normalize_headings(content)

    # Étape 5: Nettoyage final
    content = final_cleanup(content)

    return content


def find_content_start(lines: List[str]) -> int:
    """
    Trouve l'index de la première ligne du vrai contenu.
    Cherche après la table des matières.
    """
    toc_found = False
    toc_end_index = 0

    for i, line in enumerate(lines):
        line_stripped = line.strip()

        # Détecter la table des matières
        for marker in TOC_END_MARKERS:
            if marker.lower() in line_stripped.lower():
                toc_found = True
                toc_end_index = i
                break

    if toc_found:
        # Chercher le premier vrai titre après la TOC
        for i in range(toc_end_index + 1, len(lines)):
            line = lines[i].strip()

            # Ignorer les lignes vides et les lignes qui ressemblent à des entrées de TOC
            if not line:
                continue

            # Vérifier si c'est un vrai titre de chapitre
            if is_chapter_heading(line):
                return i

            # Ignorer les lignes de TOC (titre + numéro de page)
            if re.match(r'^[IVXLCDM\d]+\..*\d+\s*$', line):
                continue
            if re.match(r'^\*\*[IVXLCDM\d]+\.', line):
                continue

    # Si pas de TOC trouvée, chercher le premier titre de chapitre
    for i, line in enumerate(lines):
        if is_chapter_heading(line.strip()):
            return i

    return 0


def is_chapter_heading(line: str) -> bool:
    """
    Vérifie si une ligne est un titre de chapitre.
    Un titre Markdown H1/H2 qui ne finit PAS par un numéro de page.
    """
    # Doit commencer par # ou ##
    if not re.match(r'^#{1,2}\s+', line):
        return False

    # Ne doit pas finir par un numéro (entrée de TOC)
    if re.search(r'\s+\d+\s*$', line):
        return False

    # Ne doit pas être vide après le #
    title_text = re.sub(r'^#{1,2}\s+', '', line).strip()
    if not title_text or len(title_text) < 3:
        return False

    return True


def clean_tables(content: str) -> str:
    """Nettoie et normalise les tableaux Markdown."""
    lines = content.split('\n')
    result = []
    in_table = False
    table_lines = []

    for line in lines:
        stripped = line.strip()

        # Détecter un tableau (format html2text: "cell| cell" ou "---|---")
        is_table_line = False
        if '|' in stripped:
            # Ligne de tableau si elle contient | et n'est pas un titre Markdown
            if not stripped.startswith('#'):
                # Vérifier si c'est une ligne de séparateur ou de données
                if re.match(r'^[-|\s:]+$', stripped) or '|' in stripped:
                    is_table_line = True

        if is_table_line:
            if not in_table:
                in_table = True
                table_lines = []
            table_lines.append(line)
        else:
            if in_table:
                # Fin du tableau, le traiter
                processed_table = process_table(table_lines)
                result.extend(processed_table)
                result.append('')  # Ligne vide après le tableau
                in_table = False
                table_lines = []

            result.append(line)

    # Si on finit dans un tableau
    if in_table and table_lines:
        processed_table = process_table(table_lines)
        result.extend(processed_table)

    return '\n'.join(result)


def process_table(table_lines: List[str]) -> List[str]:
    """
    Traite un tableau pour le normaliser au format Markdown standard.
    Gère le format html2text (cell| cell) et le format standard (| cell | cell |)
    """
    if not table_lines:
        return []

    # Parser les lignes du tableau
    rows = []
    separator_idx = -1

    for i, line in enumerate(table_lines):
        line = line.strip()
        if not line:
            continue

        # Vérifier si c'est un séparateur (---|---)
        if re.match(r'^[-|\s:]+$', line) and '-' in line:
            separator_idx = len(rows)
            continue

        # Extraire les cellules
        # Format: "cell| cell" ou "| cell | cell |"
        if line.startswith('|'):
            line = line[1:]
        if line.endswith('|'):
            line = line[:-1]

        cells = [c.strip() for c in line.split('|')]
        if cells:
            rows.append(cells)

    if not rows:
        return []

    # Trouver le nombre max de colonnes
    max_cols = max(len(row) for row in rows)

    if max_cols == 0:
        return []

    # Normaliser toutes les lignes au même nombre de colonnes
    for row in rows:
        while len(row) < max_cols:
            row.append('')

    # Construire le tableau final
    result = []

    # Header (première ligne)
    header = '| ' + ' | '.join(rows[0]) + ' |'
    result.append(header)

    # Séparateur
    separator = '| ' + ' | '.join(['---'] * max_cols) + ' |'
    result.append(separator)

    # Lignes de données (à partir de la 2ème ligne)
    for row in rows[1:]:
        line = '| ' + ' | '.join(row) + ' |'
        result.append(line)

    return result


def normalize_headings(content: str) -> str:
    """
    Normalise les titres :
    - Assure un format cohérent
    - Corrige les niveaux si nécessaire
    """
    lines = content.split('\n')
    result = []

    for line in lines:
        # Détecter les titres
        match = re.match(r'^(#{1,6})\s+(.+)$', line)
        if match:
            level = len(match.group(1))
            title_text = match.group(2).strip()

            # Nettoyer le texte du titre
            # Enlever les ** au début/fin
            title_text = re.sub(r'^\*\*(.+)\*\*$', r'\1', title_text)

            # Reconstruire le titre
            line = '#' * level + ' ' + title_text

        result.append(line)

    return '\n'.join(result)


def final_cleanup(content: str) -> str:
    """Nettoyage final du contenu."""
    # Supprimer les lignes vides multiples
    content = re.sub(r'\n{3,}', '\n\n', content)

    # Supprimer les espaces en fin de ligne
    content = re.sub(r' +$', '', content, flags=re.MULTILINE)

    # Supprimer le contenu vide au début
    content = content.lstrip('\n')

    # Assurer une ligne vide à la fin
    content = content.rstrip() + '\n'

    # Supprimer les lignes qui ne contiennent que des caractères spéciaux
    lines = content.split('\n')
    cleaned_lines = []
    for line in lines:
        # Ignorer les lignes qui ne sont que des tirets, underscores, etc.
        if re.match(r'^[-_=\s|]+$', line) and '|' not in line:
            continue
        cleaned_lines.append(line)

    return '\n'.join(cleaned_lines)


# ------------------------------
# Chargement Excel
# ------------------------------
def load_targets_from_excel(excel_path: Path) -> Tuple[List[str], List[str]]:
    """Charge les fichiers à traiter depuis Excel en appliquant les filtres configurés."""
    df = pd.read_excel(excel_path, engine="openpyxl")

    # Construire le masque de filtre dynamiquement
    mask = pd.Series([True] * len(df))  # Commencer avec tout à True

    for col_idx, expected_value in EXCEL_FILTERS:
        if expected_value is None:
            # Filtre désactivé, on skip
            continue

        col_data = df.iloc[:, col_idx]

        # Pour les valeurs string, normaliser (strip + upper)
        if isinstance(expected_value, str):
            col_normalized = col_data.astype(str).str.strip().str.upper()
            expected_normalized = expected_value.strip().upper()
            mask = mask & (col_normalized == expected_normalized)
        else:
            # Pour les valeurs numériques
            mask = mask & (col_data == expected_value)

    # Extraire les colonnes EDB et NDC
    edb_col = df.iloc[:, COL_EDB]
    ndc_col = df.iloc[:, COL_NDC]

    edb = edb_col[mask].dropna().astype(str).tolist()
    ndc = ndc_col[mask].dropna().astype(str).tolist()

    def clean(lst):
        return [t.strip().strip('"').strip("'") for t in lst if t.strip()]

    # Log des filtres appliqués
    active_filters = [(col, val) for col, val in EXCEL_FILTERS if val is not None]
    if active_filters:
        filter_desc = ", ".join([f"col{col}={val}" for col, val in active_filters])
        logger.info(f"Filtres appliqués: {filter_desc}")
    else:
        logger.info("Aucun filtre appliqué (tous les fichiers seront traités)")

    return clean(ndc), clean(edb)


# ------------------------------
# Programme principal
# ------------------------------
def main() -> int:
    cwd = Path(".").resolve()
    excel_path = cwd / EXCEL_NAME

    base_out = cwd / OUTPUT_DIRNAME
    log_dir = base_out / LOG_DIRNAME
    out_ndc = base_out / SUBDIR_NDC
    out_edb = base_out / SUBDIR_EDB
    tmp_conv = base_out / "_tmp_doc_conversion"

    for d in [base_out, log_dir, out_ndc, out_edb, tmp_conv]:
        d.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        logger.error(f"Fichier Excel introuvable: {excel_path}")
        return 2

    ndc_list, edb_list = load_targets_from_excel(excel_path)

    logger.info(f"Fichiers NDC: {len(ndc_list)}")
    logger.info(f"Fichiers EDB: {len(edb_list)}")

    if not ndc_list and not edb_list:
        logger.info("Aucun fichier à traiter.")
        return 0

    report_rows = []
    stats = {"ok": 0, "error": 0, "missing": 0}

    def process_file(name: str, mode: str):
        src = cwd / name
        ext = src.suffix.lower()

        if not ext:
            src = src.with_suffix(".docx")
            ext = ".docx"

        out_dir = out_ndc if mode == "ndc" else out_edb

        if not src.exists():
            for try_ext in [".docx", ".doc", ".DOCX", ".DOC"]:
                alt = src.with_suffix(try_ext)
                if alt.exists():
                    src = alt
                    ext = try_ext.lower()
                    break

        if not src.exists():
            stats["missing"] += 1
            report_rows.append((mode, name, str(src), "", "MISSING", "Fichier introuvable"))
            logger.warning(f"Introuvable: {src}")
            return

        working_docx = src
        if ext == ".doc":
            converted = convert_doc_to_docx(src, tmp_conv)
            if not converted:
                stats["error"] += 1
                report_rows.append((mode, name, str(src), "", "CONVERT_ERROR", "Conversion .doc échouée"))
                logger.error(f"Conversion échouée: {src.name}")
                return
            working_docx = converted

        try:
            md_content = docx_to_markdown(working_docx)

            out_name = Path(name).stem + ".md"
            out_path = out_dir / out_name
            out_path.write_text(md_content, encoding="utf-8")

            stats["ok"] += 1
            report_rows.append((mode, name, str(src), str(out_path), "OK", ""))
            logger.info(f"[OK] ({mode}) {src.name} -> {out_path.name}")

        except Exception as ex:
            stats["error"] += 1
            err_msg = f"{type(ex).__name__}: {ex}"
            report_rows.append((mode, name, str(src), "", "ERROR", err_msg))

            trace = traceback.format_exc()
            log_file = log_dir / (Path(name).stem + f".{mode}.error.log")
            log_file.write_text(trace, encoding="utf-8")

            logger.error(f"[ERREUR] ({mode}) {src.name}: {err_msg}")

    for f in ndc_list:
        process_file(f, mode="ndc")

    for f in edb_list:
        process_file(f, mode="edb")

    report_df = pd.DataFrame(
        report_rows,
        columns=["type", "source_excel", "input_path", "output_md", "status", "error"]
    )
    report_path = base_out / "conversion_report.csv"
    report_df.to_csv(report_path, index=False, encoding="utf-8")

    logger.info("")
    logger.info("=== Résumé ===")
    logger.info(f"OK: {stats['ok']}")
    logger.info(f"Manquants: {stats['missing']}")
    logger.info(f"Erreurs: {stats['error']}")
    logger.info(f"Rapport: {report_path}")

    return 0 if stats["error"] == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
