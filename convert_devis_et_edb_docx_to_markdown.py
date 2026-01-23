#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversion DOC/DOCX/PDF -> Markdown (dataset pour training LLM)
- NDC (colonne G) + EDB (colonne F)
- Filtre Excel : B=1, C=1, D=OUI, E=NON
- Utilise Mammoth pour conversion DOCX (ignore headers/footers automatiquement)
- Utilise PyMuPDF pour conversion PDF
- Ignore page de garde, synthèse, table des matières
- Préserve titres, paragraphes, listes, tableaux
- Format Markdown homogène pour training

Dépendances:
  pip install pandas openpyxl mammoth html2text pymupdf4llm
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
import pymupdf

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
    (1, None),       # Colonne B = 1        (mettre None pour désactiver)
    (2, None),       # Colonne C = 1        (mettre None pour désactiver)
    (3, "OUI"),   # Colonne D = "OUI"    (mettre None pour désactiver)
    (4, None),   # Colonne E = "NON"    (mettre None pour désactiver)
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


# ------------------------------
# Conversion PDF -> Markdown
# ------------------------------
# Configuration PDF - marges pour ignorer headers/footers (en points, 72pt = 1 inch)
PDF_MARGIN_TOP = 60      # Marge haute pour ignorer header
PDF_MARGIN_BOTTOM = 50   # Marge basse pour ignorer footer/numéro de page
PDF_MARGIN_LEFT = 36     # Marge gauche
PDF_MARGIN_RIGHT = 36    # Marge droite


def pdf_to_markdown(pdf_path: Path) -> str:
    """
    Convertit un fichier PDF en Markdown propre.
    Utilise PyMuPDF directement avec extraction par zone pour ignorer headers/footers.
    """
    import pymupdf

    doc = pymupdf.open(str(pdf_path))
    all_text = []

    for page_num, page in enumerate(doc):
        # Définir la zone de clip pour ignorer headers/footers
        rect = page.rect
        clip_rect = pymupdf.Rect(
            PDF_MARGIN_LEFT,
            PDF_MARGIN_TOP,
            rect.width - PDF_MARGIN_RIGHT,
            rect.height - PDF_MARGIN_BOTTOM
        )

        # Extraire le texte uniquement dans la zone de contenu
        text = page.get_text("text", clip=clip_rect)

        if text.strip():
            all_text.append(text)

    doc.close()

    # Joindre toutes les pages
    content = '\n\n'.join(all_text)

    # Post-traitement spécifique PDF
    content = post_process_pdf(content)

    # Convertir en format Markdown (titres, listes, etc.)
    content = convert_text_to_markdown(content)

    # Post-traitement commun
    content = post_process_markdown(content)

    return content


def convert_text_to_markdown(content: str) -> str:
    """
    Convertit le texte brut extrait du PDF en Markdown.
    Détecte les titres, listes, etc.
    """
    lines = content.split('\n')
    result = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            result.append('')
            continue

        # Détecter les titres numérotés (1. Titre, 1.1. Sous-titre, etc.)
        title_match = re.match(r'^(\d+\.(?:\d+\.)*)\s*(.+)$', stripped)
        if title_match:
            num = title_match.group(1)
            title_text = title_match.group(2)
            # Niveau basé sur le nombre de points
            level = num.count('.')
            if level == 1:
                result.append(f'## {num} {title_text}')
            elif level == 2:
                result.append(f'### {num} {title_text}')
            else:
                result.append(f'#### {num} {title_text}')
            continue

        # Détecter les titres en majuscules seules (CONTEXTE, PÉRIMÈTRE, etc.)
        if stripped.isupper() and len(stripped) > 3 and len(stripped) < 50 and not re.search(r'\d', stripped):
            result.append(f'## {stripped.title()}')
            continue

        # Détecter les listes à puces
        if stripped.startswith(('•', '-', '–', '▪', '●', '○')):
            bullet_text = re.sub(r'^[•\-–▪●○]\s*', '', stripped)
            result.append(f'- {bullet_text}')
            continue

        # Texte normal
        result.append(stripped)

    return '\n'.join(result)


def post_process_pdf(content: str) -> str:
    """
    Post-traitement spécifique pour les PDF.
    Supprime les éléments typiques des PDF qui ne sont pas dans les DOCX.
    """
    lines = content.split('\n')
    cleaned_lines = []

    # Patterns à supprimer spécifiques aux PDF
    skip_patterns = [
        # Numéros de page (ex: "1 / 17", "2/17", "Page 1", etc.)
        r'^\s*\d+\s*/\s*\d+\s*$',
        r'^\s*Page\s+\d+\s*$',
        r'^\s*-\s*\d+\s*-\s*$',
    ]

    # Détecter l'en-tête répété le plus fréquent
    header_candidates = {}
    for line in lines:
        stripped = line.strip()
        # Lignes courtes en majuscules = potentiel header
        if stripped and len(stripped) < 60 and stripped.isupper():
            header_candidates[stripped] = header_candidates.get(stripped, 0) + 1

    # Headers répétés = apparaissent plus de 3 fois
    repeated_headers = {h for h, count in header_candidates.items() if count >= 3}

    for line in lines:
        stripped = line.strip()

        # Ignorer les headers répétés
        if stripped in repeated_headers:
            continue

        # Ignorer les patterns de numéros de page
        skip = False
        for pattern in skip_patterns:
            if re.match(pattern, stripped, re.IGNORECASE):
                skip = True
                break

        if skip:
            continue

        cleaned_lines.append(line)

    return '\n'.join(cleaned_lines)


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
    Un vrai chapitre est un titre suivi de contenu textuel (pas juste d'autres titres).
    """

    # Fonction pour vérifier si un titre est suivi de vrai contenu
    def has_content_after(start_idx: int) -> bool:
        """
        Vérifie si après le titre il y a du contenu textuel (pas juste des titres).
        Dans une TOC: ## 1. Titre suivi de ## 2. Autre (même niveau)
        Dans le contenu: ## 1. Titre suivi de ### 1.1 Sous-titre (niveau inférieur) ou texte
        """
        current_line = lines[start_idx].strip()
        current_level = current_line.count('#') if current_line.startswith('#') else 0

        content_lines = 0

        for j in range(start_idx + 1, min(start_idx + 20, len(lines))):
            line = lines[j].strip()
            if not line:
                continue

            # Si c'est un titre markdown
            if line.startswith('#'):
                line_level = line.count('#') if line.startswith('#') else 0
                # Même niveau ou supérieur = probablement TOC
                if line_level > 0 and line_level <= current_level:
                    # C'est une entrée TOC au même niveau
                    if content_lines == 0:
                        return False
                # Niveau inférieur (### après ##) = sous-section, c'est du contenu
                continue

            # Si c'est une numérotation seule (ex: "1. Contexte") sans #
            if re.match(r'^\d+\.\d*\s+[A-Z]', line) and len(line) < 50:
                # Pourrait être TOC ou sous-titre, continuer
                continue

            # Ligne de texte substantiel
            if len(line) > 40:
                content_lines += 1
                if content_lines >= 1:
                    return True

        return content_lines >= 1

    # Méthode 1: Chercher après un marqueur de TOC explicite
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
        # Chercher le premier vrai titre après la TOC qui a du contenu
        for i in range(toc_end_index + 1, len(lines)):
            line = lines[i].strip()

            if not line:
                continue

            # Vérifier si c'est un titre avec du vrai contenu après
            if is_chapter_heading(line) and has_content_after(i):
                return i

    # Méthode 2: Détecter la fin de la TOC en trouvant le premier titre avec contenu
    # La TOC est une série de titres sans contenu entre eux
    consecutive_titles = 0
    last_title_idx = -1

    for i, line in enumerate(lines):
        line_stripped = line.strip()

        if not line_stripped:
            continue

        is_title = line_stripped.startswith('#') or re.match(r'^\d+\.', line_stripped)

        if is_title:
            consecutive_titles += 1
            last_title_idx = i
        else:
            # Si on a eu plusieurs titres consécutifs puis du contenu, c'était la TOC
            if consecutive_titles >= 5:
                # Chercher le premier titre avec contenu à partir d'ici
                for j in range(last_title_idx, len(lines)):
                    if lines[j].strip().startswith('#') and has_content_after(j):
                        return j
            consecutive_titles = 0

    # Méthode 3: Fallback - chercher simplement le premier titre avec contenu
    for i, line in enumerate(lines):
        line_stripped = line.strip()
        if is_chapter_heading(line_stripped) and has_content_after(i):
            return i

    return 0


def is_chapter_heading(line: str) -> bool:
    """
    Vérifie si une ligne est un titre de chapitre principal.
    Un titre Markdown H1/H2/H3 qui est soit :
    - Un titre connu (Description du projet, Périmètre, etc.)
    - Un titre avec numérotation (romaine ou décimale) substantielle
    """
    # Doit commencer par #, ## ou ### (H1, H2 ou H3)
    match = re.match(r'^(#{1,3})\s+(.+)$', line)
    if not match:
        return False

    title_text = match.group(2).strip()

    # Ne doit pas finir par un numéro seul (entrée de TOC avec numéro de page)
    if re.search(r'\s+\d+\s*$', title_text):
        return False

    # Supprimer le bold ** si présent
    title_text = re.sub(r'^\*\*(.+)\*\*$', r'\1', title_text)

    # Doit avoir du contenu substantiel
    if not title_text or len(title_text) < 5:
        return False

    # Titres de chapitres connus (premier vrai chapitre après la TOC)
    known_chapter_starts = [
        r'^Description\s+du\s+projet',
        r'^Introduction',
        r'^Contexte\s+(?:du|et)',
        r'^Pr[ée]sentation',
        r'^Objectifs?\s+(?:du|et)',
    ]

    for pattern in known_chapter_starts:
        if re.search(pattern, title_text, re.IGNORECASE):
            return True

    # Vérifier numérotation romaine ou décimale
    # Ex: "II. I. Description du projet", "1. Contexte", "IV.1. Sous-titre"
    has_numbering = re.match(
        r'^[IVXLCDM]+\.?\s+[IVXLCDM]*\.?\s*\d*\.?\s*[A-ZÀ-Ý]',  # Numérotation romaine (II., III., etc.)
        title_text
    ) or re.match(
        r'^\d+\.?\s+[A-ZÀ-Ý]',  # Numérotation décimale (1., 2., etc.)
        title_text
    )

    # Ignorer les titres de section préliminaire (I.1, I.2, I.3, I.4 = avant le vrai contenu)
    is_preliminary = re.match(r'^I\.\d+\.?\s+', title_text)

    return has_numbering and not is_preliminary


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
        """
        Nettoie la liste de chemins de fichiers.
        Gère les cellules avec plusieurs fichiers séparés par des pipes (|) et/ou espaces.
        """
        result = []
        for cell in lst:
            cell = cell.strip()
            if not cell:
                continue

            # Séparer par pipe d'abord
            parts = cell.split('|')

            for part in parts:
                part = part.strip().strip('"').strip("'")
                if part and part.lower() not in ('nan', 'none', ''):
                    result.append(part)

        return result

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
            for try_ext in [".docx", ".doc", ".DOCX", ".DOC", ".pdf", ".PDF"]:
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

        working_file = src
        is_pdf = ext == ".pdf"

        if ext == ".doc":
            converted = convert_doc_to_docx(src, tmp_conv)
            if not converted:
                stats["error"] += 1
                report_rows.append((mode, name, str(src), "", "CONVERT_ERROR", "Conversion .doc échouée"))
                logger.error(f"Conversion échouée: {src.name}")
                return
            working_file = converted

        try:
            # Utiliser la bonne fonction selon le type de fichier
            if is_pdf:
                md_content = pdf_to_markdown(working_file)
            else:
                md_content = docx_to_markdown(working_file)

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
