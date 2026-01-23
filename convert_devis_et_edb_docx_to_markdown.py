#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversion DOC/DOCX -> Markdown (dataset)
- NDC (colonne G) + EDB (colonne F)
- Filtre Excel : B=1, C=1, D=OUI, E=NON
- Ignore images
- Ignore page de garde / synthèse / tables des matières
- Conserve titres, paragraphes, sauts de lignes, listes multi-niveaux
- Tableaux en Markdown (format homogène pour training)
- Déduplication conservative (doublons consécutifs)
- Rapports CSV + logs d'erreurs

Dépendances:
  pip install -U pandas openpyxl python-docx lxml
Optionnel (pour lire .doc):
  pip install pywin32 (si MS Word installé)
  ou avoir LibreOffice accessible via 'soffice' dans le PATH
"""

from __future__ import annotations

import re
import os
import html
import shutil
import subprocess
import traceback
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple, Union
from enum import Enum, auto

import pandas as pd
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn
from lxml import etree

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ------------------------------
# Configuration
# ------------------------------
EXCEL_NAME = "couverture_EDB_NDC_par_RITM.xlsx"

COL_B = 1
COL_C = 2
COL_D = 3
COL_E = 4
COL_F = 5  # EDB
COL_G = 6  # NDC

FILTER_B_EQ = 1
FILTER_C_EQ = 1
FILTER_D_EQ = "OUI"
FILTER_E_EQ = "NON"

OUTPUT_DIRNAME = "dataset_markdown"
LOG_DIRNAME = "_logs"
SUBDIR_NDC = "ndc"
SUBDIR_EDB = "edb"

# Namespace XML Word
WORD_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}

# Patterns pour identifier les sections à ignorer
TOC_KEYWORDS = ["table des matières", "table des matieres", "sommaire", "table of contents"]
SKIP_SECTION_KEYWORDS = ["synthèse", "synthese", "résumé", "resume", "abstract", "executive summary"]

# Titres connus (normalisés en minuscules sans ponctuation)
KNOWN_SECTION_TITLES = {
    # Niveau 1 - Sections principales NDC
    "description du projet": 1,
    "périmètre du projet": 1,
    "perimetre du projet": 1,
    "securité et réglementation": 1,
    "sécurité et réglementation": 1,
    "securite et reglementation": 1,
    "contraintes et risques": 1,
    "description technique de la solution": 1,
    "démarche pour la mise en œuvre": 1,
    "demarche pour la mise en oeuvre": 1,
    "offre de service en fonctionnement": 1,
    "evaluation financière": 1,
    "évaluation financière": 1,
    "evaluation financiere": 1,
    "gestion de la documentation du projet": 1,
    "rse impact co2": 1,
    "introduction": 1,
    "conclusion": 1,
    "annexes": 1,

    # Niveau 2 - Sous-sections
    "contexte": 2,
    "le besoin exprimé": 2,
    "le besoin exprime": 2,
    "besoin exprimé": 2,
    "besoin exprime": 2,
    "objectifs du projet": 2,
    "objectifs": 2,
    "périmètre": 2,
    "perimetre": 2,
    "hors périmètre": 2,
    "hors perimetre": 2,
    "projet partenaire": 2,
    "projet interne": 2,
    "contrôle psee": 2,
    "controle psee": 2,
    "contraintes et prérequis": 2,
    "contraintes et prerequis": 2,
    "contraintes": 2,
    "prérequis": 2,
    "prerequis": 2,
    "risques projet": 2,
    "risques": 2,
    "description de la solution": 2,
    "architecture": 2,
    "composants et dimensionnement": 2,
    "dimensionnement": 2,
    "lotissement": 2,
    "livrables du projet": 2,
    "livrables": 2,
    "jalons clés du projet": 2,
    "jalons cles du projet": 2,
    "jalons": 2,
    "macro planning": 2,
    "macro-planning": 2,
    "planning": 2,
    "détails des contributions": 2,
    "details des contributions": 2,
    "contributions": 2,
    "comitologie": 2,
    "validation de la solution": 2,
    "niveaux de service": 2,
    "coûts du projet": 2,
    "couts du projet": 2,
    "coûts": 2,
    "couts": 2,
    "facturation": 2,
    "coûts de fonctionnement": 2,
    "couts de fonctionnement": 2,
    "coûts de construction": 2,
    "couts de construction": 2,
}


# ------------------------------
# Classes utilitaires
# ------------------------------
class DocState(Enum):
    """État de la machine d'état pour le parsing du document."""
    PREAMBLE = auto()
    TOC = auto()
    SKIP_SECTION = auto()
    BODY = auto()


@dataclass
class NumberingInfo:
    """Information sur la numérotation d'un paragraphe."""
    num_id: int
    ilvl: int
    num_format: str


# ------------------------------
# Fonction XPath corrigée
# ------------------------------
def lxml_xpath(element, expr: str) -> list:
    """Execute XPath avec lxml et les namespaces Word."""
    try:
        return etree.XPath(expr, namespaces=WORD_NS)(element)
    except Exception:
        return []


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
# Extraction de la numérotation
# ------------------------------
class NumberingExtractor:
    """Extrait les informations de numérotation du document Word."""

    def __init__(self, doc: Document):
        self.doc = doc
        self.formats: Dict[Tuple[int, int], str] = {}
        self._extract_numbering()

    def _extract_numbering(self):
        """Parse numbering.xml."""
        try:
            numbering_part = self.doc.part.numbering_part
            if numbering_part is None:
                return

            numbering_xml = numbering_part.element
            abstract_formats: Dict[str, Dict[str, str]] = {}

            for abstract_num in lxml_xpath(numbering_xml, ".//w:abstractNum"):
                abs_id = abstract_num.get(qn("w:abstractNumId"))
                if not abs_id:
                    continue

                level_formats = {}
                for lvl in lxml_xpath(abstract_num, "./w:lvl"):
                    ilvl = lvl.get(qn("w:ilvl"))
                    num_fmt_elem = lxml_xpath(lvl, "./w:numFmt/@w:val")
                    if num_fmt_elem:
                        level_formats[ilvl] = num_fmt_elem[0]

                abstract_formats[abs_id] = level_formats

            num_to_abstract: Dict[str, str] = {}
            for num in lxml_xpath(numbering_xml, ".//w:num"):
                num_id = num.get(qn("w:numId"))
                abs_id_refs = lxml_xpath(num, "./w:abstractNumId/@w:val")
                if num_id and abs_id_refs:
                    num_to_abstract[num_id] = abs_id_refs[0]

            for num_id, abs_id in num_to_abstract.items():
                if abs_id in abstract_formats:
                    for ilvl, fmt in abstract_formats[abs_id].items():
                        self.formats[(int(num_id), int(ilvl))] = fmt

        except Exception:
            pass

    def get_list_info(self, paragraph: Paragraph) -> Optional[NumberingInfo]:
        """Retourne les infos de liste pour un paragraphe."""
        p_elem = paragraph._p
        pPr = p_elem.pPr
        if pPr is None:
            return None

        numPr = pPr.numPr
        if numPr is None:
            style_name = (paragraph.style.name or "").lower() if paragraph.style else ""
            if any(kw in style_name for kw in ["list", "puce", "bullet", "enum", "num"]):
                is_bullet = any(kw in style_name for kw in ["puce", "bullet"])
                return NumberingInfo(0, 0, "bullet" if is_bullet else "decimal")
            return None

        num_id_elem = lxml_xpath(numPr, "./w:numId/@w:val")
        ilvl_elem = lxml_xpath(numPr, "./w:ilvl/@w:val")

        if not num_id_elem:
            return None

        num_id = int(num_id_elem[0])
        ilvl = int(ilvl_elem[0]) if ilvl_elem else 0
        fmt = self.formats.get((num_id, ilvl), "decimal")

        return NumberingInfo(num_id, ilvl, fmt)


# ------------------------------
# Extraction du texte
# ------------------------------
class TextExtractor:
    """Extrait le texte des paragraphes Word."""

    @staticmethod
    def has_drawing(run_element) -> bool:
        """Vérifie si un run contient une image."""
        return bool(
            lxml_xpath(run_element, ".//w:drawing") or
            lxml_xpath(run_element, ".//w:pict") or
            lxml_xpath(run_element, ".//pic:pic")
        )

    @staticmethod
    def extract_paragraph_text(paragraph: Paragraph, include_formatting: bool = True) -> str:
        """Extrait le texte d'un paragraphe."""
        p_elem = paragraph._p
        chunks: List[str] = []

        for run in lxml_xpath(p_elem, ".//w:r"):
            if TextExtractor.has_drawing(run):
                continue

            text_elements = lxml_xpath(run, ".//w:t")
            text = "".join([(t.text or "") for t in text_elements])

            if not text:
                if lxml_xpath(run, ".//w:br"):
                    chunks.append("\n")
                continue

            text = text.replace("\u00A0", " ").replace("\u200B", "")

            if include_formatting:
                is_bold = bool(lxml_xpath(run, "./w:rPr/w:b[not(@w:val='false') and not(@w:val='0')]"))
                is_italic = bool(lxml_xpath(run, "./w:rPr/w:i[not(@w:val='false') and not(@w:val='0')]"))

                if is_bold and is_italic:
                    text = f"***{text}***"
                elif is_bold:
                    text = f"**{text}**"
                elif is_italic:
                    text = f"*{text}*"

            chunks.append(text)

        result = "".join(chunks).strip()
        result = re.sub(r'\*{4,}', '**', result)
        return result

    @staticmethod
    def extract_raw_text(paragraph: Paragraph) -> str:
        """Extrait le texte brut sans formatage."""
        return TextExtractor.extract_paragraph_text(paragraph, include_formatting=False)


# ------------------------------
# Détection des titres
# ------------------------------
class HeadingDetector:
    """Détecte et classifie les titres."""

    # Pattern numérotation romaine
    ROMAN_PATTERN = re.compile(
        r"^\s*([IVXLCDM]+)\.?\s*(\d+)?\.?\s*(.+)$",
        re.IGNORECASE
    )

    # Pattern numérotation décimale
    DECIMAL_PATTERN = re.compile(
        r"^\s*(\d+)(?:\.(\d+))?(?:\.(\d+))?(?:\.(\d+))?\.?\s+(.+)$"
    )

    @staticmethod
    def normalize_title(text: str) -> str:
        """Normalise un titre pour comparaison."""
        text = text.lower().strip()
        text = re.sub(r'[:\-–—.,;!?]+$', '', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()

    @staticmethod
    def get_style_heading_level(paragraph: Paragraph) -> Optional[int]:
        """Retourne le niveau de titre basé sur le style Word."""
        if not paragraph.style:
            return None

        style_name = paragraph.style.name or ""
        style_lower = style_name.lower()

        match = re.match(r"^(heading|titre|title)\s*(\d+)", style_lower)
        if match:
            level = int(match.group(2))
            return min(6, max(1, level))

        p_elem = paragraph._p
        pPr = p_elem.pPr
        if pPr is not None:
            outline_lvl = lxml_xpath(pPr, "./w:outlineLvl/@w:val")
            if outline_lvl:
                level = int(outline_lvl[0]) + 1
                return min(6, max(1, level))

        return None

    @staticmethod
    def get_pattern_heading_level(text: str) -> Optional[Tuple[int, str]]:
        """Détecte un titre par pattern de numérotation."""
        text = text.strip()
        if not text or len(text) > 200:
            return None

        # Numérotation romaine: I. Titre, II.1. Titre, etc.
        match = HeadingDetector.ROMAN_PATTERN.match(text)
        if match:
            roman = match.group(1).upper()
            sub_num = match.group(2)
            title = match.group(3).strip()

            # Ignorer si ça ressemble à une ligne de TOC (finit par un numéro)
            if re.search(r'\d+\s*$', title) and len(title) < 80:
                return None

            if sub_num:
                return (2, text)
            else:
                return (1, text)

        # Numérotation décimale: 1. Titre, 1.1 Titre, etc.
        match = HeadingDetector.DECIMAL_PATTERN.match(text)
        if match:
            groups = [g for g in match.groups()[:-1] if g]
            title = match.groups()[-1].strip()

            if re.search(r'\d+\s*$', title) and len(title) < 80:
                return None

            level = len(groups)
            return (min(6, level), text)

        return None

    @staticmethod
    def get_known_title_level(text: str) -> Optional[int]:
        """Retourne le niveau si c'est un titre connu."""
        normalized = HeadingDetector.normalize_title(text)
        return KNOWN_SECTION_TITLES.get(normalized)

    @staticmethod
    def detect_heading(paragraph: Paragraph, text: str) -> Optional[Tuple[int, str]]:
        """Détecte si un paragraphe est un titre."""
        if not text or not text.strip():
            return None

        text = text.strip()

        # 1. Style Word
        style_level = HeadingDetector.get_style_heading_level(paragraph)
        if style_level is not None:
            return (style_level, text)

        # 2. Pattern de numérotation
        pattern_result = HeadingDetector.get_pattern_heading_level(text)
        if pattern_result:
            return pattern_result

        # 3. Titre connu
        known_level = HeadingDetector.get_known_title_level(text)
        if known_level is not None:
            return (known_level, text)

        return None


# ------------------------------
# Filtrage des sections
# ------------------------------
class SectionFilter:
    """Filtre les sections à ignorer."""

    @staticmethod
    def is_toc_heading(text: str) -> bool:
        """Vérifie si le texte est un titre de TOC."""
        normalized = text.strip().lower()
        return any(kw in normalized for kw in TOC_KEYWORDS)

    @staticmethod
    def is_skip_section_heading(text: str) -> bool:
        """Vérifie si c'est une section à ignorer."""
        normalized = text.strip().lower()
        return any(kw in normalized for kw in SKIP_SECTION_KEYWORDS)

    @staticmethod
    def is_toc_line(text: str) -> bool:
        """Vérifie si une ligne ressemble à une entrée de TOC."""
        if not text or not text.strip():
            return False

        text = text.strip()

        # Tabulation
        if "\t" in text:
            return True

        # Points de suite
        if "..." in text or "…" in text:
            return True

        # Ligne qui finit par un numéro de page (caractéristique TOC)
        if re.search(r'\d+\s*$', text) and len(text) < 150:
            # Vérifie si c'est une numérotation suivie d'un titre et d'un numéro de page
            if re.match(r'^[IVXLCDM\d]+\.', text):
                return True

        return False

    @staticmethod
    def is_page_number_only(text: str) -> bool:
        """Vérifie si le texte est juste un numéro de page."""
        return bool(re.match(r'^\s*\d+\s*$', text.strip()))

    @staticmethod
    def looks_like_cover_page_element(text: str) -> bool:
        """Vérifie si ça ressemble à un élément de page de garde."""
        text_lower = text.strip().lower()

        cover_patterns = [
            r"^confidentiel",
            r"^document\s+interne",
            r"^usage\s+interne",
            r"^version\s*:?\s*[\d\.]+",
            r"^date\s*:?\s*\d",
            r"^auteur\s*:",
            r"^rédacteur\s*:",
            r"^redacteur\s*:",
            r"^statut\s*:",
            r"^référence\s*:",
            r"^reference\s*:",
            r"^diffusion\s*:",
            r"^\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4}$",
        ]

        for pattern in cover_patterns:
            if re.match(pattern, text_lower):
                return True

        return False


# ------------------------------
# Conversion des tableaux (Markdown uniquement)
# ------------------------------
class TableConverter:
    """Convertit les tableaux Word en Markdown."""

    @staticmethod
    def extract_cell_text(tc_element) -> str:
        """Extrait le texte d'une cellule."""
        paragraphs = lxml_xpath(tc_element, ".//w:p")
        parts = []

        for p in paragraphs:
            text_parts = []
            for run in lxml_xpath(p, ".//w:r"):
                if lxml_xpath(run, ".//w:drawing") or lxml_xpath(run, ".//w:pict"):
                    continue
                texts = lxml_xpath(run, ".//w:t")
                text_parts.append("".join([(t.text or "") for t in texts]))

            para_text = "".join(text_parts).strip()
            if para_text:
                parts.append(para_text)

        return " ".join(parts).replace("\u00A0", " ").replace("\n", " ").strip()

    @staticmethod
    def get_cell_colspan(tc_element) -> int:
        """Retourne le colspan d'une cellule."""
        grid_span = lxml_xpath(tc_element, "./w:tcPr/w:gridSpan/@w:val")
        return int(grid_span[0]) if grid_span else 1

    @staticmethod
    def is_vmerge_continue(tc_element) -> bool:
        """Vérifie si la cellule est une continuation de fusion verticale."""
        vmerge = lxml_xpath(tc_element, "./w:tcPr/w:vMerge")
        if vmerge:
            vmerge_val = lxml_xpath(tc_element, "./w:tcPr/w:vMerge/@w:val")
            if not vmerge_val or vmerge_val[0] != "restart":
                return True
        return False

    @staticmethod
    def table_to_markdown(table: Table) -> str:
        """Convertit un tableau en Markdown."""
        tbl_elem = table._tbl
        rows = lxml_xpath(tbl_elem, "./w:tr")

        if not rows:
            return ""

        md_rows = []
        max_cols = 0

        for row in rows:
            cells = lxml_xpath(row, "./w:tc")
            row_texts = []
            col_count = 0

            for cell in cells:
                colspan = TableConverter.get_cell_colspan(cell)
                is_continue = TableConverter.is_vmerge_continue(cell)

                if is_continue:
                    text = ""
                else:
                    text = TableConverter.extract_cell_text(cell)

                # Escape le pipe pour Markdown
                text = text.replace("|", "\\|")
                row_texts.append(text)

                # Ajouter des cellules vides pour colspan
                for _ in range(colspan - 1):
                    row_texts.append("")

                col_count += colspan

            max_cols = max(max_cols, col_count)
            md_rows.append(row_texts)

        # Normaliser les colonnes
        for row in md_rows:
            while len(row) < max_cols:
                row.append("")

        if not md_rows or max_cols == 0:
            return ""

        # Construire le Markdown
        lines = []

        # Header
        header = "| " + " | ".join(md_rows[0]) + " |"
        lines.append(header)

        # Séparateur
        sep = "| " + " | ".join(["---"] * max_cols) + " |"
        lines.append(sep)

        # Lignes de données
        for row in md_rows[1:]:
            line = "| " + " | ".join(row) + " |"
            lines.append(line)

        return "\n".join(lines)

    @staticmethod
    def contains_toc(table: Table) -> bool:
        """Vérifie si le tableau contient une TOC."""
        tbl_elem = table._tbl
        for tc in lxml_xpath(tbl_elem, ".//w:tc"):
            text = TableConverter.extract_cell_text(tc)
            if SectionFilter.is_toc_heading(text):
                return True
        return False


# ------------------------------
# Itération sur les blocs
# ------------------------------
def iter_document_blocks(doc: Document) -> Iterable[Union[Paragraph, Table]]:
    """Itère sur les paragraphes et tableaux."""
    body = doc.element.body

    for child in body.iterchildren():
        tag = child.tag

        if tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif tag == qn("w:tbl"):
            yield Table(child, doc)


# ------------------------------
# Convertisseur principal
# ------------------------------
class DocxToMarkdownConverter:
    """Convertisseur DOCX vers Markdown."""

    def __init__(self, docx_path: Path, mode: str = "ndc"):
        self.docx_path = docx_path
        self.mode = mode
        self.doc = Document(str(docx_path))
        self.numbering = NumberingExtractor(self.doc)

        self.state = DocState.PREAMBLE
        self.md_lines: List[str] = []
        self.in_list = False
        self.preamble_line_count = 0
        self.body_started = False

    def convert(self) -> str:
        """Effectue la conversion."""
        for block in iter_document_blocks(self.doc):
            self._process_block(block)

        return self._finalize()

    def _process_block(self, block: Union[Paragraph, Table]):
        """Traite un bloc."""
        if isinstance(block, Table):
            self._process_table(block)
        else:
            self._process_paragraph(block)

    def _process_table(self, table: Table):
        """Traite un tableau."""
        if TableConverter.contains_toc(table):
            self.state = DocState.TOC
            return

        if self.state != DocState.BODY:
            return

        self._close_list()

        table_md = TableConverter.table_to_markdown(table)
        if table_md:
            self.md_lines.append("")
            self.md_lines.append(table_md)
            self.md_lines.append("")

    def _process_paragraph(self, paragraph: Paragraph):
        """Traite un paragraphe."""
        text = TextExtractor.extract_paragraph_text(paragraph)
        raw_text = TextExtractor.extract_raw_text(paragraph)

        if not raw_text.strip():
            if self.state == DocState.BODY:
                self._close_list()
                self.md_lines.append("")
            return

        if self.state == DocState.PREAMBLE:
            self._handle_preamble(paragraph, text, raw_text)
        elif self.state == DocState.TOC:
            self._handle_toc(paragraph, text, raw_text)
        elif self.state == DocState.SKIP_SECTION:
            self._handle_skip_section(paragraph, text, raw_text)
        else:
            self._handle_body(paragraph, text, raw_text)

    def _handle_preamble(self, paragraph: Paragraph, text: str, raw_text: str):
        """Gère le préambule."""
        self.preamble_line_count += 1

        if SectionFilter.is_toc_heading(raw_text):
            self.state = DocState.TOC
            return

        if SectionFilter.is_skip_section_heading(raw_text):
            self.state = DocState.SKIP_SECTION
            return

        heading = HeadingDetector.detect_heading(paragraph, raw_text)
        if heading:
            level, title = heading
            if level <= 2 and not SectionFilter.looks_like_cover_page_element(title):
                if not SectionFilter.is_toc_line(raw_text):
                    self.state = DocState.BODY
                    self.body_started = True
                    self._add_heading(level, text)
                    return

        if self.preamble_line_count > 100:
            if len(raw_text) > 50 and not SectionFilter.looks_like_cover_page_element(raw_text):
                self.state = DocState.BODY
                self.body_started = True
                self._handle_body(paragraph, text, raw_text)

    def _handle_toc(self, paragraph: Paragraph, text: str, raw_text: str):
        """Gère la TOC."""
        if SectionFilter.is_toc_line(raw_text):
            return

        if SectionFilter.is_page_number_only(raw_text):
            return

        heading = HeadingDetector.detect_heading(paragraph, raw_text)
        if heading:
            level, title = heading
            if not SectionFilter.is_toc_heading(raw_text):
                if not SectionFilter.is_toc_line(raw_text):
                    self.state = DocState.BODY
                    self.body_started = True
                    self._add_heading(level, text)
                    return

        if len(raw_text) > 80 and not SectionFilter.is_toc_line(raw_text):
            self.state = DocState.BODY
            self.body_started = True
            self._handle_body(paragraph, text, raw_text)

    def _handle_skip_section(self, paragraph: Paragraph, text: str, raw_text: str):
        """Gère une section à ignorer."""
        heading = HeadingDetector.detect_heading(paragraph, raw_text)
        if heading:
            level, title = heading
            if level == 1 and not SectionFilter.is_skip_section_heading(raw_text):
                self.state = DocState.BODY
                self.body_started = True
                self._add_heading(level, text)
                return

        if SectionFilter.is_toc_heading(raw_text):
            self.state = DocState.TOC

    def _handle_body(self, paragraph: Paragraph, text: str, raw_text: str):
        """Gère le contenu principal."""
        if SectionFilter.looks_like_cover_page_element(raw_text) and self.preamble_line_count < 10:
            return

        self.body_started = True

        if SectionFilter.is_skip_section_heading(raw_text):
            self.state = DocState.SKIP_SECTION
            return

        heading = HeadingDetector.detect_heading(paragraph, raw_text)
        if heading:
            self._close_list()
            level, title = heading
            self._add_heading(level, text)
            return

        list_info = self.numbering.get_list_info(paragraph)
        if list_info:
            self._add_list_item(list_info, text)
            return

        self._close_list()
        self.md_lines.append(text)
        self.md_lines.append("")

    def _add_heading(self, level: int, text: str):
        """Ajoute un titre."""
        self.md_lines.append("")
        prefix = "#" * level
        self.md_lines.append(f"{prefix} {text}")
        self.md_lines.append("")

    def _add_list_item(self, list_info: NumberingInfo, text: str):
        """Ajoute un élément de liste."""
        if not self.in_list:
            self.md_lines.append("")
            self.in_list = True

        indent = "  " * list_info.ilvl
        marker = "-" if list_info.num_format == "bullet" else "1."
        self.md_lines.append(f"{indent}{marker} {text}")

    def _close_list(self):
        """Ferme une liste."""
        if self.in_list:
            self.md_lines.append("")
            self.in_list = False

    def _finalize(self) -> str:
        """Finalise le Markdown."""
        self._close_list()

        # Dédupliquer
        deduped = []
        prev_key = None
        for line in self.md_lines:
            key = re.sub(r'\s+', ' ', line.strip())
            if key and key == prev_key:
                continue
            deduped.append(line)
            prev_key = key if key else prev_key

        result = "\n".join(deduped)
        result = re.sub(r'\n{4,}', '\n\n\n', result)
        result = result.lstrip('\n')
        result = result.rstrip() + '\n'

        return result


# ------------------------------
# Chargement Excel
# ------------------------------
def load_targets_from_excel(excel_path: Path) -> Tuple[List[str], List[str]]:
    """Charge les fichiers à traiter."""
    df = pd.read_excel(excel_path, engine="openpyxl")

    b = df.iloc[:, COL_B]
    c = df.iloc[:, COL_C]
    d = df.iloc[:, COL_D]
    e = df.iloc[:, COL_E]
    f = df.iloc[:, COL_F]
    g = df.iloc[:, COL_G]

    d_norm = d.astype(str).str.strip().str.upper()
    e_norm = e.astype(str).str.strip().str.upper()

    mask = (b == FILTER_B_EQ) & (c == FILTER_C_EQ) & (d_norm == FILTER_D_EQ) & (e_norm == FILTER_E_EQ)

    edb = f[mask].dropna().astype(str).tolist()
    ndc = g[mask].dropna().astype(str).tolist()

    def clean(lst):
        return [t.strip().strip('"').strip("'") for t in lst if t.strip()]

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
            converter = DocxToMarkdownConverter(working_docx, mode=mode)
            md_content = converter.convert()

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
