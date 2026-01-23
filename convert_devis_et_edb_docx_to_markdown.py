
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Conversion DOC/DOCX -> Markdown (dataset)
- NDC (colonne G) + EDB (colonne F)
- Même filtre Excel : B=1, C=1, D=OUI, E=NON
- Ignore images
- Ignore page de garde / synthèse / tables des matières (Table des matières ou SOMMAIRE)
- Logiques NDC/EDB SIMILAIRES :
  * On ignore l'avant-propos tant qu'on n'a pas trouvé (a) une Table des matières, ou (b) un début de corps probable
  * Si on trouve une TdM : on saute les lignes de TdM puis on démarre le BODY
  * Si on ne trouve jamais de TdM : on démarre le BODY dès qu'on voit un début de corps probable
- Conserve titres, paragraphes, sauts de lignes, listes multi-niveaux
- Conserve tableaux en HTML (supporte rowspan/colspan)
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
import sys
import html
import shutil
import subprocess
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple, Union
from collections import defaultdict

import pandas as pd
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.ns import qn
from lxml import etree

# ------------------------------
# Configuration
# ------------------------------
EXCEL_NAME = "couverture_EDB_NDC_par_RITM.xlsx"

# Colonnes Excel par lettre -> index (A=0)
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

# Titres fréquents (ndc)
KNOWN_H1 = {
    "description du projet",
    "périmètre du projet",
    "securité et réglementation",
    "sécurité et réglementation",
    "contraintes et risques",
    "description technique de la solution",
    "démarche pour la mise en œuvre",
    "demarche pour la mise en oeuvre",
    "offre de service en fonctionnement",
    "evaluation financière",
    "évaluation financière",
    "gestion de la documentation du projet",
    "rse - impact co2",
    "rse – impact co2",
}
KNOWN_H2 = {
    "contexte",
    "le besoin exprimé",
    "objectifs du projet",
    "périmètre",
    "hors périmètre",
    "hors perimetre",
    "projet partenaire",
    "projet interne",
    "contrôle psee",
    "controle psee",
    "contraintes et prérequis",
    "contraintes et prerequis",
    "risques projet",
    "description de la solution",
    "architecture",
    "composants et dimensionnement",
    "lotissement",
    "livrables du projet",
    "jalons clés du projet",
    "jalons cles du projet",
    "macro-planning",
    "macro planning",
    "détails des contributions",
    "details des contributions",
    "comitologie",
    "validation de la solution",
    "niveaux de service",
    "coûts du projet",
    "couts du projet",
    "facturation",
    "coûts de fonctionnement",
    "couts de fonctionnement",
}

# Roman numeral headings (NDC)
ROMAN_RE = re.compile(r"^\s*(?P<roman>M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3}))\.\s*(?P<title>.+?)\s*$", re.I)
ROMAN_SUB_RE = re.compile(r"^\s*(?P<roman>M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3}))\.(?P<sub>\d+)\.\s*(?P<title>.+?)\s*$", re.I)

# Digit headings (EDB)
DIGIT_H1_RE = re.compile(r"^\s*(?P<n>\d+)\s+(?P<title>.+?)\s*$")
DIGIT_H1_DOT_RE = re.compile(r"^\s*(?P<n>\d+)\.\s+(?P<title>.+?)\s*$")
DIGIT_SUB_RE = re.compile(r"^\s*(?P<num>\d+(?:\.\d+){1,})\s+(?P<title>.+?)\s*$")
DIGIT_SUB_DOT_RE = re.compile(r"^\s*(?P<num>\d+(?:\.\d+){1,})\.\s+(?P<title>.+?)\s*$")

TOC_HEADINGS = {"table des matières", "table des matieres", "sommaire"}
SUMMARY_HEADINGS = {"synthèse", "synthese"}

# TOC line detection:
# - ndc: I. ... 4
# - edb: 1 Introduction 4 / 2.4.1 Socles ... 5
TOC_LINE_RE = re.compile(
    r"^\s*(?:"
    r"(?:[IVXLCDM]+\.\d*\.?\s+.+?\s+\d+)"
    r"|"
    r"(?:\d+(?:\.\d+)*\.?\s+.+?\s+\d+)"
    r")\s*$",
    re.I
)

# ------------------------------
# XPath helper
# ------------------------------
def xp(el, expr: str, ns: Optional[dict] = None):
    """
    Execute XPath on python-docx oxml elements in a compatible way.
    BaseOxmlElement.xpath() may not accept namespaces=...
    """
    if ns:
        return etree.XPath(expr, namespaces=ns)(el)
    return el.xpath(expr)

# ------------------------------
# DOCX block iterator (p/tbl in order)
# ------------------------------
def iter_block_items(doc: Document) -> Iterable[Union[Paragraph, Table]]:
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)

# ------------------------------
# Optional DOC -> DOCX conversion
# ------------------------------
def convert_doc_to_docx(input_doc: Path, workdir: Path) -> Optional[Path]:
    """
    Try convert .doc -> .docx
    Priority:
      1) MS Word via COM (pywin32)
      2) LibreOffice headless (soffice)
    """
    workdir.mkdir(parents=True, exist_ok=True)
    out_docx = workdir / (input_doc.stem + ".docx")

    if os.name == "nt":
        try:
            import win32com.client  # type: ignore
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(input_doc))
            doc.SaveAs(str(out_docx), FileFormat=16)  # wdFormatXMLDocument
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
# Numbering (bullets/ordered) detection
# ------------------------------
def _get_numPr(paragraph: Paragraph):
    p = paragraph._p
    pPr = p.pPr
    if pPr is None:
        return None
    return pPr.numPr

def _read_numbering_formats(doc: Document) -> Dict[Tuple[int, int], str]:
    fmts: Dict[Tuple[int, int], str] = {}
    try:
        numbering = doc.part.numbering_part.element
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        abstract_map: Dict[str, Dict[str, str]] = {}
        for absnum in xp(numbering, ".//w:abstractNum", ns):
            abs_id = absnum.get(qn("w:abstractNumId"))
            lvl_map: Dict[str, str] = {}
            for lvl in xp(absnum, "./w:lvl", ns):
                ilvl = lvl.get(qn("w:ilvl"))
                numFmt = xp(lvl, "./w:numFmt/@w:val", ns)
                if numFmt:
                    lvl_map[ilvl] = numFmt[0]
            if abs_id:
                abstract_map[abs_id] = lvl_map

        num_to_abs: Dict[str, str] = {}
        for num in xp(numbering, ".//w:num", ns):
            num_id = num.get(qn("w:numId"))
            abs_id = xp(num, "./w:abstractNumId/@w:val", ns)
            if num_id and abs_id:
                num_to_abs[num_id] = abs_id[0]

        for num_id, abs_id in num_to_abs.items():
            lvl_map = abstract_map.get(abs_id, {})
            for ilvl, fmt in lvl_map.items():
                fmts[(int(num_id), int(ilvl))] = fmt

    except Exception:
        pass

    return fmts

def paragraph_list_info(paragraph: Paragraph, num_fmts: Dict[Tuple[int, int], str]) -> Optional[Tuple[str, int]]:
    numPr = _get_numPr(paragraph)
    if numPr is None:
        style = (paragraph.style.name or "").lower() if paragraph.style else ""
        if "list" in style or "puce" in style or "bullet" in style or "num" in style:
            level = 0
            kind = "bullet" if ("puce" in style or "bullet" in style) else "ordered"
            return (kind, level)
        return None

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    numId = xp(numPr, "./w:numId/@w:val", ns)
    ilvl = xp(numPr, "./w:ilvl/@w:val", ns)
    if not numId or not ilvl:
        return None

    num_id = int(numId[0])
    lvl = int(ilvl[0])
    fmt = num_fmts.get((num_id, lvl), "").lower()
    kind = "bullet" if fmt == "bullet" else "ordered"
    return kind, lvl

# ------------------------------
# Text extraction (ignore images, keep line breaks)
# ------------------------------
def run_has_drawing(run_element) -> bool:
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    }
    return bool(xp(run_element, ".//w:drawing | .//w:pict | .//pic:pic", ns))

def paragraph_to_text(paragraph: Paragraph) -> str:
    p = paragraph._p
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    chunks: List[str] = []

    for r in xp(p, "./w:r", ns):
        if run_has_drawing(r):
            continue

        brs = xp(r, "./w:br", ns)
        texts = xp(r, "./w:t", ns)
        bold = bool(xp(r, "./w:rPr/w:b", ns))
        italic = bool(xp(r, "./w:rPr/w:i", ns))
        t = "".join([(tt.text or "") for tt in texts])

        if t:
            t = t.replace("\u00A0", " ")
            if bold and italic:
                t = f"***{t}***"
            elif bold:
                t = f"**{t}**"
            elif italic:
                t = f"*{t}*"
            chunks.append(t)

        if brs:
            chunks.extend(["<br/>"] * len(brs))

    if not chunks and paragraph.text:
        return paragraph.text.replace("\u00A0", " ").strip()

    out = "".join(chunks)
    out = re.sub(r"(?:<br/>){3,}", "<br/><br/>", out)
    return out.strip()

# ------------------------------
# Heading + TOC skipping
# ------------------------------
def normalize_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = s.strip(" :\t")
    return s

def is_toc_heading(text: str) -> bool:
    nk = normalize_key(text)
    if nk in TOC_HEADINGS:
        return True
    if nk.startswith("sommaire"):
        return True
    if nk.startswith("table des mati"):
        return True
    return False

def is_summary_heading(text: str) -> bool:
    nk = normalize_key(text)
    return nk in SUMMARY_HEADINGS or nk.startswith("synth")

def looks_like_toc_line(text: str) -> bool:
    if not text:
        return True
    t = text.strip()
    if "\t" in t:
        return True
    if TOC_LINE_RE.match(t):
        return True
    # dotted leaders “.... 12”
    if re.search(r"\.{3,}\s*\d+\s*$", t):
        return True
    # ending page number (common in toc)
    if re.search(r"\s+\d+\s*$", t) and len(t) < 140:
        return True
    return False

def heading_level(paragraph: Paragraph, text: str) -> Optional[int]:
    """
    Heading detection:
      1) Word heading styles
      2) Roman numerals (ndc)
      3) Digit numbering (edb)
      4) Known H1/H2 titles
    """
    t = (text or "").strip()
    if not t:
        return None

    style_name = (paragraph.style.name or "") if paragraph.style else ""
    style_low = style_name.lower()

    # Word heading styles: "Heading 1", "Titre 1", "Heading 2", ...
    m = re.match(r"^(heading|titre)\s+(\d+)\b", style_low)
    if m:
        lvl = int(m.group(2))
        return max(1, min(6, lvl))

    # Roman headings (ndc)
    if ROMAN_SUB_RE.match(t):
        return 2
    if ROMAN_RE.match(t):
        return 1

    # Digit headings (edb)
    if DIGIT_SUB_DOT_RE.match(t) or DIGIT_SUB_RE.match(t):
        m2 = DIGIT_SUB_DOT_RE.match(t) or DIGIT_SUB_RE.match(t)
        if m2:
            depth = m2.group("num").count(".")
            return min(6, 2 + depth)  # 1.1 -> H3? ici H2+depth => 1.1(H3), 1.1.1(H4)
    if DIGIT_H1_DOT_RE.match(t) or DIGIT_H1_RE.match(t):
        return 1

    nk = normalize_key(t)
    if nk in KNOWN_H1:
        return 1
    if nk in KNOWN_H2:
        return 2

    return None

def is_probable_body_start(paragraph: Paragraph, text: str) -> bool:
    lvl = heading_level(paragraph, text)
    if lvl is not None:
        return True
    if normalize_key(text) in KNOWN_H1:
        return True
    return False

# ------------------------------
# Tables to HTML (supports colspan/rowspan)
# ------------------------------
@dataclass
class CellMeta:
    row: int
    col: int
    text: str
    colspan: int = 1
    rowspan: int = 1

def _tc_props(tc) -> Tuple[int, Optional[str]]:
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    gridSpan = xp(tc, "./w:tcPr/w:gridSpan/@w:val", ns)
    gridSpan = int(gridSpan[0]) if gridSpan else 1
    vMerge = xp(tc, "./w:tcPr/w:vMerge/@w:val", ns)
    if vMerge:
        return gridSpan, vMerge[0]
    vMerge_no_val = xp(tc, "./w:tcPr/w:vMerge", ns)
    if vMerge_no_val:
        return gridSpan, "continue"
    return gridSpan, None

def _extract_tc_text(tc) -> str:
    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    }
    parts: List[str] = []
    ps = xp(tc, ".//w:p", ns)
    for pi, p in enumerate(ps):
        chunk: List[str] = []
        for r in xp(p, "./w:r", ns):
            if xp(r, ".//w:drawing | .//w:pict | .//pic:pic", ns):
                continue
            brs = xp(r, "./w:br", ns)
            texts = xp(r, "./w:t", ns)
            t = "".join([(tt.text or "") for tt in texts]).replace("\u00A0", " ")
            if t:
                chunk.append(t)
            if brs:
                chunk.extend(["<br/>"] * len(brs))
        txt = "".join(chunk).strip()
        if txt:
            parts.append(txt)
        if pi < len(ps) - 1:
            parts.append("<br/>")
    out = "".join(parts)
    out = re.sub(r"(?:<br/>){3,}", "<br/><br/>", out)
    return out.strip()

def table_contains_toc_heading(table: Table) -> bool:
    tbl = table._tbl
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for tc in xp(tbl, ".//w:tc", ns):
        txt = _extract_tc_text(tc)
        if txt and is_toc_heading(txt):
            return True
    return False

def table_to_html(table: Table) -> str:
    tbl = table._tbl
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    trs = xp(tbl, "./w:tr", ns)

    rows_tc = []
    max_cols = 0
    for tr in trs:
        tcs = xp(tr, "./w:tc", ns)
        row = []
        col_sum = 0
        for tc in tcs:
            colspan, vmerge = _tc_props(tc)
            row.append((tc, colspan, vmerge))
            col_sum += colspan
        rows_tc.append(row)
        max_cols = max(max_cols, col_sum)

    grid: List[List[Optional[CellMeta]]] = [[None] * max_cols for _ in range(len(rows_tc))]

    for r, row in enumerate(rows_tc):
        c = 0
        for tc, colspan, vmerge in row:
            while c < max_cols and grid[r][c] is not None:
                c += 1
            if c >= max_cols:
                break

            txt = _extract_tc_text(tc)

            if vmerge == "continue":
                if r == 0:
                    owner = CellMeta(r, c, txt, colspan=colspan, rowspan=1)
                    for k in range(colspan):
                        if c + k < max_cols:
                            grid[r][c + k] = owner
                else:
                    owner = grid[r - 1][c]
                    if owner is None:
                        owner = CellMeta(r, c, txt, colspan=colspan, rowspan=1)
                    else:
                        owner.rowspan += 1
                    for k in range(colspan):
                        if c + k < max_cols:
                            grid[r][c + k] = owner
                c += colspan
                continue

            owner = CellMeta(r, c, txt, colspan=colspan, rowspan=1)
            for k in range(colspan):
                if c + k < max_cols:
                    grid[r][c + k] = owner
            c += colspan

    lines = ["<table>"]
    for r in range(len(rows_tc)):
        lines.append("  <tr>")
        col = 0
        while col < max_cols:
            owner = grid[r][col]
            if owner is None:
                col += 1
                continue
            if owner.row == r and owner.col == col:
                attrs = []
                if owner.colspan > 1:
                    attrs.append(f'colspan="{owner.colspan}"')
                if owner.rowspan > 1:
                    attrs.append(f'rowspan="{owner.rowspan}"')
                safe = html.escape(owner.text or "", quote=False).replace("<br/>", "<br/>")
                attr_str = (" " + " ".join(attrs)) if attrs else ""
                lines.append(f"    <td{attr_str}>{safe}</td>")
                col += owner.colspan
            else:
                col += 1
        lines.append("  </tr>")
    lines.append("</table>")
    return "\n".join(lines)

# ------------------------------
# Markdown assembly
# ------------------------------
def _md_heading(text: str, level: int) -> str:
    level = max(1, min(6, level))
    t = text.strip()

    m = ROMAN_SUB_RE.match(t)
    if m:
        t = f"{m.group('roman').upper()}.{m.group('sub')}. {m.group('title')}"
    else:
        m = ROMAN_RE.match(t)
        if m:
            t = f"{m.group('roman').upper()}. {m.group('title')}"

    return f"{'#' * level} {t}"

def dedupe_consecutive(lines: List[str]) -> List[str]:
    out: List[str] = []
    prev_key = None
    for line in lines:
        key = re.sub(r"\s+", " ", line.strip())
        if key and prev_key == key:
            continue
        out.append(line)
        prev_key = key if key else prev_key
    return out

# ------------------------------
# Core conversion with unified logic
# ------------------------------
def docx_to_markdown(docx_path: Path, mode: str) -> str:
    """
    mode: 'ndc' or 'edb'
    Logique unifiée:
      - SKIP_PRE: ignorer préambule jusqu'à TOC ou début de corps
      - SKIP_TOC: ignorer TdM (heading + lignes) jusqu'au premier contenu non-TdM
      - BODY: extraction markdown
    """
    doc = Document(str(docx_path))
    num_fmts = _read_numbering_formats(doc)

    md_lines: List[str] = []
    state = "SKIP_PRE"
    in_list = False

    def close_list_if_needed():
        nonlocal in_list
        if in_list:
            md_lines.append("")
            in_list = False

    for block in iter_block_items(doc):
        # --------- TABLE blocks (pré-BODY: TdM parfois dans des tables)
        if isinstance(block, Table):
            if state != "BODY":
                # Détection TdM dans table (même logique NDC/EDB)
                if table_contains_toc_heading(block):
                    state = "SKIP_TOC"
                continue

            # BODY: convertir la table
            close_list_if_needed()
            md_lines.append(table_to_html(block))
            md_lines.append("")
            continue

        # --------- PARAGRAPH blocks
        text = paragraph_to_text(block)

        if not text.strip():
            if state == "BODY":
                md_lines.append("")
            continue

        # -- SKIP_PRE : ignorer page de garde / synthèse / avant-propos
        if state == "SKIP_PRE":
            if is_summary_heading(text):
                continue

            if is_toc_heading(text):
                state = "SKIP_TOC"
                continue

            # Fallback: début de corps probable => BODY
            if is_probable_body_start(block, text):
                state = "BODY"
                # ne pas continue: on veut traiter ce paragraphe

            else:
                continue

        # -- SKIP_TOC : ignorer heading TdM + lignes TdM (comme NDC)
        if state == "SKIP_TOC":
            if is_toc_heading(text):
                continue
            if looks_like_toc_line(text):
                continue
            # première ligne non-TdM => BODY
            state = "BODY"
            # ne pas continue: on traite ce paragraphe comme du BODY

        # -- BODY processing
        lvl = heading_level(block, text)
        if lvl is not None:
            close_list_if_needed()
            md_lines.append(_md_heading(text, lvl))
            md_lines.append("")
            continue

        li = paragraph_list_info(block, num_fmts)
        if li:
            kind, level = li
            indent = " " * level
            bullet = "-" if kind == "bullet" else "1."
            if not in_list:
                in_list = True
            md_lines.append(f"{indent}{bullet} {text.strip()}")
            continue

        if in_list:
            close_list_if_needed()

        md_lines.append(text.strip())
        md_lines.append("")

    md_lines = dedupe_consecutive(md_lines)
    out = "\n".join(md_lines)
    out = re.sub(r"\n{4,}", "\n\n\n", out).strip() + "\n"
    return out

# ------------------------------
# Excel driver
# ------------------------------
def load_targets_from_excel(excel_path: Path) -> Tuple[List[str], List[str]]:
    df = pd.read_excel(excel_path, engine="openpyxl")

    b = df.iloc[:, COL_B]
    c = df.iloc[:, COL_C]
    d = df.iloc[:, COL_D]
    e = df.iloc[:, COL_E]
    f = df.iloc[:, COL_F]  # EDB
    g = df.iloc[:, COL_G]  # NDC

    d_norm = d.astype(str).str.strip().str.upper()
    e_norm = e.astype(str).str.strip().str.upper()

    mask = (b == FILTER_B_EQ) & (c == FILTER_C_EQ) & (d_norm == FILTER_D_EQ) & (e_norm == FILTER_E_EQ)

    edb = f[mask].dropna().astype(str).tolist()
    ndc = g[mask].dropna().astype(str).tolist()

    def clean(lst):
        out = []
        for t in lst:
            t = t.strip().strip('"').strip("'")
            if t:
                out.append(t)
        return out

    return clean(ndc), clean(edb)

def main() -> int:
    cwd = Path(".").resolve()
    excel_path = cwd / EXCEL_NAME

    base_out = cwd / OUTPUT_DIRNAME
    log_dir = base_out / LOG_DIRNAME
    out_ndc = base_out / SUBDIR_NDC
    out_edb = base_out / SUBDIR_EDB
    tmp_conv = base_out / "_tmp_doc_conversion"

    base_out.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)
    out_ndc.mkdir(parents=True, exist_ok=True)
    out_edb.mkdir(parents=True, exist_ok=True)
    tmp_conv.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        print(f"[ERREUR] Fichier Excel introuvable: {excel_path}")
        return 2

    ndc_list, edb_list = load_targets_from_excel(excel_path)

    print(f"Traitement NDC (col G) : {len(ndc_list)} fichiers")
    print(f"Traitement EDB (col F) : {len(edb_list)} fichiers")

    if not ndc_list and not edb_list:
        print("[INFO] Aucun fichier à traiter après filtrage Excel.")
        return 0

    report_rows = []
    ok = 0
    ko = 0

    def process_one(name: str, mode: str):
        nonlocal ok, ko, report_rows

        src = (cwd / name)
        ext = src.suffix.lower()

        # If no suffix, assume .docx
        if not ext:
            src = src.with_suffix(".docx")
            ext = ".docx"

        out_dir = out_ndc if mode == "ndc" else out_edb

        # Resolve missing extension cases
        if not src.exists():
            if ext != ".docx":
                alt = src.with_suffix(".docx")
                if alt.exists():
                    src = alt
                    ext = ".docx"

        if not src.exists():
            ko += 1
            report_rows.append((mode, name, str(src), "", "MISSING", "Fichier introuvable"))
            print(f"[WARN] Introuvable: {src}")
            return

        # Convert .doc if needed
        working_docx = None
        if ext == ".doc":
            converted = convert_doc_to_docx(src, tmp_conv)
            if not converted:
                ko += 1
                report_rows.append((mode, name, str(src), "", "UNSUPPORTED_DOC",
                                    "Impossible de convertir .doc -> .docx (Word/LibreOffice indisponible ?)"))
                print(f"[WARN] .doc non convertible: {src.name}")
                return
            working_docx = converted
        else:
            working_docx = src

        try:
            md = docx_to_markdown(working_docx, mode=mode)
            out_name = Path(name).stem + ".md"
            out_path = out_dir / out_name
            out_path.write_text(md, encoding="utf-8")

            ok += 1
            report_rows.append((mode, name, str(src), str(out_path), "OK", ""))
            print(f"[OK] ({mode}) {src.name} -> {out_path.name}")

        except Exception as ex:
            ko += 1
            err = f"{type(ex).__name__}: {ex}"
            report_rows.append((mode, name, str(src), "", "ERROR", err))
            trace = traceback.format_exc()
            (log_dir / (Path(name).stem + f".{mode}.error.log")).write_text(trace, encoding="utf-8")
            print(f"[ERREUR] ({mode}) {src.name}: {err}")

    for f in ndc_list:
        process_one(f, mode="ndc")

    for f in edb_list:
        process_one(f, mode="edb")

    rep = pd.DataFrame(report_rows, columns=["type", "source_excel", "input_path", "output_md", "status", "error"])
    rep_path = base_out / "conversion_report.csv"
    rep.to_csv(rep_path, index=False, encoding="utf-8")

    print("\n--- Résumé ---")
    print(f"Traités OK : {ok}")
    print(f"En erreur : {ko}")
    print(f"Sorties NDC : {out_ndc}")
    print(f"Sorties EDB : {out_edb}")
    print(f"Rapport : {rep_path}")
    if ko:
        print(f"Logs : {log_dir}")

    return 0 if ko == 0 else 1

if __name__ == "__main__":
    raise SystemExit(main())
