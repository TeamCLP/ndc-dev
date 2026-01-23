#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Construit des datasets JSONL (train/val) pour fine-tuning Mistral Instruct.
Apparie les fichiers EDB (Expression de Besoin) et NDC (Note de Cadrage) par référence.

Dépendances:
  pip install (aucune externe, utilise stdlib)

Usage:
  python build_dataset_jsonl.py --edb_dir edb --ndc_dir ndc --report
"""

import json
import re
import argparse
import random
import logging
from pathlib import Path
from collections import defaultdict
from typing import Dict, Tuple, List, Optional

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ==============================================================================
# CONFIGURATION - MODIFIEZ ICI
# ==============================================================================

# ------------------------------
# DOSSIERS PAR DÉFAUT
# ------------------------------
DEFAULT_EDB_DIR = "edb"
DEFAULT_NDC_DIR = "ndc"
DEFAULT_TRAIN_OUT = "train_dataset.jsonl"
DEFAULT_VAL_OUT = "val_dataset.jsonl"

# ------------------------------
# SPLIT TRAIN/VAL
# ------------------------------
DEFAULT_TRAIN_RATIO = 0.9  # 90% train, 10% val
DEFAULT_SEED = 42  # Pour reproductibilité

# ------------------------------
# FILTRES CONTENU
# ------------------------------
# Ces filtres s'appliquent aux fichiers Markdown

# Longueur minimale du contenu (en caractères)
# Les paires avec EDB ou NDC plus court seront ignorées
MIN_CONTENT_CHARS = 100

# Longueur maximale du contenu (en caractères, 0 = pas de limite)
# Utile pour éviter les documents trop longs qui dépassent le contexte du modèle
MAX_CONTENT_CHARS = 0  # 0 = pas de limite

# Nombre minimum de lignes non vides
MIN_CONTENT_LINES = 5

# Patterns à exclure (si le contenu matche un de ces patterns, la paire est ignorée)
# Exemple: fichiers templates non remplis, erreurs de conversion, etc.
EXCLUDE_CONTENT_PATTERNS = [
    r'^#?\s*$',  # Fichiers vides ou avec seulement un titre vide
    r'^\s*ERREUR',  # Fichiers avec erreur de conversion
    r'^\s*\[TEMPLATE\]',  # Templates non remplis
]

# Patterns requis (le contenu DOIT matcher au moins un de ces patterns)
# Mettre une liste vide [] pour désactiver
REQUIRE_CONTENT_PATTERNS = []  # Ex: [r'Description du projet', r'Contexte']

# ------------------------------
# FILTRES NOM DE FICHIER
# ------------------------------
# Regex pour extraire la référence depuis le nom de fichier
# Format attendu: CAGIPRITM<digits> - ...
# Accepte séparateurs: -, – (en-dash), — (em-dash), _ ; avec ou sans espaces
REF_FROM_FILENAME_PATTERN = r"^(CAGIPRITM\d+)\s*[-–—_]\s*.*$"

# Patterns de noms de fichiers à exclure
EXCLUDE_FILENAME_PATTERNS = [
    r'^\..*',  # Fichiers cachés
    r'^_.*',   # Fichiers temporaires
    r'.*\.backup\.md$',
    r'.*\.old\.md$',
]

# ------------------------------
# FORMAT DATASET
# ------------------------------
# Format des messages pour le fine-tuning
# Options: "mistral_instruct", "chatml", "alpaca"
DATASET_FORMAT = "mistral_instruct"

# System prompt (optionnel, laisser vide "" pour ne pas en avoir)
SYSTEM_PROMPT = ""

# Exemple de system prompt:
# SYSTEM_PROMPT = "Tu es un assistant spécialisé dans la rédaction de Notes de Cadrage (NDC) à partir d'Expressions de Besoin (EDB). Tu dois produire des documents professionnels, structurés et complets."


# ==============================================================================
# FIN CONFIGURATION
# ==============================================================================

# Compile les patterns une fois
REF_FROM_FILENAME_RE = re.compile(REF_FROM_FILENAME_PATTERN, re.IGNORECASE)
EXCLUDE_FILENAME_RES = [re.compile(p, re.IGNORECASE) for p in EXCLUDE_FILENAME_PATTERNS]
EXCLUDE_CONTENT_RES = [re.compile(p, re.MULTILINE | re.IGNORECASE) for p in EXCLUDE_CONTENT_PATTERNS]
REQUIRE_CONTENT_RES = [re.compile(p, re.MULTILINE | re.IGNORECASE) for p in REQUIRE_CONTENT_PATTERNS]


def read_text(path: Path) -> str:
    """Lit un fichier en UTF-8 (fallback latin-1) et normalise les fins de ligne."""
    try:
        txt = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        txt = path.read_text(encoding="latin-1")
    return txt.replace("\r\n", "\n").replace("\r", "\n")


def extract_ref_from_filename(path: Path) -> Optional[str]:
    """Extrait la référence depuis le nom de fichier."""
    stem = path.stem.strip()
    m = REF_FROM_FILENAME_RE.match(stem)
    return m.group(1).upper() if m else None


def should_exclude_filename(path: Path) -> bool:
    """Vérifie si le fichier doit être exclu selon son nom."""
    name = path.name
    for pattern in EXCLUDE_FILENAME_RES:
        if pattern.match(name):
            return True
    return False


def validate_content(content: str) -> Tuple[bool, str]:
    """
    Valide le contenu selon les filtres configurés.
    Retourne (is_valid, reason).
    """
    # Vérifier longueur minimale
    if MIN_CONTENT_CHARS > 0 and len(content) < MIN_CONTENT_CHARS:
        return False, f"trop court ({len(content)} < {MIN_CONTENT_CHARS} chars)"

    # Vérifier longueur maximale
    if MAX_CONTENT_CHARS > 0 and len(content) > MAX_CONTENT_CHARS:
        return False, f"trop long ({len(content)} > {MAX_CONTENT_CHARS} chars)"

    # Vérifier nombre de lignes
    non_empty_lines = [l for l in content.split('\n') if l.strip()]
    if MIN_CONTENT_LINES > 0 and len(non_empty_lines) < MIN_CONTENT_LINES:
        return False, f"pas assez de lignes ({len(non_empty_lines)} < {MIN_CONTENT_LINES})"

    # Vérifier patterns d'exclusion
    for pattern in EXCLUDE_CONTENT_RES:
        if pattern.search(content):
            return False, f"matche pattern d'exclusion"

    # Vérifier patterns requis
    if REQUIRE_CONTENT_RES:
        found = any(pattern.search(content) for pattern in REQUIRE_CONTENT_RES)
        if not found:
            return False, "ne matche aucun pattern requis"

    return True, ""


def index_folder_by_ref(folder: Path) -> Tuple[Dict[str, Tuple[Path, str]], Dict[str, List[Path]], List[Path], List[Tuple[Path, str]]]:
    """
    Indexe tous les .md par référence.
    Retourne:
      - index: ref -> (path, content) (1 fichier retenu par ref)
      - duplicates: ref -> [paths...] si plusieurs fichiers ont la même ref
      - no_ref: liste des fichiers sans ref détectée
      - invalid: liste de (path, reason) pour fichiers filtrés
    """
    index: Dict[str, Tuple[Path, str]] = {}
    duplicates: Dict[str, List[Path]] = defaultdict(list)
    no_ref: List[Path] = []
    invalid: List[Tuple[Path, str]] = []

    for p in sorted(folder.rglob("*.md")):
        # Exclure selon le nom
        if should_exclude_filename(p):
            continue

        ref = extract_ref_from_filename(p)
        if not ref:
            no_ref.append(p)
            continue

        content = read_text(p).strip()

        # Valider le contenu
        is_valid, reason = validate_content(content)
        if not is_valid:
            invalid.append((p, reason))
            continue

        if ref in index:
            duplicates[ref].append(p)
        else:
            index[ref] = (p, content)

    # Pour lisibilité: inclure aussi le premier fichier dans la liste de doublons
    for ref, paths in list(duplicates.items()):
        paths.insert(0, index[ref][0])

    return index, duplicates, no_ref, invalid


def make_record(edb_text: str, ndc_text: str) -> dict:
    """
    Crée un enregistrement au format configuré.
    """
    if DATASET_FORMAT == "mistral_instruct":
        # Format Mistral Instruct (v0.2+)
        # https://huggingface.co/mistralai/Mistral-7B-Instruct-v0.2
        messages = []
        if SYSTEM_PROMPT:
            messages.append({"role": "system", "content": SYSTEM_PROMPT})
        messages.append({"role": "user", "content": edb_text})
        messages.append({"role": "assistant", "content": ndc_text})
        return {"messages": messages}

    elif DATASET_FORMAT == "chatml":
        # Format ChatML
        messages = []
        if SYSTEM_PROMPT:
            messages.append({"role": "system", "content": SYSTEM_PROMPT})
        messages.append({"role": "user", "content": edb_text})
        messages.append({"role": "assistant", "content": ndc_text})
        return {"messages": messages}

    elif DATASET_FORMAT == "alpaca":
        # Format Alpaca
        record = {
            "instruction": edb_text,
            "output": ndc_text
        }
        if SYSTEM_PROMPT:
            record["input"] = SYSTEM_PROMPT
        return record

    else:
        # Format par défaut (messages simples)
        return {
            "messages": [
                {"role": "user", "content": edb_text},
                {"role": "assistant", "content": ndc_text}
            ]
        }


def write_jsonl(path: Path, records: List[dict]) -> None:
    """Écrit les enregistrements dans un fichier JSONL."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")


def print_config_summary():
    """Affiche un résumé de la configuration active."""
    logger.info("=== Configuration ===")
    logger.info(f"Format dataset: {DATASET_FORMAT}")
    logger.info(f"System prompt: {'Oui' if SYSTEM_PROMPT else 'Non'}")
    logger.info(f"Min chars: {MIN_CONTENT_CHARS}")
    logger.info(f"Max chars: {MAX_CONTENT_CHARS if MAX_CONTENT_CHARS > 0 else 'illimité'}")
    logger.info(f"Min lignes: {MIN_CONTENT_LINES}")
    logger.info(f"Patterns exclusion: {len(EXCLUDE_CONTENT_PATTERNS)}")
    logger.info(f"Patterns requis: {len(REQUIRE_CONTENT_PATTERNS)}")


def main():
    parser = argparse.ArgumentParser(
        description="Construit 2 datasets JSONL (train/val) pour Mistral Instruct depuis edb/ et ndc/ appariés par ref dans le nom de fichier."
    )
    parser.add_argument("--edb_dir", default=DEFAULT_EDB_DIR,
                        help=f"Dossier EDB (entrée user). Défaut: {DEFAULT_EDB_DIR}")
    parser.add_argument("--ndc_dir", default=DEFAULT_NDC_DIR,
                        help=f"Dossier NDC (sortie assistant). Défaut: {DEFAULT_NDC_DIR}")

    parser.add_argument("--train_out", default=DEFAULT_TRAIN_OUT,
                        help=f"Fichier train JSONL. Défaut: {DEFAULT_TRAIN_OUT}")
    parser.add_argument("--val_out", default=DEFAULT_VAL_OUT,
                        help=f"Fichier val JSONL. Défaut: {DEFAULT_VAL_OUT}")

    parser.add_argument("--train_ratio", type=float, default=DEFAULT_TRAIN_RATIO,
                        help=f"Proportion train. Défaut: {DEFAULT_TRAIN_RATIO} ({DEFAULT_TRAIN_RATIO*100:.0f}%%)")
    parser.add_argument("--seed", type=int, default=DEFAULT_SEED,
                        help=f"Seed pour split reproductible. Défaut: {DEFAULT_SEED}")

    parser.add_argument("--min_chars", type=int, default=None,
                        help=f"Override MIN_CONTENT_CHARS. Défaut: {MIN_CONTENT_CHARS}")
    parser.add_argument("--max_chars", type=int, default=None,
                        help=f"Override MAX_CONTENT_CHARS. Défaut: {MAX_CONTENT_CHARS}")

    parser.add_argument("--report", action="store_true",
                        help="Affiche un rapport détaillé (manquants/doublons/filtrés).")
    parser.add_argument("--dry-run", action="store_true",
                        help="N'écrit pas les fichiers, affiche seulement les stats.")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Mode verbose (plus de détails).")

    args = parser.parse_args()

    # Override config depuis args
    global MIN_CONTENT_CHARS, MAX_CONTENT_CHARS
    if args.min_chars is not None:
        MIN_CONTENT_CHARS = args.min_chars
    if args.max_chars is not None:
        MAX_CONTENT_CHARS = args.max_chars

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    print_config_summary()

    edb_dir = Path(args.edb_dir).resolve()
    ndc_dir = Path(args.ndc_dir).resolve()
    train_path = Path(args.train_out).resolve()
    val_path = Path(args.val_out).resolve()

    if not edb_dir.exists():
        logger.error(f"Dossier EDB introuvable: {edb_dir}")
        return 1
    if not ndc_dir.exists():
        logger.error(f"Dossier NDC introuvable: {ndc_dir}")
        return 1

    if not (0.0 < args.train_ratio < 1.0):
        logger.error("--train_ratio doit être entre 0 et 1 (ex: 0.9).")
        return 1

    logger.info(f"EDB dir: {edb_dir}")
    logger.info(f"NDC dir: {ndc_dir}")

    # Indexer les dossiers
    logger.info("Indexation des fichiers EDB...")
    edb_index, edb_dups, edb_no_ref, edb_invalid = index_folder_by_ref(edb_dir)

    logger.info("Indexation des fichiers NDC...")
    ndc_index, ndc_dups, ndc_no_ref, ndc_invalid = index_folder_by_ref(ndc_dir)

    edb_refs = set(edb_index.keys())
    ndc_refs = set(ndc_index.keys())

    common_refs = sorted(edb_refs & ndc_refs)
    only_edb = sorted(edb_refs - ndc_refs)
    only_ndc = sorted(ndc_refs - edb_refs)

    # Construire toutes les paires valides
    pairs = []
    for ref in common_refs:
        edb_path, edb_txt = edb_index[ref]
        ndc_path, ndc_txt = ndc_index[ref]
        pairs.append((ref, edb_txt, ndc_txt))

    # Shuffle + split
    rng = random.Random(args.seed)
    rng.shuffle(pairs)

    n_total = len(pairs)
    n_train = int(round(n_total * args.train_ratio))
    # Pour éviter val vide sur petits datasets
    n_train = min(max(n_train, 1), n_total - 1) if n_total >= 2 else n_total

    train_pairs = pairs[:n_train]
    val_pairs = pairs[n_train:]

    train_records = [make_record(edb, ndc) for (_, edb, ndc) in train_pairs]
    val_records = [make_record(edb, ndc) for (_, edb, ndc) in val_pairs]

    # Écrire les fichiers (sauf dry-run)
    if not args.dry_run:
        write_jsonl(train_path, train_records)
        write_jsonl(val_path, val_records)

    # Résumé
    print()
    print("=" * 50)
    print("RÉSUMÉ")
    print("=" * 50)
    print(f"EDB refs détectées      : {len(edb_index)}")
    print(f"NDC refs détectées      : {len(ndc_index)}")
    print(f"EDB filtrés (invalides) : {len(edb_invalid)}")
    print(f"NDC filtrés (invalides) : {len(ndc_invalid)}")
    print(f"Paires appariées        : {len(common_refs)}")
    print(f"Paires utilisables      : {n_total}")
    print()
    print(f"Train : {len(train_records):5d} ({args.train_ratio*100:.1f}%)")
    print(f"Val   : {len(val_records):5d} ({(1-args.train_ratio)*100:.1f}%)")
    print(f"Seed  : {args.seed}")
    print()

    if args.dry_run:
        print("[DRY-RUN] Aucun fichier écrit.")
    else:
        print(f"Train -> {train_path}")
        print(f"Val   -> {val_path}")

    # Rapport détaillé
    if args.report:
        print()
        print("=" * 50)
        print("RAPPORT DÉTAILLÉ")
        print("=" * 50)

        if edb_invalid:
            print(f"\nEDB filtrés ({len(edb_invalid)}):")
            for p, reason in edb_invalid[:30]:
                print(f"  - {p.name}: {reason}")
            if len(edb_invalid) > 30:
                print(f"  ... et {len(edb_invalid) - 30} autres")

        if ndc_invalid:
            print(f"\nNDC filtrés ({len(ndc_invalid)}):")
            for p, reason in ndc_invalid[:30]:
                print(f"  - {p.name}: {reason}")
            if len(ndc_invalid) > 30:
                print(f"  ... et {len(ndc_invalid) - 30} autres")

        if edb_no_ref:
            print(f"\nEDB sans ref dans le NOM ({len(edb_no_ref)}):")
            for p in edb_no_ref[:30]:
                print(f"  - {p.name}")
            if len(edb_no_ref) > 30:
                print(f"  ... et {len(edb_no_ref) - 30} autres")

        if ndc_no_ref:
            print(f"\nNDC sans ref dans le NOM ({len(ndc_no_ref)}):")
            for p in ndc_no_ref[:30]:
                print(f"  - {p.name}")
            if len(ndc_no_ref) > 30:
                print(f"  ... et {len(ndc_no_ref) - 30} autres")

        if edb_dups:
            print(f"\nDoublons EDB ({len(edb_dups)} refs):")
            for ref, paths in list(edb_dups.items())[:20]:
                print(f"  - {ref}:")
                for p in paths:
                    print(f"      * {p.name}")
            if len(edb_dups) > 20:
                print(f"  ... et {len(edb_dups) - 20} autres refs")

        if ndc_dups:
            print(f"\nDoublons NDC ({len(ndc_dups)} refs):")
            for ref, paths in list(ndc_dups.items())[:20]:
                print(f"  - {ref}:")
                for p in paths:
                    print(f"      * {p.name}")
            if len(ndc_dups) > 20:
                print(f"  ... et {len(ndc_dups) - 20} autres refs"
                      )

        print(f"\nUniquement dans EDB ({len(only_edb)}):")
        if only_edb:
            print(f"  {', '.join(only_edb[:20])}")
            if len(only_edb) > 20:
                print(f"  ... et {len(only_edb) - 20} autres")
        else:
            print("  (aucun)")

        print(f"\nUniquement dans NDC ({len(only_ndc)}):")
        if only_ndc:
            print(f"  {', '.join(only_ndc[:20])}")
            if len(only_ndc) > 20:
                print(f"  ... et {len(only_ndc) - 20} autres")
        else:
            print("  (aucun)")

        # Aperçu des refs en validation
        if val_pairs:
            print(f"\nRefs en validation ({len(val_pairs)}):")
            val_refs = [ref for (ref, _, _) in val_pairs]
            print(f"  {', '.join(val_refs[:20])}")
            if len(val_refs) > 20:
                print(f"  ... et {len(val_refs) - 20} autres")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
