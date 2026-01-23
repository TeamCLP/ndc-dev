
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import re
import argparse
import random
from pathlib import Path
from collections import defaultdict
from typing import Dict, Tuple, List, Optional

# Format attendu dans le NOM de fichier (sans extension) :
# CAGIPRITM<digits> - ...
# Accepte séparateurs: -, – (en-dash), — (em-dash), _ ; avec ou sans espaces
REF_FROM_FILENAME_RE = re.compile(
    r"^(CAGIPRITM\d+)\s*[-–—_]\s*.*$",
    re.IGNORECASE
)

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

def index_folder_by_ref(folder: Path) -> Tuple[Dict[str, Tuple[Path, str]], Dict[str, List[Path]], List[Path]]:
    """
    Indexe tous les .md par référence.
    Retourne:
      - index: ref -> (path, content) (1 fichier retenu par ref)
      - duplicates: ref -> [paths...] si plusieurs fichiers ont la même ref
      - no_ref: liste des fichiers sans ref détectée
    """
    index: Dict[str, Tuple[Path, str]] = {}
    duplicates: Dict[str, List[Path]] = defaultdict(list)
    no_ref: List[Path] = []

    for p in sorted(folder.rglob("*.md")):
        if p.name.startswith("."):
            continue

        ref = extract_ref_from_filename(p)
        if not ref:
            no_ref.append(p)
            continue

        content = read_text(p).strip()

        if ref in index:
            duplicates[ref].append(p)
        else:
            index[ref] = (p, content)

    # Pour lisibilité: inclure aussi le premier fichier dans la liste de doublons
    for ref, paths in list(duplicates.items()):
        paths.insert(0, index[ref][0])

    return index, duplicates, no_ref

def make_record(edb_text: str, ndc_text: str) -> dict:
    return {
        "messages": [
            {"role": "user", "content": f"[INST] {edb_text} [/INST]"},
            {"role": "assistant", "content": ndc_text}
        ]
    }

def write_jsonl(path: Path, records: List[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        for r in records:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

def main():
    parser = argparse.ArgumentParser(
        description="Construit 2 datasets JSONL (train/val) pour Mistral Instruct depuis edb/ et ndc/ appariés par ref dans le nom de fichier."
    )
    parser.add_argument("--edb_dir", default="edb", help="Dossier EDB (entrée user). Défaut: edb")
    parser.add_argument("--ndc_dir", default="ndc", help="Dossier NDC (sortie assistant). Défaut: ndc")

    parser.add_argument("--train_out", default="train_dataset.jsonl", help="Fichier train JSONL. Défaut: train_dataset.jsonl")
    parser.add_argument("--val_out", default="val_dataset.jsonl", help="Fichier val JSONL. Défaut: val_dataset.jsonl")

    parser.add_argument("--train_ratio", type=float, default=0.9, help="Proportion train. Défaut: 0.9 (90%%)")
    parser.add_argument("--seed", type=int, default=42, help="Seed pour split reproductible. Défaut: 42")
    parser.add_argument("--min_chars", type=int, default=1, help="Ignore les paires trop courtes. Défaut: 1")
    parser.add_argument("--report", action="store_true", help="Affiche un rapport détaillé (manquants/doublons/sans ref).")

    args = parser.parse_args()

    edb_dir = Path(args.edb_dir).resolve()
    ndc_dir = Path(args.ndc_dir).resolve()
    train_path = Path(args.train_out).resolve()
    val_path = Path(args.val_out).resolve()

    if not edb_dir.exists():
        raise FileNotFoundError(f"Dossier EDB introuvable: {edb_dir}")
    if not ndc_dir.exists():
        raise FileNotFoundError(f"Dossier NDC introuvable: {ndc_dir}")

    if not (0.0 < args.train_ratio < 1.0):
        raise ValueError("--train_ratio doit être entre 0 et 1 (ex: 0.9).")

    edb_index, edb_dups, edb_no_ref = index_folder_by_ref(edb_dir)
    ndc_index, ndc_dups, ndc_no_ref = index_folder_by_ref(ndc_dir)

    edb_refs = set(edb_index.keys())
    ndc_refs = set(ndc_index.keys())

    common_refs = sorted(edb_refs & ndc_refs)
    only_edb = sorted(edb_refs - ndc_refs)
    only_ndc = sorted(ndc_refs - edb_refs)

    # Construire toutes les paires
    pairs = []
    skipped_short = 0
    for ref in common_refs:
        edb_path, edb_txt = edb_index[ref]
        ndc_path, ndc_txt = ndc_index[ref]

        if len(edb_txt) < args.min_chars or len(ndc_txt) < args.min_chars:
            skipped_short += 1
            continue

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

    write_jsonl(train_path, train_records)
    write_jsonl(val_path, val_records)

    print("✅ Génération terminée")
    print(f"EDB refs détectées      : {len(edb_index)}")
    print(f"NDC refs détectées      : {len(ndc_index)}")
    print(f"Paires appariées (brut) : {len(common_refs)}")
    print(f"Paires utilisables      : {n_total}")
    if skipped_short:
        print(f"Paires ignorées (trop court) : {skipped_short}")
    print(f"Train: {len(train_records)} ({args.train_ratio*100:.1f}%) -> {train_path}")
    print(f"Val  : {len(val_records)} ({(1-args.train_ratio)*100:.1f}%) -> {val_path}")
    print(f"Seed : {args.seed}")

    if args.report:
        print("\n=== RAPPORT ===")

        if edb_no_ref:
            print(f"\nEDB sans ref dans le NOM ({len(edb_no_ref)}):")
            for p in edb_no_ref[:50]:
                print(f" - {p.name}")
            if len(edb_no_ref) > 50:
                print(" ...")

        if ndc_no_ref:
            print(f"\nNDC sans ref dans le NOM ({len(ndc_no_ref)}):")
            for p in ndc_no_ref[:50]:
                print(f" - {p.name}")
            if len(ndc_no_ref) > 50:
                print(" ...")

        if edb_dups:
            print(f"\nDoublons EDB ({len(edb_dups)} refs):")
            for ref, paths in list(edb_dups.items())[:30]:
                print(f" - {ref}:")
                for p in paths:
                    print(f"    * {p.name}")
            if len(edb_dups) > 30:
                print(" ...")

        if ndc_dups:
            print(f"\nDoublons NDC ({len(ndc_dups)} refs):")
            for ref, paths in list(ndc_dups.items())[:30]:
                print(f" - {ref}:")
                for p in paths:
                    print(f"    * {p.name}")
            if len(ndc_dups) > 30:
                print(" ...")

        print(f"\nUniquement dans EDB ({len(only_edb)}): {', '.join(only_edb) if only_edb else '(aucun)'}")
        print(f"Uniquement dans NDC ({len(only_ndc)}): {', '.join(only_ndc) if only_ndc else '(aucun)'}")

        # Afficher un aperçu des refs envoyées en val (utile)
        if val_pairs:
            print("\nRefs en validation:")
            print(", ".join([ref for (ref, _, _) in val_pairs]))

if __name__ == "__main__":
    main()
