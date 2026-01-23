#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import base64
import gzip
import argparse
from pathlib import Path

def parse_b64txt(path: Path):
    lines = path.read_text(encoding="utf-8").splitlines()
    gzip_on = False
    b64_lines = []

    for line in lines:
        if line.startswith("#"):
            if line.startswith("# GZIP="):
                gzip_on = line.strip().endswith("1")
            continue
        if line.strip():
            b64_lines.append(line.strip())

    if not b64_lines:
        raise ValueError(f"Aucune donnée Base64 trouvée dans {path}")

    b64_str = "".join(b64_lines)
    return b64_str, gzip_on

def decode_b64txt_to_file(in_path: Path, out_path: Path) -> None:
    b64_str, gzip_on = parse_b64txt(in_path)
    data = base64.b64decode(b64_str.encode("ascii"))

    if gzip_on:
        data = gzip.decompress(data)

    out_path.write_bytes(data)

def main():
    p = argparse.ArgumentParser(description="Decode un fichier Base64 texte (copier/coller) vers un JSONL.")
    p.add_argument("input", help="Fichier .b64.txt (collé depuis Windows)")
    p.add_argument("--output", default="", help="Nom du fichier de sortie (par défaut: retire .b64.txt)")
    args = p.parse_args()

    in_path = Path(args.input).resolve()
    if not in_path.exists():
        raise FileNotFoundError(f"Introuvable: {in_path}")

    if args.output:
        out_path = Path(args.output).resolve()
    else:
        # train_dataset.jsonl.b64.txt -> train_dataset.jsonl
        name = in_path.name
        if name.endswith(".b64.txt"):
            name = name[:-len(".b64.txt")]
        else:
            name = name + ".decoded"
        out_path = in_path.with_name(name)

    decode_b64txt_to_file(in_path, out_path)
    print(f"✅ Reconversion terminée : {out_path}")

if __name__ == "__main__":
    main()
