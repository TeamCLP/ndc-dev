import os
import shutil
import pandas as pd

# ================================
# PARAMÈTRES
# ================================
DATA_FOLDER = r"C:\Users\U38FP75\Downloads\Data"
OUTPUT_FOLDER = os.path.join(DATA_FOLDER, "clean2")
EXCEL_PATH = os.path.join(DATA_FOLDER, "analyse_documents.xlsx")
OUTPUT_EXCEL = "analyse_documents_enrichi.xlsx"

WORD_EXT = {".doc", ".docx"}

# ================================
# PRÉPARATION
# ================================
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

df = pd.read_excel(EXCEL_PATH, engine="openpyxl")

# ================================
# FICHIERS PRÉSENTS DANS DATA
# ================================
files_on_disk = {
    f for f in os.listdir(DATA_FOLDER)
    if os.path.isfile(os.path.join(DATA_FOLDER, f))
}

# ================================
# INDEX DISQUE : Base_Name → extensions
# ================================
disk_index = {}

for filename in files_on_disk:
    base, ext = os.path.splitext(filename)
    disk_index.setdefault(base, set()).add(ext.lower())

# ================================
# COLONNES DE SORTIE
# ================================
df["Statut_Fichier"] = "SUPPRIME"
df["Nom_Fichier_Clean2"] = ""

files_to_keep = set()

# ================================
# PIPELINE PRINCIPAL
# ================================
for idx, row in df.iterrows():
    filename = row["Filename_Original"]

    if filename not in files_on_disk:
        continue

    base, ext = os.path.splitext(filename)
    ext = ext.lower()

    # ---- RÈGLE 1 : AUTRE
    if row["Type_Document"] == "AUTRE":
        df.at[idx, "Statut_Fichier"] = "SUPPRIME"
        continue

    # ---- RÈGLE 2 : PDF supprimé si Word existe
    if ext == ".pdf" and disk_index.get(base, set()) & WORD_EXT:
        df.at[idx, "Statut_Fichier"] = "SUPPRIME"
        continue

    # ---- FICHIER CONSERVÉ
    files_to_keep.add(filename)
    df.at[idx, "Statut_Fichier"] = "CONSERVE"

    # ---- RENOMMAGE
    ritm = (
        row["Reference"]
        if str(row["Reference"]).startswith("CAGIPRITM")
        else row["RITM_Parent"]
    )

    parts = filename.split("-", 1)
    rest = parts[1] if len(parts) > 1 else filename

    new_name = f"{ritm}-{row['Type_Document']}-{rest}"
    df.at[idx, "Nom_Fichier_Clean2"] = new_name

    # ---- COPIE
    src = os.path.join(DATA_FOLDER, filename)
    dst = os.path.join(OUTPUT_FOLDER, new_name)
    shutil.copy2(src, dst)

# ================================
# EXPORT EXCEL ENRICHI
# ================================
df.to_excel(
    OUTPUT_EXCEL,
    index=False,
    engine="openpyxl"
)

print("✅ Pipeline de nettoyage terminé sans erreur")
print(f"📁 Fichiers conservés et renommés : {OUTPUT_FOLDER}")
print(f"📄 Fichier Excel enrichi : {OUTPUT_EXCEL}")
