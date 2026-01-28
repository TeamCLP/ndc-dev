import os
import pandas as pd

# ================================
# PARAMÈTRES
# ================================
CLEAN_FOLDER = r"C:\Users\U38FP75\Downloads\Data\clean2"
INPUT_EXCEL = "analyse_documents_enrichi.xlsx"
OUTPUT_EXCEL = os.path.join(CLEAN_FOLDER, "couverture_EDB_NDC_par_RITM.xlsx")

# ================================
# FICHIERS PRÉSENTS DANS CLEAN2
# ================================
files_in_clean2 = {
    f for f in os.listdir(CLEAN_FOLDER)
    if os.path.isfile(os.path.join(CLEAN_FOLDER, f))
}

# ================================
# CHARGEMENT EXCEL ENRICHI
# ================================
df = pd.read_excel(INPUT_EXCEL, engine="openpyxl")

# ================================
# EXTRACTION DES RITM (A + G)
# ================================
ritm_cols = ["Reference", "RITM_Parent"]

ritms = (
    df[ritm_cols]
    .melt(value_name="RITM")
    .dropna()
)

ritms = ritms[
    ritms["RITM"].str.startswith("CAGIPRITM", na=False)
]["RITM"].drop_duplicates()

# ================================
# FILTRE FICHIERS CONSERVÉS EDB / NDC
# ================================
df_docs = df[
    (df["Statut_Fichier"] == "CONSERVE") &
    (df["Type_Document"].isin(["EDB", "NDC"])) &
    (df["Nom_Fichier_Clean2"].isin(files_in_clean2))
].copy()

# ================================
# ANALYSE PAR RITM
# ================================
results = []

for ritm in ritms:
    subset = df_docs[
        (df_docs["Reference"] == ritm) |
        (df_docs["RITM_Parent"] == ritm)
    ]

    edb_files = subset[
        subset["Type_Document"] == "EDB"
    ]["Nom_Fichier_Clean2"].tolist()

    ndc_files = subset[
        subset["Type_Document"] == "NDC"
    ]["Nom_Fichier_Clean2"].tolist()

    # ✅ NOUVEAU CONTRÔLE PDF
    has_pdf = any(
        f.lower().endswith(".pdf")
        for f in edb_files + ndc_files
    )

    results.append({
        "RITM": ritm,
        "Nb_EDB": len(edb_files),
        "Nb_NDC": len(ndc_files),
        "Couple_EDB_NDC": "OUI" if edb_files and ndc_files else "NON",
        "Presence_PDF_EDB_NDC": "OUI" if has_pdf else "NON",
        "Documents_EDB": " | ".join(edb_files) if edb_files else "",
        "Documents_NDC": " | ".join(ndc_files) if ndc_files else ""
    })

# ================================
# EXPORT EXCEL
# ================================
pd.DataFrame(results).sort_values("RITM").to_excel(
    OUTPUT_EXCEL,
    index=False,
    engine="openpyxl"
)

print("✅ Contrôle EDB / NDC + présence PDF terminé")
print(f"📄 Fichier généré : {OUTPUT_EXCEL}")
