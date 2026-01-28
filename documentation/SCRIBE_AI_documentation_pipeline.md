# SCRIBE AI — Documentation Pipeline de Données

> **De la collecte des données à l'entraînement du modèle**

📁 Repository : [github.com/TeamCLP/ndc-dev](https://github.com/TeamCLP/ndc-dev)

---

## Table des matières

1. [Vue d'ensemble](#1-vue-densemble)
2. [Schéma du Pipeline](#2-schéma-du-pipeline)
3. [Inventaire des Programmes](#3-inventaire-des-programmes)
4. [Étape 1 — Collecte](#4-étape-1--collecte-des-données)
5. [Étape 2 — Classification](#5-étape-2--classification)
6. [Étape 3 — Nettoyage](#6-étape-3--nettoyage--déduplication)
7. [Étape 4 — Appariement](#7-étape-4--appariement-edb-ndc)
8. [Étape 5 — Conversion](#8-étape-5--conversion-markdown)
9. [Étape 6 — Préparation Dataset](#9-étape-6--préparation-du-dataset)
10. [Étape 7 — Entraînement](#10-étape-7--entraînement-fine-tuning)
11. [Comparaison train.py vs train2.py](#11-comparaison-trainpy-vs-train2py)
12. [Structure du Repository](#12-structure-du-repository)

---

## 1. Vue d'ensemble

**SCRIBE AI** automatise la génération de Notes de Cadrage (NDC) à partir d'Expressions de Besoins (EDB) en utilisant un modèle **Mistral 7B fine-tuné** sur des données historiques internes du domaine bancaire.

### Deux approches d'entraînement

| Version | Script | Description | Cas d'usage |
|---------|--------|-------------|-------------|
| **V1** | `train.py` | Génération par **champs individuels** avec balises `<START>/<END>` | Interface interactive |
| **V2** | `train2.py` | Génération de **documents complets** (EDB → Devis Markdown) | Génération batch |

---

## 2. Schéma du Pipeline

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                           PIPELINE SCRIBE AI                                │
└─────────────────────────────────────────────────────────────────────────────┘

  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
  │  1.COLLECTE │───▶│2.CLASSIF.   │───▶│ 3.NETTOYAGE │───▶│4.APPARIEMENT│
  │  SQL+DL     │    │  Scoring    │    │  Dédupe     │    │  EDB-NDC    │
  └─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
         │                                                        │
         ▼                                                        ▼
  ┌─────────────┐    ┌─────────────┐    ┌─────────────┐    ┌─────────────┐
  │ 7.TRAINING  │◀───│ 6.DATASET   │◀───│5.CONVERSION │◀───│   Couples   │
  │ train.py    │    │  JSONL      │    │  Markdown   │    │   EDB-NDC   │
  │ train2.py   │    │             │    │             │    │             │
  └─────────────┘    └─────────────┘    └─────────────┘    └─────────────┘
         │
         ▼
  ┌─────────────┐
  │   MODÈLE    │
  │  FINE-TUNÉ  │
  │   + LoRA    │
  └─────────────┘
```

---

## 3. Inventaire des Programmes

| # | Étape | Programme | Statut | Description |
|---|-------|-----------|--------|-------------|
| 1 | Collecte | `extraction_sql.sql` | 🔴 À déposer | Requête SQL pour identifier les documents |
| 2 | Collecte | `download_files.py` | 🔴 À déposer | Téléchargement depuis l'outil interne |
| 3 | Classification | `classify_documents.py` | 🔴 À déposer | Scoring automatique EDB vs NDC vs AUTRE |
| 4 | Nettoyage | `pipeline_nettoyage_complet.py` | 🟢 Présent | Déduplication + renommage RITM |
| 5 | Appariement | `couples_edb_ndc.py` | 🟢 Présent | Matching EDB-NDC par RITM |
| 6 | Conversion | `extract.py` | 🟢 Présent | Extraction contenu Word |
| 7 | Conversion | `convert_devis_et_edb_docx_to_markdown.py` | 🟢 Présent | Conversion DOCX → Markdown |
| 8 | Préparation | `build_dataset_jsonl.py` | 🟢 Présent | Construction dataset JSONL |
| 9 | Entraînement | `train.py` | 🟢 Présent | Fine-tuning V1 — Champs |
| 10 | Entraînement | `train2.py` | 🟢 Présent | Fine-tuning V2 — Documents |

---

## 4. Étape 1 — Collecte des Données

> ⏳ **Scripts à déposer**

### 1.1 extraction_sql.sql

**Rôle :** Requête SQL pour identifier les documents EDB et NDC dans l'outil interne.

**Sortie :** Liste des fichiers à télécharger avec métadonnées (référence RITM, type, chemin)

### 1.2 download_files.py

**Rôle :** Téléchargement des fichiers identifiés.

**Sortie :** Dossier `Data/` contenant les fichiers bruts (PDF + Word)

---

## 5. Étape 2 — Classification

> ⏳ **Script à déposer**

### classify_documents.py

**Rôle :** Classification automatique des documents en EDB, NDC ou AUTRE via scoring.

**Entrée :** Dossier `Data/`

**Sortie :** `analyse_documents.xlsx`

**Colonnes de sortie :**
- `Filename_Original` — Nom du fichier
- `Reference` — Référence RITM extraite
- `RITM_Parent` — RITM parent si applicable
- `Type_Document` — EDB | NDC | AUTRE
- `Score_EDB` / `Score_NDC` — Scores de classification

---

## 6. Étape 3 — Nettoyage & Déduplication

### pipeline_nettoyage_complet.py 🟢

**Entrées :**
- `analyse_documents.xlsx`
- Dossier `Data/`

**Sorties :**
- Dossier `clean2/`
- `analyse_documents_enrichi.xlsx`

**Règles de traitement :**

| Règle | Action |
|-------|--------|
| Type = `AUTRE` | → Supprimé |
| PDF existe ET Word avec même nom | → PDF supprimé, Word conservé |
| Fichier conservé | → Renommé `{RITM}-{TYPE}-{nom}.ext` |
| Fichier renommé | → Copié vers `clean2/` |

**Colonnes ajoutées :**
- `Statut_Fichier` : CONSERVE | SUPPRIME
- `Nom_Fichier_Clean2` : Nouveau nom

**Exécution :**
```bash
python pipeline_nettoyage_complet.py
```

---

## 7. Étape 4 — Appariement EDB-NDC

### couples_edb_ndc.py 🟢

**Entrées :**
- `analyse_documents_enrichi.xlsx`
- Dossier `clean2/`

**Sortie :** `couverture_EDB_NDC_par_RITM.xlsx`

**Logique :**
1. Extraction des RITM uniques
2. Pour chaque RITM : comptage EDB et NDC
3. Identification des couples complets (≥1 EDB + ≥1 NDC)
4. Détection présence PDF

**Colonnes du rapport :**

| Colonne | Description |
|---------|-------------|
| `RITM` | Référence unique |
| `Nb_EDB` | Nombre d'EDB |
| `Nb_NDC` | Nombre de NDC |
| `Couple_EDB_NDC` | OUI si couple complet |
| `Presence_PDF_EDB_NDC` | OUI si PDF présent |
| `Documents_EDB` | Liste fichiers EDB |
| `Documents_NDC` | Liste fichiers NDC |

**Exécution :**
```bash
python couples_edb_ndc.py
```

---

## 8. Étape 5 — Conversion Markdown

### extract.py 🟢

**Rôle :** Extraction de contenu textuel depuis documents Word.

### convert_devis_et_edb_docx_to_markdown.py 🟢

**Entrée :** Fichiers DOCX des couples EDB-NDC

**Sortie :** Fichiers `.md`

**Outil :** Docling (IBM)

---

## 9. Étape 6 — Préparation du Dataset

### build_dataset_jsonl.py 🟢

**Entrée :** Fichiers Markdown

**Sorties :**
- `dataset/train_dataset.jsonl`
- `dataset/val_dataset.jsonl`

**Format JSONL — Mistral Instruct :**

```json
{
  "messages": [
    {
      "role": "user",
      "content": "[INST] <TASK>contexte_proj</TASK>\n<CONTEXT>\nclient: Banque ABC\n</CONTEXT> [/INST]"
    },
    {
      "role": "assistant",
      "content": "<START>Le projet s'inscrit dans le cadre...<END>"
    }
  ]
}
```

---

## 10. Étape 7 — Entraînement (Fine-tuning)

### train.py — V1 Champs individuels 🟢

**Cas d'usage :** Génération de champs individuels avec balises `<START>/<END>`

**Configuration :**

| Paramètre | Valeur |
|-----------|--------|
| Modèle | `mistralai/Mistral-7B-Instruct-v0.3` |
| Output dir | `/home/quentin/mistral-banking` |
| Max prompt | 1 536 tokens |
| Max response | 768 tokens |
| Max total | 2 304 tokens |
| LoRA r / alpha | 128 / 256 |
| Batch size | 8 |
| Gradient accum | 4 |
| Learning rate | 3e-5 |
| Epochs | 2 |
| Precision | bfloat16 |

**Commandes :**
```bash
python train.py                              # From scratch
python train.py --resume                     # Reprendre dernier checkpoint
python train.py --resume-from /path/to/ckpt  # Checkpoint spécifique
```

---

### train2.py — V2 Documents complets 🟢

**Cas d'usage :** Génération de documents Markdown complets (EDB → Devis)

**Configuration :**

| Paramètre | Valeur |
|-----------|--------|
| Modèle | `mistralai/Mistral-7B-Instruct-v0.3` |
| Output dir | `/home/quentin/mistral-devis` |
| Max prompt | **6 144 tokens** |
| Max response | **8 192 tokens** |
| Max total | **14 336 tokens** |
| LoRA r / alpha | 128 / 256 |
| Batch size | **2** |
| Gradient accum | **16** |
| Learning rate | **2e-5** |
| Epochs | **3** |
| Precision | bfloat16 |

**Commandes :**
```bash
python train2.py                             # From scratch
python train2.py --resume                    # Reprendre
python train2.py --output-dir /path/to/out   # Dossier personnalisé
python train2.py --max-prompt-length 4096    # Longueurs personnalisées
```

---

### Monitoring TensorBoard

```bash
# V1
tensorboard --logdir=/home/quentin/runs/mistral-banking

# V2
tensorboard --logdir=/home/quentin/runs/mistral-devis
```

**Métriques :**
- `train/loss` — Loss d'entraînement
- `eval/loss` — Loss de validation
- `val_gen/tag_rate` — Taux balises correctes (V1)
- `val_gen/avg_generation_time` — Temps génération
- `val_gen/avg_tokens_generated` — Tokens générés (V2)

---

## 11. Comparaison train.py vs train2.py

| Aspect | train.py (V1) | train2.py (V2) |
|--------|---------------|----------------|
| **Objectif** | Champs individuels | Documents complets |
| **Format entrée** | `<TASK>...<CONTEXT>` | EDB complète (MD) |
| **Format sortie** | `<START>...<END>` | Devis complet (MD) |
| **Max prompt** | 1 536 tokens | 6 144 tokens |
| **Max response** | 768 tokens | 8 192 tokens |
| **Batch size** | 8 | 2 |
| **Gradient accum** | 4 | 16 |
| **Learning rate** | 3e-5 | 2e-5 |
| **Epochs** | 2 | 3 |
| **Validation** | Tags START/END | Aperçu devis |
| **Cas d'usage** | Interface interactive | Génération batch |

---

## 12. Structure du Repository

```
ndc-dev/
├── README.md
├── requirements.txt
├── install.sh
├── run_train.sh
│
├── ─── PRÉPARATION ───
├── pipeline_nettoyage_complet.py
├── couples_edb_ndc.py
├── extract.py
├── convert_devis_et_edb_docx_to_markdown.py
├── build_dataset_jsonl.py
│
├── ─── ENTRAÎNEMENT ───
├── train.py                    # V1 - Champs
├── train2.py                   # V2 - Documents
│
└── dataset/
    ├── train_dataset.jsonl
    └── val_dataset.jsonl
```

### Dépendances

```
torch==2.9.1
transformers==4.57.3
peft==0.18.1
datasets==4.4.2
accelerate==1.12.0
pandas
openpyxl
docling
```

### Prérequis matériels

| Ressource | Minimum | Recommandé |
|-----------|---------|------------|
| GPU VRAM | 24 Go (V1) | 80 Go (H100) |
| CUDA | 12.0+ | 12.8+ |
| Python | 3.10+ | 3.12 |

---

## TODO — Scripts à déposer

- [ ] Script d'extraction SQL
- [ ] Script de téléchargement
- [ ] Script de classification/scoring
- [x] pipeline_nettoyage_complet.py
- [x] couples_edb_ndc.py
- [x] extract.py
- [x] convert_devis_et_edb_docx_to_markdown.py
- [x] build_dataset_jsonl.py
- [x] train.py
- [x] train2.py

---

*SCRIBE AI — Documentation Pipeline de Données*
*Dernière mise à jour : Janvier 2025*
