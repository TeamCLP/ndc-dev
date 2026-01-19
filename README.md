# SCRIBE AI - Générateur Intelligent de Notes de Cadrage

SCRIBE AI est un assistant intelligent de génération de documents de cadrage projet (Notes de Cadrage - NDC) utilisant un modèle de langage fine-tuné. Conçu pour automatiser la rédaction de spécifications projet dans le domaine bancaire et entreprise.

## Fonctionnalités

- **Génération intelligente de contenu** : 23+ champs générables avec conscience du contexte
- **Génération de tableaux** : Contraintes, risques, phases, livrables, jalons, coûts
- **Export Word** : Document DOCX complet avec mise en forme professionnelle
- **Système de feedback** : Notation et corrections pour amélioration continue
- **Interface responsive** : Application web moderne et intuitive

## Stack Technique

### Backend
- **Framework** : FastAPI (Python)
- **Modèle IA** : Mistral-7B-Instruct-v0.3 fine-tuné avec LoRA
- **Entraînement** : PyTorch + Hugging Face Transformers + PEFT
- **Précision** : bfloat16 (optimisé pour GPU H100 80GB)

### Frontend
- **Architecture** : JavaScript ES6 (SPA vanilla)
- **Export** : docx.js pour génération Word
- **Graphiques** : Chart.js
- **Composants** : Système de composants custom (Sidebar, Section, ArrayField, Modal)

## Structure du Projet

```
ndc-dev/
├── api.py                    # API FastAPI backend
├── train.py                  # Script d'entraînement du modèle
├── extract.py                # Extracteur de documents Word
├── requirements.txt          # Dépendances Python
├── run.sh                    # Script de démarrage des services
├── run_train.sh              # Lanceur d'entraînement
├── install.sh                # Script d'installation
│
├── dataset/
│   ├── train_dataset.jsonl   # Données d'entraînement
│   └── val_dataset.jsonl     # Données de validation
│
└── html/                     # Application frontend
    ├── index.html            # Point d'entrée HTML
    ├── serv.py               # Serveur HTTP avec proxy
    ├── js/
    │   ├── app.js            # Application principale
    │   ├── config.js         # Configuration API et champs
    │   ├── core/
    │   │   ├── Api.js        # Client API
    │   │   ├── Router.js     # Routage client
    │   │   └── Store.js      # Gestion d'état
    │   ├── components/       # Composants UI
    │   └── services/         # Services métier
    │       ├── GenerationService.js
    │       ├── ExportService.js
    │       ├── ValidationService.js
    │       ├── ProjectService.js
    │       └── GanttService.js
    └── css/                  # Styles
```

## Installation

### Prérequis
- Python 3.12+
- CUDA 12.8+ (GPU NVIDIA recommandé)
- 80GB+ VRAM pour inférence optimale (H100)

### Installation automatique

```bash
./install.sh
```

### Installation manuelle

```bash
# Créer l'environnement conda
conda create -n ndc-dev python=3.12
conda activate ndc-dev

# Installer les dépendances
pip install -r requirements.txt
```

## Démarrage

### Lancer les services

```bash
./run.sh
```

Cela démarre :
- **API Backend** sur le port `5000`
- **Serveur Frontend** sur le port `6006`

### Accès à l'application

Ouvrir dans le navigateur : `http://localhost:6006`

## API Endpoints

| Méthode | Endpoint | Description |
|---------|----------|-------------|
| GET | `/` | Informations du modèle |
| GET | `/health` | État de santé + infos GPU |
| GET | `/model/info` | Configuration détaillée |
| POST | `/generate` | Générer un champ |
| POST | `/generate_multiple` | Générer plusieurs champs (max 5) |
| POST | `/validate` | Valider une valeur |
| POST | `/feedback/correction` | Soumettre une correction |
| POST | `/feedback/rating` | Soumettre une notation |
| GET | `/feedback/insights/{field}` | Insights par champ |
| GET | `/feedback/summary` | Résumé des feedbacks |

### Exemple de requête

```bash
curl -X POST http://localhost:5000/generate \
  -H "Content-Type: application/json" \
  -d '{
    "field": "contexte_proj",
    "context": {
      "client": "Banque XYZ",
      "secteur": "Bancaire",
      "complexite": "medium"
    },
    "strategy": "related"
  }'
```

### Stratégies de génération

| Stratégie | Description |
|-----------|-------------|
| `minimal` | Contexte basique avec 2 champs pertinents |
| `related` | Complexité moyenne avec 3-4 champs (défaut) |
| `full` | Contexte complet avec 5-7 champs |

## Entraînement du Modèle

### Lancer l'entraînement

```bash
./run_train.sh

# Ou directement
python train.py

# Reprendre depuis un checkpoint
python train.py --resume

# Reprendre depuis un checkpoint spécifique
python train.py --resume-from /path/to/checkpoint
```

### Configuration d'entraînement

| Paramètre | Valeur |
|-----------|--------|
| Modèle de base | `mistralai/Mistral-7B-Instruct-v0.3` |
| LoRA rank (r) | 128 |
| LoRA alpha | 256 |
| Batch size | 8 |
| Gradient accumulation | 4 |
| Learning rate | 3e-5 |
| Epochs | 2 |

### Format du dataset

```json
{
  "messages": [
    {
      "role": "user",
      "content": "[INST] <TASK>contexte_proj</TASK>\n<CONTEXT>\nclient: Banque ABC\nsecteur: Bancaire\n</CONTEXT> [/INST]"
    },
    {
      "role": "assistant",
      "content": "<START>Le projet s'inscrit dans le cadre...</END>"
    }
  ]
}
```

## Champs Générables

### Champs textuels
- `contexte_proj` - Contexte du projet
- `besoin` - Besoin exprimé
- `objectifs` - Objectifs du projet
- `perimetre` - Périmètre
- `horsPerimetre` - Hors périmètre
- `descriptionSolution` - Description de la solution
- `architecture` - Architecture technique
- `composantsDimensionnement` - Composants et dimensionnement
- `conditionsHorsCrash` - Conditions hors crash site
- `conditionsCrashSite` - Conditions crash site
- `resilienceApplicative` - Résilience applicative
- `praPlanDegrade` - PRA / Plan dégradé
- `sauvegardes` - Sauvegardes
- `administrationSupervision` - Administration / Supervision
- `impactCO2` - Impact CO2 (RSE)
- `modalitesPartage` - Modalités de partage

### Champs tableaux
| Champ | Colonnes |
|-------|----------|
| `contraintes` | type, description, criticité, mitigation |
| `risques` | risque, probabilité, impact, planActions, responsable |
| `phases` | phase, description, dateDébut, dateFin, équipes |
| `livrables` | nom, description, date, responsable |
| `jalons` | nom, type, date, critères |
| `coutsConstruction` | profil, nombre_jh, tjm, total, code |
| `coutsFonctionnement` | poste, quantite, coutUnitaire, total, code |

## Configuration

### Variables d'environnement

| Variable | Description | Défaut |
|----------|-------------|--------|
| `API_HOST` | Hôte du serveur API | `0.0.0.0` |
| `API_PORT` | Port du serveur API | `5000` |
| `PORT` | Port du serveur frontend | `6006` |
| `HOSTNAME` | Nom du pod (Kubernetes) | auto |
| `PATH_PREFIX` | Préfixe URL | auto |

### Paramètres de génération

| Paramètre | Valeur |
|-----------|--------|
| Max prompt length | 1536 tokens |
| Max new tokens | 768 tokens |
| Temperature | 0.7 |
| Top-p | 0.9 |
| Repetition penalty | 1.2 |

## Déploiement

### Docker / Kubernetes

L'application supporte le déploiement Kubernetes avec détection automatique du préfixe de pod :

```
/scribe-ai/{pod-name}/url-1/api/*
```

### Support Proxy

Configuration proxy intégrée pour environnements d'entreprise (HTTP/HTTPS).

## Monitoring

### Logs de génération

Les générations sont loggées dans `generation_logs.jsonl` :

```json
{
  "timestamp": "2026-01-19T14:30:45.123456",
  "field": "contexte_proj",
  "prompt": "[INST] ...",
  "raw_response": "...",
  "parsed_response": "...",
  "error": null
}
```

### TensorBoard

Pendant l'entraînement :

```bash
tensorboard --logdir=/home/quentin/runs/mistral-banking
```

Métriques disponibles : `train/loss`, `val/loss`, `val_gen/tag_rate`

## Dépendances Principales

```
torch==2.9.1
transformers==4.57.3
peft==0.18.1
fastapi==0.128.0
uvicorn==0.40.0
datasets==4.4.2
accelerate==1.12.0
```

## Licence

Projet propriétaire - Usage interne uniquement.

## Auteur

Développé avec SCRIBE AI
