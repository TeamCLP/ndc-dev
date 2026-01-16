#!/usr/bin/env python3
"""
API FastAPI pour SCRIBE AI - Version FULL PRECISION avec SDPA
Optimisée pour GPU haute mémoire (H100 80Go)
"""

import os
import json
import torch
import logging
import re
import time
import random
import numpy as np
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime, timedelta
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field

from transformers import AutoModelForCausalLM, AutoTokenizer

# Configuration du logging principal
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Configuration du logging des générations dans un fichier
generation_logger = logging.getLogger('generation')
generation_handler = logging.FileHandler('generation_logs.jsonl', mode='a')
generation_handler.setLevel(logging.INFO)
generation_logger.addHandler(generation_handler)
generation_logger.setLevel(logging.INFO)

# ===== CONFIGURATION =====
class Settings:
    # MODÈLE FULL PRECISION
    MODEL_PATH = "/home/quentin/mistral-banking/checkpoint-200"
    
    HOST = os.getenv("API_HOST", "0.0.0.0")
    PORT = int(os.getenv("API_PORT", "5000"))
    
    # Paramètres de génération augmentés pour full precision
    MAX_NEW_TOKENS = 768  # Augmenté de 512
    MAX_PROMPT_LENGTH = 1536  # Augmenté de 1024
    TEMPERATURE = 0.7
    TOP_P = 0.9
    REPETITION_PENALTY = 1.2
    
    DEVICE = "cuda" if torch.cuda.is_available() else "cpu"
    
    # Configuration full precision
    USE_SDPA = True  # Scaled Dot Product Attention (PyTorch natif)
    TORCH_DTYPE = "bfloat16"  # bfloat16 pour stabilité

settings = Settings()

# ===== MODÈLES PYDANTIC =====
class GenerateRequest(BaseModel):
    field: str
    context: Dict[str, Any]
    strategy: str = "related"

class GenerateMultipleRequest(BaseModel):
    fields: List[str]
    context: Dict[str, Any]
    strategy: str = "related"

class ValidateRequest(BaseModel):
    field: str
    value: Any
    context: Dict[str, Any] = {}

class FeedbackCorrectionRequest(BaseModel):
    field: str
    original: Any
    corrected: Any
    context: Dict[str, Any] = {}
    reason: Optional[str] = None

class FeedbackRatingRequest(BaseModel):
    field: str
    value: Any
    rating: int = Field(ge=1, le=5)
    context: Dict[str, Any] = {}
    comment: Optional[str] = None

# ===== CONSTANTES IDENTIQUES À L'ENTRAÎNEMENT =====
ALWAYS_INCLUDED = ['client', 'secteur', 'typeProjet', 'complexite', 'libelle', 'annee']

FIELD_GENERATION_ORDER = [
    "contexte_proj", "besoin", "objectifs", "perimetre", "horsPerimetre",
    "contraintes", "risques", "descriptionSolution", "architecture", 
    "composantsDimensionnement", "phases", "jalons", "livrables",
    "conditionsHorsCrash", "conditionsCrashSite", "resilienceApplicative",
    "praPlanDegrade", "sauvegardes", "administrationSupervision",
    "impactCO2", "modalitesPartage", "coutsConstruction", "coutsFonctionnement"
]

RELEVANT_CONTEXT = {
    'contexte_proj': ['besoin', 'objectifs'],
    'besoin': ['contexte_proj', 'objectifs', 'perimetre'],
    'objectifs': ['contexte_proj', 'besoin', 'perimetre'],
    'perimetre': ['contexte_proj', 'besoin', 'objectifs', 'horsPerimetre'],
    'horsPerimetre': ['perimetre', 'contexte_proj'],
    'descriptionSolution': ['besoin', 'objectifs', 'architecture'],
    'architecture': ['descriptionSolution', 'besoin', 'composantsDimensionnement'],
    'composantsDimensionnement': ['architecture', 'perimetre', 'plageService'],
    'conditionsHorsCrash': ['architecture', 'plageService', 'resilienceApplicative'],
    'conditionsCrashSite': ['conditionsHorsCrash', 'dicp', 'dima', 'pdma'],
    'resilienceApplicative': ['architecture', 'conditionsHorsCrash', 'praPlanDegrade'],
    'praPlanDegrade': ['resilienceApplicative', 'dicp', 'dima'],
    'sauvegardes': ['pdma', 'rgpd', 'administrationSupervision'],
    'administrationSupervision': ['plageService', 'tauxContingence'],
    'impactCO2': ['composantsDimensionnement', 'architecture'],
    'modalitesPartage': ['rgpd', 'psee', 'lienDocumentation'],
    'contraintes': ['contexte_proj', 'besoin', 'complexite'],
    'risques': ['contraintes', 'contexte_proj', 'objectifs'],
    'phases': ['perimetre', 'objectifs', 'complexite'],
    'livrables': ['phases', 'objectifs', 'perimetre'],
    'jalons': ['phases', 'livrables', 'objectifs'],
    'coutsConstruction': ['phases', 'perimetre', 'complexite'],
    'coutsFonctionnement': ['coutsConstruction', 'architecture', 'administrationSupervision']
}

TABLE_FIELDS = {
    'contraintes': ['type', 'description', 'criticité', 'mitigation'],
    'risques': ['risque', 'probabilité', 'impact', 'planActions', 'responsable'],
    'phases': ['phase', 'description', 'dateDébut', 'dateFin', 'équipes'],
    'livrables': ['nom', 'description', 'date', 'responsable'],
    'jalons': ['nom', 'type', 'date', 'critères'],
    'coutsConstruction': ['profil', 'nombre_jh', 'tjm', 'total', 'code'],
    'coutsFonctionnement': ['poste', 'quantite', 'coutUnitaire', 'total', 'code']
}

ADDITIONAL_FIELDS_POOL = [
    'contexte_proj', 'besoin', 'objectifs', 'perimetre', 'horsPerimetre',
    'descriptionSolution', 'architecture', 'contraintes', 'risques',
    'phases', 'composantsDimensionnement'
]

# ===== LOGGING DES GÉNÉRATIONS =====
def log_generation(field: str, prompt: str, raw_response: str, parsed_response: Any, 
                  error: Optional[str] = None):
    """Log complet de chaque génération dans un fichier JSONL"""
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "field": field,
        "prompt": prompt,
        "raw_response": raw_response,
        "parsed_response": parsed_response,
        "error": error
    }
    generation_logger.info(json.dumps(log_entry, ensure_ascii=False))

# ===== GESTIONNAIRE DU MODÈLE =====
class ModelManager:
    def __init__(self):
        self.model = None
        self.tokenizer = None
        self.device = settings.DEVICE
        self.is_loaded = False
        
    def load_model(self):
        """Charge le modèle FULL PRECISION avec SDPA"""
        if self.is_loaded:
            return
            
        logger.info(f"🔄 Chargement du modèle full precision depuis {settings.MODEL_PATH}")
        logger.info(f"   dtype: {settings.TORCH_DTYPE}")
        logger.info(f"   SDPA: {settings.USE_SDPA}")
        
        try:
            # Tokenizer
            self.tokenizer = AutoTokenizer.from_pretrained(settings.MODEL_PATH)
            self.tokenizer.pad_token = self.tokenizer.eos_token
            self.tokenizer.padding_side = "right"
            
            # Déterminer le dtype
            torch_dtype = getattr(torch, settings.TORCH_DTYPE)
            
            # Configuration du modèle
            model_kwargs = {
                "device_map": "auto",
                "trust_remote_code": True,
                "torch_dtype": torch_dtype,
            }
            
            # Activer SDPA si configuré
            if settings.USE_SDPA:
                model_kwargs["attn_implementation"] = "sdpa"
                logger.info("📦 SDPA (PyTorch natif) activé")
            
            # Chargement du modèle
            logger.info("📦 Chargement en full precision (bfloat16)...")
            self.model = AutoModelForCausalLM.from_pretrained(
                settings.MODEL_PATH,
                **model_kwargs
            )
            
            self.model.eval()
            self.model.config.use_cache = True  # Activer le cache KV pour l'inférence
            
            self.is_loaded = True
            
            # Stats mémoire
            if torch.cuda.is_available():
                allocated = torch.cuda.memory_allocated() / 1e9
                reserved = torch.cuda.memory_reserved() / 1e9
                logger.info(f"✅ Modèle chargé avec succès!")
                logger.info(f"   Mémoire GPU - Allouée: {allocated:.2f} Go, Réservée: {reserved:.2f} Go")
            
            logger.info(f"   Type: Mistral-7B Banking (Full Precision)")
            logger.info(f"   Device: {self.device}")
            logger.info(f"   dtype: {settings.TORCH_DTYPE}")
            logger.info(f"   SDPA: {settings.USE_SDPA}")
            
        except Exception as e:
            logger.error(f"❌ Erreur lors du chargement du modèle: {e}")
            raise
    
    def count_tokens(self, text: str) -> int:
        """Compte le nombre de tokens dans un texte"""
        tokens = self.tokenizer(text, return_tensors=None)
        return len(tokens['input_ids'])
    
    def generate(self, prompt: str) -> str:
        """Génère une réponse avec les paramètres optimisés"""
        if not self.is_loaded:
            raise RuntimeError("Modèle non chargé")
        
        self.model.eval()
        
        # Seed aléatoire pour variabilité
        seed = random.randint(0, 1000000)
        torch.manual_seed(seed)
        if torch.cuda.is_available():
            torch.cuda.manual_seed(seed)
        
        inputs = self.tokenizer(
            prompt, 
            return_tensors="pt", 
            truncation=True, 
            max_length=settings.MAX_PROMPT_LENGTH
        )
        inputs = {k: v.to(self.device) for k, v in inputs.items()}
        
        start_time = time.time()
        with torch.no_grad():
            outputs = self.model.generate(
                **inputs,
                max_new_tokens=settings.MAX_NEW_TOKENS,
                temperature=settings.TEMPERATURE,
                do_sample=True,
                top_p=settings.TOP_P,
                pad_token_id=self.tokenizer.pad_token_id,
                eos_token_id=self.tokenizer.eos_token_id,
                repetition_penalty=settings.REPETITION_PENALTY,
            )
        
        generation_time = time.time() - start_time
        
        generated_ids = outputs[0][inputs['input_ids'].shape[1]:]
        response = self.tokenizer.decode(generated_ids, skip_special_tokens=True)
        
        tokens_per_sec = len(generated_ids) / generation_time if generation_time > 0 else 0
        logger.info(f"⏱️ Génération en {generation_time:.2f}s ({tokens_per_sec:.1f} tokens/s, seed: {seed})")
        
        content = response
        if '<START>' in response and '<END>' in response:
            try:
                content = response.split('<START>')[1].split('<END>')[0].strip()
            except:
                logger.warning("Erreur d'extraction des tags START/END")
        
        return content

model_manager = ModelManager()

# ===== FONCTIONS DE FORMATAGE =====
def format_table_for_context(data: List[Dict], field_name: str) -> str:
    """Formate un tableau en format pipe-separated comme dans data_convert.py"""
    if not data:
        return ""
    
    columns = TABLE_FIELDS.get(field_name, [])
    rows = []
    
    for item in data:
        row = []
        for i, col in enumerate(columns):
            value = item.get(f'col{i}', '')
            row.append(str(value))
        rows.append("|".join(row))
    
    return "\n".join(rows)

def map_frontend_to_model_fields(frontend_context: Dict[str, Any]) -> Dict[str, Any]:
    """Mappe les noms de champs du frontend vers ceux du modèle"""
    mapping = {
        'sector': 'secteur',
        'project_type': 'typeProjet',
        'complexity': 'complexite',
        'contexte': 'contexte_proj',
        'besoin': 'besoin',
        'objectifs': 'objectifs',
        'lots': 'phases',
    }
    
    model_context = {}
    for key, value in frontend_context.items():
        model_key = mapping.get(key, key)
        model_context[model_key] = value
    
    return model_context

def build_context_identical_to_training(target_field: str, frontend_context: Dict[str, Any], 
                                       complexity: str = "medium") -> Tuple[str, List[str]]:
    """Construit le contexte EXACTEMENT comme dans data_convert.py"""
    context = map_frontend_to_model_fields(frontend_context)
    
    context_parts = []
    used_fields = []
    
    for field in ALWAYS_INCLUDED:
        if field in context and context[field]:
            if field in TABLE_FIELDS and isinstance(context[field], list):
                value = format_table_for_context(context[field], field)
                if value:
                    context_parts.append(f"{field}:\n{value}")
            else:
                context_parts.append(f"{field}: {context[field]}")
            used_fields.append(field)
    
    available_fields = [f for f in FIELD_GENERATION_ORDER if f != target_field]
    
    if complexity == "minimal":
        relevant = RELEVANT_CONTEXT.get(target_field, [])[:2]
        for field in relevant:
            if field in context and field not in used_fields:
                if field in TABLE_FIELDS and isinstance(context[field], list):
                    value = format_table_for_context(context[field], field)
                    if value:
                        context_parts.append(f"{field}:\n{value}")
                        used_fields.append(field)
                else:
                    context_parts.append(f"{field}: {context[field]}")
                    used_fields.append(field)
                    
    elif complexity == "medium" or complexity == "related":
        relevant = RELEVANT_CONTEXT.get(target_field, [])[:3]
        for field in relevant:
            if field in context and field not in used_fields:
                if field in TABLE_FIELDS and isinstance(context[field], list):
                    value = format_table_for_context(context[field], field)
                    if value:
                        context_parts.append(f"{field}:\n{value}")
                        used_fields.append(field)
                else:
                    context_parts.append(f"{field}: {context[field]}")
                    used_fields.append(field)
        
        random_fields = [f for f in available_fields if f not in relevant and f not in used_fields]
        if random_fields:
            random_field = random.choice(random_fields)
            if random_field in context:
                if random_field in TABLE_FIELDS and isinstance(context[random_field], list):
                    value = format_table_for_context(context[random_field], random_field)
                    if value:
                        context_parts.append(f"{random_field}:\n{value}")
                        used_fields.append(random_field)
                else:
                    context_parts.append(f"{random_field}: {context[random_field]}")
                    used_fields.append(random_field)
                    
    else:  # full
        relevant = RELEVANT_CONTEXT.get(target_field, [])[:2]
        for field in relevant:
            if field in context and field not in used_fields:
                if field in TABLE_FIELDS and isinstance(context[field], list):
                    value = format_table_for_context(context[field], field)
                    if value:
                        context_parts.append(f"{field}:\n{value}")
                        used_fields.append(field)
                else:
                    context_parts.append(f"{field}: {context[field]}")
                    used_fields.append(field)
        
        random_fields = [f for f in available_fields if f not in relevant and f not in used_fields]
        random.shuffle(random_fields)
        for field in random_fields[:random.randint(3, 5)]:
            if field in context:
                if field in TABLE_FIELDS and isinstance(context[field], list):
                    value = format_table_for_context(context[field], field)
                    if value:
                        context_parts.append(f"{field}:\n{value}")
                        used_fields.append(field)
                else:
                    context_parts.append(f"{field}: {context[field]}")
                    used_fields.append(field)
    
    context_str = "\n".join(context_parts)
    return context_str, used_fields

def build_prompt_identical_to_training(task: str, context_str: str) -> str:
    """Construit le prompt EXACT utilisé pendant l'entraînement"""
    return f"[INST] <TASK>{task}</TASK>\n<CONTEXT>\n{context_str}\n</CONTEXT> [/INST]"

# ===== PARSING DES RÉPONSES =====
def calculate_days_between_dates(date_start: str, date_end: str) -> int:
    """Calcule le nombre de jours entre deux dates"""
    try:
        if '/' in date_start:
            parts = date_start.split('/')
            if len(parts) == 3:
                start = datetime(int(parts[2]), int(parts[1]), int(parts[0]))
        else:
            start = datetime.strptime(date_start[:10], '%Y-%m-%d')
            
        if '/' in date_end:
            parts = date_end.split('/')
            if len(parts) == 3:
                end = datetime(int(parts[2]), int(parts[1]), int(parts[0]))
        else:
            end = datetime.strptime(date_end[:10], '%Y-%m-%d')
        
        days = 0
        current = start
        while current <= end:
            if current.weekday() < 5:
                days += 1
            current += timedelta(days=1)
        
        return max(1, days)
    except:
        return 20

def parse_phases(content: str) -> List[Dict]:
    """Parse les phases en calculant la durée entre les dates"""
    phases = []
    lines = content.strip().split('\n')
    
    logger.info(f"Parsing phases - lignes trouvées: {len(lines)}")
    
    for i, line in enumerate(lines[:8]):
        if not line.strip():
            continue
            
        parts = [p.strip() for p in line.split('|')]
        logger.info(f"Phase ligne {i}: {len(parts)} colonnes - {parts}")
        
        if len(parts) >= 5:
            phase_name = parts[0]
            description = parts[1]
            date_start = parts[2]
            date_end = parts[3]
            
            duration = calculate_days_between_dates(date_start, date_end)
            
            phase = {
                "col0": phase_name,
                "col1": description,
                "col2": str(duration),
                "col3": date_start,
                "col4": date_end,
                "col5": parts[4] if len(parts) > 5 else "Équipe projet"
            }
            
            logger.info(f"Phase parsée: {phase}")
            phases.append(phase)
        else:
            phase = {
                "col0": parts[0] if len(parts) > 0 else f"Phase {i+1}",
                "col1": parts[1] if len(parts) > 1 else "Description à définir",
                "col2": "20",
                "col3": "",
                "col4": "",
                "col5": "Équipe projet"
            }
            phases.append(phase)
        
        if len(phases) >= 8:
            break
    
    return phases

def parse_jalons(content: str) -> List[Dict]:
    """Parse les jalons"""
    jalons = []
    lines = content.strip().split('\n')
    
    logger.info(f"Parsing jalons - lignes trouvées: {len(lines)}")
    
    for i, line in enumerate(lines[:5]):
        if not line.strip():
            continue
            
        parts = [p.strip() for p in line.split('|')]
        logger.info(f"Jalon ligne {i}: {len(parts)} colonnes - {parts}")
        
        if len(parts) >= 3:
            jalon = {
                "col0": parts[0],
                "col1": parts[1] if len(parts) > 1 else "Validation",
                "col2": extract_date(parts[2]) if len(parts) > 2 else "",
                "col3": parts[3] if len(parts) > 3 else "Critères à définir"
            }
            
            logger.info(f"Jalon parsé: {jalon}")
            jalons.append(jalon)
        else:
            jalon = {
                "col0": parts[0] if len(parts) > 0 else f"Jalon {i+1}",
                "col1": "Validation",
                "col2": "",
                "col3": "Critères à définir"
            }
            jalons.append(jalon)
        
        if len(jalons) >= 5:
            break
    
    return jalons

def parse_table_content(content: str, table_type: str) -> List[Dict]:
    """Parse le contenu d'un tableau selon son type"""
    
    if table_type == "contraintes":
        return parse_contraintes(content)
    elif table_type == "risques":
        return parse_risques(content)
    elif table_type == "phases":
        return parse_phases(content)
    elif table_type == "livrables":
        return parse_livrables(content)
    elif table_type == "jalons":
        return parse_jalons(content)
    elif table_type == "coutsConstruction":
        return parse_couts_construction(content)
    elif table_type == "coutsFonctionnement":
        return parse_couts_fonctionnement(content)
    
    return []

def parse_contraintes(content: str) -> List[Dict]:
    """Parse les contraintes"""
    contraintes = []
    lines = content.strip().split('\n')
    
    for line in lines[:5]:
        if not line.strip():
            continue
            
        parts = [p.strip() for p in line.split('|')]
        
        if len(parts) >= 3:
            contrainte = {
                "col0": parts[0],
                "col1": parts[1],
                "col2": map_criticality(parts[2]),
                "col3": parts[3] if len(parts) > 3 else "À définir"
            }
            contraintes.append(contrainte)
    
    return contraintes

def parse_risques(content: str) -> List[Dict]:
    """Parse les risques"""
    risques = []
    lines = content.strip().split('\n')
    
    for line in lines[:6]:
        if not line.strip():
            continue
            
        parts = [p.strip() for p in line.split('|')]
        
        if len(parts) >= 4:
            risque = {
                "col0": parts[0],
                "col1": map_probability(parts[1]),
                "col2": map_impact(parts[2]),
                "col3": parts[3],
                "col4": parts[4] if len(parts) > 4 else "Chef de projet"
            }
            risques.append(risque)
    
    return risques

def parse_livrables(content: str) -> List[Dict]:
    """Parse les livrables"""
    livrables = []
    lines = content.strip().split('\n')
    
    for line in lines[:6]:
        if not line.strip():
            continue
            
        parts = [p.strip() for p in line.split('|')]
        
        if len(parts) >= 3:
            livrable = {
                "col0": parts[0],
                "col1": parts[1] if len(parts) > 1 else "Document",
                "col2": extract_date(parts[2]) if len(parts) > 2 else "",
                "col3": parts[3] if len(parts) > 3 else "NOTRE ENTREPRISE"
            }
            livrables.append(livrable)
    
    return livrables

def parse_couts_construction(content: str) -> List[Dict]:
    """Parse les coûts de construction"""
    couts = []
    lines = content.strip().split('\n')
    
    for line in lines[:7]:
        if not line.strip():
            continue
            
        parts = [p.strip() for p in line.split('|')]
        
        if len(parts) >= 3:
            nb_jh = extract_number(parts[1])
            tjm = extract_number(parts[2])
            total = nb_jh * tjm
            
            cout = {
                "col0": parts[0],
                "col1": str(nb_jh),
                "col2": str(tjm),
                "col3": str(total),
                "col4": parts[4] if len(parts) > 4 else ""
            }
            couts.append(cout)
    
    return couts

def parse_couts_fonctionnement(content: str) -> List[Dict]:
    """Parse les coûts de fonctionnement"""
    couts = []
    lines = content.strip().split('\n')
    
    for line in lines[:7]:
        if not line.strip():
            continue
            
        parts = [p.strip() for p in line.split('|')]
        
        if len(parts) >= 3:
            quantite = extract_number(parts[1])
            prix_mensuel = extract_number(parts[2]) / 12
            total_annuel = quantite * prix_mensuel * 12
            
            cout = {
                "col0": parts[0],
                "col1": parts[1],
                "col2": prix_mensuel,
                "col3": parts[3],
                "col4": parts[4] if len(parts) > 4 else ""
            }
            couts.append(cout)
    
    return couts

# ===== FONCTIONS UTILITAIRES =====
def map_criticality(text: str) -> str:
    text = text.lower().strip()
    if any(word in text for word in ["faible", "bas", "low"]):
        return "Faible"
    elif any(word in text for word in ["élevé", "elevé", "haut", "high", "critique"]):
        return "Élevé"
    return "Moyen"

def map_probability(text: str) -> str:
    text = text.lower().strip()
    if any(word in text for word in ["faible", "bas", "low"]):
        return "Faible"
    elif any(word in text for word in ["élevé", "elevé", "haut", "high"]):
        return "Élevée"
    return "Moyenne"

def map_impact(text: str) -> str:
    text = text.lower().strip()
    if any(word in text for word in ["faible", "bas", "low"]):
        return "Faible"
    elif any(word in text for word in ["élevé", "elevé", "haut", "high", "majeur"]):
        return "Élevé"
    return "Moyen"

def extract_date(text: str) -> str:
    """Extrait une date ou retourne vide"""
    if not text or len(text.strip()) < 6:
        return ""
    
    match = re.search(r'(\d{4})-(\d{2})-(\d{2})', text)
    if match:
        return match.group(0)
    
    match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', text)
    if match:
        day, month, year = match.groups()
        return f"{day.zfill(2)}/{month.zfill(2)}/{year}"
    
    return ""

def extract_number(text: str) -> int:
    """Extrait un nombre"""
    text = str(text).replace(' ', '').replace(',', '').replace('€', '').replace('k', '000')
    match = re.search(r'(\d+)', text)
    if match:
        return int(match.group(1))
    return 0

# ===== APPLICATION FASTAPI =====
@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("🚀 Démarrage de l'API SCRIBE AI - FULL PRECISION")
    logger.info(f"📁 Modèle: {settings.MODEL_PATH}")
    logger.info(f"🔧 dtype: {settings.TORCH_DTYPE}, SDPA: {settings.USE_SDPA}")
    logger.info(f"📝 Les logs de génération sont dans: generation_logs.jsonl")
    try:
        model_manager.load_model()
    except Exception as e:
        logger.error(f"❌ Erreur au chargement: {e}")
        logger.info("⚠️  API démarre sans modèle")
    
    yield
    
    logger.info("🛑 Arrêt de l'API...")
    if torch.cuda.is_available():
        torch.cuda.empty_cache()

app = FastAPI(
    title="SCRIBE AI API - Full Precision",
    version="5.0.0",
    lifespan=lifespan
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.middleware("http")
async def disable_cache(request: Request, call_next):
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate, private"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response

# ===== ENDPOINTS =====
@app.get("/")
async def root():
    return {
        "name": "SCRIBE AI API",
        "version": "5.0.0",
        "model": "Mistral-7B Banking (Full Precision)",
        "model_type": "full_precision",
        "model_loaded": model_manager.is_loaded,
        "model_path": settings.MODEL_PATH,
        "config": {
            "dtype": settings.TORCH_DTYPE,
            "sdpa": settings.USE_SDPA,
            "max_prompt_length": settings.MAX_PROMPT_LENGTH,
            "max_new_tokens": settings.MAX_NEW_TOKENS
        },
        "logs": "Check generation_logs.jsonl"
    }

@app.get("/health")
async def health_check():
    gpu_info = {}
    if torch.cuda.is_available():
        gpu_info = {
            "name": torch.cuda.get_device_name(0),
            "memory_allocated_gb": round(torch.cuda.memory_allocated() / 1e9, 2),
            "memory_reserved_gb": round(torch.cuda.memory_reserved() / 1e9, 2),
            "memory_total_gb": round(torch.cuda.get_device_properties(0).total_memory / 1e9, 2)
        }
    
    return {
        "status": "ok",
        "model_loaded": model_manager.is_loaded,
        "device": settings.DEVICE,
        "model_type": "full_precision",
        "dtype": settings.TORCH_DTYPE,
        "sdpa": settings.USE_SDPA,
        "gpu": gpu_info,
        "timestamp": datetime.now().isoformat()
    }

@app.post("/generate")
async def generate_field(request: GenerateRequest):
    """Génère du contenu pour un champ avec logging complet"""
    prompt = None
    raw_response = None
    parsed_response = None
    error_msg = None
    
    try:
        if not model_manager.is_loaded:
            raise HTTPException(status_code=503, detail="Modèle non chargé")
        
        field = request.field
        task = "phases" if field == "lots" else field
        task = "contexte_proj" if task == "contexte" else task
        
        logger.info(f"🎯 Génération pour {field} -> tâche {task}")
        
        complexity_map = {
            "minimal": "minimal",
            "related": "medium",
            "smart": "medium",
            "full": "full",
            "debug": "full"
        }
        complexity = complexity_map.get(request.strategy, "medium")
        
        context_str, used_fields = build_context_identical_to_training(
            task, 
            request.context, 
            complexity
        )
        
        prompt = build_prompt_identical_to_training(task, context_str)
        
        token_count = model_manager.count_tokens(prompt)
        logger.info(f"📊 Prompt: {token_count} tokens")
        
        raw_response = model_manager.generate(prompt)
        
        table_fields = ["contraintes", "risques", "lots", "livrables", "jalons", 
                       "coutsConstruction", "coutsFonctionnement"]
        
        if field in table_fields:
            parse_type = "phases" if field == "lots" else field
            parsed_response = parse_table_content(raw_response, parse_type)
        else:
            parsed_response = raw_response
        
        log_generation(field, prompt, raw_response, parsed_response)
        
        return {
            "success": True,
            "field": field,
            "value": parsed_response,
            "validation": {
                "valid": True,
                "warnings": []
            },
            "metadata": {
                "task": task,
                "strategy": request.strategy,
                "complexity": complexity,
                "context_fields": used_fields,
                "token_count": token_count,
                "model_type": "full_precision",
                "dtype": settings.TORCH_DTYPE,
                "timestamp": datetime.now().isoformat()
            }
        }
        
    except HTTPException:
        raise
    except Exception as e:
        error_msg = str(e)
        logger.error(f"❌ Erreur génération: {e}", exc_info=True)
        
        if prompt:
            log_generation(field, prompt, raw_response or "", parsed_response or "", error_msg)
        
        raise HTTPException(status_code=500, detail=error_msg)

@app.post("/generate_multiple")
async def generate_multiple_fields(request: GenerateMultipleRequest):
    """Génère du contenu pour plusieurs champs"""
    results = {}
    errors = {}
    
    for field in request.fields[:5]:
        try:
            req = GenerateRequest(
                field=field,
                context=request.context,
                strategy=request.strategy
            )
            response = await generate_field(req)
            results[field] = response["value"]
        except Exception as e:
            errors[field] = str(e)
    
    return {
        "success": len(errors) == 0,
        "results": results,
        "errors": errors if errors else None,
        "timestamp": datetime.now().isoformat()
    }

@app.post("/validate")
async def validate_field(request: ValidateRequest):
    """Valide un champ"""
    warnings = []
    
    if request.field == "annee":
        try:
            year = int(request.value)
            if year < 2024 or year > 2030:
                warnings.append("L'année semble inhabituellement éloignée")
        except:
            warnings.append("Format d'année invalide")
    
    return {
        "success": True,
        "validation": {
            "valid": len(warnings) == 0,
            "errors": [],
            "warnings": warnings
        }
    }

# ===== FEEDBACK ENDPOINTS =====
feedback_store = {
    "corrections": [],
    "ratings": [],
    "stats": {}
}

@app.post("/feedback/correction")
async def submit_correction(request: FeedbackCorrectionRequest):
    feedback_store["corrections"].append({
        "field": request.field,
        "original": request.original,
        "corrected": request.corrected,
        "reason": request.reason,
        "timestamp": datetime.now().isoformat()
    })
    
    logger.info(f"📝 Correction enregistrée pour {request.field}")
    
    return {
        "success": True,
        "message": "Correction enregistrée avec succès",
        "insights": {
            "recommendations": [
                f"Améliorer la génération pour le champ {request.field}",
                "Prendre en compte le feedback utilisateur"
            ],
            "field": request.field,
            "improvement_noted": True
        }
    }

@app.post("/feedback/rating")
async def submit_rating(request: FeedbackRatingRequest):
    feedback_store["ratings"].append({
        "field": request.field,
        "rating": request.rating,
        "comment": request.comment,
        "timestamp": datetime.now().isoformat()
    })
    
    logger.info(f"⭐ Note {request.rating}/5 pour {request.field}")
    
    message = "Merci pour votre retour !"
    if request.rating < 3:
        message = "Merci pour votre retour, nous allons améliorer ce point."
    
    return {
        "success": True,
        "message": message,
        "rating": request.rating,
        "field": request.field
    }

@app.get("/feedback/insights/{field}")
async def get_field_insights(field: str):
    field_ratings = [r["rating"] for r in feedback_store["ratings"] if r["field"] == field]
    field_corrections = [c for c in feedback_store["corrections"] if c["field"] == field]
    
    avg_rating = sum(field_ratings) / len(field_ratings) if field_ratings else 0
    
    return {
        "success": True,
        "field": field,
        "insights": {
            "average_rating": round(avg_rating, 1),
            "total_corrections": len(field_corrections),
            "total_ratings": len(field_ratings),
            "common_issues": [],
            "recommendations": [
                "Continuer à collecter des feedbacks" if len(field_ratings) < 5
                else "Analyser les patterns de correction"
            ]
        }
    }

@app.get("/feedback/summary")
async def get_feedback_summary():
    total_ratings = len(feedback_store["ratings"])
    total_corrections = len(feedback_store["corrections"])
    
    all_ratings = [r["rating"] for r in feedback_store["ratings"]]
    global_avg = sum(all_ratings) / len(all_ratings) if all_ratings else 0
    
    return {
        "success": True,
        "summary": {
            "total_corrections": total_corrections,
            "total_ratings": total_ratings,
            "total_feedback": total_corrections + total_ratings,
            "global_quality_score": round(global_avg, 1),
            "quality_status": "good" if global_avg >= 3.5 else "needs_improvement"
        },
        "message": "Qualité globale satisfaisante" if global_avg >= 3.5 
                  else "Des améliorations sont nécessaires",
        "last_updated": datetime.now().isoformat()
    }

@app.get("/model/info")
async def get_model_info():
    """Endpoint pour obtenir des informations sur le modèle"""
    gpu_info = {}
    if torch.cuda.is_available():
        gpu_info = {
            "name": torch.cuda.get_device_name(0),
            "memory_total_gb": round(torch.cuda.get_device_properties(0).total_memory / 1e9, 2),
            "memory_allocated_gb": round(torch.cuda.memory_allocated() / 1e9, 2),
            "memory_peak_gb": round(torch.cuda.max_memory_allocated() / 1e9, 2)
        }
    
    return {
        "model_type": "Mistral-7B Banking Full Precision",
        "path": settings.MODEL_PATH,
        "loaded": model_manager.is_loaded,
        "device": settings.DEVICE,
        "config": {
            "dtype": settings.TORCH_DTYPE,
            "sdpa": settings.USE_SDPA,
            "max_prompt_length": settings.MAX_PROMPT_LENGTH,
            "max_new_tokens": settings.MAX_NEW_TOKENS,
            "temperature": settings.TEMPERATURE,
            "top_p": settings.TOP_P,
            "repetition_penalty": settings.REPETITION_PENALTY
        },
        "gpu": gpu_info,
        "description": "Modèle Mistral-7B fine-tuné en full precision (bfloat16) avec SDPA",
        "capabilities": {
            "fields": list(FIELD_GENERATION_ORDER),
            "table_fields": list(TABLE_FIELDS.keys())
        }
    }

# ===== MAIN =====
if __name__ == "__main__":
    import uvicorn
    
    logger.info(f"🚀 Démarrage sur {settings.HOST}:{settings.PORT}")
    logger.info(f"📁 Model path: {settings.MODEL_PATH}")
    logger.info(f"🔧 dtype: {settings.TORCH_DTYPE}, SDPA: {settings.USE_SDPA}")
    logger.info(f"📝 Logs de génération dans: generation_logs.jsonl")
    
    uvicorn.run(
        app,
        host=settings.HOST,
        port=settings.PORT,
        log_level="info"
    )
