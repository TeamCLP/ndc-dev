#!/usr/bin/env python3
"""
Script d'entraînement optimisé pour Mistral 7B Instruct v0.3
VERSION V2 - DOCUMENTS MARKDOWN COMPLETS (Expression de besoin -> Devis)

Optimisations pour GPU H100 (80 Go VRAM):
- Modèle en bfloat16
- Pas de quantization
- Contexte très long pour documents complets
- LoRA rank élevé
- SDPA natif PyTorch
"""

import os
import json
import torch
import random
import logging
import numpy as np
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any
from dataclasses import dataclass, field
from tqdm import tqdm
from collections import defaultdict
import argparse

import transformers
from transformers import (
    AutoModelForCausalLM,
    AutoTokenizer,
    TrainingArguments,
)
from peft import (
    LoraConfig,
    get_peft_model,
    TaskType,
    PeftModel,
)
from datasets import Dataset, DatasetDict
from torch.utils.tensorboard import SummaryWriter

# Configuration du logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('training_v2.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


@dataclass
class ModelConfig:
    """Configuration du modèle - optimisée pour documents longs sur H100"""
    model_name: str = "mistralai/Mistral-7B-Instruct-v0.3"

    # LoRA - valeurs élevées pour capacité maximale
    lora_r: int = 128
    lora_alpha: int = 256
    lora_dropout: float = 0.1
    lora_target_modules: List[str] = field(default_factory=lambda: [
        "q_proj", "k_proj", "v_proj", "o_proj",
        "gate_proj", "up_proj", "down_proj"
    ])

    # Limites de tokens - TRÈS augmentées pour documents complets
    # Expression de besoin markdown: ~2000-4000 tokens
    # Devis complet markdown: ~3000-6000 tokens
    max_prompt_length: int = 6144  # ~6k tokens pour l'expression de besoin
    max_response_length: int = 8192  # ~8k tokens pour le devis
    max_length: int = 14336  # Total ~14k tokens

    # Full precision
    use_quantization: bool = False
    torch_dtype: str = "bfloat16"

    # SDPA pour performance
    use_sdpa: bool = True

    # Paths
    output_dir: str = "/home/quentin/mistral-devis"
    tensorboard_dir: str = "/home/quentin/runs/mistral-devis"

    seed: int = 42


@dataclass
class TrainingConfig:
    """Configuration d'entraînement - optimisée pour H100 avec longs documents"""
    # Batch sizes réduits à cause des séquences longues
    per_device_train_batch_size: int = 2  # Réduit car séquences très longues
    per_device_eval_batch_size: int = 2
    gradient_accumulation_steps: int = 16  # Augmenté pour compenser petit batch
    gradient_checkpointing: bool = True  # Important pour les longues séquences

    # Learning rate
    learning_rate: float = 2e-5  # Légèrement réduit pour stabilité
    num_train_epochs: int = 3
    warmup_ratio: float = 0.05
    weight_decay: float = 0.01

    # bfloat16
    bf16: bool = True
    fp16: bool = False

    # Optimiseur
    optim: str = "adamw_torch"

    # Logging
    logging_steps: int = 1
    eval_steps: int = 50
    save_steps: int = 50
    save_total_limit: int = 20

    # Fichiers de données
    train_file: str = "dataset/train_dataset.jsonl"
    val_file: str = "dataset/val_dataset.jsonl"

    # Reprise d'entraînement
    resume_from_checkpoint: Optional[str] = None

    # Dataloader
    dataloader_num_workers: int = 4
    dataloader_pin_memory: bool = True


class RobustDataCollator:
    """Data collator robuste pour séquences longues"""
    def __init__(self, tokenizer, pad_to_multiple_of=8):
        self.tokenizer = tokenizer
        self.pad_to_multiple_of = pad_to_multiple_of

    def __call__(self, features):
        input_ids = [f["input_ids"] for f in features]
        attention_mask = [f["attention_mask"] for f in features]
        labels = [f["labels"] for f in features]

        max_length = max(len(ids) for ids in input_ids)

        if self.pad_to_multiple_of:
            max_length = ((max_length + self.pad_to_multiple_of - 1)
                         // self.pad_to_multiple_of) * self.pad_to_multiple_of

        batch_input_ids = []
        batch_attention_mask = []
        batch_labels = []

        for i in range(len(input_ids)):
            padding_length = max_length - len(input_ids[i])

            batch_input_ids.append(
                input_ids[i] + [self.tokenizer.pad_token_id] * padding_length
            )
            batch_attention_mask.append(
                attention_mask[i] + [0] * padding_length
            )
            batch_labels.append(
                labels[i] + [-100] * padding_length
            )

        return {
            "input_ids": torch.tensor(batch_input_ids, dtype=torch.long),
            "attention_mask": torch.tensor(batch_attention_mask, dtype=torch.long),
            "labels": torch.tensor(batch_labels, dtype=torch.long)
        }


class DocumentDatasetProcessor:
    """Processeur de dataset pour documents markdown complets"""

    def __init__(self, tokenizer, config: ModelConfig):
        self.tokenizer = tokenizer
        self.config = config
        self.stats = defaultdict(int)
        self.debug_examples = 3

    def validate_entry(self, entry: Dict) -> bool:
        """Valide une entrée - version simplifiée sans balises START/END"""
        if 'messages' not in entry or len(entry['messages']) != 2:
            return False

        if (entry['messages'][0].get('role') != 'user' or
            entry['messages'][1].get('role') != 'assistant'):
            return False

        # Vérifier que les contenus ne sont pas vides
        user_content = entry['messages'][0].get('content', '').strip()
        assistant_content = entry['messages'][1].get('content', '').strip()

        if not user_content or not assistant_content:
            return False

        return True

    def load_dataset(self, file_path: str) -> List[Dict]:
        """Charge le dataset"""
        data = []

        with open(file_path, 'r', encoding='utf-8') as f:
            for i, line in enumerate(f):
                try:
                    entry = json.loads(line)
                    if self.validate_entry(entry):
                        data.append(entry)
                        self.stats['valid'] += 1
                    else:
                        self.stats['invalid'] += 1
                        if self.stats['invalid'] <= 5:
                            logger.warning(f"Entrée invalide ligne {i+1}")
                except Exception as e:
                    logger.error(f"Erreur ligne {i+1}: {e}")
                    self.stats['error'] += 1

        logger.info(f"Dataset {file_path} chargé: {self.stats['valid']} valides, "
                   f"{self.stats['invalid']} invalides, {self.stats['error']} erreurs")

        return data

    def preprocess_example(self, messages: List[Dict], example_idx: int = -1) -> Dict:
        """Preprocessing pour documents markdown complets"""
        user_content = messages[0]['content']  # Expression de besoin
        assistant_content = messages[1]['content']  # Devis complet

        # Format Mistral Instruct
        prompt = f"[INST] {user_content} [/INST]"
        response = assistant_content

        # Tokenizer le prompt (expression de besoin)
        prompt_encoding = self.tokenizer(
            prompt,
            max_length=self.config.max_prompt_length,
            truncation=True,
            add_special_tokens=False,
            return_tensors=None
        )

        # Tokenizer la réponse (devis)
        response_encoding = self.tokenizer(
            response,
            max_length=self.config.max_response_length,
            truncation=True,
            add_special_tokens=False,
            return_tensors=None
        )

        # Tokens spéciaux
        bos_token = [self.tokenizer.bos_token_id] if self.tokenizer.bos_token_id is not None else []
        eos_token = [self.tokenizer.eos_token_id] if self.tokenizer.eos_token_id is not None else []

        input_ids = bos_token + prompt_encoding['input_ids'] + response_encoding['input_ids'] + eos_token

        # Tronquer si nécessaire
        if len(input_ids) > self.config.max_length:
            input_ids = input_ids[:self.config.max_length]
            self.stats['truncated'] += 1

        attention_mask = [1] * len(input_ids)

        # Labels : masquer BOS + prompt, apprendre uniquement la réponse
        prompt_length = len(bos_token) + len(prompt_encoding['input_ids'])
        labels = [-100] * prompt_length + input_ids[prompt_length:]

        if len(labels) != len(input_ids):
            labels = labels[:len(input_ids)]

        # Debug sur les premiers exemples
        if example_idx < self.debug_examples and example_idx >= 0:
            num_learn = sum(1 for l in labels if l != -100)
            ratio = num_learn / len(labels) * 100 if len(labels) > 0 else 0

            logger.info(f"\nExemple {example_idx}:")
            logger.info(f"  Longueur totale: {len(input_ids)} tokens")
            logger.info(f"  Prompt (expression de besoin): {len(prompt_encoding['input_ids'])} tokens")
            logger.info(f"  Réponse (devis): {len(response_encoding['input_ids'])} tokens")
            logger.info(f"  Tokens à apprendre: {num_learn} ({ratio:.1f}%)")

            if num_learn > 0:
                first_learn_idx = next(i for i, l in enumerate(labels) if l != -100)
                learn_preview = self.tokenizer.decode(
                    input_ids[first_learn_idx:min(first_learn_idx+50, len(input_ids))],
                    skip_special_tokens=False
                )
                logger.info(f"  Début du devis: {learn_preview[:100]}...")

        return {
            'input_ids': input_ids,
            'attention_mask': attention_mask,
            'labels': labels
        }

    def preprocess_function(self, examples: Dict) -> Dict:
        """Fonction de preprocessing pour map"""
        results = {
            'input_ids': [],
            'attention_mask': [],
            'labels': []
        }

        for idx, messages in enumerate(examples['messages']):
            example_idx = self.stats['processed'] + idx
            processed = self.preprocess_example(messages, example_idx)

            results['input_ids'].append(processed['input_ids'])
            results['attention_mask'].append(processed['attention_mask'])
            results['labels'].append(processed['labels'])

        self.stats['processed'] += len(examples['messages'])

        return results

    def prepare_datasets_from_files(self, train_file: str, val_file: str) -> Tuple[DatasetDict, List[List[Dict]]]:
        """Prépare les datasets depuis des fichiers séparés"""
        self.stats = defaultdict(int)
        train_data = self.load_dataset(train_file)

        self.stats = defaultdict(int)
        val_data = self.load_dataset(val_file)

        if not train_data:
            raise ValueError(f"Aucune donnée valide trouvée dans {train_file}!")
        if not val_data:
            raise ValueError(f"Aucune donnée valide trouvée dans {val_file}!")

        logger.info(f"Chargé: {len(train_data)} train, {len(val_data)} validation")

        val_messages_original = [item['messages'] for item in val_data]

        train_messages = [item['messages'] for item in train_data]
        val_messages = [item['messages'] for item in val_data]

        train_dataset = Dataset.from_dict({'messages': train_messages})
        val_dataset = Dataset.from_dict({'messages': val_messages})

        self.stats['processed'] = 0
        self.stats['truncated'] = 0

        logger.info("Préprocessing du dataset d'entraînement...")
        train_dataset = train_dataset.map(
            self.preprocess_function,
            batched=True,
            batch_size=50,  # Plus petit batch pour documents longs
            remove_columns=['messages'],
            desc="Préprocessing train",
            num_proc=4
        )

        truncated_train = self.stats['truncated']
        self.stats['processed'] = 0
        self.stats['truncated'] = 0

        logger.info("Préprocessing du dataset de validation...")
        val_dataset = val_dataset.map(
            self.preprocess_function,
            batched=True,
            batch_size=50,
            remove_columns=['messages'],
            desc="Préprocessing validation",
            num_proc=4
        )

        truncated_val = self.stats['truncated']

        if truncated_train > 0 or truncated_val > 0:
            logger.warning(f"Séquences tronquées: {truncated_train} train, {truncated_val} val")

        datasets = DatasetDict({
            'train': train_dataset,
            'validation': val_dataset
        })

        # Statistiques de longueur
        train_lengths = [len(ex['input_ids']) for ex in train_dataset]
        logger.info(f"Longueurs train - Min: {min(train_lengths)}, Max: {max(train_lengths)}, "
                   f"Moy: {np.mean(train_lengths):.1f}, Médiane: {np.median(train_lengths):.1f}")

        return datasets, val_messages_original


class ValidationTrainer(transformers.Trainer):
    """Trainer avec génération sur validation"""

    def __init__(self, *args, tb_writer=None, val_examples=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.tb_writer = tb_writer
        self.val_examples = val_examples or []
        self.generation_history = []
        self.logged_initial_loss = False

    def compute_loss(self, model, inputs, return_outputs=False, num_items_in_batch=None):
        """Compute loss avec monitoring"""
        outputs = model(**inputs)
        loss = outputs.loss

        if not self.logged_initial_loss and self.state.global_step == 0:
            logger.info(f"\n Loss initiale: {loss.item():.4f}")
            if loss.item() > 15:
                logger.info("Loss initiale élevée - normal pour documents longs")
            elif loss.item() < 3:
                logger.warning("Loss initiale très basse - vérifiez le dataset")
            else:
                logger.info("Loss initiale dans la normale")
            self.logged_initial_loss = True

        if self.tb_writer:
            self.tb_writer.add_scalar("train/loss", loss.item(), self.state.global_step)
            if self.state.global_step % 10 == 0:
                logger.info(f"Step {self.state.global_step}: Loss = {loss.item():.4f}")

        return (loss, outputs) if return_outputs else loss

    def evaluate(self, *args, **kwargs):
        """Évaluation avec génération et sauvegarde"""
        output = super().evaluate(*args, **kwargs)

        if self.state.global_step > 0:
            generation_results = self.generate_validation_examples(num_examples=5)
            self.save_generation_results(generation_results)

        return output

    def save_generation_results(self, results):
        """Sauvegarde les résultats de génération"""
        if not results:
            return

        generation_dir = os.path.join(self.args.output_dir, "generation_results")
        os.makedirs(generation_dir, exist_ok=True)

        filename = f"generation_step_{self.state.global_step}.json"
        filepath = os.path.join(generation_dir, filename)

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        logger.info(f"Résultats de génération sauvegardés: {filepath}")

    def generate_validation_examples(self, num_examples: int = 3):
        """Génère des exemples de devis depuis la validation"""
        if not self.val_examples:
            logger.warning("Pas d'exemples de validation")
            return []

        logger.info(f"\n=== Génération VALIDATION (Step {self.state.global_step}) ===")

        model = self.model
        tokenizer = self.tokenizer

        selected_examples = random.sample(
            self.val_examples,
            min(num_examples, len(self.val_examples))
        )

        model.eval()
        results = []

        with torch.no_grad():
            for i, messages in enumerate(selected_examples):
                user_content = messages[0]['content']  # Expression de besoin
                expected = messages[1]['content']  # Devis attendu

                prompt = f"[INST] {user_content} [/INST]"
                inputs = tokenizer(
                    prompt,
                    return_tensors="pt",
                    truncation=True,
                    max_length=self.args.model_config.max_prompt_length
                )
                inputs = {k: v.to(model.device) for k, v in inputs.items()}

                start_time = time.time()
                outputs = model.generate(
                    **inputs,
                    max_new_tokens=min(2048, self.args.model_config.max_response_length),  # Limité pour la validation
                    temperature=0.7,
                    do_sample=True,
                    top_p=0.9,
                    pad_token_id=tokenizer.pad_token_id,
                    eos_token_id=tokenizer.eos_token_id,
                    repetition_penalty=1.1,
                )
                generation_time = time.time() - start_time

                generated_ids = outputs[0][inputs['input_ids'].shape[1]:]
                generated = tokenizer.decode(generated_ids, skip_special_tokens=True)

                # Aperçu du résultat
                preview = generated[:200].replace('\n', ' ')
                logger.info(f"\nVal {i+1}:")
                logger.info(f"  Expression de besoin: {user_content[:100]}...")
                logger.info(f"  Devis généré: {preview}...")
                logger.info(f"  Tokens générés: {len(generated_ids)}, Temps: {generation_time:.2f}s")

                result = {
                    'step': self.state.global_step,
                    'index': i,
                    'timestamp': datetime.now().isoformat(),
                    'expression_besoin': user_content,
                    'devis_attendu': expected,
                    'devis_genere': generated,
                    'generation_time_seconds': round(generation_time, 3),
                    'tokens_generated': len(generated_ids),
                }

                results.append(result)

        if results and self.tb_writer:
            avg_gen_time = sum(r['generation_time_seconds'] for r in results) / len(results)
            avg_tokens = sum(r['tokens_generated'] for r in results) / len(results)
            self.tb_writer.add_scalar("val_gen/avg_generation_time", avg_gen_time, self.state.global_step)
            self.tb_writer.add_scalar("val_gen/avg_tokens_generated", avg_tokens, self.state.global_step)

        self.generation_history.extend(results)

        return results


def check_initial_loss(model, dataset, data_collator, tokenizer):
    """Check informatif de la loss initiale"""
    logger.info("\n=== CHECK DE LA LOSS (INFORMATIF) ===")

    try:
        indices = list(range(min(4, len(dataset))))  # Moins d'exemples car plus longs
        batch_examples = [dataset[i] for i in indices]

        batch = data_collator(batch_examples)
        batch = {k: v.to(model.device) for k, v in batch.items()}

        total_tokens = 0
        total_learn = 0

        for i in range(batch['labels'].shape[0]):
            labels = batch['labels'][i]
            num_tokens = len(labels)
            num_learn = (labels != -100).sum().item()
            total_tokens += num_tokens
            total_learn += num_learn

        avg_learn_ratio = total_learn / total_tokens * 100 if total_tokens > 0 else 0
        logger.info(f"Ratio moyen d'apprentissage: {avg_learn_ratio:.1f}%")

        model.eval()
        with torch.no_grad():
            outputs = model(**batch)
            loss = outputs.loss.item()

        logger.info(f"Loss initiale: {loss:.4f}")

    except Exception as e:
        logger.warning(f"Check impossible: {str(e)}")


def setup_model_and_tokenizer(config: ModelConfig, resume_from_checkpoint: Optional[str] = None):
    """Setup modèle et tokenizer"""
    logger.info("Chargement du modèle...")
    logger.info(f"  dtype: {config.torch_dtype}")
    logger.info(f"  SDPA: {config.use_sdpa}")

    torch_dtype = getattr(torch, config.torch_dtype)

    model_kwargs = {
        "device_map": "auto",
        "trust_remote_code": True,
        "torch_dtype": torch_dtype,
    }

    if config.use_sdpa:
        model_kwargs["attn_implementation"] = "sdpa"
        logger.info("SDPA activé")

    # Tokenizer
    tokenizer = AutoTokenizer.from_pretrained(config.model_name)
    tokenizer.pad_token = tokenizer.eos_token
    tokenizer.padding_side = "right"

    if resume_from_checkpoint:
        logger.info(f"Chargement depuis checkpoint: {resume_from_checkpoint}")

        base_model = AutoModelForCausalLM.from_pretrained(
            config.model_name,
            **model_kwargs
        )

        model = PeftModel.from_pretrained(
            base_model,
            resume_from_checkpoint,
            is_trainable=True
        )

        model.enable_input_require_grads()
        logger.info("Modèle chargé depuis checkpoint")

    else:
        logger.info("Création d'un nouveau modèle")

        model = AutoModelForCausalLM.from_pretrained(
            config.model_name,
            **model_kwargs
        )

        lora_config = LoraConfig(
            r=config.lora_r,
            lora_alpha=config.lora_alpha,
            target_modules=config.lora_target_modules,
            lora_dropout=config.lora_dropout,
            bias="none",
            task_type=TaskType.CAUSAL_LM,
        )

        model = get_peft_model(model, lora_config)
        model.enable_input_require_grads()

    model.print_trainable_parameters()

    if torch.cuda.is_available():
        allocated = torch.cuda.memory_allocated() / 1e9
        reserved = torch.cuda.memory_reserved() / 1e9
        logger.info(f"Mémoire GPU - Allouée: {allocated:.2f} Go, Réservée: {reserved:.2f} Go")

    return model, tokenizer


def find_latest_checkpoint(output_dir: str) -> Optional[str]:
    """Trouve le dernier checkpoint"""
    checkpoint_dirs = []
    if os.path.exists(output_dir):
        for item in os.listdir(output_dir):
            item_path = os.path.join(output_dir, item)
            if os.path.isdir(item_path) and item.startswith("checkpoint-"):
                try:
                    step = int(item.split("-")[1])
                    checkpoint_dirs.append((step, item_path))
                except:
                    pass

    if checkpoint_dirs:
        checkpoint_dirs.sort(key=lambda x: x[0], reverse=True)
        latest_checkpoint = checkpoint_dirs[0][1]
        logger.info(f"Dernier checkpoint trouvé: {latest_checkpoint}")
        return latest_checkpoint

    return None


def main():
    """Fonction principale"""
    parser = argparse.ArgumentParser(description="Entraînement Mistral V2 - Documents Markdown")
    parser.add_argument("--resume", action="store_true", help="Reprendre depuis le dernier checkpoint")
    parser.add_argument("--resume-from", type=str, help="Reprendre depuis un checkpoint spécifique")
    parser.add_argument("--train-file", type=str, default="dataset/train_dataset.jsonl", help="Fichier d'entraînement")
    parser.add_argument("--val-file", type=str, default="dataset/val_dataset.jsonl", help="Fichier de validation")
    parser.add_argument("--no-sdpa", action="store_true", help="Désactiver SDPA")
    parser.add_argument("--output-dir", type=str, help="Dossier de sortie")
    parser.add_argument("--max-prompt-length", type=int, help="Longueur max du prompt")
    parser.add_argument("--max-response-length", type=int, help="Longueur max de la réponse")
    args = parser.parse_args()

    # Configuration
    model_config = ModelConfig()
    training_config = TrainingConfig()

    # Mise à jour avec les arguments
    if args.train_file:
        training_config.train_file = args.train_file
    if args.val_file:
        training_config.val_file = args.val_file
    if args.no_sdpa:
        model_config.use_sdpa = False
    if args.output_dir:
        model_config.output_dir = args.output_dir
        model_config.tensorboard_dir = os.path.join(args.output_dir, "runs")
    if args.max_prompt_length:
        model_config.max_prompt_length = args.max_prompt_length
    if args.max_response_length:
        model_config.max_response_length = args.max_response_length
        model_config.max_length = model_config.max_prompt_length + model_config.max_response_length

    # Gestion de la reprise
    resume_from_checkpoint = None
    if args.resume_from:
        resume_from_checkpoint = args.resume_from
    elif args.resume:
        latest = find_latest_checkpoint(model_config.output_dir)
        if latest:
            resume_from_checkpoint = latest
        else:
            logger.warning("Aucun checkpoint trouvé, démarrage from scratch")

    # Créer les dossiers
    os.makedirs(model_config.output_dir, exist_ok=True)
    os.makedirs(model_config.tensorboard_dir, exist_ok=True)

    # Seeds
    random.seed(model_config.seed)
    np.random.seed(model_config.seed)
    torch.manual_seed(model_config.seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(model_config.seed)

    # TensorBoard
    tb_writer = SummaryWriter(model_config.tensorboard_dir)

    logger.info("=" * 60)
    logger.info("=== ENTRAINEMENT V2 - DOCUMENTS MARKDOWN COMPLETS ===")
    logger.info("=== Expression de besoin -> Devis ===")
    logger.info("=" * 60)
    logger.info(f"Version transformers: {transformers.__version__}")
    logger.info(f"GPU: {torch.cuda.get_device_name(0) if torch.cuda.is_available() else 'N/A'}")
    if torch.cuda.is_available():
        logger.info(f"VRAM totale: {torch.cuda.get_device_properties(0).total_memory / 1e9:.1f} Go")

    if resume_from_checkpoint:
        logger.info(f"Reprise depuis: {resume_from_checkpoint}")
    else:
        logger.info("Entraînement from scratch")

    # Model & tokenizer
    model, tokenizer = setup_model_and_tokenizer(model_config, resume_from_checkpoint)

    # Dataset
    dataset_processor = DocumentDatasetProcessor(tokenizer, model_config)

    if not os.path.exists(training_config.train_file):
        logger.error(f"{training_config.train_file} non trouvé!")
        return
    if not os.path.exists(training_config.val_file):
        logger.error(f"{training_config.val_file} non trouvé!")
        return

    datasets, val_messages_original = dataset_processor.prepare_datasets_from_files(
        training_config.train_file,
        training_config.val_file
    )

    logger.info(f"Exemples validation pour génération: {len(val_messages_original)}")

    # Data collator
    data_collator = RobustDataCollator(tokenizer, pad_to_multiple_of=8)

    # Check initial loss
    if not resume_from_checkpoint:
        check_initial_loss(model, datasets['train'], data_collator, tokenizer)
    logger.info("\nPoursuite de l'entraînement...")

    # Training arguments
    training_args = TrainingArguments(
        output_dir=model_config.output_dir,
        overwrite_output_dir=False if resume_from_checkpoint else True,

        # Batch
        per_device_train_batch_size=training_config.per_device_train_batch_size,
        per_device_eval_batch_size=training_config.per_device_eval_batch_size,
        gradient_accumulation_steps=training_config.gradient_accumulation_steps,

        # Optimisation
        optim=training_config.optim,
        learning_rate=training_config.learning_rate,
        weight_decay=training_config.weight_decay,
        lr_scheduler_type="cosine",
        warmup_ratio=training_config.warmup_ratio,

        # Epochs
        num_train_epochs=training_config.num_train_epochs,

        # Mémoire
        gradient_checkpointing=training_config.gradient_checkpointing,

        # Precision
        bf16=training_config.bf16,
        fp16=training_config.fp16,

        # Logging
        logging_dir=model_config.tensorboard_dir,
        logging_steps=training_config.logging_steps,
        logging_first_step=True,

        # Évaluation
        eval_strategy="steps",
        eval_steps=training_config.eval_steps,

        # Sauvegarde
        save_strategy="steps",
        save_steps=training_config.save_steps,
        save_total_limit=training_config.save_total_limit,

        # Best model
        metric_for_best_model="eval_loss",
        greater_is_better=False,
        load_best_model_at_end=True,

        # Dataloader
        dataloader_num_workers=training_config.dataloader_num_workers,
        dataloader_pin_memory=training_config.dataloader_pin_memory,

        # Autres
        report_to=["tensorboard"],
        remove_unused_columns=False,

        # Stabilité
        max_grad_norm=1.0,
        seed=model_config.seed,
    )

    training_args.model_config = model_config

    # Créer le trainer
    trainer = ValidationTrainer(
        model=model,
        args=training_args,
        train_dataset=datasets['train'],
        eval_dataset=datasets['validation'],
        tokenizer=tokenizer,
        data_collator=data_collator,
        tb_writer=tb_writer,
        val_examples=val_messages_original,
    )

    # Infos
    effective_batch = training_config.per_device_train_batch_size * training_config.gradient_accumulation_steps
    total_steps = (len(datasets['train']) // effective_batch) * training_config.num_train_epochs

    logger.info("\n=== CONFIGURATION V2 ===")
    logger.info(f"Train: {len(datasets['train'])} exemples")
    logger.info(f"Validation: {len(datasets['validation'])} exemples")
    logger.info(f"Batch effectif: {effective_batch}")
    logger.info(f"Steps estimés: ~{total_steps}")
    logger.info(f"Prompt max (expression de besoin): {model_config.max_prompt_length} tokens")
    logger.info(f"Response max (devis): {model_config.max_response_length} tokens")
    logger.info(f"Total max: {model_config.max_length} tokens")
    logger.info(f"LoRA r: {model_config.lora_r}")
    logger.info(f"LoRA alpha: {model_config.lora_alpha}")
    logger.info(f"dtype: {model_config.torch_dtype}")
    logger.info(f"SDPA: {model_config.use_sdpa}")
    logger.info(f"Gradient checkpointing: {training_config.gradient_checkpointing}")

    try:
        logger.info("\n=== DEBUT DE L'ENTRAINEMENT ===")

        if resume_from_checkpoint:
            train_result = trainer.train(resume_from_checkpoint=resume_from_checkpoint)
        else:
            train_result = trainer.train()

        # Sauvegarder
        logger.info("\nSauvegarde...")
        trainer.save_model()
        tokenizer.save_pretrained(model_config.output_dir)

        # Historique complet
        if trainer.generation_history:
            history_path = os.path.join(model_config.output_dir, "generation_history_complete.json")
            with open(history_path, 'w', encoding='utf-8') as f:
                json.dump(trainer.generation_history, f, indent=2, ensure_ascii=False)
            logger.info(f"Historique complet sauvegardé: {history_path}")

        # Métriques
        logger.info("\n=== RESULTATS ===")
        logger.info(f"Loss finale: {train_result.metrics.get('train_loss', 'N/A'):.4f}")

        # Génération finale
        final_results = trainer.generate_validation_examples(num_examples=10)
        trainer.save_generation_results(final_results)

        # Stats finales mémoire
        if torch.cuda.is_available():
            max_allocated = torch.cuda.max_memory_allocated() / 1e9
            logger.info(f"Pic mémoire GPU: {max_allocated:.2f} Go")

        logger.info("\nTerminé!")

    except KeyboardInterrupt:
        logger.warning("\nInterrompu")
        trainer.save_model(f"{model_config.output_dir}/interrupted")

    except Exception as e:
        logger.error(f"\nErreur: {str(e)}", exc_info=True)
        raise

    finally:
        if tb_writer:
            tb_writer.close()

        if torch.cuda.is_available():
            torch.cuda.empty_cache()


if __name__ == "__main__":
    main()
