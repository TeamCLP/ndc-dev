#!/bin/bash

#===============================================================================
# Script de démarrage des services NDC
# - Active l'environnement conda
# - Vérifie la présence du modèle mistral-banking
# - Télécharge le modèle si nécessaire
# - Arrête les processus existants
# - Lance l'API et le serveur HTML en parallèle
#===============================================================================

# Couleurs pour les logs
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

log_info() { echo -e "${GREEN}[INFO]${NC} $1"; }
log_warn() { echo -e "${YELLOW}[WARN]${NC} $1"; }
log_error() { echo -e "${RED}[ERROR]${NC} $1"; }

# Configuration
PROXY_URL="http://10.246.42.30:8080"
MODEL_DIR="/home/quentin/mistral-banking"
CHECKPOINT_DIR="$MODEL_DIR/checkpoint-100"
HUGGINGFACE_REPO="TeamCLP/ndc-dev"

#===============================================================================
# 1. Configuration de l'environnement
#===============================================================================
setup_environment() {
    log_info "Configuration de l'environnement..."
    
    # Configuration du proxy
    export HTTPS_PROXY="$PROXY_URL"
    export HTTP_PROXY="$PROXY_URL"
    export https_proxy="$PROXY_URL"
    export http_proxy="$PROXY_URL"
    
    # Activation de conda
    source ~/.bashrc
    source "$HOME/miniconda3/etc/profile.d/conda.sh"
    conda activate ndc-dev
    
    log_info "Environnement conda activé"
}

#===============================================================================
# 2. Vérification et téléchargement du modèle
#===============================================================================
check_and_download_model() {
    log_info "Vérification de la présence du modèle..."
    
    # Vérifier si le dossier existe et contient au moins un checkpoint
    if [[ -d "$MODEL_DIR" ]] && ls "$MODEL_DIR"/checkpoint-* 1> /dev/null 2>&1; then
        log_info "Modèle trouvé dans $MODEL_DIR"
        return 0
    fi
    
    log_warn "Modèle non trouvé ou incomplet. Téléchargement en cours..."
    
    # Créer les dossiers s'ils n'existent pas
    mkdir -p "$CHECKPOINT_DIR"
    
    # Téléchargement avec huggingface-hub (conda déjà activé)
    log_info "Téléchargement du modèle depuis HuggingFace dans $CHECKPOINT_DIR..."
    
    # Utilisation de huggingface_hub pour télécharger directement dans checkpoint-100
    python -c "
import os
from huggingface_hub import snapshot_download

# Configuration du proxy pour les requêtes Python
os.environ['HTTPS_PROXY'] = '$PROXY_URL'
os.environ['HTTP_PROXY'] = '$PROXY_URL'

try:
    snapshot_download(
        repo_id='$HUGGINGFACE_REPO',
        local_dir='$CHECKPOINT_DIR',
        local_dir_use_symlinks=False
    )
    print('Téléchargement terminé avec succès')
except Exception as e:
    print(f'Erreur lors du téléchargement: {e}')
    exit(1)
"
    
    if [ $? -eq 0 ]; then
        log_info "Modèle téléchargé avec succès dans $CHECKPOINT_DIR"
    else
        log_error "Échec du téléchargement du modèle"
        exit 1
    fi
}

#===============================================================================
# 3. Arrêt des processus existants
#===============================================================================
stop_existing_processes() {
    log_info "Arrêt des processus existants..."
    
    pkill -9 -f "run_train.sh" 2>/dev/null || true
    pkill -9 -f "tensorboard" 2>/dev/null || true
    
    log_info "Processus existants arrêtés"
}

#===============================================================================
# 4. Lancement des services
#===============================================================================
start_services() {
    log_info "Démarrage des services..."
    
    # Lancement de l'API en arrière-plan
    python api.py &
    API_PID=$!
    
    # Lancement du serveur HTML en arrière-plan
    python html/serv.py &
    HTML_PID=$!
    
    log_info "Services démarrés:"
    echo "  - API (PID: $API_PID)"
    echo "  - HTML Server (PID: $HTML_PID)"
    
    # Fonction pour arrêter proprement les services
    cleanup() {
        log_warn "Arrêt des services..."
        kill $API_PID $HTML_PID 2>/dev/null || true
        exit 0
    }
    
    # Capturer Ctrl+C pour arrêter proprement
    trap cleanup SIGINT SIGTERM
    
    # Attendre que les processus se terminent
    wait $API_PID $HTML_PID
}

#===============================================================================
# MAIN
#===============================================================================
main() {
    echo "=============================================="
    echo "       Démarrage des services NDC"
    echo "=============================================="
    echo ""
    
    setup_environment
    check_and_download_model
    stop_existing_processes
    start_services
}

main "$@"
