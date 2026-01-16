#!/bin/bash

#===============================================================================
# Script de démarrage des services NDC
# - Arrête les processus existants
# - Active l'environnement conda
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

# Configuration du proxy
PROXY_URL="http://10.246.42.30:8080"

#===============================================================================
# 1. Arrêt des processus existants
#===============================================================================
stop_existing_processes() {
    log_info "Arrêt des processus existants..."
    
    pkill -9 -f "run_train.sh" 2>/dev/null || true
    pkill -9 -f "tensorboard" 2>/dev/null || true
    
    log_info "Processus existants arrêtés"
}

#===============================================================================
# 2. Configuration de l'environnement
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
    
    log_info "Environnement configuré"
}

#===============================================================================
# 3. Lancement des services
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
    
    stop_existing_processes
    setup_environment
    start_services
}

main "$@"
