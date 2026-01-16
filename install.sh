#!/bin/bash

#===============================================================================
# Script d'installation automatisé pour l'environnement NDC
# - Installe Miniconda (plus léger qu'Anaconda)
# - Crée un environnement conda Python 3.12
# - Clone le repo et installe les dépendances
# - Gère le proxy HTTPS
#===============================================================================

set -e  # Arrêter en cas d'erreur

# Couleurs pour les logs
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

log_info() { echo -e "${GREEN}[INFO]${NC} $1"; }
log_warn() { echo -e "${YELLOW}[WARN]${NC} $1"; }
log_error() { echo -e "${RED}[ERROR]${NC} $1"; }

#===============================================================================
# CONFIGURATION
#===============================================================================
MINICONDA_INSTALLER="Miniconda3-latest-Linux-x86_64.sh"
MINICONDA_URL="https://repo.anaconda.com/miniconda/$MINICONDA_INSTALLER"
MINICONDA_INSTALL_DIR="$HOME/miniconda3"

CONDA_ENV_NAME="ndc-dev"
PYTHON_VERSION="3.12"

REPO_URL="https://github.com/TeamCLP/ndc-dev.git"
REPO_DIR="$HOME/ndc-dev"

# Configuration du proxy
PROXY_URL="http://10.246.42.30:8080"

#===============================================================================
# Configuration du proxy
#===============================================================================
setup_proxy() {
    log_info "Configuration du proxy..."
    export HTTPS_PROXY="$PROXY_URL"
    export HTTP_PROXY="$PROXY_URL"
    export https_proxy="$PROXY_URL"
    export http_proxy="$PROXY_URL"
    
    # Ne pas configurer CONDA_PROXY_SERVERS ici car cela cause l'erreur
    unset CONDA_PROXY_SERVERS
    
    log_info "Proxy configuré: $PROXY_URL"
}

#===============================================================================
# 1. Installation de Miniconda
#===============================================================================
install_miniconda() {
    if [ -d "$MINICONDA_INSTALL_DIR" ]; then
        log_warn "Miniconda déjà installé dans $MINICONDA_INSTALL_DIR"
        return 0
    fi
    
    log_info "Téléchargement de Miniconda..."
    
    cd /tmp
    if [ ! -f "$MINICONDA_INSTALLER" ]; then
        # Utilisation de curl avec proxy
        curl -L --proxy "$PROXY_URL" -o "$MINICONDA_INSTALLER" "$MINICONDA_URL"
        
        # Alternative avec wget si curl échoue
        if [ $? -ne 0 ]; then
            log_warn "Échec avec curl, tentative avec wget..."
            wget --progress=bar:force -e use_proxy=yes -e https_proxy="$PROXY_URL" "$MINICONDA_URL" -O "$MINICONDA_INSTALLER"
        fi
    fi
    
    log_info "Installation de Miniconda..."
    bash "$MINICONDA_INSTALLER" -b -p "$MINICONDA_INSTALL_DIR"
    
    log_info "Initialisation de conda..."
    "$MINICONDA_INSTALL_DIR/bin/conda" init bash
    
    # Configuration du proxy pour conda (méthode correcte)
    log_info "Configuration du proxy pour conda..."
    "$MINICONDA_INSTALL_DIR/bin/conda" config --set proxy_servers.https "$PROXY_URL"
    "$MINICONDA_INSTALL_DIR/bin/conda" config --set proxy_servers.http "$PROXY_URL"
    
    # Accepter les conditions d'utilisation
    log_info "Acceptation des conditions d'utilisation conda..."
    "$MINICONDA_INSTALL_DIR/bin/conda" config --set channel_priority strict
    "$MINICONDA_INSTALL_DIR/bin/conda" tos accept --override-channels --channel https://repo.anaconda.com/pkgs/main || true
    "$MINICONDA_INSTALL_DIR/bin/conda" tos accept --override-channels --channel https://repo.anaconda.com/pkgs/r || true
    
    rm -f "$MINICONDA_INSTALLER"
    
    log_info "Miniconda installé avec succès"
}

#===============================================================================
# 2. Création de l'environnement conda
#===============================================================================
create_conda_env() {
    source "$MINICONDA_INSTALL_DIR/etc/profile.d/conda.sh"
    
    if conda env list | grep -q "^${CONDA_ENV_NAME} "; then
        log_warn "L'environnement '${CONDA_ENV_NAME}' existe déjà"
        read -p "Voulez-vous le recréer ? (y/N) " -n 1 -r
        echo
        if [[ $REPLY =~ ^[Yy]$ ]]; then
            conda env remove -n "$CONDA_ENV_NAME" -y
        else
            return 0
        fi
    fi
    
    log_info "Création de l'environnement conda '${CONDA_ENV_NAME}' avec Python ${PYTHON_VERSION}..."
    
    # Utiliser conda-forge pour éviter les problèmes de ToS
    conda create -n "$CONDA_ENV_NAME" python="$PYTHON_VERSION" -c conda-forge -y
    
    log_info "Environnement conda créé"
}

#===============================================================================
# 3. Clone du repository
#===============================================================================
clone_repo() {
    if [ -d "$REPO_DIR" ]; then
        log_warn "Le répertoire $REPO_DIR existe déjà"
        read -p "Voulez-vous le supprimer et re-cloner ? (y/N) " -n 1 -r
        echo
        if [[ $REPLY =~ ^[Yy]$ ]]; then
            rm -rf "$REPO_DIR"
        else
            log_info "Mise à jour du repo existant..."
            cd "$REPO_DIR"
            git pull
            return 0
        fi
    fi
    
    log_info "Clonage du repository..."
    
    # Configuration du proxy pour git
    git config --global http.proxy "$PROXY_URL"
    git config --global https.proxy "$PROXY_URL"
    
    git clone "$REPO_URL" "$REPO_DIR"
    
    log_info "Repository cloné dans $REPO_DIR"
}

#===============================================================================
# 4. Installation des dépendances
#===============================================================================
install_dependencies() {
    source "$MINICONDA_INSTALL_DIR/etc/profile.d/conda.sh"
    conda activate "$CONDA_ENV_NAME"
    
    cd "$REPO_DIR"
    
    # Configuration du proxy pour pip dans le répertoire utilisateur
    mkdir -p ~/.pip
    cat > ~/.pip/pip.conf << EOF
[global]
proxy = $PROXY_URL
EOF
    
    if [ ! -f "requirements.txt" ]; then
        log_error "requirements.txt non trouvé dans $REPO_DIR"
        exit 1
    fi
    
    if [ -f "environment.yml" ]; then
        log_info "Installation des dépendances conda depuis environment.yml..."
        conda env update -f environment.yml --prune
    fi
    
    log_info "Installation des dépendances pip depuis requirements.txt..."
    pip install --upgrade pip
    pip install -r requirements.txt
    
    log_info "Dépendances installées"
}

#===============================================================================
# 5. Création du script de lancement
#===============================================================================
create_launcher() {
    LAUNCHER="$REPO_DIR/run_train.sh"
    
    cat > "$LAUNCHER" << EOF
#!/bin/bash
# Script de lancement pour train.py

# Configuration du proxy
export HTTPS_PROXY="$PROXY_URL"
export HTTP_PROXY="$PROXY_URL"
export https_proxy="$PROXY_URL"
export http_proxy="$PROXY_URL"

source "$MINICONDA_INSTALL_DIR/etc/profile.d/conda.sh"
conda activate "$CONDA_ENV_NAME"

cd "$REPO_DIR"
python train.py "\$@"
EOF
    
    chmod +x "$LAUNCHER"
    
    log_info "Script de lancement créé: $LAUNCHER"
}

#===============================================================================
# 6. Nettoyage de la configuration proxy (optionnel)
#===============================================================================
cleanup_proxy_config() {
    read -p "Voulez-vous nettoyer la configuration proxy globale de git ? (y/N) " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        git config --global --unset http.proxy 2>/dev/null || true
        git config --global --unset https.proxy 2>/dev/null || true
        log_info "Configuration proxy git nettoyée"
    fi
}

#===============================================================================
# MAIN
#===============================================================================
main() {
    echo "=============================================="
    echo "  Installation de l'environnement NDC-Dev"
    echo "=============================================="
    echo ""
    
    setup_proxy
    install_miniconda
    create_conda_env
    clone_repo
    install_dependencies
    create_launcher
    cleanup_proxy_config
    
    echo ""
    echo "=============================================="
    log_info "Installation terminée avec succès!"
    echo "=============================================="
    echo ""
    echo "Pour lancer l'entraînement:"
    echo ""
    echo "  source ~/.bashrc"
    echo "  conda activate ${CONDA_ENV_NAME}"
    echo "  cd ${REPO_DIR}"
    echo "  python train.py"
    echo ""
    echo "Ou via le launcher: ${REPO_DIR}/run_train.sh"
    echo ""
    echo "Note: Le proxy est configuré automatiquement dans le launcher"
}

main "$@"
