#!/bin/bash

#===============================================================================
# Script d'installation automatisé pour l'environnement NDC
# - Configure le proxy
# - Installe Anaconda
# - Crée un environnement conda Python 3.12
# - Clone le repo et installe les dépendances
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
# CONFIGURATION - À MODIFIER SELON VOS BESOINS
#===============================================================================
PROXY_HOST="x.x.x.x"
PROXY_PORT="8080"
PROXY_URL="http://${PROXY_HOST}:${PROXY_PORT}"

ANACONDA_VERSION="2024.02-1"  # Version stable récente
ANACONDA_INSTALLER="Anaconda3-${ANACONDA_VERSION}-Linux-x86_64.sh"
ANACONDA_URL="https://repo.anaconda.com/archive/${ANACONDA_INSTALLER}"
ANACONDA_INSTALL_DIR="$HOME/anaconda3"

CONDA_ENV_NAME="ndc-dev"
PYTHON_VERSION="3.12"

REPO_URL="https://github.com/TeamCLP/ndc-dev.git"
REPO_DIR="$HOME/ndc-dev"

#===============================================================================
# 1. Configuration du Proxy
#===============================================================================
setup_proxy() {
    log_info "Configuration du proxy..."
    
    export HTTP_PROXY="${PROXY_URL}"
    export HTTPS_PROXY="${PROXY_URL}"
    export http_proxy="${PROXY_URL}"
    export https_proxy="${PROXY_URL}"
    export NO_PROXY="localhost,127.0.0.1"
    export no_proxy="localhost,127.0.0.1"
    
    # Persister dans .bashrc si pas déjà présent
    if ! grep -q "HTTPS_PROXY" "$HOME/.bashrc" 2>/dev/null; then
        log_info "Ajout du proxy dans ~/.bashrc..."
        cat >> "$HOME/.bashrc" << EOF

# Proxy configuration
export HTTP_PROXY="${PROXY_URL}"
export HTTPS_PROXY="${PROXY_URL}"
export http_proxy="${PROXY_URL}"
export https_proxy="${PROXY_URL}"
export NO_PROXY="localhost,127.0.0.1"
export no_proxy="localhost,127.0.0.1"
EOF
    fi
    
    # Configuration git pour le proxy
    git config --global http.proxy "${PROXY_URL}"
    git config --global https.proxy "${PROXY_URL}"
    
    log_info "Proxy configuré: ${PROXY_URL}"
}

#===============================================================================
# 2. Installation d'Anaconda
#===============================================================================
install_anaconda() {
    if [ -d "$ANACONDA_INSTALL_DIR" ]; then
        log_warn "Anaconda déjà installé dans $ANACONDA_INSTALL_DIR"
        return 0
    fi
    
    log_info "Téléchargement d'Anaconda ${ANACONDA_VERSION}..."
    
    cd /tmp
    if [ ! -f "$ANACONDA_INSTALLER" ]; then
        wget --progress=bar:force -e use_proxy=yes -e https_proxy="${PROXY_URL}" \
            "${ANACONDA_URL}" -O "$ANACONDA_INSTALLER"
    fi
    
    log_info "Installation d'Anaconda..."
    bash "$ANACONDA_INSTALLER" -b -p "$ANACONDA_INSTALL_DIR"
    
    # Initialiser conda
    log_info "Initialisation de conda..."
    "$ANACONDA_INSTALL_DIR/bin/conda" init bash
    
    # Nettoyer l'installeur
    rm -f "$ANACONDA_INSTALLER"
    
    log_info "Anaconda installé avec succès"
}

#===============================================================================
# 3. Activation de conda et configuration
#===============================================================================
setup_conda() {
    log_info "Configuration de conda..."
    
    # Charger conda
    source "$ANACONDA_INSTALL_DIR/etc/profile.d/conda.sh"
    
    # Configurer le proxy pour conda
    conda config --set proxy_servers.http "${PROXY_URL}"
    conda config --set proxy_servers.https "${PROXY_URL}"
    
    # Configurer pip pour utiliser le proxy globalement
    mkdir -p "$HOME/.config/pip"
    cat > "$HOME/.config/pip/pip.conf" << EOF
[global]
proxy = ${PROXY_URL}
trusted-host = pypi.org
               pypi.python.org
               files.pythonhosted.org
EOF
    
    log_info "Conda et pip configurés pour le proxy"
}

#===============================================================================
# 4. Création de l'environnement conda
#===============================================================================
create_conda_env() {
    source "$ANACONDA_INSTALL_DIR/etc/profile.d/conda.sh"
    
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
    conda create -n "$CONDA_ENV_NAME" python="${PYTHON_VERSION}" -y
    
    log_info "Environnement conda créé"
}

#===============================================================================
# 5. Clone du repository
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
    git clone "$REPO_URL" "$REPO_DIR"
    
    log_info "Repository cloné dans $REPO_DIR"
}

#===============================================================================
# 6. Installation des dépendances
#===============================================================================
install_dependencies() {
    source "$ANACONDA_INSTALL_DIR/etc/profile.d/conda.sh"
    conda activate "$CONDA_ENV_NAME"
    
    cd "$REPO_DIR"
    
    if [ ! -f "requirements.txt" ]; then
        log_error "requirements.txt non trouvé dans $REPO_DIR"
        exit 1
    fi
    
    log_info "Installation des dépendances conda (si environment.yml existe)..."
    if [ -f "environment.yml" ]; then
        conda env update -f environment.yml --prune
    fi
    
    log_info "Installation des dépendances pip depuis requirements.txt..."
    pip install --upgrade pip
    pip install -r requirements.txt
    
    log_info "Dépendances installées"
}

#===============================================================================
# 7. Création du script de lancement
#===============================================================================
create_launcher() {
    LAUNCHER="$REPO_DIR/run_train.sh"
    
    cat > "$LAUNCHER" << EOF
#!/bin/bash
# Script de lancement pour train.py

# Configuration du proxy
export HTTP_PROXY="${PROXY_URL}"
export HTTPS_PROXY="${PROXY_URL}"
export http_proxy="${PROXY_URL}"
export https_proxy="${PROXY_URL}"

# Activation de l'environnement conda
source "$ANACONDA_INSTALL_DIR/etc/profile.d/conda.sh"
conda activate "$CONDA_ENV_NAME"

# Lancement de l'entraînement
cd "$REPO_DIR"
python train.py "\$@"
EOF
    
    chmod +x "$LAUNCHER"
    
    log_info "Script de lancement créé: $LAUNCHER"
}

#===============================================================================
# MAIN
#===============================================================================
main() {
    echo "=============================================="
    echo "  Installation de l'environnement NDC-Dev"
    echo "=============================================="
    echo ""
    
    # Vérifier qu'on n'est pas root
    if [ "$EUID" -eq 0 ]; then
        log_error "Ne pas exécuter ce script en tant que root"
        exit 1
    fi
    
    setup_proxy
    install_anaconda
    setup_conda
    create_conda_env
    clone_repo
    install_dependencies
    create_launcher
    
    echo ""
    echo "=============================================="
    log_info "Installation terminée avec succès!"
    echo "=============================================="
    echo ""
    echo "Pour utiliser l'environnement:"
    echo ""
    echo "  1. Recharger votre shell ou exécuter:"
    echo "     source ~/.bashrc"
    echo ""
    echo "  2. Activer l'environnement:"
    echo "     conda activate ${CONDA_ENV_NAME}"
    echo ""
    echo "  3. Lancer l'entraînement:"
    echo "     cd ${REPO_DIR}"
    echo "     python train.py"
    echo ""
    echo "  Ou utilisez le script de lancement:"
    echo "     ${REPO_DIR}/run_train.sh"
    echo ""
}

# Exécution
main "$@"
