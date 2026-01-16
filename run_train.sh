#!/bin/bash
# Script de lancement pour train.py

# Configuration du proxy
export HTTPS_PROXY="http://10.246.42.30:8080"
export HTTP_PROXY="http://10.246.42.30:8080"
export https_proxy="http://10.246.42.30:8080"
export http_proxy="http://10.246.42.30:8080"

source "/root/miniconda3/etc/profile.d/conda.sh"
conda activate "ndc-dev"

cd "/home/quentin/ndc-dev"
python train.py "$@"
