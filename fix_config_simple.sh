#!/bin/bash

echo "🔄 Restauration et modification ultra-simple..."

# 1. Restaurer le backup
cp ./html/js/config.js.backup ./html/js/config.js
cp ./html/js/config.js.backup ./html/js/.ipynb_checkpoints/config-checkpoint.js

# 2. Ajouter une fonction de détection en haut du fichier
sed -i '1i\
// Détection automatique du préfixe\
const detectPrefix = () => {\
    const match = window.location.pathname.match(/^(\\/scribe-ai\\/[^\\/]+\\/url-1)/);\
    return match ? match[1] : "/scribe-ai/test/url-1";\
};\
const DYNAMIC_PREFIX = detectPrefix();\
' ./html/js/config.js

# 3. Remplacer l'URL fixe par une URL dynamique
sed -i "s|'https://runai-poc-100.mlops.cagip.group.gca/scribe-ai/scribe2/url-1/api'|\`https://runai-poc-100.mlops.cagip.group.gca\${DYNAMIC_PREFIX}/api\`|g" ./html/js/config.js

# Faire pareil pour le checkpoint
sed -i '1i\
// Détection automatique du préfixe\
const detectPrefix = () => {\
    const match = window.location.pathname.match(/^(\\/scribe-ai\\/[^\\/]+\\/url-1)/);\
    return match ? match[1] : "/scribe-ai/test/url-1";\
};\
const DYNAMIC_PREFIX = detectPrefix();\
' ./html/js/.ipynb_checkpoints/config-checkpoint.js

sed -i "s|'https://runai-poc-100.mlops.cagip.group.gca/scribe-ai/scribe2/url-1/api'|\`https://runai-poc-100.mlops.cagip.group.gca\${DYNAMIC_PREFIX}/api\`|g" ./html/js/.ipynb_checkpoints/config-checkpoint.js

echo "✅ Modification ultra-simple terminée!"
echo "🚀 Tous les exports originaux sont préservés!"
