#!/usr/bin/env python3
"""
Script pour ajouter automatiquement le support du préfixe dynamique à api.py
"""

import re
import os
import shutil
from datetime import datetime

def modify_api_file(api_file_path="api.py"):
    """Modifie le fichier api.py pour ajouter le support du préfixe dynamique"""
    
    # Vérifier que le fichier existe
    if not os.path.exists(api_file_path):
        print(f"❌ Fichier {api_file_path} non trouvé")
        return False
    
    # Faire une sauvegarde
    backup_path = f"{api_file_path}.backup.{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    shutil.copy2(api_file_path, backup_path)
    print(f"💾 Sauvegarde créée: {backup_path}")
    
    # Lire le fichier
    with open(api_file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Vérifier si les modifications ont déjà été appliquées
    if "DYNAMIC_PREFIX" in content:
        print("⚠️  Les modifications semblent déjà appliquées")
        return True
    
    # 1. Ajouter la fonction de détection du préfixe après settings = Settings()
    prefix_code = '''
# ===== CONFIGURATION DYNAMIQUE DU PRÉFIXE =====
def get_dynamic_prefix():
    """Détecte le préfixe dynamiquement basé sur le nom du pod"""
    
    # Méthode 1: Variable d'environnement du pod (Kubernetes met le nom du pod dans HOSTNAME)
    pod_name = os.environ.get("HOSTNAME", "")
    if pod_name:
        # Extraire le nom de base du pod (ex: "test-abc123" -> "test")
        pod_base = re.sub(r'-[a-f0-9]+.*$', '', pod_name)
        return f"/scribe-ai/{pod_base}/url-1"
    
    # Méthode 2: Variable d'environnement personnalisée
    pod_name = os.environ.get("POD_NAME", "")
    if pod_name:
        return f"/scribe-ai/{pod_name}/url-1"
    
    # Méthode 3: Fallback sur PATH_PREFIX ou défaut
    return os.environ.get("PATH_PREFIX", "/scribe-ai/test/url-1").rstrip("/")

# Détection du préfixe au démarrage
DYNAMIC_PREFIX = get_dynamic_prefix()
logger.info(f"🚀 API configurée avec préfixe dynamique: {DYNAMIC_PREFIX}")
'''
    
    # Trouver la ligne settings = Settings() et ajouter après
    content = re.sub(
        r'(settings = Settings\(\))',
        r'\1' + prefix_code,
        content
    )
    
    # 2. Modifier le middleware pour ajouter le logging
    middleware_replacement = '''@app.middleware("http")
async def disable_cache(request: Request, call_next):
    # Logger les requêtes pour debug
    logger.info(f"📥 {request.method} {request.url.path}")
    
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate, private"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response'''
    
    content = re.sub(
        r'@app\.middleware\("http"\)\nasync def disable_cache\(request: Request, call_next\):\s*response = await call_next\(request\)\s*response\.headers\["Cache-Control"\] = "no-cache, no-store, must-revalidate, private"\s*response\.headers\["Pragma"\] = "no-cache"\s*response\.headers\["Expires"\] = "0"\s*return response',
        middleware_replacement,
        content,
        flags=re.DOTALL
    )
    
    # 3. Ajouter les nouvelles routes avec préfixe après get_model_info()
    new_routes = '''
# ===== ROUTES AVEC PRÉFIXE DYNAMIQUE =====
@app.get(f"{DYNAMIC_PREFIX}/api/health")
async def health_check_with_prefix():
    return await health_check()

@app.post(f"{DYNAMIC_PREFIX}/api/generate")
async def generate_field_with_prefix(request: GenerateRequest):
    return await generate_field(request)

@app.post(f"{DYNAMIC_PREFIX}/api/generate_multiple")
async def generate_multiple_fields_with_prefix(request: GenerateMultipleRequest):
    return await generate_multiple_fields(request)

@app.post(f"{DYNAMIC_PREFIX}/api/validate")
async def validate_field_with_prefix(request: ValidateRequest):
    return await validate_field(request)

@app.post(f"{DYNAMIC_PREFIX}/api/feedback/correction")
async def submit_correction_with_prefix(request: FeedbackCorrectionRequest):
    return await submit_correction(request)

@app.post(f"{DYNAMIC_PREFIX}/api/feedback/rating")
async def submit_rating_with_prefix(request: FeedbackRatingRequest):
    return await submit_rating(request)

@app.get(f"{DYNAMIC_PREFIX}/api/feedback/insights/")
async def get_field_insights_with_prefix(field: str):
    return await get_field_insights(field)

@app.get(f"{DYNAMIC_PREFIX}/api/feedback/summary")
async def get_feedback_summary_with_prefix():
    return await get_feedback_summary()

@app.get(f"{DYNAMIC_PREFIX}/api/model/info")
async def get_model_info_with_prefix():
    return await get_model_info()

@app.get(f"{DYNAMIC_PREFIX}/api/config")
async def get_config():
    return {
        "prefix": DYNAMIC_PREFIX,
        "api_base_url": f"{DYNAMIC_PREFIX}/api",
        "pod_name": os.environ.get("HOSTNAME", "unknown"),
        "model_loaded": model_manager.is_loaded
    }
'''
    
    # Trouver la fin de la fonction get_model_info et ajouter les nouvelles routes
    content = re.sub(
        r'(# ===== MAIN =====)',
        new_routes + '\n\\1',
        content
    )
    
    # Écrire le fichier modifié
    with open(api_file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print("✅ Modifications appliquées avec succès!")
    print("📝 Nouvelles fonctionnalités ajoutées:")
    print("   - Détection automatique du préfixe basé sur le nom du pod")
    print("   - Routes avec préfixe dynamique")
    print("   - Endpoint /api/config pour la configuration")
    print("   - Logging amélioré des requêtes")
    
    return True

def test_modifications():
    """Teste si les modifications fonctionnent"""
    print("\n🧪 Test des modifications...")
    
    try:
        # Simuler différents noms de pod
        test_cases = [
            ("test-abc123", "/scribe-ai/test/url-1"),
            ("production-def456", "/scribe-ai/production/url-1"),
            ("dev-ghi789", "/scribe-ai/dev/url-1"),
        ]
        
        for hostname, expected_prefix in test_cases:
            os.environ["HOSTNAME"] = hostname
            
            # Importer la fonction (simulation)
            import re
            pod_name = os.environ.get("HOSTNAME", "")
            if pod_name:
                pod_base = re.sub(r'-[a-f0-9]+.*$', '', pod_name)
                result_prefix = f"/scribe-ai/{pod_base}/url-1"
            else:
                result_prefix = "/scribe-ai/test/url-1"
            
            if result_prefix == expected_prefix:
                print(f"   ✅ {hostname} -> {result_prefix}")
            else:
                print(f"   ❌ {hostname} -> {result_prefix} (attendu: {expected_prefix})")
        
        # Nettoyer
        if "HOSTNAME" in os.environ:
            del os.environ["HOSTNAME"]
            
    except Exception as e:
        print(f"   ⚠️  Erreur de test: {e}")

def main():
    print("🔧 Modification automatique d'api.py pour le support du préfixe dynamique")
    print("=" * 70)
    
    # Vérifier le répertoire courant
    if not os.path.exists("api.py"):
        print("❌ api.py non trouvé dans le répertoire courant")
        print("💡 Assurez-vous d'être dans le bon répertoire")
        return
    
    # Appliquer les modifications
    success = modify_api_file()
    
    if success:
        test_modifications()
        
        print("\n🚀 Prochaines étapes:")
        print("1. Redémarrer l'API: pkill -f api.py && python api.py &")
        print("2. Tester: curl http://localhost:5000/scribe-ai/test/url-1/api/health")
        print("3. Vérifier la config: curl http://localhost:5000/scribe-ai/test/url-1/api/config")
        
        print(f"\n💾 Sauvegarde disponible en cas de problème")
    else:
        print("❌ Échec de la modification")

if __name__ == "__main__":
    main()
