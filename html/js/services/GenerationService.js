// js/services/GenerationService.js
import { api } from '../core/Api.js';
import { store } from '../core/Store.js';
import { Config } from '../config.js';
import { getFieldHtmlId, showNotification } from '../utils/helpers.js';

/**
 * Service de génération IA - Version alignée avec le modèle entraîné
 * 
 * IMPORTANT: Ce service construit le contexte EXACTEMENT comme pendant l'entraînement
 * - Champs de base TOUJOURS inclus
 * - Mapping des noms (sector → secteur, etc.)
 * - Dépendances spécifiques par champ
 * - Formatage spécial pour les tableaux
 */
export class GenerationService {
    constructor() {
        this.isGenerating = false;
        this.feedbackTimeouts = new Map();
        this.generationCache = new Map();
        
        // Configuration identique à l'entraînement
        this.ALWAYS_INCLUDED = ['client', 'secteur', 'typeProjet', 'complexite', 'libelle', 'annee'];
        
        // Ordre de génération (pour référence)
        this.FIELD_GENERATION_ORDER = [
            "contexte_proj", "besoin", "objectifs", "perimetre", "horsPerimetre",
            "contraintes", "risques", "descriptionSolution", "architecture", 
            "composantsDimensionnement", "phases", "jalons", "livrables",
            "conditionsHorsCrash", "conditionsCrashSite", "resilienceApplicative",
            "praPlanDegrade", "sauvegardes", "administrationSupervision",
            "impactCO2", "modalitesPartage", "coutsConstruction", "coutsFonctionnement"
        ];
        
        // Dépendances par champ (IDENTIQUE à l'entraînement)
        this.RELEVANT_CONTEXT = {
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
        };
        
        // Pool de champs additionnels
        this.ADDITIONAL_FIELDS_POOL = [
            'contexte_proj', 'besoin', 'objectifs', 'perimetre', 'horsPerimetre',
            'descriptionSolution', 'architecture', 'contraintes', 'risques',
            'phases', 'composantsDimensionnement'
        ];
        
        // Champs qui sont des tableaux
        this.TABLE_FIELDS = [
            'contraintes', 'risques', 'phases', 'livrables', 'jalons',
            'coutsConstruction', 'coutsFonctionnement'
        ];
    }
    
    /**
     * Mappe les noms de champs frontend vers ceux du modèle
     */
    mapFieldNames(frontendContext) {
        const mapping = {
            'sector': 'secteur',
            'project_type': 'typeProjet',
            'complexity': 'complexite',
            'contexte': 'contexte_proj',
            'lots': 'phases'
        };
        
        const modelContext = {};
        
        for (const [key, value] of Object.entries(frontendContext)) {
            const modelKey = mapping[key] || key;
            modelContext[modelKey] = value;
        }
        
        return modelContext;
    }
    
    /**
     * Formate un tableau pour le contexte (pipe-separated)
     */
    formatTableForContext(data, fieldName) {
        if (!data || !Array.isArray(data) || data.length === 0) {
            return "";
        }
        
        const rows = [];
        
        for (const item of data) {
            const row = [];
            
            // Extraire les valeurs dans l'ordre col0, col1, col2...
            let colIndex = 0;
            while (item[`col${colIndex}`] !== undefined) {
                row.push(item[`col${colIndex}`] || '');
                colIndex++;
            }
            
            if (row.length > 0) {
                rows.push(row.join('|'));
            }
        }
        
        return rows.join('\n');
    }
    
    /**
     * Collecte le contexte EXACTEMENT comme pendant l'entraînement
     */
    collectContext(targetField = null) {
        const projectData = store.getProjectData();
        const frontendContext = {};
        
        // Collecter tous les champs non vides
        Object.entries(projectData).forEach(([key, value]) => {
            if (value !== undefined && value !== null && value !== '') {
                if (Array.isArray(value) && value.length > 0) {
                    frontendContext[key] = value;
                } else if (typeof value === 'string' && value.trim() !== '') {
                    frontendContext[key] = value;
                } else if (typeof value !== 'string') {
                    frontendContext[key] = value;
                }
            }
        });
        
        // Mapper vers les noms du modèle
        const modelContext = this.mapFieldNames(frontendContext);
        
        return modelContext;
    }
    
    /**
     * Construit le contexte selon la stratégie (IDENTIQUE à l'entraînement)
     */
    buildContextString(targetField, modelContext, strategy = 'related') {
        const contextParts = [];
        const usedFields = [];
        
        // Mapper le nom du champ cible si nécessaire
        const targetTask = this.mapFieldToTask(targetField);
        
        // 1. TOUJOURS inclure les champs de base
        for (const field of this.ALWAYS_INCLUDED) {
            if (modelContext[field]) {
                if (this.TABLE_FIELDS.includes(field)) {
                    const formattedTable = this.formatTableForContext(modelContext[field], field);
                    if (formattedTable) {
                        contextParts.push(`${field}:\n${formattedTable}`);
                        usedFields.push(field);
                    }
                } else {
                    contextParts.push(`${field}: ${modelContext[field]}`);
                    usedFields.push(field);
                }
            }
        }
        
        // 2. Déterminer la complexité selon la stratégie
        let complexity = 'medium';
        if (strategy === 'minimal') {
            complexity = 'minimal';
        } else if (strategy === 'full' || strategy === 'debug') {
            complexity = 'full';
        }
        
        // 3. Ajouter les champs selon la complexité
        const availableFields = this.FIELD_GENERATION_ORDER.filter(f => f !== targetTask);
        
        if (complexity === 'minimal') {
            // Ajouter jusqu'à 2 dépendances pertinentes
            const relevant = this.RELEVANT_CONTEXT[targetTask] || [];
            const toAdd = relevant.slice(0, 2);
            
            for (const field of toAdd) {
                if (modelContext[field] && !usedFields.includes(field)) {
                    if (this.TABLE_FIELDS.includes(field)) {
                        const formattedTable = this.formatTableForContext(modelContext[field], field);
                        if (formattedTable) {
                            contextParts.push(`${field}:\n${formattedTable}`);
                            usedFields.push(field);
                        }
                    } else {
                        contextParts.push(`${field}: ${modelContext[field]}`);
                        usedFields.push(field);
                    }
                }
            }
            
        } else if (complexity === 'medium') {
            // Ajouter jusqu'à 3 dépendances pertinentes + 1 champ aléatoire
            const relevant = this.RELEVANT_CONTEXT[targetTask] || [];
            const toAdd = relevant.slice(0, 3);
            
            for (const field of toAdd) {
                if (modelContext[field] && !usedFields.includes(field)) {
                    if (this.TABLE_FIELDS.includes(field)) {
                        const formattedTable = this.formatTableForContext(modelContext[field], field);
                        if (formattedTable) {
                            contextParts.push(`${field}:\n${formattedTable}`);
                            usedFields.push(field);
                        }
                    } else {
                        contextParts.push(`${field}: ${modelContext[field]}`);
                        usedFields.push(field);
                    }
                }
            }
            
            // Ajouter 1 champ aléatoire
            const randomFields = availableFields.filter(f => 
                !relevant.includes(f) && 
                !usedFields.includes(f) &&
                modelContext[f]
            );
            
            if (randomFields.length > 0) {
                const randomField = randomFields[Math.floor(Math.random() * randomFields.length)];
                if (this.TABLE_FIELDS.includes(randomField)) {
                    const formattedTable = this.formatTableForContext(modelContext[randomField], randomField);
                    if (formattedTable) {
                        contextParts.push(`${randomField}:\n${formattedTable}`);
                        usedFields.push(randomField);
                    }
                } else {
                    contextParts.push(`${randomField}: ${modelContext[randomField]}`);
                    usedFields.push(randomField);
                }
            }
            
        } else { // full
            // Ajouter 2 dépendances pertinentes + 3-5 champs aléatoires
            const relevant = this.RELEVANT_CONTEXT[targetTask] || [];
            const toAdd = relevant.slice(0, 2);
            
            for (const field of toAdd) {
                if (modelContext[field] && !usedFields.includes(field)) {
                    if (this.TABLE_FIELDS.includes(field)) {
                        const formattedTable = this.formatTableForContext(modelContext[field], field);
                        if (formattedTable) {
                            contextParts.push(`${field}:\n${formattedTable}`);
                            usedFields.push(field);
                        }
                    } else {
                        contextParts.push(`${field}: ${modelContext[field]}`);
                        usedFields.push(field);
                    }
                }
            }
            
            // Ajouter 3-5 champs aléatoires
            const randomFields = availableFields.filter(f => 
                !relevant.includes(f) && 
                !usedFields.includes(f) &&
                modelContext[f]
            );
            
            // Mélanger et prendre 3-5 champs
            const shuffled = [...randomFields].sort(() => Math.random() - 0.5);
            const numToAdd = Math.floor(Math.random() * 3) + 3; // 3-5
            
            for (const field of shuffled.slice(0, numToAdd)) {
                if (this.TABLE_FIELDS.includes(field)) {
                    const formattedTable = this.formatTableForContext(modelContext[field], field);
                    if (formattedTable) {
                        contextParts.push(`${field}:\n${formattedTable}`);
                        usedFields.push(field);
                    }
                } else {
                    contextParts.push(`${field}: ${modelContext[field]}`);
                    usedFields.push(field);
                }
            }
        }
        
        return {
            contextString: contextParts.join('\n'),
            usedFields: usedFields
        };
    }
    
    /**
     * Mappe le nom du champ frontend vers la tâche du modèle
     */
    mapFieldToTask(fieldName) {
        const mapping = {
            'contexte': 'contexte_proj',
            'lots': 'phases'
        };
        
        return mapping[fieldName] || fieldName;
    }
    
    /**
     * Génère du contenu pour un champ unique
     */
    async generateField(fieldName) {
        if (this.isGenerating) {
            showNotification('Une génération est déjà en cours', 'warning');
            return null;
        }
        
        // Vérifier le cache
        const cacheKey = this.getCacheKey(fieldName);
        if (this.generationCache.has(cacheKey)) {
            const cached = this.generationCache.get(cacheKey);
            if (Date.now() - cached.timestamp < 300000) { // Cache 5 minutes
                showNotification('✨ Utilisation du cache', 'info');
                store.setFieldValue(fieldName, cached.value);
                return cached.value;
            }
        }
        
        const startTime = Date.now();
        this.isGenerating = true;
        
        try {
            // Collecter le contexte complet
            const context = this.collectContext(fieldName);
            const strategy = store.state.generation.currentStrategy;
            
            // Afficher l'état de chargement
            this.setFieldLoading(fieldName, true);
            
            // Logger le contexte en dev
            console.log(`🤖 Génération ${fieldName} avec stratégie: ${strategy}`);
            
            // Appeler l'API
            const response = await api.generateField(fieldName, context, strategy);
            
            if (response.success && response.value) {
                // Mettre à jour la valeur
                store.setFieldValue(fieldName, response.value);
                
                // Mettre en cache
                this.generationCache.set(cacheKey, {
                    value: response.value,
                    timestamp: Date.now()
                });
                
                // Statistiques
                const timeTaken = Date.now() - startTime;
                store.updateGenerationStats(fieldName, timeTaken);
                
                // Afficher le feedback
                this.showGenerationFeedback(fieldName, response.value);
                
                // Avertissements de validation
                if (response.validation?.warnings?.length > 0) {
                    setTimeout(() => {
                        showNotification(`⚠️ ${response.validation.warnings.join(', ')}`, 'warning');
                    }, 1000);
                }
                
                // Log les métadonnées si disponibles
                if (response.metadata) {
                    console.log(`📊 Métadonnées de génération:`, response.metadata);
                }
                
                showNotification('✨ Génération réussie!', 'success');
                return response.value;
            } else {
                throw new Error(response.error || 'Échec de la génération');
            }
        } catch (error) {
            console.error('Erreur génération:', error);
            showNotification(`❌ ${error.message}`, 'error');
            return null;
        } finally {
            this.isGenerating = false;
            this.setFieldLoading(fieldName, false);
        }
    }
    
    /**
     * Génère du contenu pour un champ tableau
     */
    async generateArray(arrayName) {
        if (this.isGenerating) {
            showNotification('Une génération est déjà en cours', 'warning');
            return null;
        }
        
        const startTime = Date.now();
        this.isGenerating = true;
        
        try {
            const context = this.collectContext(arrayName);
            const strategy = store.state.generation.currentStrategy;
            
            this.setArrayLoading(arrayName, true);
            
            console.log(`🤖 Génération tableau ${arrayName} avec stratégie: ${strategy}`);
            
            const response = await api.generateField(arrayName, context, strategy);
            
            if (response.success && Array.isArray(response.value)) {
                // Validation et nettoyage des données
                let processedValue = this.validateArrayData(arrayName, response.value);
                
                // Remplacer complètement le tableau existant
                store.setState({
                    project: {
                        ...store.state.project,
                        [arrayName]: processedValue
                    }
                });
                
                // Stats et feedback
                const timeTaken = Date.now() - startTime;
                store.updateGenerationStats(arrayName, timeTaken);
                this.showGenerationFeedback(arrayName, processedValue);
                
                // Message de succès personnalisé
                const itemName = this.getArrayItemName(arrayName);
                showNotification(`✨ ${processedValue.length} ${itemName} générés!`, 'success');
                
                return processedValue;
            } else {
                throw new Error(response.error || 'Échec de la génération');
            }
        } catch (error) {
            console.error('Erreur génération tableau:', error);
            showNotification(`❌ ${error.message}`, 'error');
            return null;
        } finally {
            this.isGenerating = false;
            this.setArrayLoading(arrayName, false);
        }
    }
    
    /**
     * Génère plusieurs champs en une seule requête
     */
    async generateMultipleFields(fields) {
        if (this.isGenerating) {
            showNotification('Une génération est déjà en cours', 'warning');
            return null;
        }
        
        this.isGenerating = true;
        const results = { success: [], failed: [] };
        
        try {
            // Filtrer les champs déjà remplis si stratégie intelligente
            const fieldsToGenerate = fields.filter(field => {
                const value = store.getFieldValue(field);
                return !value || value.toString().trim() === '';
            });
            
            if (fieldsToGenerate.length === 0) {
                showNotification('Tous les champs sont déjà remplis', 'info');
                return results;
            }
            
            const context = this.collectContext();
            const strategy = store.state.generation.currentStrategy;
            
            // Notification de progression
            const progressId = showNotification(
                `Génération de ${fieldsToGenerate.length} champs...`, 
                'info', 
                0
            );
            
            const response = await api.generateMultipleFields(
                fieldsToGenerate, 
                context, 
                strategy
            );
            
            if (response.results) {
                Object.entries(response.results).forEach(([field, value]) => {
                    store.setFieldValue(field, value);
                    results.success.push(field);
                });
            }
            
            if (response.errors) {
                Object.entries(response.errors).forEach(([field, error]) => {
                    results.failed.push({ field, error });
                });
            }
            
            // Supprimer la notification de progression
            store.removeNotification(progressId);
            
            // Résumé
            if (results.success.length > 0) {
                showNotification(
                    `✅ ${results.success.length} champs générés avec succès`,
                    'success'
                );
            }
            
            if (results.failed.length > 0) {
                showNotification(
                    `❌ ${results.failed.length} échecs de génération`,
                    'error'
                );
            }
            
            return results;
        } catch (error) {
            console.error('Erreur génération multiple:', error);
            showNotification(`❌ ${error.message}`, 'error');
            return results;
        } finally {
            this.isGenerating = false;
        }
    }
    
    /**
     * Valide et nettoie les données de tableau
     */
    validateArrayData(arrayName, data) {
        // S'assurer que chaque élément a toutes les colonnes requises
        const structure = Config.ARRAY_STRUCTURES[arrayName];
        if (!structure) return data;
        
        return data.map((item, index) => {
            const validatedItem = {};
            
            structure.cols.forEach((col, colIndex) => {
                let value = item[col] || '';
                
                // Validation selon le type
                const type = structure.types[colIndex];
                
                if (type === 'number') {
                    // S'assurer que c'est un nombre valide
                    const num = parseFloat(value) || 0;
                    value = num.toString();
                } else if (type === 'calculated') {
                    // Recalculer les champs calculés
                    if (arrayName === 'coutsConstruction' && col === 'col3') {
                        const jours = parseFloat(item.col1) || 0;
                        const tjm = parseFloat(item.col2) || 0;
                        value = (jours * tjm).toString();
                    } else if (arrayName === 'coutsFonctionnement' && col === 'col3') {
                        const qty = parseFloat(item.col1) || 0;
                        const monthly = parseFloat(item.col2) || 0;
                        value = (qty * monthly * 12).toString();
                    }
                } else if (type === 'date' && !value) {
                    // Date par défaut si manquante
                    const futureDate = new Date();
                    futureDate.setMonth(futureDate.getMonth() + index + 1);
                    value = futureDate.toISOString().split('T')[0];
                }
                
                // Appliquer les valeurs par défaut
                if (!value && structure.defaults && structure.defaults[col]) {
                    value = structure.defaults[col];
                }
                
                validatedItem[col] = value;
            });
            
            return validatedItem;
        });
    }
    
    /**
     * Affiche le feedback de génération
     */
    showGenerationFeedback(fieldName, value) {
        // Nettoyer les timeouts existants
        if (this.feedbackTimeouts.has(fieldName)) {
            clearTimeout(this.feedbackTimeouts.get(fieldName));
        }
        
        // Créer l'élément de feedback
        const feedbackId = `feedback-${fieldName}`;
        let feedbackEl = document.getElementById(feedbackId);
        
        if (!feedbackEl) {
            const fieldId = getFieldHtmlId(fieldName);
            const fieldEl = document.getElementById(fieldId);
            
            if (fieldEl) {
                const formGroup = fieldEl.closest('.form-group');
                if (formGroup) {
                    feedbackEl = document.createElement('div');
                    feedbackEl.id = feedbackId;
                    feedbackEl.className = 'generation-feedback';
                    formGroup.appendChild(feedbackEl);
                }
            } else {
                // Pour les tableaux
                const arrayContainer = document.querySelector(`#${fieldName}-container`);
                if (arrayContainer) {
                    feedbackEl = document.createElement('div');
                    feedbackEl.id = feedbackId;
                    feedbackEl.className = 'generation-feedback';
                    arrayContainer.appendChild(feedbackEl);
                }
            }
        }
        
        if (feedbackEl) {
            feedbackEl.innerHTML = this.createFeedbackHTML(fieldName, value);
            feedbackEl.classList.add('show');
            
            // Auto-masquer après un délai
            const timeout = setTimeout(() => {
                feedbackEl.classList.remove('show');
                this.feedbackTimeouts.delete(fieldName);
            }, Config.UI.FEEDBACK_TIMEOUT);
            
            this.feedbackTimeouts.set(fieldName, timeout);
        }
    }
    
    /**
     * Crée le HTML du feedback
     */
    createFeedbackHTML(fieldName, value) {
        const isArray = Array.isArray(value);
        const quality = isArray ? 
            `${value.length} éléments générés` : 
            `${value.toString().length} caractères`;
        
        return `
            <div class="feedback-header">
                <span class="feedback-quality">✨ ${quality}</span>
            </div>
            <div class="feedback-actions">
                <button class="feedback-btn" onclick="window.rateGeneration('${fieldName}', 5, this)">
                    😍 Parfait
                </button>
                <button class="feedback-btn" onclick="window.rateGeneration('${fieldName}', 4, this)">
                    😊 Bien
                </button>
                <button class="feedback-btn" onclick="window.rateGeneration('${fieldName}', 3, this)">
                    😐 Moyen
                </button>
                <button class="feedback-btn" onclick="window.rateGeneration('${fieldName}', 2, this)">
                    😕 À améliorer
                </button>
                <button class="feedback-btn feedback-btn-correct" onclick="window.openCorrectionModal('${fieldName}')">
                    ✏️ Corriger
                </button>
            </div>
            <div id="feedback-message-${fieldName}" class="feedback-message"></div>
        `;
    }
    
    /**
     * Note une génération
     */
    async rateGeneration(fieldName, rating, buttonElement) {
        try {
            const value = store.getFieldValue(fieldName);
            const context = this.collectContext();
            
            const response = await api.submitRating(fieldName, value, rating, context);
            
            if (response.success) {
                // Mettre à jour l'UI
                if (buttonElement) {
                    const container = buttonElement.closest('.feedback-actions');
                    container.querySelectorAll('.feedback-btn').forEach(btn => {
                        btn.classList.remove('selected');
                    });
                    buttonElement.classList.add('selected');
                }
                
                // Mettre à jour les statistiques
                store.updateGenerationStats(fieldName, 0, rating);
                
                // Message de remerciement
                const messageEl = document.getElementById(`feedback-message-${fieldName}`);
                if (messageEl) {
                    const messages = {
                        5: 'Merci ! L\'IA apprend de vos retours positifs 🎯',
                        4: 'Merci ! Nous continuons à améliorer la qualité 📈',
                        3: 'Merci ! Votre feedback nous aide à progresser 🔧',
                        2: 'Merci ! Nous allons améliorer ce type de génération 🛠️',
                        1: 'Merci ! Utilisez "Corriger" pour nous montrer l\'idéal 📝'
                    };
                    messageEl.textContent = messages[rating] || 'Merci pour votre retour !';
                }
                
                showNotification('✅ Note enregistrée', 'success');
            }
        } catch (error) {
            console.error('Erreur notation:', error);
            showNotification('❌ Erreur lors de l\'envoi de la note', 'error');
        }
    }
    
    /**
     * Soumet une correction
     */
    async submitCorrection(fieldName, original, corrected, reason) {
        try {
            const context = this.collectContext();
            
            const response = await api.submitCorrection(
                fieldName, 
                original, 
                corrected, 
                context, 
                reason
            );
            
            if (response.success) {
                // Mettre à jour la valeur
                store.setFieldValue(fieldName, corrected);
                
                // Invalider le cache pour ce champ
                const cacheKey = this.getCacheKey(fieldName);
                this.generationCache.delete(cacheKey);
                
                // Afficher les insights
                if (response.insights?.recommendations) {
                    const recommendations = response.insights.recommendations.join('\n• ');
                    showNotification(`💡 L'IA a appris:\n• ${recommendations}`, 'info', 5000);
                } else {
                    showNotification('✅ Correction enregistrée', 'success');
                }
                
                return true;
            }
            
            return false;
        } catch (error) {
            console.error('Erreur correction:', error);
            showNotification('❌ Erreur lors de l\'envoi de la correction', 'error');
            return false;
        }
    }
    
    /**
     * Génère une clé de cache basée sur le contexte
     */
    getCacheKey(fieldName) {
        const context = this.collectContext(fieldName);
        // Utiliser uniquement les champs de base pour la clé
        const contextKey = JSON.stringify({
            field: fieldName,
            client: context.client,
            typeProjet: context.typeProjet,
            complexite: context.complexite
        });
        return `${fieldName}_${btoa(contextKey)}`;
    }
    
    /**
     * Obtient le nom d'un élément de tableau
     */
    getArrayItemName(arrayName) {
        const names = {
            contraintes: 'contraintes',
            risques: 'risques',
            lots: 'lots',
            phases: 'phases',
            livrables: 'livrables',
            jalons: 'jalons',
            coutsConstruction: 'lignes de coût',
            coutsFonctionnement: 'postes de coût'
        };
        
        return names[arrayName] || 'éléments';
    }
    
    /**
     * UI state helpers
     */
    setFieldLoading(fieldName, loading) {
        const fieldId = getFieldHtmlId(fieldName);
        const fieldEl = document.getElementById(fieldId);
        
        if (fieldEl) {
            const button = fieldEl.parentElement.querySelector('.btn-ai');
            if (button) {
                if (loading) {
                    button.disabled = true;
                    button.innerHTML = '<span class="spinner"></span> Génération...';
                } else {
                    button.disabled = false;
                    button.innerHTML = '🤖 Générer';
                }
            }
        }
    }
    
    setArrayLoading(arrayName, loading) {
        const button = document.querySelector(`#${arrayName}-generate`);
        if (button) {
            if (loading) {
                button.disabled = true;
                button.innerHTML = '<span class="spinner"></span> Génération...';
            } else {
                button.disabled = false;
                button.innerHTML = '🤖 Générer';
            }
        }
    }
    
    /**
     * Méthodes de nettoyage
     */
    destroy() {
        // Nettoyer les timeouts
        this.feedbackTimeouts.forEach(timeout => clearTimeout(timeout));
        this.feedbackTimeouts.clear();
        
        // Vider le cache
        this.generationCache.clear();
        
        this.isGenerating = false;
    }
}

// Créer l'instance singleton
export const generationService = new GenerationService();

// Exporter les fonctions globales pour les handlers onclick HTML
window.generateField = (fieldName) => generationService.generateField(fieldName);
window.generateArray = (arrayName) => generationService.generateArray(arrayName);
window.rateGeneration = (fieldName, rating, button) => generationService.rateGeneration(fieldName, rating, button);
