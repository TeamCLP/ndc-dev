// js/core/Api.js
import { Config } from '../config.js';

/**
 * ApiClient - Client pour l'API SCRIBE AI v3.1
 * 
 * PARFAITEMENT ALIGNÉ avec le modèle entraîné
 * 
 * Mapping automatique des champs :
 * - sector -> secteur
 * - project_type -> typeProjet  
 * - complexity -> complexite
 * - contexte -> contexte_proj
 * - lots -> phases
 * 
 * Nouveaux champs supportés :
 * - impactCO2
 * - modalitesPartage
 */
export class ApiClient {
    constructor() {
        this.baseUrl = Config.API.BASE_URL;
        this.timeout = null; // Désactivé - pas de timeout
        this.headers = {
            'Content-Type': 'application/json'
        };
        
        // Liste des champs supportés
        this.supportedFields = {
            text: [
                'contexte', 'besoin', 'objectifs', 'perimetre', 'horsPerimetre',
                'descriptionSolution', 'architecture', 'composantsDimensionnement',
                'conditionsHorsCrash', 'conditionsCrashSite', 'resilienceApplicative',
                'praPlanDegrade', 'sauvegardes', 'administrationSupervision',
                'impactCO2', 'modalitesPartage'
            ],
            table: [
                'contraintes', 'risques', 'lots', 'livrables', 'jalons',
                'coutsConstruction', 'coutsFonctionnement'
            ]
        };
        
        // Stratégies disponibles (mappées vers les complexités du modèle)
        this.strategies = {
            minimal: 'minimal',    // Juste les champs de base
            related: 'related',    // Champs pertinents (par défaut)
            smart: 'related',      // Alias pour related
            full: 'full',         // Tous les champs disponibles
            debug: 'full'         // Alias pour full
        };
    }
    
    // Update API URL
    setBaseUrl(url) {
        this.baseUrl = url;
        localStorage.setItem(Config.STORAGE.API_URL, url);
    }
    
    // Generic request method - SANS TIMEOUT
    async request(endpoint, options = {}) {
        const url = `${this.baseUrl}${endpoint}`;
        const config = {
            ...options,
            headers: {
                ...this.headers,
                ...options.headers
            }
        };
        
        try {
            const response = await fetch(url, config);
            
            if (!response.ok) {
                const error = await this.parseError(response);
                throw new ApiError(error.message || `HTTP ${response.status}`, response.status, error);
            }
            
            return await response.json();
        } catch (error) {
            if (error instanceof ApiError) {
                throw error;
            }
            
            throw new ApiError('Network error', 0, error);
        }
    }
    
    async parseError(response) {
        try {
            return await response.json();
        } catch {
            return { message: `HTTP ${response.status}` };
        }
    }
    
    // Health check
    async checkHealth() {
        return this.request(Config.API.ENDPOINTS.HEALTH);
    }
    
    /**
     * Prépare le contexte pour l'API
     * Inclut automatiquement les champs de base toujours nécessaires
     */
    prepareContext(store) {
        const project = store.getProjectData();
        const context = {};
        
        // Champs de base toujours inclus (mapping automatique)
        context.client = project.client || '';
        context.sector = project.sector || '';
        context.project_type = project.project_type || '';
        context.complexity = project.complexity || '';
        context.libelle = project.libelle || '';
        context.annee = project.annee || new Date().getFullYear().toString();
        
        // Autres champs potentiellement utiles
        const otherFields = [
            'numeroProjet', 'trigramme', 'contexte', 'besoin', 'objectifs',
            'perimetre', 'horsPerimetre', 'descriptionSolution', 'architecture',
            'composantsDimensionnement', 'plageService', 'dicp', 'dima', 'pdma',
            'rgpd', 'psee', 'tauxContingence', 'contraintes', 'risques', 'lots',
            'livrables', 'jalons', 'coutsConstruction', 'coutsFonctionnement',
            'conditionsHorsCrash', 'conditionsCrashSite', 'resilienceApplicative',
            'praPlanDegrade', 'sauvegardes', 'administrationSupervision',
            'impactCO2', 'modalitesPartage', 'lienDocumentation'
        ];
        
        // Ajouter les autres champs s'ils existent
        otherFields.forEach(field => {
            if (project[field] !== undefined && project[field] !== '') {
                context[field] = project[field];
            }
        });
        
        return context;
    }
    
    // Generation endpoints
    async generateField(field, context, strategy = 'related', options = {}) {
        // Validation du champ
        const allFields = [...this.supportedFields.text, ...this.supportedFields.table];
        if (!allFields.includes(field)) {
            console.warn(`Field '${field}' might not be supported by the model`);
        }
        
        // Normaliser la stratégie
        const normalizedStrategy = this.strategies[strategy] || strategy;
        
        return this.request(Config.API.ENDPOINTS.GENERATE, {
            method: 'POST',
            body: JSON.stringify({
                field,
                context,
                strategy: normalizedStrategy,
                ...options
            })
        });
    }
    
    async generateMultipleFields(fields, context, strategy = 'related') {
        // Validation des champs
        const unsupportedFields = fields.filter(f => {
            const allFields = [...this.supportedFields.text, ...this.supportedFields.table];
            return !allFields.includes(f);
        });
        
        if (unsupportedFields.length > 0) {
            console.warn(`Unsupported fields: ${unsupportedFields.join(', ')}`);
        }
        
        // Normaliser la stratégie
        const normalizedStrategy = this.strategies[strategy] || strategy;
        
        return this.request(Config.API.ENDPOINTS.GENERATE_MULTIPLE, {
            method: 'POST',
            body: JSON.stringify({
                fields,
                context,
                strategy: normalizedStrategy
            })
        });
    }
    
    // Validation
    async validateField(field, value, context = {}) {
        return this.request(Config.API.ENDPOINTS.VALIDATE, {
            method: 'POST',
            body: JSON.stringify({
                field,
                value,
                context
            })
        });
    }
    
    // Feedback endpoints
    async submitCorrection(field, original, corrected, context, reason = null) {
        return this.request(Config.API.ENDPOINTS.FEEDBACK_CORRECTION, {
            method: 'POST',
            body: JSON.stringify({
                field,
                original,
                corrected,
                context,
                reason
            })
        });
    }
    
    async submitRating(field, value, rating, context, comment = null) {
        return this.request(Config.API.ENDPOINTS.FEEDBACK_RATING, {
            method: 'POST',
            body: JSON.stringify({
                field,
                value,
                rating,
                context,
                comment
            })
        });
    }
    
    async getFieldInsights(field) {
        return this.request(`${Config.API.ENDPOINTS.FEEDBACK_INSIGHTS}/${field}`);
    }
    
    async getFeedbackSummary() {
        return this.request(Config.API.ENDPOINTS.FEEDBACK_SUMMARY);
    }
    
    // Helper methods pour les nouveaux champs
    
    /**
     * Génère l'impact CO2 du projet
     * @param {Object} context - Contexte du projet
     * @returns {Promise<Object>} Réponse de l'API
     */
    async generateImpactCO2(context) {
        return this.generateField('impactCO2', context, 'related');
    }
    
    /**
     * Génère les modalités de partage
     * @param {Object} context - Contexte du projet incluant rgpd, psee, etc.
     * @returns {Promise<Object>} Réponse de l'API
     */
    async generateModalitesPartage(context) {
        return this.generateField('modalitesPartage', context, 'related');
    }
    
    /**
     * Génère tous les champs de base d'un projet
     * @param {Object} context - Contexte initial du projet
     * @returns {Promise<Object>} Résultats de génération
     */
    async generateProjectBasics(context) {
        const basicFields = ['contexte', 'besoin', 'objectifs', 'perimetre'];
        return this.generateMultipleFields(basicFields, context, 'related');
    }
    
    /**
     * Génère les contraintes et risques
     * @param {Object} context - Contexte du projet
     * @returns {Promise<Object>} Résultats de génération
     */
    async generateConstraintsAndRisks(context) {
        const fields = ['contraintes', 'risques'];
        return this.generateMultipleFields(fields, context, 'related');
    }
    
    /**
     * Génère la solution technique complète
     * @param {Object} context - Contexte du projet
     * @returns {Promise<Object>} Résultats de génération
     */
    async generateTechnicalSolution(context) {
        const fields = ['descriptionSolution', 'architecture', 'composantsDimensionnement'];
        return this.generateMultipleFields(fields, context, 'related');
    }
    
    /**
     * Génère le planning complet (phases, jalons, livrables)
     * @param {Object} context - Contexte du projet
     * @returns {Promise<Object>} Résultats de génération
     */
    async generatePlanning(context) {
        const fields = ['lots', 'jalons', 'livrables'];
        return this.generateMultipleFields(fields, context, 'related');
    }
    
    /**
     * Génère les coûts complets
     * @param {Object} context - Contexte du projet
     * @returns {Promise<Object>} Résultats de génération
     */
    async generateCosts(context) {
        const fields = ['coutsConstruction', 'coutsFonctionnement'];
        return this.generateMultipleFields(fields, context, 'related');
    }
    
    // Méthodes utilitaires
    
    /**
     * Vérifie si un champ est un tableau
     */
    isTableField(field) {
        return this.supportedFields.table.includes(field);
    }
    
    /**
     * Vérifie si un champ est supporté
     */
    isFieldSupported(field) {
        const allFields = [...this.supportedFields.text, ...this.supportedFields.table];
        return allFields.includes(field);
    }
    
    /**
     * Retourne la liste de tous les champs supportés
     */
    getAllSupportedFields() {
        return [...this.supportedFields.text, ...this.supportedFields.table];
    }
    
    /**
     * Retourne les métadonnées d'un champ
     */
    getFieldMetadata(field) {
        const metadata = {
            isTable: this.isTableField(field),
            isSupported: this.isFieldSupported(field),
            category: this.supportedFields.text.includes(field) ? 'text' : 
                     this.supportedFields.table.includes(field) ? 'table' : 'unknown'
        };
        
        // Ajouter des infos spécifiques selon le champ
        if (field === 'lots') {
            metadata.note = 'Mappé vers "phases" pour le modèle';
        } else if (field === 'contexte') {
            metadata.note = 'Mappé vers "contexte_proj" pour le modèle';
        }
        
        return metadata;
    }
}

// Custom error class
export class ApiError extends Error {
    constructor(message, status = 0, details = null) {
        super(message);
        this.name = 'ApiError';
        this.status = status;
        this.details = details;
    }
    
    isNetworkError() {
        return this.status === 0;
    }
    
    isTimeout() {
        return this.status === 408;
    }
    
    isServerError() {
        return this.status >= 500;
    }
    
    isClientError() {
        return this.status >= 400 && this.status < 500;
    }
    
    isModelNotLoaded() {
        return this.status === 503;
    }
}

// Create singleton instance
export const api = new ApiClient();
