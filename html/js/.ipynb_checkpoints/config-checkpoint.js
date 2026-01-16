// Détection automatique du préfixe
const detectPrefix = () => {
    const match = window.location.pathname.match(/^(\/scribe-ai\/[^\/]+\/url-1)/);
    return match ? match[1] : "/scribe-ai/test/url-1";
};
const DYNAMIC_PREFIX = detectPrefix();

// js/config.js
export const Config = {
    // API Configuration
    API: {
        BASE_URL: `https://runai-poc-100.mlops.cagip.group.gca${DYNAMIC_PREFIX}/api`,
        ENDPOINTS: {
            HEALTH: '/health',
            GENERATE: '/generate',
            GENERATE_MULTIPLE: '/generate_multiple',
            VALIDATE: '/validate',
            FEEDBACK_CORRECTION: '/feedback/correction',
            FEEDBACK_RATING: '/feedback/rating',
            FEEDBACK_INSIGHTS: '/feedback/insights',
            FEEDBACK_SUMMARY: '/feedback/summary'
        },
        TIMEOUT: 30000
    },
    
    // Supported Fields
    FIELDS: {
        TEXT_SHORT: ['client', 'libelle', 'trigramme', 'typeProjet', 'plageService'],
        TEXT_LONG: [
            'contexte', 'besoin', 'objectifs', 'perimetre', 'horsPerimetre',
            'descriptionSolution', 'architecture', 'composantsDimensionnement',
            'conditionsHorsCrash', 'conditionsCrashSite', 'resilienceApplicative',
            'praPlanDegrade', 'sauvegardes', 'administrationSupervision'
        ],
        ENUM: ['sector', 'project_type', 'complexity', 'typeDevis', 'besoinPOC'],
        NUMERIC: ['annee', 'numeroProjet', 'numeroDevis', 'tauxContingence'],
        ARRAY: [
            'contraintes', 'risques', 'lots', 'livrables', 'jalons',
            'coutsConstruction', 'coutsFonctionnement'
        ]
    },
    
    // Field Mappings
    FIELD_ID_MAP: {
        'perimetre': 'perimetre-field'
    },
    
    // Array Structures
    ARRAY_STRUCTURES: {
        contraintes: {
            cols: ['col0', 'col1', 'col2', 'col3'],
            labels: ['Type', 'Description', 'Criticité', 'Mitigation'],
            types: ['text', 'text', 'enum', 'text'],
            enums: {
                col0: ['TECHNIQUE', 'ORGANISATIONNELLE', 'REGLEMENTAIRE', 'TEMPORELLE', 'BUDGETAIRE'],
                col2: ['Faible', 'Moyen', 'Élevé']
            }
        },
        risques: {
            cols: ['col0', 'col1', 'col2', 'col3', 'col4'],
            labels: ['Risque', 'Probabilité', 'Impact', 'Plan d\'action', 'Responsable'],
            types: ['text', 'enum', 'enum', 'text', 'text'],
            enums: {
                col1: ['Faible', 'Moyenne', 'Élevée'],
                col2: ['Faible', 'Moyen', 'Élevé']
            }
        },
        lots: {
            cols: ['col0', 'col1', 'col2', 'col3', 'col4', 'col5'],
            labels: ['Nom', 'Description', 'Durée (jours)', 'Date début', 'Date fin', 'Ressources'],
            types: ['text', 'text', 'number', 'date', 'date', 'text']
        },
        livrables: {
            cols: ['col0', 'col1', 'col2', 'col3'],
            labels: ['Nom', 'Description', 'Date', 'Responsable'],
            types: ['text', 'text', 'date', 'text'],
            defaults: { col3: 'NOTRE ENTREPRISE' }
        },
        jalons: {
            cols: ['col0', 'col1', 'col2', 'col3'],
            labels: ['Nom', 'Type', 'Date', 'Critères'],
            types: ['text', 'text', 'date', 'text']
        },
        coutsConstruction: {
            cols: ['col0', 'col1', 'col2', 'col3', 'col4'],
            labels: ['Profil', 'Charge (j.h)', 'TJM (€)', 'Total HT (€)', 'Code équipe'],
            types: ['text', 'number', 'number', 'calculated', 'text'], // Ajout de col4
            // Les suggestions peuvent toujours être utilisées par l'API
            suggestions: [
                'Chef de projet', 'Architecte solution', 'Développeur senior',
                'Développeur', 'Analyste', 'Expert technique', 'Testeur'
            ],
            tjmRanges: {
                'Chef de projet': [850, 1200],
                'Architecte solution': [900, 1300],
                'Développeur senior': [750, 1000],
                'Développeur': [600, 850],
                'Analyste': [700, 950],
                'Expert technique': [800, 1100],
                'Testeur': [550, 750]
            }
        },
        coutsFonctionnement: {
            cols: ['col0', 'col1', 'col2', 'col3', 'col4'],
            labels: ['Poste', 'Quantité', 'Prix unitaire/an (€)', 'Total annuel (€)', 'Code UO'],
            types: ['text', 'number', 'number', 'calculated', 'text'], // Ajout de col4
            defaults: { col1: '1' },
            // Indicateur pour la conversion mensuel->annuel
            monthlyToAnnualConversion: true
        }
    },
    
    // Required Fields - CORRECTION: utiliser 'perimetre' au lieu de 'perimetre-field'
    REQUIRED_FIELDS: [
        'client', 'sector', 'project_type', 'complexity', 'libelle',
        'annee', 'numeroProjet', 'trigramme', 'contexte', 'besoin',
        'objectifs', 'perimetre', 'horsPerimetre', 'descriptionSolution',
        'architecture', 'composantsDimensionnement', 'plageService',
        'conditionsHorsCrash', 'tauxContingence'
    ],
    
    // Sections Configuration
    SECTIONS: [
        {
            id: 'informations-generales',
            title: 'Informations Générales',
            icon: '📋',
            fields: ['client', 'sector', 'project_type', 'complexity', 'libelle', 
                    'annee', 'numeroProjet', 'numeroDevis', 'trigramme', 'typeDevis', 'typeProjet']
        },
        {
            id: 'description-projet',
            title: 'I. Description du Projet',
            icon: '📝',
            fields: ['contexte', 'besoin', 'objectifs']
        },
        {
            id: 'perimetre',
            title: 'II. Périmètre',
            icon: '🎯',
            fields: ['perimetre', 'horsPerimetre']
        },
        {
            id: 'securite',
            title: 'III. Sécurité & Réglementation',
            icon: '🔒',
            fields: ['dicp', 'dima', 'pdma', 'rgpd', 'psee']
        },
        {
            id: 'contraintes-risques',
            title: 'IV. Contraintes & Risques',
            icon: '⚠️',
            fields: ['contraintes', 'risques']
        },
        {
            id: 'solution-technique',
            title: 'V. Solution Technique',
            icon: '🔧',
            fields: ['descriptionSolution', 'besoinPOC', 'architecture', 'composantsDimensionnement']
        },
        {
            id: 'mise-en-oeuvre',
            title: 'VI. Mise en Œuvre',
            icon: '🚀',
            fields: ['lots', 'livrables', 'jalons']
        },
        {
            id: 'service-fonctionnement',
            title: 'VII. Service en Fonctionnement',
            icon: '⚙️',
            fields: ['plageService', 'conditionsHorsCrash', 'conditionsCrashSite',
                    'resilienceApplicative', 'praPlanDegrade', 'sauvegardes', 'administrationSupervision']
        },
        {
            id: 'evaluation-financiere',
            title: 'VIII. Évaluation Financière',
            icon: '💰',
            fields: ['tauxContingence', 'coutsConstruction', 'coutsFonctionnement']
        },
        {
            id: 'rse',
            title: 'IX. RSE - Impact CO2',
            icon: '🌱',
            fields: ['impactCO2']
        },
        {
            id: 'documentation',
            title: 'X. Documentation',
            icon: '📚',
            fields: ['modalitesPartage', 'lienDocumentation']
        }
    ],
    
    // UI Configuration
    UI: {
        NOTIFICATION_DURATION: 3000,
        FEEDBACK_TIMEOUT: 30000,
        DEBOUNCE_DELAY: 500,
        ANIMATION_DURATION: 300
    },
    
    // Storage Keys
    STORAGE: {
        CURRENT_PROJECT: 'scribe_current_project',
        API_URL: 'scribe_api_url',
        GENERATION_STATS: 'scribe_generation_stats',
        USER_PREFERENCES: 'scribe_user_preferences'
    }
};

// Enum Values
export const ENUM_VALUES = {
    sector: ['Banque', 'Assurance', 'Industrie', 'Service', 'Public'],
    project_type: [
        'Conformité Réglementaire',
        'Migration Cloud',
        'Transformation Digitale',
        'Modernisation Legacy',
        'Sécurisation SI',
        'Innovation Data',
        'Automatisation Process',
        'Core Banking',
        'Trading Platform'
    ],
    complexity: ['Simple', 'Standard', 'Complexe', 'Très complexe'],
    typeDevis: ['Etude', 'Realisation', 'Etude_Realisation'],
    besoinPOC: ['oui', 'non'],
    rgpd: ['oui', 'non'],
    psee: ['oui', 'non']
};

// Strategy Options
export const GENERATION_STRATEGIES = {
    smart: { label: 'IA Intelligente', description: 'Utilise le contexte optimal' },
    minimal: { label: 'IA Minimale', description: 'Contexte minimal pour rapidité' },
    related: { label: 'IA Contextuelle', description: 'Utilise les champs liés' }, // Default strategy
    random: { label: 'IA Variée', description: 'Génération créative' }
};

// Export Word Configuration
export const WORD_EXPORT_CONFIG = {
    margins: {
        top: 1134,    // 2cm
        right: 1134,  // 2cm
        bottom: 1134, // 2cm
        left: 1134    // 2cm
    },
    styles: {
        heading1: {
            size: 40,  // 20pt
            bold: true,
            color: "2C3E50",
            font: "Calibri"
        },
        heading2: {
            size: 32,  // 16pt
            bold: true,
            color: "34495E",
            font: "Calibri"
        },
        heading3: {
            size: 28,  // 14pt
            bold: true,
            color: "7F8C8D",
            font: "Calibri"
        },
        normal: {
            size: 22,  // 11pt
            font: "Calibri"
        }
    }
};
