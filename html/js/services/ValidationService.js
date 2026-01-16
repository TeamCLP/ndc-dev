// js/services/ValidationService.js
import { Config, ENUM_VALUES } from '../config.js';
import { store } from '../core/Store.js';
import { api } from '../core/Api.js';
import { showNotification, isValidEmail, isValidUrl } from '../utils/helpers.js';

export class ValidationService {
    constructor() {
        this.validationRules = this.defineValidationRules();
        this.customValidators = this.defineCustomValidators();
    }
    
    defineValidationRules() {
        return {
            // Text fields
            client: {
                required: true,
                minLength: 2,
                maxLength: 100,
                pattern: /^[a-zA-ZÀ-ÿ0-9\s\-&.,]+$/,
                message: 'Le nom du client doit contenir au moins 2 caractères'
            },
            libelle: {
                required: true,
                minLength: 5,
                maxLength: 200,
                message: 'Le libellé du projet doit être descriptif (5-200 caractères)'
            },
            trigramme: {
                required: true,
                pattern: /^[A-Z]{3}$/,
                transform: 'uppercase',
                message: 'Le trigramme doit contenir exactement 3 lettres majuscules'
            },
            
            // Numeric fields
            annee: {
                required: true,
                pattern: /^20\d{2}$/,
                min: 2020,
                max: 2030,
                message: 'L\'année doit être entre 2020 et 2030'
            },
            numeroProjet: {
                required: true,
                pattern: /^\d{6}$/,
                message: 'Le numéro de projet doit contenir 6 chiffres'
            },
            numeroDevis: {
                required: true,
                pattern: /^\d{3}$/,
                message: 'Le numéro de devis doit contenir 3 chiffres'
            },
            tauxContingence: {
                required: true,
                type: 'number',
                min: 10,
                max: 30,
                message: 'Le taux de contingence doit être entre 10% et 30%'
            },
            
            // Long text fields
            contexte: {
                required: true,
                minLength: 50,
                maxLength: 3000,
                message: 'Le contexte doit contenir entre 50 et 3000 caractères'
            },
            besoin: {
                required: true,
                minLength: 30,
                maxLength: 2000,
                message: 'Le besoin doit être clairement exprimé (30-2000 caractères)'
            },
            objectifs: {
                required: true,
                minLength: 50,
                mustInclude: ['business', 'technique'],
                message: 'Les objectifs doivent inclure les aspects business et techniques'
            },
            perimetre: {
                required: true,
                minLength: 50,
                message: 'Le périmètre doit être détaillé (min 50 caractères)'
            },
            horsPerimetre: {
                required: true,
                minLength: 30,
                message: 'Précisez ce qui est hors périmètre (min 30 caractères)'
            },
            
            // Technical fields
            descriptionSolution: {
                required: true,
                minLength: 100,
                message: 'La description de la solution doit être détaillée (min 100 caractères)'
            },
            architecture: {
                required: true,
                minLength: 100,
                message: 'L\'architecture doit être décrite en détail (min 100 caractères)'
            },
            composantsDimensionnement: {
                required: true,
                minLength: 50,
                message: 'Listez les composants et leur dimensionnement'
            },
            
            // Service fields
            plageService: {
                required: true,
                pattern: /\d+j\/\d+|24\/7|7j\/7/,
                message: 'Format attendu: 5j/7 8h-18h ou 24/7'
            },
            conditionsHorsCrash: {
                required: true,
                minLength: 50,
                message: 'Décrivez les conditions normales de fonctionnement'
            },
            
            // Optional fields with validation when filled
            dima: {
                type: 'number',
                min: 0,
                max: 365,
                message: 'DIMA doit être entre 0 et 365 jours'
            },
            pdma: {
                type: 'number',
                min: 0,
                max: 365,
                message: 'PDMA doit être entre 0 et 365 jours'
            },
            lienDocumentation: {
                type: 'url',
                message: 'Le lien doit être une URL valide'
            },
            
            // Enum fields
            sector: {
                required: true,
                enum: ENUM_VALUES.sector,
                message: 'Sélectionnez un secteur d\'activité'
            },
            project_type: {
                required: true,
                enum: ENUM_VALUES.project_type,
                message: 'Sélectionnez un type de projet'
            },
            complexity: {
                required: true,
                enum: ENUM_VALUES.complexity,
                message: 'Évaluez la complexité du projet'
            },
            typeDevis: {
                required: true,
                enum: ENUM_VALUES.typeDevis,
                message: 'Sélectionnez le type de devis'
            }
        };
    }
    
    defineCustomValidators() {
        return {
            // Validate coherence between fields
            checkDateCoherence: (project) => {
                const errors = [];
                
                if (project.lots && project.lots.length > 0) {
                    project.lots.forEach((lot, index) => {
                        if (lot.col3 && lot.col4) {
                            const start = new Date(lot.col3);
                            const end = new Date(lot.col4);
                            
                            if (end < start) {
                                errors.push({
                                    field: `lots[${index}]`,
                                    message: `Lot ${index + 1}: la date de fin est antérieure à la date de début`
                                });
                            }
                        }
                    });
                }
                
                return errors;
            },
            
            // Validate financial calculations
            checkFinancialCoherence: (project) => {
                const errors = [];
                
                if (project.coutsConstruction) {
                    project.coutsConstruction.forEach((cout, index) => {
                        const jours = parseFloat(cout.col1) || 0;
                        const tjm = parseFloat(cout.col2) || 0;
                        const total = parseFloat(cout.col3) || 0;
                        const expected = jours * tjm;
                        
                        if (Math.abs(total - expected) > 1) {
                            errors.push({
                                field: `coutsConstruction[${index}]`,
                                message: `Ligne ${index + 1}: erreur de calcul (${jours} × ${tjm} ≠ ${total})`
                            });
                        }
                    });
                }
                
                return errors;
            },
            
            // Check TJM ranges
            checkTJMRanges: (project) => {
                const warnings = [];
                const tjmRanges = Config.ARRAY_STRUCTURES.coutsConstruction.tjmRanges;
                
                if (project.coutsConstruction) {
                    project.coutsConstruction.forEach((cout, index) => {
                        const profile = cout.col0;
                        const tjm = parseFloat(cout.col2) || 0;
                        
                        if (profile && tjmRanges[profile]) {
                            const [min, max] = tjmRanges[profile];
                            
                            if (tjm < min || tjm > max) {
                                warnings.push({
                                    field: `coutsConstruction[${index}]`,
                                    message: `${profile}: TJM inhabituel (${tjm}€, range habituel: ${min}-${max}€)`
                                });
                            }
                        }
                    });
                }
                
                return warnings;
            },
            
            // Check array minimum items
            checkArrayMinimums: (project) => {
                const warnings = [];
                const minimums = {
                    contraintes: 2,
                    risques: 2,
                    lots: 3,
                    livrables: 3,
                    jalons: 3
                };
                
                Object.entries(minimums).forEach(([field, min]) => {
                    const items = project[field];
                    if (items && items.length < min) {
                        warnings.push({
                            field,
                            message: `Il est recommandé d'avoir au moins ${min} ${field}`
                        });
                    }
                });
                
                return warnings;
            }
        };
    }
    
    // Validate single field
    async validateField(fieldName, value) {
        const rule = this.validationRules[fieldName];
        if (!rule) return { valid: true };
        
        const errors = [];
        const warnings = [];
        
        // Transform value if needed
        if (rule.transform === 'uppercase' && typeof value === 'string') {
            value = value.toUpperCase();
        }
        
        // Required check
        if (rule.required && (!value || value.toString().trim() === '')) {
            errors.push(rule.message || `Le champ ${fieldName} est requis`);
            return { valid: false, errors, warnings };
        }
        
        // Skip other validations if empty and not required
        if (!value || value.toString().trim() === '') {
            return { valid: true, errors, warnings };
        }
        
        // Type validation
        if (rule.type === 'number') {
            const num = parseFloat(value);
            if (isNaN(num)) {
                errors.push('Doit être un nombre');
            } else {
                if (rule.min !== undefined && num < rule.min) {
                    errors.push(`Valeur minimale: ${rule.min}`);
                }
                if (rule.max !== undefined && num > rule.max) {
                    errors.push(`Valeur maximale: ${rule.max}`);
                }
            }
        } else if (rule.type === 'url') {
            if (!isValidUrl(value)) {
                errors.push('URL invalide');
            }
        }
        
        // String validations
        if (typeof value === 'string') {
            if (rule.minLength && value.length < rule.minLength) {
                errors.push(`Minimum ${rule.minLength} caractères`);
            }
            if (rule.maxLength && value.length > rule.maxLength) {
                errors.push(`Maximum ${rule.maxLength} caractères`);
            }
            if (rule.pattern && !rule.pattern.test(value)) {
                errors.push(rule.message || 'Format invalide');
            }
            if (rule.mustInclude) {
                const valueLower = value.toLowerCase();
                const missing = rule.mustInclude.filter(word => !valueLower.includes(word));
                if (missing.length > 0) {
                    warnings.push(`Devrait mentionner: ${missing.join(', ')}`);
                }
            }
        }
        
        // Enum validation
        if (rule.enum && !rule.enum.includes(value)) {
            errors.push(`Valeur invalide. Options: ${rule.enum.join(', ')}`);
        }
        
        return {
            valid: errors.length === 0,
            errors,
            warnings
        };
    }
    
    // Validate entire project
    async validateProject() {
        const project = store.getProjectData();
        const results = {
            valid: true,
            errors: [],
            warnings: [],
            fieldErrors: {},
            fieldWarnings: {}
        };
        
        // Validate individual fields
        for (const [fieldName, rule] of Object.entries(this.validationRules)) {
            const value = project[fieldName];
            const validation = await this.validateField(fieldName, value);
            
            if (!validation.valid) {
                results.valid = false;
                results.fieldErrors[fieldName] = validation.errors;
                results.errors.push(`${fieldName}: ${validation.errors.join(', ')}`);
            }
            
            if (validation.warnings && validation.warnings.length > 0) {
                results.fieldWarnings[fieldName] = validation.warnings;
                results.warnings.push(...validation.warnings);
            }
        }
        
        // Run custom validators
        const dateErrors = this.customValidators.checkDateCoherence(project);
        const financialErrors = this.customValidators.checkFinancialCoherence(project);
        const tjmWarnings = this.customValidators.checkTJMRanges(project);
        const arrayWarnings = this.customValidators.checkArrayMinimums(project);
        
        if (dateErrors.length > 0) {
            results.valid = false;
            results.errors.push(...dateErrors.map(e => e.message));
        }
        
        if (financialErrors.length > 0) {
            results.valid = false;
            results.errors.push(...financialErrors.map(e => e.message));
        }
        
        results.warnings.push(...tjmWarnings.map(w => w.message));
        results.warnings.push(...arrayWarnings.map(w => w.message));
        
        return results;
    }
    
    // Validate before export
    async validateForExport() {
        const validation = await this.validateProject();
        
        if (!validation.valid) {
            const errorMessage = 'Le projet contient des erreurs:\n\n' + 
                                validation.errors.slice(0, 5).join('\n') +
                                (validation.errors.length > 5 ? `\n... et ${validation.errors.length - 5} autres erreurs` : '');
            
            showNotification(errorMessage, 'error', 10000);
            return false;
        }
        
        if (validation.warnings.length > 0) {
            const warningMessage = 'Avertissements:\n' + validation.warnings.slice(0, 3).join('\n');
            showNotification(warningMessage, 'warning', 5000);
        }
        
        return true;
    }
    
    // Real-time field validation
    setupRealtimeValidation() {
        // This would be called to set up real-time validation on form fields
        document.addEventListener('change', async (e) => {
            if (e.target.classList.contains('form-control')) {
                const fieldId = e.target.id;
                const fieldName = this.getFieldNameFromId(fieldId);
                
                if (fieldName && this.validationRules[fieldName]) {
                    const validation = await this.validateField(fieldName, e.target.value);
                    this.updateFieldValidationUI(e.target, validation);
                }
            }
        });
    }
    
    updateFieldValidationUI(fieldElement, validation) {
        // Remove existing validation classes
        fieldElement.classList.remove('has-error', 'has-warning');
        
        // Remove existing error messages
        const formGroup = fieldElement.closest('.form-group');
        if (formGroup) {
            const existingError = formGroup.querySelector('.form-error');
            if (existingError) {
                existingError.remove();
            }
        }
        
        // Add validation feedback
        if (!validation.valid) {
            fieldElement.classList.add('has-error');
            
            if (formGroup && validation.errors.length > 0) {
                const errorDiv = document.createElement('div');
                errorDiv.className = 'form-error';
                errorDiv.textContent = validation.errors[0];
                formGroup.appendChild(errorDiv);
            }
        } else if (validation.warnings && validation.warnings.length > 0) {
            fieldElement.classList.add('has-warning');
        }
    }
    
    getFieldNameFromId(fieldId) {
        // Handle special cases
        if (fieldId === 'perimetre-field') return 'perimetre';
        return fieldId;
    }
    
    // Validate with API
    async validateFieldWithAPI(fieldName, value, context = {}) {
        try {
            const response = await api.validateField(fieldName, value, context);
            return response;
        } catch (error) {
            console.error('API validation error:', error);
            // Fallback to local validation
            return this.validateField(fieldName, value);
        }
    }
    
    // Get validation summary
    getValidationSummary() {
        const project = store.getProjectData();
        const summary = {
            totalFields: 0,
            completedFields: 0,
            requiredFields: 0,
            completedRequired: 0,
            arrays: {}
        };
        
        // Check regular fields
        Object.entries(this.validationRules).forEach(([fieldName, rule]) => {
            summary.totalFields++;
            
            if (rule.required) {
                summary.requiredFields++;
            }
            
            const value = project[fieldName];
            if (value && value.toString().trim() !== '') {
                summary.completedFields++;
                if (rule.required) {
                    summary.completedRequired++;
                }
            }
        });
        
        // Check arrays
        Config.FIELDS.ARRAY.forEach(arrayName => {
            const items = project[arrayName] || [];
            summary.arrays[arrayName] = {
                count: items.length,
                complete: items.length > 0
            };
        });
        
        summary.completionRate = Math.round((summary.completedFields / summary.totalFields) * 100);
        summary.requiredCompletionRate = Math.round((summary.completedRequired / summary.requiredFields) * 100);
        
        return summary;
    }
}