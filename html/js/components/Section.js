// js/components/Section.js
import { Config, ENUM_VALUES } from '../config.js';
import { store } from '../core/Store.js';
import { ArrayField } from './ArrayField.js';
import { getFieldHtmlId, isFieldRequired, getFieldType, formatCurrency } from '../utils/helpers.js';

export class Section {
    constructor(config, container) {
        this.config = config;
        this.container = container;
        this.arrayFields = new Map();
        this.init();
    }
    
    init() {
        this.render();
        this.bindEvents();
        
        // Subscribe to store changes
        this.unsubscribe = store.subscribe((state, oldState) => {
            // Update field values when store changes
            this.config.fields.forEach(fieldName => {
                const oldValue = oldState.project[fieldName];
                const newValue = state.project[fieldName];
                
                if (oldValue !== newValue && !Array.isArray(newValue)) {
                    this.updateFieldValue(fieldName, newValue);
                }
            });
            
            // Update financial totals when relevant arrays change
            if (this.config.id === 'evaluation-financiere') {
                const arraysChanged = 
                    state.project.coutsConstruction !== oldState.project.coutsConstruction ||
                    state.project.coutsFonctionnement !== oldState.project.coutsFonctionnement ||
                    state.project.tauxContingence !== oldState.project.tauxContingence;
                
                if (arraysChanged) {
                    this.updateFinancialTotals();
                }
            }
        });
    }
    
    render() {
        const content = this.getContentForSection();
        this.container.innerHTML = content;
        
        // Initialize array fields
        this.initializeArrayFields();
        
        // Load initial values
        this.loadValues();
    }
    
    getContentForSection() {
        switch (this.config.id) {
            case 'informations-generales':
                return this.renderInformationsGenerales();
            case 'description-projet':
                return this.renderDescriptionProjet();
            case 'perimetre':
                return this.renderPerimetre();
            case 'securite':
                return this.renderSecurite();
            case 'contraintes-risques':
                return this.renderContraintesRisques();
            case 'solution-technique':
                return this.renderSolutionTechnique();
            case 'mise-en-oeuvre':
                return this.renderMiseEnOeuvre();
            case 'service-fonctionnement':
                return this.renderServiceFonctionnement();
            case 'evaluation-financiere':
                return this.renderEvaluationFinanciere();
            case 'rse':
                return this.renderRSE();
            case 'documentation':
                return this.renderDocumentation();
            default:
                return '<p>Section non implémentée</p>';
        }
    }
    
    renderInformationsGenerales() {
        return `
            <div class="section-header">
                <h1>Informations Générales</h1>
                <p class="section-description">Renseignez les informations de base du projet pour initialiser la génération intelligente</p>
            </div>
            
            <div class="stats-banner">
                <div class="stat-card">
                    <div class="stat-value" id="statTime">-</div>
                    <div class="stat-label">Temps économisé</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value" id="statFields">0</div>
                    <div class="stat-label">Champs complétés</div>
                </div>
                <div class="stat-card">
                    <div class="stat-value" id="statQuality">-</div>
                    <div class="stat-label">Qualité moyenne</div>
                </div>
            </div>
            
            <div class="card">
                <div class="row">
                    ${this.renderField('client', 'Client', 'text', 'Nom du client', true)}
                    ${this.renderField('sector', 'Secteur d\'activité', 'select', null, true, ENUM_VALUES.sector)}
                </div>
                
                <div class="row">
                    ${this.renderField('project_type', 'Type de Projet', 'select', null, true, ENUM_VALUES.project_type)}
                    ${this.renderField('complexity', 'Complexité', 'select', 'Évaluez la complexité technique et organisationnelle', true, ENUM_VALUES.complexity)}
                </div>
            </div>
            
            <div class="card">
                ${this.renderField('libelle', 'Libellé du Projet', 'text', 'Titre complet et descriptif du projet', true)}
                
                <div class="row">
                    ${this.renderField('annee', 'Année', 'text', '2025', true, null, 4)}
                    ${this.renderField('numeroProjet', 'Numéro de Projet', 'text', 'Ex: 121', true)}
                    ${this.renderField('numeroDevis', 'Numéro de Devis', 'text', 'Ex: 01', true)}
                </div>
                
                <div class="row">
                    ${this.renderField('trigramme', 'Trigramme Équipe', 'text', 'Ex: INF', true, null, 3, 'text-transform: uppercase;')}
                    ${this.renderField('typeDevis', 'Type de Devis', 'select', null, true, ENUM_VALUES.typeDevis)}
                    ${this.renderField('typeProjet', 'Type de Projet', 'text', 'Ex: Migration, Refonte, Intégration...', false)}
                </div>
            </div>
        `;
    }
    
    renderDescriptionProjet() {
        return `
            <div class="section-header">
                <h1>I. Description du Projet</h1>
                <p class="section-description">Décrivez le contexte, les besoins et les objectifs du projet</p>
            </div>
            
            <div class="card">
                ${this.renderField('contexte', 'Contexte', 'textarea', 'Décrivez la situation actuelle du client et l\'origine de la demande...', true, null, null, null, 6)}
            </div>
            
            <div class="card">
                ${this.renderField('besoin', 'Besoin Exprimé', 'textarea', 'Décrivez précisément ce que le client souhaite accomplir...', true, null, null, null, 4)}
            </div>
            
            <div class="card">
                ${this.renderField('objectifs', 'Objectifs du Projet', 'textarea', '• Objectifs business:\n  - \n  - \n\n• Objectifs techniques:\n  - \n  - ', true, null, null, null, 6)}
            </div>
        `;
    }
    
    renderPerimetre() {
        return `
            <div class="section-header">
                <h1>II. Périmètre du Projet</h1>
                <p class="section-description">Définissez précisément ce qui est inclus et exclu du projet</p>
            </div>
            
            <div class="card">
                ${this.renderField('perimetre', 'Périmètre', 'textarea', 'Précisez les éléments inclus dans le projet...', true, null, null, null, 6)}
            </div>
            
            <div class="card">
                ${this.renderField('horsPerimetre', 'Hors Périmètre', 'textarea', 'Précisez explicitement ce qui n\'est pas inclus...', true, null, null, null, 4)}
            </div>
        `;
    }
    
    renderSecurite() {
        return `
            <div class="section-header">
                <h1>III. Sécurité et Réglementation</h1>
                <p class="section-description">Définissez les exigences de sécurité et de conformité</p>
            </div>
            
            <div class="card">
                <div class="alert alert-info">
                    <span>ℹ️</span>
                    <div>
                        <strong>Section manuelle</strong><br>
                        Cette section est généralement complétée en collaboration avec les équipes sécurité et conformité.
                        Les champs de cette section ne sont pas générables par l'IA.
                    </div>
                </div>
                
                ${this.renderField('dicp', 'DICP Souhaité', 'text', 'Ex: Disponibilité 99.9%, Intégrité Élevée, Confidentialité Critique...', false)}
                
                <div class="row">
                    ${this.renderField('dima', 'DIMA (jours)', 'number', 'Durée max d\'interruption', false, null, null, null, null, 0)}
                    ${this.renderField('pdma', 'PDMA (jours)', 'number', 'Perte de données max', false, null, null, null, null, 0)}
                </div>
                
                <div class="row">
                    ${this.renderField('rgpd', 'Analyse RGPD/AIVP requise ?', 'select', null, false, ENUM_VALUES.rgpd)}
                    ${this.renderField('psee', 'Validation PSEE effectuée ?', 'select', null, false, ENUM_VALUES.psee)}
                </div>
            </div>
        `;
    }
    
    renderContraintesRisques() {
        return `
            <div class="section-header">
                <h1>IV. Contraintes et Risques</h1>
                <p class="section-description">Identifiez les contraintes du projet et analysez les risques potentiels</p>
            </div>
            
            <div id="contraintes-container"></div>
            <div id="risques-container"></div>
        `;
    }
    
    renderSolutionTechnique() {
        return `
            <div class="section-header">
                <h1>V. Description Technique de la Solution</h1>
                <p class="section-description">Détaillez la solution technique proposée et son architecture</p>
            </div>
            
            <div class="card">
                ${this.renderField('descriptionSolution', 'Description de la Solution', 'textarea', 'Décrivez la solution technique proposée, ses principes et ses avantages...', true, null, null, null, 6)}
                ${this.renderField('besoinPOC', 'Besoin d\'un POC ?', 'select', 'Un Proof of Concept est-il nécessaire pour valider la solution ?', false, ENUM_VALUES.besoinPOC)}
            </div>
            
            <div class="card">
                ${this.renderField('architecture', 'Architecture', 'textarea', 'Décrivez l\'architecture technique cible...', true, null, null, null, 6)}
                
                <div class="form-group">
                    <label>Schéma d'Architecture</label>
                    <div class="architecture-upload-container">
                        <div id="architecturePreview" class="architecture-preview" style="display: none;">
                            <img id="architectureImage" src="" alt="Schéma d'architecture">
                            <button class="btn btn-danger btn-small" onclick="window.removeArchitectureSchema()">
                                🗑️ Supprimer
                            </button>
                        </div>
                        <div id="architectureUpload" class="architecture-upload">
                            <input type="file" id="architectureFileInput" accept="image/*" style="display: none;" onchange="window.handleArchitectureUpload(event)">
                            <button class="btn btn-secondary" onclick="document.getElementById('architectureFileInput').click()">
                                📎 Joindre un schéma
                            </button>
                            <p class="form-help">Formats acceptés: PNG, JPG, GIF (max 5MB)</p>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="card">
                ${this.renderField('composantsDimensionnement', 'Composants et Dimensionnement', 'textarea', 'Listez les composants de l\'infrastructure et leur dimensionnement...', true, null, null, null, 6)}
            </div>
        `;
    }
    
    renderMiseEnOeuvre() {
        return `
            <div class="section-header">
                <h1>VI. Démarche pour la Mise en Œuvre</h1>
                <p class="section-description">Définissez le découpage en lots, les livrables et le planning du projet</p>
            </div>
            
            <div id="lots-container"></div>
            <div id="livrables-container"></div>
            <div id="jalons-container"></div>
            
            <div class="card">
                <h3>📊 Visualisation du Planning</h3>
                <div class="gantt-actions">
                    <button class="btn btn-info" onclick="window.generateGantt()">
                        📊 Générer le diagramme de Gantt
                    </button>
                    <p class="form-help">Générez un diagramme de Gantt basé sur les lots et jalons définis ci-dessus</p>
                </div>
                <div id="ganttPreview" class="gantt-preview" style="display: none;">
                    <img id="ganttImage" src="" alt="Diagramme de Gantt">
                </div>
            </div>
        `;
    }
    
    renderServiceFonctionnement() {
        return `
            <div class="section-header">
                <h1>VII. Offre de Service en Fonctionnement</h1>
                <p class="section-description">Définissez les niveaux de service et les conditions de fonctionnement</p>
            </div>
            
            <div class="card">
                ${this.renderField('plageService', 'Plage de Service', 'text', 'Ex: 5j/7 8h-18h, 7j/7 24h/24', true)}
            </div>
            
            <div class="card">
                ${this.renderField('conditionsHorsCrash', 'Conditions de Fonctionnement Hors Crash Site', 'textarea', 'Décrivez les conditions normales de fonctionnement...', true, null, null, null, 4)}
                ${this.renderField('conditionsCrashSite', 'Conditions de Fonctionnement en Cas de Crash Site', 'textarea', 'Décrivez les conditions en mode dégradé...', false, null, null, null, 4)}
            </div>
            
            <div class="card">
                ${this.renderField('resilienceApplicative', 'Résilience Applicative', 'textarea', 'Décrivez les mécanismes de résilience...', false, null, null, null, 4)}
                ${this.renderField('praPlanDegrade', 'Plan de Reprise d\'Activité / Gestion du Mode Dégradé', 'textarea', 'Décrivez le PRA et les procédures de reprise...', false, null, null, null, 4)}
            </div>
            
            <div class="card">
                ${this.renderField('sauvegardes', 'Politique de Sauvegarde', 'textarea', 'Décrivez la stratégie de sauvegarde...', false, null, null, null, 4)}
                ${this.renderField('administrationSupervision', 'Administration et Supervision', 'textarea', 'Décrivez les outils et processus de supervision...', false, null, null, null, 4)}
            </div>
        `;
    }
    
    renderEvaluationFinanciere() {
        return `
            <div class="section-header">
                <h1>VIII. Évaluation Financière</h1>
                <p class="section-description">Estimez les coûts de construction et de fonctionnement du projet</p>
            </div>
            
            <div class="card">
                ${this.renderField('tauxContingence', 'Taux de Contingence (%)', 'number', 'Généralement entre 10% et 30% selon la complexité du projet', true, null, null, null, null, 10, 30)}
            </div>
            
            <div id="coutsConstruction-container"></div>
            <div id="coutsFonctionnement-container"></div>
            
            <div class="alert alert-info">
                <span>💡</span>
                <div>
                    <strong>Astuce : Coûts de décommissionnement</strong><br>
                    Dans les coûts de fonctionnement, vous pouvez saisir des valeurs négatives pour représenter les économies ou gains liés au décommissionnement d'anciens systèmes. Ces lignes apparaîtront en rouge avec la mention "(Décommissionnement)".<br><br>
                    <strong>Exemples :</strong>
                    <ul style="margin-top: 8px; margin-bottom: 0;">
                        <li>Décommissionnement serveur legacy : -2000€/an</li>
                        <li>Économie licence Oracle : -50000€/an</li>
                        <li>Réduction maintenance ancienne app : -15000€/an</li>
                    </ul>
                </div>
            </div>
            
            <div class="card">
                <h3>Synthèse financière</h3>
                <div class="financial-summary">
                    <div class="summary-row">
                        <span>Total Hors Contingence</span>
                        <strong id="totalHorsContingence">0 €</strong>
                    </div>
                    <div class="summary-row">
                        <span>Contingence <span id="contingencePercent">15</span>%</span>
                        <strong id="contingence">0 €</strong>
                    </div>
                    <div class="summary-row total">
                        <span>Total Avec Contingence</span>
                        <strong id="totalAvecContingence">0 €</strong>
                    </div>
                    <div class="summary-row">
                        <span>Total Fonctionnement /an</span>
                        <strong id="totalFonctionnement">0 €</strong>
                    </div>
                    <div class="summary-row">
                        <span>Total Fonctionnement 3 ans</span>
                        <strong id="totalFonctionnement3ans">0 €</strong>
                    </div>
                    <div class="summary-row total">
                        <span>TCO 3 ans</span>
                        <strong id="tco3ans">0 €</strong>
                    </div>
                </div>
            </div>
        `;
    }
    
    renderRSE() {
        return `
            <div class="section-header">
                <h1>IX. RSE - Impact CO2</h1>
                <p class="section-description">Évaluez l'impact environnemental du projet</p>
            </div>
            
            <div class="card">
                <div class="alert alert-success">
                    <span>🌱</span>
                    <div>
                        <strong>Génération IA disponible</strong><br>
                        L'IA peut vous aider à estimer l'impact CO2 du projet en fonction de l'architecture technique et des composants déployés.
                    </div>
                </div>
                
                ${this.renderField('impactCO2', 'Estimation Impact CO2', 'textarea', 'L\'IA va estimer l\'impact CO2 basé sur l\'architecture, les composants et le dimensionnement du projet...', false, null, null, null, 6)}
            </div>
        `;
    }
    
    renderDocumentation() {
        return `
            <div class="section-header">
                <h1>X. Gestion de la Documentation</h1>
                <p class="section-description">Définissez les modalités de gestion documentaire du projet</p>
            </div>
            
            <div class="card">
                <div class="alert alert-success">
                    <span>📚</span>
                    <div>
                        <strong>Génération IA disponible</strong><br>
                        L'IA peut générer les modalités de partage en fonction du contexte RGPD et PSEE du projet.
                    </div>
                </div>
                
                ${this.renderField('modalitesPartage', 'Modalités de Partage', 'textarea', 'L\'IA va définir les modalités de partage adaptées selon les contraintes de confidentialité et réglementaires...', false, null, null, null, 4)}
                ${this.renderField('lienDocumentation', 'Lien SharePoint / Confluence', 'url', 'https://...', false)}
            </div>
        `;
    }
    
    renderField(name, label, type, placeholder, required = false, options = null, maxLength = null, style = null, rows = null, min = null, max = null) {
        const fieldId = getFieldHtmlId(name);
        
        // Liste des champs qui peuvent être générés par l'API (mise à jour avec impactCO2 et modalitesPartage)
        const generatableFields = [
            // Champs texte générables
            'contexte', 'besoin', 'objectifs', 'perimetre', 'horsPerimetre',
            'descriptionSolution', 'architecture', 'composantsDimensionnement',
            'conditionsHorsCrash', 'conditionsCrashSite', 'resilienceApplicative',
            'praPlanDegrade', 'sauvegardes', 'administrationSupervision',
            'impactCO2', 'modalitesPartage' // AJOUT DES DEUX NOUVEAUX CHAMPS
            // Arrays générables (pas besoin ici car gérés séparément)
            // 'contraintes', 'risques', 'lots', 'livrables', 'jalons',
            // 'coutsConstruction', 'coutsFonctionnement'
        ];
        
        // Afficher le bouton AI uniquement pour les champs qui peuvent réellement être générés
        const showAIButton = generatableFields.includes(name);
        
        let input = '';
        
        switch (type) {
            case 'select':
                input = `
                    <select class="form-control" id="${fieldId}">
                        <option value="">Sélectionner...</option>
                        ${options ? options.map(opt => `<option value="${opt}">${opt}</option>`).join('') : ''}
                    </select>
                `;
                break;
                
            case 'textarea':
                input = `<textarea class="form-control" id="${fieldId}" rows="${rows || 4}" placeholder="${placeholder || ''}"></textarea>`;
                break;
                
            case 'number':
                input = `<input type="number" class="form-control" id="${fieldId}" 
                    ${placeholder ? `placeholder="${placeholder}"` : ''} 
                    ${min !== null ? `min="${min}"` : ''} 
                    ${max !== null ? `max="${max}"` : ''}>`;
                break;
                
            case 'url':
                input = `<input type="url" class="form-control" id="${fieldId}" placeholder="${placeholder || ''}">`;
                break;
                
            default:
                input = `<input type="text" class="form-control" id="${fieldId}" 
                    ${placeholder ? `placeholder="${placeholder}"` : ''} 
                    ${maxLength ? `maxlength="${maxLength}"` : ''}
                    ${style ? `style="${style}"` : ''}>`;
        }
        
        const help = placeholder && type !== 'textarea' && type !== 'text' ? `<p class="form-help">${placeholder}</p>` : '';
        
        return `
            <div class="form-group">
                <label${required ? ' class="required"' : ''}>${label}</label>
                ${showAIButton ? `
                    <div class="input-group">
                        ${input}
                        <button class="btn btn-ai" onclick="generateField('${name}')">
                            🤖 Générer
                        </button>
                    </div>
                ` : input}
                ${help}
                <div class="generation-feedback" id="feedback-${name}"></div>
            </div>
        `;
    }
    
    initializeArrayFields() {
        // Initialize array fields based on section
        const arrayFields = {
            'contraintes-risques': ['contraintes', 'risques'],
            'mise-en-oeuvre': ['lots', 'livrables', 'jalons'],
            'evaluation-financiere': ['coutsConstruction', 'coutsFonctionnement']
        };
        
        const sectionArrays = arrayFields[this.config.id] || [];
        
        sectionArrays.forEach(fieldName => {
            const container = document.getElementById(`${fieldName}-container`);
            if (container) {
                const arrayField = new ArrayField(fieldName, container);
                this.arrayFields.set(fieldName, arrayField);
            }
        });
    }
    
    bindEvents() {
        // Field change events
        this.container.addEventListener('change', (e) => {
            if (e.target.classList.contains('form-control')) {
                const fieldId = e.target.id;
                const fieldName = this.getFieldNameFromId(fieldId);
                if (fieldName) {
                    store.setFieldValue(fieldName, e.target.value);
                    
                    // Special handling for tauxContingence
                    if (fieldName === 'tauxContingence') {
                        this.updateContingencyDisplay();
                    }
                }
            }
        });
    }
    
    loadValues() {
        const project = store.getProjectData();
        
        this.config.fields.forEach(fieldName => {
            if (!Config.FIELDS.ARRAY.includes(fieldName)) {
                const value = project[fieldName];
                if (value !== undefined) {
                    this.updateFieldValue(fieldName, value);
                }
            }
        });
        
        // Update financial totals if on financial section
        if (this.config.id === 'evaluation-financiere') {
            this.updateFinancialTotals();
        }
        
        // Load architecture schema if exists
        if (this.config.id === 'solution-technique' && project.architectureSchema) {
            this.displayArchitectureSchema(project.architectureSchema);
        }
        
        // Load Gantt diagram if exists
        if (this.config.id === 'mise-en-oeuvre' && project.ganttDiagram) {
            this.displayGanttDiagram(project.ganttDiagram);
        }
    }
    
    updateFieldValue(fieldName, value) {
        const fieldId = getFieldHtmlId(fieldName);
        const field = this.container.querySelector(`#${fieldId}`);
        if (field) {
            field.value = value || '';
        }
    }
    
    getFieldNameFromId(fieldId) {
        // Reverse lookup
        if (fieldId === 'perimetre-field') return 'perimetre';
        
        // Check if it's a direct field name
        if (this.config.fields.includes(fieldId)) {
            return fieldId;
        }
        
        return null;
    }
    
    updateContingencyDisplay() {
        const percentEl = document.getElementById('contingencePercent');
        if (percentEl) {
            percentEl.textContent = store.state.project.tauxContingence || '15';
        }
        this.updateFinancialTotals();
    }
    
    updateFinancialTotals() {
        const totals = store.calculateFinancialTotals();
        
        const elements = {
            'totalHorsContingence': totals.totalHorsContingence,
            'contingence': totals.contingence,
            'totalAvecContingence': totals.totalAvecContingence,
            'totalFonctionnement': totals.totalFonctionnement,
            'totalFonctionnement3ans': totals.totalFonctionnement3ans,
            'tco3ans': totals.tco3ans
        };
        
        Object.entries(elements).forEach(([id, value]) => {
            const el = document.getElementById(id);
            if (el) {
                el.textContent = formatCurrency(value);
            }
        });
    }
    
    displayArchitectureSchema(base64Image) {
        const preview = document.getElementById('architecturePreview');
        const upload = document.getElementById('architectureUpload');
        const image = document.getElementById('architectureImage');
        
        if (preview && upload && image) {
            image.src = base64Image;
            preview.style.display = 'block';
            upload.style.display = 'none';
        }
    }
    
    displayGanttDiagram(base64Image) {
        const preview = document.getElementById('ganttPreview');
        const image = document.getElementById('ganttImage');
        
        if (preview && image) {
            image.src = base64Image;
            preview.style.display = 'block';
        }
    }
    
    destroy() {
        if (this.unsubscribe) {
            this.unsubscribe();
        }
        
        // Destroy array fields
        this.arrayFields.forEach(field => field.destroy());
        this.arrayFields.clear();
    }
}
