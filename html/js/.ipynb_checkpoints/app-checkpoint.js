// js/app.js
import { Config } from './config.js';
import { store } from './core/Store.js';
import { api } from './core/Api.js';
import { generationService } from './services/GenerationService.js';
import { ganttService } from './services/GanttService.js';
import { Sidebar } from './components/Sidebar.js';
import { Router } from './core/Router.js';
import { ProjectService } from './services/ProjectService.js';
import { ExportService } from './services/ExportService.js';
import { ValidationService } from './services/ValidationService.js';
import { Modal } from './components/Modal.js';
import { showNotification } from './utils/helpers.js';
import { MusicPlayer } from './components/MusicPlayer.js';

class ScribeApp {
    constructor() {
        this.components = {};
        this.services = {};
        this.initialized = false;
    }
    
    async init() {
        try {
            console.log('🚀 Initializing SCRIBE AI...');
            
            // Initialize UI
            this.initializeUI();
            
            // Initialize services
            this.initializeServices();
            
            // Initialize components
            this.initializeComponents();
            
            // Initialize router
            this.router = new Router();
            
            // Check API connection
            await this.checkApiConnection();
            
            // Setup event listeners
            this.setupEventListeners();
            
            // Load initial data
            this.loadInitialData();
            
            // Hide loading screen
            this.hideLoadingScreen();
            
            // Navigate to first section
            this.router.navigateToSection('informations-generales');
            
            this.initialized = true;
            console.log('✅ SCRIBE AI initialized successfully');
            
        } catch (error) {
            console.error('❌ Initialization error:', error);
            showNotification('Erreur lors de l\'initialisation', 'error');
        }
    }
    
    initializeUI() {
        // Setup actions bar
        const actionsBar = document.getElementById('actionsBar');
        if (actionsBar) {
            actionsBar.innerHTML = `
                <div class="actions-left">
                    <button class="btn btn-danger" id="newProjectBtn" onclick="window.newProject()">
                        🆕 Nouveau projet
                    </button>
                    <button class="btn btn-primary" id="saveProjectBtn">
                        💾 Sauvegarder
                    </button>
                    <button class="btn btn-secondary" id="loadProjectBtn">
                        📂 Charger
                    </button>
                    <div class="tooltip">
                        <select id="contextStrategy" class="form-control" style="width: 150px;">
                            <option value="smart">IA Intelligente</option>
                            <option value="minimal">IA Minimale</option>
                            <option value="related">IA Contextuelle</option>
                            <option value="random">IA Variée</option>
                        </select>
                        <span class="tooltiptext">
                            Stratégie de génération IA
                        </span>
                    </div>
                </div>
                <div class="actions-right">
                    <button class="btn btn-info" id="showInsightsBtn">
                        📊 Insights IA
                    </button>
                    <button class="btn btn-success" id="exportWordBtn">
                        📄 Exporter Word
                    </button>
                </div>
            `;
        }
        
        // Initialize modals container
        this.initializeModals();
    }
    
    initializeServices() {
        // Services are already imported as singletons
        this.services = {
            generation: generationService,
            project: new ProjectService(),
            export: new ExportService(),
            validation: new ValidationService()
        };
    }
    
    initializeComponents() {
        // Initialize sidebar
        const sidebarEl = document.getElementById('sidebar');
        if (sidebarEl) {
            this.components.sidebar = new Sidebar(sidebarEl);
        }
    
        // Initialize music player (en dehors du if sidebar)
        try {
            this.components.musicPlayer = new MusicPlayer();
            console.log('✅ MusicPlayer initialisé dans app.js');
        } catch (error) {
            console.error('❌ Erreur initialisation MusicPlayer:', error);
        }
        
        // Initialize modals
        this.components.correctionModal = new Modal('correctionModal', {
            title: 'Corriger la génération',
            size: 'medium'
        });
        
        this.components.insightsModal = new Modal('insightsModal', {
            title: '📊 Tableau de bord des insights IA',
            size: 'large'
        });
        
        this.components.apiConfigModal = new Modal('apiConfigModal', {
            title: 'Configuration API',
            size: 'small'
        });
    }
    
    initializeModals() {
        const modalsContainer = document.getElementById('modalsContainer');
        if (modalsContainer) {
            modalsContainer.innerHTML = `
                <!-- Correction Modal -->
                <div id="correctionModal"></div>
                
                <!-- Insights Modal -->
                <div id="insightsModal"></div>
                
                <!-- API Config Modal -->
                <div id="apiConfigModal"></div>
            `;
        }
    }
    
    setupEventListeners() {
        // Menu toggle
        const menuToggle = document.getElementById('menuToggle');
        if (menuToggle) {
            menuToggle.addEventListener('click', () => store.toggleSidebar());
        }
        
        // Project actions
        document.getElementById('saveProjectBtn')?.addEventListener('click', 
            () => this.services.project.saveProject());
        
        document.getElementById('loadProjectBtn')?.addEventListener('click', 
            () => this.services.project.loadProject());
        
        // Export
        document.getElementById('exportWordBtn')?.addEventListener('click', 
            () => this.services.export.exportToWord());
        
        // Strategy change
        document.getElementById('contextStrategy')?.addEventListener('change', 
            (e) => store.setGenerationStrategy(e.target.value));
        
        // Insights
        document.getElementById('showInsightsBtn')?.addEventListener('click', 
            () => this.showInsightsDashboard());
        
        document.getElementById('insightsButton')?.addEventListener('click', 
            () => this.showInsightsDashboard());
        
        // Store state changes
        store.subscribe((state, oldState) => {
            // Update progress when project data changes
            if (state.project !== oldState.project) {
                this.updateProgress();
            }
            
            // Update sidebar state
            if (state.ui.sidebarOpen !== oldState.ui.sidebarOpen) {
                const sidebar = document.getElementById('sidebar');
                if (sidebar) {
                    sidebar.classList.toggle('open', state.ui.sidebarOpen);
                }
            }
        });
        
        // Window resize
        window.addEventListener('resize', () => this.handleResize());
        
        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            // Ctrl+S to save
            if (e.ctrlKey && e.key === 's') {
                e.preventDefault();
                this.services.project.saveProject();
            }
            // Ctrl+E to export
            if (e.ctrlKey && e.key === 'e') {
                e.preventDefault();
                this.services.export.exportToWord();
            }
            // Ctrl+N for new project
            if (e.ctrlKey && e.key === 'n') {
                e.preventDefault();
                window.newProject();
            }
        });
    }
    
    async checkApiConnection() {
        try {
            const response = await api.checkHealth();
            if (response.status === 'ok') {
                showNotification('✅ API SCRIBE connectée', 'success');
            }
        } catch (error) {
            console.warn('API non disponible:', error);
            setTimeout(() => {
                showNotification('⚠️ API SCRIBE non disponible. Vérifiez la configuration.', 'warning');
            }, 2000);
        }
    }
    
    loadInitialData() {
        // Set initial strategy
        const strategySelect = document.getElementById('contextStrategy');
        if (strategySelect) {
            strategySelect.value = store.state.generation.currentStrategy;
            // Ensure the select has 'related' selected if it's the default
            if (strategySelect.value !== store.state.generation.currentStrategy) {
                // If for some reason the value wasn't set, try again
                setTimeout(() => {
                    const select = document.getElementById('contextStrategy');
                    if (select) {
                        select.value = 'related';
                    }
                }, 100);
            }
        }
        
        // Update initial progress
        this.updateProgress();
    }
    
    updateProgress() {
        const progress = store.calculateProgress();
        
        // Update sidebar progress if component exists
        if (this.components.sidebar) {
            this.components.sidebar.updateProgress(progress);
        }
        
        // Update stats
        const completedFields = this.getCompletedFieldsCount();
        const statFieldsEl = document.getElementById('statFields');
        if (statFieldsEl) {
            statFieldsEl.textContent = completedFields;
        }
        
        // Update time saved
        this.updateTimeSaved();
        
        // Update quality
        this.updateAverageQuality();
    }
    
    getCompletedFieldsCount() {
        let count = 0;
        const project = store.state.project;
        
        // Check regular fields
        Config.REQUIRED_FIELDS.forEach(field => {
            const actualField = field === 'perimetre-field' ? 'perimetre' : field;
            if (project[actualField] && project[actualField].toString().trim()) {
                count++;
            }
        });
        
        // Check arrays
        Config.FIELDS.ARRAY.forEach(arrayName => {
            if (project[arrayName] && project[arrayName].length > 0) {
                count++;
            }
        });
        
        return count;
    }
    
    updateTimeSaved() {
        const stats = store.state.generation.stats;
        const avgTimePerField = 5; // minutes
        const timeSaved = stats.totalGenerations * avgTimePerField;
        
        const statTimeEl = document.getElementById('statTime');
        if (statTimeEl) {
            statTimeEl.textContent = timeSaved > 60 
                ? `${Math.floor(timeSaved / 60)}h ${timeSaved % 60}min`
                : `${timeSaved}min`;
        }
    }
    
    updateAverageQuality() {
        const stats = store.state.generation.stats;
        const allRatings = Object.values(stats.ratings).flat();
        
        const statQualityEl = document.getElementById('statQuality');
        if (statQualityEl) {
            if (allRatings.length === 0) {
                statQualityEl.textContent = '-';
            } else {
                const avg = allRatings.reduce((a, b) => a + b, 0) / allRatings.length;
                statQualityEl.textContent = `⭐ ${avg.toFixed(1)}`;
            }
        }
    }
    
    async showInsightsDashboard() {
        try {
            const response = await api.getFeedbackSummary();
            if (response.success) {
                this.updateInsightsDashboard(response.summary);
            }
        } catch (error) {
            // Use local stats as fallback
            this.updateInsightsDashboardLocal();
        }
        
        this.components.insightsModal.show();
    }
    
    updateInsightsDashboard(summary) {
        // Update content in insights modal
        const content = `
            <div class="insights-stats">
                <div class="insight-stat-card">
                    <div class="insight-stat-value">${summary.total_corrections || '0'}</div>
                    <div class="insight-stat-label">Corrections totales</div>
                </div>
                <div class="insight-stat-card">
                    <div class="insight-stat-value">${this.calculateAverageQuality(summary.average_quality_scores)}</div>
                    <div class="insight-stat-label">Qualité moyenne</div>
                </div>
                <div class="insight-stat-card">
                    <div class="insight-stat-value">${this.getTopField(summary.top_fields_with_feedback)}</div>
                    <div class="insight-stat-label">Champ le plus corrigé</div>
                </div>
                <div class="insight-stat-card">
                    <div class="insight-stat-value">${this.calculateImprovementRate()}</div>
                    <div class="insight-stat-label">Taux d'amélioration</div>
                </div>
            </div>
            
            <div class="insights-details">
                <h4>Apprentissages par champ</h4>
                <div id="fieldInsightsList" class="field-insights-list">
                    ${this.renderFieldInsights(summary)}
                </div>
            </div>
            
            <div class="insights-recommendations">
                <h4>Recommandations globales</h4>
                <div id="globalRecommendations" class="recommendations-list">
                    ${this.renderGlobalRecommendations(summary)}
                </div>
            </div>
        `;
        
        this.components.insightsModal.setContent(content);
    }
    
    updateInsightsDashboardLocal() {
        const stats = store.state.generation.stats;
        
        const content = `
            <div class="insights-stats">
                <div class="insight-stat-card">
                    <div class="insight-stat-value">0</div>
                    <div class="insight-stat-label">Corrections totales</div>
                </div>
                <div class="insight-stat-card">
                    <div class="insight-stat-value">${this.calculateLocalAverageQuality()}</div>
                    <div class="insight-stat-label">Qualité moyenne</div>
                </div>
                <div class="insight-stat-card">
                    <div class="insight-stat-value">${this.getTopRatedField()}</div>
                    <div class="insight-stat-label">Meilleur champ</div>
                </div>
                <div class="insight-stat-card">
                    <div class="insight-stat-value">${stats.totalGenerations}</div>
                    <div class="insight-stat-label">Générations totales</div>
                </div>
            </div>
            
            <div class="alert alert-info">
                <span>ℹ️</span>
                <div>Données locales uniquement. Connectez-vous à l'API pour des insights complets.</div>
            </div>
        `;
        
        this.components.insightsModal.setContent(content);
    }
    
    calculateAverageQuality(scores) {
        if (!scores || Object.keys(scores).length === 0) return '-';
        const values = Object.values(scores);
        const avg = values.reduce((a, b) => a + b, 0) / values.length;
        return avg.toFixed(1);
    }
    
    calculateLocalAverageQuality() {
        const stats = store.state.generation.stats;
        const allRatings = Object.values(stats.ratings).flat();
        
        if (allRatings.length === 0) return '-';
        const avg = allRatings.reduce((a, b) => a + b, 0) / allRatings.length;
        return `⭐ ${avg.toFixed(1)}`;
    }
    
    getTopField(topFields) {
        if (!topFields || topFields.length === 0) return '-';
        return topFields[0][0];
    }
    
    getTopRatedField() {
        const stats = store.state.generation.stats;
        let topField = '-';
        let topAvg = 0;
        
        Object.entries(stats.ratings).forEach(([field, ratings]) => {
            if (ratings.length > 0) {
                const avg = ratings.reduce((a, b) => a + b, 0) / ratings.length;
                if (avg > topAvg) {
                    topAvg = avg;
                    topField = field;
                }
            }
        });
        
        return topField;
    }
    
    calculateImprovementRate() {
        const stats = store.state.generation.stats;
        const allRatings = Object.values(stats.ratings).flat();
        
        if (allRatings.length < 5) return '-';
        
        const recent = allRatings.slice(-5);
        const older = allRatings.slice(0, 5);
        
        const recentAvg = recent.reduce((a, b) => a + b, 0) / recent.length;
        const olderAvg = older.reduce((a, b) => a + b, 0) / older.length;
        
        const improvement = ((recentAvg - olderAvg) / olderAvg * 100).toFixed(0);
        return improvement > 0 ? `+${improvement}%` : `${improvement}%`;
    }
    
    renderFieldInsights(summary) {
        if (!summary.average_quality_scores) return '';
        
        return Object.entries(summary.average_quality_scores).map(([field, score]) => `
            <div class="field-insight-card">
                <div class="field-insight-header">
                    <span class="field-insight-name">${field}</span>
                    <span class="field-insight-score">⭐ ${score.toFixed(1)}/5</span>
                </div>
                <div class="field-insight-stats">
                    <span>Corrections: ${summary.corrections_by_field?.[field] || 0}</span>
                    <span>Évaluations: ${summary.ratings_by_field?.[field] || 0}</span>
                </div>
            </div>
        `).join('');
    }
    
    renderGlobalRecommendations(summary) {
        const recommendations = [];
        
        if (summary.fields_needing_improvement?.length > 0) {
            recommendations.push(`
                <div class="recommendation-item">
                    <span>⚠️</span>
                    <span>Améliorer la génération pour: ${summary.fields_needing_improvement.join(', ')}</span>
                </div>
            `);
        }
        
        return recommendations.join('') || '<p>Aucune recommandation pour le moment.</p>';
    }
    
    hideLoadingScreen() {
        const loadingScreen = document.getElementById('loadingScreen');
        if (loadingScreen) {
            loadingScreen.style.opacity = '0';
            setTimeout(() => {
                loadingScreen.style.display = 'none';
            }, 500);
        }
    }
    
    handleResize() {
        // Close sidebar on mobile if needed
        if (window.innerWidth < 768 && store.state.ui.sidebarOpen) {
            store.setState({ ui: { ...store.state.ui, sidebarOpen: false } });
        }
    }
}

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    window.scribeApp = new ScribeApp();
    window.scribeApp.init();
});

// Export global functions for correction modal
window.openCorrectionModal = (fieldName) => {
    if (!window.scribeApp || !window.scribeApp.initialized) {
        showNotification('⚠️ Application non initialisée', 'warning');
        return;
    }
    
    const correctionModal = window.scribeApp.components.correctionModal;
    if (!correctionModal) {
        showNotification('⚠️ Modal de correction non disponible', 'warning');
        return;
    }
    
    const value = store.getFieldValue(fieldName);
    
    const content = `
        <div class="form-group">
            <label>Valeur originale</label>
            <textarea class="form-control" id="originalValue" rows="4" disabled>${value}</textarea>
        </div>
        <div class="form-group">
            <label>Valeur corrigée</label>
            <textarea class="form-control" id="correctedValue" rows="4">${value}</textarea>
        </div>
        <div class="form-group">
            <label>Raison de la correction (optionnel)</label>
            <input type="text" class="form-control" id="correctionReason" 
                   placeholder="Ex: Manque de détails sur la sécurité">
        </div>
    `;
    
    correctionModal.setContent(content);
    correctionModal.setFooter(`
        <button class="btn btn-secondary" onclick="window.scribeApp.components.correctionModal.hide()">
            Annuler
        </button>
        <button class="btn btn-primary" onclick="window.submitCorrection('${fieldName}')">
            Valider la correction
        </button>
    `);
    
    correctionModal.show();
};

window.submitCorrection = async (fieldName) => {
    if (!window.scribeApp || !window.scribeApp.initialized) {
        showNotification('⚠️ Application non initialisée', 'warning');
        return;
    }
    
    const original = document.getElementById('originalValue').value;
    const corrected = document.getElementById('correctedValue').value;
    const reason = document.getElementById('correctionReason').value;
    
    if (original === corrected) {
        showNotification('⚠️ Aucune modification détectée', 'warning');
        return;
    }
    
    const success = await generationService.submitCorrection(fieldName, original, corrected, reason);
    if (success) {
        window.scribeApp.components.correctionModal.hide();
    }
};

// Fonction pour nouveau projet
window.newProject = () => {
    if (!window.scribeApp || !window.scribeApp.initialized) {
        showNotification('⚠️ Application non initialisée', 'warning');
        return;
    }
    
    if (confirm('⚠️ ATTENTION !\n\nCette action va effacer TOUTES les données du projet actuel.\n\nÊtes-vous sûr de vouloir continuer ?')) {
        window.scribeApp.services.project.clearProject();
    }
};

// Fonction pour gérer l'upload du schéma d'architecture
window.handleArchitectureUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;
    
    // Vérifier la taille (max 5MB)
    if (file.size > 5 * 1024 * 1024) {
        showNotification('❌ Le fichier est trop volumineux (max 5MB)', 'error');
        return;
    }
    
    // Vérifier le type
    if (!file.type.startsWith('image/')) {
        showNotification('❌ Veuillez sélectionner une image', 'error');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = (e) => {
        const base64Image = e.target.result;
        
        // Sauvegarder dans le store
        store.setFieldValue('architectureSchema', base64Image);
        
        // Afficher l'aperçu
        const preview = document.getElementById('architecturePreview');
        const upload = document.getElementById('architectureUpload');
        const image = document.getElementById('architectureImage');
        
        if (preview && upload && image) {
            image.src = base64Image;
            preview.style.display = 'block';
            upload.style.display = 'none';
        }
        
        showNotification('✅ Schéma d\'architecture ajouté', 'success');
    };
    
    reader.onerror = () => {
        showNotification('❌ Erreur lors de la lecture du fichier', 'error');
    };
    
    reader.readAsDataURL(file);
};

// Fonction pour supprimer le schéma d'architecture
window.removeArchitectureSchema = () => {
    if (confirm('Êtes-vous sûr de vouloir supprimer le schéma d\'architecture ?')) {
        // Supprimer du store
        store.setFieldValue('architectureSchema', '');
        
        // Masquer l'aperçu
        const preview = document.getElementById('architecturePreview');
        const upload = document.getElementById('architectureUpload');
        
        if (preview && upload) {
            preview.style.display = 'none';
            upload.style.display = 'block';
        }
        
        // Réinitialiser l'input file
        const fileInput = document.getElementById('architectureFileInput');
        if (fileInput) {
            fileInput.value = '';
        }
        
        showNotification('✅ Schéma supprimé', 'success');
    }
};

// Fonction pour générer le diagramme de Gantt
window.generateGantt = async () => {
    if (!window.scribeApp || !window.scribeApp.initialized) {
        showNotification('⚠️ Application non initialisée', 'warning');
        return;
    }
    
    if (!ganttService.hasGanttData()) {
        showNotification('⚠️ Veuillez d\'abord définir des lots avec des dates ou des jalons', 'warning');
        return;
    }
    
    const ganttImage = await ganttService.generateGantt();
    
    if (ganttImage) {
        // Afficher l'aperçu dans la section
        const preview = document.getElementById('ganttPreview');
        const image = document.getElementById('ganttImage');
        
        if (preview && image) {
            image.src = ganttImage;
            preview.style.display = 'block';
        }
    }
};
