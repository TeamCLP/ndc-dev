// js/services/ProjectService.js
import { store } from '../core/Store.js';
import { showNotification, downloadFile, uploadFile } from '../utils/helpers.js';

export class ProjectService {
    constructor() {
        this.autosaveInterval = null;
        this.autosaveEnabled = true;
        this.autosaveDelay = 30000; // 30 seconds
        
        this.startAutosave();
    }
    
    startAutosave() {
        if (this.autosaveEnabled) {
            this.autosaveInterval = setInterval(() => {
                this.autosave();
            }, this.autosaveDelay);
        }
    }
    
    stopAutosave() {
        if (this.autosaveInterval) {
            clearInterval(this.autosaveInterval);
            this.autosaveInterval = null;
        }
    }
    
    autosave() {
        try {
            const projectData = store.getProjectData();
            const dataStr = JSON.stringify({
                data: projectData,
                timestamp: new Date().toISOString()
            });
            
            // Check size limit (5MB for localStorage)
            if (dataStr.length > 5 * 1024 * 1024) {
                console.warn('Project too large for autosave');
                return;
            }
            
            localStorage.setItem('scribe_autosave', dataStr);
            console.log('Project autosaved');
        } catch (error) {
            console.error('Autosave error:', error);
            // If quota exceeded, try to clear old data
            if (error.name === 'QuotaExceededError') {
                try {
                    localStorage.removeItem('scribe_autosave');
                    console.warn('Cleared autosave due to quota exceeded');
                } catch (e) {
                    console.error('Failed to clear autosave:', e);
                }
            }
        }
    }
    
    async saveProject() {
        try {
            const projectData = store.getProjectData();
            
            // Generate filename
            const fileName = this.generateFileName(projectData);
            
            // Prepare data with metadata
            const saveData = {
                version: '2.0',
                timestamp: new Date().toISOString(),
                metadata: {
                    client: projectData.client || 'Unknown',
                    project: projectData.libelle || 'Untitled',
                    progress: store.calculateProgress()
                },
                data: projectData,
                stats: store.state.generation.stats
            };
            
            // Convert to JSON
            const jsonStr = JSON.stringify(saveData, null, 2);
            
            // Download file
            downloadFile(jsonStr, fileName, 'application/json');
            
            showNotification('✅ Projet sauvegardé!', 'success');
            
            // Update last save time
            this.updateLastSaveTime();
            
        } catch (error) {
            console.error('Save error:', error);
            showNotification('❌ Erreur lors de la sauvegarde', 'error');
        }
    }
    
    async loadProject() {
        try {
            const result = await uploadFile('.json');
            
            // Parse JSON
            const loadedData = JSON.parse(result.content);
            
            // Check version compatibility
            if (loadedData.version && loadedData.version.startsWith('2.')) {
                // Version 2.x format
                this.loadV2Project(loadedData);
            } else {
                // Legacy format
                this.loadLegacyProject(loadedData);
            }
            
            showNotification('✅ Projet chargé avec succès!', 'success');
            
        } catch (error) {
            console.error('Load error:', error);
            showNotification('❌ Erreur lors du chargement: ' + error.message, 'error');
        }
    }
    
    loadV2Project(loadedData) {
        // Load project data
        if (loadedData.data) {
            store.setState({ project: loadedData.data });
        }
        
        // Load generation stats if available
        if (loadedData.stats) {
            store.setState({
                generation: {
                    ...store.state.generation,
                    stats: loadedData.stats
                }
            });
        }
        
        // Navigate to first section
        window.dispatchEvent(new CustomEvent('navigate', {
            detail: { section: 'informations-generales' }
        }));
    }
    
    loadLegacyProject(projectData) {
        // Direct load for legacy format
        store.setState({ project: projectData });
        
        // Navigate to first section
        window.dispatchEvent(new CustomEvent('navigate', {
            detail: { section: 'informations-generales' }
        }));
    }
    
    generateFileName(projectData) {
        const client = projectData.client || 'projet';
        const date = new Date().toISOString().split('T')[0];
        const time = new Date().toTimeString().split(' ')[0].replace(/:/g, '-');
        
        return `note_cadrage_${client}_${date}_${time}.json`;
    }
    
    updateLastSaveTime() {
        localStorage.setItem('scribe_last_save', new Date().toISOString());
    }
    
    getLastSaveTime() {
        const lastSave = localStorage.getItem('scribe_last_save');
        return lastSave ? new Date(lastSave) : null;
    }
    
    // Import/Export specific sections
    async exportSection(sectionId) {
        try {
            const section = Config.SECTIONS.find(s => s.id === sectionId);
            if (!section) throw new Error('Section not found');
            
            const projectData = store.getProjectData();
            const sectionData = {};
            
            // Extract fields for this section
            section.fields.forEach(field => {
                if (projectData[field] !== undefined) {
                    sectionData[field] = projectData[field];
                }
            });
            
            const exportData = {
                version: '2.0',
                type: 'section',
                sectionId: sectionId,
                sectionTitle: section.title,
                timestamp: new Date().toISOString(),
                data: sectionData
            };
            
            const jsonStr = JSON.stringify(exportData, null, 2);
            const fileName = `section_${sectionId}_${new Date().toISOString().split('T')[0]}.json`;
            
            downloadFile(jsonStr, fileName);
            showNotification(`✅ Section "${section.title}" exportée`, 'success');
            
        } catch (error) {
            console.error('Section export error:', error);
            showNotification('❌ Erreur lors de l\'export de la section', 'error');
        }
    }
    
    async importSection() {
        try {
            const result = await uploadFile('.json');
            const importData = JSON.parse(result.content);
            
            if (importData.type !== 'section') {
                throw new Error('Ce fichier n\'est pas une section exportée');
            }
            
            // Merge section data
            const currentProject = store.getProjectData();
            const mergedProject = {
                ...currentProject,
                ...importData.data
            };
            
            store.setState({ project: mergedProject });
            
            // Navigate to imported section
            if (importData.sectionId) {
                window.dispatchEvent(new CustomEvent('navigate', {
                    detail: { section: importData.sectionId }
                }));
            }
            
            showNotification(`✅ Section "${importData.sectionTitle}" importée`, 'success');
            
        } catch (error) {
            console.error('Section import error:', error);
            showNotification('❌ Erreur lors de l\'import: ' + error.message, 'error');
        }
    }
    
    // Project templates
    async saveAsTemplate() {
        try {
            const projectData = store.getProjectData();
            
            // Remove client-specific data
            const templateData = { ...projectData };
            const fieldsToRemove = ['client', 'numeroProjet', 'numeroDevis', 'annee'];
            fieldsToRemove.forEach(field => delete templateData[field]);
            
            const template = {
                version: '2.0',
                type: 'template',
                name: prompt('Nom du template:') || 'Template sans nom',
                description: prompt('Description du template:') || '',
                timestamp: new Date().toISOString(),
                data: templateData
            };
            
            const jsonStr = JSON.stringify(template, null, 2);
            const fileName = `template_${template.name.replace(/\s+/g, '_')}.json`;
            
            downloadFile(jsonStr, fileName);
            showNotification('✅ Template sauvegardé', 'success');
            
        } catch (error) {
            console.error('Template save error:', error);
            showNotification('❌ Erreur lors de la sauvegarde du template', 'error');
        }
    }
    
    async loadFromTemplate() {
        try {
            const result = await uploadFile('.json');
            const template = JSON.parse(result.content);
            
            if (template.type !== 'template') {
                throw new Error('Ce fichier n\'est pas un template');
            }
            
            if (confirm(`Charger le template "${template.name}"?\n${template.description}\n\nCeci remplacera les données actuelles.`)) {
                // Keep current client info
                const currentProject = store.getProjectData();
                const preservedFields = {
                    client: currentProject.client,
                    numeroProjet: currentProject.numeroProjet,
                    numeroDevis: currentProject.numeroDevis,
                    annee: currentProject.annee
                };
                
                // Load template data
                const mergedData = {
                    ...template.data,
                    ...preservedFields
                };
                
                store.setState({ project: mergedData });
                showNotification(`✅ Template "${template.name}" chargé`, 'success');
            }
            
        } catch (error) {
            console.error('Template load error:', error);
            showNotification('❌ Erreur lors du chargement du template: ' + error.message, 'error');
        }
    }
    
    // Recovery methods
    recoverFromAutosave() {
        try {
            const autosave = localStorage.getItem('scribe_autosave');
            if (!autosave) {
                showNotification('Aucune sauvegarde automatique trouvée', 'warning');
                return false;
            }
            
            const { data, timestamp } = JSON.parse(autosave);
            const saveDate = new Date(timestamp);
            
            if (confirm(`Restaurer la sauvegarde automatique du ${saveDate.toLocaleString('fr-FR')}?`)) {
                store.setState({ project: data });
                showNotification('✅ Projet restauré depuis la sauvegarde automatique', 'success');
                return true;
            }
            
            return false;
        } catch (error) {
            console.error('Recovery error:', error);
            showNotification('❌ Erreur lors de la récupération', 'error');
            return false;
        }
    }
    
    clearProject() {
        if (confirm('Êtes-vous sûr de vouloir effacer toutes les données du projet actuel?')) {
            store.clearProject();
            showNotification('✅ Projet réinitialisé', 'success');
            
            // Navigate to first section
            window.dispatchEvent(new CustomEvent('navigate', {
                detail: { section: 'informations-generales' }
            }));
        }
    }
    
    // Validation before save
    validateProject() {
        const errors = [];
        const warnings = [];
        const project = store.getProjectData();
        
        // Check required fields
        const requiredFields = ['client', 'libelle', 'annee', 'numeroProjet'];
        requiredFields.forEach(field => {
            if (!project[field] || project[field].toString().trim() === '') {
                errors.push(`Le champ "${field}" est requis`);
            }
        });
        
        // Check financial coherence
        if (project.coutsConstruction && project.coutsConstruction.length > 0) {
            project.coutsConstruction.forEach((cout, index) => {
                const jours = parseFloat(cout.col1) || 0;
                const tjm = parseFloat(cout.col2) || 0;
                const total = parseFloat(cout.col3) || 0;
                
                if (Math.abs(jours * tjm - total) > 1) {
                    warnings.push(`Ligne ${index + 1} des coûts de construction: incohérence de calcul`);
                }
            });
        }
        
        return { errors, warnings, isValid: errors.length === 0 };
    }
    
    destroy() {
        this.stopAutosave();
    }
}