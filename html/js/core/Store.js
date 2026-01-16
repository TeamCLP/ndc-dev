// js/core/Store.js
export class Store {
    constructor() {
        this.state = {
            project: {
                // Informations générales
                client: '',
                sector: '',
                project_type: '',
                complexity: '',
                libelle: '',
                annee: new Date().getFullYear().toString(),
                numeroProjet: '',
                numeroDevis: '001',
                trigramme: '',
                typeDevis: '',
                typeProjet: '',
                
                // Description
                contexte: '',
                besoin: '',
                objectifs: '',
                
                // Périmètre
                perimetre: '',
                horsPerimetre: '',
                
                // Sécurité
                dicp: '',
                dima: '',
                pdma: '',
                rgpd: 'non',
                psee: 'non',
                
                // Solution technique
                descriptionSolution: '',
                besoinPOC: 'non',
                architecture: '',
                composantsDimensionnement: '',
                
                // Service
                plageService: '',
                conditionsHorsCrash: '',
                conditionsCrashSite: '',
                resilienceApplicative: '',
                praPlanDegrade: '',
                sauvegardes: '',
                administrationSupervision: '',
                
                // Financier
                tauxContingence: '15',
                
                // RSE
                impactCO2: '',
                
                // Documentation
                modalitesPartage: '',
                lienDocumentation: '',
                
                // Arrays
                contraintes: [],
                risques: [],
                lots: [],
                livrables: [],
                jalons: [],
                coutsConstruction: [],
                coutsFonctionnement: [],
                
                // Diagrams
                ganttDiagram: '',
                architectureSchema: ''
            },
            
            ui: {
                currentSection: 'informations-generales',
                sidebarOpen: true,
                loading: false,
                notifications: []
            },
            
            generation: {
                stats: {
                    totalGenerations: 0,
                    totalTime: 0,
                    ratings: {},
                    fieldCount: {}
                },
                currentStrategy: 'related', // Changed from 'smart' to 'related' (contextuelle)
                pendingFields: []
            },
            
            validation: {
                errors: {},
                warnings: {}
            }
        };
        
        this.listeners = new Map();
        this.loadFromLocalStorage();
    }
    
    // State Management
    getState() {
        return JSON.parse(JSON.stringify(this.state));
    }
    
    setState(updates, silent = false) {
        const oldState = this.getState();
        this.state = this.deepMerge(this.state, updates);
        
        if (!silent) {
            this.notifyListeners(oldState, this.state);
        }
        
        this.saveToLocalStorage();
    }
    
    getProjectData() {
        return JSON.parse(JSON.stringify(this.state.project));
    }
    
    setProjectData(data) {
        this.setState({ project: { ...this.state.project, ...data } });
    }
    
    // Field Management
    getFieldValue(fieldName) {
        if (fieldName === 'perimetre-field') {
            return this.state.project.perimetre;
        }
        return this.state.project[fieldName];
    }
    
    setFieldValue(fieldName, value) {
        const actualFieldName = fieldName === 'perimetre-field' ? 'perimetre' : fieldName;
        this.setState({
            project: {
                ...this.state.project,
                [actualFieldName]: value
            }
        });
    }
    
    // Array Management
    getArrayItems(arrayName) {
        return this.state.project[arrayName] || [];
    }
    
    addArrayItem(arrayName, item) {
        const currentArray = this.getArrayItems(arrayName);
        
        // Pour coutsFonctionnement, s'assurer que col2 est stocké en mensuel
        if (arrayName === 'coutsFonctionnement') {
            const quantite = parseFloat(item.col1) || 0;
            const prixMensuel = parseFloat(item.col2) || 0;
            item.col3 = (quantite * prixMensuel * 12).toString();
            // S'assurer que col4 existe
            if (item.col4 === undefined) {
                item.col4 = '';
            }
        }
        
        // Pour coutsConstruction, s'assurer que col4 existe
        if (arrayName === 'coutsConstruction' && item.col4 === undefined) {
            item.col4 = '';
        }
        
        this.setState({
            project: {
                ...this.state.project,
                [arrayName]: [...currentArray, item]
            }
        });
    }
    
    updateArrayItem(arrayName, index, updates) {
        const currentArray = this.getArrayItems(arrayName);
        const updatedArray = [...currentArray];
        
        // Pour coutsFonctionnement, recalculer col3 si col1 ou col2 change
        if (arrayName === 'coutsFonctionnement' && (updates.col1 !== undefined || updates.col2 !== undefined)) {
            const currentItem = updatedArray[index];
            const quantite = parseFloat(updates.col1 !== undefined ? updates.col1 : currentItem.col1) || 0;
            const prixMensuel = parseFloat(updates.col2 !== undefined ? updates.col2 : currentItem.col2) || 0;
            updates.col3 = (quantite * prixMensuel * 12).toString();
        }
        
        updatedArray[index] = { ...updatedArray[index], ...updates };
        
        this.setState({
            project: {
                ...this.state.project,
                [arrayName]: updatedArray
            }
        });
    }
    
    removeArrayItem(arrayName, index) {
        const currentArray = this.getArrayItems(arrayName);
        const updatedArray = currentArray.filter((_, i) => i !== index);
        
        this.setState({
            project: {
                ...this.state.project,
                [arrayName]: updatedArray
            }
        });
    }
    
    // UI State
    setCurrentSection(sectionId) {
        this.setState({
            ui: { ...this.state.ui, currentSection: sectionId }
        });
    }
    
    toggleSidebar() {
        this.setState({
            ui: { ...this.state.ui, sidebarOpen: !this.state.ui.sidebarOpen }
        });
    }
    
    setLoading(loading) {
        this.setState({
            ui: { ...this.state.ui, loading }
        });
    }
    
    // Notifications
    addNotification(message, type = 'info', duration = 3000) {
        const id = Date.now();
        const notification = { id, message, type, duration };
        
        this.setState({
            ui: {
                ...this.state.ui,
                notifications: [...this.state.ui.notifications, notification]
            }
        });
        
        if (duration > 0) {
            setTimeout(() => this.removeNotification(id), duration);
        }
        
        return id;
    }
    
    removeNotification(id) {
        this.setState({
            ui: {
                ...this.state.ui,
                notifications: this.state.ui.notifications.filter(n => n.id !== id)
            }
        });
    }
    
    // Generation Stats
    updateGenerationStats(field, timeTaken, rating = null) {
        const stats = { ...this.state.generation.stats };
        
        stats.totalGenerations++;
        stats.totalTime += timeTaken;
        stats.fieldCount[field] = (stats.fieldCount[field] || 0) + 1;
        
        if (rating !== null) {
            if (!stats.ratings[field]) {
                stats.ratings[field] = [];
            }
            stats.ratings[field].push(rating);
        }
        
        this.setState({
            generation: { ...this.state.generation, stats }
        });
    }
    
    setGenerationStrategy(strategy) {
        this.setState({
            generation: { ...this.state.generation, currentStrategy: strategy }
        });
    }
    
    // Validation
    setFieldError(field, error) {
        this.setState({
            validation: {
                ...this.state.validation,
                errors: {
                    ...this.state.validation.errors,
                    [field]: error
                }
            }
        });
    }
    
    clearFieldError(field) {
        const errors = { ...this.state.validation.errors };
        delete errors[field];
        
        this.setState({
            validation: { ...this.state.validation, errors }
        });
    }
    
    // Persistence
    saveToLocalStorage() {
        try {
            localStorage.setItem('scribe_current_project', JSON.stringify(this.state.project));
            localStorage.setItem('scribe_generation_stats', JSON.stringify(this.state.generation.stats));
        } catch (e) {
            console.error('Error saving to localStorage:', e);
        }
    }
    
    loadFromLocalStorage() {
        try {
            const savedProject = localStorage.getItem('scribe_current_project');
            if (savedProject) {
                this.state.project = { ...this.state.project, ...JSON.parse(savedProject) };
            }
            
            const savedStats = localStorage.getItem('scribe_generation_stats');
            if (savedStats) {
                this.state.generation.stats = JSON.parse(savedStats);
            }
        } catch (e) {
            console.error('Error loading from localStorage:', e);
        }
    }
    
    clearProject() {
        const emptyProject = {};
        Object.keys(this.state.project).forEach(key => {
            if (Array.isArray(this.state.project[key])) {
                emptyProject[key] = [];
            } else if (key === 'annee') {
                emptyProject[key] = new Date().getFullYear().toString();
            } else if (key === 'numeroDevis') {
                emptyProject[key] = '001';
            } else if (key === 'tauxContingence') {
                emptyProject[key] = '15';
            } else if (['rgpd', 'psee', 'besoinPOC'].includes(key)) {
                emptyProject[key] = 'non';
            } else if (['ganttDiagram', 'architectureSchema'].includes(key)) {
                emptyProject[key] = '';
            } else {
                emptyProject[key] = '';
            }
        });
        
        this.setState({ project: emptyProject });
    }
    
    // Listeners
    subscribe(listener) {
        const id = Date.now() + Math.random();
        this.listeners.set(id, listener);
        return () => this.listeners.delete(id);
    }
    
    notifyListeners(oldState, newState) {
        this.listeners.forEach(listener => {
            try {
                listener(newState, oldState);
            } catch (e) {
                console.error('Error in state listener:', e);
            }
        });
    }
    
    // Progress Calculation
    calculateProgress() {
        const required = [
            'client', 'sector', 'project_type', 'complexity', 'libelle',
            'annee', 'numeroProjet', 'trigramme', 'contexte', 'besoin',
            'objectifs', 'perimetre', 'horsPerimetre', 'descriptionSolution',
            'architecture', 'composantsDimensionnement', 'plageService',
            'conditionsHorsCrash', 'tauxContingence'
        ];
        
        const arrays = ['contraintes', 'risques', 'lots', 'livrables', 'jalons', 
                       'coutsConstruction', 'coutsFonctionnement'];
        
        let completed = 0;
        let total = required.length + arrays.length;
        
        required.forEach(field => {
            if (this.state.project[field] && this.state.project[field].toString().trim()) {
                completed++;
            }
        });
        
        arrays.forEach(arrayName => {
            if (this.state.project[arrayName] && this.state.project[arrayName].length > 0) {
                completed++;
            }
        });
        
        return Math.round((completed / total) * 100);
    }
    
    // Financial Calculations - UPDATED FOR MONTHLY STORAGE AND NEW COLUMNS
    calculateFinancialTotals() {
        const project = this.state.project;
        
        // Construction costs
        let totalConstruction = 0;
        if (project.coutsConstruction) {
            project.coutsConstruction.forEach(item => {
                // col3 contient le total HT, col4 est le code équipe (ignoré pour les calculs)
                totalConstruction += parseFloat(item.col3) || 0;
            });
        }
        
        // Contingency
        const tauxContingence = parseFloat(project.tauxContingence) || 15;
        const contingence = Math.round(totalConstruction * (tauxContingence / 100));
        const totalAvecContingence = totalConstruction + contingence;
        
        // Operating costs - col2 is stored as monthly, calculate annual total
        let totalFonctionnement = 0;
        if (project.coutsFonctionnement) {
            project.coutsFonctionnement.forEach(item => {
                // col0: poste (texte)
                // col1: quantité
                // col2: prix mensuel unitaire (stocké en mensuel)
                // col3: montant annuel total (calculé)
                // col4: code UO (ignoré pour les calculs)
                const quantite = parseFloat(item.col1) || 0;
                const prixMensuelUnitaire = parseFloat(item.col2) || 0;
                
                // Calcul du coût annuel : quantité * prix mensuel unitaire * 12 mois
                const coutAnnuel = quantite * prixMensuelUnitaire * 12;
                totalFonctionnement += coutAnnuel;
            });
        }
        
        totalFonctionnement = Math.round(totalFonctionnement);
        const totalFonctionnement3ans = totalFonctionnement * 3;
        const tco3ans = totalAvecContingence + totalFonctionnement3ans;
        
        return {
            totalHorsContingence: totalConstruction,
            contingence,
            totalAvecContingence,
            totalFonctionnement,
            totalFonctionnement3ans,
            tco3ans
        };
    }
    
    // Utility Methods
    deepMerge(target, source) {
        const output = { ...target };
        
        Object.keys(source).forEach(key => {
            if (source[key] && typeof source[key] === 'object' && !Array.isArray(source[key])) {
                if (key in target) {
                    output[key] = this.deepMerge(target[key], source[key]);
                } else {
                    output[key] = source[key];
                }
            } else {
                output[key] = source[key];
            }
        });
        
        return output;
    }
}

// Create singleton instance
export const store = new Store();
