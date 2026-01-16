// js/components/ArrayField.js
import { store } from '../core/Store.js';
import { Config, ENUM_VALUES } from '../config.js';
import { showNotification } from '../utils/helpers.js';

export class ArrayField {
    constructor(fieldName, container) {
        this.fieldName = fieldName;
        this.container = container;
        this.config = Config.ARRAY_STRUCTURES[fieldName];
        
        if (!this.config) {
            console.error(`No configuration found for array field: ${fieldName}`);
            return;
        }
        
        this.init();
    }
    
    init() {
        this.render();
        this.bindEvents();
        
        // Subscribe to store changes
        this.unsubscribe = store.subscribe((state, oldState) => {
            if (state.project[this.fieldName] !== oldState.project[this.fieldName]) {
                this.renderItems();
            }
        });
    }
    
    render() {
        this.container.innerHTML = `
            <div class="card">
                <h3>${this.getTitle()}</h3>
                <div class="array-controls">
                    <button class="btn btn-secondary" id="${this.fieldName}-add">
                        ➕ Ajouter ${this.getItemName()}
                    </button>
                    <button class="btn btn-ai" id="${this.fieldName}-generate">
                        🤖 Générer ${this.getItemsName()}
                    </button>
                    <button class="btn btn-danger" id="${this.fieldName}-clear">
                        🗑️
                    </button>
                    <span class="quality-indicator" id="quality-${this.fieldName}" style="display: none;">
                        ⭐ -
                    </span>
                </div>
                <div class="generation-feedback" id="feedback-${this.fieldName}"></div>
                <div id="${this.fieldName}-list" class="array-list"></div>
            </div>
        `;
        
        this.renderItems();
    }
    
    renderItems() {
        const items = store.getArrayItems(this.fieldName);
        const listContainer = document.getElementById(`${this.fieldName}-list`);
        
        if (!listContainer) return;
        
        listContainer.innerHTML = '';
        
        items.forEach((item, index) => {
            const itemEl = this.createItemElement(item, index);
            listContainer.appendChild(itemEl);
        });
        
        // Update quality indicator if we have ratings
        this.updateQualityIndicator();
    }
    
    createItemElement(item, index) {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'array-item';
        itemDiv.dataset.index = index;
        
        // Check if this item has negative values (for coutsFonctionnement)
        if (this.fieldName === 'coutsFonctionnement') {
            const totalValue = parseFloat(item.col3) || 0;
            if (totalValue < 0) {
                itemDiv.classList.add('has-negative-value');
            }
        }
        
        const headerTitle = this.getItemHeader(item, index);
        
        itemDiv.innerHTML = `
            <div class="array-item-header">
                <h4>${headerTitle}</h4>
                <div class="array-item-actions">
                    <button class="btn btn-small btn-danger" data-action="remove" data-index="${index}">
                        🗑️ Supprimer
                    </button>
                </div>
            </div>
            <div class="array-fields">
                ${this.renderFields(item, index)}
            </div>
        `;
        
        return itemDiv;
    }
    
    renderFields(item, index) {
        const fields = [];
        
        // Group fields by row if needed
        const rows = this.getFieldRows();
        
        rows.forEach(row => {
            const rowFields = row.map(colIndex => {
                const col = this.config.cols[colIndex];
                const label = this.config.labels[colIndex];
                const type = this.config.types[colIndex];
                let value = item[col] || '';
                
                // For coutsFonctionnement col2, the stored value is monthly but we display annual
                if (this.fieldName === 'coutsFonctionnement' && col === 'col2' && value) {
                    // Convert monthly to annual for display
                    const monthlyValue = parseFloat(value) || 0;
                    value = (monthlyValue * 12).toString();
                }
                
                return this.renderField(col, label, type, value, index);
            });
            
            if (row.length > 1) {
                fields.push(`<div class="array-field-row">${rowFields.join('')}</div>`);
            } else {
                fields.push(rowFields[0]);
            }
        });
        
        return fields.join('');
    }
    
    renderField(col, label, type, value, itemIndex) {
        let input = '';
        const isFinancialField = (this.fieldName === 'coutsConstruction' || this.fieldName === 'coutsFonctionnement') && 
                                (type === 'number' || type === 'calculated');
        
        // Check if this specific column should have currency
        const shouldHaveCurrency = isFinancialField && (
            (this.fieldName === 'coutsConstruction' && (col === 'col2' || col === 'col3')) ||
            (this.fieldName === 'coutsFonctionnement' && (col === 'col2' || col === 'col3'))
        );
        
        // Check if value is negative for styling
        const isNegative = parseFloat(value) < 0;
        const negativeClass = isNegative && this.fieldName === 'coutsFonctionnement' ? ' negative-value' : '';
        
        switch (type) {
            case 'enum':
                const options = this.config.enums?.[col] || [];
                input = `
                    <select data-col="${col}" data-index="${itemIndex}">
                        <option value="">Sélectionner...</option>
                        ${options.map(opt => 
                            `<option value="${opt}" ${value === opt ? 'selected' : ''}>${opt}</option>`
                        ).join('')}
                    </select>
                `;
                break;
                
            case 'number':
                // For financial fields, format the value
                const formattedValue = isFinancialField && value ? this.formatNumberForDisplay(value) : value;
                input = `<input type="text" 
                        data-col="${col}" 
                        data-index="${itemIndex}" 
                        data-type="number"
                        value="${formattedValue}"
                        ${isFinancialField ? `class="financial-input${negativeClass}"` : ''}
                        placeholder="${isFinancialField ? '0' : ''}">`;
                break;
                
            case 'date':
                input = `<input type="date" data-col="${col}" data-index="${itemIndex}" value="${value}">`;
                break;
                
            case 'calculated':
                const calculatedValue = isFinancialField && value ? this.formatNumberForDisplay(value) : value;
                input = `<input type="text" 
                        data-col="${col}" 
                        data-index="${itemIndex}" 
                        data-type="calculated"
                        value="${calculatedValue}" 
                        readonly 
                        style="background: var(--background-color);"
                        class="financial-input${negativeClass}">`;
                break;
                
            case 'text':
            default:
                if (label.toLowerCase().includes('description') || 
                    label.toLowerCase().includes('critère') ||
                    label.toLowerCase().includes('plan')) {
                    input = `<textarea rows="2" data-col="${col}" data-index="${itemIndex}">${value}</textarea>`;
                } else {
                    input = `<input type="text" data-col="${col}" data-index="${itemIndex}" value="${value}">`;
                }
        }
        
        // Add currency wrapper only for specific financial fields
        const fieldClass = shouldHaveCurrency ? 'array-field currency-field' : 'array-field';
        
        return `
            <div class="${fieldClass}">
                <label>${label}</label>
                ${input}
            </div>
        `;
    }
    
    bindEvents() {
        // Add button
        const addBtn = document.getElementById(`${this.fieldName}-add`);
        if (addBtn) {
            addBtn.addEventListener('click', () => this.addItem());
        }
        
        // Generate button
        const generateBtn = document.getElementById(`${this.fieldName}-generate`);
        if (generateBtn) {
            generateBtn.addEventListener('click', () => {
                window.generateArray(this.fieldName);
            });
        }
        
        // Clear button
        const clearBtn = document.getElementById(`${this.fieldName}-clear`);
        if (clearBtn) {
            clearBtn.addEventListener('click', () => this.clearArray());
        }
        
        // Delegate events for dynamic content
        this.container.addEventListener('click', (e) => {
            if (e.target.dataset.action === 'remove') {
                const index = parseInt(e.target.dataset.index);
                this.removeItem(index);
            }
        });
        
        this.container.addEventListener('change', (e) => {
            if (e.target.dataset.col && e.target.dataset.index !== undefined) {
                const col = e.target.dataset.col;
                const index = parseInt(e.target.dataset.index);
                let value = e.target.value;
                
                // Handle financial inputs
                if (e.target.classList.contains('financial-input') || e.target.dataset.type === 'number') {
                    value = this.parseFinancialValue(value);
                }
                
                // For coutsFonctionnement col2, convert annual input to monthly for storage
                if (this.fieldName === 'coutsFonctionnement' && col === 'col2') {
                    const annualValue = parseFloat(value) || 0;
                    value = (annualValue / 12).toString();
                }
                
                this.updateItem(index, col, value);
            }
        });
        
        // Input event for real-time calculation and formatting
        this.container.addEventListener('input', (e) => {
            if (e.target.dataset.col && e.target.dataset.index !== undefined) {
                const col = e.target.dataset.col;
                const index = parseInt(e.target.dataset.index);
                
                // Handle financial input formatting
                if (e.target.classList.contains('financial-input') || e.target.dataset.type === 'number') {
                    // Allow typing but validate
                    this.handleFinancialInput(e.target);
                }
            }
        });
        
        // Format on blur for financial inputs
        this.container.addEventListener('blur', (e) => {
            if ((e.target.classList.contains('financial-input') || e.target.dataset.type === 'number') && 
                e.target.dataset.col && e.target.dataset.index !== undefined) {
                const col = e.target.dataset.col;
                const index = parseInt(e.target.dataset.index);
                let value = this.parseFinancialValue(e.target.value);
                
                // Format the display
                e.target.value = this.formatNumberForDisplay(value);
                
                // Update negative value styling
                if (this.fieldName === 'coutsFonctionnement') {
                    const isNegative = parseFloat(value) < 0;
                    if (isNegative) {
                        e.target.classList.add('negative-value');
                    } else {
                        e.target.classList.remove('negative-value');
                    }
                }
                
                // For coutsFonctionnement col2, convert annual to monthly for storage
                if (this.fieldName === 'coutsFonctionnement' && col === 'col2') {
                    const annualValue = parseFloat(value) || 0;
                    value = (annualValue / 12).toString();
                }
                
                // Update the store with parsed value
                this.updateItem(index, col, value);
            }
        }, true);
        
        // Focus event to prepare for editing
        this.container.addEventListener('focus', (e) => {
            if ((e.target.classList.contains('financial-input') || e.target.dataset.type === 'number') && !e.target.readOnly) {
                // Select all text for easy replacement
                e.target.select();
            }
        }, true);
        
        // Keyboard navigation for financial inputs
        this.container.addEventListener('keydown', (e) => {
            if ((e.target.classList.contains('financial-input') || e.target.dataset.type === 'number') && !e.target.readOnly) {
                // Arrow up/down to increment/decrement
                if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
                    e.preventDefault();
                    const currentValue = this.parseFinancialValue(e.target.value) || 0;
                    const increment = e.shiftKey ? 100 : (e.ctrlKey ? 1000 : 10);
                    const newValue = e.key === 'ArrowUp' ? 
                        currentValue + increment : 
                        currentValue - increment; // Allow negative values
                    
                    e.target.value = this.formatNumberForDisplay(newValue.toString());
                    
                    // Trigger change event
                    const changeEvent = new Event('change', { bubbles: true });
                    e.target.dispatchEvent(changeEvent);
                }
            }
        });
    }
    
    // Financial input handling methods
    handleFinancialInput(input) {
        let value = input.value;
        
        // Allow only numbers, spaces, minus sign and common separators
        value = value.replace(/[^\d\s,.-]/g, '');
        
        // Ensure minus sign is only at the beginning
        const parts = value.split('-');
        if (parts.length > 2) {
            value = '-' + parts.slice(1).join('').replace(/-/g, '');
        } else if (parts.length === 2 && parts[0] !== '') {
            value = parts.join('');
        }
        
        // Prevent multiple decimal points
        const decimalParts = value.split(/[,.]/).filter(p => p);
        if (decimalParts.length > 2) {
            // Keep only first decimal separator
            const integerPart = decimalParts[0];
            const decimalPart = decimalParts.slice(1).join('');
            value = integerPart + '.' + decimalPart.substring(0, 2);
        }
        
        input.value = value;
    }
    
    parseFinancialValue(value) {
        if (!value) return '0';
        
        // Remove all non-numeric characters except decimal point and minus sign
        let cleanValue = value.toString()
            .replace(/[^\d.-]/g, '')
            .replace(/,/g, '.');
        
        // Handle multiple decimal points
        const parts = cleanValue.split('.');
        if (parts.length > 2) {
            cleanValue = parts[0] + '.' + parts.slice(1).join('');
        }
        
        // Parse to number and back to string to normalize
        const parsed = parseFloat(cleanValue) || 0;
        return parsed.toString();
    }
    
    formatNumberForDisplay(value) {
        if (!value || value === '0') return '0';
        
        const num = parseFloat(value);
        if (isNaN(num)) return '0';
        
        // Format with space as thousand separator
        return num.toLocaleString('fr-FR', {
            minimumFractionDigits: 0,
            maximumFractionDigits: 2
        });
    }
    
    addItem() {
        const newItem = this.createEmptyItem();
        store.addArrayItem(this.fieldName, newItem);
    }
    
    removeItem(index) {
        if (confirm('Êtes-vous sûr de vouloir supprimer cet élément ?')) {
            store.removeArrayItem(this.fieldName, index);
        }
    }
    
    clearArray() {
        const itemCount = store.getArrayItems(this.fieldName).length;
        if (itemCount === 0) {
            showNotification('Le tableau est déjà vide', 'info');
            return;
        }
        
        if (confirm(`Êtes-vous sûr de vouloir supprimer tous les éléments (${itemCount} élément${itemCount > 1 ? 's' : ''}) ?`)) {
            store.setState({
                project: {
                    ...store.state.project,
                    [this.fieldName]: []
                }
            });
            showNotification(`✅ Tableau "${this.getTitle()}" vidé`, 'success');
        }
    }
    
    updateItem(index, col, value) {
        const updates = { [col]: value };
        
        // Handle calculated fields
        if (this.fieldName === 'coutsConstruction' && ['col1', 'col2'].includes(col)) {
            const items = store.getArrayItems(this.fieldName);
            const item = items[index];
            
            const jours = parseFloat(col === 'col1' ? value : item.col1) || 0;
            const tjm = parseFloat(col === 'col2' ? value : item.col2) || 0;
            updates.col3 = (jours * tjm).toString();
        } else if (this.fieldName === 'coutsFonctionnement' && ['col1', 'col2'].includes(col)) {
            // Pour coutsFonctionnement, col2 est stocké en mensuel, col3 est le total annuel
            const items = store.getArrayItems(this.fieldName);
            const item = items[index];
            
            const quantite = parseFloat(col === 'col1' ? value : item.col1) || 0;
            // col2 is already in monthly value when we reach here
            const prixMensuelUnitaire = parseFloat(col === 'col2' ? value : item.col2) || 0;
            // Total annuel = quantité × prix mensuel × 12
            updates.col3 = (quantite * prixMensuelUnitaire * 12).toString();
        }
        
        // Special handling for lots dates
        if (this.fieldName === 'lots' && col === 'col2') {
            // Duration changed, update end date if start date exists
            const items = store.getArrayItems(this.fieldName);
            const item = items[index];
            
            if (item.col3) { // If start date exists
                const startDate = new Date(item.col3);
                const duration = parseInt(value) || 0;
                
                // Calculate end date considering only working days
                let endDate = new Date(startDate);
                let daysAdded = 0;
                
                while (daysAdded < duration) {
                    endDate.setDate(endDate.getDate() + 1);
                    const dayOfWeek = endDate.getDay();
                    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Not Sunday or Saturday
                        daysAdded++;
                    }
                }
                
                updates.col4 = endDate.toISOString().split('T')[0];
            }
        } else if (this.fieldName === 'lots' && col === 'col3') {
            // Start date changed, update end date if duration exists
            const items = store.getArrayItems(this.fieldName);
            const item = items[index];
            
            if (item.col2 && value) { // If duration exists and start date is provided
                const startDate = new Date(value);
                const duration = parseInt(item.col2) || 0;
                
                // Calculate end date considering only working days
                let endDate = new Date(startDate);
                let daysAdded = 0;
                
                while (daysAdded < duration) {
                    endDate.setDate(endDate.getDate() + 1);
                    const dayOfWeek = endDate.getDay();
                    if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Not Sunday or Saturday
                        daysAdded++;
                    }
                }
                
                updates.col4 = endDate.toISOString().split('T')[0];
            }
        }
        
        store.updateArrayItem(this.fieldName, index, updates);
        
        // Update visual styling for negative values in coutsFonctionnement
        if (this.fieldName === 'coutsFonctionnement' && (col === 'col1' || col === 'col2' || col === 'col3')) {
            setTimeout(() => {
                const itemEl = this.container.querySelector(`.array-item[data-index="${index}"]`);
                if (itemEl) {
                    const items = store.getArrayItems(this.fieldName);
                    const item = items[index];
                    const totalValue = parseFloat(item.col3) || 0;
                    
                    if (totalValue < 0) {
                        itemEl.classList.add('has-negative-value');
                    } else {
                        itemEl.classList.remove('has-negative-value');
                    }
                    
                    // Update all financial inputs in this item
                    itemEl.querySelectorAll('.financial-input').forEach(input => {
                        const inputValue = this.parseFinancialValue(input.value);
                        if (parseFloat(inputValue) < 0) {
                            input.classList.add('negative-value');
                        } else {
                            input.classList.remove('negative-value');
                        }
                    });
                }
            }, 10);
        }
    }
    
    createEmptyItem() {
        const item = {};
        
        // Initialize all columns
        this.config.cols.forEach((col, index) => {
            const type = this.config.types[index];
            
            if (this.config.defaults?.[col]) {
                item[col] = this.config.defaults[col];
            } else if (type === 'number' || type === 'calculated') {
                item[col] = '0';
            } else if (type === 'date') {
                item[col] = '';
            } else {
                item[col] = '';
            }
        });
        
        // Set default values for coutsConstruction
        if (this.fieldName === 'coutsConstruction') {
            item.col0 = '';      // Profil
            item.col1 = '0';     // Charge (j.h)
            item.col2 = '0';     // TJM (€)
            item.col3 = '0';     // Total HT (€)
            item.col4 = '';      // Code équipe (nouveau)
        }
        
        // Set default values for coutsFonctionnement
        if (this.fieldName === 'coutsFonctionnement') {
            item.col0 = '';      // Poste
            item.col1 = '1';     // Quantité par défaut
            item.col2 = '0';     // Prix mensuel unitaire (stocké en mensuel)
            item.col3 = '0';     // Montant annuel total (calculé)
            item.col4 = '';      // Code UO (nouveau)
        }
        
        return item;
    }
    
    getFieldRows() {
        // Define row layouts for different array types
        const layouts = {
            contraintes: [[0], [2], [1], [3]],
            risques: [[0], [1, 2], [3], [4]],
            lots: [[0], [1], [2, 3, 4], [5]],
            livrables: [[0, 2], [1], [3]],
            jalons: [[0, 1, 2], [3]],
            coutsConstruction: [[0, 1, 2, 3], [4]],  // Nouvelle ligne pour le code équipe
            coutsFonctionnement: [[0], [1, 2, 3], [4]]  // Nouvelle ligne pour le code UO
        };
        
        return layouts[this.fieldName] || this.config.cols.map((_, i) => [i]);
    }
    
    getTitle() {
        const titles = {
            contraintes: 'Contraintes du projet',
            risques: 'Risques identifiés',
            lots: 'Lotissement du projet',
            livrables: 'Livrables du projet',
            jalons: 'Jalons clés',
            coutsConstruction: 'Coûts de Construction',
            coutsFonctionnement: 'Coûts de Fonctionnement'
        };
        
        return titles[this.fieldName] || this.fieldName;
    }
    
    getItemName() {
        const names = {
            contraintes: 'une contrainte',
            risques: 'un risque',
            lots: 'un lot',
            livrables: 'un livrable',
            jalons: 'un jalon',
            coutsConstruction: 'une ligne',
            coutsFonctionnement: 'une ligne'
        };
        
        return names[this.fieldName] || 'un élément';
    }
    
    getItemsName() {
        const names = {
            contraintes: 'les contraintes',
            risques: 'les risques',
            lots: 'les lots',
            livrables: 'les livrables',
            jalons: 'les jalons',
            coutsConstruction: 'les coûts',
            coutsFonctionnement: 'les coûts'
        };
        
        return names[this.fieldName] || 'les éléments';
    }
    
    getItemHeader(item, index) {
        // Pour coutsFonctionnement, afficher le poste (col0) s'il existe
        if (this.fieldName === 'coutsFonctionnement' && item.col0) {
            return item.col0;
        }
        
        if (this.fieldName === 'lots' && item.col0) {
            return item.col0;
        }
        
        const prefixes = {
            contraintes: 'Contrainte',
            risques: 'Risque',
            lots: 'Lot',
            livrables: 'Livrable',
            jalons: 'Jalon',
            coutsConstruction: 'Ligne',
            coutsFonctionnement: 'Poste'
        };
        
        const prefix = prefixes[this.fieldName] || 'Élément';
        return `${prefix} ${index + 1}`;
    }
    
    updateQualityIndicator() {
        const stats = store.state.generation.stats;
        const ratings = stats.ratings[this.fieldName];
        
        if (ratings && ratings.length > 0) {
            const avg = ratings.reduce((a, b) => a + b, 0) / ratings.length;
            const indicator = document.getElementById(`quality-${this.fieldName}`);
            
            if (indicator) {
                indicator.textContent = `⭐ ${avg.toFixed(1)}`;
                indicator.style.display = 'inline-flex';
            }
        }
    }
    
    destroy() {
        if (this.unsubscribe) {
            this.unsubscribe();
        }
    }
}
