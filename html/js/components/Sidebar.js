// js/components/Sidebar.js
import { Config } from '../config.js';
import { store } from '../core/Store.js';

export class Sidebar {
    constructor(container) {
        this.container = container;
        this.sections = Config.SECTIONS;
        this.init();
    }
    
    init() {
        this.render();
        this.bindEvents();
        
        // Subscribe to store changes
        this.unsubscribe = store.subscribe((state, oldState) => {
            // Update active section
            if (state.ui.currentSection !== oldState.ui.currentSection) {
                this.updateActiveSection(state.ui.currentSection);
            }
            
            // Update completion status
            if (state.project !== oldState.project) {
                this.updateCompletionStatus();
            }
        });
    }
    
    render() {
        const navItems = this.sections.map(section => `
            <button class="nav-item" data-section="${section.id}">
                <span class="nav-icon">${section.icon}</span>
                ${section.title}
            </button>
        `).join('');
        
        this.container.innerHTML = `
            <div class="sidebar-header">
                <span class="sidebar-logo">🤖</span>
                <h2>SCRIBE AI</h2>
            </div>
            <div class="nav-sections">
                ${navItems}
            </div>
            <div class="sidebar-progress">
                <div class="progress-circle">
                    <svg width="80" height="80">
                        <circle cx="40" cy="40" r="35" class="progress-circle-bg"></circle>
                        <circle cx="40" cy="40" r="35" class="progress-circle-fill" id="progressCircle"></circle>
                    </svg>
                    <div class="progress-text" id="progressText">0%</div>
                </div>
                <p style="text-align: center; color: rgba(255,255,255,0.7); font-size: 0.85rem;">
                    Progression globale
                </p>
            </div>
        `;
        
        // Set initial active section
        this.updateActiveSection(store.state.ui.currentSection);
        
        // Update completion status
        this.updateCompletionStatus();
    }
    
    bindEvents() {
        // Navigation clicks
        this.container.addEventListener('click', (e) => {
            const navItem = e.target.closest('.nav-item');
            if (navItem) {
                const sectionId = navItem.dataset.section;
                this.navigateToSection(sectionId);
            }
        });
    }
    
    navigateToSection(sectionId) {
        // Update store
        store.setCurrentSection(sectionId);
        
        // Trigger navigation event
        const event = new CustomEvent('navigate', { 
            detail: { section: sectionId } 
        });
        window.dispatchEvent(event);
        
        // Close sidebar on mobile
        if (window.innerWidth < 768) {
            store.toggleSidebar();
        }
    }
    
    updateActiveSection(sectionId) {
        // Remove active class from all items
        this.container.querySelectorAll('.nav-item').forEach(item => {
            item.classList.remove('active');
        });
        
        // Add active class to current section
        const activeItem = this.container.querySelector(`[data-section="${sectionId}"]`);
        if (activeItem) {
            activeItem.classList.add('active');
        }
    }
    
    updateCompletionStatus() {
        const project = store.state.project;
        
        this.sections.forEach(section => {
            const navItem = this.container.querySelector(`[data-section="${section.id}"]`);
            if (!navItem) return;
            
            // Check if all fields in section are completed
            const isCompleted = this.isSectionCompleted(section, project);
            
            if (isCompleted) {
                navItem.classList.add('completed');
            } else {
                navItem.classList.remove('completed');
            }
        });
    }
    
    isSectionCompleted(section, project) {
        if (!section.fields || section.fields.length === 0) {
            return false;
        }
        
        return section.fields.every(fieldName => {
            // Handle field ID mapping
            const actualFieldName = fieldName === 'perimetre-field' ? 'perimetre' : fieldName;
            const value = project[actualFieldName];
            
            // Check based on field type
            if (Array.isArray(value)) {
                return value.length > 0;
            } else if (typeof value === 'string') {
                return value.trim() !== '';
            } else {
                return !!value;
            }
        });
    }
    
    updateProgress(percentage) {
        const circle = this.container.querySelector('#progressCircle');
        const text = this.container.querySelector('#progressText');
        
        if (circle) {
            const circumference = 2 * Math.PI * 35; // radius = 35
            const offset = circumference - (percentage / 100) * circumference;
            circle.style.strokeDashoffset = offset;
        }
        
        if (text) {
            text.textContent = `${percentage}%`;
        }
    }
    
    destroy() {
        if (this.unsubscribe) {
            this.unsubscribe();
        }
    }
}