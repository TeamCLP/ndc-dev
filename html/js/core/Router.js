// js/core/Router.js
import { store } from './Store.js';
import { Config } from '../config.js';
import { Section } from '../components/Section.js';

export class Router {
    constructor() {
        this.sections = new Map();
        this.currentSection = null;
        this.init();
    }
    
    init() {
        // Listen for navigation events
        window.addEventListener('navigate', (e) => {
            this.navigateToSection(e.detail.section);
        });
        
        // Handle browser back/forward
        window.addEventListener('popstate', (e) => {
            if (e.state && e.state.section) {
                this.navigateToSection(e.state.section, false);
            }
        });
        
        // Initialize sections
        this.initializeSections();
    }
    
    initializeSections() {
        const contentArea = document.getElementById('contentArea');
        if (!contentArea) return;
        
        // Create section containers
        Config.SECTIONS.forEach(sectionConfig => {
            const sectionEl = document.createElement('div');
            sectionEl.id = sectionConfig.id;
            sectionEl.className = 'section';
            contentArea.appendChild(sectionEl);
            
            // Create section component
            const section = new Section(sectionConfig, sectionEl);
            this.sections.set(sectionConfig.id, section);
        });
    }
    
    navigateToSection(sectionId, pushState = true) {
        // Hide current section
        if (this.currentSection) {
            const currentSectionEl = document.getElementById(this.currentSection);
            if (currentSectionEl) {
                currentSectionEl.classList.remove('active');
            }
        }
        
        // Show new section
        const newSectionEl = document.getElementById(sectionId);
        if (newSectionEl) {
            newSectionEl.classList.add('active');
            this.currentSection = sectionId;
            
            // Update URL
            if (pushState) {
                const url = new URL(window.location);
                url.hash = sectionId;
                window.history.pushState(
                    { section: sectionId }, 
                    '', 
                    url.toString()
                );
            }
            
            // Update store
            store.setCurrentSection(sectionId);
            
            // Scroll to top
            window.scrollTo(0, 0);
        }
    }
    
    getCurrentSection() {
        return this.currentSection;
    }
    
    getSectionComponent(sectionId) {
        return this.sections.get(sectionId);
    }
}