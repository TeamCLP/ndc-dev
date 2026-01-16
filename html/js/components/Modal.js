// js/components/Modal.js
export class Modal {
    constructor(id, options = {}) {
        this.id = id;
        this.options = {
            title: options.title || 'Modal',
            size: options.size || 'medium', // small, medium, large
            closable: options.closable !== false,
            keyboard: options.keyboard !== false,
            backdrop: options.backdrop !== false
        };
        
        this.isOpen = false;
        this.container = null;
        this.contentEl = null;
        this.footerEl = null;
        
        this.init();
    }
    
    init() {
        this.container = document.getElementById(this.id);
        if (!this.container) {
            console.error(`Modal container #${this.id} not found`);
            return;
        }
        
        this.render();
        this.bindEvents();
    }
    
    render() {
        const sizeClass = `modal-${this.options.size}`;
        
        this.container.className = 'modal';
        this.container.innerHTML = `
            <div class="modal-content ${sizeClass}">
                <div class="modal-header">
                    <h3>${this.options.title}</h3>
                    ${this.options.closable ? '<button class="modal-close">×</button>' : ''}
                </div>
                <div class="modal-body"></div>
                <div class="modal-footer"></div>
            </div>
        `;
        
        this.contentEl = this.container.querySelector('.modal-body');
        this.footerEl = this.container.querySelector('.modal-footer');
    }
    
    bindEvents() {
        // Close button
        if (this.options.closable) {
            const closeBtn = this.container.querySelector('.modal-close');
            if (closeBtn) {
                closeBtn.addEventListener('click', () => this.hide());
            }
        }
        
        // Backdrop click
        if (this.options.backdrop) {
            this.container.addEventListener('click', (e) => {
                if (e.target === this.container) {
                    this.hide();
                }
            });
        }
        
        // Keyboard events
        if (this.options.keyboard) {
            this.keyboardHandler = (e) => {
                if (e.key === 'Escape' && this.isOpen) {
                    this.hide();
                }
            };
        }
    }
    
    show() {
        if (this.isOpen) return;
        
        this.container.classList.add('active');
        this.isOpen = true;
        
        // Add keyboard listener
        if (this.keyboardHandler) {
            document.addEventListener('keydown', this.keyboardHandler);
        }
        
        // Emit show event
        this.emit('show');
        
        // Focus first focusable element
        setTimeout(() => {
            const focusable = this.container.querySelector(
                'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
            );
            if (focusable) {
                focusable.focus();
            }
        }, 100);
    }
    
    hide() {
        if (!this.isOpen) return;
        
        this.container.classList.remove('active');
        this.isOpen = false;
        
        // Remove keyboard listener
        if (this.keyboardHandler) {
            document.removeEventListener('keydown', this.keyboardHandler);
        }
        
        // Emit hide event
        this.emit('hide');
    }
    
    setContent(content) {
        if (!this.contentEl) return;
        
        if (typeof content === 'string') {
            this.contentEl.innerHTML = content;
        } else if (content instanceof HTMLElement) {
            this.contentEl.innerHTML = '';
            this.contentEl.appendChild(content);
        }
    }
    
    setFooter(content) {
        if (!this.footerEl) return;
        
        if (typeof content === 'string') {
            this.footerEl.innerHTML = content;
        } else if (content instanceof HTMLElement) {
            this.footerEl.innerHTML = '';
            this.footerEl.appendChild(content);
        }
    }
    
    setTitle(title) {
        const titleEl = this.container.querySelector('.modal-header h3');
        if (titleEl) {
            titleEl.textContent = title;
        }
    }
    
    setSize(size) {
        const content = this.container.querySelector('.modal-content');
        if (content) {
            // Remove existing size classes
            content.classList.remove('modal-small', 'modal-medium', 'modal-large');
            // Add new size class
            content.classList.add(`modal-${size}`);
        }
    }
    
    // Event system
    emit(eventName, data = {}) {
        const event = new CustomEvent(`modal:${eventName}`, {
            detail: { modal: this, ...data }
        });
        this.container.dispatchEvent(event);
    }
    
    on(eventName, handler) {
        this.container.addEventListener(`modal:${eventName}`, handler);
    }
    
    off(eventName, handler) {
        this.container.removeEventListener(`modal:${eventName}`, handler);
    }
    
    destroy() {
        if (this.keyboardHandler) {
            document.removeEventListener('keydown', this.keyboardHandler);
        }
        this.container.innerHTML = '';
    }
}