// js/utils/helpers.js
import { Config } from '../config.js';
import { store } from '../core/Store.js';

// Field ID mapping
export function getFieldHtmlId(fieldName) {
    return Config.FIELD_ID_MAP[fieldName] || fieldName;
}

// Notifications
export function showNotification(message, type = 'info', duration = Config.UI.NOTIFICATION_DURATION) {
    return store.addNotification(message, type, duration);
}

// Debounce function
export function debounce(func, wait = Config.UI.DEBOUNCE_DELAY) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// Format currency
export function formatCurrency(amount) {
    return new Intl.NumberFormat('fr-FR', {
        style: 'currency',
        currency: 'EUR',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(amount);
}

// Format number
export function formatNumber(num) {
    return new Intl.NumberFormat('fr-FR').format(num);
}

// Parse date
export function parseDate(dateStr) {
    if (!dateStr) return null;
    const date = new Date(dateStr);
    return isNaN(date.getTime()) ? null : date;
}

// Format date
export function formatDate(date) {
    if (!date) return '';
    if (typeof date === 'string') {
        date = parseDate(date);
    }
    if (!date) return '';
    
    return date.toLocaleDateString('fr-FR', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
    });
}

// Calculate working days between dates
export function calculateWorkingDays(startDate, endDate) {
    if (!startDate || !endDate) return 0;
    
    const start = parseDate(startDate);
    const end = parseDate(endDate);
    
    if (!start || !end || start > end) return 0;
    
    let count = 0;
    const current = new Date(start);
    
    while (current <= end) {
        const dayOfWeek = current.getDay();
        if (dayOfWeek !== 0 && dayOfWeek !== 6) { // Not Sunday or Saturday
            count++;
        }
        current.setDate(current.getDate() + 1);
    }
    
    return count;
}

// Deep clone object
export function deepClone(obj) {
    if (obj === null || typeof obj !== 'object') return obj;
    if (obj instanceof Date) return new Date(obj.getTime());
    if (obj instanceof Array) return obj.map(item => deepClone(item));
    if (obj instanceof Object) {
        const clonedObj = {};
        for (const key in obj) {
            if (obj.hasOwnProperty(key)) {
                clonedObj[key] = deepClone(obj[key]);
            }
        }
        return clonedObj;
    }
}

// Download file
export function downloadFile(content, filename, mimeType = 'application/json') {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// Upload file
export function uploadFile(accept = '.json') {
    return new Promise((resolve, reject) => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = accept;
        
        input.onchange = (event) => {
            const file = event.target.files[0];
            if (!file) {
                reject(new Error('No file selected'));
                return;
            }
            
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const content = e.target.result;
                    resolve({ content, file });
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsText(file);
        };
        
        input.click();
    });
}

// Validate email
export function isValidEmail(email) {
    const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return re.test(email);
}

// Validate URL
export function isValidUrl(url) {
    try {
        new URL(url);
        return true;
    } catch {
        return false;
    }
}

// Sanitize HTML
export function sanitizeHtml(html) {
    const div = document.createElement('div');
    div.textContent = html;
    return div.innerHTML;
}

// Generate unique ID
export function generateId() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

// Check if field is required
export function isFieldRequired(fieldName) {
    const actualFieldName = fieldName === 'perimetre-field' ? 'perimetre' : fieldName;
    return Config.REQUIRED_FIELDS.includes(fieldName) || Config.REQUIRED_FIELDS.includes(actualFieldName);
}

// Get field type
export function getFieldType(fieldName) {
    for (const [type, fields] of Object.entries(Config.FIELDS)) {
        if (fields.includes(fieldName)) {
            return type.toLowerCase();
        }
    }
    return 'unknown';
}

// Create loading spinner
export function createSpinner() {
    const spinner = document.createElement('span');
    spinner.className = 'spinner';
    return spinner;
}

// Scroll to element
export function scrollToElement(element, offset = 100) {
    if (!element) return;
    
    const top = element.getBoundingClientRect().top + window.pageYOffset - offset;
    window.scrollTo({
        top,
        behavior: 'smooth'
    });
}

// Get section for field
export function getSectionForField(fieldName) {
    for (const section of Config.SECTIONS) {
        if (section.fields.includes(fieldName)) {
            return section.id;
        }
    }
    return null;
}

// Calculate project duration
export function calculateProjectDuration(lots) {
    if (!lots || lots.length === 0) return 0;
    
    let minDate = null;
    let maxDate = null;
    
    lots.forEach(lot => {
        if (lot.col3) { // Start date
            const date = parseDate(lot.col3);
            if (date && (!minDate || date < minDate)) {
                minDate = date;
            }
        }
        if (lot.col4) { // End date
            const date = parseDate(lot.col4);
            if (date && (!maxDate || date > maxDate)) {
                maxDate = date;
            }
        }
    });
    
    if (minDate && maxDate) {
        return calculateWorkingDays(minDate, maxDate);
    }
    
    return 0;
}

// Export all notification DOM management
let notificationContainer = null;
let subscriptionInitialized = false;

function getNotificationContainer() {
    if (!notificationContainer) {
        notificationContainer = document.createElement('div');
        notificationContainer.className = 'notification-container';
        document.body.appendChild(notificationContainer);
    }
    return notificationContainer;
}

// Subscribe to store notifications (only once)
if (!subscriptionInitialized) {
    subscriptionInitialized = true;
    store.subscribe((state, oldState) => {
        if (state.ui.notifications !== oldState.ui.notifications) {
            updateNotifications(state.ui.notifications);
        }
    });
}

function updateNotifications(notifications) {
    const container = getNotificationContainer();
    
    // Remove notifications that are no longer in state
    Array.from(container.children).forEach(child => {
        const id = parseInt(child.dataset.notificationId);
        if (!notifications.find(n => n.id === id)) {
            child.style.animation = 'slideOut 0.3s ease';
            setTimeout(() => child.remove(), 300);
        }
    });
    
    // Add new notifications
    notifications.forEach(notification => {
        if (!container.querySelector(`[data-notification-id="${notification.id}"]`)) {
            const notifEl = createNotificationElement(notification);
            container.appendChild(notifEl);
        }
    });
}

function createNotificationElement(notification) {
    const div = document.createElement('div');
    div.className = `notification notification-${notification.type}`;
    div.dataset.notificationId = notification.id;
    div.innerHTML = `<span>${notification.message}</span>`;
    return div;
}