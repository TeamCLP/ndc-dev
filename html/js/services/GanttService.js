// js/services/GanttService.js
import { store } from '../core/Store.js';
import { showNotification, parseDate } from '../utils/helpers.js';

/**
 * Service de génération de diagrammes de Gantt optimisé
 * Utilise Canvas natif pour éviter les dépendances externes
 */
export class GanttService {
    constructor() {
        this.chartInstance = null;
        this.modal = null;
        this.currentZoom = 1;
        this.viewMode = 'month'; // month, week, day
    }
    
    /**
     * Fonction utilitaire pour formater les dates
     */
    formatDate(date) {
        if (!date) return '';
        if (typeof date === 'string') {
            date = parseDate(date);
        }
        if (!date || isNaN(date.getTime())) return '';
        
        return date.toLocaleDateString('fr-FR', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
        });
    }
    
    /**
     * Génère le diagramme de Gantt
     */
    async generateGantt() {
        const project = store.getProjectData();
        
        // Vérifier les données
        if (!this.hasGanttData()) {
            showNotification('⚠️ Aucun lot ou jalon défini pour générer le Gantt', 'warning');
            return null;
        }
        
        try {
            // Créer/afficher le modal
            this.showGanttModal();
            
            // Attendre que le DOM soit prêt
            await new Promise(resolve => setTimeout(resolve, 100));
            
            // Préparer les données
            const ganttData = this.prepareGanttData(project);
            
            // Générer avec Canvas
            const ganttImage = await this.generateCanvasGantt(ganttData);
            
            if (ganttImage) {
                store.setFieldValue('ganttDiagram', ganttImage);
                showNotification('✅ Diagramme de Gantt généré', 'success');
                
                // Afficher l'aperçu dans la section
                const preview = document.getElementById('ganttPreview');
                const image = document.getElementById('ganttImage');
                if (preview && image) {
                    image.src = ganttImage;
                    preview.style.display = 'block';
                }
            }
            
            return ganttImage;
            
        } catch (error) {
            console.error('Gantt generation error:', error);
            showNotification('❌ Erreur: ' + error.message, 'error');
            return null;
        }
    }
    
    /**
     * Prépare les données pour le Gantt
     */
    prepareGanttData(project) {
        const tasks = [];
        let taskId = 1;
        
        // Traiter les lots
        if (project.lots && Array.isArray(project.lots)) {
            project.lots.forEach((lot, index) => {
                if (lot && typeof lot === 'object' && lot.col3 && lot.col4) {
                    const start = parseDate(lot.col3);
                    const end = parseDate(lot.col4);
                    
                    if (start && end && !isNaN(start.getTime()) && !isNaN(end.getTime())) {
                        tasks.push({
                            id: taskId++,
                            name: lot.col0 || `Lot ${index + 1}`,
                            start: start,
                            end: end,
                            type: 'task',
                            description: lot.col1 || '',
                            duration: parseInt(lot.col2) || Math.ceil((end - start) / (1000 * 60 * 60 * 24)),
                            resources: lot.col5 || ''
                        });
                    }
                }
            });
        }
        
        // Traiter les jalons
        if (project.jalons && Array.isArray(project.jalons)) {
            project.jalons.forEach((jalon, index) => {
                if (jalon && typeof jalon === 'object' && jalon.col2) {
                    const date = parseDate(jalon.col2);
                    
                    if (date && !isNaN(date.getTime())) {
                        tasks.push({
                            id: taskId++,
                            name: jalon.col0 || `Jalon ${index + 1}`,
                            start: date,
                            end: date,
                            type: 'milestone',
                            description: jalon.col3 || ''
                        });
                    }
                }
            });
        }
        
        if (tasks.length === 0) {
            throw new Error('Aucune tâche valide trouvée');
        }
        
        // Trier par date de début
        tasks.sort((a, b) => a.start - b.start);
        
        // Calculer les dates min/max
        const dates = tasks.flatMap(t => t.type === 'milestone' ? [t.start] : [t.start, t.end]);
        const minDate = new Date(Math.min(...dates));
        const maxDate = new Date(Math.max(...dates));
        
        // Ajouter un buffer
        minDate.setDate(minDate.getDate() - 7);
        maxDate.setDate(maxDate.getDate() + 7);
        
        return {
            tasks,
            projectName: project.libelle || 'Projet',
            client: project.client || '',
            minDate,
            maxDate
        };
    }
    
    /**
     * Génère le Gantt avec Canvas
     */
    async generateCanvasGantt(data) {
        const container = document.getElementById('ganttContainer');
        if (!container) return null;
        
        // Créer le canvas
        let canvas = document.getElementById('ganttCanvas');
        if (!canvas) {
            canvas = document.createElement('canvas');
            canvas.id = 'ganttCanvas';
            container.appendChild(canvas);
        }
        
        const ctx = canvas.getContext('2d');
        
        // Configuration adaptative selon le mode de vue
        const config = this.getConfig();
        
        // Calculer les dimensions
        const containerWidth = container.clientWidth;
        const containerHeight = Math.max(700, config.headerHeight + (data.tasks.length * config.rowHeight) + 120);
        
        // Définir la taille du canvas avec le zoom
        canvas.width = containerWidth * this.currentZoom;
        canvas.height = containerHeight;
        
        // Définir la taille d'affichage
        canvas.style.width = containerWidth + 'px';
        canvas.style.height = containerHeight + 'px';
        
        // Appliquer le zoom
        if (this.currentZoom !== 1) {
            ctx.scale(this.currentZoom, 1);
        }
        
        // Dessiner le Gantt
        this.drawGantt(ctx, data, config, containerWidth);
        
        // Retourner l'image base64
        return canvas.toDataURL('image/png');
    }
    
    /**
     * Configuration selon le mode de vue
     */
    getConfig() {
        const baseConfig = {
            padding: { top: 30, right: 30, bottom: 60, left: 250 },
            headerHeight: 100,
            rowHeight: 45,
            barHeight: 28,
            milestoneSize: 14,
            fontSize: {
                title: 20,
                subtitle: 14,
                header: 12,
                task: 13,
                date: 11,
                legend: 12
            },
            colors: {
                background: '#FFFFFF',
                headerBg: '#F5F7FA',
                headerBorder: '#E5E7EB',
                gridLight: '#FAFBFC',
                gridLine: '#E5E7EB',
                border: '#D1D5DB',
                text: '#1F2937',
                textLight: '#6B7280',
                textMuted: '#9CA3AF',
                tasks: [
                    { fill: '#3B82F6', stroke: '#2563EB' }, // Bleu
                    { fill: '#10B981', stroke: '#059669' }, // Vert
                    { fill: '#F59E0B', stroke: '#D97706' }, // Orange
                    { fill: '#8B5CF6', stroke: '#7C3AED' }, // Violet
                    { fill: '#EF4444', stroke: '#DC2626' }, // Rouge
                    { fill: '#06B6D4', stroke: '#0891B2' }, // Cyan
                    { fill: '#F97316', stroke: '#EA580C' }, // Orange foncé
                    { fill: '#84CC16', stroke: '#65A30D' }, // Vert lime
                    { fill: '#EC4899', stroke: '#DB2777' }, // Rose
                    { fill: '#6366F1', stroke: '#4F46E5' }  // Indigo
                ],
                taskText: '#FFFFFF',
                milestone: '#F87171',
                milestoneStroke: '#EF4444',
                today: '#FCD34D',
                todayLine: '#F59E0B',
                weekend: '#F3F4F6'
            }
        };
        
        // Ajuster selon le mode de vue
        switch (this.viewMode) {
            case 'day':
                baseConfig.columnWidth = 60;
                break;
            case 'week':
                baseConfig.columnWidth = 25;
                break;
            case 'month':
            default:
                baseConfig.columnWidth = 5;
        }
        
        return baseConfig;
    }
    
    /**
     * Dessine le Gantt complet
     */
    drawGantt(ctx, data, config, canvasWidth) {
        // Fond blanc
        ctx.fillStyle = config.colors.background;
        ctx.fillRect(0, 0, canvasWidth, ctx.canvas.height);
        
        // Titre et infos
        this.drawHeader(ctx, data, config, canvasWidth);
        
        // Timeline
        const timelineWidth = canvasWidth - config.padding.left - config.padding.right;
        const startY = config.padding.top + 60;
        
        this.drawTimeline(ctx, data, config, startY, timelineWidth);
        
        // Grille et tâches
        const gridStartY = startY + config.headerHeight;
        this.drawGrid(ctx, data, config, gridStartY, timelineWidth);
        this.drawTasks(ctx, data, config, gridStartY, timelineWidth);
        
        // Légende
        this.drawLegend(ctx, config, ctx.canvas.height - 40, canvasWidth);
    }
    
    /**
     * Dessine l'en-tête
     */
    drawHeader(ctx, data, config, width) {
        // Titre principal
        ctx.font = `600 ${config.fontSize.title}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
        ctx.fillStyle = config.colors.text;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        
        const title = `${data.projectName} - Diagramme de Gantt`;
        ctx.fillText(title, width / 2, config.padding.top);
        
        // Sous-titre avec client et dates
        ctx.font = `${config.fontSize.subtitle}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
        ctx.fillStyle = config.colors.textLight;
        const subtitle = `${data.client ? data.client + ' | ' : ''}${this.formatDate(data.minDate)} - ${this.formatDate(data.maxDate)}`;
        ctx.fillText(subtitle, width / 2, config.padding.top + 28);
    }
    
    /**
     * Dessine la timeline
     */
    drawTimeline(ctx, data, config, startY, timelineWidth) {
        const startX = config.padding.left;
        const endX = startX + timelineWidth;
        
        // Fond de l'en-tête avec gradient subtil
        const gradient = ctx.createLinearGradient(0, startY, 0, startY + config.headerHeight);
        gradient.addColorStop(0, config.colors.headerBg);
        gradient.addColorStop(1, '#FFFFFF');
        ctx.fillStyle = gradient;
        ctx.fillRect(startX, startY, timelineWidth, config.headerHeight);
        
        // Calculer les périodes selon le mode
        const periods = this.calculatePeriods(data.minDate, data.maxDate);
        const totalDays = Math.ceil((data.maxDate - data.minDate) / (1000 * 60 * 60 * 24));
        const pixelsPerDay = timelineWidth / totalDays;
        
        // Grouper par année et mois
        const yearGroups = this.groupPeriodsByYear(periods, data.minDate, pixelsPerDay, startX);
        const monthGroups = this.groupPeriodsByMonth(periods, data.minDate, pixelsPerDay, startX);
        
        // Dessiner les années en premier (ligne du haut)
        ctx.font = `600 ${config.fontSize.header + 2}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        
        yearGroups.forEach((year, index) => {
            // Fond alterné pour les années
            if (index % 2 === 0) {
                ctx.fillStyle = config.colors.gridLight;
                ctx.fillRect(year.x, startY, year.width, 30);
            }
            
            // Ligne de séparation verticale
            if (index > 0) {
                ctx.strokeStyle = config.colors.border;
                ctx.lineWidth = 2;
                ctx.beginPath();
                ctx.moveTo(year.x, startY);
                ctx.lineTo(year.x, startY + 30);
                ctx.stroke();
            }
            
            // Année
            ctx.fillStyle = config.colors.text;
            ctx.font = `600 ${config.fontSize.header + 2}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
            ctx.fillText(year.label, year.x + year.width / 2, startY + 15);
        });
        
        // Ligne horizontale de séparation
        ctx.strokeStyle = config.colors.headerBorder;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(startX, startY + 30);
        ctx.lineTo(endX, startY + 30);
        ctx.stroke();
        
        // Dessiner les mois (deuxième ligne)
        ctx.font = `500 ${config.fontSize.header}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
        
        monthGroups.forEach((month, index) => {
            // Ligne de séparation verticale
            if (index > 0) {
                ctx.strokeStyle = config.colors.headerBorder;
                ctx.lineWidth = 1;
                ctx.beginPath();
                ctx.moveTo(month.x, startY + 30);
                ctx.lineTo(month.x, startY + config.headerHeight);
                ctx.stroke();
            }
            
            // Nom du mois (sans l'année)
            ctx.fillStyle = config.colors.text;
            ctx.fillText(month.shortLabel, month.x + month.width / 2, startY + 55);
        });
        
        // Dessiner les sous-périodes (jours/semaines) si nécessaire
        if (this.viewMode !== 'month') {
            ctx.font = `${config.fontSize.date}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
            periods.forEach((period, index) => {
                const x = startX + ((period.start - data.minDate) / (1000 * 60 * 60 * 24)) * pixelsPerDay;
                const width = period.days * pixelsPerDay;
                
                // Ligne verticale fine
                if (index > 0) {
                    ctx.strokeStyle = config.colors.gridLine;
                    ctx.lineWidth = 0.5;
                    ctx.beginPath();
                    ctx.moveTo(x, startY + 75);
                    ctx.lineTo(x, startY + config.headerHeight);
                    ctx.stroke();
                }
                
                // Texte de la sous-période
                ctx.fillStyle = config.colors.textMuted;
                ctx.fillText(period.shortLabel || period.label, x + width / 2, startY + 85);
            });
        }
        
        // Bordure de l'en-tête
        ctx.strokeStyle = config.colors.headerBorder;
        ctx.lineWidth = 1;
        ctx.strokeRect(startX, startY, timelineWidth, config.headerHeight);
    }
    
    /**
     * Groupe les périodes par année
     */
    groupPeriodsByYear(periods, minDate, pixelsPerDay, startX) {
        const years = [];
        let currentYear = null;
        
        periods.forEach(period => {
            const yearKey = period.start.getFullYear();
            
            if (!currentYear || currentYear.key !== yearKey) {
                if (currentYear) years.push(currentYear);
                
                currentYear = {
                    key: yearKey,
                    label: yearKey.toString(),
                    start: period.start,
                    x: startX + ((period.start - minDate) / (1000 * 60 * 60 * 24)) * pixelsPerDay,
                    width: 0
                };
            }
            
            currentYear.width += period.days * pixelsPerDay;
        });
        
        if (currentYear) years.push(currentYear);
        return years;
    }
    
    /**
     * Groupe les périodes par mois
     */
    groupPeriodsByMonth(periods, minDate, pixelsPerDay, startX) {
        const months = [];
        let currentMonth = null;
        
        periods.forEach(period => {
            const monthKey = `${period.start.getFullYear()}-${period.start.getMonth()}`;
            
            if (!currentMonth || currentMonth.key !== monthKey) {
                if (currentMonth) months.push(currentMonth);
                
                currentMonth = {
                    key: monthKey,
                    label: period.start.toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' }),
                    shortLabel: period.start.toLocaleDateString('fr-FR', { month: 'long' }),
                    start: period.start,
                    x: startX + ((period.start - minDate) / (1000 * 60 * 60 * 24)) * pixelsPerDay,
                    width: 0
                };
            }
            
            currentMonth.width += period.days * pixelsPerDay;
        });
        
        if (currentMonth) months.push(currentMonth);
        return months;
    }
    
    /**
     * Calcule les périodes pour la timeline
     */
    calculatePeriods(minDate, maxDate) {
        const periods = [];
        
        if (this.viewMode === 'day') {
            // Afficher chaque jour
            let current = new Date(minDate);
            while (current <= maxDate) {
                periods.push({
                    start: new Date(current),
                    days: 1,
                    label: current.toLocaleDateString('fr-FR', { day: 'numeric', month: 'short' }),
                    shortLabel: current.getDate().toString()
                });
                current.setDate(current.getDate() + 1);
            }
        } else if (this.viewMode === 'week') {
            // Afficher par semaine
            let current = new Date(minDate);
            current.setDate(current.getDate() - current.getDay() + 1); // Lundi
            
            while (current <= maxDate) {
                const weekEnd = new Date(current);
                weekEnd.setDate(weekEnd.getDate() + 6);
                
                periods.push({
                    start: new Date(current),
                    days: 7,
                    label: `Sem ${this.getWeekNumber(current)}`,
                    shortLabel: `S${this.getWeekNumber(current)}`
                });
                current.setDate(current.getDate() + 7);
            }
        } else {
            // Afficher par mois (par défaut) - mais diviser en jours pour le calcul
            let current = new Date(minDate);
            while (current <= maxDate) {
                periods.push({
                    start: new Date(current),
                    days: 1,
                    label: current.toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' })
                });
                current.setDate(current.getDate() + 1);
            }
        }
        
        return periods;
    }
    
    /**
     * Obtient le numéro de semaine
     */
    getWeekNumber(date) {
        const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
        const dayNum = d.getUTCDay() || 7;
        d.setUTCDate(d.getUTCDate() + 4 - dayNum);
        const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
        return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    }
    
    /**
     * Dessine la grille
     */
    drawGrid(ctx, data, config, startY, timelineWidth) {
        const startX = config.padding.left;
        
        // Colonne des tâches (à gauche)
        ctx.fillStyle = config.colors.gridLight;
        ctx.fillRect(0, startY, startX, data.tasks.length * config.rowHeight);
        
        // Lignes horizontales et fonds alternés
        data.tasks.forEach((task, index) => {
            const y = startY + index * config.rowHeight;
            
            // Fond alterné pour la zone du graphique
            if (index % 2 === 1) {
                ctx.fillStyle = config.colors.gridLight;
                ctx.fillRect(startX, y, timelineWidth, config.rowHeight);
            }
            
            // Ligne horizontale
            ctx.strokeStyle = config.colors.gridLine;
            ctx.lineWidth = 1;
            ctx.beginPath();
            ctx.moveTo(0, y + config.rowHeight);
            ctx.lineTo(ctx.canvas.width, y + config.rowHeight);
            ctx.stroke();
        });
        
        // Bordure verticale entre labels et graphique
        ctx.strokeStyle = config.colors.border;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(startX, startY);
        ctx.lineTo(startX, startY + data.tasks.length * config.rowHeight);
        ctx.stroke();
    }
    
    /**
     * Dessine les tâches
     */
    drawTasks(ctx, data, config, startY, timelineWidth) {
        const startX = config.padding.left;
        const totalDays = Math.ceil((data.maxDate - data.minDate) / (1000 * 60 * 60 * 24));
        const pixelsPerDay = timelineWidth / totalDays;
        
        // Ajouter ombres pour Canvas
        ctx.shadowColor = 'rgba(0, 0, 0, 0.1)';
        ctx.shadowBlur = 4;
        ctx.shadowOffsetX = 0;
        ctx.shadowOffsetY = 2;
        
        // Compteur pour les tâches (lots) pour assigner les couleurs
        let taskColorIndex = 0;
        
        data.tasks.forEach((task, index) => {
            const y = startY + index * config.rowHeight;
            const centerY = y + config.rowHeight / 2;
            
            // Nom de la tâche (sans ombre)
            ctx.shadowColor = 'transparent';
            ctx.font = `500 ${config.fontSize.task}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
            ctx.fillStyle = config.colors.text;
            ctx.textAlign = 'left';
            ctx.textBaseline = 'middle';
            
            // Tronquer le nom si trop long
            const maxWidth = config.padding.left - 40;
            let taskName = task.name;
            while (ctx.measureText(taskName).width > maxWidth && taskName.length > 3) {
                taskName = taskName.substring(0, taskName.length - 1);
            }
            if (taskName !== task.name) taskName += '...';
            
            ctx.fillText(taskName, 20, centerY);
            
            // Réactiver les ombres pour les éléments graphiques
            ctx.shadowColor = 'rgba(0, 0, 0, 0.1)';
            
            if (task.type === 'milestone') {
                // Dessiner un losange pour le jalon
                const x = startX + ((task.start - data.minDate) / (1000 * 60 * 60 * 24)) * pixelsPerDay;
                
                ctx.fillStyle = config.colors.milestone;
                ctx.strokeStyle = config.colors.milestoneStroke;
                ctx.lineWidth = 2;
                
                ctx.beginPath();
                ctx.moveTo(x, centerY - config.milestoneSize);
                ctx.lineTo(x + config.milestoneSize, centerY);
                ctx.lineTo(x, centerY + config.milestoneSize);
                ctx.lineTo(x - config.milestoneSize, centerY);
                ctx.closePath();
                ctx.fill();
                ctx.stroke();
                
                // Date du jalon (sans ombre)
                ctx.shadowColor = 'transparent';
                ctx.font = `${config.fontSize.date}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
                ctx.fillStyle = config.colors.textLight;
                ctx.textAlign = 'center';
                ctx.fillText(this.formatDate(task.start), x, centerY + config.milestoneSize + 15);
            } else {
                // Dessiner une barre pour la tâche avec couleur unique
                const taskColor = config.colors.tasks[taskColorIndex % config.colors.tasks.length];
                taskColorIndex++;
                
                const startXPos = startX + ((task.start - data.minDate) / (1000 * 60 * 60 * 24)) * pixelsPerDay;
                const endXPos = startX + ((task.end - data.minDate) / (1000 * 60 * 60 * 24)) * pixelsPerDay;
                const barWidth = endXPos - startXPos;
                const barY = centerY - config.barHeight / 2;
                
                // Coins arrondis pour la barre
                const radius = 4;
                
                ctx.fillStyle = taskColor.fill;
                ctx.strokeStyle = taskColor.stroke;
                ctx.lineWidth = 1.5;
                
                // Dessiner la barre avec coins arrondis
                ctx.beginPath();
                ctx.moveTo(startXPos + radius, barY);
                ctx.lineTo(endXPos - radius, barY);
                ctx.quadraticCurveTo(endXPos, barY, endXPos, barY + radius);
                ctx.lineTo(endXPos, barY + config.barHeight - radius);
                ctx.quadraticCurveTo(endXPos, barY + config.barHeight, endXPos - radius, barY + config.barHeight);
                ctx.lineTo(startXPos + radius, barY + config.barHeight);
                ctx.quadraticCurveTo(startXPos, barY + config.barHeight, startXPos, barY + config.barHeight - radius);
                ctx.lineTo(startXPos, barY + radius);
                ctx.quadraticCurveTo(startXPos, barY, startXPos + radius, barY);
                ctx.closePath();
                ctx.fill();
                ctx.stroke();
                
                // Durée sur la barre (si assez de place)
                ctx.shadowColor = 'transparent';
                if (barWidth > 50) {
                    ctx.font = `600 ${config.fontSize.date}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
                    ctx.fillStyle = config.colors.taskText;
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    ctx.fillText(`${task.duration}j`, startXPos + barWidth / 2, centerY);
                }
                
                // Dates de début et fin
                ctx.font = `${config.fontSize.date}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
                ctx.fillStyle = config.colors.textMuted;
                ctx.textAlign = 'left';
                ctx.textBaseline = 'bottom';
                ctx.fillText(this.formatDate(task.start), startXPos, barY - 4);
                
                if (barWidth > 120) {
                    ctx.textAlign = 'right';
                    ctx.fillText(this.formatDate(task.end), endXPos, barY - 4);
                }
            }
        });
        
        // Désactiver les ombres
        ctx.shadowColor = 'transparent';
        
        // Ligne "aujourd'hui" si dans la période
        const today = new Date();
        if (today >= data.minDate && today <= data.maxDate) {
            const todayX = startX + ((today - data.minDate) / (1000 * 60 * 60 * 24)) * pixelsPerDay;
            
            // Zone colorée pour aujourd'hui
            ctx.fillStyle = config.colors.today + '20';
            ctx.fillRect(todayX - 10, startY, 20, data.tasks.length * config.rowHeight);
            
            // Ligne verticale
            ctx.strokeStyle = config.colors.todayLine;
            ctx.lineWidth = 2;
            ctx.setLineDash([8, 4]);
            ctx.beginPath();
            ctx.moveTo(todayX, startY);
            ctx.lineTo(todayX, startY + data.tasks.length * config.rowHeight);
            ctx.stroke();
            ctx.setLineDash([]);
            
            // Label "Aujourd'hui"
            ctx.font = `600 ${config.fontSize.date}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
            ctx.fillStyle = config.colors.todayLine;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillText("Aujourd'hui", todayX, startY - 5);
        }
    }
    
    /**
     * Dessine la légende
     */
    drawLegend(ctx, config, y, width) {
        // Fond de la légende
        ctx.fillStyle = config.colors.headerBg;
        ctx.fillRect(0, y - 25, width, 50);
        
        ctx.font = `${config.fontSize.legend}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
        ctx.textBaseline = 'middle';
        
        const legendY = y;
        
        // Légende pour les tâches (plusieurs couleurs)
        let legendX = 30;
        ctx.fillStyle = config.colors.text;
        ctx.textAlign = 'left';
        ctx.fillText('Tâches/Lots:', legendX, legendY);
        legendX += ctx.measureText('Tâches/Lots:').width + 15;
        
        // Afficher quelques exemples de couleurs
        for (let i = 0; i < Math.min(4, config.colors.tasks.length); i++) {
            const taskColor = config.colors.tasks[i];
            
            // Rectangle avec coins arrondis
            ctx.fillStyle = taskColor.fill;
            ctx.strokeStyle = taskColor.stroke;
            ctx.lineWidth = 1.5;
            const barW = 20;
            const barH = 12;
            const radius = 2;
            
            ctx.beginPath();
            ctx.moveTo(legendX + radius, legendY - barH/2);
            ctx.lineTo(legendX + barW - radius, legendY - barH/2);
            ctx.quadraticCurveTo(legendX + barW, legendY - barH/2, legendX + barW, legendY - barH/2 + radius);
            ctx.lineTo(legendX + barW, legendY + barH/2 - radius);
            ctx.quadraticCurveTo(legendX + barW, legendY + barH/2, legendX + barW - radius, legendY + barH/2);
            ctx.lineTo(legendX + radius, legendY + barH/2);
            ctx.quadraticCurveTo(legendX, legendY + barH/2, legendX, legendY + barH/2 - radius);
            ctx.lineTo(legendX, legendY - barH/2 + radius);
            ctx.quadraticCurveTo(legendX, legendY - barH/2, legendX + radius, legendY - barH/2);
            ctx.closePath();
            ctx.fill();
            ctx.stroke();
            
            legendX += barW + 10;
        }
        
        // Légende pour les jalons
        legendX = width / 2 - 50;
        ctx.fillStyle = config.colors.text;
        ctx.fillText('Jalons:', legendX, legendY);
        legendX += ctx.measureText('Jalons:').width + 10;
        
        // Losange
        const size = 8;
        ctx.fillStyle = config.colors.milestone;
        ctx.strokeStyle = config.colors.milestoneStroke;
        ctx.lineWidth = 2;
        
        ctx.beginPath();
        ctx.moveTo(legendX + size, legendY - size);
        ctx.lineTo(legendX + size * 2, legendY);
        ctx.lineTo(legendX + size, legendY + size);
        ctx.lineTo(legendX, legendY);
        ctx.closePath();
        ctx.fill();
        ctx.stroke();
        
        // Date et signature
        ctx.textAlign = 'right';
        ctx.fillStyle = config.colors.textLight;
        ctx.font = `${config.fontSize.date}px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial`;
        ctx.fillText(`Généré le ${this.formatDate(new Date())} - SCRIBE AI`, width - 30, legendY);
    }
    
    /**
     * Affiche le modal du Gantt
     */
    showGanttModal() {
        // Supprimer l'ancien modal s'il existe
        if (this.modal) {
            this.modal.remove();
        }
        
        // Créer le nouveau modal
        this.modal = document.createElement('div');
        this.modal.className = 'modal active gantt-modal';
        this.modal.innerHTML = `
            <div class="modal-content" style="width: 90%; max-width: 1400px; height: 85vh;">
                <div class="modal-header" style="padding: 15px 20px; background: #F5F7FA; border-bottom: 1px solid #E5E7EB;">
                    <h3 style="margin: 0; color: #1F2937; font-weight: 600;">📊 Diagramme de Gantt</h3>
                    <div style="display: flex; gap: 10px; align-items: center;">
                        <select id="ganttViewMode" class="form-control" style="width: 120px; border: 1px solid #D1D5DB; border-radius: 6px;">
                            <option value="month">Mois</option>
                            <option value="week">Semaine</option>
                            <option value="day">Jour</option>
                        </select>
                        <button class="btn btn-small" style="padding: 6px 12px;" onclick="window.ganttService.zoomIn()">🔍+</button>
                        <button class="btn btn-small" style="padding: 6px 12px;" onclick="window.ganttService.zoomOut()">🔍-</button>
                        <button class="btn btn-small" style="padding: 6px 12px;" onclick="window.ganttService.resetZoom()">100%</button>
                        <button class="modal-close" style="background: none; border: none; font-size: 24px; color: #6B7280; cursor: pointer;" onclick="window.ganttService.closeModal()">×</button>
                    </div>
                </div>
                <div class="modal-body" style="padding: 0; overflow: hidden;">
                    <div id="ganttWrapper" style="width: 100%; height: 100%; overflow: auto; background: #F9FAFB;">
                        <div id="ganttContainer" style="min-width: 100%; background: white; padding: 20px; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);">
                            <!-- Le canvas sera inséré ici -->
                        </div>
                    </div>
                </div>
                <div class="modal-footer" style="padding: 12px 20px; background: #F5F7FA; border-top: 1px solid #E5E7EB; display: flex; justify-content: space-between; align-items: center;">
                    <span id="ganttInfo" style="color: #6B7280; font-size: 14px;"></span>
                    <div style="display: flex; gap: 10px;">
                        <button class="btn btn-secondary" style="padding: 8px 16px; border-radius: 6px;" onclick="window.ganttService.exportExcel()">
                            📊 Export Excel
                        </button>
                        <button class="btn btn-success" style="padding: 8px 16px; border-radius: 6px; background: #10B981; border-color: #10B981;" onclick="window.ganttService.downloadPNG()">
                            💾 Télécharger PNG
                        </button>
                    </div>
                </div>
            </div>
        `;
        
        document.body.appendChild(this.modal);
        
        // Event listeners
        document.getElementById('ganttViewMode').addEventListener('change', (e) => {
            this.viewMode = e.target.value;
            this.regenerateChart();
        });
        
        // Mettre à jour les infos
        this.updateInfo();
    }
    
    /**
     * Met à jour les informations affichées
     */
    updateInfo() {
        const info = document.getElementById('ganttInfo');
        if (!info) return;
        
        const project = store.getProjectData();
        const taskCount = (project.lots?.length || 0) + (project.jalons?.length || 0);
        info.textContent = `${taskCount} éléments | Zoom: ${Math.round(this.currentZoom * 100)}%`;
    }
    
    /**
     * Zoom avant
     */
    zoomIn() {
        this.currentZoom = Math.min(this.currentZoom * 1.2, 3);
        this.regenerateChart();
    }
    
    /**
     * Zoom arrière
     */
    zoomOut() {
        this.currentZoom = Math.max(this.currentZoom / 1.2, 0.5);
        this.regenerateChart();
    }
    
    /**
     * Réinitialise le zoom
     */
    resetZoom() {
        this.currentZoom = 1;
        this.regenerateChart();
    }
    
    /**
     * Régénère le graphique
     */
    async regenerateChart() {
        const project = store.getProjectData();
        const ganttData = this.prepareGanttData(project);
        await this.generateCanvasGantt(ganttData);
        this.updateInfo();
    }
    
    /**
     * Export Excel
     */
    exportExcel() {
        const project = store.getProjectData();
        const data = this.prepareGanttData(project);
        
        let csv = 'Type,Nom,Description,Début,Fin,Durée (jours),Ressources\n';
        
        data.tasks.forEach(task => {
            const type = task.type === 'milestone' ? 'Jalon' : 'Tâche';
            const duration = task.type === 'milestone' ? '0' : task.duration;
            csv += `"${type}","${task.name}","${task.description}","${this.formatDate(task.start)}","${this.formatDate(task.end)}","${duration}","${task.resources}"\n`;
        });
        
        const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `gantt_${project.libelle || 'projet'}_${new Date().toISOString().split('T')[0]}.csv`;
        link.click();
        
        showNotification('✅ Export Excel généré', 'success');
    }
    
    /**
     * Télécharge le PNG (optimisé pour A4)
     */
    async downloadPNG() {
        const canvas = document.getElementById('ganttCanvas');
        if (!canvas) {
            showNotification('❌ Aucun diagramme à télécharger', 'error');
            return;
        }
        
        // Créer un nouveau canvas pour l'export A4
        const exportCanvas = document.createElement('canvas');
        const exportCtx = exportCanvas.getContext('2d');
        
        // Dimensions A4 en pixels (300 DPI)
        const a4Width = 2480;  // 210mm at 300 DPI
        const a4Height = 3508; // 297mm at 300 DPI
        const margin = 100;    // Marges
        
        // Calculer le ratio pour ajuster à la largeur A4
        const contentWidth = a4Width - (2 * margin);
        const ratio = contentWidth / canvas.width;
        const contentHeight = canvas.height * ratio;
        
        // Si le contenu est trop haut, ajuster pour tenir sur une page
        let finalRatio = ratio;
        if (contentHeight > (a4Height - 2 * margin)) {
            finalRatio = (a4Height - 2 * margin) / canvas.height;
        }
        
        // Définir la taille du canvas d'export
        exportCanvas.width = a4Width;
        exportCanvas.height = Math.min(a4Height, contentHeight + 2 * margin);
        
        // Fond blanc
        exportCtx.fillStyle = '#FFFFFF';
        exportCtx.fillRect(0, 0, exportCanvas.width, exportCanvas.height);
        
        // Dessiner le contenu redimensionné
        exportCtx.drawImage(
            canvas,
            margin,
            margin,
            canvas.width * finalRatio,
            canvas.height * finalRatio
        );
        
        // Télécharger
        const link = document.createElement('a');
        const project = store.getProjectData();
        link.download = `gantt_${project.libelle || 'projet'}_${new Date().toISOString().split('T')[0]}.png`;
        link.href = exportCanvas.toDataURL('image/png');
        link.click();
        
        showNotification('✅ Image PNG téléchargée (format A4)', 'success');
    }
    
    /**
     * Ferme le modal
     */
    closeModal() {
        if (this.modal) {
            this.modal.remove();
            this.modal = null;
        }
    }
    
    /**
     * Vérifie si des données sont disponibles
     */
    hasGanttData() {
        const project = store.getProjectData();
        const hasLots = project.lots && Array.isArray(project.lots) && 
                       project.lots.some(lot => lot && lot.col3 && lot.col4);
        const hasJalons = project.jalons && Array.isArray(project.jalons) && 
                         project.jalons.some(jalon => jalon && jalon.col2);
        return hasLots || hasJalons;
    }
}

// Créer l'instance singleton
export const ganttService = new GanttService();

// Exposer globalement
window.ganttService = ganttService;