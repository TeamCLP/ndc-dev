// js/services/ExportService.js
import { store } from '../core/Store.js';
import { Config, WORD_EXPORT_CONFIG } from '../config.js';
import { showNotification, formatCurrency, formatDate } from '../utils/helpers.js';

// SaveAs implementation if not loaded from external library
if (typeof saveAs === 'undefined') {
    window.saveAs = function(blob, filename) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    };
}

// Système de thèmes pour faciliter les changements de couleurs
const THEMES = {
    default: {
        // 🎨 NOUVELLE PALETTE BASÉE SUR VOTRE DOCUMENT
        primary: "2C3E50",        // Bleu-gris foncé (comme dans le document)
        primaryDark: "1A252F",    // Version plus foncée
        primaryLight: "DDDDDD",   // Gris clair pour les bordures (6 caractères)
        accent: "34495E",         // Gris-bleu moyen
        background: "F5F5F5",     // Fond gris très clair
        danger: "E74C3C",         // Rouge moderne
        dangerLight: "FADBD8",    // Rouge très clair
        warning: "F39C12",        // Orange moderne
        text: "2C3E50",           // Même que primary pour cohérence
        textLight: "7F8C8D",      // Gris moyen
        white: "FFFFFF",          // Blanc
        coverAccent: "95A5A6"     // Gris argenté pour accents
    },
    green: {
        // Palette verte originale
        primary: "0F766E",        
        primaryDark: "134E4A",    
        primaryLight: "D1FAE5",   
        accent: "059669",         
        background: "F0FDFA",     
        danger: "DC2626",
        dangerLight: "FEF2F2",         
        warning: "F59E0B",        
        text: "1E293B",           
        textLight: "64748B",      
        white: "FFFFFF",
        coverAccent: "14B8A6"     
    },
    corporate: {
        primary: "1E40AF",        // Bleu foncé
        primaryDark: "1E3A8A",    // Bleu très foncé
        primaryLight: "DBEAFE",   // Bleu clair
        accent: "3B82F6",         // Bleu vif
        background: "EFF6FF",     // Fond bleu très clair
        danger: "DC2626",         // Rouge
        dangerLight: "FEF2F2",    // Rouge très clair
        warning: "F59E0B",        // Orange
        text: "1E293B",           // Gris foncé
        textLight: "64748B",      // Gris moyen
        white: "FFFFFF",
        coverAccent: "60A5FA"     // Bleu clair pour la page de garde
    },
    modern: {
        primary: "7C3AED",        // Violet
        primaryDark: "6D28D9",    // Violet foncé
        primaryLight: "EDE9FE",   // Violet clair
        accent: "8B5CF6",         // Violet moyen
        background: "F5F3FF",     // Fond violet très clair
        danger: "DC2626",         // Rouge
        dangerLight: "FEF2F2",    // Rouge très clair
        warning: "F59E0B",        // Orange
        text: "1E293B",           // Gris foncé
        textLight: "64748B",      // Gris moyen
        white: "FFFFFF",
        coverAccent: "A78BFA"     // Violet clair pour la page de garde
    },
    elegant: {
        primary: "0F172A",        // Noir bleuté
        primaryDark: "020617",    // Noir profond
        primaryLight: "E2E8F0",   // Gris très clair
        accent: "475569",         // Gris ardoise
        background: "F8FAFC",     // Fond gris très clair
        danger: "DC2626",         // Rouge
        dangerLight: "FEF2F2",    // Rouge très clair
        warning: "F59E0B",        // Orange
        text: "0F172A",           // Noir bleuté
        textLight: "64748B",      // Gris moyen
        white: "FFFFFF",
        coverAccent: "94A3B8"     // Gris bleu pour la page de garde
    },
    darkGreen: {
        // 🌲 PALETTE VERT FONCÉ ÉLÉGANTE
        primary: "0D3B2E",        // Vert forêt très foncé (vert sapin profond)
        primaryDark: "071F17",    // Vert presque noir
        primaryLight: "E6F2ED",   // Vert très clair (menthe pâle)
        accent: "1B5E3F",         // Vert émeraude foncé
        background: "F0F7F4",     // Fond vert très pâle
        danger: "B91C1C",         // Rouge foncé élégant
        dangerLight: "FEF2F2",    // Rouge très clair
        warning: "D97706",        // Orange brûlé
        text: "0A2A1F",           // Vert très foncé pour le texte
        textLight: "4B7C65",      // Vert moyen pour texte secondaire
        white: "FFFFFF",          // Blanc
        coverAccent: "6B9F88"     // Vert sauge pour accents
    }
};

export class ExportService {
    constructor() {
        // Check if docx library is loaded
        this.checkDocxLibrary();
        // Thème par défaut
        this.currentTheme = 'default';
        this.colors = THEMES[this.currentTheme];
    }
    
    // Méthode pour changer de thème facilement
    setTheme(themeName) {
        if (THEMES[themeName]) {
            this.currentTheme = themeName;
            this.colors = THEMES[themeName];
        } else {
            console.warn(`Theme '${themeName}' not found. Using default theme.`);
            this.currentTheme = 'default';
            this.colors = THEMES.default;
        }
    }
    
    // Méthode pour personnaliser les couleurs individuellement
    setCustomColors(customColors) {
        this.colors = { ...this.colors, ...customColors };
    }
    
    checkDocxLibrary() {
        if (typeof docx === 'undefined') {
            console.error('docx library not loaded');
            return false;
        }
        return true;
    }
    
    async exportToWord() {
        if (!this.checkDocxLibrary()) {
            showNotification('❌ Erreur: Bibliothèque d\'export non chargée', 'error');
            return;
        }
        
        // Créer la notification avec un ID pour pouvoir la supprimer après
        const notificationId = showNotification('📄 Génération du document Word en cours...', 'info', 0);
        
        try {
            // Collect all data
            const data = this.prepareExportData();
            
            // Generate document
            await this.generateWordDocument(data);
            
            // Supprimer la notification de chargement
            store.removeNotification(notificationId);
            
        } catch (error) {
            console.error('Export error:', error);
            // Supprimer la notification de chargement en cas d'erreur
            store.removeNotification(notificationId);
            showNotification('❌ Erreur lors de l\'export: ' + error.message, 'error');
        }
    }
    
    // Méthode publique pour exporter avec un thème spécifique
    async exportToWordWithTheme(themeName) {
        this.setTheme(themeName);
        return this.exportToWord();
    }
    
    // Méthode publique pour exporter avec des couleurs personnalisées
    async exportToWordWithCustomColors(customColors) {
        this.setCustomColors(customColors);
        return this.exportToWord();
    }
    
    prepareExportData() {
        const projectData = store.getProjectData();
        const financialTotals = store.calculateFinancialTotals();
        
        return {
            ...projectData,
            ...financialTotals,
            exportDate: new Date(),
            progress: store.calculateProgress()
        };
    }
    
async generateWordDocument(data) {
        // Generate document title
        const title = this.generateDocumentTitle(data);
        const documentRef = `${data.annee || new Date().getFullYear()}-${data.numeroProjet || 'XXXXXX'}-${data.numeroDevis || '001'}`;
        
        // Create document with proper section configuration
        const doc = new docx.Document({
            creator: "SCRIBE AI",
            title: title,
            description: "Note de cadrage générée par SCRIBE AI",
            styles: this.createEnhancedDocumentStyles(),
            numbering: this.createNumberingConfig(),
            features: {
                updateFields: true
            },
            settings: {
                updateFields: true,
                updateFieldsOnOpen: true,
                trackRevisions: false,
                evenAndOddHeaders: false
            },
            sections: [
                // Section 1: Cover page (no header/footer)
                {
                    properties: {
                        page: {
                            margin: {
                                top: 1134,    // 2cm
                                right: 1134,  // 2cm
                                bottom: 1134, // 2cm
                                left: 1134    // 2cm
                            }
                        }
                    },
                    children: this.createCoverPage(data, title)
                },
                // Section 2: Table of contents
                {
                    properties: {
                        page: {
                            margin: {
                                top: 1417,    // 2.5cm pour l'en-tête compact
                                right: 1134,  // 2cm
                                bottom: 1417, // 2.5cm pour le pied de page
                                left: 1134,   // 2cm
                                header: 400,  // 0.7cm (réduit)
                                footer: 400   // 0.7cm (réduit)
                            }
                        }
                    },
                    headers: {
                        default: this.createSimpleHeader(documentRef)
                    },
                    footers: {
                        default: this.createFooter(data, 2)
                    },
                    children: [
                        this.createHeading("Table des matières", 1, false),
                        new docx.TableOfContents("Table des matières", {
                            hyperlink: true,
                            headingStyleRange: "1-3",
                            stylesWithLevels: [
                                { styleName: "Heading1", level: 1 },
                                { styleName: "Heading2", level: 2 },
                                { styleName: "Heading3", level: 3 }
                            ]
                        })
                    ]
                },
                // Section 3: Main content
                {
                    properties: {
                        page: {
                            margin: {
                                top: 1417,    // 2.5cm pour l'en-tête compact
                                right: 1134,  // 2cm
                                bottom: 1417, // 2.5cm pour le pied de page
                                left: 1134,   // 2cm
                                header: 400,  // 0.7cm (réduit)
                                footer: 400   // 0.7cm (réduit)
                            }
                        }
                    },
                    headers: {
                        default: this.createFullHeader(documentRef, data)
                    },
                    footers: {
                        default: this.createFooter(data, 3)
                    },
                    children: [
                        new docx.Paragraph({ text: "" }), // Espace après l'en-tête
                        ...this.createSummary(data, title),
                        ...this.createProjectDescription(data),
                        ...this.createScope(data),
                        ...this.createSecurity(data),
                        ...this.createConstraintsAndRisks(data),
                        ...await this.createTechnicalSolution(data),
                        ...await this.createImplementation(data),
                        ...this.createOperationalService(data),
                        ...this.createFinancialEvaluation(data),
                        ...this.createRSE(data),
                        ...this.createDocumentation(data)
                    ]
                }
            ]
        });
        
        // Generate and save
        const blob = await docx.Packer.toBlob(doc);
        saveAs(blob, `${title}.docx`);
        
        showNotification('✅ Document Word généré avec succès!', 'success');
        
        // Track export
        const exportTime = Date.now() - data.exportDate.getTime();
        store.updateGenerationStats('export', exportTime);
    }
    
    createSimpleHeader(documentRef) {
        return new docx.Header({
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: "📄 ",
                            size: 16
                        }),
                        new docx.TextRun({
                            text: `${documentRef}`,
                            size: 16,
                            color: this.colors.primary
                        })
                    ],
                    alignment: docx.AlignmentType.CENTER,
                    border: {
                        bottom: {
                            color: this.colors.primaryLight,
                            space: 1,
                            size: 3,
                            style: docx.BorderStyle.SINGLE
                        }
                    },
                    spacing: { after: 100 }
                })
            ]
        });
    }
    
    createFullHeader(documentRef, data) {
        return new docx.Header({
            children: [
                new docx.Table({
                    width: {
                        size: 100,
                        type: docx.WidthType.PERCENTAGE
                    },
                    borders: {
                        top: { style: docx.BorderStyle.NONE },
                        bottom: { style: docx.BorderStyle.SINGLE, size: 3, color: this.colors.primaryLight },
                        left: { style: docx.BorderStyle.NONE },
                        right: { style: docx.BorderStyle.NONE },
                        insideHorizontal: { style: docx.BorderStyle.NONE },
                        insideVertical: { style: docx.BorderStyle.NONE }
                    },
                    rows: [
                        new docx.TableRow({
                            height: {
                                value: 400,
                                rule: docx.HeightRule.EXACT
                            },
                            children: [
                                new docx.TableCell({
                                    children: [
                                        new docx.Paragraph({
                                            children: [
                                                new docx.TextRun({
                                                    text: "🏢 CA-GIP",
                                                    size: 18,
                                                    color: this.colors.primary
                                                })
                                            ],
                                            alignment: docx.AlignmentType.LEFT,
                                            spacing: { after: 0 }
                                        })
                                    ],
                                    borders: this.noBorders(),
                                    verticalAlign: docx.VerticalAlign.CENTER,
                                    width: { size: 35, type: docx.WidthType.PERCENTAGE }
                                }),
                                new docx.TableCell({
                                    children: [
                                        new docx.Paragraph({
                                            children: [
                                                new docx.TextRun({
                                                    text: data.libelle || "Note de cadrage",
                                                    size: 16,
                                                    color: this.colors.textLight,
                                                    italics: true
                                                })
                                            ],
                                            alignment: docx.AlignmentType.CENTER,
                                            spacing: { after: 0 }
                                        })
                                    ],
                                    borders: this.noBorders(),
                                    verticalAlign: docx.VerticalAlign.CENTER,
                                    width: { size: 35, type: docx.WidthType.PERCENTAGE }
                                }),
                                new docx.TableCell({
                                    children: [
                                        new docx.Paragraph({
                                            children: [
                                                new docx.TextRun({
                                                    text: "📄 ",
                                                    size: 16
                                                }),
                                                new docx.TextRun({
                                                    text: documentRef,
                                                    size: 16,
                                                    color: this.colors.primaryDark
                                                })
                                            ],
                                            alignment: docx.AlignmentType.RIGHT,
                                            spacing: { after: 0 }
                                        })
                                    ],
                                    borders: this.noBorders(),
                                    verticalAlign: docx.VerticalAlign.CENTER,
                                    width: { size: 30, type: docx.WidthType.PERCENTAGE }
                                })
                            ]
                        })
                    ],
                    margins: {
                        bottom: 100
                    }
                })
            ]
        });
    }
    
    createFooter(data, startingPage = 1) {
        return new docx.Footer({
            children: [
                new docx.Table({
                    width: {
                        size: 100,
                        type: docx.WidthType.PERCENTAGE
                    },
                    borders: {
                        top: { style: docx.BorderStyle.SINGLE, size: 3, color: this.colors.primaryLight },
                        bottom: { style: docx.BorderStyle.NONE },
                        left: { style: docx.BorderStyle.NONE },
                        right: { style: docx.BorderStyle.NONE }
                    },
                    rows: [
                        new docx.TableRow({
                            height: {
                                value: 300,
                                rule: docx.HeightRule.EXACT
                            },
                            children: [
                                new docx.TableCell({
                                    children: [
                                        new docx.Paragraph({
                                            children: [
                                                new docx.TextRun({
                                                    text: "📅 " + formatDate(data.exportDate),
                                                    size: 14,
                                                    color: this.colors.textLight
                                                })
                                            ],
                                            alignment: docx.AlignmentType.LEFT,
                                            spacing: { before: 50, after: 0 }
                                        })
                                    ],
                                    borders: this.noBorders(),
                                    verticalAlign: docx.VerticalAlign.CENTER,
                                    width: { size: 33, type: docx.WidthType.PERCENTAGE }
                                }),
                                new docx.TableCell({
                                    children: [
                                        new docx.Paragraph({
                                            children: [
                                                new docx.TextRun({
                                                    text: "Page ",
                                                    size: 14,
                                                    color: this.colors.textLight
                                                }),
                                                new docx.SimpleField(`PAGE`),
                                                new docx.TextRun({
                                                    text: " sur ",
                                                    size: 14,
                                                    color: this.colors.textLight
                                                }),
                                                new docx.SimpleField(`NUMPAGES`)
                                            ],
                                            alignment: docx.AlignmentType.CENTER,
                                            spacing: { before: 50, after: 0 }
                                        })
                                    ],
                                    borders: this.noBorders(),
                                    verticalAlign: docx.VerticalAlign.CENTER,
                                    width: { size: 34, type: docx.WidthType.PERCENTAGE }
                                }),
                                new docx.TableCell({
                                    children: [
                                        new docx.Paragraph({
                                            children: [
                                                new docx.TextRun({
                                                    text: "🔒 CONFIDENTIEL",
                                                    bold: true,
                                                    size: 14,
                                                    color: this.colors.danger
                                                })
                                            ],
                                            alignment: docx.AlignmentType.RIGHT,
                                            spacing: { before: 50, after: 0 }
                                        })
                                    ],
                                    borders: this.noBorders(),
                                    verticalAlign: docx.VerticalAlign.CENTER,
                                    width: { size: 33, type: docx.WidthType.PERCENTAGE }
                                })
                            ]
                        })
                    ]
                })
            ]
        });
    }
    
    noBorders() {
        return {
            top: { style: docx.BorderStyle.NONE },
            bottom: { style: docx.BorderStyle.NONE },
            left: { style: docx.BorderStyle.NONE },
            right: { style: docx.BorderStyle.NONE }
        };
    }
    
    generateDocumentTitle(data) {
        const numeroDevis = data.numeroDevis || '001';
        const typeProjet = data.typeProjet || data.typeDevis || 'Note_de_Cadrage';
        return `${data.client || 'CLIENT'}_${data.annee || new Date().getFullYear()}-${data.numeroProjet || 'XXXXXX'}-${numeroDevis}_${data.trigramme || 'XXX'}_${data.libelle || 'Projet'}_${typeProjet}`;
    }
    
    createEnhancedDocumentStyles() {
        return {
            default: {
                heading1: {
                    run: {
                        size: 40, // 20pt
                        bold: true,
                        color: this.colors.primary,
                        font: "Calibri Light"
                    },
                    paragraph: {
                        spacing: {
                            before: 0,
                            after: 240,
                            line: 360
                        },
                        outlineLevel: 0,
                        keepNext: true,
                        keepLines: true,
                        border: {
                            bottom: {
                                color: this.colors.primaryLight,
                                space: 1,
                                size: 12,
                                style: docx.BorderStyle.SINGLE
                            }
                        }
                    }
                },
                heading2: {
                    run: {
                        size: 32, // 16pt
                        bold: true,
                        color: this.colors.primaryDark,
                        font: "Calibri"
                    },
                    paragraph: {
                        spacing: {
                            before: 360,
                            after: 180,
                            line: 360
                        },
                        outlineLevel: 1,
                        keepNext: true,
                        keepLines: true
                    }
                },
                heading3: {
                    run: {
                        size: 28, // 14pt
                        bold: true,
                        color: this.colors.accent,
                        font: "Calibri"
                    },
                    paragraph: {
                        spacing: {
                            before: 240,
                            after: 120,
                            line: 360
                        },
                        outlineLevel: 2,
                        keepNext: true,
                        keepLines: true
                    }
                },
                document: {
                    run: {
                        size: 22, // 11pt
                        font: "Calibri",
                        color: this.colors.text
                    },
                    paragraph: {
                        spacing: {
                            line: 360,
                            after: 200
                        },
                        alignment: docx.AlignmentType.JUSTIFIED
                    }
                }
            },
            paragraphStyles: [
                {
                    id: "Heading1",
                    name: "Heading 1",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: {
                        size: 40,
                        bold: true,
                        color: this.colors.primary,
                        font: "Calibri Light"
                    },
                    paragraph: {
                        spacing: {
                            before: 0,
                            after: 240,
                            line: 360
                        },
                        outlineLevel: 0,
                        keepNext: true,
                        keepLines: true,
                        border: {
                            bottom: {
                                color: this.colors.primaryLight,
                                space: 1,
                                size: 12,
                                style: docx.BorderStyle.SINGLE
                            }
                        }
                    }
                },
                {
                    id: "Heading2",
                    name: "Heading 2",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: {
                        size: 32,
                        bold: true,
                        color: this.colors.primaryDark,
                        font: "Calibri"
                    },
                    paragraph: {
                        spacing: {
                            before: 360,
                            after: 180,
                            line: 360
                        },
                        outlineLevel: 1,
                        keepNext: true,
                        keepLines: true
                    }
                },
                {
                    id: "Heading3",
                    name: "Heading 3",
                    basedOn: "Normal",
                    next: "Normal",
                    quickFormat: true,
                    run: {
                        size: 28,
                        bold: true,
                        color: this.colors.accent,
                        font: "Calibri"
                    },
                    paragraph: {
                        spacing: {
                            before: 240,
                            after: 120,
                            line: 360
                        },
                        outlineLevel: 2,
                        keepNext: true,
                        keepLines: true
                    }
                },
                {
                    id: "CoverTitle",
                    name: "Cover Title",
                    basedOn: "Normal",
                    run: {
                        size: 52,
                        bold: true,
                        color: this.colors.primary,
                        font: "Calibri Light"
                    },
                    paragraph: {
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { before: 400, after: 300 }
                    }
                },
                {
                    id: "CoverSubtitle",
                    name: "Cover Subtitle",
                    basedOn: "Normal",
                    run: {
                        size: 32,
                        color: this.colors.primaryDark,
                        font: "Calibri Light"
                    },
                    paragraph: {
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { before: 150, after: 300 }
                    }
                }
            ]
        };
    }
    
    // Document sections creation methods
createCoverPage(data, title) {
    const coverElements = [];
    
    // Barre supérieure décorative
    coverElements.push(
        new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            children: [
                new docx.TextRun({
                    text: " ",
                    size: 24
                })
            ],
            shading: {
                fill: this.colors.primary,
                type: docx.ShadingType.CLEAR
            },
            spacing: { before: 0, after: 0 },
            border: {
                bottom: { style: docx.BorderStyle.NONE }
            }
        })
    );
    
    // Espacement
    coverElements.push(
        new docx.Paragraph({
            text: "",
            spacing: { after: 600 }
        })
    );
    
    // Type de document avec style badge moderne
    coverElements.push(
        new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            children: [
                new docx.TextRun({
                    text: "DOCUMENT STRATÉGIQUE",
                    size: 16,
                    color: this.colors.accent,
                    font: "Calibri",
                    allCaps: true,
                    characterSpacing: 120
                })
            ],
            spacing: { after: 300 }
        })
    );
    
    // Titre principal avec effet moderne
    coverElements.push(
        new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            children: [
                new docx.TextRun({
                    text: "NOTE DE",
                    size: 60,
                    color: this.colors.text,
                    font: "Calibri Light"
                })
            ],
            spacing: { after: 0 }
        }),
        new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            children: [
                new docx.TextRun({
                    text: "CADRAGE",
                    bold: true,
                    size: 72,
                    color: this.colors.primary,
                    font: "Calibri",
                    allCaps: true,
                    characterSpacing: 200
                })
            ],
            spacing: { after: 400 }
        })
    );
    
    // Ligne décorative moderne avec gradient simulé
    const decorativeLine = new docx.Table({
        width: {
            size: 60,
            type: docx.WidthType.PERCENTAGE
        },
        alignment: docx.AlignmentType.CENTER,
        borders: {
            top: { style: docx.BorderStyle.NONE },
            bottom: { style: docx.BorderStyle.NONE },
            left: { style: docx.BorderStyle.NONE },
            right: { style: docx.BorderStyle.NONE }
        },
        rows: [
            new docx.TableRow({
                height: {
                    value: 60,
                    rule: docx.HeightRule.EXACT
                },
                children: [
                    new docx.TableCell({
                        children: [new docx.Paragraph({ text: "" })],
                        shading: { fill: this.colors.primaryLight },
                        borders: this.noBorders(),
                        width: { size: 20, type: docx.WidthType.PERCENTAGE }
                    }),
                    new docx.TableCell({
                        children: [new docx.Paragraph({ text: "" })],
                        shading: { fill: this.colors.accent },
                        borders: this.noBorders(),
                        width: { size: 60, type: docx.WidthType.PERCENTAGE }
                    }),
                    new docx.TableCell({
                        children: [new docx.Paragraph({ text: "" })],
                        shading: { fill: this.colors.primary },
                        borders: this.noBorders(),
                        width: { size: 20, type: docx.WidthType.PERCENTAGE }
                    })
                ]
            })
        ]
    });
    
    coverElements.push(decorativeLine);
    
    // Espacement
    coverElements.push(
        new docx.Paragraph({
            text: "",
            spacing: { after: 600 }
        })
    );
    
    // Carte projet moderne avec ombrage
    const projectInfoCard = new docx.Table({
        width: {
            size: 85,
            type: docx.WidthType.PERCENTAGE
        },
        alignment: docx.AlignmentType.CENTER,
        borders: {
            top: { style: docx.BorderStyle.SINGLE, size: 6, color: this.colors.primary },
            bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: this.colors.primary },
            left: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
            right: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight }
        },
        rows: [
            // Nom du projet
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: data.libelle || "Projet",
                                        bold: true,
                                        size: 32,
                                        color: this.colors.primaryDark,
                                        font: "Calibri"
                                    })
                                ],
                                spacing: { before: 200, after: 100 }
                            })
                        ],
                        borders: this.noBorders(),
                        shading: {
                            fill: this.colors.background,
                            type: docx.ShadingType.CLEAR
                        }
                    })
                ]
            }),
            // Client
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: "Pour",
                                        size: 18,
                                        color: this.colors.textLight,
                                        font: "Calibri Light",
                                        italics: true
                                    })
                                ],
                                spacing: { after: 50 }
                            }),
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: data.client || "CLIENT",
                                        size: 28,
                                        bold: true,
                                        color: this.colors.text,
                                        font: "Calibri"
                                    })
                                ],
                                spacing: { after: 200 }
                            })
                        ],
                        borders: this.noBorders()
                    })
                ]
            })
        ]
    });
    
    coverElements.push(projectInfoCard);
    
    // Espacement
    coverElements.push(
        new docx.Paragraph({
            text: "",
            spacing: { after: 500 }
        })
    );
    
    // Métadonnées en grille moderne
    const metadataGrid = new docx.Table({
        width: {
            size: 85,
            type: docx.WidthType.PERCENTAGE
        },
        alignment: docx.AlignmentType.CENTER,
        borders: {
            top: { style: docx.BorderStyle.NONE },
            bottom: { style: docx.BorderStyle.NONE },
            left: { style: docx.BorderStyle.NONE },
            right: { style: docx.BorderStyle.NONE },
            insideHorizontal: { style: docx.BorderStyle.NONE },
            insideVertical: { style: docx.BorderStyle.NONE }
        },
        rows: [
            new docx.TableRow({
                children: [
                    // Référence
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: "RÉFÉRENCE",
                                        size: 12,
                                        color: this.colors.textLight,
                                        font: "Calibri",
                                        allCaps: true,
                                        characterSpacing: 80
                                    })
                                ],
                                spacing: { after: 30 }
                            }),
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: `${data.annee || new Date().getFullYear()}-${data.numeroProjet || 'XXXXXX'}-${data.numeroDevis || '001'}`,
                                        size: 22,
                                        bold: true,
                                        color: this.colors.primary,
                                        font: "Consolas"
                                    })
                                ],
                                spacing: { after: 100 }
                            })
                        ],
                        borders: {
                            right: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight }
                        },
                        width: { size: 33.33, type: docx.WidthType.PERCENTAGE }
                    }),
                    // Type
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: "TYPE",
                                        size: 12,
                                        color: this.colors.textLight,
                                        font: "Calibri",
                                        allCaps: true,
                                        characterSpacing: 80
                                    })
                                ],
                                spacing: { after: 30 }
                            }),
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: (data.project_type || data.typeDevis || "À définir").toUpperCase(),
                                        size: 22,
                                        bold: true,
                                        color: this.colors.accent,
                                        font: "Calibri"
                                    })
                                ],
                                spacing: { after: 100 }
                            })
                        ],
                        borders: {
                            right: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight }
                        },
                        width: { size: 33.34, type: docx.WidthType.PERCENTAGE }
                    }),
                    // Date
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: "DATE",
                                        size: 12,
                                        color: this.colors.textLight,
                                        font: "Calibri",
                                        allCaps: true,
                                        characterSpacing: 80
                                    })
                                ],
                                spacing: { after: 30 }
                            }),
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: formatDate(data.exportDate),
                                        size: 22,
                                        bold: true,
                                        color: this.colors.primaryDark,
                                        font: "Calibri"
                                    })
                                ],
                                spacing: { after: 100 }
                            })
                        ],
                        borders: this.noBorders(),
                        width: { size: 33.33, type: docx.WidthType.PERCENTAGE }
                    })
                ]
            })
        ]
    });
    
    coverElements.push(metadataGrid);
    
    // Espacement
    coverElements.push(
        new docx.Paragraph({
            text: "",
            spacing: { after: 400 }
        })
    );
    
    // Mention confidentiel élégante
    coverElements.push(
        new docx.Paragraph({
            alignment: docx.AlignmentType.CENTER,
            children: [
                new docx.TextRun({
                    text: "— ",
                    size: 20,
                    color: this.colors.danger
                }),
                new docx.TextRun({
                    text: "CONFIDENTIEL",
                    size: 20,
                    bold: true,
                    color: this.colors.danger,
                    font: "Calibri",
                    characterSpacing: 100
                }),
                new docx.TextRun({
                    text: " —",
                    size: 20,
                    color: this.colors.danger
                })
            ],
            spacing: { after: 400 }
        })
    );
    
    // Footer moderne avec logo stylisé
    const footerTable = new docx.Table({
        width: {
            size: 100,
            type: docx.WidthType.PERCENTAGE
        },
        borders: {
            top: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            bottom: { style: docx.BorderStyle.NONE },
            left: { style: docx.BorderStyle.NONE },
            right: { style: docx.BorderStyle.NONE }
        },
        rows: [
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: "CA-GIP",
                                        size: 26,
                                        bold: true,
                                        color: this.colors.primary,
                                        font: "Calibri Light",
                                        characterSpacing: 150
                                    })
                                ],
                                spacing: { before: 150, after: 50 }
                            }),
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: "Excellence • Innovation • Confiance",
                                        size: 16,
                                        italics: true,
                                        color: this.colors.textLight,
                                        font: "Calibri Light"
                                    })
                                ],
                                spacing: { after: 100 }
                            }),
                            new docx.Paragraph({
                                alignment: docx.AlignmentType.CENTER,
                                children: [
                                    new docx.TextRun({
                                        text: `Version 1.0 | Équipe ${data.trigramme || "XXX"}`,
                                        size: 14,
                                        color: this.colors.textLight,
                                        font: "Calibri Light"
                                    })
                                ],
                                spacing: { after: 50 }
                            })
                        ],
                        borders: this.noBorders()
                    })
                ]
            })
        ]
    });
    
    coverElements.push(footerTable);
    
    return coverElements;
}
    
    createInfoRow(label, value, highlight = false) {
        return new docx.TableRow({
            children: [
                new docx.TableCell({
                    children: [
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: label,
                                    bold: true,
                                    size: 20,
                                    color: this.colors.primaryDark
                                })
                            ],
                            alignment: docx.AlignmentType.RIGHT
                        })
                    ],
                    shading: {
                        fill: this.colors.background,
                        type: docx.ShadingType.CLEAR
                    },
                    width: { size: 40, type: docx.WidthType.PERCENTAGE },
                    margins: {
                        top: 180,
                        bottom: 180,
                        left: 200,
                        right: 200
                    }
                }),
                new docx.TableCell({
                    children: [
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: value,
                                    size: 20,
                                    bold: highlight,
                                    color: highlight ? this.colors.danger : this.colors.text
                                })
                            ],
                            alignment: docx.AlignmentType.LEFT
                        })
                    ],
                    shading: highlight ? {
                        fill: this.colors.dangerLight,
                        type: docx.ShadingType.CLEAR
                    } : null,
                    width: { size: 60, type: docx.WidthType.PERCENTAGE },
                    margins: {
                        top: 180,
                        bottom: 180,
                        left: 200,
                        right: 200
                    }
                })
            ]
        });
    }

createSummary(data, title) {
    const summaryElements = [];
    
    summaryElements.push(
        this.createHeading("Synthèse", 1, false)
    );
    
    // Tableau principal du résumé exécutif
    const summaryTable = new docx.Table({
        width: {
            size: 100,
            type: docx.WidthType.PERCENTAGE
        },
        borders: {
            top: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
            bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
            left: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
            right: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
            insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            insideVertical: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight }
        },
        rows: [
            // Ligne "RÉSUMÉ EXÉCUTIF"
            new docx.TableRow({
                children: [
                    this.createTableCell("📊 RÉSUMÉ EXÉCUTIF", {
                        columnSpan: 2,
                        shading: this.colors.primary,
                        bold: true,
                        alignment: docx.AlignmentType.CENTER,
                        textColor: this.colors.white,
                        size: 24
                    })
                ],
                height: {
                    value: 600,
                    rule: docx.HeightRule.MINIMUM
                }
            }),
            // Ligne titre du projet
            new docx.TableRow({
                children: [
                    this.createTableCell(data.libelle || title, {
                        columnSpan: 2,
                        shading: this.colors.primary,
                        bold: true,
                        alignment: docx.AlignmentType.CENTER,
                        textColor: this.colors.white,
                        size: 20
                    })
                ],
                height: {
                    value: 500,
                    rule: docx.HeightRule.MINIMUM
                }
            }),
            // Ligne Entité / Porteur
            new docx.TableRow({
                children: [
                    this.createTableCell("🏢 Entité", { 
                        shading: this.colors.background, 
                        bold: true 
                    }),
                    this.createTableCell("👥 Porteur", { 
                        shading: this.colors.background, 
                        bold: true 
                    })
                ]
            }),
            // Ligne Client / Cluster
            new docx.TableRow({
                children: [
                    this.createTableCell(`${data.client || 'À définir'}`, { bold: true }),
                    this.createTableCell(`Cluster ${data.trigramme || 'XXX'}`, { bold: true })
                ]
            }),
            // Ligne Coûts headers
            new docx.TableRow({
                children: [
                    this.createTableCell("💶 Coûts de mise en œuvre", { 
                        bold: true, 
                        shading: this.colors.background 
                    }),
                    this.createTableCell("📈 Coûts de fonctionnement", { 
                        bold: true, 
                        shading: this.colors.background 
                    })
                ]
            }),
            // Ligne des montants principaux
            new docx.TableRow({
                height: {
                    value: 600,
                    rule: docx.HeightRule.MINIMUM
                },
                children: [
                    this.createTableCell([
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: formatCurrency(data.totalHorsContingence || 0),
                                    bold: true,
                                    size: 28,
                                    color: this.colors.primary
                                })
                            ],
                            alignment: docx.AlignmentType.CENTER,
                            spacing: { before: 50, after: 30 }
                        }),
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: `(${formatCurrency(data.totalAvecContingence || 0)} avec contingence)`,
                                    italics: true,
                                    size: 16,
                                    color: this.colors.accent
                                })
                            ],
                            alignment: docx.AlignmentType.CENTER,
                            spacing: { after: 50 }
                        })
                    ], { verticalAlign: docx.VerticalAlign.CENTER }),
                    this.createTableCell([
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: formatCurrency(data.totalFonctionnement || 0) + " /an",
                                    bold: true,
                                    size: 28,
                                    color: this.colors.primary
                                })
                            ],
                            alignment: docx.AlignmentType.CENTER,
                            spacing: { before: 50, after: 30 }
                        }),
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: `(${formatCurrency(data.totalFonctionnement3ans || 0)} sur 3 ans)`,
                                    italics: true,
                                    size: 16,
                                    color: this.colors.accent
                                })
                            ],
                            alignment: docx.AlignmentType.CENTER,
                            spacing: { after: 50 }
                        })
                    ], { verticalAlign: docx.VerticalAlign.CENTER })
                ]
            }),
            // Ligne TCO
            new docx.TableRow({
                children: [
                    this.createTableCell("💰 TCO 3 ans", {
                        shading: this.colors.primaryDark,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    }),
                    this.createTableCell(formatCurrency(data.tco3ans || 0), {
                        shading: this.colors.primaryDark,
                        bold: true,
                        size: 24,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    })
                ],
                height: {
                    value: 500,
                    rule: docx.HeightRule.MINIMUM
                }
            })
        ]
    });
    
    summaryElements.push(summaryTable);
    
    // Espacement réduit
    summaryElements.push(
        new docx.Paragraph({
            text: "",
            spacing: { after: 300 }
        })
    );
    
    // Indicateurs clés en format compact
    summaryElements.push(this.createHeading("Indicateurs clés", 3));
    
    const metricsTable = new docx.Table({
        width: {
            size: 100,
            type: docx.WidthType.PERCENTAGE
        },
        borders: {
            top: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            bottom: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            left: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            right: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            insideHorizontal: { style: docx.BorderStyle.NONE },
            insideVertical: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight }
        },
        rows: [
            new docx.TableRow({
                height: {
                    value: 600,
                    rule: docx.HeightRule.MINIMUM
                },
                children: [
                    this.createCompactMetricCell("⚡ Complexité", data.complexity || "standard", this.getComplexityColor(data.complexity)),
                    this.createCompactMetricCell("🏭 Secteur", data.sector || "À définir", this.colors.primary),
                    this.createCompactMetricCell("📝 Type", data.typeDevis || "À définir", this.colors.primary),
                    this.createCompactMetricCell("🧪 POC requis", data.besoinPOC || "non", data.besoinPOC === "oui" ? this.colors.warning : this.colors.accent)
                ]
            })
        ]
    });
    
    summaryElements.push(metricsTable);
    
    // Informations complémentaires
    summaryElements.push(this.createHeading("Informations du document", 3));
    
    const infoTable = new docx.Table({
        width: {
            size: 100,
            type: docx.WidthType.PERCENTAGE
        },
        borders: {
            top: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            bottom: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            left: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            right: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
            insideVertical: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight }
        },
        rows: [
            new docx.TableRow({
                children: [
                    this.createCompactInfoCell("📋 Référence", `${data.annee || new Date().getFullYear()}-${data.numeroProjet || 'XXXXXX'}-${data.numeroDevis || '001'}`),
                    this.createCompactInfoCell("👥 Équipe", data.trigramme || "XXX"),
                    this.createCompactInfoCell("📅 Date de diffusion", formatDate(data.exportDate)),
                    this.createCompactInfoCell("⏰ Fin de validité", formatDate(new Date(Date.now() + 30*24*60*60*1000)))
                ]
            })
        ]
    });
    
    summaryElements.push(infoTable);
    
    return summaryElements;
}

// Version compacte de createMetricCell
createCompactMetricCell(label, value, color = null) {
    const cellColor = color || this.colors.primary;
    
    return new docx.TableCell({
        children: [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: label,
                        size: 14,
                        color: this.colors.textLight
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 40 }
            }),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: value.toUpperCase(),
                        bold: true,
                        size: 18,
                        color: cellColor
                    })
                ],
                alignment: docx.AlignmentType.CENTER
            })
        ],
        borders: {
            top: { style: docx.BorderStyle.NONE },
            bottom: { style: docx.BorderStyle.NONE },
            left: { style: docx.BorderStyle.NONE },
            right: { style: docx.BorderStyle.NONE }
        },
        shading: {
            fill: this.colors.background,
            type: docx.ShadingType.CLEAR
        },
        margins: {
            top: 100,
            bottom: 100
        },
        verticalAlign: docx.VerticalAlign.CENTER
    });
}

// Version compacte de createInfoCell
createCompactInfoCell(label, value) {
    return new docx.TableCell({
        children: [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: label + " : ",
                        size: 14,
                        color: this.colors.textLight
                    }),
                    new docx.TextRun({
                        text: value,
                        bold: true,
                        size: 14,
                        color: this.colors.text
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { before: 60, after: 60 }
            })
        ],
        width: { size: 25, type: docx.WidthType.PERCENTAGE },
        verticalAlign: docx.VerticalAlign.CENTER
    });
}

// Ajout de la méthode manquante createTableCell avec size optionnel
createTableCell(content, options = {}) {
    let children;
    
    if (Array.isArray(content)) {
        children = content;
    } else {
        let alignment = options.alignment;
        if (!alignment) {
            if (options.shading === this.colors.primary || options.shading === this.colors.primaryDark) {
                alignment = docx.AlignmentType.CENTER;
            } else {
                alignment = docx.AlignmentType.LEFT;
            }
        }
        
        children = [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: content || '',
                        bold: options.bold,
                        size: options.size || 20,
                        font: "Calibri",
                        color: options.textColor || (options.shading === this.colors.primary || options.shading === this.colors.primaryDark ? this.colors.white : this.colors.text)
                    })
                ],
                alignment: alignment,
                spacing: { after: 0 }
            })
        ];
    }
    
    const cellOptions = {
        children: children,
        margins: {
            top: 120,
            bottom: 120,
            left: 120,
            right: 120
        },
        verticalAlign: options.verticalAlign || (
            (options.shading === this.colors.primary || options.shading === this.colors.primaryDark) 
                ? docx.VerticalAlign.CENTER 
                : docx.VerticalAlign.TOP
        )
    };
    
    if (options.shading) {
        cellOptions.shading = {
            fill: options.shading,
            type: docx.ShadingType.CLEAR
        };
    }
    
    if (options.columnSpan) {
        cellOptions.columnSpan = options.columnSpan;
    }
    
    return new docx.TableCell(cellOptions);
}

createMetricCell(label, value, color = null) {
    // Si aucune couleur n'est spécifiée, utiliser la couleur principale du thème
    const cellColor = color || this.colors.primary;
    
    return new docx.TableCell({
        children: [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: label,
                        size: 18,
                        color: this.colors.textLight
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 80 }
            }),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: value.toUpperCase(),
                        bold: true,
                        size: 24,
                        color: cellColor
                    })
                ],
                alignment: docx.AlignmentType.CENTER
            })
        ],
        borders: {
            top: { style: docx.BorderStyle.SINGLE, size: 3, color: this.colors.primaryLight },
            bottom: { style: docx.BorderStyle.SINGLE, size: 3, color: this.colors.primaryLight },
            left: { style: docx.BorderStyle.SINGLE, size: 3, color: this.colors.primaryLight },
            right: { style: docx.BorderStyle.SINGLE, size: 3, color: this.colors.primaryLight }
        },
        shading: {
            fill: this.colors.background,
            type: docx.ShadingType.CLEAR
        },
        margins: {
            top: 250,
            bottom: 250
        },
        verticalAlign: docx.VerticalAlign.CENTER  // Ajout de l'alignement vertical centré
    });
}
getComplexityColor(complexity) {
        const colors = {
            'simple': this.colors.accent,       // Utilise la couleur d'accent du thème
            'standard': this.colors.primary,    // Utilise la couleur principale
            'complexe': this.colors.warning,    // Orange pour le warning
            'très complexe': this.colors.danger // Rouge pour le danger
        };
        return colors[complexity] || this.colors.primaryDark;
    }    
    createProjectDescription(data) {
        const sections = [];
        
        sections.push(
            this.createHeading("I. Description du projet", 1, true),
            this.createHeading("I.1. Contexte", 2)
        );
        
        if (!data.contexte || data.contexte.trim() === '') {
            sections.push(
                this.createParagraph("PARAGRAPHE OBLIGATOIRE : A COMPLETER SI LE CONTEXTE DU PROJET LE NECESSITE.", { bold: true }),
                this.createParagraph("À compléter")
            );
        } else {
            sections.push(...this.parseAndCreateParagraphs(data.contexte));
        }
        
        sections.push(this.createHeading("I.2. Le besoin exprimé", 2));
        
        if (!data.besoin || data.besoin.trim() === '') {
            sections.push(
                this.createParagraph("PARAGRAPHE OBLIGATOIRE : A COMPLETER SI LE CONTEXTE DU PROJET LE NECESSITE.", { bold: true }),
                this.createParagraph("À compléter")
            );
        } else {
            sections.push(...this.parseAndCreateParagraphs(data.besoin));
        }
        
        sections.push(this.createHeading("I.3. Objectifs du projet", 2));
        
        if (!data.objectifs || data.objectifs.trim() === '') {
            sections.push(
                this.createParagraph("PARAGRAPHE OBLIGATOIRE : A COMPLETER SI LE CONTEXTE DU PROJET LE NECESSITE.", { bold: true }),
                this.createParagraph("À compléter")
            );
        } else {
            sections.push(...this.parseAndCreateParagraphs(data.objectifs));
        }
        
        return sections;
    }
    
    createScope(data) {
        const sections = [];
        
        sections.push(
            this.createHeading("II. Périmètre du projet", 1, true),
            this.createHeading("II.1. Périmètre", 2)
        );
        
        if (!data.perimetre || data.perimetre.trim() === '') {
            sections.push(
                this.createParagraph("PARAGRAPHE OBLIGATOIRE : A COMPLETER SI LE CONTEXTE DU PROJET LE NECESSITE.", { bold: true }),
                this.createParagraph("À compléter")
            );
        } else {
            sections.push(...this.parseAndCreateParagraphs(data.perimetre));
        }
        
        sections.push(this.createHeading("II.2. Hors périmètre", 2));
        
        if (!data.horsPerimetre || data.horsPerimetre.trim() === '') {
            sections.push(
                this.createParagraph("PARAGRAPHE OBLIGATOIRE : A COMPLETER SI LE CONTEXTE DU PROJET LE NECESSITE.", { bold: true }),
                this.createParagraph("À compléter")
            );
        } else {
            sections.push(...this.parseAndCreateParagraphs(data.horsPerimetre));
        }
        
        return sections;
    }
    
    createSecurity(data) {
        return [
            this.createHeading("III. Sécurité et Réglementation", 1, true),
            this.createHeading("III.1. Projet Partenaire", 2),
            this.createParagraph(`DICP souhaité : ${data.dicp || 'À définir'}`),
            this.createParagraph(`Continuité d'activité: DIMA ${data.dima || 'X'} jours \\ PDMA ${data.pdma || 'X'} jours`),
            this.createHeading("III.2. Projet Interne", 2),
            this.createParagraph(`Analyse RGPD/AIVP requise: ${data.rgpd || 'non'}`),
            this.createHeading("III.3. Contrôle PSEE", 2),
            this.createParagraph(`Validation PSEE effectuée: ${data.psee || 'non'}`)
        ];
    }
    
    createConstraintsAndRisks(data) {
        const sections = [
            this.createHeading("IV. Contraintes et risques", 1, true),
            this.createHeading("IV.1. Contraintes et prérequis", 2)
        ];
        
        if (data.contraintes && data.contraintes.length > 0) {
            sections.push(this.createEnhancedTable(
                ['Type', 'Description', 'Criticité', 'Mitigation'],
                data.contraintes.map(c => [
                    c.col0 || '',
                    c.col1 || '',
                    c.col2 || '',
                    c.col3 || ''
                ])
            ));
        } else {
            sections.push(this.createParagraph("Aucune contrainte identifiée"));
        }
        
        sections.push(this.createHeading("IV.2. Risques projet", 2));
        
        if (data.risques && data.risques.length > 0) {
            sections.push(this.createEnhancedTable(
                ['Risque', 'Probabilité', 'Impact', 'Plan d\'action', 'Responsable'],
                data.risques.map(r => [
                    r.col0 || '',
                    r.col1 || '',
                    r.col2 || '',
                    r.col3 || '',
                    r.col4 || ''
                ])
            ));
        } else {
            sections.push(this.createParagraph("Aucun risque identifié"));
        }
        
        return sections;
    }
    async createTechnicalSolution(data) {
        const sections = [
            this.createHeading("V. Description technique de la solution", 1, true),
            this.createHeading("V.1. Description de la solution", 2)
        ];
        
        if (data.descriptionSolution) {
            sections.push(...this.parseAndCreateParagraphs(data.descriptionSolution));
        } else {
            sections.push(this.createParagraph("À compléter"));
        }
        
        sections.push(
            this.createHeading("Besoin d'un POC", 3),
            this.createParagraph(`POC nécessaire: ${data.besoinPOC || 'non'}`),
            this.createHeading("V.2. Architecture", 2)
        );
        
        if (data.architecture) {
            sections.push(...this.parseAndCreateParagraphs(data.architecture));
        } else {
            sections.push(this.createParagraph("À compléter"));
        }
        
        // Ajouter le schéma d'architecture s'il existe
        if (data.architectureSchema) {
            sections.push(
                this.createHeading("Schéma d'architecture", 3),
                await this.createImageFromBase64(data.architectureSchema, "Schéma d'architecture technique")
            );
        }
        
        sections.push(this.createHeading("V.3. Composants et dimensionnement", 2));
        
        if (data.composantsDimensionnement) {
            sections.push(...this.parseAndCreateParagraphs(data.composantsDimensionnement));
        } else {
            sections.push(this.createParagraph("À compléter"));
        }
        
        return sections;
    }
    
    async createImplementation(data) {
        const sections = [
            this.createHeading("VI. Démarche pour la mise en œuvre", 1, true),
            this.createHeading("VI.1. Lotissement", 2)
        ];
        
        if (data.lots && data.lots.length > 0) {
            sections.push(this.createEnhancedTable(
                ['Lot', 'Description', 'Durée', 'Début', 'Fin', 'Ressources'],
                data.lots.map(l => [
                    l.col0 || '',
                    l.col1 || '',
                    `${l.col2 || ''} j`,
                    formatDate(l.col3) || '',
                    formatDate(l.col4) || '',
                    l.col5 || ''
                ])
            ));
        } else {
            sections.push(this.createParagraph("Lotissement à définir"));
        }
        
        sections.push(this.createHeading("VI.2. Livrables du projet", 2));
        
        if (data.livrables && data.livrables.length > 0) {
            sections.push(this.createEnhancedTable(
                ['Livrable', 'Description', 'Date', 'Responsable'],
                data.livrables.map(l => [
                    l.col0 || '',
                    l.col1 || '',
                    formatDate(l.col2) || '',
                    l.col3 || ''
                ])
            ));
        } else {
            sections.push(this.createParagraph("Livrables à définir"));
        }
        
        sections.push(this.createHeading("VI.3. Jalons clés du projet", 2));
        
        if (data.jalons && data.jalons.length > 0) {
            sections.push(this.createEnhancedTable(
                ['Jalon', 'Type', 'Date', 'Critères'],
                data.jalons.map(j => [
                    j.col0 || '',
                    j.col1 || '',
                    formatDate(j.col2) || '',
                    j.col3 || ''
                ])
            ));
        } else {
            sections.push(this.createParagraph("Jalons à définir"));
        }
        
        // Ajouter le diagramme de Gantt s'il existe
        if (data.ganttDiagram) {
            sections.push(
                this.createHeading("VI.4. Planning Gantt", 2),
                this.createParagraph("Le diagramme ci-dessous présente la planification temporelle des lots et jalons du projet."),
                await this.createImageFromBase64(data.ganttDiagram, "Diagramme de Gantt du projet")
            );
        }
        
        return sections;
    }
    
    createOperationalService(data) {
        const sections = [
            this.createHeading("VII. Offre de service en fonctionnement", 1, true),
            this.createHeading("VII.1. Validation de la solution", 2),
            this.createParagraph("L'infrastructure retenue est validée par le comité de validation technique des infrastructures."),
            this.createHeading("VII.2. Niveaux de service", 2),
            this.createHeading("Conditions de fonctionnement hors crash site", 3)
        ];
        
        if (data.conditionsHorsCrash) {
            sections.push(...this.parseAndCreateParagraphs(data.conditionsHorsCrash));
        } else {
            sections.push(this.createParagraph("À compléter"));
        }
        
        sections.push(this.createHeading("Conditions de fonctionnement en cas de crash site", 3));
        
        if (data.conditionsCrashSite) {
            sections.push(...this.parseAndCreateParagraphs(data.conditionsCrashSite));
        } else {
            sections.push(this.createParagraph("À compléter"));
        }
        
        sections.push(
            this.createHeading("La plage de service", 3),
            this.createParagraph(`Plage de service: ${data.plageService || 'À définir'}`),
            this.createHeading("Résilience applicative", 3)
        );
        
        if (data.resilienceApplicative) {
            sections.push(...this.parseAndCreateParagraphs(data.resilienceApplicative));
        } else {
            sections.push(this.createParagraph("À compléter"));
        }
        
        sections.push(this.createHeading("Plan de reprise d'activité / Gestion du dégradé", 3));
        
        if (data.praPlanDegrade) {
            sections.push(...this.parseAndCreateParagraphs(data.praPlanDegrade));
        } else {
            sections.push(this.createParagraph("À compléter"));
        }
        
        sections.push(this.createHeading("Sauvegardes", 3));
        
        if (data.sauvegardes) {
            sections.push(...this.parseAndCreateParagraphs(data.sauvegardes));
        } else {
            sections.push(this.createParagraph("À compléter"));
        }
        
        sections.push(this.createHeading("Administration / Supervision", 3));
        
        if (data.administrationSupervision) {
            sections.push(...this.parseAndCreateParagraphs(data.administrationSupervision));
        } else {
            sections.push(this.createParagraph("À compléter"));
        }
        
        return sections;
    }
    
createFinancialEvaluation(data) {
    const sections = [
        this.createHeading("VIII. Évaluation financière", 1, true)
    ];
    
    // Introduction avec style
    sections.push(
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: "Cette section présente l'évaluation financière complète du projet, incluant les coûts de construction (BUILD) et de fonctionnement (RUN).",
                    size: 22,
                    italics: true,
                    color: this.colors.textLight
                })
            ],
            spacing: { after: 300 }
        }),
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: "💡 Note importante : ",
                    bold: true,
                    size: 20,
                    color: this.colors.warning
                }),
                new docx.TextRun({
                    text: "Les coûts présentés sont hors charges transverses.",
                    size: 20,
                    italics: true,
                    color: this.colors.text
                })
            ],
            spacing: { after: 400 },
            shading: {
                fill: this.colors.background,
                type: docx.ShadingType.CLEAR
            },
            indent: { left: 200, right: 200 }
        })
    );
    
    // VIII.1. Coûts du projet
    sections.push(
        this.createHeading("VIII.1. Coûts de construction (BUILD)", 2),
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: `Taux de contingence appliqué : `,
                    size: 20
                }),
                new docx.TextRun({
                    text: `${data.tauxContingence || '15'}%`,
                    bold: true,
                    size: 20,
                    color: this.colors.primary
                })
            ],
            spacing: { after: 300 }
        })
    );
    
    // Tableau des coûts de construction
    if (data.coutsConstruction && data.coutsConstruction.length > 0) {
        // Créer toutes les lignes d'abord
        const constructionRows = [];
        
        // En-tête
        constructionRows.push(
            new docx.TableRow({
                children: [
                    this.createTableCell("Profil", {
                        shading: this.colors.primaryDark,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    }),
                    this.createTableCell("Code équipe", {
                        shading: this.colors.primaryDark,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    }),
                    this.createTableCell("Charge (j.h)", {
                        shading: this.colors.primaryDark,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    }),
                    this.createTableCell("TJM (€)", {
                        shading: this.colors.primaryDark,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    }),
                    this.createTableCell("Total HT (€)", {
                        shading: this.colors.primaryDark,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    })
                ],
                tableHeader: true
            })
        );
        
        // Lignes de données
        data.coutsConstruction.forEach((cout, index) => {
            const isEven = index % 2 === 0;
            constructionRows.push(
                new docx.TableRow({
                    children: [
                        this.createTableCell(cout.col0 || '', {
                            shading: isEven ? this.colors.white : this.colors.background
                        }),
                        this.createTableCell(cout.col4 || '', {
                            shading: isEven ? this.colors.white : this.colors.background,
                            alignment: docx.AlignmentType.CENTER
                        }),
                        this.createTableCell(cout.col1 || '', {
                            shading: isEven ? this.colors.white : this.colors.background,
                            alignment: docx.AlignmentType.CENTER
                        }),
                        this.createTableCell(formatCurrency(parseFloat(cout.col2) || 0), {
                            shading: isEven ? this.colors.white : this.colors.background,
                            alignment: docx.AlignmentType.RIGHT
                        }),
                        this.createTableCell(formatCurrency(parseFloat(cout.col3) || 0), {
                            shading: isEven ? this.colors.white : this.colors.background,
                            alignment: docx.AlignmentType.RIGHT,
                            bold: true
                        })
                    ]
                })
            );
        });
        
        // Lignes de totaux
        constructionRows.push(
            // Sous-total
            new docx.TableRow({
                children: [
                    this.createTableCell("Sous-total hors contingence", {
                        columnSpan: 4,
                        shading: this.colors.primaryLight,
                        bold: true,
                        alignment: docx.AlignmentType.RIGHT
                    }),
                    this.createTableCell(formatCurrency(data.totalHorsContingence || 0), {
                        shading: this.colors.primaryLight,
                        bold: true,
                        alignment: docx.AlignmentType.RIGHT,
                        size: 22
                    })
                ]
            }),
            // Contingence
            new docx.TableRow({
                children: [
                    this.createTableCell(`Contingence (${data.tauxContingence || '15'}%)`, {
                        columnSpan: 4,
                        shading: this.colors.background,
                        italics: true,
                        alignment: docx.AlignmentType.RIGHT,
                        textColor: this.colors.accent
                    }),
                    this.createTableCell(formatCurrency(data.contingence || 0), {
                        shading: this.colors.background,
                        italics: true,
                        alignment: docx.AlignmentType.RIGHT,
                        textColor: this.colors.accent
                    })
                ]
            }),
            // Total
            new docx.TableRow({
                children: [
                    this.createTableCell("TOTAL CONSTRUCTION", {
                        columnSpan: 4,
                        shading: this.colors.primary,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.RIGHT,
                        size: 20
                    }),
                    this.createTableCell(formatCurrency(data.totalAvecContingence || 0), {
                        shading: this.colors.primary,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.RIGHT,
                        size: 24
                    })
                ]
            })
        );
        
        // Créer la table avec toutes les lignes
        const constructionTable = new docx.Table({
            width: {
                size: 100,
                type: docx.WidthType.PERCENTAGE
            },
            borders: {
                top: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                left: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                right: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
                insideVertical: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight }
            },
            rows: constructionRows
        });
        
        sections.push(constructionTable);
    } else {
        sections.push(this.createParagraph("Coûts de construction à définir"));
    }
    
    // Espacement
    sections.push(new docx.Paragraph({ text: "", spacing: { after: 600 } }));
    
    // VIII.2. Coûts de fonctionnement
    sections.push(
        this.createHeading("VIII.2. Coûts de fonctionnement (RUN)", 2),
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: "Coûts opérationnels récurrents annuels",
                    size: 20,
                    italics: true,
                    color: this.colors.textLight
                })
            ],
            spacing: { after: 300 }
        })
    );
    
    // Tableau des coûts de fonctionnement
    if (data.coutsFonctionnement && data.coutsFonctionnement.length > 0) {
        // Créer toutes les lignes d'abord
        const fonctionnementRows = [];
        
        // En-tête
        fonctionnementRows.push(
            new docx.TableRow({
                children: [
                    this.createTableCell("Poste", {
                        shading: this.colors.accent,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    }),
                    this.createTableCell("Code UO", {
                        shading: this.colors.accent,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    }),
                    this.createTableCell("Quantité", {
                        shading: this.colors.accent,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    }),
                    this.createTableCell("Prix unitaire/an (€)", {
                        shading: this.colors.accent,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    }),
                    this.createTableCell("Total annuel (€)", {
                        shading: this.colors.accent,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.CENTER
                    })
                ],
                tableHeader: true
            })
        );
        
        // Lignes de données
        data.coutsFonctionnement.forEach((cout, index) => {
            const isEven = index % 2 === 0;
            const quantite = parseFloat(cout.col1) || 0;
            const prixMensuel = parseFloat(cout.col2) || 0;
            const prixAnnuel = prixMensuel * 12;
            const totalAnnuel = quantite * prixAnnuel;
            
            fonctionnementRows.push(
                new docx.TableRow({
                    children: [
                        this.createTableCell(cout.col0 || '', {
                            shading: isEven ? this.colors.white : this.colors.background
                        }),
                        this.createTableCell(cout.col4 || '', {
                            shading: isEven ? this.colors.white : this.colors.background,
                            alignment: docx.AlignmentType.CENTER
                        }),
                        this.createTableCell(cout.col1 || '', {
                            shading: isEven ? this.colors.white : this.colors.background,
                            alignment: docx.AlignmentType.CENTER
                        }),
                        this.createTableCell(formatCurrency(prixAnnuel), {
                            shading: isEven ? this.colors.white : this.colors.background,
                            alignment: docx.AlignmentType.RIGHT
                        }),
                        this.createTableCell(formatCurrency(totalAnnuel), {
                            shading: isEven ? this.colors.white : this.colors.background,
                            alignment: docx.AlignmentType.RIGHT,
                            bold: true
                        })
                    ]
                })
            );
        });
        
        // Lignes de totaux
        fonctionnementRows.push(
            // Total annuel
            new docx.TableRow({
                children: [
                    this.createTableCell("Total Fonctionnement /an", {
                        columnSpan: 4,
                        shading: this.colors.primaryLight,
                        bold: true,
                        alignment: docx.AlignmentType.RIGHT
                    }),
                    this.createTableCell(formatCurrency(data.totalFonctionnement || 0), {
                        shading: this.colors.primaryLight,
                        bold: true,
                        alignment: docx.AlignmentType.RIGHT,
                        size: 22
                    })
                ]
            }),
            // Total 3 ans
            new docx.TableRow({
                children: [
                    this.createTableCell("Total Fonctionnement 3 ans", {
                        columnSpan: 4,
                        shading: this.colors.background,
                        bold: true,
                        alignment: docx.AlignmentType.RIGHT
                    }),
                    this.createTableCell(formatCurrency(data.totalFonctionnement3ans || 0), {
                        shading: this.colors.background,
                        bold: true,
                        alignment: docx.AlignmentType.RIGHT,
                        textColor: this.colors.accent
                    })
                ]
            }),
            // TCO
            new docx.TableRow({
                children: [
                    this.createTableCell("TCO 3 ANS (BUILD + RUN)", {
                        columnSpan: 4,
                        shading: this.colors.primaryDark,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.RIGHT,
                        size: 20
                    }),
                    this.createTableCell(formatCurrency(data.tco3ans || 0), {
                        shading: this.colors.primaryDark,
                        bold: true,
                        textColor: this.colors.white,
                        alignment: docx.AlignmentType.RIGHT,
                        size: 24
                    })
                ]
            })
        );
        
        // Créer la table avec toutes les lignes
        const fonctionnementTable = new docx.Table({
            width: {
                size: 100,
                type: docx.WidthType.PERCENTAGE
            },
            borders: {
                top: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                left: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                right: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
                insideVertical: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight }
            },
            rows: fonctionnementRows
        });
        
        sections.push(fonctionnementTable);
    } else {
        sections.push(this.createParagraph("Coûts de fonctionnement à définir"));
    }
    
    // Espacement
    sections.push(new docx.Paragraph({ text: "", spacing: { after: 600 } }));
    
    // VIII.3. Synthèse financière
    sections.push(this.createHeading("VIII.3. Synthèse financière", 2));
    
    // Tableau de synthèse visuelle
    const synthesisSummary = new docx.Table({
        width: {
            size: 100,
            type: docx.WidthType.PERCENTAGE
        },
        borders: {
            top: { style: docx.BorderStyle.NONE },
            bottom: { style: docx.BorderStyle.NONE },
            left: { style: docx.BorderStyle.NONE },
            right: { style: docx.BorderStyle.NONE },
            insideHorizontal: { style: docx.BorderStyle.NONE },
            insideVertical: { style: docx.BorderStyle.NONE }
        },
        rows: [
            new docx.TableRow({
                children: [
                    // Carte Investissement
                    this.createSummaryCard(
                        "💼 INVESTISSEMENT",
                        formatCurrency(data.totalAvecContingence || 0),
                        "Coût unique",
                        this.colors.primary
                    ),
                    // Espace
                    new docx.TableCell({
                        children: [new docx.Paragraph({ text: "" })],
                        borders: this.noBorders(),
                        width: { size: 5, type: docx.WidthType.PERCENTAGE }
                    }),
                    // Carte Fonctionnement
                    this.createSummaryCard(
                        "🔄 FONCTIONNEMENT",
                        formatCurrency(data.totalFonctionnement || 0) + "/an",
                        formatCurrency(data.totalFonctionnement3ans || 0) + " sur 3 ans",
                        this.colors.accent
                    ),
                    // Espace
                    new docx.TableCell({
                        children: [new docx.Paragraph({ text: "" })],
                        borders: this.noBorders(),
                        width: { size: 5, type: docx.WidthType.PERCENTAGE }
                    }),
                    // Carte TCO
                    this.createSummaryCard(
                        "💰 TCO TOTAL",
                        formatCurrency(data.tco3ans || 0),
                        "Sur 3 ans",
                        this.colors.primaryDark
                    )
                ]
            })
        ]
    });
    
    sections.push(synthesisSummary);
    
    // Note de fin
    sections.push(
        new docx.Paragraph({
            children: [
                new docx.TextRun({
                    text: "📊 ",
                    size: 20
                }),
                new docx.TextRun({
                    text: "Analyse financière : ",
                    bold: true,
                    size: 20,
                    color: this.colors.text
                }),
                new docx.TextRun({
                    text: `Le projet représente un investissement initial de ${formatCurrency(data.totalAvecContingence || 0)} avec des coûts opérationnels de ${formatCurrency(data.totalFonctionnement || 0)} par an. Le coût total de possession sur 3 ans s'élève à ${formatCurrency(data.tco3ans || 0)}.`,
                    size: 20,
                    color: this.colors.text
                })
            ],
            spacing: { before: 600, after: 200 },
            alignment: docx.AlignmentType.JUSTIFIED
        })
    );
    
    return sections;
}

// Méthode helper pour créer une carte de synthèse
createSummaryCard(title, amount, subtitle, bgColor) {
    return new docx.TableCell({
        children: [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: title,
                        bold: true,
                        size: 16,
                        color: this.colors.white,
                        font: "Calibri"
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { before: 150, after: 100 }
            }),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: amount,
                        bold: true,
                        size: 32,
                        color: this.colors.white,
                        font: "Calibri Light"
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: subtitle,
                        size: 14,
                        color: this.colors.white,
                        font: "Calibri Light",
                        italics: true
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 150 }
            })
        ],
        shading: {
            fill: bgColor,
            type: docx.ShadingType.CLEAR
        },
        borders: {
            top: { style: docx.BorderStyle.SINGLE, size: 6, color: this.colors.white },
            bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: this.colors.white },
            left: { style: docx.BorderStyle.SINGLE, size: 6, color: this.colors.white },
            right: { style: docx.BorderStyle.SINGLE, size: 6, color: this.colors.white }
        },
        margins: {
            top: 200,
            bottom: 200,
            left: 150,
            right: 150
        },
        width: { size: 30, type: docx.WidthType.PERCENTAGE }
    });
}
    
    createRSE(data) {
        const sections = [
            this.createHeading("IX. RSE - Impact CO2", 1, true)
        ];
        
        if (data.impactCO2) {
            sections.push(...this.parseAndCreateParagraphs(data.impactCO2));
        } else {
            sections.push(this.createParagraph("Impact CO2 à évaluer avec les équipes RSE."));
        }
        
        return sections;
    }
    
    createDocumentation(data) {
        const sections = [
            this.createHeading("X. Gestion de la documentation du projet", 1, true)
        ];
        
        sections.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({ text: "Modalités de partage:", bold: true }),
                ],
                spacing: { after: 100 }
            })
        );
        
        if (data.modalitesPartage) {
            sections.push(...this.parseAndCreateParagraphs(data.modalitesPartage));
        } else {
            sections.push(this.createParagraph("La gestion documentaire reste à définir."));
        }
        
        sections.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({ text: "Lien documentation:", bold: true }),
                ],
                spacing: { after: 100 }
            })
        );
        sections.push(this.createParagraph(data.lienDocumentation || "À définir"));
        
        return sections;
    }
    
    // Enhanced helper methods
    createHeading(text, level, pageBreak = false) {
        const headingLevel = level === 1 ? docx.HeadingLevel.HEADING_1 :
                           level === 2 ? docx.HeadingLevel.HEADING_2 :
                           docx.HeadingLevel.HEADING_3;
        
        const styleId = level === 1 ? "Heading1" :
                       level === 2 ? "Heading2" :
                       "Heading3";
        
        return new docx.Paragraph({
            text: text,
            heading: headingLevel,
            style: styleId,
            pageBreakBefore: pageBreak,
            spacing: {
                before: pageBreak && level === 1 ? 0 : (level === 1 ? 480 : level === 2 ? 360 : 240),
                after: level === 1 ? 240 : level === 2 ? 180 : 120
            }
        });
    }
    
// 1. Fonction createNumberingConfig simplifiée
createNumberingConfig() {
    return {
        config: [
            {
                reference: "bullet-list",
                levels: [
                    {
                        level: 0,
                        format: docx.LevelFormat.BULLET,
                        text: "•",
                        alignment: docx.AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: 720, hanging: 360 }
                            }
                        }
                    },
                    {
                        level: 1,
                        format: docx.LevelFormat.BULLET,
                        text: "○",
                        alignment: docx.AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: 1440, hanging: 360 }
                            }
                        }
                    },
                    {
                        level: 2,
                        format: docx.LevelFormat.BULLET,
                        text: "▪",
                        alignment: docx.AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: 2160, hanging: 360 }
                            }
                        }
                    }
                ]
            }
        ]
    };
}

// 2. Fonction parseAndCreateParagraphs avec approche simple et fiable
parseAndCreateParagraphs(text) {
    if (!text) return [this.createParagraph("")];
    
    const paragraphs = [];
    const lines = text.split('\n');
    
    lines.forEach(line => {
        // Conserver l'indentation originale
        const originalLine = line;
        const trimmedLine = line.trim();
        
        // Calculer le niveau d'indentation (2 espaces = 1 niveau)
        const leadingSpaces = originalLine.match(/^(\s*)/)[1].length;
        const indentLevel = Math.floor(leadingSpaces / 2);
        
        // Vérifier les différents types de listes
        const bulletMatch = trimmedLine.match(/^[•\-\*]\s*(.*)$/);
        const numberedMatch = trimmedLine.match(/^(\d+)[\.\)]\s*(.*)$/);
        const letterMatch = trimmedLine.match(/^([a-zA-Z])[\.\)]\s*(.*)$/);
        const romanMatch = trimmedLine.match(/^([ivxIVX]+)[\.\)]\s*(.*)$/);
        
        if (bulletMatch) {
            // Puce standard à plusieurs niveaux
            const bulletText = bulletMatch[1];
            const level = Math.min(indentLevel, 2); // Limiter à 3 niveaux (0, 1, 2)
            
            paragraphs.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: bulletText,
                            size: 22,
                            font: "Calibri",
                            color: this.colors.text
                        })
                    ],
                    bullet: {
                        level: level
                    },
                    spacing: { 
                        after: level === 0 ? 100 : 80,
                        line: 360 
                    }
                })
            );
        }
        else if (numberedMatch || letterMatch || romanMatch) {
            // Toutes les listes numérotées converties en paragraphes simples
            let fullText = "";
            
            if (numberedMatch) {
                fullText = numberedMatch[1] + ". " + numberedMatch[2];
            } else if (letterMatch) {
                fullText = letterMatch[1] + ") " + letterMatch[2];
            } else if (romanMatch) {
                fullText = romanMatch[1] + ". " + romanMatch[2];
            }
            
            // Créer un paragraphe simple avec tout le texte
            paragraphs.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: fullText,
                            size: 22,
                            font: "Calibri",
                            color: this.colors.text
                        })
                    ],
                    indent: {
                        left: 720 + (indentLevel * 720)  // Indentation progressive simple
                    },
                    spacing: { 
                        after: 100,
                        line: 360 
                    },
                    alignment: docx.AlignmentType.JUSTIFIED
                })
            );
        }
        else if (trimmedLine) {
            // Paragraphe normal
            const isBold = trimmedLine.startsWith('**') && trimmedLine.endsWith('**');
            const isHeading = trimmedLine.match(/^#+\s/);
            
            if (isBold) {
                // Texte en gras
                const boldText = trimmedLine.replace(/\*\*/g, '');
                paragraphs.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: boldText,
                                bold: true,
                                size: 24,
                                font: "Calibri",
                                color: this.colors.text
                            })
                        ],
                        spacing: { before: 200, after: 150, line: 360 }
                    })
                );
            }
            else if (isHeading) {
                // Titre avec # ## ###
                const headingLevel = trimmedLine.match(/^(#+)/)[1].length;
                const headingText = trimmedLine.replace(/^#+\s/, '');
                const fontSize = headingLevel === 1 ? 28 : headingLevel === 2 ? 24 : 22;
                
                paragraphs.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: headingText,
                                bold: true,
                                size: fontSize,
                                font: "Calibri",
                                color: this.colors.primaryDark
                            })
                        ],
                        spacing: { before: 300, after: 200, line: 360 }
                    })
                );
            }
            else {
                // Paragraphe normal avec indentation si nécessaire
                paragraphs.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: trimmedLine,
                                size: 22,
                                font: "Calibri",
                                color: this.colors.text
                            })
                        ],
                        spacing: { after: 200, line: 360 },
                        indent: indentLevel > 0 ? { left: indentLevel * 720 } : undefined,
                        alignment: docx.AlignmentType.JUSTIFIED
                    })
                );
            }
        }
        // Les lignes vides sont ignorées, ce qui créera naturellement des espaces entre les paragraphes
    });
    
    return paragraphs.length > 0 ? paragraphs : [this.createParagraph("")];
}
    
    createParagraph(text, options = {}) {
        const children = [];
        
        if (options.bold) {
            children.push(new docx.TextRun({
                text: text || '',
                bold: true,
                size: options.size || 22,
                font: options.font || "Calibri",
                color: options.color || this.colors.text
            }));
        } else {
            children.push(new docx.TextRun({
                text: text || '',
                size: options.size || 22,
                font: options.font || "Calibri",
                color: options.color || this.colors.text
            }));
        }
        
        return new docx.Paragraph({
            children: children,
            spacing: options.spacing || { after: 200, line: 360 },
            alignment: options.alignment || docx.AlignmentType.JUSTIFIED,
            pageBreakBefore: options.pageBreakBefore
        });
    }
    
createTableCell(content, options = {}) {
    let children;
    
    if (Array.isArray(content)) {
        children = content;
    } else {
        // Déterminer l'alignement automatiquement
        let alignment = options.alignment;
        if (!alignment) {
            // Si c'est une cellule d'en-tête (avec fond coloré primary ou primaryDark), centrer automatiquement
            if (options.shading === this.colors.primary || options.shading === this.colors.primaryDark) {
                alignment = docx.AlignmentType.CENTER;
            } else {
                alignment = docx.AlignmentType.LEFT;
            }
        }
        
        children = [
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: content || '',
                        bold: options.bold,
                        size: 20,
                        font: "Calibri",
                        color: options.textColor || (options.shading === this.colors.primary || options.shading === this.colors.primaryDark ? this.colors.white : this.colors.text)
                    })
                ],
                alignment: alignment,
                spacing: { after: 0 }
            })
        ];
    }
    
    const cellOptions = {
        children: children,
        margins: {
            top: 120,
            bottom: 120,
            left: 120,
            right: 120
        },
        // Utiliser l'alignement vertical des options s'il est fourni, sinon utiliser l'auto-détection
        verticalAlign: options.verticalAlign || (
            (options.shading === this.colors.primary || options.shading === this.colors.primaryDark) 
                ? docx.VerticalAlign.CENTER 
                : docx.VerticalAlign.TOP
        )
    };
    
    if (options.shading) {
        cellOptions.shading = {
            fill: options.shading,
            type: docx.ShadingType.CLEAR
        };
    }
    
    if (options.columnSpan) {
        cellOptions.columnSpan = options.columnSpan;
    }
    
    return new docx.TableCell(cellOptions);
}    
    createEnhancedTable(headers, data) {
        const rows = [];
        
        // Header row with primary color background
        rows.push(new docx.TableRow({
            children: headers.map(h => this.createTableCell(h, { 
                shading: this.colors.primary, 
                bold: true,
                textColor: this.colors.white
            })),
            tableHeader: true,
            height: {
                value: 500,
                rule: docx.HeightRule.MINIMUM
            }
        }));
        
        // Data rows with alternating background
        data.forEach((rowData, index) => {
            const shading = index % 2 === 0 ? this.colors.white : this.colors.background;
            rows.push(new docx.TableRow({
                children: rowData.map(cell => this.createTableCell(cell, { 
                    shading,
                    textColor: this.colors.text
                }))
            }));
        });
        
        return new docx.Table({
            width: {
                size: 100,
                type: docx.WidthType.PERCENTAGE
            },
            borders: {
                top: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                left: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                right: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
                insideVertical: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight }
            },
            rows: rows
        });
    }
    
    createFinancialTable(headers, data, totals, columnCount = 4) {
        const rows = [];
        
        // Header row with dark primary background
        rows.push(new docx.TableRow({
            children: headers.map(h => this.createTableCell(h, { 
                shading: this.colors.primaryDark, 
                bold: true,
                textColor: this.colors.white
            })),
            tableHeader: true,
            height: {
                value: 500,
                rule: docx.HeightRule.MINIMUM
            }
        }));
        
        // Data rows with alternating background
        data.forEach((rowData, index) => {
            const shading = index % 2 === 0 ? this.colors.white : this.colors.background;
            rows.push(new docx.TableRow({
                children: rowData.map(cell => this.createTableCell(cell, { 
                    shading,
                    textColor: this.colors.text
                }))
            }));
        });
        
        // Total rows with gradient effect
        totals.forEach((total, index) => {
            const cells = [];
            const isLast = index === totals.length - 1;
            
            // First cell spans multiple columns
            cells.push(this.createTableCell(total.label, {
                columnSpan: columnCount - 1,
                shading: isLast ? this.colors.primary : this.colors.primaryLight,
                bold: true,
                textColor: isLast ? this.colors.white : this.colors.primaryDark
            }));
            
            // Last cell contains the value
            cells.push(this.createTableCell(total.value, {
                shading: isLast ? this.colors.primary : this.colors.primaryLight,
                bold: true,
                alignment: docx.AlignmentType.RIGHT,
                textColor: isLast ? this.colors.white : this.colors.primaryDark
            }));
            
            rows.push(new docx.TableRow({ children: cells }));
        });
        
        return new docx.Table({
            width: {
                size: 100,
                type: docx.WidthType.PERCENTAGE
            },
            borders: {
                top: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                left: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                right: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                insideHorizontal: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight },
                insideVertical: { style: docx.BorderStyle.SINGLE, size: 1, color: this.colors.primaryLight }
            },
            rows: rows
        });
    }
    
    // Nouvelle méthode pour créer une image à partir d'un base64
    async createImageFromBase64(base64Data, altText = "Image") {
        try {
            // Vérifier le format base64
            const base64Match = base64Data.match(/^data:image\/([a-zA-Z]*);base64,(.*)$/);
            if (!base64Match) {
                console.error('Invalid base64 image format');
                return this.createParagraph("[Image non disponible - format invalide]", { italics: true });
            }
            
            const mimeType = `image/${base64Match[1]}`;
            const imageData = base64Match[2];
            
            // Convertir base64 en ArrayBuffer
            const binaryString = atob(imageData);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            
            // Créer une image temporaire pour obtenir les dimensions réelles
            const img = new Image();
            const dimensionsPromise = new Promise((resolve) => {
                img.onload = () => {
                    resolve({ width: img.width, height: img.height });
                };
                img.onerror = () => {
                    resolve({ width: 600, height: 400 }); // Dimensions par défaut en cas d'erreur
                };
            });
            
            // Charger l'image
            img.src = base64Data;
            const originalDimensions = await dimensionsPromise;
            
            // Calculer les dimensions finales en conservant les proportions
            const maxWidth = 600; // Largeur maximale dans le document (environ 21cm - marges)
            const maxHeight = 800; // Hauteur maximale raisonnable
            
            let finalWidth = originalDimensions.width;
            let finalHeight = originalDimensions.height;
            
            // Si l'image est trop large
            if (finalWidth > maxWidth) {
                const ratio = maxWidth / finalWidth;
                finalWidth = maxWidth;
                finalHeight = Math.round(originalDimensions.height * ratio);
            }
            
            // Si l'image est encore trop haute après redimensionnement
            if (finalHeight > maxHeight) {
                const ratio = maxHeight / finalHeight;
                finalHeight = maxHeight;
                finalWidth = Math.round(finalWidth * ratio);
            }
            
            // Pour les diagrammes de Gantt, forcer une largeur importante
            if (altText.toLowerCase().includes('gantt')) {
                finalWidth = Math.min(680, originalDimensions.width); // Un peu plus large pour les Gantt
                finalHeight = Math.round((originalDimensions.height / originalDimensions.width) * finalWidth);
            }
            
            // Créer l'image avec docx
            const image = new docx.ImageRun({
                data: bytes.buffer,
                transformation: {
                    width: finalWidth,
                    height: finalHeight
                },
                type: mimeType
            });
            
            // Retourner un paragraphe contenant l'image, centré
            return new docx.Paragraph({
                children: [image],
                alignment: docx.AlignmentType.CENTER,
                spacing: { before: 200, after: 200 }
            });
            
        } catch (error) {
            console.error('Error creating image from base64:', error);
            // En cas d'erreur, créer un placeholder élégant
            return new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `📷 ${altText}`,
                        italics: true,
                        size: 24,
                        color: this.colors.primary
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { before: 200, after: 200 },
                border: {
                    top: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                    bottom: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                    left: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight },
                    right: { style: docx.BorderStyle.SINGLE, size: 2, color: this.colors.primaryLight }
                }
            });
        }
    }
}
