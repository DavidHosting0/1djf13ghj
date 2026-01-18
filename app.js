/**
 * Hauptanwendungslogik
 * Koordiniert UI, Dateiparsing und CSV-Export
 */

class App {
    constructor() {
        this.excelParser = new ExcelParser();
        this.csvExporter = new CSVExporter();
        this.currentData = null;

        this.initializeElements();
        this.attachEventListeners();
    }

    /**
     * Initialisiert alle DOM-Elemente
     */
    initializeElements() {
        // Upload Section
        this.uploadSection = document.getElementById('uploadSection');
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.selectFileBtn = document.getElementById('selectFileBtn');

        // Processing Section
        this.processingSection = document.getElementById('processingSection');
        this.statusTitle = document.getElementById('statusTitle');
        this.statusMessage = document.getElementById('statusMessage');
        this.progressFill = document.getElementById('progressFill');

        // Result Section
        this.resultSection = document.getElementById('resultSection');
        this.entryCount = document.getElementById('entryCount');
        this.fileName = document.getElementById('fileName');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.resetBtn = document.getElementById('resetBtn');

        // Error Section
        this.errorSection = document.getElementById('errorSection');
        this.errorMessage = document.getElementById('errorMessage');
        this.errorResetBtn = document.getElementById('errorResetBtn');
    }

    /**
     * Fügt Event-Listener hinzu
     */
    attachEventListeners() {
        // Datei-Auswahl Button
        this.selectFileBtn.addEventListener('click', () => {
            this.fileInput.click();
        });

        // File Input Change
        this.fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                this.handleFile(file);
            }
        });

        // Drag & Drop
        this.uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.uploadArea.classList.add('dragover');
        });

        this.uploadArea.addEventListener('dragleave', () => {
            this.uploadArea.classList.remove('dragover');
        });

        this.uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            this.uploadArea.classList.remove('dragover');
            
            const file = e.dataTransfer.files[0];
            if (file) {
                this.handleFile(file);
            }
        });

        // Download Button
        this.downloadBtn.addEventListener('click', () => {
            this.downloadCSV();
        });

        // Reset Buttons
        this.resetBtn.addEventListener('click', () => {
            this.reset();
        });

        this.errorResetBtn.addEventListener('click', () => {
            this.reset();
        });
    }

    /**
     * Verarbeitet eine hochgeladene Datei
     * @param {File} file - Die hochgeladene Datei
     */
    async handleFile(file) {
        // Validierung
        if (!this.excelParser.isValidFileType(file)) {
            this.showError('Bitte wählen Sie eine gültige Excel-Datei (.xlsx oder .xls) aus.');
            return;
        }

        // UI auf Processing umstellen
        this.showProcessing();

        try {
            // Datei parsen
            const result = await this.excelParser.parseFile(file);
            
            // Daten speichern
            this.currentData = result;

            // UI auf Result umstellen
            this.showResult(result);

        } catch (error) {
            console.error('Fehler beim Verarbeiten der Datei:', error);
            this.showError(error.message || 'Ein unbekannter Fehler ist aufgetreten.');
        }
    }

    /**
     * Zeigt den Processing-Zustand an
     */
    showProcessing() {
        this.hideAllSections();
        this.processingSection.classList.remove('hidden');
        this.statusTitle.textContent = 'Datei wird verarbeitet...';
        this.statusMessage.textContent = 'Bitte warten Sie einen Moment.';
        this.progressFill.style.width = '0%';
        
        // Progress-Animation starten
        setTimeout(() => {
            this.progressFill.style.width = '100%';
        }, 100);
    }

    /**
     * Zeigt den Result-Zustand an
     * @param {Object} result - Das Ergebnis mit rows und fileName
     */
    showResult(result) {
        this.hideAllSections();
        this.resultSection.classList.remove('hidden');
        this.entryCount.textContent = result.rows.length;
        this.fileName.textContent = result.fileName;
        
        // Vorschau-Tabelle füllen
        this.populatePreviewTable(result.rows);
    }

    /**
     * Füllt die Vorschau-Tabelle mit Daten
     * @param {Array<Object>} rows - Die zu zeigenden Zeilen
     */
    populatePreviewTable(rows) {
        const tbody = document.getElementById('previewTableBody');
        tbody.innerHTML = '';
        
        // Zeige maximal 50 Zeilen in der Vorschau
        const maxRows = Math.min(rows.length, 50);
        
        for (let i = 0; i < maxRows; i++) {
            const row = rows[i];
            const tr = document.createElement('tr');
            
            const columns = [
                'Id', 'BookingNumber', 'OTANumber', 'Name', 
                'NumberOfAdults', 'NumberOfTeens', 'NumberOfChildren', 'NumberOfBabys',
                'DateFrom', 'DateTo'
            ];
            
            columns.forEach(col => {
                const td = document.createElement('td');
                const value = row[col] !== undefined && row[col] !== null ? row[col] : '';
                td.textContent = value;
                tr.appendChild(td);
            });
            
            tbody.appendChild(tr);
        }
        
        // Wenn mehr Zeilen vorhanden sind, zeige Hinweis
        if (rows.length > maxRows) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 10;
            td.className = 'preview-more';
            td.textContent = `... und ${rows.length - maxRows} weitere Zeilen`;
            tr.appendChild(td);
            tbody.appendChild(tr);
        }
    }

    /**
     * Zeigt einen Fehler an
     * @param {string} message - Die Fehlermeldung
     */
    showError(message) {
        this.hideAllSections();
        this.errorSection.classList.remove('hidden');
        this.errorMessage.textContent = message;
    }

    /**
     * Versteckt alle Sektionen
     */
    hideAllSections() {
        this.uploadSection.classList.add('hidden');
        this.processingSection.classList.add('hidden');
        this.resultSection.classList.add('hidden');
        this.errorSection.classList.add('hidden');
    }

    /**
     * Startet den CSV-Download
     */
    downloadCSV() {
        if (!this.currentData || !this.currentData.rows || this.currentData.rows.length === 0) {
            this.showError('Keine Daten zum Exportieren verfügbar.');
            return;
        }

        try {
            const csvContent = this.csvExporter.createCSV(this.currentData.rows);
            
            const baseFileName = this.currentData.fileName
                .replace(/\.(xlsx|xls)$/i, '')
                .replace(/[^a-z0-9]/gi, '_');
            
            this.csvExporter.downloadCSV(csvContent, baseFileName);
        } catch (error) {
            console.error('Fehler beim CSV-Export:', error);
            this.showError('Fehler beim Erstellen der CSV-Datei. Bitte versuchen Sie es erneut.');
        }
    }

    /**
     * Setzt die Anwendung zurück
     */
    reset() {
        this.currentData = null;
        this.fileInput.value = '';
        this.hideAllSections();
        this.uploadSection.classList.remove('hidden');
    }
}

// App initialisieren, sobald das DOM geladen ist
document.addEventListener('DOMContentLoaded', () => {
    new App();
});


