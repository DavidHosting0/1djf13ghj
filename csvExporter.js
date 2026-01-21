/**
 * Excel Export Module
 * Verantwortlich für die Generierung und den Download von Excel-Dateien mit Formatierung
 */

class CSVExporter {
    /**
     * CSV-Spaltenreihenfolge (exakt wie spezifiziert)
     */
    constructor() {
        this.csvColumns = [
            'Id',
            'BookingNumber',
            'OTANumber',
            'Name',
            'NumberOfAdults',
            'NumberOfTeens',
            'NumberOfChildren',
            'NumberOfBabys',
            'DateFrom',
            'DateTo'
        ];
        // Semikolon als Trennzeichen für deutsche Excel-Versionen
        this.delimiter = ';';
    }

    /**
     * Erstellt eine Excel-Datei aus den transformierten Zeilen
     * @param {Array<Object>} rows - Array von transformierten Zeilenobjekten
     * @returns {ArrayBuffer} - Excel-Datei als ArrayBuffer
     */
    createExcel(rows) {
        if (!Array.isArray(rows) || rows.length === 0) {
            throw new Error('Keine Daten zum Exportieren verfügbar.');
        }

        // Neues Workbook erstellen
        const wb = XLSX.utils.book_new();
        
        // Daten für das Worksheet vorbereiten
        const worksheetData = [
            this.csvColumns, // Kopfzeile
            ...rows.map(row => {
                return this.csvColumns.map(column => {
                    return row[column] !== undefined && row[column] !== null ? row[column] : '';
                });
            })
        ];
        
        // Worksheet erstellen
        const ws = XLSX.utils.aoa_to_sheet(worksheetData);
        
        // Spaltenbreiten setzen
        const colWidths = [
            { wch: 15 }, // Id
            { wch: 15 }, // BookingNumber
            { wch: 15 }, // OTANumber
            { wch: 25 }, // Name
            { wch: 12 }, // NumberOfAdults
            { wch: 12 }, // NumberOfTeens
            { wch: 12 }, // NumberOfChildren
            { wch: 12 }, // NumberOfBabys
            { wch: 12 }, // DateFrom
            { wch: 12 }  // DateTo
        ];
        ws['!cols'] = colWidths;
        
        // Zellformatierung: Id, BookingNumber und OTANumber zentrieren
        const centerAlignment = { alignment: { horizontal: 'center', vertical: 'center' } };
        
        // Bereich für alle Zellen bestimmen
        const range = XLSX.utils.decode_range(ws['!ref']);
        
        // Erste drei Spalten (Id, BookingNumber, OTANumber) zentrieren
        for (let row = 0; row <= range.e.r; row++) {
            for (let col = 0; col <= 2; col++) { // Spalten A, B, C (0, 1, 2)
                const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
                if (!ws[cellAddress]) {
                    ws[cellAddress] = { t: 's', v: '' };
                }
                ws[cellAddress].s = {
                    alignment: { horizontal: 'center', vertical: 'center' }
                };
            }
        }
        
        // Worksheet zum Workbook hinzufügen
        XLSX.utils.book_append_sheet(wb, ws, 'Daten');
        
        // Excel-Datei als ArrayBuffer erstellen
        return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    }

    /**
     * Konvertiert eine Excel-Datei zu CSV
     * @param {ArrayBuffer} excelData - Die Excel-Datei als ArrayBuffer
     * @returns {string} - CSV-String
     */
    excelToCSV(excelData) {
        // Excel-Datei lesen
        const wb = XLSX.read(excelData, { type: 'array' });
        
        // Erste Tabelle verwenden
        const firstSheetName = wb.SheetNames[0];
        const ws = wb.Sheets[firstSheetName];
        
        // Als CSV konvertieren (Formatierung geht verloren, aber Daten bleiben)
        return XLSX.utils.sheet_to_csv(ws, { 
            FS: this.delimiter,
            RS: '\r\n'
        });
    }

    /**
     * Startet den Download einer CSV-Datei
     * Erstellt zuerst eine Excel-Datei mit Formatierung, konvertiert sie dann zu CSV
     * @param {ArrayBuffer} excelData - Die Excel-Datei als ArrayBuffer
     */
    downloadCSV(excelData) {
        // Excel-Datei zu CSV konvertieren
        const csvContent = this.excelToCSV(excelData);
        
        // CSV-Datei herunterladen
        this.downloadCSVDirect(csvContent);
    }

    /**
     * Erstellt eine CSV-Datei direkt (ohne Formatierung)
     * @param {Array<Object>} rows - Array von transformierten Zeilenobjekten
     * @returns {string} - CSV-String
     */
    createCSV(rows) {
        if (!Array.isArray(rows) || rows.length === 0) {
            throw new Error('Keine Daten zum Exportieren verfügbar.');
        }

        // CSV-Escape-Funktion für Werte, die Trennzeichen, Anführungszeichen oder Zeilenumbrüche enthalten
        const escapeCSVValue = (value) => {
            const stringValue = String(value);
            const delimiter = this.delimiter;
            // Wenn der Wert Trennzeichen, Anführungszeichen oder Zeilenumbrüche enthält, in Anführungszeichen setzen
            if (stringValue.includes(delimiter) || stringValue.includes('"') || stringValue.includes('\n') || stringValue.includes('\r')) {
                // Anführungszeichen verdoppeln (CSV-Escape)
                return `"${stringValue.replace(/"/g, '""')}"`;
            }
            return stringValue;
        };

        // Kopfzeile erstellen
        const header = this.csvColumns.map(col => escapeCSVValue(col)).join(this.delimiter);

        // Datenzeilen erstellen
        const dataRows = rows.map(row => {
            return this.csvColumns.map(column => {
                const value = row[column] !== undefined && row[column] !== null ? row[column] : '';
                return escapeCSVValue(value);
            }).join(this.delimiter);
        });

        // CSV zusammenfügen mit Windows-Zeilenumbrüchen (CRLF) für bessere Excel-Kompatibilität
        return [header, ...dataRows].join('\r\n');
    }

    /**
     * Startet den Download einer CSV-Datei direkt
     * @param {string} csvContent - Der CSV-String
     */
    downloadCSVDirect(csvContent) {
        // BOM für UTF-8 hinzufügen (für korrekte Anzeige in Excel)
        const BOM = '\uFEFF';
        const csvWithBOM = BOM + csvContent;

        // Blob erstellen
        const blob = new Blob([csvWithBOM], { type: 'text/csv;charset=utf-8;' });
        
        // Download-Link erstellen
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        
        // Dateiname: Anreise_DD.MM.YYYY (z.B. Anreise_19.01.2026)
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        const dateString = `${day}.${month}.${year}`;
        link.download = `Anreise_${dateString}.csv`;
        
        // Link temporär zum DOM hinzufügen, klicken und entfernen
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // URL nach kurzer Zeit freigeben (für bessere Performance)
        setTimeout(() => {
            URL.revokeObjectURL(url);
        }, 100);
    }

    /**
     * Generiert einen Zeitstempel für den Dateinamen
     * @returns {string} - Format: YYYYMMDD_HHMMSS
     */
    getTimestamp() {
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');
        
        return `${year}${month}${day}_${hours}${minutes}${seconds}`;
    }
}
