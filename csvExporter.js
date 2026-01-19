/**
 * CSV Export Module
 * Verantwortlich für die Generierung und den Download von CSV-Dateien
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
     * Erstellt eine CSV-Datei aus den transformierten Zeilen
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
     * Startet den Download einer CSV-Datei
     * @param {string} csvContent - Der CSV-String
     * @param {string} fileName - Der gewünschte Dateiname (ohne Extension)
     */
    downloadCSV(csvContent, fileName = 'export') {
        // BOM für UTF-8 hinzufügen (für korrekte Anzeige in Excel)
        const BOM = '\uFEFF';
        const csvWithBOM = BOM + csvContent;

        // Blob erstellen
        const blob = new Blob([csvWithBOM], { type: 'text/csv;charset=utf-8;' });
        
        // Download-Link erstellen
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        
        // Dateiname: Anreise_heutigesDatum (z.B. Anreise_20260118)
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        const dateString = `${year}${month}${day}`;
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
