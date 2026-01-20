/**
 * Excel Parser Module
 * Verantwortlich für das Parsen von Excel-Dateien und die Transformation von Hotel-Export-Daten
 */

class ExcelParser {
    constructor() {
        // Mapping: Excel-Spaltenname → CSV-Feldname
        // Unterstützt mehrere mögliche Spaltennamen (erster gefundener wird verwendet)
        this.columnMapping = {
            'BookingNumber': {
                possibleNames: ['Reservation Number', 'Reservation Num'],
                type: 'string'
            },
            'OTANumber': {
                possibleNames: ['Voucher'],
                type: 'string'
            },
            'Name': {
                possibleNames: ['Main Guest'],
                type: 'string'
            },
            'NumberOfAdults': {
                possibleNames: ['AD', 'Adults'],
                type: 'number'
            },
            'NumberOfChildren': {
                possibleNames: ['CH', 'Children'],
                type: 'number'
            },
            'DateFrom': {
                possibleNames: ['Arrival', 'Arrival Date'],
                type: 'string'
            },
            'DateTo': {
                possibleNames: ['Departure', 'Departure Date'],
                type: 'string'
            }
        };

        // Pflichtspalte (Zeilen ohne diesen Wert werden verworfen)
        this.requiredColumn = 'BookingNumber';
    }

    /**
     * Parst eine Excel-Datei und transformiert die Daten gemäß Mapping
     * @param {File} file - Die hochgeladene Excel-Datei
     * @returns {Promise<{rows: Array<Object>, fileName: string}>}
     * @throws {Error} Wenn die Datei nicht gelesen werden kann oder erforderliche Spalten fehlen
     */
    async parseFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Erste Tabelle verwenden
                    const firstSheetName = workbook.SheetNames[0];
                    if (!firstSheetName) {
                        throw new Error('Die Excel-Datei enthält keine Tabellen.');
                    }

                    const worksheet = workbook.Sheets[firstSheetName];
                    // Worksheet-Referenz für direkten Zugriff auf Zellen behalten
                    this.worksheet = worksheet;
                    
                    // Zuerst die Kopfzeile finden
                    // raw: true verwenden, damit wir die ursprünglichen Werte haben für Datumskonvertierung
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        defval: '',
                        raw: true
                    });

                    if (jsonData.length === 0) {
                        throw new Error('Die Excel-Datei ist leer.');
                    }

                    // Kopfzeile finden
                    const headerRow = jsonData[0];
                    if (!Array.isArray(headerRow) || headerRow.length === 0) {
                        throw new Error('Die Excel-Datei enthält keine gültige Kopfzeile.');
                    }

                    // Spaltenindizes finden
                    const columnIndices = this.findColumnIndices(headerRow);
                    
                    // Prüfen, ob die Pflichtspalte vorhanden ist
                    if (columnIndices[this.requiredColumn] === -1) {
                        const availableColumns = headerRow.filter(h => h).join(', ');
                        const possibleNames = this.columnMapping[this.requiredColumn].possibleNames.join('" oder "');
                        throw new Error(
                            `Die erforderliche Spalte "${possibleNames}" wurde nicht gefunden. ` +
                            `Verfügbare Spalten: ${availableColumns}`
                        );
                    }

                    // Datenzeilen verarbeiten (ab Zeile 2, da Zeile 1 die Kopfzeile ist)
                    const processedRows = [];
                    let rowNumber = 1; // Für Id-Zählung

                    for (let i = 1; i < jsonData.length; i++) {
                        const row = jsonData[i];
                        
                        if (!Array.isArray(row)) {
                            continue;
                        }

                        // Prüfen, ob BookingNumber (Reservation Number) vorhanden ist
                        const bookingNumberIndex = columnIndices[this.requiredColumn];
                        const bookingNumber = this.getCellValue(row, bookingNumberIndex);
                        
                        // Zeilen ohne BookingNumber verwerfen
                        if (!bookingNumber || bookingNumber.trim() === '') {
                            continue;
                        }

                        // Zeile transformieren
                        const transformedRow = this.transformRow(row, columnIndices, rowNumber, i);
                        processedRows.push(transformedRow);
                        rowNumber++;
                    }

                    if (processedRows.length === 0) {
                        const possibleNames = this.columnMapping[this.requiredColumn].possibleNames.join('" oder "');
                        throw new Error(
                            'Keine gültigen Zeilen gefunden. Stellen Sie sicher, dass die Spalte "' +
                            possibleNames + '" Werte enthält.'
                        );
                    }

                    resolve({
                        rows: processedRows,
                        fileName: file.name
                    });

                } catch (error) {
                    reject(error);
                }
            };

            reader.onerror = () => {
                reject(new Error('Fehler beim Lesen der Datei. Bitte versuchen Sie es erneut.'));
            };

            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Findet die Indizes aller benötigten Spalten in der Kopfzeile
     * @param {Array} headerRow - Die Kopfzeile als Array
     * @returns {Object} - Objekt mit CSV-Feldnamen als Keys und Indizes als Values (-1 wenn nicht gefunden)
     */
    findColumnIndices(headerRow) {
        const indices = {};
        
        for (const [csvFieldName, config] of Object.entries(this.columnMapping)) {
            let foundIndex = -1;
            
            // Versuche jeden möglichen Spaltennamen zu finden
            for (const possibleName of config.possibleNames) {
                const index = this.findColumnIndex(headerRow, possibleName);
                if (index !== -1) {
                    foundIndex = index;
                    break; // Ersten gefundenen verwenden
                }
            }
            
            indices[csvFieldName] = foundIndex;
        }

        return indices;
    }

    /**
     * Findet den Index einer Spalte in der Kopfzeile (case-insensitive)
     * @param {Array} headerRow - Die Kopfzeile als Array
     * @param {string} columnName - Der Name der gesuchten Spalte
     * @returns {number} - Der Index der Spalte oder -1 wenn nicht gefunden
     */
    findColumnIndex(headerRow, columnName) {
        const normalizedColumnName = columnName.trim().toLowerCase();
        
        for (let i = 0; i < headerRow.length; i++) {
            const headerValue = String(headerRow[i] || '').trim().toLowerCase();
            if (headerValue === normalizedColumnName) {
                return i;
            }
        }

        return -1;
    }

    /**
     * Extrahiert einen Zellwert aus einer Zeile
     * @param {Array} row - Die Datenzeile
     * @param {number} index - Der Spaltenindex
     * @param {number} rowIndex - Der Zeilenindex (0-basiert, inkl. Header)
     * @param {string} columnName - Der Name der Spalte (für Datumsbehandlung)
     * @returns {string} - Der Zellwert als String (leer wenn nicht vorhanden)
     */
    getCellValue(row, index, rowIndex = -1, columnName = '') {
        if (index === -1 || !row || row[index] === undefined || row[index] === null) {
            return '';
        }

        const value = row[index];
        const stringValue = String(value).trim();

        // Wenn es ein Datumsfeld ist, versuche das ursprüngliche Format aus dem Worksheet zu lesen
        if ((columnName === 'DateFrom' || columnName === 'DateTo') && 
            rowIndex > 0 && 
            this.worksheet) {
            
            // Excel-Spaltenbuchstaben berechnen (A, B, C, ...)
            const colLetter = this.numberToColumnLetter(index);
            const cellAddress = colLetter + (rowIndex + 1); // Excel ist 1-basiert
            
            const cell = this.worksheet[cellAddress];
            if (cell) {
                // Wenn die Zelle ein formatiertes Datum hat (w), verwende das Originalformat
                if (cell.w) {
                    return cell.w;
                }
                
                // Prüfe, ob es eine Zahl ist, die wie ein Excel-Datum aussieht
                const numValue = parseFloat(stringValue);
                if (!isNaN(numValue) && numValue > 1 && numValue < 1000000) {
                    // Wenn die Zelle ein Datum ist (t === 'd' oder Zahl im Excel-Datumsbereich)
                    if (cell.t === 'd' || (cell.v && typeof cell.v === 'number')) {
                        // Excel-Datum konvertieren (Tage seit 1900-01-01)
                        const excelDate = cell.v || numValue;
                        const date = this.excelDateToJSDate(excelDate);
                        // Format als YYYY-MM-DD
                        return this.formatDate(date);
                    }
                }
            } else {
                // Wenn Zelle nicht gefunden wurde, aber Wert eine Zahl ist, versuche Konvertierung
                const numValue = parseFloat(stringValue);
                if (!isNaN(numValue) && numValue > 1 && numValue < 1000000) {
                    // Möglicherweise ein Excel-Datum
                    const date = this.excelDateToJSDate(numValue);
                    return this.formatDate(date);
                }
            }
        }

        return stringValue;
    }

    /**
     * Konvertiert eine Spaltennummer (0-basiert) zu einem Excel-Spaltenbuchstaben (A, B, C, ...)
     * @param {number} num - Spaltennummer (0-basiert)
     * @returns {string} - Excel-Spaltenbuchstabe
     */
    numberToColumnLetter(num) {
        let result = '';
        while (num >= 0) {
            result = String.fromCharCode(65 + (num % 26)) + result;
            num = Math.floor(num / 26) - 1;
        }
        return result;
    }

    /**
     * Konvertiert ein Excel-Datum (Seriennummer) zu einem JavaScript-Datum
     * @param {number} excelDate - Excel-Seriennummer (Tage seit 1900-01-01)
     * @returns {Date} - JavaScript-Datum
     */
    excelDateToJSDate(excelDate) {
        // Excel-Datum startet am 1. Januar 1900, aber Excel hat einen Bug:
        // Es behandelt 1900 als Schaltjahr, obwohl es keines ist
        // Deshalb müssen wir 1 Tag abziehen, außer für Daten vor März 1900
        const excelEpoch = new Date(1899, 11, 30); // 30. Dezember 1899
        const jsDate = new Date(excelEpoch.getTime() + excelDate * 24 * 60 * 60 * 1000);
        return jsDate;
    }

    /**
     * Formatiert ein Datum als String (DD.MM.YYYY)
     * @param {Date} date - JavaScript-Datum
     * @returns {string} - Formatierter Datumsstring
     */
    formatDate(date) {
        if (!date || isNaN(date.getTime())) {
            return '';
        }
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${day}.${month}.${year}`;
    }

    /**
     * Transformiert eine Datenzeile gemäß Mapping
     * @param {Array} row - Die Datenzeile
     * @param {Object} columnIndices - Die Spaltenindizes (mit CSV-Feldnamen als Keys)
     * @param {number} id - Die fortlaufende ID für diese Zeile
     * @param {number} rowIndex - Der Zeilenindex im Worksheet (0-basiert, inkl. Header)
     * @returns {Object} - Transformierte Zeile mit allen CSV-Feldern
     */
    transformRow(row, columnIndices, id, rowIndex) {
        const transformed = {
            Id: '', // Wird später auf BookingNumber gesetzt
            BookingNumber: '',
            OTANumber: '',
            Name: '',
            NumberOfAdults: '',
            NumberOfTeens: 0,
            NumberOfChildren: '',
            NumberOfBabys: 0,
            DateFrom: '',
            DateTo: ''
        };

        // Mapping durchführen
        for (const [csvFieldName, config] of Object.entries(this.columnMapping)) {
            const columnIndex = columnIndices[csvFieldName];
            
            if (columnIndex === -1) {
                // Spalte nicht gefunden, Standardwert beibehalten
                continue;
            }
            
            // Ursprünglicher Wert aus dem Array
            const rawValue = row[columnIndex];
            
            // Spezielle Behandlung für Datumsfelder
            if (csvFieldName === 'DateFrom' || csvFieldName === 'DateTo') {
                let dateValue = '';
                
                // Versuche zuerst, die Zelle direkt aus dem Worksheet zu lesen
                // Dies funktioniert unabhängig von der Spaltenposition, da columnIndex
                // basierend auf dem Spaltennamen gefunden wurde
                if (rowIndex > 0 && this.worksheet && columnIndex !== -1) {
                    const colLetter = this.numberToColumnLetter(columnIndex);
                    const cellAddress = colLetter + (rowIndex + 1);
                    const cell = this.worksheet[cellAddress];
                    
                    if (cell) {
                        if (cell.w) {
                            // Verwende das formatierte Datum aus Excel, aber prüfe ob es konvertiert werden muss
                            // Wenn cell.w bereits ein Datum im richtigen Format ist, verwende es
                            // Sonst konvertiere cell.v falls vorhanden
                            const wValue = String(cell.w).trim();
                            // Prüfe ob es bereits im DD.MM.YYYY Format ist
                            if (/^\d{2}\.\d{2}\.\d{4}$/.test(wValue)) {
                                dateValue = wValue;
                            } else if (cell.v !== undefined && cell.v !== null && typeof cell.v === 'number' && cell.v > 1 && cell.v < 1000000) {
                                // Konvertiere die Seriennummer zu DD.MM.YYYY
                                const date = this.excelDateToJSDate(cell.v);
                                dateValue = this.formatDate(date);
                            } else {
                                dateValue = wValue;
                            }
                        } else if (cell.v !== undefined && cell.v !== null) {
                            // Wenn cell.v eine Zahl im Excel-Datumsbereich ist, konvertiere sie
                            if (typeof cell.v === 'number' && cell.v > 1 && cell.v < 1000000) {
                                const date = this.excelDateToJSDate(cell.v);
                                dateValue = this.formatDate(date);
                            } else {
                                dateValue = String(cell.v);
                            }
                        }
                    }
                }
                
                // Fallback: Wenn noch kein Wert gesetzt wurde, prüfe den rawValue aus dem Array
                // Dieser Fallback funktioniert immer, unabhängig von der Spaltenposition,
                // da rawValue direkt aus row[columnIndex] kommt, wobei columnIndex über
                // den Spaltennamen gefunden wurde
                if (!dateValue && rawValue !== undefined && rawValue !== null && rawValue !== '') {
                    const numValue = typeof rawValue === 'number' ? rawValue : parseFloat(rawValue);
                    if (!isNaN(numValue) && numValue > 1 && numValue < 1000000) {
                        // Es ist eine Excel-Seriennummer, konvertiere sie zu einem Datum
                        const date = this.excelDateToJSDate(numValue);
                        dateValue = this.formatDate(date);
                    } else {
                        // Wenn es keine Zahl ist, verwende den Wert direkt (könnte bereits formatiert sein)
                        dateValue = String(rawValue).trim();
                    }
                }
                
                transformed[csvFieldName] = dateValue;
                continue;
            }
            
            // Standard-Verarbeitung für alle anderen Felder
            let cellValue = this.getCellValue(row, columnIndex, rowIndex, csvFieldName);

            // Spezielle Behandlung für BookingNumber: führende Nullen entfernen (insbesondere die erste 0)
            if (csvFieldName === 'BookingNumber') {
                if (typeof cellValue === 'string') {
                    // Nur eine führende 0 entfernen
                    if (cellValue.startsWith('0')) {
                        cellValue = cellValue.substring(1);
                    }
                }
            }
            
            // Spezielle Behandlung für numerische Felder
            if (config.type === 'number') {
                // Versuche als Zahl zu parsen, sonst leer lassen
                const numValue = parseInt(cellValue, 10);
                transformed[csvFieldName] = isNaN(numValue) ? '' : numValue;
            } else {
                transformed[csvFieldName] = cellValue;
            }
        }

        // ID gleich BookingNumber setzen
        transformed.Id = transformed.BookingNumber || '';

        return transformed;
    }

    /**
     * Validiert, ob eine Datei ein gültiges Excel-Format hat
     * @param {File} file - Die zu validierende Datei
     * @returns {boolean}
     */
    isValidFileType(file) {
        const validTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel' // .xls
        ];
        const validExtensions = ['.xlsx', '.xls'];
        
        const hasValidType = validTypes.includes(file.type);
        const hasValidExtension = validExtensions.some(ext => 
            file.name.toLowerCase().endsWith(ext)
        );

        return hasValidType || hasValidExtension;
    }
}
