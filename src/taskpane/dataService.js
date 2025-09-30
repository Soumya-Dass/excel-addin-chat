/* global Excel */

// DataService - Advanced Excel data reading with full Office.js capabilities
class DataService {
    constructor() {
        this.CONFIG = {
            // No artificial limits - use actual data boundaries
            MAX_ANALYSIS_CELLS: 100000, // Safety limit for performance (100k cells)
            CHUNK_SIZE: 10000, // For processing large datasets in chunks
            SAMPLE_RATIO: 0.1, // 10% sampling for very large datasets
        };
        this.currentWorksheetData = null;
        this.workbookStructure = null;
    }

    // Advanced Excel data reading with full Office.js capabilities
    async readCurrentWorksheetDataEnhanced(shouldUseSelection = true, includeAllSheets = false) {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                try {
                    const worksheet = context.workbook.worksheets.getActiveWorksheet();
                    worksheet.load('name');
                    
                    let range;
                    let isSelection = false;
                    let isSampled = false;
                    let dataMetadata = {};
                    
                    // First, gather advanced worksheet information
                    const tables = worksheet.tables;
                    const pivotTables = worksheet.pivotTables;
                    const charts = worksheet.charts;
                    tables.load(['name', 'id']);
                    pivotTables.load(['name', 'id']);
                    charts.load(['name', 'id']);
                    
                    if (shouldUseSelection) {
                        // Try to use selected range
                        const selectedRange = context.workbook.getSelectedRange();
                        selectedRange.load(['values', 'formulas', 'rowCount', 'columnCount', 'address', 'numberFormat']);
                        await context.sync();
                        
                        // Check if selection is valid and reasonable size
                        if (selectedRange.rowCount > 1 || selectedRange.columnCount > 1 || 
                            (selectedRange.values[0][0] !== null && selectedRange.values[0][0] !== "")) {
                            
                            const totalCells = selectedRange.rowCount * selectedRange.columnCount;
                            if (totalCells <= this.CONFIG.MAX_ANALYSIS_CELLS) {
                                range = selectedRange;
                                isSelection = true;
                                console.log(`Using selection: ${range.address} (${totalCells} cells)`);
                            } else {
                                console.log(`Selection too large (${totalCells} cells), using smart sampling instead`);
                                range = await this.getSmartSampledRange(context, worksheet, selectedRange);
                                isSelection = true;
                                isSampled = true;
                            }
                        }
                    }
                    
                    // If not using selection, use dynamic used range detection
                    if (!range) {
                        const usedRange = worksheet.getUsedRange();
                        if (usedRange) {
                            usedRange.load(['values', 'formulas', 'rowCount', 'columnCount', 'address', 'numberFormat']);
                            await context.sync();
                            
                            const totalCells = usedRange.rowCount * usedRange.columnCount;
                            console.log(`Detected used range: ${usedRange.address} (${totalCells} cells)`);
                            
                            if (totalCells <= this.CONFIG.MAX_ANALYSIS_CELLS) {
                                range = usedRange;
                            } else {
                                // For very large ranges, use smart sampling
                                console.log(`Used range too large (${totalCells} cells), applying smart sampling`);
                                range = await this.getSmartSampledRange(context, worksheet, usedRange);
                                isSampled = true;
                            }
                        }
                    }
                    
                    // Collect metadata about worksheet objects
                    await context.sync();
                    dataMetadata = {
                        tables: tables.items.map(t => ({ name: t.name, id: t.id })),
                        pivotTables: pivotTables.items.map(p => ({ name: p.name, id: p.id })),
                        charts: charts.items.map(c => ({ name: c.name, id: c.id })),
                        hasNamedRanges: await this.detectNamedRanges(context)
                    };
                    
                    if (!range || range.rowCount === 0 || range.columnCount === 0) {
                        this.currentWorksheetData = {
                            worksheetName: worksheet.name,
                            structuredData: { type: 'empty', headers: [], dataRows: [], keyRows: [] },
                            dataMetadata: dataMetadata,
                            summary: 'The current worksheet appears to be empty.'
                        };
                        resolve(this.currentWorksheetData);
                        return;
                    }
                    
                    // Get comprehensive data including formulas and formatting
                    const rawValues = range.values;
                    const rawFormulas = range.formulas;
                    const numberFormats = range.numberFormat;
                    const finalAddress = range.address;
                    
                    console.log(`Processing range: ${finalAddress} (${range.rowCount} rows × ${range.columnCount} columns)`);
                    
                    // Advanced data structure analysis with formulas and formatting
                    const structuredData = this.analyzeAdvancedTableStructure(rawValues, rawFormulas, numberFormats);
                    
                    // Collect additional workbook structure if requested
                    let workbookData = null;
                    if (includeAllSheets) {
                        workbookData = await this.analyzeWorkbookStructure(context);
                    }
                    
                    this.currentWorksheetData = {
                        worksheetName: worksheet.name,
                        address: finalAddress,
                        totalRows: range.rowCount,
                        totalCols: range.columnCount,
                        rawData: rawValues,
                        rawFormulas: rawFormulas,
                        numberFormats: numberFormats,
                        structuredData: structuredData,
                        dataMetadata: dataMetadata,
                        workbookData: workbookData,
                        isSelection: isSelection,
                        isSampled: isSampled,
                        summary: this.generateDataSummary(worksheet.name, finalAddress, range, isSelection, isSampled, structuredData, dataMetadata)
                    };
                    
                    resolve(this.currentWorksheetData);
                    
                } catch (error) {
                    reject(new Error('Failed to read Excel data: ' + error.message));
                }
            });
        });
    }

    // Smart sampling for very large datasets
    async getSmartSampledRange(context, worksheet, originalRange) {
        const totalRows = originalRange.rowCount;
        const totalCols = originalRange.columnCount;
        
        // Calculate sampling strategy
        const maxSampleRows = Math.min(Math.floor(Math.sqrt(this.CONFIG.MAX_ANALYSIS_CELLS)), totalRows);
        const maxSampleCols = Math.min(Math.floor(this.CONFIG.MAX_ANALYSIS_CELLS / maxSampleRows), totalCols);
        
        // Always include first few rows (headers) and last few rows (totals)
        const headerRows = Math.min(5, totalRows);
        const footerRows = Math.min(3, totalRows - headerRows);
        const middleRows = Math.max(0, maxSampleRows - headerRows - footerRows);
        
        // Sample columns proportionally 
        const sampleCols = Math.min(maxSampleCols, totalCols);
        
        console.log(`Smart sampling: ${maxSampleRows} rows × ${sampleCols} cols from ${totalRows} × ${totalCols}`);
        
        // Create a composite range with headers, sample middle, and footers
        const ranges = [];
        
        // Headers
        if (headerRows > 0) {
            const startCol = this.getColumnLetter(1);
            const endCol = this.getColumnLetter(sampleCols);
            ranges.push(`${startCol}1:${endCol}${headerRows}`);
        }
        
        // Sample middle rows
        if (middleRows > 0) {
            const step = Math.floor((totalRows - headerRows - footerRows) / middleRows);
            for (let i = 0; i < middleRows; i++) {
                const rowNum = headerRows + 1 + (i * step);
                const startCol = this.getColumnLetter(1);
                const endCol = this.getColumnLetter(sampleCols);
                ranges.push(`${startCol}${rowNum}:${endCol}${rowNum}`);
            }
        }
        
        // Footer rows
        if (footerRows > 0) {
            const startRow = totalRows - footerRows + 1;
            const startCol = this.getColumnLetter(1);
            const endCol = this.getColumnLetter(sampleCols);
            ranges.push(`${startCol}${startRow}:${endCol}${totalRows}`);
        }
        
        // For now, return a simplified range (top-left portion)
        // TODO: Implement true composite range sampling
        const simplifiedEndCol = this.getColumnLetter(sampleCols);
        const simplifiedRange = worksheet.getRange(`A1:${simplifiedEndCol}${maxSampleRows}`);
        simplifiedRange.load(['values', 'formulas', 'rowCount', 'columnCount', 'address', 'numberFormat']);
        await context.sync();
        
        simplifiedRange._isSampled = true;
        simplifiedRange._originalSize = { rows: totalRows, cols: totalCols };
        
        return simplifiedRange;
    }

    // Detect named ranges in the workbook
    async detectNamedRanges(context) {
        try {
            const namedItems = context.workbook.names;
            namedItems.load(['name', 'type']);
            await context.sync();
            return namedItems.items.length > 0;
        } catch (error) {
            console.log('Could not detect named ranges:', error);
            return false;
        }
    }

    // Analyze workbook structure across all sheets
    async analyzeWorkbookStructure(context) {
        try {
            const worksheets = context.workbook.worksheets;
            worksheets.load(['name', 'position']);
            await context.sync();
            
            const sheetInfo = [];
            for (const sheet of worksheets.items) {
                const usedRange = sheet.getUsedRange();
                if (usedRange) {
                    usedRange.load(['rowCount', 'columnCount', 'address']);
                    await context.sync();
                    
                    sheetInfo.push({
                        name: sheet.name,
                        position: sheet.position,
                        dataRange: usedRange.address,
                        rowCount: usedRange.rowCount,
                        columnCount: usedRange.columnCount,
                        totalCells: usedRange.rowCount * usedRange.columnCount
                    });
                } else {
                    sheetInfo.push({
                        name: sheet.name,
                        position: sheet.position,
                        dataRange: 'Empty',
                        rowCount: 0,
                        columnCount: 0,
                        totalCells: 0
                    });
                }
            }
            
            return {
                totalSheets: worksheets.items.length,
                sheets: sheetInfo,
                totalDataCells: sheetInfo.reduce((sum, sheet) => sum + sheet.totalCells, 0)
            };
        } catch (error) {
            console.log('Could not analyze workbook structure:', error);
            return null;
        }
    }

    // Advanced table structure analysis with formulas and formatting
    analyzeAdvancedTableStructure(rawValues, rawFormulas, numberFormats) {
        if (!rawValues || rawValues.length === 0) {
            return { type: 'empty', headers: [], dataRows: [], keyRows: [] };
        }
        
        const result = {
            type: 'advanced_financial_table',
            headers: [],
            columnHeaders: [],
            dataRows: [],
            keyRows: [],
            totalRows: [],
            quarterlyData: [],
            rowLabels: [],
            formulaAnalysis: this.analyzeFormulas(rawFormulas),
            formatAnalysis: this.analyzeNumberFormats(numberFormats),
            dataTypes: this.analyzeDataTypes(rawValues, numberFormats)
        };
        
        // Clean and process data with type detection
        const cleanData = rawValues.map((row, rowIndex) => 
            row.map((cell, colIndex) => {
                if (cell === null || cell === undefined) return '';
                if (typeof cell === 'string') return cell.trim();
                
                // Enhanced data type detection
                const format = numberFormats && numberFormats[rowIndex] && numberFormats[rowIndex][colIndex];
                return {
                    value: cell,
                    originalType: typeof cell,
                    format: format,
                    isFormula: rawFormulas && rawFormulas[rowIndex] && rawFormulas[rowIndex][colIndex] && 
                              typeof rawFormulas[rowIndex][colIndex] === 'string' && 
                              rawFormulas[rowIndex][colIndex].startsWith('='),
                    detectedType: this.detectCellType(cell, format)
                };
            })
        );
        
        // Find header row with enhanced detection
        let headerRowIndex = this.findHeaderRow(cleanData);
        console.log(`Header detection: Found header at row ${headerRowIndex}`);
        console.log(`Sample data preview:`, cleanData.slice(0, 3).map((row, i) => 
            `Row ${i}: [${row.slice(0, 5).map(cell => typeof cell === 'object' ? cell.value : cell).join(', ')}]`
        ));
        
        if (headerRowIndex !== -1) {
            result.columnHeaders = cleanData[headerRowIndex].map(cell => 
                typeof cell === 'object' ? cell.value : cell
            );
        }
        
        // Process data rows with enhanced analysis (including header row)
        console.log(`Processing ALL data rows starting from row 0 (including header)`);
        for (let i = 0; i < cleanData.length; i++) {
            const row = cleanData[i];
            const rowLabel = typeof row[0] === 'object' ? row[0].value : row[0];
            
            if (!rowLabel || rowLabel === '') continue;
            
            const rowValues = row.slice(1).map(cell => 
                typeof cell === 'object' ? cell.value : cell
            );
            
            const rowData = {
                rowIndex: i,
                label: rowLabel,
                values: rowValues,
                originalRow: row, // Keep enhanced data
                isTotal: this.isLikelyTotalRow(rowLabel),
                isSubtotal: this.isLikelySubtotalRow(rowLabel),
                category: this.categorizeRowLabel(rowLabel),
                hasFormulas: row.some(cell => 
                    typeof cell === 'object' && cell.isFormula
                ),
                dataTypes: row.map(cell => 
                    typeof cell === 'object' ? cell.detectedType : 'text'
                )
            };
            
            result.dataRows.push(rowData);
            result.rowLabels.push(rowLabel);
            
            // Enhanced key row identification
            if (rowData.isTotal || this.isKeyFinancialRow(rowLabel)) {
                result.keyRows.push(rowData);
            }
            
            if (rowData.isTotal) {
                result.totalRows.push(rowData);
            }
            
            // Extract quarterly data with enhanced detection
            if (this.hasQuarterlyPattern(result.columnHeaders)) {
                const quarterlyRow = this.extractQuarterlyData(rowData, result.columnHeaders);
                if (quarterlyRow.quarters.length > 0) {
                    result.quarterlyData.push(quarterlyRow);
                }
            }
        }
        
        return result;
    }

    // Analyze formulas in the dataset
    analyzeFormulas(rawFormulas) {
        if (!rawFormulas) return { hasFormulas: false, formulaCount: 0, types: [] };
        
        let formulaCount = 0;
        const formulaTypes = new Set();
        
        rawFormulas.forEach(row => {
            row.forEach(cell => {
                if (typeof cell === 'string' && cell.startsWith('=')) {
                    formulaCount++;
                    // Extract function names
                    const matches = cell.match(/([A-Z]+)\(/g);
                    if (matches) {
                        matches.forEach(match => {
                            formulaTypes.add(match.slice(0, -1)); // Remove the '('
                        });
                    }
                }
            });
        });
        
        return {
            hasFormulas: formulaCount > 0,
            formulaCount: formulaCount,
            types: Array.from(formulaTypes)
        };
    }

    // Analyze number formats
    analyzeNumberFormats(numberFormats) {
        if (!numberFormats) return { hasFormatting: false, types: [] };
        
        const formatTypes = new Set();
        
        numberFormats.forEach(row => {
            row.forEach(format => {
                if (format && format !== 'General') {
                    formatTypes.add(format);
                }
            });
        });
        
        return {
            hasFormatting: formatTypes.size > 0,
            types: Array.from(formatTypes)
        };
    }

    // Enhanced data type detection
    analyzeDataTypes(rawValues, numberFormats) {
        const types = { text: 0, number: 0, date: 0, currency: 0, percentage: 0, boolean: 0 };
        
        rawValues.forEach((row, rowIndex) => {
            row.forEach((cell, colIndex) => {
                const format = numberFormats && numberFormats[rowIndex] && numberFormats[rowIndex][colIndex];
                const detectedType = this.detectCellType(cell, format);
                if (types.hasOwnProperty(detectedType)) {
                    types[detectedType]++;
                }
            });
        });
        
        return types;
    }

    // Detect individual cell type
    detectCellType(value, format) {
        if (value === null || value === undefined || value === '') return 'empty';
        
        if (typeof value === 'boolean') return 'boolean';
        if (typeof value === 'number') {
            if (format) {
                if (format.includes('$') || format.includes('currency')) return 'currency';
                if (format.includes('%')) return 'percentage';
                if (format.includes('d') || format.includes('m') || format.includes('y')) return 'date';
            }
            return 'number';
        }
        
        if (typeof value === 'string') {
            // Try to detect dates
            if (!isNaN(Date.parse(value))) return 'date';
            // Try to detect numbers
            if (!isNaN(parseFloat(value)) && isFinite(value)) return 'number';
        }
        
        return 'text';
    }

    // Enhanced header row detection
    findHeaderRow(cleanData) {
        for (let i = 0; i < Math.min(5, cleanData.length); i++) {
            const row = cleanData[i];
            
            // Look for quarterly patterns
            const quarterlyPattern = row.filter(cell => {
                const val = typeof cell === 'object' ? cell.value : cell;
                return typeof val === 'string' && /^[1-4]Q\d{2}$/i.test(val);
            }).length;
            
            if (quarterlyPattern >= 2) return i;
            
            // Look for text-heavy rows (potential headers)
            const textCellCount = row.filter(cell => {
                const val = typeof cell === 'object' ? cell.value : cell;
                return typeof val === 'string' && val.length > 0 && isNaN(val);
            }).length;
            
            if (textCellCount >= 3) return i;
        }
        
        return 0; // Default to first row
    }

    // Enhanced financial row detection
    isKeyFinancialRow(label) {
        if (typeof label !== 'string') return false;
        
        const keyPatterns = [
            /revenue|sales|income/i,
            /expense|cost|opex|capex/i,
            /profit|loss|ebitda|ebit/i,
            /cash|flow|fcf/i,
            /margin|ratio/i,
            /growth|change/i
        ];
        
        return keyPatterns.some(pattern => pattern.test(label));
    }

    // Check for quarterly patterns in headers
    hasQuarterlyPattern(headers) {
        return headers.some(h => /^[1-4]Q\d{2}$/i.test(String(h)));
    }

    // Generate comprehensive data summary
    generateDataSummary(worksheetName, address, range, isSelection, isSampled, structuredData, dataMetadata) {
        let summary = `${isSelection ? 'Selected range' : 'Worksheet'} "${isSelection ? address : worksheetName}" `;
        summary += `contains ${range.rowCount} rows and ${range.columnCount} columns`;
        
        if (isSampled) {
            const originalSize = range._originalSize;
            if (originalSize) {
                summary += ` (sampled from ${originalSize.rows} × ${originalSize.cols})`;
            } else {
                summary += ' (intelligently sampled)';
            }
        }
        
        // Add structure information
        if (structuredData.keyRows && structuredData.keyRows.length > 0) {
            summary += `. Found ${structuredData.keyRows.length} key financial rows`;
        }
        
        if (structuredData.quarterlyData && structuredData.quarterlyData.length > 0) {
            summary += `, ${structuredData.quarterlyData.length} quarterly data series`;
        }
        
        // Add metadata
        if (dataMetadata.tables && dataMetadata.tables.length > 0) {
            summary += `, ${dataMetadata.tables.length} Excel tables`;
        }
        
        if (dataMetadata.pivotTables && dataMetadata.pivotTables.length > 0) {
            summary += `, ${dataMetadata.pivotTables.length} pivot tables`;
        }
        
        if (dataMetadata.charts && dataMetadata.charts.length > 0) {
            summary += `, ${dataMetadata.charts.length} charts`;
        }
        
        summary += '.';
        
        return summary;
    }

    // Helper functions (keeping existing ones and adding new ones)
    isLikelyTotalRow(label) {
        if (typeof label !== 'string') return false;
        const totalKeywords = /^(total|sum|grand total|net|aggregate|consolidated)/i;
        return totalKeywords.test(label.trim());
    }

    isLikelySubtotalRow(label) {
        if (typeof label !== 'string') return false;
        const subtotalKeywords = /subtotal|sub-total|sub total/i;
        return subtotalKeywords.test(label.trim());
    }

    categorizeRowLabel(label) {
        if (typeof label !== 'string') return 'data';
        return label.trim();
    }

    extractQuarterlyData(rowData, headers) {
        const quarters = [];
        
        headers.forEach((header, index) => {
            if (typeof header === 'string' && /^[1-4]Q\d{2}$/i.test(header)) {
                const value = rowData.values[index - 1]; // -1 because values excludes label column
                if (value !== null && value !== undefined && value !== '') {
                    quarters.push({
                        quarter: header,
                        value: parseFloat(value) || value,
                        columnIndex: index
                    });
                }
            }
        });
        
        return {
            label: rowData.label,
            labelType: rowData.category,
            quarters: quarters,
            isTotal: rowData.isTotal
        };
    }

    // Utility functions for Excel column manipulation
    getColumnLetter(columnNumber) {
        let columnLetter = '';
        while (columnNumber > 0) {
            const remainder = (columnNumber - 1) % 26;
            columnLetter = String.fromCharCode(65 + remainder) + columnLetter;
            columnNumber = Math.floor((columnNumber - 1) / 26);
        }
        return columnLetter;
    }

    getColumnNumber(columnLetter) {
        let columnNumber = 0;
        for (let i = 0; i < columnLetter.length; i++) {
            columnNumber = columnNumber * 26 + (columnLetter.charCodeAt(i) - 64);
        }
        return columnNumber;
    }

    // New method to read multiple worksheets
    async readAllWorksheets() {
        return new Promise((resolve, reject) => {
            Excel.run(async (context) => {
                try {
                    const workbookData = await this.analyzeWorkbookStructure(context);
                    const allSheetsData = [];
                    
                    const worksheets = context.workbook.worksheets;
                    worksheets.load(['name']);
                    await context.sync();
                    
                    for (const sheet of worksheets.items) {
                        try {
                            context.workbook.worksheets.getItem(sheet.name).activate();
                            await context.sync();
                            
                            const sheetData = await this.readCurrentWorksheetDataEnhanced(false, false);
                            allSheetsData.push(sheetData);
                        } catch (error) {
                            console.log(`Error reading sheet ${sheet.name}:`, error);
                            allSheetsData.push({
                                worksheetName: sheet.name,
                                error: error.message,
                                structuredData: { type: 'error' }
                            });
                        }
                    }
                    
                    resolve({
                        workbookData: workbookData,
                        allSheets: allSheetsData,
                        summary: `Read ${allSheetsData.length} worksheets from workbook`
                    });
                    
                } catch (error) {
                    reject(new Error('Failed to read all worksheets: ' + error.message));
                }
            });
        });
    }

    // Getter for current worksheet data
    getCurrentWorksheetData() {
        return this.currentWorksheetData;
    }

    // Clear current data
    clearCurrentData() {
        this.currentWorksheetData = null;
        this.workbookStructure = null;
    }
}

// Export for use in other modules
if (typeof window !== 'undefined') {
    window.DataService = DataService;
}