import {
    extractMonthAbbreviation,
    buildSourceSheetName,
    dayToColumnIndex,
    parseDateRange,
} from './helpers';

const sourceSpreadsheetId = '1B5kwiYri3x3qNde1gnpi9K9Bwk3Ruh8JtGuKptBz12Q'; // ID file sumber "Copy of YUMBENTO PTC AGUSTUS 2025"
const targetSheetName = "des'25"; // perlu diganti sesuai bulan yg dikerjakan sesuai nama sheet tujuan

function onOpen(): void {
    const ui = SpreadsheetApp.getUi();
    // Add a new menu to the spreadsheet.
    ui.createMenu('Custom Menu')
        .addItem('Clear Ranges', 'hapusNilaiDariRentang')
        .addItem('Isi Formula Subtotal dan Profit', 'setAllFormulas')
        .addItem('isi formula sticker', 'isiFormulasticker')
        .addItem('Profit Sticker', 'isiFormulaProfitSticker')
        .addItem(
            'Proses Data Penjualan (Multi-Tanggal)',
            'processMultipleDates'
        )
        .addSeparator()
        .addItem('Masukkan data penjualan (Manual)', 'copyDynamicRange')
        .addItem('masukkan barang baru (Manual)', 'processYellowRows')
        .addToUi();
}

function copyDynamicRange(): void {
    const ui = SpreadsheetApp.getUi();

    // --- Input dari user ---
    const sourceSheetName = ui
        .prompt('Masukkan nama sheet sumber (contoh: 1agt, 2agt):')
        .getResponseText();
    const endRow = parseInt(
        ui
            .prompt('Masukkan baris akhir sumber (contoh: 356 atau 376):')
            .getResponseText(),
        10
    );
    const targetColumn = ui
        .prompt('Masukkan kolom tujuan (contoh: G):')
        .getResponseText();

    // --- Spreadsheet Sumber ---
    const sourceColumn = 'D'; // kolom sumber selalu D
    const startRow = 2; // selalu mulai dari baris 2

    // Buat range string, contoh: "D2:D356"
    const sourceRangeString =
        sourceColumn + startRow + ':' + sourceColumn + endRow;
    const targetRangeString =
        targetColumn + startRow + ':' + targetColumn + endRow;

    // Ambil data dari sheet sumber
    const sourceSheet =
        SpreadsheetApp.openById(sourceSpreadsheetId).getSheetByName(
            sourceSheetName
        );
    if (!sourceSheet) {
        ui.alert('Sheet sumber tidak ditemukan: ' + sourceSheetName);
        return;
    }
    const sourceRange = sourceSheet.getRange(sourceRangeString).getValues();

    // Tempelkan ke sheet tujuan dengan baris sama
    const targetSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
    if (!targetSheet) {
        ui.alert('Sheet tujuan tidak ditemukan: ' + targetSheetName);
        return;
    }
    targetSheet.getRange(targetRangeString).setValues(sourceRange);
}

function hapusNilaiDariRentang(): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const ranges: string[] = [
        'D',
        'G',
        'J',
        'M',
        'P',
        'S',
        'V',
        'Y',
        'AB',
        'AE',
        'AH',
        'AK',
        'AN',
        'AQ',
        'AT',
        'AW',
        'AZ',
        'BC',
        'BF',
        'BI',
        'BL',
        'BO',
        'BR',
        'BU',
        'BX',
        'CA',
        'CD',
        'CG',
        'CJ',
        'CM',
        'CP',
    ];
    ranges.forEach((col) => {
        sheet.getRange(`${col}2:${col}437`).clearContent();
    });
}

function setFormulasBatch(endRow: number): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startRow: number = 2;
    const columns: string[] = [
        'E',
        'H',
        'K',
        'N',
        'Q',
        'T',
        'W',
        'Z',
        'AC',
        'AF',
        'AI',
        'AL',
        'AO',
        'AR',
        'AU',
        'AX',
        'BA',
        'BD',
        'BG',
        'BJ',
        'BM',
        'BP',
        'BS',
        'BV',
        'BY',
        'CB',
        'CE',
        'CH',
        'CK',
        'CN',
        'CQ',
    ]; // List of columns where formulas will be set
    const refColumns: string[] = [
        'D',
        'G',
        'J',
        'M',
        'P',
        'S',
        'V',
        'Y',
        'AB',
        'AE',
        'AH',
        'AK',
        'AN',
        'AQ',
        'AT',
        'AW',
        'AZ',
        'BC',
        'BF',
        'BI',
        'BL',
        'BO',
        'BR',
        'BU',
        'BX',
        'CA',
        'CD',
        'CG',
        'CJ',
        'CM',
        'CP',
    ]; // List of reference columns

    for (let colIndex: number = 0; colIndex < columns.length; colIndex++) {
        const formulas: string[][] = [];
        const col: string = columns[colIndex];
        const refCol: string = refColumns[colIndex];
        for (let i: number = startRow; i <= endRow; i++) {
            formulas.push(['=' + refCol + i + '*$C' + i]);
        }
        sheet
            .getRange(col + startRow + ':' + col + endRow)
            .setFormulas(formulas);
    }
}

function setProfitBatch(endRow: number): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startRow: number = 2;
    const columns: string[] = [
        'F',
        'I',
        'L',
        'O',
        'R',
        'U',
        'X',
        'AA',
        'AD',
        'AG',
        'AJ',
        'AM',
        'AP',
        'AS',
        'AV',
        'AY',
        'BB',
        'BE',
        'BH',
        'BK',
        'BN',
        'BQ',
        'BT',
        'BW',
        'BZ',
        'CC',
        'CF',
        'CI',
        'CL',
        'CO',
        'CR',
    ]; // List of reference columns
    const refColumns: string[] = [
        'E',
        'H',
        'K',
        'N',
        'Q',
        'T',
        'W',
        'Z',
        'AC',
        'AF',
        'AI',
        'AL',
        'AO',
        'AR',
        'AU',
        'AX',
        'BA',
        'BD',
        'BG',
        'BJ',
        'BM',
        'BP',
        'BS',
        'BV',
        'BY',
        'CB',
        'CE',
        'CH',
        'CK',
        'CN',
        'CQ',
    ]; // List of columns where formulas will be set
    const twoColumns: string[] = [
        'D',
        'G',
        'J',
        'M',
        'P',
        'S',
        'V',
        'Y',
        'AB',
        'AE',
        'AH',
        'AK',
        'AN',
        'AQ',
        'AT',
        'AW',
        'AZ',
        'BC',
        'BF',
        'BI',
        'BL',
        'BO',
        'BR',
        'BU',
        'BX',
        'CA',
        'CD',
        'CG',
        'CJ',
        'CM',
        'CP',
    ]; // List of reference columns

    for (let colIndex: number = 0; colIndex < columns.length; colIndex++) {
        const formulas: string[][] = [];
        const col: string = columns[colIndex];
        const refCol: string = refColumns[colIndex];
        const twoCol: string = twoColumns[colIndex];
        for (let i: number = startRow; i <= endRow; i++) {
            formulas.push([
                '=' + refCol + i + '- (' + twoCol + i + '*$B' + i + ')',
            ]);
        }
        sheet
            .getRange(col + startRow + ':' + col + endRow)
            .setFormulas(formulas);
    }
}

function setAllFormulas(): void {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Get last non-empty row in column A (product names)
    const lastProductRow = getLastNonEmptyRow(sheet, 1);

    // If no products found, show error
    if (lastProductRow < 2) {
        SpreadsheetApp.getUi().alert('Tidak ada produk ditemukan.');
        return;
    }

    // Set both formulas with detected end row
    setFormulasBatch(lastProductRow);
    setProfitBatch(lastProductRow);

    SpreadsheetApp.getUi().alert(
        `Formula subtotal dan profit berhasil diisi untuk baris 2 sampai ${lastProductRow}.`
    );
}

function isiFormulasticker(): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const columns: string[] = ['E', 'H', 'K']; // Kolom spesifik
    const lastColumn: string = 'CN'; // Kolom terakhir

    // Mengisi kolom yang spesifik dulu
    columns.forEach(function (col: string): void {
        sheet.getRange(col + '297').setFormula('=' + col + '285');
    });

    // Mengisi dari kolom berikutnya (L) sampai CN
    const startCol: number = sheet.getRange('L1').getColumn();
    const endCol: number = sheet.getRange(lastColumn + '1').getColumn();

    for (let col: number = startCol; col <= endCol; col++) {
        const colLetter: string = sheet
            .getRange(1, col)
            .getA1Notation()
            .replace(/\d/g, ''); // Dapatkan huruf kolom
        sheet.getRange(colLetter + '297').setFormula('=' + colLetter + '285');
    }
}

function isiFormulaProfitSticker(): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startColumn: string = 'F'; // Kolom awal
    const endColumn: string = 'CR'; // Kolom akhir
    const step: number = 3; // Selisih antar kolom (F, I, L, ...)

    const startColIndex: number = sheet.getRange(startColumn + '1').getColumn();
    const endColIndex: number = sheet.getRange(endColumn + '1').getColumn();

    for (let col: number = startColIndex; col <= endColIndex; col += step) {
        const cell234: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(
            297,
            col
        ); // Baris 234, kolom ke-X
        const cell216: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(
            285,
            col
        ); // Baris 216, kolom ke-X
        cell234.setFormula('=' + cell216.getA1Notation()); // Masukkan formula
    }
}

function processYellowRows(): void {
    try {
        // Get source sheet name from user
        const ui: GoogleAppsScript.Base.Ui = SpreadsheetApp.getUi();
        const response: GoogleAppsScript.Base.PromptResponse = ui.prompt(
            'Source Sheet Name',
            'Please enter the source sheet name:',
            ui.ButtonSet.OK_CANCEL
        );

        if (response.getSelectedButton() !== ui.Button.OK) {
            return; // User cancelled
        }

        const sourceSheetName: string = response.getResponseText().trim();
        if (!sourceSheetName) {
            ui.alert('Error', 'Sheet name cannot be empty!', ui.ButtonSet.OK);
            return;
        }

        const sourceSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet =
            SpreadsheetApp.openById(sourceSpreadsheetId);
        const destSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet =
            SpreadsheetApp.getActiveSpreadsheet();

        // Get the sheets
        const sourceSheet: GoogleAppsScript.Spreadsheet.Sheet | null =
            sourceSpreadsheet.getSheetByName(sourceSheetName);
        const destSheet: GoogleAppsScript.Spreadsheet.Sheet | null =
            destSpreadsheet.getSheetByName(targetSheetName);

        if (!sourceSheet) {
            ui.alert(
                'Error',
                `Sheet "${sourceSheetName}" not found in source spreadsheet!`,
                ui.ButtonSet.OK
            );
            return;
        }

        if (!destSheet) {
            ui.alert(
                'Error',
                `Sheet "sept'25" not found in destination spreadsheet!`,
                ui.ButtonSet.OK
            );
            return;
        }

        // Start scanning from row 2
        let currentRow: number = 2;
        let processedCount: number = 0;

        while (true) {
            // Get the value in column A
            const cellA: GoogleAppsScript.Spreadsheet.Range =
                sourceSheet.getRange(currentRow, 1);
            const valueA: unknown = cellA.getValue();

            // Check if we've reached an empty cell - stop processing
            if (!valueA || valueA === '') {
                break;
            }

            // Check if the background color is yellow
            const backgroundColor: string = cellA.getBackground();

            // Common yellow color codes in Google Sheets
            const yellowColors: string[] = [
                '#ffff00',
                '#ffff99',
                '#fff2cc',
                '#fffacd',
                '#ffd966',
                '#f9cb9c',
            ];
            const isYellow: boolean = yellowColors.some(
                (color: string): boolean =>
                    backgroundColor.toLowerCase() === color.toLowerCase()
            );

            if (isYellow) {
                // Get values from columns A and B
                const valueB: unknown = sourceSheet
                    .getRange(currentRow, 2)
                    .getValue();

                // Insert new row at the same position in destination sheet
                destSheet.insertRowBefore(currentRow);

                // Set values in columns A and C of the destination sheet
                destSheet.getRange(currentRow, 1).setValue(valueA); // Column A
                destSheet.getRange(currentRow, 3).setValue(valueB); // Column C

                processedCount++;

                console.log(
                    `Processed row ${currentRow}: A="${valueA}", B="${valueB}"`
                );
            }

            currentRow++;

            // Safety check to prevent infinite loops
            if (currentRow > 10000) {
                ui.alert(
                    'Warning',
                    'Stopped processing at row 10000 to prevent infinite loop.',
                    ui.ButtonSet.OK
                );
                break;
            }
        }

        // Show completion message
        ui.alert(
            'Complete',
            `Processing completed!\n` +
                `Rows processed: ${processedCount}\n` +
                `Last row checked: ${currentRow - 1}`,
            ui.ButtonSet.OK
        );
    } catch (error: unknown) {
        const errorMessage =
            error instanceof Error ? error.message : String(error);
        console.error('Error in processYellowRows:', error);
        SpreadsheetApp.getUi().alert(
            'Error',
            'An error occurred: ' + errorMessage,
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}

// Helper: Find last non-empty row in a column
function getLastNonEmptyRow(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    column: number = 1
): number {
    const lastRow = sheet.getLastRow();

    if (lastRow === 0) {
        return 0;
    }

    // Only fetch data up to lastRow instead of getMaxRows() (often 10000+)
    const data = sheet.getRange(1, column, lastRow, 1).getValues();

    // Find first empty cell (original behavior: returns row number of first empty)
    for (let i = 0; i < data.length; i++) {
        if (data[i][0] === '' || data[i][0] === null) {
            return i;
        }
    }
    return data.length;
}

// Interface for validation result
interface ValidationResult {
    isValid: boolean;
    sourceCount: number;
    targetCount: number;
}

// Validate that product count matches between source and target sheets
function validateProductCount(
    sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
    targetSheet: GoogleAppsScript.Spreadsheet.Sheet
): ValidationResult {
    const startRow = 2;

    // Get the last non-empty row for EACH sheet independently
    const sourceEndRow = getLastNonEmptyRow(sourceSheet, 1);
    const targetEndRow = getLastNonEmptyRow(targetSheet, 1);

    // Read product names up to each sheet's own last row
    const sourceNames =
        sourceEndRow >= startRow
            ? sourceSheet
                  .getRange(startRow, 1, sourceEndRow - startRow + 1, 1)
                  .getValues()
            : [];
    const targetNames =
        targetEndRow >= startRow
            ? targetSheet
                  .getRange(startRow, 1, targetEndRow - startRow + 1, 1)
                  .getValues()
            : [];

    // Count non-empty product names
    const sourceCount = sourceNames.filter(
        (row) => String(row[0]).trim() !== ''
    ).length;
    const targetCount = targetNames.filter(
        (row) => String(row[0]).trim() !== ''
    ).length;

    return {
        isValid: sourceCount === targetCount,
        sourceCount,
        targetCount,
    };
}

// Process yellow rows for a single day (extracted from processYellowRows, no UI prompts)
function processYellowRowsForDay(
    sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
    destSheet: GoogleAppsScript.Spreadsheet.Sheet
): number {
    const startRow = 2;
    const lastRow = sourceSheet.getLastRow();

    if (lastRow < startRow) {
        return 0;
    }

    const numRows = lastRow - startRow + 1;

    // BATCH READ: Get all values from columns A and B at once (1 API call)
    const allValues = sourceSheet.getRange(startRow, 1, numRows, 2).getValues();

    // BATCH READ: Get all background colors from column A (1 API call)
    const allBackgrounds = sourceSheet
        .getRange(startRow, 1, numRows, 1)
        .getBackgrounds();

    const yellowColors = new Set([
        '#ffff00',
        '#ffff99',
        '#fff2cc',
        '#fffacd',
        '#ffd966',
        '#f9cb9c',
    ]);

    let processedCount = 0;

    // Single loop: check and write immediately (same as original)
    for (let i = 0; i < numRows; i++) {
        const valueA = allValues[i][0];

        // Stop at first empty cell in column A (matches original behavior)
        if (!valueA || valueA === '') {
            break;
        }

        const backgroundColor = allBackgrounds[i][0].toLowerCase();
        const isYellow = yellowColors.has(backgroundColor);

        if (isYellow) {
            const valueB = allValues[i][1];
            const currentRow = startRow + i;

            destSheet.insertRowBefore(currentRow);
            destSheet.getRange(currentRow, 1).setValue(valueA);
            destSheet.getRange(currentRow, 3).setValue(valueB);

            processedCount++;
        }
    }

    return processedCount;
}

// Copy sales data for a single day (extracted from copyDynamicRange, no UI prompts)
function copySalesDataForDay(
    sourceSheet: GoogleAppsScript.Spreadsheet.Sheet,
    targetSheet: GoogleAppsScript.Spreadsheet.Sheet,
    endRow: number,
    targetColumnIndex: number
): number {
    const startRow = 2;
    const sourceColumn = 4; // Column D

    const sourceData = sourceSheet
        .getRange(startRow, sourceColumn, endRow - startRow + 1, 1)
        .getValues();

    targetSheet
        .getRange(startRow, targetColumnIndex, endRow - startRow + 1, 1)
        .setValues(sourceData);

    return endRow - startRow + 1;
}

// Interface for processing summary
interface ProcessingSummary {
    daysProcessed: number;
    totalNewProducts: number;
    totalSalesRowsCopied: number;
    warnings: string[];
}

// Main combined function: process multiple dates with validation
function processMultipleDates(): void {
    const ui = SpreadsheetApp.getUi();

    // Step 1: Get date range from user
    const response = ui.prompt(
        'Masukkan Rentang Tanggal',
        'Format: "1-3" untuk tanggal 1 sampai 3, atau "5" untuk tanggal 5 saja:',
        ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
        return;
    }

    // Step 2: Parse date range
    const dateRange = parseDateRange(response.getResponseText());
    if (!dateRange) {
        ui.alert(
            'Error',
            'Format tanggal tidak valid. Gunakan format "1-3" atau "5".',
            ui.ButtonSet.OK
        );
        return;
    }

    // Step 3: Extract month abbreviation
    let monthAbbr: string;
    try {
        monthAbbr = extractMonthAbbreviation(targetSheetName);
    } catch {
        ui.alert(
            'Error',
            `Tidak dapat mengekstrak bulan dari targetSheetName: ${targetSheetName}`,
            ui.ButtonSet.OK
        );
        return;
    }

    // Step 4: Open spreadsheets once
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

    if (!targetSheet) {
        ui.alert(
            'Error',
            `Sheet tujuan "${targetSheetName}" tidak ditemukan!`,
            ui.ButtonSet.OK
        );
        return;
    }

    // Step 5: Initialize summary
    const summary: ProcessingSummary = {
        daysProcessed: 0,
        totalNewProducts: 0,
        totalSalesRowsCopied: 0,
        warnings: [],
    };

    // Step 6: Process each day
    for (let day = dateRange.start; day <= dateRange.end; day++) {
        const sourceSheetName = buildSourceSheetName(day, monthAbbr);
        const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

        // Check if source sheet exists
        if (!sourceSheet) {
            summary.warnings.push(
                `Sheet "${sourceSheetName}" tidak ditemukan, dilewati.`
            );
            continue;
        }

        // Get last non-empty row
        const endRow = getLastNonEmptyRow(sourceSheet, 1);
        if (endRow < 2) {
            summary.warnings.push(
                `Sheet "${sourceSheetName}" kosong atau hanya memiliki header.`
            );
            continue;
        }

        // Process yellow rows (add new products)
        const newProductsCount = processYellowRowsForDay(
            sourceSheet,
            targetSheet
        );

        // Note: Source sheet is unchanged by processYellowRowsForDay (only target is modified)
        // so we can reuse endRow instead of re-fetching

        // Validate product count
        const validationResult = validateProductCount(sourceSheet, targetSheet);

        if (!validationResult.isValid) {
            const errorMsg =
                `Validasi gagal pada tanggal ${day} (sheet: ${sourceSheetName})!\n\n` +
                `Jumlah barang tidak sama:\n` +
                `  - Source: ${validationResult.sourceCount} barang\n` +
                `  - Target: ${validationResult.targetCount} barang\n\n` +
                `Proses DIHENTIKAN. Perbaiki data terlebih dahulu.`;

            ui.alert('Validasi Gagal', errorMsg, ui.ButtonSet.OK);
            return;
        }

        // Calculate target column
        const targetColumnIndex = dayToColumnIndex(day);

        // Copy sales data
        const rowsCopied = copySalesDataForDay(
            sourceSheet,
            targetSheet,
            endRow,
            targetColumnIndex
        );

        // Update summary
        summary.daysProcessed++;
        summary.totalNewProducts += newProductsCount;
        summary.totalSalesRowsCopied += rowsCopied;
    }

    // Step 7: Show summary
    let summaryMsg = `Proses selesai!\n\n`;
    summaryMsg += `Jumlah hari diproses: ${summary.daysProcessed}\n`;
    summaryMsg += `Total barang baru ditambahkan: ${summary.totalNewProducts}\n`;
    summaryMsg += `Total baris penjualan disalin: ${summary.totalSalesRowsCopied}\n`;

    if (summary.warnings.length > 0) {
        summaryMsg += `\nPeringatan:\n`;
        summary.warnings.forEach((w) => {
            summaryMsg += `- ${w}\n`;
        });
    }

    ui.alert('Ringkasan Proses', summaryMsg, ui.ButtonSet.OK);
}
