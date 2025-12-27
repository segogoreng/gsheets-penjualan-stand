const sourceSpreadsheetId = '1RDYyrHPdoI_hYOxqT52HjIvgk5TponJkSIglx_fO0VE'; // ID file sumber "Copy of YUMBENTO PTC AGUSTUS 2025"
const targetSpreadsheetId = '1eIOD7Xl_wcMmZS83jTog1ilHM3shkWxg-0xfyw4PdoU'; // ID file tujuan "Copy of Rekap yumbento PTC 2025"
const targetSheetName = "des'25"; // perlu diganti sesuai bulan yg dikerjakan sesuai nama sheet tujuan

function onOpen(): void {
    const ui = SpreadsheetApp.getUi();
    // Add a new menu to the spreadsheet.
    ui.createMenu('Custom Menu')
        .addItem('Clear Ranges', 'hapusNilaiDariRentang')
        .addItem('subtotal', 'setFormulasBatch')
        .addItem('profit', 'setProfitBatch')
        .addItem('isi formula sticker', 'isiFormulasticker')
        .addItem('Profit Sticker', 'isiFormulaProfitSticker')
        .addItem('Masukkan data penjualan', 'copyDynamicRange')
        .addItem('masukkan barang baru', 'processYellowRows')

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
        SpreadsheetApp.openById(targetSpreadsheetId).getSheetByName(
            targetSheetName
        );
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

function setFormulasBatch(): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startRow: number = 2;
    const endRow: number = 380;
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

function setProfitBatch(): void {
    const sheet: GoogleAppsScript.Spreadsheet.Sheet =
        SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startRow: number = 2;
    const endRow: number = 380;
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
            SpreadsheetApp.openById(targetSpreadsheetId);

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
            const cellA: GoogleAppsScript.Spreadsheet.Range = sourceSheet.getRange(
                currentRow,
                1
            );
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
