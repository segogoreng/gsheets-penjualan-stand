const sourceSpreadsheetId = '1RDYyrHPdoI_hYOxqT52HjIvgk5TponJkSIglx_fO0VE'; // ID file sumber "Copy of YUMBENTO PTC AGUSTUS 2025"
const targetSpreadsheetId = '1eIOD7Xl_wcMmZS83jTog1ilHM3shkWxg-0xfyw4PdoU'; // ID file tujuan "Copy of Rekap yumbento PTC 2025"
const targetSheetName = "des'25"; // perlu diganti sesuai bulan yg dikerjakan sesuai nama sheet tujuan

function onOpen() {
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

function copyDynamicRange() {
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
    const sourceRange = SpreadsheetApp.openById(sourceSpreadsheetId)
        .getSheetByName(sourceSheetName)
        .getRange(sourceRangeString)
        .getValues();

    // Tempelkan ke sheet tujuan dengan baris sama
    SpreadsheetApp.openById(targetSpreadsheetId)
        .getSheetByName(targetSheetName)
        .getRange(targetRangeString)
        .setValues(sourceRange);
}

function hapusNilaiDariRentang() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const ranges = [
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

function setFormulasBatch() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startRow = 2;
    const endRow = 380;
    const columns = [
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
    const refColumns = [
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

    for (let colIndex = 0; colIndex < columns.length; colIndex++) {
        const formulas = [];
        const col = columns[colIndex];
        const refCol = refColumns[colIndex];
        for (let i = startRow; i <= endRow; i++) {
            formulas.push(['=' + refCol + i + '*$C' + i]);
        }
        sheet
            .getRange(col + startRow + ':' + col + endRow)
            .setFormulas(formulas);
    }
}

function setProfitBatch() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startRow = 2;
    const endRow = 380;
    const columns = [
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
    const refColumns = [
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
    const twoColumns = [
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

    for (let colIndex = 0; colIndex < columns.length; colIndex++) {
        const formulas = [];
        const col = columns[colIndex];
        const refCol = refColumns[colIndex];
        const twoCol = twoColumns[colIndex];
        for (let i = startRow; i <= endRow; i++) {
            formulas.push([
                '=' + refCol + i + '- (' + twoCol + i + '*$B' + i + ')',
            ]);
        }
        sheet
            .getRange(col + startRow + ':' + col + endRow)
            .setFormulas(formulas);
    }
}

function isiFormulasticker() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const columns = ['E', 'H', 'K']; // Kolom spesifik
    const lastColumn = 'CN'; // Kolom terakhir

    // Mengisi kolom yang spesifik dulu
    columns.forEach(function (col) {
        sheet.getRange(col + '297').setFormula('=' + col + '285');
    });

    // Mengisi dari kolom berikutnya (L) sampai CN
    const startCol = sheet.getRange('L1').getColumn();
    const endCol = sheet.getRange(lastColumn + '1').getColumn();

    for (let col = startCol; col <= endCol; col++) {
        const colLetter = sheet
            .getRange(1, col)
            .getA1Notation()
            .replace(/\d/g, ''); // Dapatkan huruf kolom
        sheet.getRange(colLetter + '297').setFormula('=' + colLetter + '285');
    }
}

function isiFormulaProfitSticker() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startColumn = 'F'; // Kolom awal
    const endColumn = 'CR'; // Kolom akhir
    const step = 3; // Selisih antar kolom (F, I, L, ...)

    const startColIndex = sheet.getRange(startColumn + '1').getColumn();
    const endColIndex = sheet.getRange(endColumn + '1').getColumn();

    for (let col = startColIndex; col <= endColIndex; col += step) {
        const cell234 = sheet.getRange(297, col); // Baris 234, kolom ke-X
        const cell216 = sheet.getRange(285, col); // Baris 216, kolom ke-X
        cell234.setFormula('=' + cell216.getA1Notation()); // Masukkan formula
    }
}

function processYellowRows() {
    try {
        // Get source sheet name from user
        const ui = SpreadsheetApp.getUi();
        const response = ui.prompt(
            'Source Sheet Name',
            'Please enter the source sheet name:',
            ui.ButtonSet.OK_CANCEL
        );

        if (response.getSelectedButton() !== ui.Button.OK) {
            return; // User cancelled
        }

        const sourceSheetName = response.getResponseText().trim();
        if (!sourceSheetName) {
            ui.alert('Error', 'Sheet name cannot be empty!', ui.ButtonSet.OK);
            return;
        }

        const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
        const destSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);

        // Get the sheets
        const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
        const destSheet = destSpreadsheet.getSheetByName(targetSheetName);

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
                `Sheet "sept\'25" not found in destination spreadsheet!`,
                ui.ButtonSet.OK
            );
            return;
        }

        // Start scanning from row 2
        let currentRow = 2;
        let processedCount = 0;

        while (true) {
            // Get the value in column A
            const cellA = sourceSheet.getRange(currentRow, 1);
            const valueA = cellA.getValue();

            // Check if we've reached an empty cell - stop processing
            if (!valueA || valueA === '') {
                break;
            }

            // Check if the background color is yellow
            const backgroundColor = cellA.getBackground();

            // Common yellow color codes in Google Sheets
            const yellowColors = [
                '#ffff00',
                '#ffff99',
                '#fff2cc',
                '#fffacd',
                '#ffd966',
                '#f9cb9c',
            ];
            const isYellow = yellowColors.some(
                (color) => backgroundColor.toLowerCase() === color.toLowerCase()
            );

            if (isYellow) {
                // Get values from columns A and B
                const valueB = sourceSheet.getRange(currentRow, 2).getValue();

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
    } catch (error) {
        console.error('Error in processYellowRows:', error);
        SpreadsheetApp.getUi().alert(
            'Error',
            'An error occurred: ' + error.toString(),
            SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
}
