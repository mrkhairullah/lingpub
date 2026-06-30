const CONFIG = {
    SHEETS: {
        MAINTENANCE_REPORT_DATA: {
            NAME: 'Data Laporan Pemeliharaan',
            POSITION: {
                COL: {
                    REPORT_NUMBER: 'A',
                    REPORT_ID: 'B',
                    REPORT_STATUS: 'C',
                    REPORT_DATE: 'D',
                    REPORTER_EMAIL: 'E',
                    REPORTER_NAME: 'F',
                    REPORTER_POSITION: 'G',
                    APP: 'H',
                    USER_ROLES: 'I',
                    DEVICE_TYPES: 'J',
                    WEBSITE_THEME: 'K',
                    ENVIRONMENT_TYPE: 'L',
                    REPORT_GROUP: 'M',
                    MUST_LOGGED_IN: 'N',
                    REPORT_TYPE: 'O',
                    REPORT_LEVEL: 'P',
                    REPORT_TITLE: 'Q',
                    REPORT_STEP: 'R',
                    REPORT_DESCRIPTION: 'S',
                    REPORT_SOLUTION: 'T',
                    REPORT_ATTACHMENT_LINK: 'U',
                },
            },
        },
        MAINTENANCE_PROGRESS_REPORT_DATA: {
            NAME: 'Data Laporan Progres Pemeliharaan',
            POSITION: {
                COL: {
                    REPORT_NUMBER: 'A',
                    REPORT_ID: 'B',
                    REPORT_DATE: 'C',
                    REPORTER_EMAIL: 'D',
                    REPORTER_NAME: 'E',
                    REPORTER_POSITION: 'F',
                    MAINTENANCE_REPORT_ID: 'G',
                    REPORT_STATUS: 'H',
                    REPORT_DESCRIPTION: 'I',
                    REPORT_ATTACHMENT_LINK: 'J',
                },
            },
        },
        TEMPLATE_MAINTENANCE_REPORT: {
            NAME: 'Lihat Laporan',
            POSITION: {
                CELL: { INPUT_REPORT_ID: 'D4' },
                CONTENT_AREA: 'A1:J47',
            },
        },
        TEMPLATE_MAINTENANCE_PROGRESS_REPORT: {
            NAME: 'Template Laporan Progres Pemeliharaan',
            POSITION: {
                CELL: { INPUT_REPORT_ID: 'H2' },
                CONTENT_AREA: 'A1:I19',
            },
        },
    },

    EXPORT: {
        TEMP_SHEET_PREFIX: '_temp_export_',
        PDF_FOLDER_ID: '1CxhK80vs6FW4zzcggm8_nj6pzRja8PNn',
        TEMPLATE_RENDER_DELAY_MS: 5000,
        IMAGE_LOAD_DELAY_MS: 10000,
    },

    STATUS_OPTIONS: ['Belum Diperiksa', 'Sedang Diperiksa', 'Sudah Diperiksa'],
};

// ================================================================
// DATA FETCHING
// ================================================================

function _filterReports(sheet, filter) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const COL = CONFIG.SHEETS.MAINTENANCE_REPORT_DATA.POSITION.COL;
    const idCol = _colToIndex(COL.REPORT_ID);
    const statusCol = _colToIndex(COL.REPORT_STATUS);
    const dateCol = _colToIndex(COL.REPORT_DATE);
    const titleCol = _colToIndex(COL.REPORT_TITLE);
    const startCol = Math.min(idCol, statusCol, dateCol, titleCol);
    const numCols = Math.max(idCol, statusCol, dateCol, titleCol) - startCol + 1;
    const data = sheet.getRange(2, startCol, lastRow - 1, numCols).getValues();
    const idOffset = idCol - startCol;
    const statusOffset = statusCol - startCol;
    const dateOffset = dateCol - startCol;
    const titleOffset = titleCol - startCol;

    return data.reduce((acc, row) => {
        const id = row[idOffset];
        const status = row[statusOffset];
        const date = row[dateOffset];
        const title = row[titleOffset];

        if (!id) return acc;
        const normalizedDate = date instanceof Date ? _normalizeDate(date) : null;
        if (!normalizedDate) return acc;
        if (normalizedDate < filter.dateStart || normalizedDate > filter.dateEnd) return acc;
        if (filter.status && status !== filter.status) return acc;

        acc.push({
            id: String(id),
            date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            status,
            title: title || '',
        });
        return acc;
    }, []);
}

function _findProgressIds(sheet, maintenanceReportId) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const COL = CONFIG.SHEETS.MAINTENANCE_PROGRESS_REPORT_DATA.POSITION.COL;
    const idCol = _colToIndex(COL.REPORT_ID);
    const maintenanceIdCol = _colToIndex(COL.MAINTENANCE_REPORT_ID);
    const startCol = Math.min(idCol, maintenanceIdCol);
    const numCols = Math.max(idCol, maintenanceIdCol) - startCol + 1;
    const data = sheet.getRange(2, startCol, lastRow - 1, numCols).getValues();
    const idOffset = idCol - startCol;
    const maintenanceIdOffset = maintenanceIdCol - startCol;

    return data
        .filter((row) => String(row[maintenanceIdOffset]) === String(maintenanceReportId) && row[idOffset])
        .map((row) => row[idOffset]);
}

// ================================================================
// TEMPLATE & PDF
// ================================================================

function _processTemplateToTempSheet(ss, templateSheet, id, inputCell, contentArea, newSheetName) {
    const targetCell = _topLeftCell(inputCell);

    // 1. Render Template di Template Asli
    templateSheet.getRange(targetCell).setValue(id);
    SpreadsheetApp.flush();
    Utilities.sleep(CONFIG.EXPORT.TEMPLATE_RENDER_DELAY_MS);

    // 2. Simpan formula gambar sebelum dibekukan
    const sourceRange = templateSheet.getRange(contentArea);
    const formulas = sourceRange.getFormulas();

    // 3. Duplikasi Template
    const tempSheet = templateSheet.copyTo(ss);
    tempSheet.setName(newSheetName);
    const targetRange = tempSheet.getRange(contentArea);

    // 4. BEKUKAN DATA (Ini menggantikan setValue yang sering error)
    // Ini menyalin nilai (values) dari Template ke TempSheet secara instan.
    // Cara ini aman terhadap merged cells karena menggunakan engine Copy-Paste bawaan.
    sourceRange.copyTo(targetRange, { contentsOnly: true });

    // 5. Restorasi formula IMAGE() yang sempat hilang
    // Kita hanya menulis formula ke sel yang benar-benar membutuhkan gambar
    for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
            const f = formulas[r][c];
            if (f && f.trim().toUpperCase().startsWith('=IMAGE(')) {
                targetRange.getCell(r + 1, c + 1).setFormula(f);
            }
        }
    }

    return tempSheet;
}

function _cleanupLeftoverTempSheets(ss) {
    ss.getSheets()
        .filter((s) => s.getName().startsWith(CONFIG.EXPORT.TEMP_SHEET_PREFIX))
        .forEach((s) => {
            try {
                ss.deleteSheet(s);
            } catch (_) {}
        });
}

function _exportAllVisibleToPdf(ss, paperSize, orientation) {
    const url = [
        `https://docs.google.com/spreadsheets/d/${ss.getId()}/export`,
        '?format=pdf',
        `&size=${paperSize}`,
        `&portrait=${orientation === 'portrait'}`,
        '&scale=2',
        '&sheetnames=false',
        '&printtitle=false',
        '&pagenumbers=false',
        '&gridlines=false',
        '&fzr=false',
    ].join('');

    const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true,
    });

    if (response.getResponseCode() !== 200) {
        throw new Error('Gagal mengekspor PDF. Response code: ' + response.getResponseCode());
    }

    return response.getBlob().setName('Laporan_Pemeliharaan_' + _formatTimestamp() + '.pdf');
}

function _openInNewTab(blob) {
    const file = DriveApp.getFolderById(CONFIG.EXPORT.PDF_FOLDER_ID).createFile(blob);
    const html = HtmlService.createHtmlOutput(
        `<script>window.open('${file.getUrl()}','_blank');google.script.host.close();</script>`
    )
        .setWidth(10)
        .setHeight(10);
    SpreadsheetApp.getUi().showModalDialog(html, 'Membuka PDF...');
}
