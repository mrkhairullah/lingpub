function showExportReportModal() {
    const template = HtmlService.createTemplateFromFile('ExportReportModal');
    const html = template.evaluate().setWidth(1400).setHeight(900);
    ui.showModalDialog(html, 'Ekspor Laporan');
}

function getFilteredReports(filter) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const dateStart = _parseDate(filter.tanggalAwal);
    const dateEnd = _parseDate(filter.tanggalAkhir);
    const status = typeof filter.status === 'string' ? filter.status.trim() : '';

    if (!dateStart) throw new Error('Tanggal Awal tidak valid.');
    if (!dateEnd) throw new Error('Tanggal Akhir tidak valid.');
    if (dateStart > dateEnd) {
        throw new Error('Tanggal Awal tidak boleh lebih besar dari Tanggal Akhir.');
    }
    if (status && CONFIG.STATUS_OPTIONS.indexOf(status) === -1) {
        throw new Error('Status laporan tidak valid.');
    }

    return _filterReports(_getSheet(ss, CONFIG.SHEETS.MAINTENANCE_REPORT_DATA.NAME), {
        dateStart: _normalizeDate(dateStart),
        dateEnd: _normalizeDate(dateEnd),
        status: status,
    });
}

function processExport(payload) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ids = Array.isArray(payload.ids) ? payload.ids : [];

    if (ids.length === 0) {
        throw new Error('Tidak ada laporan yang dipilih. Pilih minimal satu laporan.');
    }

    const includeProgress = payload.includeProgress === true;
    const paperSize = payload.paperSize || 'A4';
    const orientation = payload.orientation || 'portrait';

    const originalSheets = ss.getSheets();
    const tempSheets = [];

    const maintenanceCfg = CONFIG.SHEETS.TEMPLATE_MAINTENANCE_REPORT.POSITION;
    const progressCfg = CONFIG.SHEETS.TEMPLATE_MAINTENANCE_PROGRESS_REPORT.POSITION;
    const maintenanceTemplate = _getSheet(ss, CONFIG.SHEETS.TEMPLATE_MAINTENANCE_REPORT.NAME);
    const progressTemplate = includeProgress
        ? _getSheet(ss, CONFIG.SHEETS.TEMPLATE_MAINTENANCE_PROGRESS_REPORT.NAME)
        : null;
    const progressData = includeProgress ? _getSheet(ss, CONFIG.SHEETS.MAINTENANCE_PROGRESS_REPORT_DATA.NAME) : null;

    try {
        _cleanupLeftoverTempSheets(ss);

        ids.forEach((id, i) => {
            const tempMaint = _processTemplateToTempSheet(
                ss,
                maintenanceTemplate,
                id,
                maintenanceCfg.CELL.INPUT_REPORT_ID,
                maintenanceCfg.CONTENT_AREA,
                CONFIG.EXPORT.TEMP_SHEET_PREFIX + `M_${i + 1}`
            );
            tempSheets.push(tempMaint);

            if (includeProgress) {
                const progIds = _findProgressIds(progressData, id);
                progIds.forEach((progId, j) => {
                    const tempProg = _processTemplateToTempSheet(
                        ss,
                        progressTemplate,
                        progId,
                        progressCfg.CELL.INPUT_REPORT_ID,
                        progressCfg.CONTENT_AREA,
                        CONFIG.EXPORT.TEMP_SHEET_PREFIX + `P_${i + 1}_${j + 1}`
                    );
                    tempSheets.push(tempProg);
                });
            }
        });

        SpreadsheetApp.flush();
        Utilities.sleep(CONFIG.EXPORT.IMAGE_LOAD_DELAY_MS);

        originalSheets.forEach((s) => {
            try {
                s.hideSheet();
            } catch (_) {}
        });

        SpreadsheetApp.flush();
        _openInNewTab(_exportAllVisibleToPdf(ss, paperSize, orientation));
    } catch (e) {
        throw new Error('Gagal mengekspor PDF: ' + e.message);
    } finally {
        try {
            maintenanceTemplate.getRange(_topLeftCell(maintenanceCfg.CELL.INPUT_REPORT_ID)).clearContent();
        } catch (_) {}
        try {
            if (progressTemplate) {
                progressTemplate.getRange(_topLeftCell(progressCfg.CELL.INPUT_REPORT_ID)).clearContent();
            }
        } catch (_) {}
        originalSheets.forEach((s) => {
            try {
                s.showSheet();
            } catch (_) {}
        });
        tempSheets.forEach((s) => {
            try {
                ss.deleteSheet(s);
            } catch (_) {}
        });
    }
}
