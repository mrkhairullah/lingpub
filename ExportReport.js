const CONFIG = {
  SHEETS: {
    MAINTENANCE_REPORT_DATA: {
      NAME: "Data Laporan Pemeliharaan",
      POSITION: {
        COL: {
          REPORT_NUMBER: "A",
          REPORT_ID: "B",
          REPORT_STATUS: "C",
          REPORT_DATE: "D",
          REPORTER_EMAIL: "E",
          REPORTER_NAME: "F",
          REPORTER_POSITION: "G",
          APP: "H",
          USER_ROLES: "I",
          DEVICE_TYPES: "J",
          WEBSITE_THEME: "K",
          ENVIRONMENT_TYPE: "L",
          REPORT_GROUP: "M",
          MUST_LOGGED_IN: "N",
          REPORT_TYPE: "O",
          REPORT_LEVEL: "P",
          REPORT_TITLE: "Q",
          REPORT_STEP: "R",
          REPORT_DESCRIPTION: "S",
          REPORT_SOLUTION: "T",
          REPORT_ATTACHMENT_LINK: "U",
        },
      },
    },
    MAINTENANCE_PROGRESS_REPORT_DATA: {
      NAME: "Data Laporan Progres Pemeliharaan",
      POSITION: {
        COL: {
          REPORT_NUMBER: "A",
          REPORT_ID: "B",
          REPORT_DATE: "C",
          REPORTER_EMAIL: "D",
          REPORTER_NAME: "E",
          REPORTER_POSITION: "F",
          MAINTENANCE_REPORT_ID: "G",
          REPORT_STATUS: "H",
          REPORT_DESCRIPTION: "I",
          REPORT_ATTACHMENT_LINK: "J",
        },
      },
    },
    TEMPLATE_MAINTENANCE_REPORT: {
      NAME: "Lihat Laporan",
      POSITION: {
        CELL: { INPUT_REPORT_ID: "D4" },
        CONTENT_AREA: "A1:J47",
      },
    },
    TEMPLATE_MAINTENANCE_PROGRESS_REPORT: {
      NAME: "Template Laporan Progres Pemeliharaan",
      POSITION: {
        CELL: { INPUT_REPORT_ID: "H2" },
        CONTENT_AREA: "A1:I19",
      },
    },
    EXPORT_REPORT: {
      NAME: "Ekspor Laporan",
      POSITION: {
        ROW: { DATA_START: 7 },
        COL: {
          REPORT_NUMBER: "A",
          REPORT_CHECKBOX: "B",
          REPORT_ID: "C",
          REPORT_DATE: "F",
          REPORT_STATUS: "H",
        },
      },
    },
  },

  EXPORT: {
    TEMP_SHEET_PREFIX: "_temp_export_",
    PDF_FOLDER_ID: "1CxhK80vs6FW4zzcggm8_nj6pzRja8PNn",
    TEMPLATE_RENDER_DELAY_MS: 5000,
    IMAGE_LOAD_DELAY_MS: 10000,
  },

  STATUS_OPTIONS: ["Belum Diperiksa", "Belum Diterapkan", "Diterapkan"],
};

// ================================================================
// TRIGGERS
// ================================================================

function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    _clearTable(_getSheet(ss, CONFIG.SHEETS.EXPORT_REPORT.NAME));
  } catch (_) {}
}

// ================================================================
// ENTRY POINTS
// ================================================================

function openFilterModal() {
  const html = HtmlService.createHtmlOutputFromFile("ExportFilterModal")
    .setWidth(440)
    .setHeight(430);
  SpreadsheetApp.getUi().showModalDialog(html, "Filter & Pengaturan Ekspor");
}

function loadReportsFromFilter(filter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const dateStart = _parseDate(filter.tanggalAwal);
    const dateEnd = _parseDate(filter.tanggalAkhir);

    if (!dateStart) throw new Error("Tanggal Awal tidak valid.");
    if (!dateEnd) throw new Error("Tanggal Akhir tidak valid.");
    if (dateStart > dateEnd)
      throw new Error(
        "Tanggal Awal tidak boleh lebih besar dari Tanggal Akhir.",
      );

    const reports = _filterReports(
      _getSheet(ss, CONFIG.SHEETS.MAINTENANCE_REPORT_DATA.NAME),
      {
        dateStart: _normalizeDate(dateStart),
        dateEnd: _normalizeDate(dateEnd),
        status: typeof filter.status === "string" ? filter.status.trim() : "",
      },
    );

    const exportSheet = _getSheet(ss, CONFIG.SHEETS.EXPORT_REPORT.NAME);
    _clearTable(exportSheet);

    if (reports.length === 0) {
      ui.alert(
        "Data Tidak Ditemukan",
        "Tidak ada laporan yang sesuai dengan rentang tanggal dan kriteria yang dipilih.",
        ui.ButtonSet.OK,
      );
      return;
    }

    _fillTable(exportSheet, reports);

    PropertiesService.getDocumentProperties().setProperties({
      includeProgress: String(filter.includeProgress === true),
      paperSize: filter.paperSize || "A4",
      orientation: filter.orientation || "portrait",
    });
  } catch (e) {
    ui.alert("Gagal memuat data:\n" + e.message);
  }
}

function processExport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const ids = _readCheckedIds(ss);

  if (ids.length === 0) {
    ui.alert("Tidak ada laporan yang dipilih. Centang minimal satu laporan.");
    return;
  }

  const props = PropertiesService.getDocumentProperties();
  const includeProgress = props.getProperty("includeProgress") === "true";
  const paperSize = props.getProperty("paperSize") || "A4";
  const orientation = props.getProperty("orientation") || "portrait";

  const originalSheets = ss.getSheets();
  const tempSheets = [];

  const maintenanceCfg = CONFIG.SHEETS.TEMPLATE_MAINTENANCE_REPORT.POSITION;
  const progressCfg =
    CONFIG.SHEETS.TEMPLATE_MAINTENANCE_PROGRESS_REPORT.POSITION;
  const maintenanceTemplate = _getSheet(
    ss,
    CONFIG.SHEETS.TEMPLATE_MAINTENANCE_REPORT.NAME,
  );
  const progressTemplate = includeProgress
    ? _getSheet(ss, CONFIG.SHEETS.TEMPLATE_MAINTENANCE_PROGRESS_REPORT.NAME)
    : null;
  const progressData = includeProgress
    ? _getSheet(ss, CONFIG.SHEETS.MAINTENANCE_PROGRESS_REPORT_DATA.NAME)
    : null;

  try {
    _cleanupLeftoverTempSheets(ss);

    for (let i = 0; i < ids.length; i++) {
      const id = ids[i];

      const tempMaint = _processTemplateToTempSheet(
        ss,
        maintenanceTemplate,
        id,
        maintenanceCfg.CELL.INPUT_REPORT_ID,
        maintenanceCfg.CONTENT_AREA,
        CONFIG.EXPORT.TEMP_SHEET_PREFIX + `M_${i + 1}`,
      );
      tempSheets.push(tempMaint);

      if (includeProgress) {
        const progIds = _findProgressIds(progressData, id);
        for (let j = 0; j < progIds.length; j++) {
          const tempProg = _processTemplateToTempSheet(
            ss,
            progressTemplate,
            progIds[j],
            progressCfg.CELL.INPUT_REPORT_ID,
            progressCfg.CONTENT_AREA,
            CONFIG.EXPORT.TEMP_SHEET_PREFIX + `P_${i + 1}_${j + 1}`,
          );
          tempSheets.push(tempProg);
        }
      }
    }

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
    ui.alert("Gagal mengekspor PDF:\n" + e.message + "\n\nDetail:\n" + e.stack);
  } finally {
    try {
      maintenanceTemplate
        .getRange(_topLeftCell(maintenanceCfg.CELL.INPUT_REPORT_ID))
        .clearContent();
    } catch (_) {}
    try {
      if (progressTemplate)
        progressTemplate
          .getRange(_topLeftCell(progressCfg.CELL.INPUT_REPORT_ID))
          .clearContent();
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
    try {
      ss.setActiveSheet(_getSheet(ss, CONFIG.SHEETS.EXPORT_REPORT.NAME));
    } catch (_) {}
  }
}

function resetExportTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _clearTable(_getSheet(ss, CONFIG.SHEETS.EXPORT_REPORT.NAME));
  ss.toast("Tabel daftar laporan berhasil direset.", "Reset Berhasil", 3);
}

// ================================================================
// TABLE OPERATIONS
// ================================================================

function _clearTable(sheet) {
  const ROW = CONFIG.SHEETS.EXPORT_REPORT.POSITION.ROW;
  const maxRows = sheet.getMaxRows();
  if (maxRows < ROW.DATA_START) return;

  const lastCol = Math.max(
    ...Object.values(CONFIG.SHEETS.EXPORT_REPORT.POSITION.COL).map(_colToIndex),
  );
  sheet
    .getRange(ROW.DATA_START, 1, maxRows - ROW.DATA_START + 1, lastCol)
    .clearContent()
    .clearDataValidations();
}

function _fillTable(sheet, reports) {
  if (reports.length === 0) return;

  const COL = CONFIG.SHEETS.EXPORT_REPORT.POSITION.COL;
  const ROW = CONFIG.SHEETS.EXPORT_REPORT.POSITION.ROW;
  const checkboxRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();

  reports.forEach((report, index) => {
    const row = ROW.DATA_START + index;
    sheet.getRange(row, _colToIndex(COL.REPORT_NUMBER)).setValue(index + 1);
    sheet
      .getRange(row, _colToIndex(COL.REPORT_CHECKBOX))
      .setValue(false)
      .setDataValidation(checkboxRule);
    sheet.getRange(row, _colToIndex(COL.REPORT_ID)).setValue(report.id);
    sheet.getRange(row, _colToIndex(COL.REPORT_DATE)).setValue(report.date);
    sheet.getRange(row, _colToIndex(COL.REPORT_STATUS)).setValue(report.status);
  });
}

function _readCheckedIds(ss) {
  const sheet = _getSheet(ss, CONFIG.SHEETS.EXPORT_REPORT.NAME);
  const COL = CONFIG.SHEETS.EXPORT_REPORT.POSITION.COL;
  const ROW = CONFIG.SHEETS.EXPORT_REPORT.POSITION.ROW;
  const lastRow = sheet.getLastRow();
  if (lastRow < ROW.DATA_START) return [];

  const checkboxCol = _colToIndex(COL.REPORT_CHECKBOX);
  const idCol = _colToIndex(COL.REPORT_ID);
  const startCol = Math.min(checkboxCol, idCol);
  const numCols = Math.max(checkboxCol, idCol) - startCol + 1;
  const data = sheet
    .getRange(ROW.DATA_START, startCol, lastRow - ROW.DATA_START + 1, numCols)
    .getValues();
  const checkboxOffset = checkboxCol - startCol;
  const idOffset = idCol - startCol;

  return data
    .filter((row) => row[checkboxOffset] === true && row[idOffset])
    .map((row) => row[idOffset]);
}

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
  const startCol = Math.min(idCol, statusCol, dateCol);
  const numCols = Math.max(idCol, statusCol, dateCol) - startCol + 1;
  const data = sheet.getRange(2, startCol, lastRow - 1, numCols).getValues();
  const idOffset = idCol - startCol;
  const statusOffset = statusCol - startCol;
  const dateOffset = dateCol - startCol;

  return data.reduce((acc, row) => {
    const id = row[idOffset];
    const status = row[statusOffset];
    const date = row[dateOffset];

    if (!id) return acc;
    const normalizedDate = date instanceof Date ? _normalizeDate(date) : null;
    if (!normalizedDate) return acc;
    if (normalizedDate < filter.dateStart || normalizedDate > filter.dateEnd)
      return acc;
    if (filter.status && status !== filter.status) return acc;

    acc.push({ id, date, status });
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
    .filter(
      (row) =>
        String(row[maintenanceIdOffset]) === String(maintenanceReportId) &&
        row[idOffset],
    )
    .map((row) => row[idOffset]);
}

// ================================================================
// TEMPLATE & PDF
// ================================================================

function _processTemplateToTempSheet(
  ss,
  templateSheet,
  id,
  inputCell,
  contentArea,
  newSheetName,
) {
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
      if (f && f.trim().toUpperCase().startsWith("=IMAGE(")) {
        // Tulis ulang formula gambar agar merender kembali
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
    "?format=pdf",
    `&size=${paperSize}`,
    `&portrait=${orientation === "portrait"}`,
    "&scale=2",
    "&sheetnames=false",
    "&printtitle=false",
    "&pagenumbers=false",
    "&gridlines=false",
    "&fzr=false",
  ].join("");

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(
      "Gagal mengekspor PDF. Response code: " + response.getResponseCode(),
    );
  }

  return response
    .getBlob()
    .setName("Laporan_Pemeliharaan_" + _formatTimestamp() + ".pdf");
}

function _openInNewTab(blob) {
  const file = DriveApp.getFolderById(CONFIG.EXPORT.PDF_FOLDER_ID).createFile(
    blob,
  );
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${file.getUrl()}','_blank');google.script.host.close();</script>`,
  )
    .setWidth(10)
    .setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, "Membuka PDF...");
}

// ================================================================
// UTILITIES
// ================================================================

function _colToIndex(letter) {
  return letter.toUpperCase().charCodeAt(0) - 64;
}

function _topLeftCell(a1Notation) {
  return a1Notation.split(":")[0];
}

function _parseDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === "number") {
    const d = new Date(Math.round((value - 25569) * 86400 * 1000));
    if (!isNaN(d.getTime())) return d;
  }
  if (typeof value === "string" && value.trim()) {
    const d = new Date(value.trim());
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function _normalizeDate(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function _getSheet(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet "${name}" tidak ditemukan.`);
  return sheet;
}

function _formatTimestamp() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmmss",
  );
}
