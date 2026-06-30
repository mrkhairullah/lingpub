const ui = SpreadsheetApp.getUi();

function onOpen() {
    createCustomMenu();
}

function createCustomMenu() {
    ui.createMenu('Aksi').addItem('Ekspor Laporan', 'showExportReportModal').addToUi();
}
