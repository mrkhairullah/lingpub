// ================================================================
// HTML INCLUDE
// ================================================================

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ================================================================
// SHEET HELPERS
// ================================================================

function _getSheet(ss, name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) throw new Error(`Sheet "${name}" tidak ditemukan.`);
    return sheet;
}

function _colToIndex(letter) {
    return letter.toUpperCase().charCodeAt(0) - 64;
}

function _topLeftCell(a1Notation) {
    return a1Notation.split(':')[0];
}

// ================================================================
// DATE HELPERS
// ================================================================

function _parseDate(value) {
    if (value instanceof Date && !isNaN(value.getTime())) return value;
    if (typeof value === 'number') {
        const d = new Date(Math.round((value - 25569) * 86400 * 1000));
        if (!isNaN(d.getTime())) return d;
    }
    if (typeof value === 'string' && value.trim()) {
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

function _formatTimestamp() {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
}
