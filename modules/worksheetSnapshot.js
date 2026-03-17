/**
 * worksheetSnapshot.js
 *
 * Captures the full used-range state before each forward step and restores it
 * verbatim on backward navigation. No undo_code is ever executed.
 *
 * Key design rule:
 *   capture() ALWAYS snapshots the entire used range — never a subset.
 *   The `ranges` hint from sheet_context is ignored for correctness: scoping
 *   the snapshot to only the touched cells means restore() would clear the
 *   full sheet but only write back a subset, leaving everything else blank.
 *
 * Snapshot shape:
 * {
 *   address:      string,       // used-range address, sheet prefix stripped
 *   formulas:     any[][],
 *   numberFormat: any[][],
 *   fill:         any[][],
 *   fontColor:    any[][],
 *   fontBold:     any[][],
 *   fontSize:     any[][],
 *   alignment:    any[][],
 * }
 *
 * MAX_CELLS = 5 000. If the used range exceeds this, capture is skipped and
 * a warning is logged. Navigation will still work but back-steps won't restore.
 */
const WorksheetSnapshot = (() => {

    const MAX_CELLS = 5000;

    // ── Public API ────────────────────────────────────────────────────────────

    /**
     * Capture the full used-range state of the active worksheet.
     * @returns {object|null}
     */
    async function capture() {
        let snapshot = null;
        try {
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();

                // getUsedRangeOrNullObject avoids throwing on empty sheets
                const used = sheet.getUsedRangeOrNullObject();
                used.load('isNullObject');
                await ctx.sync();
                if (used.isNullObject) return;   // empty sheet — nothing to capture

                used.load([
                    'address',
                    'rowCount',
                    'columnCount',
                ]);
                await ctx.sync();

                const cells = used.rowCount * used.columnCount;
                if (cells > MAX_CELLS) {
                    _warn(`Used range has ${cells} cells (cap ${MAX_CELLS}). Snapshot skipped — back navigation won't restore this step.`);
                    return;
                }

                // Load all properties in one sync
                used.load([
                    'formulas',
                    'numberFormat',
                    'format/fill/color',
                    'format/font/color',
                    'format/font/bold',
                    'format/font/size',
                    'format/horizontalAlignment',
                ]);
                await ctx.sync();

                snapshot = {
                    address:      _stripSheet(used.address),
                    formulas:     _copy2d(used.formulas),
                    numberFormat: _copy2d(used.numberFormat),
                    fill:         _copy2d(used.format.fill.color),
                    fontColor:    _copy2d(used.format.font.color),
                    fontBold:     _copy2d(used.format.font.bold),
                    fontSize:     _copy2d(used.format.font.size),
                    alignment:    _copy2d(used.format.horizontalAlignment),
                };
            });
        } catch (err) {
            _warn(`capture failed: ${err.message}`);
        }
        return snapshot;
    }

    /**
     * Restore the worksheet to a previously captured snapshot.
     *
     * Strategy:
     *   1. Clear the entire snapshot region (removes anything added by the step).
     *   2. Write formulas + numberFormat as 2-D arrays (one call each — fast).
     *   3. Write fill / font / alignment per-cell (Office.js requires scalar
     *      writes for these; we batch all of them before a single ctx.sync).
     *
     * @param {object} snapshot — returned by capture()
     */
    async function restore(snapshot) {
        if (!snapshot) throw new Error('No snapshot to restore.');

        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();
            const addr  = snapshot.address;
            const rng   = sheet.getRange(addr);

            // ── 1. Clear the snapshot region completely ───────────────────────
            rng.clear('All');
            await ctx.sync();

            const rows = snapshot.formulas.length;
            const cols = snapshot.formulas[0]?.length || 0;
            if (!rows || !cols) return;

            // ── 2. Restore formulas and number formats (2-D array writes) ─────
            rng.formulas     = snapshot.formulas;
            rng.numberFormat = snapshot.numberFormat;

            // ── 3. Restore fill / font / alignment per-cell ───────────────────
            // All property sets are queued before ctx.sync so they go in one batch.
            const { startRow, startCol } = _parseTopLeft(addr);

            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const cellAddr = _cellAddr(startRow + r, startCol + c);
                    const cell     = sheet.getRange(cellAddr);

                    // Fill
                    const fill = _v(snapshot.fill, r, c);
                    if (fill && fill !== '' && fill !== 'null') {
                        cell.format.fill.color = fill;
                    } else {
                        cell.format.fill.clear();
                    }

                    // Font
                    const fc   = _v(snapshot.fontColor, r, c);
                    const bold = _v(snapshot.fontBold,  r, c);
                    const size = _v(snapshot.fontSize,  r, c);
                    if (fc   !== null && fc   !== undefined) cell.format.font.color = fc || null;
                    if (bold !== null && bold !== undefined) cell.format.font.bold  = !!bold;
                    if (size) cell.format.font.size = size;

                    // Alignment
                    const align = _v(snapshot.alignment, r, c);
                    if (align) cell.format.horizontalAlignment = align;
                }
            }

            await ctx.sync();
        });
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    function _stripSheet(addr) {
        const i = addr?.indexOf('!');
        return (i >= 0) ? addr.slice(i + 1) : (addr || '');
    }

    function _parseTopLeft(addr) {
        const m = addr.match(/^([A-Z]+)(\d+)/i);
        if (!m) return { startRow: 0, startCol: 0 };
        return { startRow: parseInt(m[2], 10) - 1, startCol: _colIndex(m[1].toUpperCase()) };
    }

    function _cellAddr(row, col) {
        let s = '', n = col;
        do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
        return s + (row + 1);
    }

    function _colIndex(letters) {
        let idx = 0;
        for (const ch of letters) idx = idx * 26 + (ch.charCodeAt(0) - 64);
        return idx - 1;
    }

    // Safe 2-D value accessor — handles both scalar and array returns from Office.js
    function _v(data, r, c) {
        if (!Array.isArray(data)) return data;
        const row = Array.isArray(data[r]) ? data[r] : data;
        return Array.isArray(row) ? row[c] : row;
    }

    function _copy2d(v) {
        if (Array.isArray(v)) return v.map(_copy2d);
        return v;
    }

    function _warn(msg) {
        try { ExecutionEngine.log('err', `[Snapshot] ${msg}`); } catch (_) { console.warn('[WorksheetSnapshot]', msg); }
    }

    return { capture, restore };
})();
