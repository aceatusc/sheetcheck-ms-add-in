/**
 * worksheetSnapshot.js
 *
 * Captures cell content (formulas + values) before each forward step and
 * restores it verbatim on backward navigation.
 *
 * Formatting (fill, font, alignment, numberFormat) is intentionally NOT
 * captured or restored — it is owned by the code segments themselves and
 * does not need snapshot-based undo.
 *
 * Snapshot shape:
 * {
 *   address:  string,      // used-range address, sheet prefix stripped
 *   rows:     number,
 *   cols:     number,
 *   formulas: any[][],     // raw formula strings or primitive values
 *   values:   any[][],     // computed values — used for non-formula cells
 * }
 *
 * MAX_CELLS = 5000. If exceeded, capture is skipped and a warning is logged.
 */
const WorksheetSnapshot = (() => {

    const MAX_CELLS = 5000;

    // ── capture ───────────────────────────────────────────────────────────────

    async function capture() {
        let snapshot = null;
        try {
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();

                const used = sheet.getUsedRangeOrNullObject();
                used.load('isNullObject');
                await ctx.sync();
                if (used.isNullObject) return;

                used.load(['address', 'rowCount', 'columnCount']);
                await ctx.sync();

                const rows  = used.rowCount;
                const cols  = used.columnCount;
                const cells = rows * cols;

                if (cells > MAX_CELLS) {
                    _warn(`Used range ${rows}×${cols} = ${cells} cells exceeds cap (${MAX_CELLS}). Snapshot skipped.`);
                    return;
                }

                used.load(['formulas', 'values']);
                await ctx.sync();

                snapshot = {
                    address: _stripSheet(used.address),
                    rows,
                    cols,
                    formulas: _copy2d(used.formulas),
                    values:   _copy2d(used.values),
                };
            });
        } catch (err) {
            _warn(`capture failed: ${err.message}`);
        }
        return snapshot;
    }

    // ── restore ───────────────────────────────────────────────────────────────

    async function restore(snapshot) {
        if (!snapshot) throw new Error('No snapshot to restore.');

        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();

            // Clear the entire current used range so any content added by
            // later steps (outside snapshot.address) is removed first.
            const currentUsed = sheet.getUsedRangeOrNullObject();
            currentUsed.load('isNullObject');
            await ctx.sync();
            if (!currentUsed.isNullObject) currentUsed.clear('Contents');
            sheet.getRange(snapshot.address).clear('Contents');
            await ctx.sync();

            const { rows, cols } = snapshot;
            if (!rows || !cols) return;

            const { startRow, startCol } = _parseTopLeft(snapshot.address);

            // Write each cell individually to avoid the batch formulas-array
            // restriction where Office.js rejects empty strings in the array.
            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const formula = snapshot.formulas[r][c];
                    const value   = snapshot.values[r][c];

                    if (typeof formula === 'string' && formula.startsWith('=')) {
                        // Formula cell — restore the formula string
                        sheet.getRange(_cellAddr(startRow + r, startCol + c))
                             .formulas = [[formula]];
                    } else if (value !== null && value !== undefined && value !== '') {
                        // Value cell — restore the computed value
                        sheet.getRange(_cellAddr(startRow + r, startCol + c))
                             .values = [[value]];
                    }
                    // Empty cell — already cleared above, nothing to write
                }
            }

            await ctx.sync();
        });
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    function _stripSheet(addr) {
        const i = addr?.indexOf('!');
        return (i >= 0) ? addr.slice(i + 1) : (addr || '');
    }

    function _parseTopLeft(addr) {
        const m = addr.match(/^([A-Z]+)(\d+)/i);
        if (!m) return { startRow: 0, startCol: 0 };
        return {
            startRow: parseInt(m[2], 10) - 1,
            startCol: _colIndex(m[1].toUpperCase()),
        };
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

    function _copy2d(v) {
        if (Array.isArray(v)) return v.map(r => Array.isArray(r) ? [...r] : r);
        return v;
    }

    function _warn(msg) {
        try { ExecutionEngine.log('err', `[Snapshot] ${msg}`); }
        catch (_) { console.warn('[WorksheetSnapshot]', msg); }
    }

    return { capture, restore };
})();
