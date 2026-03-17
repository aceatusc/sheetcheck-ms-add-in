/**
 * worksheetSnapshot.js
 *
 * Captures the full visual + data state of the active worksheet's used range
 * before each step executes, and restores it verbatim on undo.
 *
 * Snapshot shape:
 * {
 *   address:      string,          // e.g. "Sheet1!A1:E8"
 *   formulas:     string[][],      // raw formulas (falls back to values where no formula)
 *   numberFormat: string[][],
 *   fill:         (string|null)[][], // fill colour per cell, null = no fill
 *   fontColor:    (string|null)[][],
 *   fontBold:     boolean[][],
 *   fontSize:     number[][],
 *   alignment:    string[][],      // horizontal alignment
 * }
 *
 * Why used-range only:
 *   getUsedRange() is a single cheap Office.js call. It covers every cell the
 *   LLM segments will ever touch for a typical spreadsheet task. Full-worksheet
 *   snapshots would be prohibitively large and slow for an add-in.
 *
 * Why formulas not values:
 *   Restoring values would silently drop formulas. We load formulas and write
 *   them back so =SUM(...) cells stay as formulas after an undo.
 */
const WorksheetSnapshot = (() => {

    /**
     * Capture the current state of the active worksheet's used range.
     * Returns a snapshot object, or null if the sheet is empty / Excel.run fails.
     */
    async function capture() {
        let snapshot = null;
        try {
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                let range;
                try {
                    range = sheet.getUsedRange();
                } catch (_) {
                    // Sheet is empty — nothing to snapshot
                    return;
                }

                range.load([
                    'address',
                    'formulas',
                    'numberFormat',
                    'format/fill/color',
                    'format/font/color',
                    'format/font/bold',
                    'format/font/size',
                    'format/horizontalAlignment',
                ]);
                await ctx.sync();

                // Cell-level font/fill properties require loading per-cell
                // for full fidelity. We use the range-level arrays which Office.js
                // returns as 2-D arrays matching the range dimensions.
                snapshot = {
                    address:      range.address,
                    formulas:     _deepCopy(range.formulas),
                    numberFormat: _deepCopy(range.numberFormat),
                    fill:         _deepCopy(range.format.fill.color),
                    fontColor:    _deepCopy(range.format.font.color),
                    fontBold:     _deepCopy(range.format.font.bold),
                    fontSize:     _deepCopy(range.format.font.size),
                    alignment:    _deepCopy(range.format.horizontalAlignment),
                };
            });
        } catch (err) {
            // Non-fatal: if capture fails, undo simply won't be available
            console.warn('[WorksheetSnapshot] capture failed:', err.message);
        }
        return snapshot;
    }

    /**
     * Restore the worksheet to a previously captured snapshot.
     * Writes formulas, formats, fill, font, and alignment back in one Excel.run.
     * @param {object} snapshot  — object returned by capture()
     */
    async function restore(snapshot) {
        if (!snapshot) throw new Error('No snapshot to restore.');

        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();

            // Strip the sheet-name prefix from address (e.g. "Sheet1!A1:E8" → "A1:E8")
            const addr = snapshot.address.includes('!')
                ? snapshot.address.split('!')[1]
                : snapshot.address;

            const range = sheet.getRange(addr);

            // ── 1. Clear everything in the range first ────────────────────────
            // This handles cells that were added by the step but are outside the
            // original snapshot — they'll be left with defaults after the restore.
            range.clear('All');
            await ctx.sync();

            // ── 2. Restore formulas / values ──────────────────────────────────
            range.formulas = snapshot.formulas;

            // ── 3. Restore number format ──────────────────────────────────────
            range.numberFormat = snapshot.numberFormat;

            // ── 4. Restore fill colour ────────────────────────────────────────
            // Office.js fill.color is a scalar on a Range (applies to whole range)
            // but snapshot stores a 2-D array from the per-cell load.
            // We must restore cell-by-cell when values differ across the range.
            _restorePerCell(ctx, sheet, addr, snapshot, (cell, r, c) => {
                const fill = _val(snapshot.fill, r, c);
                if (fill && fill !== '') {
                    cell.format.fill.color = fill;
                } else {
                    cell.format.fill.clear();
                }
            });

            // ── 5. Restore font properties ────────────────────────────────────
            _restorePerCell(ctx, sheet, addr, snapshot, (cell, r, c) => {
                const fc   = _val(snapshot.fontColor, r, c);
                const bold = _val(snapshot.fontBold,  r, c);
                const size = _val(snapshot.fontSize,  r, c);
                if (fc !== undefined)   cell.format.font.color = fc || null;
                if (bold !== undefined) cell.format.font.bold  = !!bold;
                if (size !== undefined && size) cell.format.font.size = size;
            });

            // ── 6. Restore alignment ──────────────────────────────────────────
            _restorePerCell(ctx, sheet, addr, snapshot, (cell, r, c) => {
                const align = _val(snapshot.alignment, r, c);
                if (align) cell.format.horizontalAlignment = align;
            });

            await ctx.sync();
        });
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    /**
     * Office.js returns scalar strings for range-level font/fill properties
     * when all cells share the same value, and 2-D arrays when they differ.
     * Normalise both cases into a consistent 2-D array accessor.
     */
    function _val(data, row, col) {
        if (Array.isArray(data)) {
            const r = Array.isArray(data[row]) ? data[row] : data;
            return Array.isArray(r) ? r[col] : r;
        }
        return data; // scalar — same value for every cell
    }

    /**
     * Iterate every cell in a range by address and call cb(cellRange, row, col).
     * We batch all property sets into one ctx.sync at the end (done by caller).
     */
    function _restorePerCell(ctx, sheet, addr, snapshot, cb) {
        const rows = snapshot.formulas.length;
        const cols = snapshot.formulas[0]?.length || 0;

        // Parse top-left cell from address (handles both "A1:E8" and "A1" forms)
        const match = addr.match(/^([A-Z]+)(\d+)/);
        if (!match) return;
        const startCol = _colIndex(match[1]);  // 0-based
        const startRow = parseInt(match[2], 10) - 1; // 0-based

        for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
                const cellAddr = _cellAddr(startRow + r, startCol + c);
                const cell = sheet.getRange(cellAddr);
                cb(cell, r, c);
            }
        }
    }

    /** Convert 0-based row/col to Excel address string, e.g. (0,0) → "A1". */
    function _cellAddr(row, col) {
        let colStr = '';
        let n = col;
        do {
            colStr = String.fromCharCode(65 + (n % 26)) + colStr;
            n = Math.floor(n / 26) - 1;
        } while (n >= 0);
        return colStr + (row + 1);
    }

    /** Convert column letter(s) to 0-based index, e.g. "A" → 0, "Z" → 25, "AA" → 26. */
    function _colIndex(letters) {
        let idx = 0;
        for (const ch of letters) {
            idx = idx * 26 + (ch.charCodeAt(0) - 64);
        }
        return idx - 1;
    }

    function _deepCopy(v) {
        if (Array.isArray(v)) return v.map(_deepCopy);
        return v;
    }

    return { capture, restore };
})();
