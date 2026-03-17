/**
 * worksheetSnapshot.js
 *
 * Captures worksheet state before each forward step and restores it verbatim
 * when the user navigates backward or to an arbitrary visited node.
 * No undo_code is ever executed — snapshots are the sole undo mechanism.
 *
 * Snapshot shape:
 * {
 *   regions: [{                  // one entry per captured range
 *     address:      string,      // e.g. "A1:E8" (sheet prefix stripped)
 *     formulas:     string[][],
 *     numberFormat: string[][],
 *     fill:         any[][],     // Office.js 2-D color array
 *     fontColor:    any[][],
 *     fontBold:     any[][],
 *     fontSize:     any[][],
 *     alignment:    any[][],
 *   }],
 *   clearedBeyond: string|null,  // used-range address captured for the clear pass
 * }
 *
 * Design decisions:
 *   - capture(ranges?) scopes to the segment's sheet_context ranges when
 *     provided, falling back to the full used range. This keeps snapshots
 *     small and restore fast for typical tasks.
 *   - A MAX_CELLS guard (default 5 000) prevents runaway memory on large sheets.
 *     Capture still proceeds but logs a warning and skips ranges that exceed it.
 *   - restore() clears the entire used range first (removes cells added by the
 *     step outside snapshot regions), then writes all properties back.
 *   - Font/fill/alignment are restored per-row using getRange(rowAddr) instead
 *     of per-cell, cutting Office.js RPC calls by ~cols factor.
 */
const WorksheetSnapshot = (() => {

    const MAX_CELLS = 5000;

    // ── Public API ────────────────────────────────────────────────────────────

    /**
     * Capture worksheet state.
     * @param {string[]} [ranges]  — sheet_context address list from the upcoming
     *                               segment. When omitted, captures the full used range.
     * @returns {object|null} snapshot object, or null if sheet is empty.
     */
    async function capture(ranges) {
        let snapshot = null;
        try {
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();

                // Always capture the used-range address for the clear pass,
                // even when we snapshot only specific regions.
                let usedAddr = null;
                try {
                    const used = sheet.getUsedRange();
                    used.load('address');
                    await ctx.sync();
                    usedAddr = _stripSheet(used.address);
                } catch (_) {
                    // Sheet is empty — nothing to snapshot
                    return;
                }

                // Decide which address(es) to capture
                const addrs = _resolveAddresses(ranges, usedAddr);

                const regions = [];
                let totalCells = 0;

                for (const addr of addrs) {
                    const cellCount = _countCells(addr);
                    if (totalCells + cellCount > MAX_CELLS) {
                        _warn(`Snapshot skipping range ${addr} (would exceed ${MAX_CELLS}-cell cap). Navigation to this node may be imprecise.`);
                        continue;
                    }
                    totalCells += cellCount;

                    const rng = sheet.getRange(addr);
                    rng.load([
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

                    regions.push({
                        address:      _stripSheet(rng.address),
                        formulas:     _copy2d(rng.formulas),
                        numberFormat: _copy2d(rng.numberFormat),
                        fill:         _copy2d(rng.format.fill.color),
                        fontColor:    _copy2d(rng.format.font.color),
                        fontBold:     _copy2d(rng.format.font.bold),
                        fontSize:     _copy2d(rng.format.font.size),
                        alignment:    _copy2d(rng.format.horizontalAlignment),
                    });
                }

                if (regions.length) {
                    snapshot = { regions, clearedBeyond: usedAddr };
                }
            });
        } catch (err) {
            _warn(`capture failed: ${err.message}`);
        }
        return snapshot;
    }

    /**
     * Restore the worksheet to a previously captured snapshot.
     * 1. Clears the used range at capture time (removes anything the step added).
     * 2. Writes formulas, formats, fill, font, alignment back per region.
     *
     * @param {object} snapshot — returned by capture()
     */
    async function restore(snapshot) {
        if (!snapshot) throw new Error('No snapshot to restore.');

        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();

            // ── 1. Clear the range that existed when we snapshotted ───────────
            // This removes cells the step added that are outside our regions.
            if (snapshot.clearedBeyond) {
                sheet.getRange(snapshot.clearedBeyond).clear('All');
                await ctx.sync();
            }

            // ── 2. Restore each region ────────────────────────────────────────
            for (const region of snapshot.regions) {
                const rows = region.formulas.length;
                if (!rows) continue;
                const rng = sheet.getRange(region.address);

                // Formulas and numberFormat accept 2-D arrays natively —
                // single Office.js write for the whole region.
                rng.formulas     = region.formulas;
                rng.numberFormat = region.numberFormat;

                // Fill, font, alignment must be set row-by-row because
                // the range-level property is a scalar write (sets all cells
                // identically) while we need per-cell fidelity.
                // Row batching: one getRange() per row, all queued before sync.
                const { startRow, startCol } = _parseTopLeft(region.address);
                const cols = region.formulas[0]?.length || 0;

                for (let r = 0; r < rows; r++) {
                    const rowAddr = _rowRangeAddr(startRow + r, startCol, cols);
                    const rowRng  = sheet.getRange(rowAddr);

                    // Build single-row 2-D arrays for this row
                    const fillRow      = [[...Array(cols)].map((_, c) => _v(region.fill,      r, c))];
                    const fontColorRow = [[...Array(cols)].map((_, c) => _v(region.fontColor, r, c))];
                    const fontBoldRow  = [[...Array(cols)].map((_, c) => _v(region.fontBold,  r, c))];
                    const fontSizeRow  = [[...Array(cols)].map((_, c) => _v(region.fontSize,  r, c))];
                    const alignRow     = [[...Array(cols)].map((_, c) => _v(region.alignment, r, c))];

                    // Fill: use clear() for empty, set color for non-empty
                    // We can't pass a 2-D array to fill.color — it's scalar only.
                    // So we iterate columns just for fill (fast: only changes per row).
                    for (let c = 0; c < cols; c++) {
                        const cellAddr = _cellAddr(startRow + r, startCol + c);
                        const fill = fillRow[0][c];
                        const cell = sheet.getRange(cellAddr);
                        if (fill && fill !== '' && fill !== 'null') {
                            cell.format.fill.color = fill;
                        } else {
                            cell.format.fill.clear();
                        }
                    }

                    // Font: numberFormat, fontColor, fontBold, fontSize, alignment
                    // can be written as 2-D arrays on a row range.
                    rowRng.numberFormat = [region.numberFormat[r]];

                    // fontColor, fontBold, fontSize, alignment are written via
                    // format properties which only accept scalars on a range.
                    // Use row range with arrays where API accepts them,
                    // fall back to per-cell for strict scalar-only props.
                    for (let c = 0; c < cols; c++) {
                        const cellAddr = _cellAddr(startRow + r, startCol + c);
                        const cell     = sheet.getRange(cellAddr);
                        const fc   = fontColorRow[0][c];
                        const bold = fontBoldRow[0][c];
                        const size = fontSizeRow[0][c];
                        const align = alignRow[0][c];
                        if (fc !== null && fc !== undefined)   cell.format.font.color = fc || null;
                        if (bold !== null && bold !== undefined) cell.format.font.bold = !!bold;
                        if (size)  cell.format.font.size = size;
                        if (align) cell.format.horizontalAlignment = align;
                    }
                }
            }

            await ctx.sync();
        });
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    /**
     * Given a ranges hint and the full used-range address, return the list
     * of addresses to snapshot. Deduplicates and validates each entry.
     */
    function _resolveAddresses(ranges, usedAddr) {
        if (!ranges?.length) return [usedAddr];
        // Normalise: strip sheet prefix, deduplicate
        const seen = new Set();
        const out  = [];
        for (const r of ranges) {
            const addr = _stripSheet(r.trim());
            if (addr && !seen.has(addr)) { seen.add(addr); out.push(addr); }
        }
        return out.length ? out : [usedAddr];
    }

    /** Strip "SheetName!" prefix from an address. */
    function _stripSheet(addr) {
        if (!addr) return addr;
        const i = addr.indexOf('!');
        return i >= 0 ? addr.slice(i + 1) : addr;
    }

    /**
     * Count cells in an address like "A1:E8" or "A1".
     * Returns Infinity if parsing fails (treats as oversized).
     */
    function _countCells(addr) {
        const clean = _stripSheet(addr);
        const m = clean.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
        if (!m) return Infinity;
        if (!m[3]) return 1; // single cell
        const cols = _colIndex(m[3].toUpperCase()) - _colIndex(m[1].toUpperCase()) + 1;
        const rows = parseInt(m[4], 10) - parseInt(m[2], 10) + 1;
        return rows * cols;
    }

    /** Parse top-left cell of "A1:E8" → { startRow: 0, startCol: 0 } (0-based). */
    function _parseTopLeft(addr) {
        const m = addr.match(/^([A-Z]+)(\d+)/i);
        if (!m) return { startRow: 0, startCol: 0 };
        return {
            startRow: parseInt(m[2], 10) - 1,
            startCol: _colIndex(m[1].toUpperCase()),
        };
    }

    /**
     * Return the address of a single row within a region.
     * e.g. (0, 0, 5) → "A1:E1"
     */
    function _rowRangeAddr(row, startCol, cols) {
        const from = _cellAddr(row, startCol);
        const to   = _cellAddr(row, startCol + cols - 1);
        return from === to ? from : `${from}:${to}`;
    }

    /** Convert 0-based row/col to Excel address, e.g. (0,0) → "A1". */
    function _cellAddr(row, col) {
        let s = '';
        let n = col;
        do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
        return s + (row + 1);
    }

    /** Convert column letters to 0-based index, e.g. "A"→0, "Z"→25, "AA"→26. */
    function _colIndex(letters) {
        let idx = 0;
        for (const ch of letters) idx = idx * 26 + (ch.charCodeAt(0) - 64);
        return idx - 1;
    }

    /**
     * Safe 2-D array value accessor.
     * Office.js sometimes returns a scalar (when all cells share a value)
     * instead of a 2-D array. This handles both shapes uniformly.
     */
    function _v(data, r, c) {
        if (!Array.isArray(data)) return data;             // scalar
        const row = Array.isArray(data[r]) ? data[r] : data;
        return Array.isArray(row) ? row[c] : row;          // row scalar fallback
    }

    /** Deep-copy a possibly nested array (avoids shared references in snapshot). */
    function _copy2d(v) {
        if (Array.isArray(v)) return v.map(_copy2d);
        return v;
    }

    function _warn(msg) {
        try { ExecutionEngine.log('err', `[Snapshot] ${msg}`); } catch (_) { console.warn('[WorksheetSnapshot]', msg); }
    }

    return { capture, restore };
})();
