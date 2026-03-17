/**
 * worksheetSnapshot.js
 *
 * Captures the full used-range state before each forward step and restores
 * it verbatim on backward navigation.
 *
 * Two bugs fixed vs previous version:
 *
 *  Bug 1 — address scoping: restore() was only clearing/writing the snapshot's
 *  address (e.g. "A1:E7"). If a later step added rows outside that region,
 *  going back left those rows intact. Fix: restore() clears the ENTIRE current
 *  used range before writing the snapshot back, so the sheet is fully reset.
 *
 *  Bug 2 — scalar format properties: Office.js returns a scalar (e.g. null,
 *  "#000000") when all cells in a range share the same format value, instead
 *  of a 2D array. Storing that scalar and applying it back to every cell
 *  overwrites real per-cell formatting from earlier steps. Fix: after loading
 *  range-level format properties, expand any scalar into a proper rows×cols 2D
 *  array during capture so restore() always has exact per-cell values.
 *
 * Snapshot shape:
 * {
 *   address:      string,      // used-range address at capture time, sheet prefix stripped
 *   rows:         number,
 *   cols:         number,
 *   formulas:     string[][],  // always 2D
 *   numberFormat: string[][],  // always 2D
 *   fill:         string[][],  // always 2D, null = no fill
 *   fontColor:    string[][],  // always 2D, null = default
 *   fontBold:     boolean[][], // always 2D
 *   fontSize:     number[][],  // always 2D
 *   alignment:    string[][],  // always 2D
 * }
 *
 * MAX_CELLS = 5000. Exceeded → warning logged, snapshot skipped for that step.
 */
const WorksheetSnapshot = (() => {

    const MAX_CELLS = 5000;

    // ── Public API ────────────────────────────────────────────────────────────

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

                // Load formulas + values + numberFormat
                // We need values as fallback because the batch formulas write
                // rejects empty strings — we restore cell-by-cell instead.
                used.load(['formulas', 'values', 'numberFormat']);

                // Load format properties — Office.js MAY return scalars when
                // all cells share the same value. We load them here then expand below.
                used.load([
                    'format/fill/color',
                    'format/font/color',
                    'format/font/bold',
                    'format/font/size',
                    'format/horizontalAlignment',
                ]);
                await ctx.sync();

                const addr = _stripSheet(used.address);

                // Expand any scalar format values into proper rows×cols 2D arrays.
                // This is the critical fix: a scalar null/string means "all cells
                // share this value" — we must record it per-cell so restore() can
                // write the exact right value back to each cell individually.
                snapshot = {
                    address:      addr,
                    rows,
                    cols,
                    formulas:     _copy2d(used.formulas),
                    values:       _copy2d(used.values),
                    numberFormat: _copy2d(used.numberFormat),
                    fill:         _expand(used.format.fill.color,              rows, cols),
                    fontColor:    _expand(used.format.font.color,              rows, cols),
                    fontBold:     _expand(used.format.font.bold,               rows, cols),
                    fontSize:     _expand(used.format.font.size,               rows, cols),
                    alignment:    _expand(used.format.horizontalAlignment,     rows, cols),
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
     *   1. Clear the ENTIRE current used range (not just the snapshot region) so
     *      any cells added by steps after this snapshot are wiped first.
     *   2. Also clear the snapshot region explicitly in case the current used range
     *      is smaller (e.g. user deleted data).
     *   3. Write formulas + numberFormat as 2D arrays (fast, one call each).
     *   4. Write fill / font / alignment per-cell (all queued before one ctx.sync).
     */
    async function restore(snapshot) {
        if (!snapshot) throw new Error('No snapshot to restore.');

        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();

            // ── 1. Clear everything that currently exists ─────────────────────
            // This removes cells that were added by steps occurring AFTER the
            // snapshot was taken, which lie outside snapshot.address.
            const currentUsed = sheet.getUsedRangeOrNullObject();
            currentUsed.load('isNullObject');
            await ctx.sync();
            if (!currentUsed.isNullObject) {
                currentUsed.clear('All');
            }
            // Also clear the snapshot region itself (covers case where current
            // used range shrank and misses part of what we need to restore).
            sheet.getRange(snapshot.address).clear('All');
            await ctx.sync();

            const { rows, cols } = snapshot;
            if (!rows || !cols) return;

            const { startRow, startCol } = _parseTopLeft(snapshot.address);

            // ── 2 & 3. Restore content + formatting per-cell ─────────────────
            // We restore formulas/values cell-by-cell rather than as a batch
            // 2D array write because Office.js rejects empty strings ("") in
            // the formulas array and throws for the entire write. Per-cell
            // writes skip empty cells cleanly. All writes are queued before
            // a single ctx.sync so there is still only one round-trip.

            const VALID_ALIGN = ['Left','Center','Right','Fill','Justify','CenterAcrossSelection','Distributed','General'];

            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const cell = sheet.getRange(_cellAddr(startRow + r, startCol + c));

                    // Content: write formula if starts with '=', raw value otherwise.
                    // Empty cells are already cleared — writing '' to cell.formulas throws.
                    const formula = snapshot.formulas[r][c];
                    const value   = snapshot.values[r][c];
                    if (typeof formula === 'string' && formula.startsWith('=')) {
                        cell.formulas = [[formula]];
                    } else if (value !== null && value !== undefined && value !== '') {
                        cell.values = [[value]];
                    }

                    // Number format — null/undefined/''/0 → 'General'
                    const nf = snapshot.numberFormat[r][c];
                    cell.numberFormat = [[(nf === null || nf === undefined || nf === '') ? 'General' : String(nf)]];

                    // Fill — clear() for no-fill; only set a valid hex/named string
                    const fill = snapshot.fill[r][c];
                    if (fill && typeof fill === 'string' && fill !== 'null') {
                        cell.format.fill.color = fill;
                    } else {
                        cell.format.fill.clear();
                    }

                    // Font color — null/falsy resets to automatic; string "null" also treated as reset
                    const fc = snapshot.fontColor[r][c];
                    cell.format.font.color = (fc && typeof fc === 'string' && fc !== 'null') ? fc : null;

                    // Font bold
                    cell.format.font.bold = !!snapshot.fontBold[r][c];

                    // Font size — only set if non-zero (0 means "default")
                    const size = snapshot.fontSize[r][c];
                    if (size) cell.format.font.size = size;

                    // Alignment — only valid enum strings accepted, '' throws
                    const align = snapshot.alignment[r][c];
                    cell.format.horizontalAlignment =
                        (align && VALID_ALIGN.includes(align)) ? align : 'General';
                }
            }

            await ctx.sync();
        });
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    /**
     * Expand a range-level format property into a guaranteed rows×cols 2D array.
     *
     * Office.js returns:
     *   - A proper 2D array when cells differ (the common case after styling)
     *   - A scalar when ALL cells share the same value (e.g. null fill everywhere,
     *     or bold=false everywhere)
     *
     * We always want a 2D array so restore() can index [r][c] unconditionally.
     */
    function _expand(value, rows, cols) {
        // Already a proper 2D array — deep-copy it
        if (Array.isArray(value) && Array.isArray(value[0])) {
            return value.map(row => [...row]);
        }

        // 1D array (shouldn't happen for these properties, but handle it)
        if (Array.isArray(value)) {
            return Array.from({ length: rows }, (_, r) =>
                Array.from({ length: cols }, (_, c) => value[c] ?? null)
            );
        }

        // Scalar — every cell shares this value
        return Array.from({ length: rows }, () =>
            Array.from({ length: cols }, () => value)
        );
    }

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

    function _copy2d(v) {
        if (Array.isArray(v)) return v.map(r => Array.isArray(r) ? [...r] : r);
        return v;
    }

    function _warn(msg) {
        try { ExecutionEngine.log('err', `[Snapshot] ${msg}`); } catch (_) { console.warn('[WorksheetSnapshot]', msg); }
    }

    return { capture, restore };
})();
