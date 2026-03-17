/**
 * worksheetSnapshot.js
 *
 * Captures the full used-range state before each forward step and restores
 * it verbatim on backward navigation.
 *
 * Snapshot shape:
 * {
 *   address:      string,      // used-range address at capture time, sheet prefix stripped
 *   rows:         number,
 *   cols:         number,
 *   formulas:     string[][],  // always 2D
 *   values:       any[][],     // always 2D — used for non-formula cells
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

    const VALID_ALIGN = ['Left','Center','Right','Fill','Justify',
                         'CenterAcrossSelection','Distributed','General'];

    // ── capture() ─────────────────────────────────────────────────────────────

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

                used.load([
                    'formulas', 'values', 'numberFormat',
                    'format/fill/color',
                    'format/font/color',
                    'format/font/bold',
                    'format/font/size',
                    'format/horizontalAlignment',
                ]);
                await ctx.sync();

                snapshot = {
                    address:      _stripSheet(used.address),
                    rows,
                    cols,
                    formulas:     _copy2d(used.formulas),
                    values:       _copy2d(used.values),
                    numberFormat: _copy2d(used.numberFormat),
                    fill:         _expand(used.format.fill.color,          rows, cols),
                    fontColor:    _expand(used.format.font.color,          rows, cols),
                    fontBold:     _expand(used.format.font.bold,           rows, cols),
                    fontSize:     _expand(used.format.font.size,           rows, cols),
                    alignment:    _expand(used.format.horizontalAlignment, rows, cols),
                };
            });
        } catch (err) {
            _warn(`capture failed: ${err.message}`);
        }
        return snapshot;
    }

    // ── restore() ─────────────────────────────────────────────────────────────

    async function restore(snapshot) {
        if (!snapshot) throw new Error('No snapshot to restore.');

        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();

            // ── 1. Clear everything currently on the sheet ────────────────────
            // Wipes cells added by steps that ran after this snapshot was taken.
            const currentUsed = sheet.getUsedRangeOrNullObject();
            currentUsed.load('isNullObject');
            await ctx.sync();
            if (!currentUsed.isNullObject) currentUsed.clear('All');
            sheet.getRange(snapshot.address).clear('All');
            await ctx.sync();

            const { rows, cols } = snapshot;
            if (!rows || !cols) return;

            const { startRow, startCol } = _parseTopLeft(snapshot.address);

            // ── 2. Restore content + formatting per-cell ──────────────────────
            // Per-cell writes instead of batch 2D array writes because:
            //   - rng.formulas rejects empty strings "" → throws for whole array
            //   - Per-cell writes safely skip empty cells
            // All property sets are queued before the single ctx.sync below.
            //
            // _set() wraps each write to produce a detailed error message that
            // includes the cell address, property name, and exact value rejected.

            const _set = (cellAddr, prop, val, writeFn) => {
                try {
                    writeFn();
                } catch (e) {
                    const display = val === null      ? 'null'
                                  : val === undefined ? 'undefined'
                                  : `${typeof val} ${JSON.stringify(String(val)).slice(0, 60)}`;
                    throw new Error(
                        `[Snapshot restore] cell ${cellAddr}, prop "${prop}", ` +
                        `value ${display} — ${e.message}`
                    );
                }
            };

            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const addr    = _cellAddr(startRow + r, startCol + c);
                    const cell    = sheet.getRange(addr);
                    const formula = snapshot.formulas[r][c];
                    const value   = snapshot.values[r][c];
                    const nf      = snapshot.numberFormat[r][c];
                    const fill    = snapshot.fill[r][c];
                    const fc      = snapshot.fontColor[r][c];
                    const bold    = snapshot.fontBold[r][c];
                    const size    = snapshot.fontSize[r][c];
                    const align   = snapshot.alignment[r][c];

                    // Content
                    if (typeof formula === 'string' && formula.startsWith('=')) {
                        _set(addr, 'formulas', formula,
                            () => { cell.formulas = [[formula]]; });
                    } else if (value !== null && value !== undefined && value !== '') {
                        _set(addr, 'values', value,
                            () => { cell.values = [[value]]; });
                    }

                    // Number format
                    const nfSafe = (nf === null || nf === undefined || nf === '')
                                    ? 'General' : String(nf);
                    _set(addr, 'numberFormat', nfSafe,
                        () => { cell.numberFormat = [[nfSafe]]; });

                    // Fill
                    if (fill && typeof fill === 'string' && fill !== 'null') {
                        _set(addr, 'fill.color', fill,
                            () => { cell.format.fill.color = fill; });
                    } else {
                        cell.format.fill.clear();
                    }

                    // Font color
                    const fcSafe = (fc && typeof fc === 'string' && fc !== 'null') ? fc : null;
                    _set(addr, 'font.color', fcSafe,
                        () => { cell.format.font.color = fcSafe; });

                    // Font bold
                    _set(addr, 'font.bold', bold,
                        () => { cell.format.font.bold = !!bold; });

                    // Font size
                    if (size) {
                        _set(addr, 'font.size', size,
                            () => { cell.format.font.size = size; });
                    }

                    // Alignment
                    const alignSafe = (align && VALID_ALIGN.includes(align))
                                       ? align : 'General';
                    _set(addr, 'horizontalAlignment', alignSafe,
                        () => { cell.format.horizontalAlignment = alignSafe; });
                }
            }

            await ctx.sync();
        });
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    /**
     * Expand a range-level format property into a guaranteed rows×cols 2D array.
     * Office.js returns a scalar when all cells share the same value.
     */
    function _expand(value, rows, cols) {
        if (Array.isArray(value) && Array.isArray(value[0])) {
            return value.map(row => [...row]);
        }
        if (Array.isArray(value)) {
            return Array.from({ length: rows }, () =>
                Array.from({ length: cols }, (_, c) => value[c] ?? null));
        }
        return Array.from({ length: rows }, () =>
            Array.from({ length: cols }, () => value));
    }

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
