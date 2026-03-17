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
                    fill:         _expandColor(used.format.fill.color,     rows, cols),
                    fontColor:    _expandColor(used.format.font.color,     rows, cols),
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

            // ── 2. Restore in phases, each followed by ctx.sync() ────────────
            // Phased syncs isolate which property group Office.js rejects.
            // The error message from each phase names the phase so we can
            // pinpoint the exact problem without guessing.

            const _phase = async (name, writeFn) => {
                try {
                    writeFn();
                    await ctx.sync();
                } catch (e) {
                    throw new Error(`[Snapshot restore: ${name}] ${e.message}`);
                }
            };

            // Phase A: formulas / values
            await _phase('content', () => {
                for (let r = 0; r < rows; r++) {
                    for (let c = 0; c < cols; c++) {
                        const cell    = sheet.getRange(_cellAddr(startRow + r, startCol + c));
                        const formula = snapshot.formulas[r][c];
                        const value   = snapshot.values[r][c];
                        if (typeof formula === 'string' && formula.startsWith('=')) {
                            cell.formulas = [[formula]];
                        } else if (value !== null && value !== undefined && value !== '') {
                            cell.values = [[value]];
                        }
                    }
                }
            });

            // Phase B: numberFormat
            await _phase('numberFormat', () => {
                for (let r = 0; r < rows; r++) {
                    for (let c = 0; c < cols; c++) {
                        const cell = sheet.getRange(_cellAddr(startRow + r, startCol + c));
                        const nf   = snapshot.numberFormat[r][c];
                        const nfSafe = (nf === null || nf === undefined || nf === '')
                                        ? 'General' : String(nf);
                        cell.numberFormat = [[nfSafe]];
                    }
                }
            });

            // Phase C: fill
            // Values already sanitized to '#RRGGBB' or '' at capture time.
            // Use fill.clear() for '' to properly remove fill (setting color='' throws).
            await _phase('fill', () => {
                for (let r = 0; r < rows; r++) {
                    for (let c = 0; c < cols; c++) {
                        const cell = sheet.getRange(_cellAddr(startRow + r, startCol + c));
                        const fill = snapshot.fill[r][c];
                        if (fill) {
                            cell.format.fill.color = fill;
                        } else {
                            cell.format.fill.clear();
                        }
                    }
                }
            });

            // Phase D: font color
            // Values already sanitized to '#RRGGBB' or '' at capture time.
            // Write '' (not null) to reset to theme/automatic color.
            await _phase('font.color', () => {
                for (let r = 0; r < rows; r++) {
                    for (let c = 0; c < cols; c++) {
                        const cell = sheet.getRange(_cellAddr(startRow + r, startCol + c));
                        cell.format.font.color = snapshot.fontColor[r][c] || '';
                    }
                }
            });

            // Phase E: font bold + size
            await _phase('font.bold+size', () => {
                for (let r = 0; r < rows; r++) {
                    for (let c = 0; c < cols; c++) {
                        const cell = sheet.getRange(_cellAddr(startRow + r, startCol + c));
                        cell.format.font.bold = !!snapshot.fontBold[r][c];
                        const size = snapshot.fontSize[r][c];
                        if (size) cell.format.font.size = size;
                    }
                }
            });

            // Phase F: alignment
            await _phase('alignment', () => {
                for (let r = 0; r < rows; r++) {
                    for (let c = 0; c < cols; c++) {
                        const cell      = sheet.getRange(_cellAddr(startRow + r, startCol + c));
                        const align     = snapshot.alignment[r][c];
                        const alignSafe = (align && VALID_ALIGN.includes(align))
                                           ? align : 'General';
                        cell.format.horizontalAlignment = alignSafe;
                    }
                }
            });
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

    /** Like _expand but also sanitizes every value through _sanitizeColor. */
    function _expandColor(value, rows, cols) {
        const grid = _expand(value, rows, cols);
        return grid.map(row => row.map(v => _sanitizeColor(v)));
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

    /**
     * Sanitize a color value read from Office.js into a form it will accept back.
     *
     * Office.js reads font.color / fill.color and may return:
     *   "#RRGGBB"   — valid, write as-is
     *   "FFRRGGBB"  — 8-char ARGB without #; extract last 6 chars → "#RRGGBB"
     *   null        — theme/automatic color; write "" to reset
     *   ""          — already reset; write "" to reset
     *   anything else — unknown; write "" to reset
     *
     * @param {string|null} color
     * @param {string} fallback — "" to reset, or a default hex like "#000000"
     */
    function _sanitizeColor(color, fallback = '') {
        if (!color || typeof color !== 'string') return fallback;
        // Already valid #RRGGBB
        if (/^#[0-9A-Fa-f]{6}$/.test(color)) return color;
        // 8-char ARGB "AARRGGBB" — drop the alpha channel prefix
        if (/^[0-9A-Fa-f]{8}$/.test(color)) return '#' + color.slice(2);
        // 6-char RGB without # — add it
        if (/^[0-9A-Fa-f]{6}$/.test(color)) return '#' + color;
        return fallback;
    }

    function _warn(msg) {
        try { ExecutionEngine.log('err', `[Snapshot] ${msg}`); }
        catch (_) { console.warn('[WorksheetSnapshot]', msg); }
    }

    return { capture, restore };
})();
