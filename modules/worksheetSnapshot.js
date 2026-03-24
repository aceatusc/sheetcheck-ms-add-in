/**
 * worksheetSnapshot.js
 *
 * Captures and restores the full visible state of the active worksheet
 * before/after each execution step, enabling robust undo on backward nav.
 *
 * What is captured
 * ────────────────
 *   • Content     — formulas + computed values
 *   • Formatting  — fill, font, number format, alignment, wrap, indent, borders
 *   • Dimensions  — column widths, row heights
 *   • Merges      — merged-cell ranges (must be unmerged before clear)
 *   • Charts      — type, title, data range, position, size
 *   • Validation  — data-validation rules per cell
 *   • Cond. fmt   — captured for awareness but NOT restored (Office.js write
 *                   API for conditional formatting is unavailable on Desktop);
 *                   a warning is emitted if rules were present.
 *
 * Office.js constraints handled
 * ──────────────────────────────
 *   • Borders must be written one BorderIndex at a time — no bulk assignment.
 *   • Merged ranges must be un-merged before clear() or Desktop throws.
 *   • Chart title text requires a second sync after chart.title.load().
 *   • getRowProperties / getColumnProperties are post-2019 API — we fall back
 *     to per-row/col iteration when unavailable.
 *   • Every section of capture and restore is wrapped in its own try/catch
 *     so a failure in one section never prevents the others from running.
 *
 * MAX_CELLS = 5000 (content + style both respect this cap).
 */
const WorksheetSnapshot = (() => {

    const MAX_CELLS = 5000;

    // Border sides we capture/restore.  DiagonalDown/Up are intentionally
    // omitted — they are rarely used and restoration is unreliable on Desktop.
    const BORDER_SIDES = [
        'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'EdgeTop',
        'InsideHorizontal', 'InsideVertical',
    ];

    // ─────────────────────────────────────────────────────────────────────────
    // capture
    // ─────────────────────────────────────────────────────────────────────────

    async function capture() {
        let snapshot = null;

        try {
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();

                // ── Determine used range ──────────────────────────────────────
                const used = sheet.getUsedRangeOrNullObject();
                used.load(['isNullObject', 'rowCount', 'columnCount', 'address']);
                await ctx.sync();

                if (used.isNullObject) {
                    // Sheet is blank — still capture charts/validation if any.
                    snapshot = _emptySnapshot();
                    await _captureCharts(ctx, sheet, snapshot);
                    await _captureMerges(ctx, sheet, snapshot);
                    return;
                }

                const rows  = used.rowCount;
                const cols  = used.columnCount;
                const cells = rows * cols;
                const addr  = _stripSheet(used.address);

                if (cells > MAX_CELLS) {
                    _warn(`Used range ${rows}×${cols} = ${cells} cells exceeds cap (${MAX_CELLS}). Snapshot skipped.`);
                    return;
                }

                snapshot = _emptySnapshot();
                snapshot.address = addr;
                snapshot.rows    = rows;
                snapshot.cols    = cols;

                // ── Content (formulas + values) ───────────────────────────────
                try {
                    used.load(['formulas', 'values']);
                    await ctx.sync();
                    snapshot.formulas = _copy2d(used.formulas);
                    snapshot.values   = _copy2d(used.values);
                } catch (e) {
                    _warn(`content capture failed: ${e.message}`);
                }

                // ── Cell formatting ───────────────────────────────────────────
                await _captureFormatting(ctx, used, snapshot, rows, cols);

                // ── Row heights + column widths ───────────────────────────────
                await _captureDimensions(ctx, sheet, snapshot, rows, cols);

                // ── Merged cells ──────────────────────────────────────────────
                await _captureMerges(ctx, sheet, snapshot);

                // ── Charts ────────────────────────────────────────────────────
                await _captureCharts(ctx, sheet, snapshot);

                // ── Data validation ───────────────────────────────────────────
                await _captureValidation(ctx, used, snapshot, rows, cols);

                // ── Conditional formatting (read-only — capture count only) ───
                try {
                    const cf = used.conditionalFormats;
                    cf.load('items/type');
                    await ctx.sync();
                    snapshot.conditionalFormatCount = cf.items.length;
                    if (cf.items.length) {
                        _warn(`Sheet has ${cf.items.length} conditional format rule(s). ` +
                              `These cannot be restored by snapshot (Office.js limitation).`);
                    }
                } catch (_) { /* non-fatal */ }
            });
        } catch (err) {
            _warn(`capture failed: ${err.message}`);
            snapshot = null;
        }

        return snapshot;
    }

    // ── Formatting capture ────────────────────────────────────────────────────

    async function _captureFormatting(ctx, used, snapshot, rows, cols) {
        try {
            // Load all scalar format properties in one batch.
            used.load([
                'format/fill/color',
                'format/font/color',
                'format/font/bold',
                'format/font/italic',
                'format/font/size',
                'format/font/name',
                'format/font/strikethrough',
                'format/font/underline',
                'format/horizontalAlignment',
                'format/verticalAlignment',
                'format/wrapText',
                'format/indentLevel',
                'format/numberFormat',
            ]);
            await ctx.sync();

            snapshot.format = {
                fill:                { color: used.format.fill.color || null },
                font: {
                    color:         used.format.font.color         || null,
                    bold:          used.format.font.bold,
                    italic:        used.format.font.italic,
                    size:          used.format.font.size          || null,
                    name:          used.format.font.name          || null,
                    strikethrough: used.format.font.strikethrough,
                    underline:     used.format.font.underline     || null,
                },
                horizontalAlignment: used.format.horizontalAlignment || null,
                verticalAlignment:   used.format.verticalAlignment   || null,
                wrapText:            used.format.wrapText,
                indentLevel:         used.format.indentLevel         || 0,
                numberFormat:        used.format.numberFormat        || null,
            };
        } catch (e) {
            _warn(`format capture (bulk) failed: ${e.message}`);
        }

        // Per-cell format snapshot — needed for cells that deviate from the
        // range-level defaults (e.g. a single bold header in an otherwise
        // plain range).  We do all loads then one sync.
        try {
            const cellFormatData = [];
            const cellRefs       = [];

            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const cell = used.getCell(r, c);
                    cell.load([
                        'address',
                        'format/fill/color',
                        'format/font/color',
                        'format/font/bold',
                        'format/font/italic',
                        'format/font/size',
                        'format/font/name',
                        'format/font/strikethrough',
                        'format/font/underline',
                        'format/horizontalAlignment',
                        'format/verticalAlignment',
                        'format/wrapText',
                        'format/indentLevel',
                        'format/numberFormat',
                    ]);
                    cellRefs.push({ cell, r, c });
                }
            }
            await ctx.sync();

            for (const { cell, r, c } of cellRefs) {
                cellFormatData.push({
                    r, c,
                    fill:                { color: cell.format.fill.color || null },
                    font: {
                        color:         cell.format.font.color         || null,
                        bold:          cell.format.font.bold,
                        italic:        cell.format.font.italic,
                        size:          cell.format.font.size          || null,
                        name:          cell.format.font.name          || null,
                        strikethrough: cell.format.font.strikethrough,
                        underline:     cell.format.font.underline     || null,
                    },
                    horizontalAlignment: cell.format.horizontalAlignment || null,
                    verticalAlignment:   cell.format.verticalAlignment   || null,
                    wrapText:            cell.format.wrapText,
                    indentLevel:         cell.format.indentLevel         || 0,
                    numberFormat:        cell.format.numberFormat        || null,
                });
            }
            snapshot.cellFormats = cellFormatData;
        } catch (e) {
            _warn(`per-cell format capture failed: ${e.message}`);
        }

        // Borders — per-cell, per-side.  Load all at once then sync once.
        try {
            const borderRefs  = [];
            const borderData  = [];

            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const cell = used.getCell(r, c);
                    const sideData = {};
                    for (const side of BORDER_SIDES) {
                        const b = cell.format.borders.getItem(side);
                        b.load(['style', 'color', 'weight']);
                        sideData[side] = b;
                    }
                    borderRefs.push({ r, c, sideData });
                }
            }
            await ctx.sync();

            for (const { r, c, sideData } of borderRefs) {
                const sides = {};
                for (const side of BORDER_SIDES) {
                    sides[side] = {
                        style:  sideData[side].style  || 'None',
                        color:  sideData[side].color  || null,
                        weight: sideData[side].weight || null,
                    };
                }
                borderData.push({ r, c, sides });
            }
            snapshot.borders = borderData;
        } catch (e) {
            _warn(`border capture failed: ${e.message}`);
        }
    }

    // ── Dimension capture (row heights + col widths) ──────────────────────────

    async function _captureDimensions(ctx, sheet, snapshot, rows, cols) {
        // Try the modern getRowProperties / getColumnProperties API first
        // (ExcelApi 1.9+).  Falls back to per-index iteration on older hosts.
        try {
            const rowProps = sheet.getRowProperties(0, rows, { rowHeight: true });
            await ctx.sync();
            snapshot.rowHeights = rowProps.value.map(p => p.rowHeight ?? null);
        } catch (_) {
            // Fallback: load each row via getRangeByIndexes
            try {
                const refs = [];
                for (let r = 0; r < rows; r++) {
                    const row = sheet.getRangeByIndexes(r, 0, 1, 1).getEntireRow();
                    row.load('height');
                    refs.push(row);
                }
                await ctx.sync();
                snapshot.rowHeights = refs.map(r => r.height ?? null);
            } catch (e2) {
                _warn(`row height capture failed: ${e2.message}`);
            }
        }

        try {
            const colProps = sheet.getColumnProperties(0, cols, { columnWidth: true });
            await ctx.sync();
            snapshot.colWidths = colProps.value.map(p => p.columnWidth ?? null);
        } catch (_) {
            try {
                const refs = [];
                for (let c = 0; c < cols; c++) {
                    const col = sheet.getRangeByIndexes(0, c, 1, 1).getEntireColumn();
                    col.load('width');
                    refs.push(col);
                }
                await ctx.sync();
                snapshot.colWidths = refs.map(c => c.width ?? null);
            } catch (e2) {
                _warn(`col width capture failed: ${e2.message}`);
            }
        }
    }

    // ── Merged cells capture ──────────────────────────────────────────────────

    async function _captureMerges(ctx, sheet, snapshot) {
        try {
            // getMergedAreas isn't available everywhere; use the used range merge flags.
            // We inspect isMerged per cell and collect unique merge area addresses.
            if (!snapshot.rows || !snapshot.cols) return;

            const used = sheet.getRange(snapshot.address || 'A1');
            // Reload to get a fresh range reference if needed.
            const mergedAreas = used.getMergedAreasOrNullObject
                ? used.getMergedAreasOrNullObject()
                : null;

            if (mergedAreas) {
                mergedAreas.load('isNullObject');
                await ctx.sync();
                if (!mergedAreas.isNullObject) {
                    mergedAreas.load('areas/address');
                    await ctx.sync();
                    snapshot.merges = mergedAreas.areas
                        ? mergedAreas.areas.items.map(a => _stripSheet(a.address))
                        : [];
                    return;
                }
            }

            // Fallback: scan cells for isMerged flag and collect unique merge areas.
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load('isNullObject');
            await ctx.sync();
            if (usedRange.isNullObject) return;

            usedRange.load('mergeArea/address');
            await ctx.sync();

            // mergeArea on the whole range gives us the merged region if the whole
            // range is one merge; for individual cells we need per-cell approach.
            const cellRefs = [];
            for (let r = 0; r < snapshot.rows; r++) {
                for (let c = 0; c < snapshot.cols; c++) {
                    const cell = usedRange.getCell(r, c);
                    cell.load(['isMerged', 'mergeArea/address']);
                    cellRefs.push(cell);
                }
            }
            await ctx.sync();

            const seen    = new Set();
            const merges  = [];
            for (const cell of cellRefs) {
                if (cell.isMerged) {
                    const ma = _stripSheet(cell.mergeArea.address);
                    if (!seen.has(ma)) { seen.add(ma); merges.push(ma); }
                }
            }
            snapshot.merges = merges;
        } catch (e) {
            _warn(`merge capture failed: ${e.message}`);
        }
    }

    // ── Chart capture ─────────────────────────────────────────────────────────

    async function _captureCharts(ctx, sheet, snapshot) {
        try {
            const charts = sheet.charts;
            charts.load('items/name,items/chartType,items/left,items/top,items/width,items/height');
            await ctx.sync();

            const chartData = [];
            for (const chart of charts.items) {
                const entry = {
                    name:      chart.name      || null,
                    chartType: chart.chartType || null,
                    left:      chart.left,
                    top:       chart.top,
                    width:     chart.width,
                    height:    chart.height,
                    title:     null,
                    dataRange: null,
                };

                // Title — requires separate load
                try {
                    if (chart.title) {
                        chart.title.load('text');
                        await ctx.sync();
                        entry.title = chart.title.text || null;
                    }
                } catch (_) { /* chart may have no title */ }

                // Data range — throws for some chart types
                try {
                    const dr = chart.getDataRange();
                    dr.load('address');
                    await ctx.sync();
                    entry.dataRange = _stripSheet(dr.address);
                } catch (_) { /* some chart types don't expose a data range */ }

                chartData.push(entry);
            }
            snapshot.charts = chartData;
        } catch (e) {
            _warn(`chart capture failed: ${e.message}`);
        }
    }

    // ── Data validation capture ───────────────────────────────────────────────

    async function _captureValidation(ctx, used, snapshot, rows, cols) {
        try {
            const validationData = [];
            const cellRefs       = [];

            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const cell = used.getCell(r, c);
                    cell.dataValidation.load([
                        'rule',
                        'ignoreBlanks',
                        'showErrorAlert',
                        'errorTitle',
                        'errorMessage',
                        'showInputMessage',
                        'promptTitle',
                        'prompt',
                    ]);
                    cellRefs.push({ cell, r, c });
                }
            }
            await ctx.sync();

            for (const { cell, r, c } of cellRefs) {
                const dv = cell.dataValidation;
                // rule is null when no validation is set
                if (dv.rule && Object.keys(dv.rule).length) {
                    validationData.push({
                        r, c,
                        rule:             JSON.parse(JSON.stringify(dv.rule)),
                        ignoreBlanks:     dv.ignoreBlanks,
                        showErrorAlert:   dv.showErrorAlert,
                        errorTitle:       dv.errorTitle    || null,
                        errorMessage:     dv.errorMessage  || null,
                        showInputMessage: dv.showInputMessage,
                        promptTitle:      dv.promptTitle   || null,
                        prompt:           dv.prompt        || null,
                    });
                }
            }
            snapshot.validation = validationData;
        } catch (e) {
            _warn(`validation capture failed: ${e.message}`);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // restore
    // ─────────────────────────────────────────────────────────────────────────

    async function restore(snapshot) {
        if (!snapshot) throw new Error('No snapshot to restore.');

        await Excel.run(async (ctx) => {
            const sheet = ctx.workbook.worksheets.getActiveWorksheet();

            // ── 1. Un-merge everything first (required before clear on Desktop) ─
            await _restoreUnmerge(ctx, sheet);

            // ── 2. Delete all existing charts ────────────────────────────────
            await _deleteAllCharts(ctx, sheet);

            // ── 3. Clear content + all formatting on used range ───────────────
            try {
                const currentUsed = sheet.getUsedRangeOrNullObject();
                currentUsed.load('isNullObject');
                await ctx.sync();
                if (!currentUsed.isNullObject) {
                    currentUsed.clear('All');   // clears content + formats
                }
                if (snapshot.address) {
                    sheet.getRange(snapshot.address).clear('All');
                }
                await ctx.sync();
            } catch (e) {
                _warn(`clear failed: ${e.message}`);
            }

            if (!snapshot.rows || !snapshot.cols) {
                // Sheet was blank — re-add charts and return
                await _restoreCharts(ctx, sheet, snapshot);
                return;
            }

            const { startRow, startCol } = _parseTopLeft(snapshot.address);

            // ── 4. Content ────────────────────────────────────────────────────
            await _restoreContent(ctx, sheet, snapshot, startRow, startCol);

            // ── 5. Per-cell formatting ────────────────────────────────────────
            await _restoreCellFormats(ctx, sheet, snapshot, startRow, startCol);

            // ── 6. Borders ────────────────────────────────────────────────────
            await _restoreBorders(ctx, sheet, snapshot, startRow, startCol);

            // ── 7. Row heights + col widths ───────────────────────────────────
            await _restoreDimensions(ctx, sheet, snapshot, startRow, startCol);

            // ── 8. Re-apply merges ────────────────────────────────────────────
            await _restoreMerges(ctx, sheet, snapshot);

            // ── 9. Re-add charts ──────────────────────────────────────────────
            await _restoreCharts(ctx, sheet, snapshot);

            // ── 10. Data validation ───────────────────────────────────────────
            await _restoreValidation(ctx, sheet, snapshot, startRow, startCol);
        });
    }

    // ── Un-merge all merged cells before clearing ─────────────────────────────

    async function _restoreUnmerge(ctx, sheet) {
        try {
            const usedForMerge = sheet.getUsedRangeOrNullObject();
            usedForMerge.load('isNullObject');
            await ctx.sync();
            if (!usedForMerge.isNullObject) {
                usedForMerge.unmerge();
                await ctx.sync();
            }
        } catch (e) {
            _warn(`un-merge failed (non-fatal): ${e.message}`);
        }
    }

    // ── Delete all charts from the sheet ─────────────────────────────────────

    async function _deleteAllCharts(ctx, sheet) {
        try {
            const charts = sheet.charts;
            charts.load('items/name');
            await ctx.sync();
            // Delete in reverse to avoid index shifting
            for (let i = charts.items.length - 1; i >= 0; i--) {
                charts.items[i].delete();
            }
            await ctx.sync();
        } catch (e) {
            _warn(`chart delete failed: ${e.message}`);
        }
    }

    // ── Content restore ───────────────────────────────────────────────────────

    async function _restoreContent(ctx, sheet, snapshot, startRow, startCol) {
        if (!snapshot.formulas?.length) return;
        try {
            const { rows, cols } = snapshot;
            for (let r = 0; r < rows; r++) {
                for (let c = 0; c < cols; c++) {
                    const formula = snapshot.formulas[r]?.[c];
                    const value   = snapshot.values?.[r]?.[c];
                    const addr    = _cellAddr(startRow + r, startCol + c);

                    if (typeof formula === 'string' && formula.startsWith('=')) {
                        sheet.getRange(addr).formulas = [[formula]];
                    } else if (value !== null && value !== undefined && value !== '') {
                        sheet.getRange(addr).values = [[value]];
                    }
                }
            }
            await ctx.sync();
        } catch (e) {
            _warn(`content restore failed: ${e.message}`);
        }
    }

    // ── Per-cell format restore ───────────────────────────────────────────────

    async function _restoreCellFormats(ctx, sheet, snapshot, startRow, startCol) {
        if (!snapshot.cellFormats?.length) return;
        try {
            for (const cf of snapshot.cellFormats) {
                const addr = _cellAddr(startRow + cf.r, startCol + cf.c);
                const cell = sheet.getRange(addr);

                if (cf.fill?.color)           cell.format.fill.color           = cf.fill.color;
                if (cf.font?.color)           cell.format.font.color           = cf.font.color;
                if (cf.font?.bold   != null)  cell.format.font.bold            = cf.font.bold;
                if (cf.font?.italic != null)  cell.format.font.italic          = cf.font.italic;
                if (cf.font?.size)            cell.format.font.size            = cf.font.size;
                if (cf.font?.name)            cell.format.font.name            = cf.font.name;
                if (cf.font?.strikethrough != null) cell.format.font.strikethrough = cf.font.strikethrough;
                if (cf.font?.underline)       cell.format.font.underline       = cf.font.underline;
                if (cf.horizontalAlignment)   cell.format.horizontalAlignment  = cf.horizontalAlignment;
                if (cf.verticalAlignment)     cell.format.verticalAlignment    = cf.verticalAlignment;
                if (cf.wrapText != null)      cell.format.wrapText             = cf.wrapText;
                if (cf.indentLevel)           cell.format.indentLevel          = cf.indentLevel;
                if (cf.numberFormat)          cell.format.numberFormat         = cf.numberFormat;
            }
            // One sync for all cells — much better on Desktop
            await ctx.sync();
        } catch (e) {
            _warn(`cell format restore failed: ${e.message}`);
        }
    }

    // ── Border restore ────────────────────────────────────────────────────────

    async function _restoreBorders(ctx, sheet, snapshot, startRow, startCol) {
        if (!snapshot.borders?.length) return;
        try {
            for (const bd of snapshot.borders) {
                const addr = _cellAddr(startRow + bd.r, startCol + bd.c);
                const cell = sheet.getRange(addr);

                for (const side of BORDER_SIDES) {
                    const s = bd.sides?.[side];
                    if (!s || s.style === 'None' || !s.style) continue;
                    try {
                        const b = cell.format.borders.getItem(side);
                        b.style = s.style;
                        if (s.color)  b.color  = s.color;
                        if (s.weight) b.weight = s.weight;
                    } catch (sideErr) {
                        // Individual border side can throw (e.g. InsideHorizontal
                        // on a single cell) — safe to skip.
                    }
                }
            }
            await ctx.sync();
        } catch (e) {
            _warn(`border restore failed: ${e.message}`);
        }
    }

    // ── Dimension restore (row heights + col widths) ──────────────────────────

    async function _restoreDimensions(ctx, sheet, snapshot, startRow, startCol) {
        // Row heights
        if (snapshot.rowHeights?.length) {
            try {
                // Try modern setRowProperties first
                const props = snapshot.rowHeights.map((h, i) => ({
                    index: startRow + i,
                    rowHeight: h,
                }));
                sheet.setRowProperties(props);
                await ctx.sync();
            } catch (_) {
                // Fallback: per-row
                try {
                    for (let i = 0; i < snapshot.rowHeights.length; i++) {
                        const h = snapshot.rowHeights[i];
                        if (h == null) continue;
                        sheet.getRangeByIndexes(startRow + i, 0, 1, 1)
                             .getEntireRow().format.rowHeight = h;
                    }
                    await ctx.sync();
                } catch (e2) {
                    _warn(`row height restore failed: ${e2.message}`);
                }
            }
        }

        // Column widths
        if (snapshot.colWidths?.length) {
            try {
                const props = snapshot.colWidths.map((w, i) => ({
                    index: startCol + i,
                    columnWidth: w,
                }));
                sheet.setColumnProperties(props);
                await ctx.sync();
            } catch (_) {
                try {
                    for (let i = 0; i < snapshot.colWidths.length; i++) {
                        const w = snapshot.colWidths[i];
                        if (w == null) continue;
                        sheet.getRangeByIndexes(0, startCol + i, 1, 1)
                             .getEntireColumn().format.columnWidth = w;
                    }
                    await ctx.sync();
                } catch (e2) {
                    _warn(`col width restore failed: ${e2.message}`);
                }
            }
        }
    }

    // ── Merge restore ─────────────────────────────────────────────────────────

    async function _restoreMerges(ctx, sheet, snapshot) {
        if (!snapshot.merges?.length) return;
        try {
            for (const addr of snapshot.merges) {
                try {
                    sheet.getRange(addr).merge(false);   // false = don't merge across
                } catch (me) {
                    _warn(`merge restore for ${addr} failed: ${me.message}`);
                }
            }
            await ctx.sync();
        } catch (e) {
            _warn(`merge restore failed: ${e.message}`);
        }
    }

    // ── Chart restore ─────────────────────────────────────────────────────────

    async function _restoreCharts(ctx, sheet, snapshot) {
        if (!snapshot.charts?.length) return;

        for (const cd of snapshot.charts) {
            try {
                let chart;
                if (cd.dataRange) {
                    // charts.add(type, dataRange, seriesBy)
                    const dataRng = sheet.getRange(cd.dataRange);
                    chart = sheet.charts.add(
                        cd.chartType || 'ColumnClustered',
                        dataRng,
                        'Auto',
                    );
                } else {
                    // No data range — add a blank chart of the same type
                    chart = sheet.charts.add(
                        cd.chartType || 'ColumnClustered',
                        sheet.getRange('A1'),   // placeholder; will be empty
                        'Auto',
                    );
                }

                // Position + size must be set before sync
                if (cd.left   != null) chart.left   = cd.left;
                if (cd.top    != null) chart.top    = cd.top;
                if (cd.width  != null) chart.width  = cd.width;
                if (cd.height != null) chart.height = cd.height;
                if (cd.name)           chart.name   = cd.name;

                await ctx.sync();

                // Title requires a second sync after the chart object exists
                if (cd.title) {
                    try {
                        chart.title.text    = cd.title;
                        chart.title.visible = true;
                        await ctx.sync();
                    } catch (_) { /* title not settable on all chart types */ }
                }
            } catch (e) {
                _warn(`chart restore for "${cd.name}" failed: ${e.message}`);
            }
        }
    }

    // ── Validation restore ────────────────────────────────────────────────────

    async function _restoreValidation(ctx, sheet, snapshot, startRow, startCol) {
        if (!snapshot.validation?.length) return;
        try {
            for (const vd of snapshot.validation) {
                try {
                    const addr = _cellAddr(startRow + vd.r, startCol + vd.c);
                    const dv   = sheet.getRange(addr).dataValidation;
                    dv.rule             = vd.rule;
                    dv.ignoreBlanks     = vd.ignoreBlanks     ?? true;
                    dv.showErrorAlert   = vd.showErrorAlert   ?? false;
                    if (vd.errorTitle)   dv.errorTitle   = vd.errorTitle;
                    if (vd.errorMessage) dv.errorMessage = vd.errorMessage;
                    dv.showInputMessage = vd.showInputMessage ?? false;
                    if (vd.promptTitle)  dv.promptTitle  = vd.promptTitle;
                    if (vd.prompt)       dv.prompt       = vd.prompt;
                } catch (ve) {
                    _warn(`validation restore for cell [${vd.r},${vd.c}] failed: ${ve.message}`);
                }
            }
            await ctx.sync();
        } catch (e) {
            _warn(`validation restore failed: ${e.message}`);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // helpers
    // ─────────────────────────────────────────────────────────────────────────

    function _emptySnapshot() {
        return {
            address:              null,
            rows:                 0,
            cols:                 0,
            formulas:             [],
            values:               [],
            format:               null,
            cellFormats:          [],
            borders:              [],
            rowHeights:           [],
            colWidths:            [],
            merges:               [],
            charts:               [],
            validation:           [],
            conditionalFormatCount: 0,
        };
    }

    function _stripSheet(addr) {
        if (!addr) return '';
        // Handle both "Sheet1!A1:B2" and "'Sheet 1'!A1:B2"
        const i = addr.lastIndexOf('!');
        return i >= 0 ? addr.slice(i + 1) : addr;
    }

    function _parseTopLeft(addr) {
        if (!addr) return { startRow: 0, startCol: 0 };
        // Strip any absolute $ signs before parsing
        const clean = addr.replace(/\$/g, '');
        const m     = clean.match(/^([A-Z]+)(\d+)/i);
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
