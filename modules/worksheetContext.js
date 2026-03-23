/**
 * worksheetContext.js
 * Reads worksheet state to provide LLM context.
 *
 * gather() supports context types:
 *   'selection'   — currently selected range (address, values, formulas)
 *   'sheet'       — used range of EVERY worksheet (address + values), keyed by
 *                   sheet name. Active sheet is always included; others are
 *                   capped at MAX_CELLS_PER_SHEET to keep the payload lean.
 *   'named-ranges'— workbook-level named ranges
 *   'styles'      — header + first data row formatting sample from active sheet
 *   'charts'      — chart objects on the active sheet
 *
 * sheetData shape (backward-compatible):
 *   {
 *     usedRange: { address, values },   // active sheet (same as before)
 *     allSheets: [                      // NEW — all sheets including active
 *       { name, address, values, rowCount, colCount },
 *       ...
 *     ],
 *     styles?: [...],
 *     charts?: [...],
 *   }
 */
const WorksheetContext = (() => {

    // Cap per non-active sheet to avoid ballooning context size.
    // Active sheet is sent in full (worksheetSnapshot already caps at 5000 cells).
    const MAX_CELLS_PER_SHEET = 2000;

    async function gather(contextTypes = ['selection']) {
        const ctx = {
            selection:   null,
            sheetData:   null,
            namedRanges: null,
            sheetNames:  null,
        };

        try {
            await Excel.run(async (exCtx) => {
                const wb    = exCtx.workbook;
                const sheet = wb.worksheets.getActiveWorksheet();

                // ── Selection ────────────────────────────────────────────────
                if (contextTypes.includes('selection')) {
                    const sel = wb.getSelectedRange();
                    sel.load(['address', 'values', 'formulas']);
                    await exCtx.sync();
                    ctx.selection = { address: sel.address, values: sel.values, formulas: sel.formulas };
                }

                // ── Sheet data (all worksheets) ───────────────────────────────
                // Always reads every worksheet when 'sheet' is requested.
                // allSheets[] gives the LLM visibility into cross-sheet references
                // (e.g. Data!A2:B6 used by VLOOKUP in Opportunities sheet).
                if (contextTypes.includes('sheet')) {
                    const allSheets = wb.worksheets;
                    allSheets.load('items/name');
                    await exCtx.sync();

                    const sheetDataList = [];
                    let activeSheetUsedRange = null;

                    for (const ws of allSheets.items) {
                        try {
                            const used = ws.getUsedRangeOrNullObject();
                            used.load('isNullObject');
                            await exCtx.sync();

                            if (used.isNullObject) {
                                sheetDataList.push({ name: ws.name, address: null, values: [], rowCount: 0, colCount: 0 });
                                continue;
                            }

                            used.load(['address', 'rowCount', 'columnCount']);
                            await exCtx.sync();

                            const rows  = used.rowCount;
                            const cols  = used.columnCount;
                            const cells = rows * cols;
                            const isActive = ws.name === sheet.name;

                            // For non-active sheets, truncate to avoid context explosion.
                            // We load the full range but slice values client-side.
                            if (!isActive && cells > MAX_CELLS_PER_SHEET) {
                                // Load only first N rows worth of data
                                const maxRows = Math.max(1, Math.floor(MAX_CELLS_PER_SHEET / cols));
                                const truncRange = ws.getRangeByIndexes(0, 0, maxRows, cols);
                                truncRange.load(['address', 'values']);
                                await exCtx.sync();
                                sheetDataList.push({
                                    name:      ws.name,
                                    address:   truncRange.address,
                                    values:    truncRange.values,
                                    rowCount:  rows,   // actual full row count
                                    colCount:  cols,
                                    truncated: true,
                                    shownRows: maxRows,
                                });
                            } else {
                                used.load(['values']);
                                await exCtx.sync();
                                const entry = {
                                    name:     ws.name,
                                    address:  used.address,
                                    values:   used.values,
                                    rowCount: rows,
                                    colCount: cols,
                                };
                                sheetDataList.push(entry);
                                if (isActive) activeSheetUsedRange = entry;
                            }
                        } catch (sheetErr) {
                            console.warn(`[WorksheetContext] sheet "${ws.name}" error:`, sheetErr.message);
                            sheetDataList.push({ name: ws.name, error: sheetErr.message });
                        }
                    }

                    // Backward-compat: sheetData.usedRange still points to active sheet
                    const activeEntry = activeSheetUsedRange
                        || sheetDataList.find(s => s.name === sheet.name)
                        || null;

                    ctx.sheetData = {
                        usedRange: activeEntry
                            ? { address: activeEntry.address, values: activeEntry.values }
                            : null,
                        allSheets: sheetDataList,
                    };
                }

                // ── Named ranges ─────────────────────────────────────────────
                if (contextTypes.includes('named-ranges')) {
                    const names = wb.names;
                    names.load('items');
                    await exCtx.sync();
                    ctx.namedRanges = names.items.map(n => ({ name: n.name, value: n.value }));
                }

                // ── Styles (sampled formatting snapshot from active sheet) ────
                if (contextTypes.includes('styles')) {
                    ctx.sheetData = ctx.sheetData || {};
                    ctx.sheetData.styles = [];
                    try {
                        const used = sheet.getUsedRangeOrNullObject();
                        used.load('isNullObject');
                        await exCtx.sync();

                        if (!used.isNullObject) {
                            used.load(['rowCount', 'columnCount', 'address']);
                            await exCtx.sync();

                            const rowsToSample = Math.min(used.rowCount, 2);
                            for (let r = 0; r < rowsToSample; r++) {
                                const colsToSample = Math.min(used.columnCount, 3);
                                for (let c = 0; c < colsToSample; c++) {
                                    const cell = used.getCell(r, c);
                                    cell.load([
                                        'address',
                                        'format/fill/color',
                                        'format/font/color',
                                        'format/font/bold',
                                        'format/font/italic',
                                        'format/font/size',
                                        'format/numberFormat',
                                        'format/horizontalAlignment',
                                    ]);
                                    await exCtx.sync();
                                    ctx.sheetData.styles.push({
                                        address:             cell.address,
                                        fillColor:           cell.format.fill.color           || null,
                                        fontColor:           cell.format.font.color           || null,
                                        fontBold:            cell.format.font.bold,
                                        fontItalic:          cell.format.font.italic,
                                        fontSize:            cell.format.font.size            || null,
                                        numberFormat:        cell.format.numberFormat         || null,
                                        horizontalAlignment: cell.format.horizontalAlignment  || null,
                                    });
                                }
                            }
                        }
                    } catch (styleErr) {
                        console.warn('[WorksheetContext] styles error:', styleErr.message);
                    }
                }

                // ── Charts (active sheet only) ───────────────────────────────
                if (contextTypes.includes('charts')) {
                    ctx.sheetData = ctx.sheetData || {};
                    ctx.sheetData.charts = [];
                    try {
                        const charts = sheet.charts;
                        charts.load('items/name,items/chartType,items/title/text');
                        await exCtx.sync();
                        for (const chart of charts.items) {
                            let dataRange = null;
                            try {
                                const dr = chart.getDataRange();
                                dr.load('address');
                                await exCtx.sync();
                                dataRange = dr.address || null;
                            } catch (_) {}
                            ctx.sheetData.charts.push({
                                name:      chart.name      || null,
                                chartType: chart.chartType || null,
                                title:     chart.title?.text || null,
                                dataRange,
                            });
                        }
                    } catch (chartErr) {
                        console.warn('[WorksheetContext] charts error:', chartErr.message);
                    }
                }

                // ── Sheet names (always) ─────────────────────────────────────
                const sheets = wb.worksheets;
                sheets.load('items/name');
                await exCtx.sync();
                ctx.sheetNames = sheets.items.map(s => s.name);
            });
        } catch (err) {
            console.warn('[WorksheetContext] gather error:', err.message);
        }

        return ctx;
    }

    return { gather };
})();
