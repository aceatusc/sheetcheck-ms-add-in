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
 *     allSheets: [                      // all sheets including active
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

        // sheetData is always initialised so that styles/charts can attach
        // to it regardless of whether 'sheet' is in contextTypes.
        ctx.sheetData = {
            usedRange: null,
            allSheets: [],
        };

        try {
            await Excel.run(async (exCtx) => {
                const wb         = exCtx.workbook;
                const activeSheet = wb.worksheets.getActiveWorksheet();

                // Load active sheet name once so we can compare later.
                // Desktop Excel requires properties to be loaded before access.
                activeSheet.load('name');
                await exCtx.sync();
                const activeSheetName = activeSheet.name;

                // ── Selection ────────────────────────────────────────────────
                if (contextTypes.includes('selection')) {
                    const sel = wb.getSelectedRange();
                    sel.load(['address', 'values', 'formulas']);
                    await exCtx.sync();
                    ctx.selection = {
                        address:  sel.address,
                        values:   sel.values,
                        formulas: sel.formulas,
                    };
                }

                // ── Sheet data (all worksheets) ───────────────────────────────
                // Always reads every worksheet when 'sheet' is requested.
                // allSheets[] gives the LLM visibility into cross-sheet references
                // (e.g. Data!A2:B6 used by VLOOKUP in Opportunities sheet).
                if (contextTypes.includes('sheet')) {
                    const allSheetsCollection = wb.worksheets;
                    allSheetsCollection.load('items/name');
                    await exCtx.sync();

                    const sheetDataList = [];
                    let activeEntry = null;

                    for (const ws of allSheetsCollection.items) {
                        const isActive = ws.name === activeSheetName;
                        try {
                            // Use getUsedRangeOrNullObject + isNullObject.
                            // On Desktop we must load isNullObject in the same
                            // batch as the initial load — do NOT rely on the
                            // property being readable before sync().
                            const used = ws.getUsedRangeOrNullObject();
                            used.load(['isNullObject', 'rowCount', 'columnCount', 'address']);
                            await exCtx.sync();

                            if (used.isNullObject) {
                                const emptyEntry = {
                                    name:     ws.name,
                                    address:  null,
                                    values:   [],
                                    rowCount: 0,
                                    colCount: 0,
                                };
                                sheetDataList.push(emptyEntry);
                                if (isActive) activeEntry = emptyEntry;
                                continue;
                            }

                            const rows  = used.rowCount;
                            const cols  = used.columnCount;
                            const cells = rows * cols;

                            if (!isActive && cells > MAX_CELLS_PER_SHEET) {
                                // Truncate non-active oversized sheets to keep context lean.
                                const maxRows    = Math.max(1, Math.floor(MAX_CELLS_PER_SHEET / cols));
                                const truncRange = ws.getRangeByIndexes(0, 0, maxRows, cols);
                                truncRange.load(['address', 'values']);
                                await exCtx.sync();

                                sheetDataList.push({
                                    name:      ws.name,
                                    address:   truncRange.address,
                                    values:    truncRange.values,
                                    rowCount:  rows,     // actual full row count
                                    colCount:  cols,
                                    truncated: true,
                                    shownRows: maxRows,
                                });
                            } else {
                                // Load values in a separate sync — Desktop Excel
                                // handles this more reliably than loading everything
                                // at once when the range object came from getUsedRangeOrNullObject.
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
                                if (isActive) activeEntry = entry;
                            }
                        } catch (sheetErr) {
                            console.warn(`[WorksheetContext] sheet "${ws.name}" error:`, sheetErr.message);
                            sheetDataList.push({ name: ws.name, error: sheetErr.message });
                        }
                    }

                    // If the active sheet was truncated it won't be in activeEntry yet;
                    // fall back to finding it in the list.
                    if (!activeEntry) {
                        activeEntry = sheetDataList.find(s => s.name === activeSheetName) || null;
                    }

                    ctx.sheetData.usedRange = activeEntry
                        ? { address: activeEntry.address, values: activeEntry.values }
                        : null;
                    ctx.sheetData.allSheets = sheetDataList;
                }

                // ── Named ranges ─────────────────────────────────────────────
                if (contextTypes.includes('named-ranges')) {
                    const names = wb.names;
                    names.load('items');
                    await exCtx.sync();
                    ctx.namedRanges = names.items.map(n => ({
                        name:  n.name,
                        value: n.value,
                    }));
                }

                // ── Styles (sampled formatting snapshot from active sheet) ────
                if (contextTypes.includes('styles')) {
                    ctx.sheetData.styles = [];
                    try {
                        const used = activeSheet.getUsedRangeOrNullObject();
                        // Batch isNullObject + dimensions in one load/sync for Desktop compat.
                        used.load(['isNullObject', 'rowCount', 'columnCount', 'address']);
                        await exCtx.sync();

                        if (!used.isNullObject) {
                            const rowsToSample = Math.min(used.rowCount, 2);
                            const colsToSample = Math.min(used.columnCount, 3);

                            // Collect all cell objects first, then load in one batch.
                            const cellsToLoad = [];
                            for (let r = 0; r < rowsToSample; r++) {
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
                                    cellsToLoad.push(cell);
                                }
                            }
                            // Single sync for all cells — much faster on Desktop.
                            await exCtx.sync();

                            for (const cell of cellsToLoad) {
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
                    } catch (styleErr) {
                        console.warn('[WorksheetContext] styles error:', styleErr.message);
                    }
                }

                // ── Charts (active sheet only) ───────────────────────────────
                if (contextTypes.includes('charts')) {
                    ctx.sheetData.charts = [];
                    try {
                        const charts = activeSheet.charts;
                        charts.load('items/name,items/chartType,items/title/text');
                        await exCtx.sync();

                        for (const chart of charts.items) {
                            let dataRange = null;
                            try {
                                const dr = chart.getDataRange();
                                dr.load('address');
                                await exCtx.sync();
                                dataRange = dr.address || null;
                            } catch (_) {
                                // getDataRange() can throw on some chart types — safe to ignore.
                            }
                            ctx.sheetData.charts.push({
                                name:      chart.name           || null,
                                chartType: chart.chartType      || null,
                                title:     chart.title?.text    || null,
                                dataRange,
                            });
                        }
                    } catch (chartErr) {
                        console.warn('[WorksheetContext] charts error:', chartErr.message);
                    }
                }

                // ── Sheet names (always) ─────────────────────────────────────
                const sheetsCollection = wb.worksheets;
                sheetsCollection.load('items/name');
                await exCtx.sync();
                ctx.sheetNames = sheetsCollection.items.map(s => s.name);
            });
        } catch (err) {
            console.warn('[WorksheetContext] gather error:', err.message);
        }

        return ctx;
    }

    return { gather };
})();
