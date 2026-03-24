/**
 * worksheetContext.js
 * Reads worksheet state to provide LLM context.
 *
 * gather() now supports two additional context types:
 *   'styles'  — samples fill/font/numberFormat from the used range header row
 *               and a representative data row, so the LLM knows the current
 *               formatting theme and doesn't override it unnecessarily.
 *   'charts'  — lists charts present on the active sheet (name, type, data
 *               range, title) so the LLM is aware of existing visualisations.
 *
 * Default call from chatManager: gather(['selection', 'sheet'])
 * Full call for richer context:  gather(['selection', 'sheet', 'styles', 'charts'])
 */
const WorksheetContext = (() => {

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

                // ── Sheet data ───────────────────────────────────────────────
                if (contextTypes.includes('sheet')) {
                    const used = sheet.getUsedRange();
                    used.load(['address', 'values']);
                    await exCtx.sync();
                    ctx.sheetData = { usedRange: { address: used.address, values: used.values } };
                }

                // ── Named ranges ─────────────────────────────────────────────
                if (contextTypes.includes('named-ranges')) {
                    const names = wb.names;
                    names.load('items');
                    await exCtx.sync();
                    ctx.namedRanges = names.items.map(n => ({ name: n.name, value: n.value }));
                }

                // ── Styles (sampled formatting snapshot) ─────────────────────
                // Loads fill colour, font colour/bold/size, number format, and
                // horizontal alignment for the header row and first data row.
                // Sampling keeps the payload small while still conveying theme.
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

                            // Sample up to 2 rows: header (row 0) + first data row (row 1)
                            const rowsToSample = Math.min(used.rowCount, 2);
                            for (let r = 0; r < rowsToSample; r++) {
                                // Sample up to 3 cells per row to keep payload small
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

                // ── Charts ───────────────────────────────────────────────────
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
                                // getDataRange() throws if the chart has no data range
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
