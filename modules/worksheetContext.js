/**
 * worksheetContext.js
 * Reads worksheet state to provide LLM context.
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

                if (contextTypes.includes('selection')) {
                    const sel = wb.getSelectedRange();
                    sel.load(['address','values','formulas']);
                    await exCtx.sync();
                    ctx.selection = { address: sel.address, values: sel.values, formulas: sel.formulas };
                }

                if (contextTypes.includes('sheet')) {
                    const used = sheet.getUsedRange();
                    used.load(['address','values']);
                    await exCtx.sync();
                    ctx.sheetData = { usedRange: { address: used.address, values: used.values } };
                }

                if (contextTypes.includes('named-ranges')) {
                    const names = wb.names;
                    names.load('items');
                    await exCtx.sync();
                    ctx.namedRanges = names.items.map(n => ({ name: n.name, value: n.value }));
                }

                // Always include sheet names
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
