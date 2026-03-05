/**
 * worksheetContext.js
 * Reads worksheet state to provide LLM context.
 *
 * PLACEHOLDER — implementation coming.
 */
const WorksheetContext = (() => {

    /**
     * Gather context from the active worksheet.
     * @param {string[]} contextTypes - e.g. ['selection', 'sheet', 'named-ranges']
     * @returns {Promise<object>} Structured context object sent to the LLM.
     */
    async function gather(contextTypes = ['selection']) {
        // PLACEHOLDER: implement per context type
        // Example shape to fill in:
        const ctx = {
            selection:   null,   // { address, values, formulas }
            sheetData:   null,   // { usedRange: { address, values } }
            namedRanges: null,   // [{ name, address }]
            sheetNames:  null,   // string[]
        };

        // TODO: use Excel.run to populate each requested field
        // await Excel.run(async (context) => { ... });

        return ctx;
    }

    return { gather };
})();
