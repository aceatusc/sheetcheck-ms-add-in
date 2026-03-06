/**
 * llmClient.js
 * Sends prompt + context to the LLM API.
 * Returns an array of CodeSegments to execute.
 *
 * PLACEHOLDER — implementation coming.
 *
 * CodeSegment shape:
 *   {
 *     id:           string,
 *     description:  string,
 *     code:         string,   // Office.js async code to run
 *     pauseAfterMs: number,   // delay before next segment
 *   }
 */
const LLMClient = (() => {

    // TODO: replace with real endpoint / model selection
    const CONFIG = {
        endpoint: 'PLACEHOLDER_API_ENDPOINT',
        model:    'PLACEHOLDER_MODEL_NAME',
        apiKey:   'PLACEHOLDER_API_KEY',  // Note: use a server-side proxy in production
    };

    /**
     * Send user message + worksheet context to the LLM.
     * @param {string} userMessage  - Raw user input from chat
     * @param {object} wsContext    - Output of WorksheetContext.gather()
     * @returns {Promise<CodeSegment[]>}
     */
    async function sendMessage(userMessage, wsContext) {
        // PLACEHOLDER: construct prompt and POST to LLM
        // const payload = buildPrompt(userMessage, wsContext);
        // const response = await fetch(CONFIG.endpoint, { method: 'POST', ... });
        // const raw = await response.json();
        // return parseSegments(raw);

        console.warn('[LLMClient] Using stub response. Implement sendMessage().');
        return _stubSegments(userMessage);
    }

    /**
     * Build the system + user prompt for the LLM.
     * PLACEHOLDER.
     */
    function buildPrompt(userMessage, wsContext) {
        // TODO: craft a prompt that instructs the LLM to return
        // a JSON array of { description, code, pauseAfterMs } objects.
        return {
            system: 'PLACEHOLDER_SYSTEM_PROMPT',
            user:   userMessage,
        };
    }

    /**
     * Parse raw LLM response text into CodeSegment[].
     * PLACEHOLDER.
     */
    function parseSegments(rawText) {
        // TODO: strip markdown fences, JSON.parse, validate shape
        return [];
    }

    // --- Stub for development only ---
    function _stubSegments(msg) {
        return [
            {
                id:           'seg-1',
                description:  'Write header row labels',
                sheet_context: ['A1:E1'],
                explanation:  'Creates the five column headers — Month, Revenue, Expenses, Profit, and Growth % — in row 1. These labels define the structure of the entire table.',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        sheet.getRange("A1:E1").values = [["Month", "Revenue", "Expenses", "Profit", "Growth %"]];
                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 1000,
            },
            {
                id:           'seg-2',
                description:  'Style header row (bold, background, font color)',
                sheet_context: ['A1:E1'],
                explanation:  'Applies a dark background (#1a1d27) with blue bold text to A1:E1, making the headers visually distinct from the data rows below.',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        const header = sheet.getRange("A1:E1");
                        header.format.fill.color       = "#1a1d27";
                        header.format.font.color       = "#4f8ef7";
                        header.format.font.bold        = true;
                        header.format.font.size        = 11;
                        header.format.horizontalAlignment = "Center";
                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 900,
            },
            {
                id:           'seg-3',
                description:  'Fill in monthly data rows',
                sheet_context: ['A2:E7', 'A2:A7'],
                explanation:  'Writes six months of raw data into A2:E7. Columns B–D hold numeric values for Revenue, Expenses, and Profit. Column E is left blank here — Growth % formulas are added in the next step.',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        const data = [
                            ["Jan", 142000, 98000,  44000, ""],
                            ["Feb", 158000, 104000, 54000, ""],
                            ["Mar", 175000, 110000, 65000, ""],
                            ["Apr", 163000, 107000, 56000, ""],
                            ["May", 191000, 115000, 76000, ""],
                            ["Jun", 210000, 121000, 89000, ""],
                        ];
                        sheet.getRange("A2:E7").values = data;
                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 1200,
            },
            {
                id:           'seg-4',
                description:  'Add Growth % formulas',
                sheet_context: ['E3:E7', 'D2:D7'],
                explanation:  'Inserts IFERROR formulas in E3:E7 that compute month-over-month profit growth: (current − previous) ÷ previous. E2 is skipped because Jan has no prior month to compare against. Results are formatted as percentages with one decimal place.',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        sheet.getRange("E3:E7").formulas = [
                            ['=IFERROR((D3-D2)/D2, "")'],
                            ['=IFERROR((D4-D3)/D3, "")'],
                            ['=IFERROR((D5-D4)/D4, "")'],
                            ['=IFERROR((D6-D5)/D5, "")'],
                            ['=IFERROR((D7-D6)/D6, "")'],
                        ];
                        sheet.getRange("E3:E7").numberFormat = [["0.0%"],["0.0%"],["0.0%"],["0.0%"],["0.0%"]];
                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 1000,
            },
            {
                id:           'seg-5',
                description:  'Format Revenue, Expenses, Profit as currency',
                sheet_context: ['B2:D7'],
                explanation:  'Applies the $#,##0 number format to the three numeric columns so values render as dollar amounts with comma separators (e.g. $142,000). This is display-only — the underlying values remain plain numbers.',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        sheet.getRange("B2:D7").numberFormat = [
                            ["$#,##0", "$#,##0", "$#,##0"],
                            ["$#,##0", "$#,##0", "$#,##0"],
                            ["$#,##0", "$#,##0", "$#,##0"],
                            ["$#,##0", "$#,##0", "$#,##0"],
                            ["$#,##0", "$#,##0", "$#,##0"],
                            ["$#,##0", "$#,##0", "$#,##0"],
                        ];
                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 700,
            },
            {
                id:           'seg-6',
                description:  'Zebra-stripe data rows',
                sheet_context: ['A2:E7'],
                explanation:  'Alternates the row background between #f5f7ff (even rows) and white (odd rows) across the full data range A2:E7. This improves readability when scanning across wide rows.',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        for (let i = 2; i <= 7; i++) {
                            const row = sheet.getRange("A" + i + ":E" + i);
                            row.format.fill.color = i % 2 === 0 ? "#f5f7ff" : "#ffffff";
                        }
                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 800,
            },
            {
                id:           'seg-7',
                description:  'Colour Profit column by value',
                sheet_context: ['D2:D7'],
                explanation:  'Reads each cell in D2:D7, then colours the font green (#1a7a4a) and bold for months with profit ≥ $60,000, or red (#b94040) for those below. This gives an instant visual signal of high vs low-performing months.',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet  = ctx.workbook.worksheets.getActiveWorksheet();
                        const profit = sheet.getRange("D2:D7");
                        profit.load("values");
                        await ctx.sync();

                        profit.values.forEach((row, i) => {
                            const cell = sheet.getRange("D" + (i + 2));
                            cell.format.font.color = row[0] >= 60000 ? "#1a7a4a" : "#b94040";
                            cell.format.font.bold  = row[0] >= 60000;
                        });
                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 900,
            },
            {
                id:           'seg-8',
                description:  'Add totals row and auto-fit columns',
                sheet_context: ['A8:E8', 'A1:E8'],
                explanation:  'Appends a TOTAL row in row 8 using SUM formulas for Revenue, Expenses, and Profit (B8:D8). The row gets the same dark styling as the header. Finally, autofitColumns() resizes all five columns to fit their widest content.',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

                        sheet.getRange("A8").values        = [["TOTAL"]];
                        sheet.getRange("B8:D8").formulas   = [["=SUM(B2:B7)", "=SUM(C2:C7)", "=SUM(D2:D7)"]];
                        sheet.getRange("B8:D8").numberFormat = [["$#,##0", "$#,##0", "$#,##0"]];

                        const totalsRow = sheet.getRange("A8:E8");
                        totalsRow.format.fill.color = "#1a1d27";
                        totalsRow.format.font.color = "#ffffff";
                        totalsRow.format.font.bold  = true;

                        sheet.getRange("A1:E8").getEntireColumn().format.autofitColumns();

                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 0,
            },
        ];
    }

    return { sendMessage };
})();
