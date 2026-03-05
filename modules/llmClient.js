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
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        // First data row has no prior row to compare — leave blank
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
                description:  'Colour Profit column by value (green / red)',
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
                description:  'Auto-fit column widths and add a totals row',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

                        // Totals row
                        sheet.getRange("A8").values        = [["TOTAL"]];
                        sheet.getRange("B8:D8").formulas   = [["=SUM(B2:B7)", "=SUM(C2:C7)", "=SUM(D2:D7)"]];
                        sheet.getRange("B8:D8").numberFormat = [["$#,##0", "$#,##0", "$#,##0"]];

                        const totalsRow = sheet.getRange("A8:E8");
                        totalsRow.format.fill.color   = "#1a1d27";
                        totalsRow.format.font.color   = "#ffffff";
                        totalsRow.format.font.bold    = true;

                        // Auto-fit all columns
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
