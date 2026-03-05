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
                description:  'Write message to A1',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        sheet.getRange("A1").values = [["Hello from AI!"]];
                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 800,
            },
            {
                id:           'seg-2',
                description:  'Highlight A1 yellow',
                code:         `
                    await Excel.run(async (ctx) => {
                        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                        sheet.getRange("A1").format.fill.color = "yellow";
                        await ctx.sync();
                    });
                `,
                pauseAfterMs: 600,
            },
        ];
    }

    return { sendMessage };
})();
