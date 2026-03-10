/**
 * llmClient.js
 * Sends the user message + worksheet context to a Python proxy
 * which handles test mode, LLM API calls, and response parsing.
 *
 * To use the stub/test data, set TEST_MODE = True in server.py.
 *
 * CodeSegment shape (returned by server):
 *   {
 *     id:            string,
 *     description:   string,
 *     sheet_context: string[],
 *     explanation:   string,
 *     code:          string,   // Office.js async code to run
 *   }
 */
const LLMClient = (() => {

    // ── Config ────────────────────────────────────────────────────────────────

    const SERVER_URL    = "https://sackend.isi.edu/sheetcheck/backend/addin/code";
    const SHARED_SECRET = "my-super-secret-2025"; // must match SHARED_SECRET in backend app/server.py

    // ── Public ────────────────────────────────────────────────────────────────

    /**
     * Send user message + worksheet context to the Python proxy.
     *
     * @param {string} userMessage  - Raw user input from chat
     * @param {object} wsContext    - Output of WorksheetContext.gather()
     * @returns {Promise<CodeSegment[]>}
     */
    async function sendMessage(userMessage, wsContext) {
        const response = await fetch(SERVER_URL, {
            method:  'POST',
            headers: {
                'Content-Type':   'application/json',
                'X-Addin-Secret': SHARED_SECRET,
            },
            body: JSON.stringify({
                message: userMessage,
                context: wsContext,
            }),
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.error || `Server error ${response.status}`);
        }

        const data = await response.json();
        return data.segments;
    }

    return { sendMessage };
})();