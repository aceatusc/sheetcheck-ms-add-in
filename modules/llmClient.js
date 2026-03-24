/**
 * llmClient.js — thin fetch wrapper for all /addin/* endpoints.
 */
const LLMClient = (() => {

    const BASE_URL      = "https://sackend.isi.edu/sheetcheck/backend/addin";
    const SHARED_SECRET = "my-super-secret-2025";

    function _headers() {
        return { 'Content-Type': 'application/json', 'X-Addin-Secret': SHARED_SECRET };
    }

    // Timeout per endpoint type (ms). Larger models on /code and /edit can
    // take several minutes; Q&A and chat are much faster.
    const TIMEOUTS = {
        '/code':            300_000,   // 5 min
        '/edit':            300_000,   // 5 min
        '/rubric/verify':    120_000,   // 2 min
        '/rubric/scaffold':  120_000,   // 2 min
        '/ask':              60_000,   // 60 s
        '/chat':             30_000,   // 30 s
    };
    const DEFAULT_TIMEOUT = 120_000;   // 2 min fallback

    async function _post(path, body) {
        const ms         = TIMEOUTS[path] ?? DEFAULT_TIMEOUT;
        const controller = new AbortController();
        const timer      = setTimeout(() => controller.abort(), ms);

        let res;
        try {
            res = await fetch(`${BASE_URL}${path}`, {
                method: 'POST', headers: _headers(), body: JSON.stringify(body),
                signal: controller.signal,
            });
        } catch (err) {
            if (err.name === 'AbortError') {
                throw new Error(`Request timed out after ${ms / 1000}s — the model is taking too long. Try again or switch to a faster model.`);
            }
            throw err;
        } finally {
            clearTimeout(timer);
        }

        if (!res.ok) {
            const e = await res.json().catch(() => ({}));
            throw new Error(e.error || `Server error ${res.status}`);
        }
        return res.json();
    }

    /** Generate code segments for a task. */
    async function generateCode(message, wsContext, rubric = null, chatHistory = []) {
        const body = { message, context: wsContext, chat_history: chatHistory };
        if (rubric) body.rubric = rubric;
        const data = await _post('/code', body);
        return data.segments;
    }

    /** Follow-up Q&A about a step. */
    async function ask(message, wsContext, step, history = [], chatHistory = []) {
        return _post('/ask', { message, context: wsContext, step, history, chat_history: chatHistory });
    }

    /** Edit a segment and regenerate the remainder of the chain.
     *  Returns an array: [editedSeg, ...regeneratedRemainder] */
    async function edit(message, wsContext, segment, remainingSegments = [], chatHistory = []) {
        const data = await _post('/edit', {
            message,
            context: wsContext,
            segment,
            remaining_segments: remainingSegments,
            chat_history: chatHistory,
        });
        return data.segments;
    }

    /** Scaffold an initial rubric for the task. */
    async function rubricScaffold(message, wsContext, chatHistory = []) {
        return _post('/rubric/scaffold', { message, context: wsContext, chat_history: chatHistory });
    }

    /** Verify worksheet against rubric. */
    async function rubricVerify(rubric, wsContext, chatHistory = []) {
        return _post('/rubric/verify', { rubric, context: wsContext, chat_history: chatHistory });
    }

    /** General chat proxy. */
    async function chat(message, wsContext, chatHistory = []) {
        return _post('/chat', { message, context: wsContext, chat_history: chatHistory });
    }

    // -- Exposed for InteractionLogger ----------------------------------------

    /** Full URL for the interactions endpoint (used with sendBeacon). */
    function interactionsUrl() {
        return `${BASE_URL}/interactions`;
    }

    /** Auth headers without Content-Type (caller sets it via Blob for sendBeacon). */
    function authHeaders() {
        return { 'X-Addin-Secret': SHARED_SECRET };
    }

    return { generateCode, ask, edit, rubricScaffold, rubricVerify, chat,
             interactionsUrl, authHeaders };
})();
