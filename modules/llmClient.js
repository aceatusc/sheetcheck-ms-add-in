/**
 * llmClient.js — thin fetch wrapper for all /addin/* endpoints.
 */
const LLMClient = (() => {

    const BASE_URL      = "https://sackend.isi.edu/sheetcheck/backend/addin";
    const SHARED_SECRET = "my-super-secret-2025";

    function _headers() {
        return { 'Content-Type': 'application/json', 'X-Addin-Secret': SHARED_SECRET };
    }

    async function _post(path, body) {
        const res = await fetch(`${BASE_URL}${path}`, {
            method: 'POST', headers: _headers(), body: JSON.stringify(body),
        });
        if (!res.ok) {
            const e = await res.json().catch(() => ({}));
            throw new Error(e.error || `Server error ${res.status}`);
        }
        return res.json();
    }

    /** Generate code segments for a task. */
    async function generateCode(message, wsContext, rubric = null) {
        const body = { message, context: wsContext };
        if (rubric) body.rubric = rubric;
        const data = await _post('/code', body);
        return data.segments;
    }

    /** Follow-up Q&A about a step. */
    async function ask(message, wsContext, step, history = []) {
        return _post('/ask', { message, context: wsContext, step, history });
    }

    /** Edit a segment and regenerate the rest of the chain from that point.
     *  Returns an array: [editedSeg, ...regeneratedRemainder] */
    // TODO: should not get remainingSegments; it should generate it as a new full path
    async function edit(message, wsContext, segment, remainingSegments = []) {
        const data = await _post('/edit', {
            message,
            context: wsContext,
            segment,
            remaining_segments: remainingSegments,
        });
        // Normalise: server returns { segments: [...] }
        return data.segments;
    }

    /** Scaffold an initial rubric for the task. */
    async function rubricScaffold(message, wsContext) {
        return _post('/rubric/scaffold', { message, context: wsContext });
    }

    /** Verify worksheet against rubric. */
    async function rubricVerify(rubric, wsContext) {
        return _post('/rubric/verify', { rubric, context: wsContext });
    }

    /** General chat proxy. */
    async function chat(message, wsContext) {
        return _post('/chat', { message, context: wsContext });
    }

    return { generateCode, ask, edit, rubricScaffold, rubricVerify, chat };
})();
