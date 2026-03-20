/**
 * chatHistory.js — Shared session-scoped chat history.
 *
 * ChatManager pushes user messages here after each send.
 * StepNavigator reads from here when making LLM calls (edit, ask)
 * so those calls have the same conversation context as /code.
 *
 * Kept intentionally minimal — a plain array with a max-length cap.
 */
const ChatHistory = (() => {
    const MAX = 10;
    let _messages = [];

    /** Append a user message. Called by ChatManager after each send. */
    function push(text) {
        _messages.push(text);
        if (_messages.length > MAX) _messages.shift();
    }

    /** Return a copy of the current history (oldest first). */
    function get() { return [..._messages]; }

    /** Clear history — call between sessions if needed. */
    function reset() { _messages = []; }

    return { push, get, reset };
})();
