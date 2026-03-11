/**
 * chatManager.js — chat UI + orchestration pipeline.
 */
const ChatManager = (() => {

    const _feed    = document.getElementById('message-feed');
    const _input   = document.getElementById('chat-input');
    const _sendBtn = document.getElementById('send-button');
    const _typing  = document.getElementById('typing-indicator');

    let _isBusy = false;

    function init() {
        _input.addEventListener('input', () => {
            _input.style.height = 'auto';
            _input.style.height = Math.min(_input.scrollHeight, 100) + 'px';
            _sendBtn.disabled   = _input.value.trim() === '' || _isBusy;
        });
        _input.addEventListener('keydown', e => {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); if (!_sendBtn.disabled) _handleSend(); }
        });
        _sendBtn.addEventListener('click', _handleSend);
        document.getElementById('execution-header')?.addEventListener('click', () => {
            const panel = document.getElementById('execution-panel');
            const icon  = document.getElementById('execution-toggle-icon');
            panel.classList.toggle('open');
            // ▲ = collapsed (click to expand), ▼ = expanded (click to collapse)
            icon.textContent = panel.classList.contains('open') ? '▼' : '▲';
        });
    }

    async function _handleSend() {
        const text = _input.value.trim();
        if (!text || _isBusy) return;
        _setBusy(true);
        _appendMessage('user', text);
        _clearInput();
        _showTyping(true);

        try {
            const wsCtx = await WorksheetContext.gather(['selection','sheet']);

            // 1. Scaffold rubric from user prompt
            let rubric = null;
            try {
                rubric = await LLMClient.rubricScaffold(text, wsCtx);
                StepNavigator.setRubric(rubric);
            } catch(e) { console.warn('[ChatManager] rubric scaffold failed:', e.message); }

            // 2. Generate code segments (rubric passed as optional context)
            const segments = await LLMClient.generateCode(text, wsCtx, rubric);
            _showTyping(false);

            // 3. Load segments into navigator so graph renders in rubric gate
            StepNavigator.loadSegments(segments);

            // 4. Show rubric gate — await user clicking "Start →"
            await StepNavigator.showRubricGate();

            // 5. Execute
            _appendMessage('agent', `Applying ${segments.length} step(s) to your sheet…`);
            await ExecutionEngine.run(segments);

        } catch(err) {
            _showTyping(false);
            _appendMessage('agent', `⚠️ ${err.message}`);
        }

        _setBusy(false);
    }

    function _appendMessage(role, text) {
        const w = document.createElement('div');
        w.className = `message ${role}`;
        const b = document.createElement('div');
        b.className = 'message-bubble'; b.textContent = text;
        const m = document.createElement('span');
        m.className = 'message-meta'; m.textContent = role === 'user' ? 'You' : 'Assistant';
        w.appendChild(b); w.appendChild(m);
        _feed.insertBefore(w, _typing);
        _feed.scrollTop = _feed.scrollHeight;
    }

    function _showTyping(v) {
        _typing.classList.toggle('visible', v);
        _feed.scrollTop = _feed.scrollHeight;
    }

    function _setBusy(b) {
        _isBusy = b;
        _sendBtn.disabled = b || _input.value.trim() === '';
        _input.disabled   = b;
    }

    function _clearInput() {
        _input.value = ''; _input.style.height = 'auto';
    }

    return { init };
})();
