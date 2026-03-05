/**
 * chatManager.js
 * Manages the chat UI: message rendering, input handling,
 * context chip toggles, and orchestrates the full pipeline:
 *   user input → WorksheetContext → LLMClient → ExecutionEngine
 */
const ChatManager = (() => {

    const _feed         = document.getElementById('message-feed');
    const _input        = document.getElementById('chat-input');
    const _sendBtn      = document.getElementById('send-button');
    const _typingEl     = document.getElementById('typing-indicator');
    const _contextChips = document.querySelectorAll('.context-chip');

    let _activeContexts = new Set(['selection']);
    let _isBusy         = false;

    function init() {
        // Auto-resize textarea
        _input.addEventListener('input', () => {
            _input.style.height = 'auto';
            _input.style.height = Math.min(_input.scrollHeight, 100) + 'px';
            _sendBtn.disabled = _input.value.trim() === '' || _isBusy;
        });

        // Send on Enter (Shift+Enter = newline)
        _input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                if (!_sendBtn.disabled) _handleSend();
            }
        });

        _sendBtn.addEventListener('click', _handleSend);

        // Context chip toggles
        _contextChips.forEach(chip => {
            chip.addEventListener('click', () => {
                const ctx = chip.dataset.context;
                if (_activeContexts.has(ctx)) {
                    _activeContexts.delete(ctx);
                    chip.classList.remove('active');
                } else {
                    _activeContexts.add(ctx);
                    chip.classList.add('active');
                }
            });
        });

        // Execution panel collapse toggle
        document.getElementById('execution-header').addEventListener('click', () => {
            const panel = document.getElementById('execution-panel');
            const icon  = document.getElementById('execution-toggle-icon');
            panel.classList.toggle('open');
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
            // 1. Gather worksheet context for the selected chips
            const wsContext = await WorksheetContext.gather([..._activeContexts]);

            // 2. Send to LLM, receive code segments
            const segments  = await LLMClient.sendMessage(text, wsContext);

            _showTyping(false);

            // 3. Show agent acknowledgement
            // PLACEHOLDER: use the LLM's prose reply when the API returns one
            _appendMessage('agent', `Got it! Applying ${segments.length} change(s) to your sheet…`);

            // 4. Execute segments with pauses so user can observe each step
            await ExecutionEngine.run(segments);

        } catch (err) {
            _showTyping(false);
            _appendMessage('agent', `⚠️ Something went wrong: ${err.message}`);
            console.error('[ChatManager] Error:', err);
        }

        _setBusy(false);
    }

    function _appendMessage(role, text) {
        const wrapper = document.createElement('div');
        wrapper.className = `message ${role}`;

        const bubble = document.createElement('div');
        bubble.className = 'message-bubble';
        bubble.textContent = text;

        const meta = document.createElement('span');
        meta.className = 'message-meta';
        meta.textContent = role === 'user' ? 'You' : 'Assistant';

        wrapper.appendChild(bubble);
        wrapper.appendChild(meta);

        // Insert before typing indicator so it stays at the bottom
        _feed.insertBefore(wrapper, _typingEl);
        _feed.scrollTop = _feed.scrollHeight;
    }

    function _showTyping(visible) {
        _typingEl.classList.toggle('visible', visible);
        _feed.scrollTop = _feed.scrollHeight;
    }

    function _setBusy(busy) {
        _isBusy           = busy;
        _sendBtn.disabled = busy || _input.value.trim() === '';
        _input.disabled   = busy;
    }

    function _clearInput() {
        _input.value        = '';
        _input.style.height = 'auto';
    }

    return { init };
})();
