/**
 * chatManager.js — chat UI + orchestration pipeline.
 */
const ChatManager = (() => {

    const _feed      = document.getElementById('message-feed');
    const _input     = document.getElementById('chat-input');
    const _sendBtn   = document.getElementById('send-button');
    const _typing    = document.getElementById('typing-indicator');
    const _typingLbl = document.getElementById('typing-label');

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
            icon.textContent = panel.classList.contains('open') ? '▼' : '▲';
        });
    }

    async function _handleSend() {
        const text = _input.value.trim();
        if (!text || _isBusy) return;
        _setBusy(true);
        _appendMessage('user', text);
        _clearInput();

        try {
            const wsCtx = await WorksheetContext.gather(['selection', 'sheet']);

            // 1 & 2. Fire rubric scaffold in parallel — user sees the gate
            //        immediately while the LLM generates requirements in the background.
            _showTyping(true, 'Creating requirements…');
            const rubricPromise = LLMClient.rubricScaffold(text, wsCtx).catch(e => {
                console.warn('[ChatManager] rubric scaffold failed:', e.message);
                return null;
            });

            // Show gate right away with empty rubric; update when ready
            RubricManager.showRubricGate();
            rubricPromise.then(rubric => {
                if (rubric) RubricManager.setRubric(rubric);
            });

            // Wait for user to click Start in the gate
            await RubricManager.waitForGate();

            // 3. Generate code segments (no rubric passed — server no longer needs it)
            _showTyping(true, 'Generating a solution…');
            const segments = await LLMClient.generateCode(text, wsCtx);
            _showTyping(false);

            // 4. Register chain and show Start button
            const chain = DagRunner.prepareChain(text, segments);

            _appendMessage('agent',
                `Ready — ${segments.length} step${segments.length !== 1 ? 's' : ''} planned.`,
                { actions: [{ label: '▶ Apply to sheet', primary: true, onClick: async (btn) => {
                    if (_isBusy) return;
                    btn.disabled    = true;
                    btn.textContent = 'Starting…';
                    _setBusy(true);
                    try {
                        _appendMessage('agent', `Applying ${segments.length} step(s)…`);
                        await DagRunner.start(chain.chainId);
                    } catch (err) {
                        _appendMessage('agent', `⚠️ ${err.message}`);
                    } finally {
                        _setBusy(false);
                        btn.textContent = '▶ Apply to sheet';
                        btn.disabled    = false;
                    }
                }}]}
            );

        } catch (err) {
            _showTyping(false);
            _appendMessage('agent', `⚠️ ${err.message}`);
        }

        _setBusy(false);
    }

    function _appendMessage(role, text, options = {}) {
        const w = document.createElement('div');
        w.className = `message ${role}`;
        const b = document.createElement('div');
        b.className   = 'message-bubble';
        b.textContent = text;

        const actions = Array.isArray(options.actions) ? options.actions : [];
        if (actions.length) {
            const row = document.createElement('div');
            row.className = 'message-actions';
            actions.forEach(a => {
                const btn = document.createElement('button');
                btn.className   = `message-action-btn${a.primary ? ' primary' : ''}`;
                btn.textContent = a.label || 'Action';
                btn.addEventListener('click', () => a.onClick?.(btn));
                row.appendChild(btn);
            });
            b.appendChild(row);
        }

        const m = document.createElement('span');
        m.className   = 'message-meta';
        m.textContent = role === 'user' ? 'You' : 'Assistant';
        w.appendChild(b); w.appendChild(m);
        _feed.insertBefore(w, _typing);
        _feed.scrollTop = _feed.scrollHeight;
    }

    function _showTyping(v, label = '') {
        _typing.classList.toggle('visible', v);
        if (_typingLbl) _typingLbl.textContent = v && label ? label : '';
        _feed.scrollTop = _feed.scrollHeight;
    }

    function _setBusy(b) {
        _isBusy       = b;
        _sendBtn.disabled = b || _input.value.trim() === '';
        _input.disabled   = b;
    }

    function _clearInput() {
        _input.value = ''; _input.style.height = 'auto';
    }

    return { init };
})();
