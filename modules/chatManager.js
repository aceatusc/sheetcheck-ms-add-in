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

        document.getElementById('verify-btn')?.addEventListener('click', () => {
            AspectManager.open();
        });

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

        ChatHistory.push(text);

        // Reveal the Specifications button on the first prompt, populate aspects
        // in the background, then open the panel once they're ready.
        const verifyBtn = document.getElementById('verify-btn');
        if (verifyBtn?.hidden) {
            verifyBtn.hidden = false;
            AspectManager.populate().then(() => AspectManager.open());
        }

        try {
            const wsCtx = await WorksheetContext.gather(['selection', 'sheet', 'styles', 'charts']);
            const runStore = DagStore.create();

            _showTyping(true, 'Generating a solution…');
            const segments = await LLMClient.generateCode(text, wsCtx, null, ChatHistory.get());
            _showTyping(false);

            const chain = DagRunner.prepareChain(text, segments, runStore);
            _appendSegmentMessage(segments, chain.chainId);

        } catch (err) {
            _showTyping(false);
        }

        _setBusy(false);
    }

    function _appendSegmentMessage(segments, chainId) {
        const w = document.createElement('div');
        w.className = 'message agent';
        const b = document.createElement('div');
        b.className = 'message-bubble';

        // Intro line
        const intro = document.createElement('p');
        intro.className   = 'segment-msg-intro';
        intro.textContent = `I'm ready to walk you through these ${segments.length} change${segments.length !== 1 ? 's' : ''} step-by-step:`;
        b.appendChild(intro);

        // Step list: bold description + muted explanation
        const list = document.createElement('div');
        list.className = 'segment-msg-list';
        segments.forEach((seg, i) => {
            const item = document.createElement('div');
            item.className = 'segment-msg-item';

            const desc = document.createElement('div');
            desc.className   = 'segment-msg-desc';
            desc.textContent = `${i + 1}. ${seg.description}`;

            const expl = document.createElement('div');
            expl.className   = 'segment-msg-expl';
            expl.textContent = seg.explanation;

            item.appendChild(desc);
            item.appendChild(expl);
            list.appendChild(item);
        });
        b.appendChild(list);

        // Apply button first
        const actions = document.createElement('div');
        actions.className = 'message-actions';
        const btn = document.createElement('button');
        btn.className   = 'message-action-btn primary';
        btn.textContent = '▶ Begin Tour';
        btn.addEventListener('click', async () => {
            btn.disabled    = true;
            btn.textContent = 'Starting…';
            try {
                await DagRunner.start(chainId);
            } catch (_) {
            } finally {
                btn.textContent = '▶ Begin Tour';
                btn.disabled    = false;
            }
        });
        actions.appendChild(btn);
        b.appendChild(actions);

        const meta = document.createElement('span');
        meta.className   = 'message-meta';
        meta.textContent = 'Assistant';
        w.appendChild(b);
        w.appendChild(meta);
        _feed.insertBefore(w, _typing);
        _feed.scrollTop = _feed.scrollHeight;
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
