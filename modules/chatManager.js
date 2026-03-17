/**
 * chatManager.js — chat UI + orchestration pipeline.
 * Now wires into DagStore to record every segment chain as a DAG.
 */
const ChatManager = (() => {

    const _feed       = document.getElementById('message-feed');
    const _input      = document.getElementById('chat-input');
    const _sendBtn    = document.getElementById('send-button');
    const _typing     = document.getElementById('typing-indicator');
    const _typingLbl  = document.getElementById('typing-label');

    let _isBusy = false;

    // Track which DAG node the worksheet is currently at
    let _currentNodeId = null;
    // Track the edgeIds for the current running chain so we can mark them executed
    let _currentEdgeIds = [];

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

        // Journey button in header
        document.getElementById('journey-btn')?.addEventListener('click', () => {
            JourneyPanel.toggle(_currentNodeId);
        });

        document.getElementById('execution-header')?.addEventListener('click', () => {
            const panel = document.getElementById('execution-panel');
            const icon  = document.getElementById('execution-toggle-icon');
            panel.classList.toggle('open');
            icon.textContent = panel.classList.contains('open') ? '▼' : '▲';
        });

        JourneyPanel.init();
    }

    async function _handleSend() {
        const text = _input.value.trim();
        if (!text || _isBusy) return;
        _setBusy(true);
        _appendMessage('user', text);
        _clearInput();

        try {
            const wsCtx = await WorksheetContext.gather(['selection','sheet']);

            // 1. Scaffold rubric
            _showTyping(true, 'Creating requirements');
            let rubric = null;
            try {
                rubric = await LLMClient.rubricScaffold(text, wsCtx);
                StepNavigator.setRubric(rubric);
            } catch(e) { console.warn('[ChatManager] rubric scaffold failed:', e.message); }

            // 2. Generate code segments
            _showTyping(true, 'Generating a solution');
            const segments = await LLMClient.generateCode(text, wsCtx, rubric);
            _showTyping(false);

            // 3. Record new chain in DAG (branch from current node if we have one)
            const { rootNodeId, nodeIds, edgeIds } = DagStore.addChain(
                text, segments, _currentNodeId || null
            );
            // Start position is the root node of this chain
            if (!_currentNodeId) _currentNodeId = rootNodeId;
            _currentEdgeIds = edgeIds;
            StepNavigator.setCurrentDagNode(_currentNodeId);

            // 4. Load segments + show rubric gate
            StepNavigator.loadSegments(segments);
            await StepNavigator.showRubricGate();

            // 5. Execute — pass DAG edge tracking callbacks
            _appendMessage('agent', `Applying ${segments.length} step(s) to your sheet…`);
            await ExecutionEngine.run(segments, {
                onStepDone(index) {
                    const edgeId = _currentEdgeIds[index];
                    if (edgeId) {
                        DagStore.markEdgeExecuted(edgeId, false);
                        _currentNodeId = nodeIds[index + 1];
                        StepNavigator.setCurrentDagNode(_currentNodeId);
                        JourneyPanel.refresh(_currentNodeId);
                    }
                },
                onStepFailed(index) {
                    const edgeId = _currentEdgeIds[index];
                    if (edgeId) DagStore.markEdgeExecuted(edgeId, true);
                },
            });

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

    function _showTyping(v, label = '') {
        _typing.classList.toggle('visible', v);
        if (_typingLbl) _typingLbl.textContent = v && label ? label : '';
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

    /** Called by StepNavigator after an edit produces a new chain. */
    function onEditChain(newSegments, fromEdgeIndex) {
        // branchFrom current node at the edit point
        const { nodeIds, edgeIds } = DagStore.branchFrom(
            _currentNodeId, '(edit)', newSegments
        );
        _currentEdgeIds = edgeIds;
        JourneyPanel.refresh(_currentNodeId);
        return { nodeIds, edgeIds };
    }

    return { init, onEditChain };
})();
