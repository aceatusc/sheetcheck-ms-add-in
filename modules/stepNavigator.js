/**
 * stepNavigator.js
 * Step navigation UI: graph, rubric, ask, edit, chat dimming.
 */
const StepNavigator = (() => {

    // ── DOM refs ──────────────────────────────────────────────────────────────
    const _overlay      = document.getElementById('step-navigator');
    const _badge        = document.getElementById('step-nav-badge');
    const _ranges       = document.getElementById('step-nav-ranges');
    const _desc         = document.getElementById('step-nav-description');
    const _expl         = document.getElementById('step-nav-explanation');
    const _counter      = document.getElementById('step-nav-counter');
    const _btnPrev      = document.getElementById('step-nav-prev');
    const _btnNext      = document.getElementById('step-nav-next');
    const _btnClose     = document.getElementById('step-nav-close');
    const _btnEdit      = document.getElementById('step-nav-edit');
    const _btnAsk       = document.getElementById('step-nav-ask');
    const _btnVerify    = document.getElementById('step-nav-verify');
    const _graphEl      = document.getElementById('step-nav-graph');
    const _askPanel     = document.getElementById('step-nav-ask-panel');
    const _askInput     = document.getElementById('step-nav-ask-input');
    const _askSend      = document.getElementById('step-nav-ask-send');
    const _askThread    = document.getElementById('step-nav-ask-thread');
    const _askChips     = document.getElementById('step-nav-ask-chips');
    const _editPanel    = document.getElementById('step-nav-edit-panel');
    const _editFeedback = document.getElementById('step-nav-edit-feedback');
    const _editSend     = document.getElementById('step-nav-edit-send');
    const _rubricPanel  = document.getElementById('step-nav-rubric');
    const _rubricHard   = document.getElementById('rubric-hard-list');
    const _rubricSoft   = document.getElementById('rubric-soft-list');
    const _rubricAdd    = document.getElementById('rubric-add-btn');
    const _rubricVerifyPanel = document.getElementById('rubric-verify-results');
    const _chatPanel    = document.getElementById('chat-panel');
    const _qaList       = document.getElementById('step-nav-qa-list');

    // ── State ─────────────────────────────────────────────────────────────────
    let _segments       = [];
    let _dagMeta        = null; // { chainId, rootNodeId, nodeIds[], edgeIds[], taskLabel }
    let _completedUpTo  = -1;
    let _currentIndex   = 0;
    let _isRunning      = false;
    let _advanceResolve = null;
    let _askHistory     = [];
    let _rubric         = { hard_requirements: [], soft_requirements: [] };
    let _activePanel    = null; // 'ask' | 'edit' | 'rubric' | null
    let _isRubricGate   = false;
    let _isVerifyGate   = false; // true while showing verification results screen
    let _failedIndices  = new Set(); // indices of segments that errored
    // Edit fork history: [{fromIndex, segments}] — old paths replaced by edits, shown dimmed
    let _branches       = [];

    // ── Public ────────────────────────────────────────────────────────────────

    function init() {
        _btnPrev.addEventListener('click', _onPrev);
        _btnNext.addEventListener('click', _onNext);
        _btnClose.addEventListener('click', dismiss);
        _btnAsk.addEventListener('click', () => _togglePanel('ask'));
        _btnEdit.addEventListener('click', () => _togglePanel('edit'));

        _askSend.addEventListener('click', _onAskSend);
        _askInput.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); _onAskSend(); }});

        _editSend.addEventListener('click', _onEditSend);
        _rubricAdd.addEventListener('click', _onRubricAdd);
    }

    function loadSegments(segments, dagMeta = null) {
        _segments       = segments;
        _dagMeta        = dagMeta;
        _completedUpTo  = -1;
        _currentIndex   = 0;
        _isRunning      = false;
        _advanceResolve = null;
        _askHistory     = [];
        _failedIndices  = new Set();
        _isVerifyGate   = false;
        _branches       = [];
        _render();
    }

    function setRubric(rubric) {
        _rubric = rubric;
        _renderRubric();
    }

    /**
     * Show rubric gate screen — navigator visible, rubric panel open,
     * Next button says "Start →". Resolves when user clicks Start.
     * chatManager awaits this before calling ExecutionEngine.run().
     */
    function showRubricGate() {
        _isRubricGate = true;
        _overlay.classList.add('visible');
        _chatPanel.classList.add('nav-active');
        // Show rubric panel, hide step content
        _overlay.classList.add('rubric-gate');
        _togglePanel('rubric');
        _renderRubric();
        // Override Next button for gate
        _btnNext.textContent = 'Start →';
        _btnNext.disabled    = false;
        _btnPrev.disabled    = true;
        _btnEdit.disabled    = true;
        _btnAsk.disabled     = true;
        _badge.textContent   = 'Review Requirements';
        _counter.textContent = '';
        _ranges.innerHTML    = '';
        _desc.textContent    = 'Review and edit the requirements below, then click Start to begin.';
        _expl.textContent    = '';
        _qaList.innerHTML    = '';

        return new Promise(resolve => { _advanceResolve = resolve; });
    }

    /**
     * Called by ExecutionEngine after all segments complete.
     * Shows rubric verification results and waits for the user to click "Done ✓".
     * @returns {Promise<void>} resolves when user dismisses the screen
     */
    async function showVerifyResults() {
        // Switch into verify-gate mode: reuse the overlay as a results card
        _isVerifyGate = true;
        _overlay.classList.remove('failed', 'running');
        _overlay.classList.add('visible', 'verify-gate');
        _chatPanel.classList.add('nav-active');

        // Configure the fixed UI elements
        _badge.textContent   = 'Requirements Check';
        _counter.textContent = '';
        _desc.textContent    = 'Verifying your spreadsheet against the requirements…';
        _expl.textContent    = '';
        _ranges.innerHTML    = '';
        _qaList.innerHTML    = '';
        _btnPrev.disabled    = true;
        _btnNext.textContent = '…';
        _btnNext.disabled    = true;
        _btnEdit.disabled    = true;
        _btnAsk.disabled     = true;

        // Open the rubric panel to show results inside it
        _togglePanel('rubric');
        _rubricVerifyPanel.innerHTML =
            '<span style="color:rgba(255,255,255,0.6);font-size:11px">Verifying requirements…</span>';

        try {
            const wsCtx = await WorksheetContext.gather(['sheet']);
            const res   = await LLMClient.rubricVerify(_rubric, wsCtx);

            _rubricVerifyPanel.innerHTML =
                '<div class="rubric-section-label" style="margin-top:8px">Verification Results</div>';

            const allItems = [
                ...(_rubric.hard_requirements || []),
                ...(_rubric.soft_requirements || []),
            ];
            let metCount = 0;

            (res.results || []).forEach(r => {
                const item = allItems.find(x => x.id === r.id);
                if (!item) return;
                const type = (_rubric.hard_requirements || []).find(x => x.id === r.id) ? 'hard' : 'soft';
                if (r.met) metCount++;

                const row = document.createElement('div');
                row.className = 'rubric-row';
                const icon = r.met
                    ? `<span class="rubric-check">✓</span>`
                    : `<span class="rubric-warn">⚠</span>`;
                row.innerHTML = `
                    <span class="rubric-badge rubric-${type}">${type === 'hard' ? 'H' : 'S'}</span>
                    <span class="rubric-label" style="user-select:none">${item.label}</span>
                    ${icon}`;

                // Clicking the icon expands reasoning + references
                const iconEl = row.querySelector('.rubric-check, .rubric-warn');
                if (iconEl) {
                    iconEl.style.cursor = 'pointer';
                    iconEl.title = 'Click to expand reasoning';
                    let open = false;
                    iconEl.addEventListener('click', () => {
                        open = !open;
                        let detail = row.querySelector('.rubric-detail');
                        if (!detail) {
                            detail = document.createElement('div');
                            detail.className = 'rubric-detail';
                            row.appendChild(detail);
                        }
                        detail.textContent = r.reasoning +
                            (r.references?.length ? ` (${r.references.join(', ')})` : '');
                        detail.style.display = open ? 'block' : 'none';
                    });
                }
                _rubricVerifyPanel.appendChild(row);
            });

            // Summary line
            const total = res.results?.length || 0;
            _desc.textContent = total
                ? `${metCount} of ${total} requirement${total !== 1 ? 's' : ''} met`
                : 'No requirements to verify.';

        } catch (err) {
            _rubricVerifyPanel.innerHTML =
                `<span style="color:var(--color-error);font-size:11px">Verification failed: ${err.message}</span>`;
            _desc.textContent = 'Could not verify requirements.';
        }

        // Enable Done button and wait for user click
        _btnNext.textContent = 'Done ✓';
        _btnNext.disabled    = false;

        return new Promise(resolve => { _advanceResolve = resolve; });
    }

    function markRunning(index) {
        _isRunning    = true;
        _currentIndex = index;
        _overlay.classList.add('running', 'visible');
        _chatPanel.classList.add('nav-active');
        _render();
    }

    async function waitForNext(index) {
        _isRunning     = false;
        _completedUpTo = Math.max(_completedUpTo, index);
        _currentIndex  = index;
        _askHistory    = [];
        _overlay.classList.remove('running');
        _overlay.classList.add('visible');
        _chatPanel.classList.add('nav-active');

        const promise = new Promise(resolve => { _advanceResolve = resolve; });
        _render();
        await _focusRanges(index);
        return promise;
    }

    /**
     * Called after a segment fails. Shows yellow error card and still
     * waits for the user to click → so they can navigate freely.
     * @returns {Promise<void>}
     */
    async function markFailed(index, errorMsg) {
        _isRunning     = false;
        _failedIndices.add(index);
        // Treat failed step as "reached" so user can navigate back to it
        _completedUpTo = Math.max(_completedUpTo, index);
        _currentIndex  = index;
        _askHistory    = [];
        _overlay.classList.remove('running');
        _overlay.classList.add('visible', 'failed');
        _chatPanel.classList.add('nav-active');

        // Temporarily show error in explanation area
        const seg = _segments[index];
        if (seg) seg._errorMsg = errorMsg;

        const promise = new Promise(resolve => { _advanceResolve = resolve; });
        _render();
        return promise;
    }

    function dismiss() {
        _overlay.classList.remove('visible', 'running', 'failed');
        _chatPanel.classList.remove('nav-active');
        _closePanel();
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    function _onNext() {
        if (_isRunning) return;

        // Verify gate: "Done ✓" click — dismiss and resolve
        if (_isVerifyGate) {
            _isVerifyGate = false;
            _overlay.classList.remove('verify-gate');
            dismiss();
            if (_advanceResolve) { const r = _advanceResolve; _advanceResolve = null; r(); }
            return;
        }

        // Rubric gate: "Start →" click
        if (_isRubricGate) {
            _isRubricGate = false;
            _overlay.classList.remove('rubric-gate');
            _closePanel();
            _overlay.classList.remove('visible');
            _chatPanel.classList.remove('nav-active');
            if (_advanceResolve) { const r = _advanceResolve; _advanceResolve = null; r(); }
            return;
        }

        if (_currentIndex < _completedUpTo) {
            _overlay.classList.remove('failed');
            _navigate(_currentIndex + 1);
            return;
        }
        if (_advanceResolve) {
            const resolve = _advanceResolve;
            _advanceResolve = null;
            const isLast = _currentIndex >= _segments.length - 1;
            if (isLast) dismiss();
            else _overlay.classList.remove('failed');
            resolve();
        }
    }

    function _onPrev() {
        if (_isRunning || _currentIndex <= 0) return;
        _overlay.classList.remove('failed');
        _navigate(_currentIndex - 1);
    }

    async function _navigate(targetIndex) {
        // Allow visiting any step that has been reached (completed or failed)
        if (targetIndex < 0 || (targetIndex > _completedUpTo && !_failedIndices.has(targetIndex))) return;
        _currentIndex  = targetIndex;
        _askHistory    = [];
        _closePanel();
        if (_failedIndices.has(targetIndex)) {
            _overlay.classList.add('failed');
        } else {
            _overlay.classList.remove('failed');
        }
        _render();
        await _focusRanges(targetIndex);
    }

    // Navigate to a segment by its id — called from graph edge click
    async function _navigateById(segId) {
        const idx = _segments.findIndex(s => s.id === segId);
        if (idx < 0) return;
        if (idx > _completedUpTo && !_failedIndices.has(idx)) return;
        await _navigate(idx);
    }

    // ── Panels ────────────────────────────────────────────────────────────────

    function _togglePanel(name) {
        if (_activePanel === name) { _closePanel(); return; }
        _activePanel = name;
        _askPanel.style.display   = name === 'ask'  ? 'flex' : 'none';
        _editPanel.style.display  = name === 'edit' ? 'flex' : 'none';
        _rubricPanel.style.display= name === 'rubric' ? 'block' : 'none';
        if (name === 'rubric') _renderRubric();
    }

    function _closePanel() {
        _activePanel = null;
        _askPanel.style.display   = 'none';
        _editPanel.style.display  = 'none';
        _rubricPanel.style.display= 'none';
    }

    // ── Ask feature ───────────────────────────────────────────────────────────

    async function _onAskSend() {
        const msg = _askInput.value.trim();
        if (!msg) return;
        _askInput.value = '';
        _askInput.disabled = true;
        _askSend.disabled  = true;
        _appendAskBubble('user', msg);

        const seg = _segments[_currentIndex];
        try {
            const wsCtx = await WorksheetContext.gather(['sheet']);
            const res   = await LLMClient.ask(msg, wsCtx,
                { description: seg.description, explanation: seg.explanation }, _askHistory);
            _askHistory.push({ q: msg, a: res.answer });
            _appendAskBubble('agent', res.answer);
            _renderAskChips(res.follow_up_questions || []);
        } catch(err) {
            _appendAskBubble('agent', `⚠️ ${err.message}`);
        } finally {
            _askInput.disabled = false;
            _askSend.disabled  = false;
        }
    }

    function _appendAskBubble(role, text) {
        const d = document.createElement('div');
        d.className = `ask-bubble ask-${role}`;
        d.textContent = text;
        _askThread.appendChild(d);
        _askThread.scrollTop = _askThread.scrollHeight;
    }

    function _renderAskChips(questions) {
        _askChips.innerHTML = '';
        questions.forEach(q => {
            const btn = document.createElement('button');
            btn.className = 'ask-chip';
            btn.textContent = q;
            btn.onclick = () => { _askInput.value = q; _onAskSend(); };
            _askChips.appendChild(btn);
        });
    }

    async function _onEditSend() {
        const msg            = _editFeedback.value.trim();
        const seg            = _segments[_currentIndex];
        const remaining      = _segments.slice(_currentIndex + 1);  // tail after edit point
        _editSend.disabled   = true;
        _editSend.textContent = '…';
        try {
            const wsCtx  = await WorksheetContext.gather(['sheet']);
            // /edit now returns an array: the edited step + regenerated remainder
            const newChain = await LLMClient.edit(msg, wsCtx, seg, remaining);
            _editFeedback.value = '';

            // Save the old tail as a visual branch before replacing it
            if (remaining.length > 0) {
                _branches.push({ fromIndex: _currentIndex, segments: remaining });
            }

            // Update DAG and splice new chain into _segments from current index onward
            if (_dagMeta?.chainId) {
                const updated = DagRunner.applyEdit(_dagMeta.chainId, _currentIndex, msg, newChain);
                _segments = updated.segments;
                _dagMeta = {
                    chainId: updated.chainId,
                    rootNodeId: updated.rootNodeId,
                    nodeIds: updated.nodeIds,
                    edgeIds: updated.edgeIds,
                    taskLabel: updated.taskLabel,
                };
            } else {
                _segments = [..._segments.slice(0, _currentIndex), ...newChain];
            }

            // Mark completed state: only steps before edit point still counted
            _completedUpTo = Math.min(_completedUpTo, _currentIndex - 1);
            // Clear failure flags on spliced-out indices
            _failedIndices = new Set([..._failedIndices].filter(i => i < _currentIndex));

            // Execute the first segment of the new chain (the edited step itself)
            const firstSeg = _segments[_currentIndex];
            if (firstSeg?.code) {
                try {
                    const fn = new (Object.getPrototypeOf(async function(){}).constructor)(firstSeg.code);
                    await fn();
                    _completedUpTo = Math.max(_completedUpTo, _currentIndex);
                    const edgeId = _dagMeta?.edgeIds?.[_currentIndex];
                    if (edgeId) DagStore.markEdgeExecuted(edgeId, false);
                    _editSend.textContent = '✓ Applied';
                } catch (execErr) {
                    _editSend.textContent = '⚠ Run failed';
                    const edgeId = _dagMeta?.edgeIds?.[_currentIndex];
                    if (edgeId) DagStore.markEdgeExecuted(edgeId, true);
                    console.error('[StepNavigator] edit run error:', execErr.message);
                }
            }

            _render();
        } catch(err) {
            _editSend.textContent = '⚠ Error';
            console.error('[StepNavigator] edit LLM error:', err);
        } finally {
            setTimeout(() => {
                _editSend.textContent = 'Apply Edit';
                _editSend.disabled = false;
            }, 1800);
        }
    }

    // ── Rubric feature ────────────────────────────────────────────────────────

    // Drag state
    let _dragId   = null;  // id of item being dragged
    let _dragType = null;  // 'hard' | 'soft'

    function _renderRubric() {
        _rubricHard.innerHTML = '';
        _rubricSoft.innerHTML = '';
        const hard = _rubric.hard_requirements || [];
        const soft = _rubric.soft_requirements || [];
        hard.forEach(r => _appendRubricRow(r, 'hard'));
        soft.forEach(r => _appendRubricRow(r, 'soft'));
        _setupDropZone(_rubricHard, 'hard');
        _setupDropZone(_rubricSoft, 'soft');

        // Show hint when both lists are empty
        if (hard.length === 0 && soft.length === 0) {
            const hint = document.createElement('div');
            hint.style.cssText = 'font-size:11px;color:rgba(255,255,255,0.4);padding:8px 2px;text-align:center';
            hint.textContent = 'No requirements yet — click "+ Add requirement" below.';
            _rubricSoft.appendChild(hint);
        }
    }

    function _appendRubricRow(req, type, verifyResult = null) {
        const list = type === 'hard' ? _rubricHard : _rubricSoft;
        const row  = document.createElement('div');
        row.className = 'rubric-row';
        row.dataset.id   = req.id;
        row.dataset.type = type;
        row.draggable = true;

        // Drag handle
        const handle = document.createElement('span');
        handle.className = 'rubric-drag';
        handle.textContent = '⠿';
        handle.title = 'Drag to reorder or move between Hard/Soft';

        // Badge
        const badge = document.createElement('span');
        badge.className = `rubric-badge rubric-${type}`;
        badge.textContent = type === 'hard' ? 'H' : 'S';

        // Inline editable label
        const labelEl = document.createElement('span');
        labelEl.className = 'rubric-label';
        labelEl.contentEditable = 'true';
        labelEl.spellcheck = false;
        labelEl.textContent = req.label;
        labelEl.addEventListener('blur', () => {
            const lst = type === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
            const item = lst.find(x => x.id === req.id);
            if (item) item.label = labelEl.textContent.trim() || item.label;
        });
        // Prevent Enter from creating newlines
        labelEl.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); labelEl.blur(); }});

        // Move button (H↔S)
        const moveBtn = document.createElement('button');
        moveBtn.className = 'rubric-move';
        moveBtn.title = `Move to ${type === 'hard' ? 'Soft' : 'Hard'}`;
        moveBtn.textContent = type === 'hard' ? '↓S' : '↑H';
        moveBtn.onclick = () => _moveRubricItem(req.id, type);

        // Delete button
        const delBtn = document.createElement('button');
        delBtn.className = 'rubric-del';
        delBtn.title = 'Remove';
        delBtn.textContent = '✕';
        delBtn.onclick = () => _deleteRubricItem(req.id, type);

        row.appendChild(handle);
        row.appendChild(badge);
        row.appendChild(labelEl);

        // Verify result icon if present
        if (verifyResult !== null) {
            const icon = document.createElement('span');
            icon.className = verifyResult.met ? 'rubric-check' : 'rubric-warn';
            icon.textContent = verifyResult.met ? '✓' : '⚠';
            icon.title = verifyResult.reasoning;
            icon.style.cursor = 'pointer';
            let open = false;
            icon.addEventListener('click', () => {
                open = !open;
                let detail = row.querySelector('.rubric-detail');
                if (!detail) {
                    detail = document.createElement('div');
                    detail.className = 'rubric-detail';
                    row.appendChild(detail);
                }
                detail.textContent = verifyResult.reasoning
                    + (verifyResult.references?.length ? ` (${verifyResult.references.join(', ')})` : '');
                detail.style.display = open ? 'block' : 'none';
            });
            row.appendChild(icon);
        }

        row.appendChild(moveBtn);
        row.appendChild(delBtn);

        // ── Drag events ──
        row.addEventListener('dragstart', e => {
            _dragId   = req.id;
            _dragType = type;
            row.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        });
        row.addEventListener('dragend', () => {
            row.classList.remove('dragging');
            document.querySelectorAll('.rubric-drop-over').forEach(el => el.classList.remove('rubric-drop-over'));
        });
        row.addEventListener('dragover', e => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            row.classList.add('rubric-drop-over');
        });
        row.addEventListener('dragleave', () => row.classList.remove('rubric-drop-over'));
        row.addEventListener('drop', e => {
            e.preventDefault();
            row.classList.remove('rubric-drop-over');
            if (!_dragId || _dragId === req.id) return;
            _dropItem(_dragId, _dragType, req.id, type);
        });

        list.appendChild(row);
    }

    /** Set up drop zone on the list container itself (for dropping at end or into empty list) */
    function _setupDropZone(listEl, targetType) {
        listEl.addEventListener('dragover', e => {
            // Only handle if hovering directly on the list (not a child row)
            if (e.target === listEl) { e.preventDefault(); listEl.classList.add('rubric-list-over'); }
        });
        listEl.addEventListener('dragleave', e => {
            if (e.target === listEl) listEl.classList.remove('rubric-list-over');
        });
        listEl.addEventListener('drop', e => {
            if (e.target !== listEl) return;
            e.preventDefault();
            listEl.classList.remove('rubric-list-over');
            if (!_dragId) return;
            // Move dragged item to end of targetType list
            _moveRubricToType(_dragId, _dragType, targetType);
        });
    }

    /** Reorder or cross-list drop: move dragId to just before targetId (possibly changing type) */
    function _dropItem(dragId, fromType, targetId, targetType) {
        const fromList = fromType === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const toList   = targetType === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;

        const fromIdx = fromList.findIndex(r => r.id === dragId);
        if (fromIdx < 0) return;
        const [item] = fromList.splice(fromIdx, 1);

        // If crossing lists, push to toList before target
        const toIdx = toList.findIndex(r => r.id === targetId);
        if (toIdx >= 0) toList.splice(toIdx, 0, item);
        else toList.push(item);

        _renderRubric();
    }

    function _moveRubricToType(dragId, fromType, toType) {
        if (fromType === toType) return;
        const fromList = fromType === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const toList   = toType   === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const idx = fromList.findIndex(r => r.id === dragId);
        if (idx < 0) return;
        const [item] = fromList.splice(idx, 1);
        toList.push(item);
        _renderRubric();
    }

    function _moveRubricItem(id, fromType) {
        const toType = fromType === 'hard' ? 'soft' : 'hard';
        _moveRubricToType(id, fromType, toType);
    }

    function _deleteRubricItem(id, type) {
        const list = type === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const idx  = list.findIndex(r => r.id === id);
        if (idx >= 0) list.splice(idx, 1);
        _renderRubric();
    }

    function _onRubricAdd() {
        // Add blank entry to soft requirements and re-render — label is editable inline
        const id = 'u' + Date.now();
        _rubric.soft_requirements.push({ id, label: '', checked: false });
        _renderRubric();
        // Focus the newly added label so user can type immediately
        requestAnimationFrame(() => {
            const newRow = _rubricSoft.querySelector(`[data-id="${id}"] .rubric-label`);
            if (newRow) { newRow.focus(); _selectAllText(newRow); }
        });
    }

    function _selectAllText(el) {
        const range = document.createRange();
        range.selectNodeContents(el);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
    }

    // ── Graph — edge-focused DAG with branch support ──────────────────────────
    //
    // Model: nodes are implicit worksheet states (indices 0..N), edges are
    // segments. "Viewing step i" means edge i (node i → node i+1) is active.
    //
    // Layout: main path is a horizontal row. Each edit fork is drawn as a
    // separate row below the branch point, connected by a curved path.
    // Clicking any reached edge navigates to it.

    function _renderGraph() {
        if (!_graphEl) return;
        _graphEl.innerHTML = '';

        if (_dagMeta?.chainId) {
            const rep = DagRunner.getGraphRepresentation(_dagMeta.chainId);
            _renderGraphFromDag(rep);
            return;
        }

        const total = _segments.length;
        if (total === 0) return;

        const SVG_NS  = 'http://www.w3.org/2000/svg';
        const NODE_R  = 5;   // state dot radius
        const EDGE_W  = 28;  // horizontal space per edge
        const ROW_H   = 26;  // vertical space between rows
        const PAD_X   = 10;
        const PAD_Y   = 10;

        // ── Build rows (legacy local-only rendering) ───────────────────────────
        // Row 0 = active path (_segments). Row 1..N = branches (oldest first).
        const rows = [];
        rows.push({ segs: _segments, fromIndex: -1 });
        _branches.forEach(b => rows.push({ segs: b.segments, fromIndex: b.fromIndex }));

        const maxEdges = Math.max(...rows.map(r => r.segs.length));
        // Node count for main row = total + 1; branches may be shorter
        const mainNodes = total + 1;
        const svgW = PAD_X * 2 + mainNodes * EDGE_W;
        const svgH = PAD_Y * 2 + rows.length * ROW_H;

        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.setAttribute('viewBox', `0 0 ${svgW} ${svgH}`);
        svg.setAttribute('width', svgW);
        svg.setAttribute('height', svgH);

        // x position of node i in the main row
        const nodeX = i => PAD_X + i * EDGE_W;
        // y centre of a row
        const rowY  = r => PAD_Y + r * ROW_H + ROW_H / 2;

        // ── Draw each row ─────────────────────────────────────────────────────
        rows.forEach((row, rowIdx) => {
            const isBranch = rowIdx > 0;
            const y        = rowY(rowIdx);
            const segs     = row.segs;
            const nodeCount = segs.length + 1;

            // For a branch row, start x offset matches where it forks from main
            const startNodeIndex = isBranch ? row.fromIndex + 1 : 0;

            // Draw state nodes
            for (let ni = 0; ni < nodeCount; ni++) {
                const absIdx = startNodeIndex + ni;
                const x      = nodeX(absIdx);
                const dot    = document.createElementNS(SVG_NS, 'circle');
                dot.setAttribute('cx', x);
                dot.setAttribute('cy', y);
                dot.setAttribute('r', NODE_R);
                dot.setAttribute('fill', isBranch
                    ? 'rgba(255,255,255,0.1)'
                    : (absIdx <= _completedUpTo + 1
                        ? 'rgba(255,255,255,0.35)'
                        : 'rgba(255,255,255,0.08)'));
                dot.setAttribute('stroke', isBranch
                    ? 'rgba(255,255,255,0.18)'
                    : 'rgba(255,255,255,0.4)');
                dot.setAttribute('stroke-width', '1');
                svg.appendChild(dot);
            }

            // Draw branch connector curve from main row to branch start
            if (isBranch) {
                const fx  = nodeX(row.fromIndex + 1);
                const fy  = rowY(0);
                const tx  = nodeX(startNodeIndex);
                const ty  = y;
                const mid = (fy + ty) / 2;
                const path = document.createElementNS(SVG_NS, 'path');
                path.setAttribute('d', `M${fx} ${fy} C${fx} ${mid} ${tx} ${mid} ${tx} ${ty}`);
                path.setAttribute('fill', 'none');
                path.setAttribute('stroke', 'rgba(255,255,255,0.12)');
                path.setAttribute('stroke-width', '1.5');
                path.setAttribute('stroke-dasharray', '3 2');
                svg.appendChild(path);
            }

            // Draw edges (one per segment in this row)
            segs.forEach((seg, edgeIdx) => {
                const x1 = nodeX(startNodeIndex + edgeIdx);
                const x2 = nodeX(startNodeIndex + edgeIdx + 1);
                const mx = (x1 + x2) / 2;

                const isActive   = !isBranch && edgeIdx === _currentIndex;
                const isExecuted = !isBranch && edgeIdx <= _completedUpTo;
                const isFailed   = !isBranch && _failedIndices.has(edgeIdx);
                const isRunning  = !isBranch && edgeIdx === _currentIndex && _isRunning;
                const isReachable = !isBranch && (isExecuted || isFailed);

                // Edge line
                const line = document.createElementNS(SVG_NS, 'line');
                line.setAttribute('x1', x1); line.setAttribute('y1', y);
                line.setAttribute('x2', x2); line.setAttribute('y2', y);

                let stroke, width, dash;
                if (isBranch) {
                    stroke = 'rgba(255,255,255,0.15)'; width = '1.5'; dash = '3 2';
                } else if (isFailed) {
                    stroke = '#f5643c'; width = isActive ? '3' : '2'; dash = 'none';
                } else if (isActive && isExecuted) {
                    stroke = '#fff'; width = '3'; dash = 'none';
                } else if (isExecuted) {
                    stroke = '#3ecf8e'; width = '2'; dash = 'none';
                } else if (isActive) {
                    stroke = '#4f8ef7'; width = '3'; dash = 'none';  // running/current unexecuted
                } else {
                    stroke = 'rgba(79,142,247,0.2)'; width = '1.5'; dash = '4 3';
                }

                line.setAttribute('stroke', stroke);
                line.setAttribute('stroke-width', width);
                if (dash !== 'none') line.setAttribute('stroke-dasharray', dash);
                svg.appendChild(line);

                // Running pulse on active edge
                if (isRunning) {
                    const pulse = document.createElementNS(SVG_NS, 'line');
                    pulse.setAttribute('x1', x1); pulse.setAttribute('y1', y);
                    pulse.setAttribute('x2', x2); pulse.setAttribute('y2', y);
                    pulse.setAttribute('stroke', 'rgba(79,142,247,0.5)');
                    pulse.setAttribute('stroke-width', '6');
                    const anim = document.createElementNS(SVG_NS, 'animate');
                    anim.setAttribute('attributeName', 'stroke-opacity');
                    anim.setAttribute('values', '0.5;0.1;0.5');
                    anim.setAttribute('dur', '1s');
                    anim.setAttribute('repeatCount', 'indefinite');
                    pulse.appendChild(anim);
                    svg.appendChild(pulse);
                }

                // Active edge glow highlight
                if (isActive && !isBranch) {
                    const glow = document.createElementNS(SVG_NS, 'line');
                    glow.setAttribute('x1', x1); glow.setAttribute('y1', y);
                    glow.setAttribute('x2', x2); glow.setAttribute('y2', y);
                    glow.setAttribute('stroke', isFailed ? 'rgba(245,100,60,0.3)' : 'rgba(255,255,255,0.2)');
                    glow.setAttribute('stroke-width', '8');
                    glow.setAttribute('stroke-linecap', 'round');
                    svg.insertBefore(glow, line); // insert behind the main line
                }

                // Invisible hit-target for clicks (wider than visible line)
                if (isReachable) {
                    const hit = document.createElementNS(SVG_NS, 'line');
                    hit.setAttribute('x1', x1); hit.setAttribute('y1', y);
                    hit.setAttribute('x2', x2); hit.setAttribute('y2', y);
                    hit.setAttribute('stroke', 'transparent');
                    hit.setAttribute('stroke-width', '14');
                    hit.style.cursor = 'pointer';
                    hit.addEventListener('click', () => _navigate(edgeIdx));
                    const title = document.createElementNS(SVG_NS, 'title');
                    title.textContent = `${edgeIdx + 1}. ${seg.description}`;
                    hit.appendChild(title);
                    svg.appendChild(hit);
                }

                // Step number label on active edge
                if (isActive && !isBranch) {
                    const label = document.createElementNS(SVG_NS, 'text');
                    label.setAttribute('x', mx);
                    label.setAttribute('y', y - NODE_R - 3);
                    label.setAttribute('text-anchor', 'middle');
                    label.setAttribute('font-size', '8');
                    label.setAttribute('fill', 'rgba(255,255,255,0.7)');
                    label.setAttribute('font-family', 'Segoe UI, system-ui, sans-serif');
                    label.textContent = `${edgeIdx + 1}`;
                    svg.appendChild(label);
                }
            });
        });

        _graphEl.innerHTML = '';
        _graphEl.style.overflowX = 'auto';
        _graphEl.appendChild(svg);
    }

    function _renderGraphFromDag(rep) {
        if (!_graphEl) return;
        _graphEl.innerHTML = '';

        const rows = Array.isArray(rep?.rows) ? rep.rows : [];
        if (rows.length === 0) return;

        const main = rows.find(r => r.kind === 'main') || rows[0];
        const total = (main.edgeIds || []).length;
        if (total === 0) return;

        const SVG_NS  = 'http://www.w3.org/2000/svg';
        const NODE_R  = 5;
        const EDGE_W  = 28;
        const ROW_H   = 26;
        const PAD_X   = 10;
        const PAD_Y   = 10;

        const mainNodes = total + 1;
        const svgW = PAD_X * 2 + mainNodes * EDGE_W;
        const svgH = PAD_Y * 2 + rows.length * ROW_H;

        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.setAttribute('viewBox', `0 0 ${svgW} ${svgH}`);
        svg.setAttribute('width', svgW);
        svg.setAttribute('height', svgH);

        const nodeX = i => PAD_X + i * EDGE_W;
        const rowY  = r => PAD_Y + r * ROW_H + ROW_H / 2;

        rows.forEach((row, rowIdx) => {
            const isBranch = row.kind !== 'main';
            const y = rowY(rowIdx);
            const edgeIds = row.edgeIds || [];
            const nodeCount = edgeIds.length + 1;
            const startNodeIndex = isBranch ? (row.fromMainNodeIndex + 1) : 0;

            // nodes
            for (let ni = 0; ni < nodeCount; ni++) {
                const absIdx = startNodeIndex + ni;
                const x = nodeX(absIdx);
                const dot = document.createElementNS(SVG_NS, 'circle');
                dot.setAttribute('cx', x);
                dot.setAttribute('cy', y);
                dot.setAttribute('r', NODE_R);
                dot.setAttribute('fill', isBranch
                    ? 'rgba(255,255,255,0.1)'
                    : (absIdx <= _completedUpTo + 1
                        ? 'rgba(255,255,255,0.35)'
                        : 'rgba(255,255,255,0.08)'));
                dot.setAttribute('stroke', isBranch ? 'rgba(255,255,255,0.18)' : 'rgba(255,255,255,0.4)');
                dot.setAttribute('stroke-width', '1');
                svg.appendChild(dot);
            }

            // connector
            if (isBranch) {
                const fx  = nodeX(row.fromMainNodeIndex + 1);
                const fy  = rowY(0);
                const tx  = nodeX(startNodeIndex);
                const ty  = y;
                const mid = (fy + ty) / 2;
                const path = document.createElementNS(SVG_NS, 'path');
                path.setAttribute('d', `M${fx} ${fy} C${fx} ${mid} ${tx} ${mid} ${tx} ${ty}`);
                path.setAttribute('fill', 'none');
                path.setAttribute('stroke', 'rgba(255,255,255,0.12)');
                path.setAttribute('stroke-width', '1.5');
                path.setAttribute('stroke-dasharray', '3 2');
                svg.appendChild(path);
            }

            // edges
            edgeIds.forEach((edgeId, edgeIdx) => {
                const edge = DagStore.getEdge(edgeId);
                const seg = edge?.segment || {};

                const x1 = nodeX(startNodeIndex + edgeIdx);
                const x2 = nodeX(startNodeIndex + edgeIdx + 1);

                const isActive = !isBranch && edgeIdx === _currentIndex;
                const isExecuted = !isBranch && (edge?.executed || edgeIdx <= _completedUpTo);
                const isFailed = !isBranch && (edge?.failed || _failedIndices.has(edgeIdx));
                const isRunning = !isBranch && edgeIdx === _currentIndex && _isRunning;
                const isReachable = !isBranch && (isExecuted || isFailed);

                const line = document.createElementNS(SVG_NS, 'line');
                line.setAttribute('x1', x1); line.setAttribute('y1', y);
                line.setAttribute('x2', x2); line.setAttribute('y2', y);

                let stroke, width, dash;
                if (isBranch) {
                    stroke = 'rgba(255,255,255,0.15)'; width = '1.5'; dash = '3 2';
                } else if (isFailed) {
                    stroke = '#f5643c'; width = isActive ? '3' : '2'; dash = 'none';
                } else if (isActive && isExecuted) {
                    stroke = '#fff'; width = '3'; dash = 'none';
                } else if (isExecuted) {
                    stroke = '#3ecf8e'; width = '2'; dash = 'none';
                } else if (isActive) {
                    stroke = '#4f8ef7'; width = '3'; dash = 'none';
                } else {
                    stroke = 'rgba(79,142,247,0.2)'; width = '1.5'; dash = '4 3';
                }

                line.setAttribute('stroke', stroke);
                line.setAttribute('stroke-width', width);
                if (dash !== 'none') line.setAttribute('stroke-dasharray', dash);
                svg.appendChild(line);

                if (isRunning) {
                    const pulse = document.createElementNS(SVG_NS, 'line');
                    pulse.setAttribute('x1', x1); pulse.setAttribute('y1', y);
                    pulse.setAttribute('x2', x2); pulse.setAttribute('y2', y);
                    pulse.setAttribute('stroke', 'rgba(79,142,247,0.5)');
                    pulse.setAttribute('stroke-width', '6');
                    const anim = document.createElementNS(SVG_NS, 'animate');
                    anim.setAttribute('attributeName', 'stroke-opacity');
                    anim.setAttribute('values', '0.5;0.1;0.5');
                    anim.setAttribute('dur', '1s');
                    anim.setAttribute('repeatCount', 'indefinite');
                    pulse.appendChild(anim);
                    svg.appendChild(pulse);
                }

                if (isActive && !isBranch) {
                    const glow = document.createElementNS(SVG_NS, 'line');
                    glow.setAttribute('x1', x1); glow.setAttribute('y1', y);
                    glow.setAttribute('x2', x2); glow.setAttribute('y2', y);
                    glow.setAttribute('stroke', isFailed ? 'rgba(245,100,60,0.3)' : 'rgba(255,255,255,0.2)');
                    glow.setAttribute('stroke-width', '8');
                    glow.setAttribute('stroke-linecap', 'round');
                    svg.insertBefore(glow, line);
                }

                if (isReachable) {
                    const hit = document.createElementNS(SVG_NS, 'line');
                    hit.setAttribute('x1', x1); hit.setAttribute('y1', y);
                    hit.setAttribute('x2', x2); hit.setAttribute('y2', y);
                    hit.setAttribute('stroke', 'transparent');
                    hit.setAttribute('stroke-width', '14');
                    hit.style.cursor = 'pointer';
                    hit.addEventListener('click', () => _navigate(edgeIdx));
                    const title = document.createElementNS(SVG_NS, 'title');
                    title.textContent = `${edgeIdx + 1}. ${seg.description || ''}`;
                    hit.appendChild(title);
                    svg.appendChild(hit);
                }
            });
        });

        _graphEl.appendChild(svg);
    }

    // ── Render card ───────────────────────────────────────────────────────────

    function _render() {
        const seg   = _segments[_currentIndex];
        const total = _segments.length;
        const idx   = _currentIndex;
        if (!seg) return;

        const isLast          = idx >= total - 1;
        const isLatest        = idx === _completedUpTo;
        const awaitingAdvance = _advanceResolve !== null;
        const isFailed        = _failedIndices.has(idx);

        _badge.textContent = _isRunning
            ? `Running ${idx+1}/${total}…`
            : isFailed
                ? `✗ Step ${idx+1} of ${total} — failed`
                : `Step ${idx+1} of ${total}`;

        _ranges.innerHTML = '';
        (seg.sheet_context || []).forEach(addr => {
            const c = document.createElement('span');
            c.className = 'range-chip'; c.textContent = addr;
            _ranges.appendChild(c);
        });

        _desc.textContent = seg.description;
        // Show error message below explanation when step failed
        if (isFailed && seg._errorMsg) {
            _expl.innerHTML = `<span style="opacity:0.75">${seg.explanation || ''}</span>` +
                `<div class="step-error-msg">⚠ ${seg._errorMsg}</div>`;
        } else {
            _expl.textContent = seg.explanation || '';
        }
        _counter.textContent = `${idx+1}/${total}`;

        // Q&A pairs
        _qaList.innerHTML = '';
        (seg.qa_pairs || []).forEach(pair => {
            const item = document.createElement('details');
            item.className = 'qa-item';
            item.innerHTML = `<summary class="qa-q">${pair.q}</summary><p class="qa-a">${pair.a}</p>`;
            _qaList.appendChild(item);
        });

        _btnPrev.disabled = _isRunning || idx <= 0;

        if (_isRunning) {
            _btnNext.textContent = '…';
            _btnNext.disabled    = true;
        } else if (isFailed) {
            // On a failed step: Next advances to next step (or finishes), always enabled
            _btnNext.textContent = isLast ? '✓' : '→';
            _btnNext.disabled    = false;
        } else if (isLatest && awaitingAdvance) {
            _btnNext.textContent = isLast ? '✓' : '→';
            _btnNext.disabled    = false;
        } else {
            _btnNext.textContent = '→';
            _btnNext.disabled    = idx >= _completedUpTo;
        }

        _btnEdit.disabled = _isRunning;
        _btnAsk.disabled  = _isRunning;

        // Always redraw graph in sync with card state — single update path
        _renderGraph();
    }

    // ── Range focus ───────────────────────────────────────────────────────────

    async function _focusRanges(index) {
        const contexts = _segments[index]?.sheet_context;
        if (!contexts?.length) return;
        try {
            await Excel.run(async (ctx) => {
                ctx.workbook.worksheets.getActiveWorksheet()
                    .getRange(contexts.join(', ')).select();
                await ctx.sync();
            });
        } catch(err) {
            console.warn('[StepNavigator] focusRanges:', err.message);
        }
    }

    return { init, loadSegments, setRubric, showRubricGate, showVerifyResults, markRunning, markFailed, waitForNext, dismiss };
})();
