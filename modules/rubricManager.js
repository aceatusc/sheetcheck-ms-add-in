/**
 * rubricManager.js
 *
 * Owns the rubric lifecycle:
 *   - Rubric state (_rubric)
 *   - Editable list UI (hard / soft requirements, drag-to-reorder)
 *   - showRubricGate()   — shown before execution starts
 *   - showVerifyResults() — shown after all segments complete
 *
 * DOM dependencies (same elements formerly used inside stepNavigator):
 *   #step-navigator, #chat-panel
 *   #step-nav-badge, #step-nav-description, #step-nav-explanation
 *   #step-nav-counter, #step-nav-ranges, #step-nav-qa-list
 *   #step-nav-prev, #step-nav-next, #step-nav-edit, #step-nav-ask
 *   #step-nav-rubric, #rubric-hard-list, #rubric-soft-list
 *   #rubric-add-btn, #rubric-verify-results
 */
const RubricManager = (() => {

    // ── DOM refs ──────────────────────────────────────────────────────────────
    const _overlay    = document.getElementById('step-navigator');
    const _chatPanel  = document.getElementById('chat-panel');
    const _badge      = document.getElementById('step-nav-badge');
    const _desc       = document.getElementById('step-nav-description');
    const _expl       = document.getElementById('step-nav-explanation');
    const _counter    = document.getElementById('step-nav-counter');
    const _ranges     = document.getElementById('step-nav-ranges');
    const _qaList     = document.getElementById('step-nav-qa-list');
    const _btnPrev    = document.getElementById('step-nav-prev');
    const _btnNext    = document.getElementById('step-nav-next');
    const _btnEdit    = document.getElementById('step-nav-edit');
    const _btnAsk     = document.getElementById('step-nav-ask');
    const _panel      = document.getElementById('step-nav-rubric');
    const _hardList   = document.getElementById('rubric-hard-list');
    const _softList   = document.getElementById('rubric-soft-list');
    const _addBtn     = document.getElementById('rubric-add-btn');
    const _verifyEl   = document.getElementById('rubric-verify-results');

    // ── State ─────────────────────────────────────────────────────────────────
    let _rubric        = { hard_requirements: [], soft_requirements: [] };
    let _advanceResolve = null;
    let _dragId        = null;
    let _dragType      = null;

    // ── Public ────────────────────────────────────────────────────────────────

    function init() {
        _addBtn?.addEventListener('click', _onAdd);
    }

    function setRubric(rubric) {
        _rubric = rubric || { hard_requirements: [], soft_requirements: [] };
        _render();
    }

    function getRubric() { return _rubric; }

    /** Open / close the rubric panel (called by StepNavigator panel toggle). */
    function showPanel(visible) {
        if (_panel) _panel.style.display = visible ? 'block' : 'none';
        if (visible) _render();
    }

    /**
     * Show the "Review Requirements" gate before execution.
     * Resolves when the user clicks Start →.
     *
     * Registers a one-shot gate callback on StepNavigator so the next →
     * click resolves the promise — no shared boolean flag needed between modules.
     */
    function showRubricGate() {
        _lockNav();
        _overlay.classList.add('visible', 'rubric-gate');
        _chatPanel.classList.add('nav-active');
        showPanel(true);
        _render();

        _btnNext.textContent = 'Start →';
        _btnNext.disabled    = false;
        _badge.textContent   = 'Review Requirements';
        _counter.textContent = '';
        _ranges.innerHTML    = '';
        _desc.textContent    = 'Review and edit the requirements below, then click Start to begin.';
        _expl.textContent    = '';
        _qaList.innerHTML    = '';

        return new Promise(resolve => {
            _advanceResolve = resolve;
            // Hand off to StepNavigator: the very next → click should fire this teardown
            StepNavigator.setGateMode(() => {
                if (_advanceResolve) { const r = _advanceResolve; _advanceResolve = null; r(); }
                StepNavigator.dismissGate('rubric');
                showPanel(false);
            });
        });
    }

    /**
     * Show verification results after all segments complete.
     * Resolves when user clicks Done ✓.
     */
    async function showVerifyResults() {
        _lockNav();
        _overlay.classList.remove('failed', 'running');
        _overlay.classList.add('visible', 'verify-gate');
        _chatPanel.classList.add('nav-active');

        _badge.textContent   = 'Requirements Check';
        _counter.textContent = '';
        _desc.textContent    = 'Verifying your spreadsheet against the requirements…';
        _expl.textContent    = '';
        _ranges.innerHTML    = '';
        _qaList.innerHTML    = '';
        _btnNext.textContent = '…';
        _btnNext.disabled    = true;

        showPanel(true);
        _verifyEl.innerHTML =
            '<span style="color:rgba(255,255,255,0.6);font-size:11px">Verifying requirements…</span>';

        try {
            const wsCtx = await WorksheetContext.gather(['sheet']);
            const res   = await LLMClient.rubricVerify(_rubric, wsCtx);

            _verifyEl.innerHTML =
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

                const row  = document.createElement('div');
                row.className = 'rubric-row';
                const icon = r.met ? `<span class="rubric-check">✓</span>`
                                   : `<span class="rubric-warn">⚠</span>`;
                row.innerHTML = `
                    <span class="rubric-badge rubric-${type}">${type === 'hard' ? 'H' : 'S'}</span>
                    <span class="rubric-label" style="user-select:none">${item.label}</span>
                    ${icon}`;

                const iconEl = row.querySelector('.rubric-check, .rubric-warn');
                if (iconEl) {
                    iconEl.style.cursor = 'pointer';
                    let open = false;
                    iconEl.addEventListener('click', () => {
                        open = !open;
                        let detail = row.querySelector('.rubric-detail');
                        if (!detail) {
                            detail = document.createElement('div');
                            detail.className = 'rubric-detail';
                            row.appendChild(detail);
                        }
                        detail.textContent = r.reasoning
                            + (r.references?.length ? ` (${r.references.join(', ')})` : '');
                        detail.style.display = open ? 'block' : 'none';
                    });
                }
                _verifyEl.appendChild(row);
            });

            const total = res.results?.length || 0;
            _desc.textContent = total
                ? `${metCount} of ${total} requirement${total !== 1 ? 's' : ''} met`
                : 'No requirements to verify.';

        } catch (err) {
            _verifyEl.innerHTML =
                `<span style="color:var(--color-error);font-size:11px">Verification failed: ${err.message}</span>`;
            _desc.textContent = 'Could not verify requirements.';
        }

        _btnNext.textContent = 'Done ✓';
        _btnNext.disabled    = false;

        return new Promise(resolve => {
            _advanceResolve = resolve;
            StepNavigator.setGateMode(() => {
                if (_advanceResolve) { const r = _advanceResolve; _advanceResolve = null; r(); }
                StepNavigator.dismissGate('verify');
                showPanel(false);
            });
        });
    }

    // ── Render ────────────────────────────────────────────────────────────────

    function _render() {
        if (!_hardList || !_softList) return;
        _hardList.innerHTML = '';
        _softList.innerHTML = '';

        const hard = _rubric.hard_requirements || [];
        const soft = _rubric.soft_requirements || [];
        hard.forEach(r => _appendRow(r, 'hard'));
        soft.forEach(r => _appendRow(r, 'soft'));
        _setupDropZone(_hardList, 'hard');
        _setupDropZone(_softList, 'soft');

        if (!hard.length && !soft.length) {
            const hint = document.createElement('div');
            hint.style.cssText = 'font-size:11px;color:rgba(255,255,255,0.4);padding:8px 2px;text-align:center';
            hint.textContent   = 'No requirements yet — click "+ Add requirement" below.';
            _softList.appendChild(hint);
        }
    }

    function _appendRow(req, type) {
        const list  = type === 'hard' ? _hardList : _softList;
        const row   = document.createElement('div');
        row.className    = 'rubric-row';
        row.dataset.id   = req.id;
        row.dataset.type = type;
        row.draggable    = true;

        const handle = document.createElement('span');
        handle.className = 'rubric-drag';
        handle.textContent = '⠿';
        handle.title = 'Drag to reorder';

        const badge = document.createElement('span');
        badge.className   = `rubric-badge rubric-${type}`;
        badge.textContent = type === 'hard' ? 'H' : 'S';

        const label = document.createElement('span');
        label.className       = 'rubric-label';
        label.contentEditable = 'true';
        label.spellcheck      = false;
        label.textContent     = req.label;
        label.addEventListener('blur', () => {
            const lst  = type === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
            const item = lst.find(x => x.id === req.id);
            if (item) item.label = label.textContent.trim() || item.label;
        });
        label.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); label.blur(); } });

        const moveBtn = document.createElement('button');
        moveBtn.className   = 'rubric-move';
        moveBtn.title       = `Move to ${type === 'hard' ? 'Soft' : 'Hard'}`;
        moveBtn.textContent = type === 'hard' ? '↓S' : '↑H';
        moveBtn.onclick     = () => _moveToType(req.id, type, type === 'hard' ? 'soft' : 'hard');

        const delBtn = document.createElement('button');
        delBtn.className   = 'rubric-del';
        delBtn.title       = 'Remove';
        delBtn.textContent = '✕';
        delBtn.onclick     = () => _delete(req.id, type);

        row.appendChild(handle);
        row.appendChild(badge);
        row.appendChild(label);
        row.appendChild(moveBtn);
        row.appendChild(delBtn);

        // Drag events
        row.addEventListener('dragstart', e => {
            _dragId = req.id; _dragType = type;
            row.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        });
        row.addEventListener('dragend', () => {
            row.classList.remove('dragging');
            document.querySelectorAll('.rubric-drop-over').forEach(el => el.classList.remove('rubric-drop-over'));
        });
        row.addEventListener('dragover',  e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; row.classList.add('rubric-drop-over'); });
        row.addEventListener('dragleave', () => row.classList.remove('rubric-drop-over'));
        row.addEventListener('drop', e => {
            e.preventDefault();
            row.classList.remove('rubric-drop-over');
            if (!_dragId || _dragId === req.id) return;
            _dropItem(_dragId, _dragType, req.id, type);
        });

        list.appendChild(row);
    }

    function _setupDropZone(listEl, targetType) {
        listEl.addEventListener('dragover',  e => { if (e.target === listEl) { e.preventDefault(); listEl.classList.add('rubric-list-over'); } });
        listEl.addEventListener('dragleave', e => { if (e.target === listEl) listEl.classList.remove('rubric-list-over'); });
        listEl.addEventListener('drop', e => {
            if (e.target !== listEl) return;
            e.preventDefault();
            listEl.classList.remove('rubric-list-over');
            if (!_dragId) return;
            _moveToType(_dragId, _dragType, targetType);
        });
    }

    function _dropItem(dragId, fromType, targetId, targetType) {
        const fromList = fromType === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const toList   = targetType === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const fromIdx  = fromList.findIndex(r => r.id === dragId);
        if (fromIdx < 0) return;
        const [item] = fromList.splice(fromIdx, 1);
        const toIdx  = toList.findIndex(r => r.id === targetId);
        if (toIdx >= 0) toList.splice(toIdx, 0, item); else toList.push(item);
        _render();
    }

    function _moveToType(id, fromType, toType) {
        if (fromType === toType) return;
        const fromList = fromType === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const toList   = toType   === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const idx = fromList.findIndex(r => r.id === id);
        if (idx < 0) return;
        const [item] = fromList.splice(idx, 1);
        toList.push(item);
        _render();
    }

    function _delete(id, type) {
        const list = type === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const idx  = list.findIndex(r => r.id === id);
        if (idx >= 0) list.splice(idx, 1);
        _render();
    }

    function _onAdd() {
        const id = 'u' + Date.now();
        _rubric.soft_requirements.push({ id, label: '', checked: false });
        _render();
        requestAnimationFrame(() => {
            const el = _softList.querySelector(`[data-id="${id}"] .rubric-label`);
            if (el) { el.focus(); _selectAll(el); }
        });
    }

    function _selectAll(el) {
        const r = document.createRange();
        r.selectNodeContents(el);
        const s = window.getSelection();
        s.removeAllRanges();
        s.addRange(r);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    function _lockNav() {
        _btnPrev.disabled = true;
        _btnEdit.disabled = true;
        _btnAsk.disabled  = true;
    }

    return { init, setRubric, getRubric, showPanel, showRubricGate, showVerifyResults };
})();
