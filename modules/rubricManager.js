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
    // All static rubric-panel elements hidden during verify results
    const _staticEls  = () => [
        ...document.querySelectorAll('#step-nav-rubric .rubric-section-label'),
        _hardList, _softList, _addBtn,
    ];

    // ── State ─────────────────────────────────────────────────────────────────
    let _rubric        = { hard_requirements: [], soft_requirements: [] };
    let _advanceResolve = null;
    let _advanceReject  = null;
    let _dragId        = null;
    let _dragType      = null;

    // ── Public ────────────────────────────────────────────────────────────────

    /** Reset all per-run state — called by chatManager before each new run. */
    function reset() {
        _rubric = { hard_requirements: [], soft_requirements: [] };
        _advanceResolve = null;
        _advanceReject  = null;
        // Clear verify results from previous run
        if (_verifyEl) _verifyEl.innerHTML = '';
        // Restore static elements hidden during verify
        _staticEls().forEach(el => { if (el) el.style.display = ''; });
        _panel?.querySelector('.rubric-loading')?.remove();
        showPanel(false);
    }

    function init() {
        _addBtn?.addEventListener('click', _onAdd);
    }

    function setRubric(rubric) {
        _rubric = rubric || { hard_requirements: [], soft_requirements: [] };
        // Clear loading placeholder from panel container
        _panel?.querySelector('.rubric-loading')?.remove();
        _render();
    }

    function getRubric() { return _rubric; }

    /** Open / close the rubric panel (called by StepNavigator panel toggle). */
    function showPanel(visible) {
        if (_panel) _panel.style.display = visible ? 'block' : 'none';
        if (visible) _render();
    }

    /** Show the rubric gate immediately. Rubric loads in parallel via setRubric(). */
    function showRubricGate() {
        _lockNav();
        _overlay.classList.add('visible', 'rubric-gate');
        _chatPanel.classList.add('nav-active');
        showPanel(true);
        _render();

        _btnNext.textContent = '→';
        _btnNext.disabled    = false;
        _badge.textContent   = 'Review Requirements';
        _counter.textContent = '';
        _ranges.innerHTML    = '';
        _desc.textContent    = 'Review and edit the requirements below, then click Start to begin.';
        _expl.textContent    = '';
        _qaList.innerHTML    = '';
        // Show loading placeholder in the panel container until rubric arrives
        _hardList.innerHTML = '';
        _softList.innerHTML = '';
        if (_panel) {
            let _loadingEl = _panel.querySelector('.rubric-loading');
            if (!_loadingEl) {
                _loadingEl = document.createElement('div');
                _loadingEl.className = 'rubric-loading';
                _loadingEl.textContent = 'Generating requirements…';
                _panel.insertBefore(_loadingEl, _panel.firstChild);
            }
        }
    }

    /** Returns a promise that resolves when user clicks →.
     *  Rejects with DismissedError if the user closes the navigator. */
    function waitForGate() {
        return new Promise((resolve, reject) => {
            _advanceResolve = resolve;
            _advanceReject  = reject;
            StepNavigator.setGateMode(() => {
                if (_advanceResolve) { const r = _advanceResolve; _advanceResolve = null; _advanceReject = null; r(); }
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
        _overlay.classList.remove('failed');
        _overlay.classList.add('visible', 'verify-gate');
        _chatPanel.classList.add('nav-active');

        _badge.textContent   = 'Requirements Check';
        _counter.textContent = '';
        _desc.textContent    = '';
        _expl.textContent    = '';
        _ranges.innerHTML    = '';
        _qaList.innerHTML    = '';

        // Register gate NOW so _renderCard enables the ✓ button immediately.
        const gatePromise = new Promise((resolve, reject) => {
            _advanceResolve = resolve;
            _advanceReject  = reject;
            StepNavigator.setGateMode(() => {
                if (_advanceResolve) { const r = _advanceResolve; _advanceResolve = null; _advanceReject = null; r(); }
                _staticEls().forEach(el => { if (el) el.style.display = ''; });
                showPanel(false);
                // Full dismiss — re-enables chat and clears all overlay state
                StepNavigator.dismiss();
            });
        });

        // Show the rubric panel, hide the editable gate elements
        showPanel(true);
        _staticEls().forEach(el => { if (el) el.style.display = 'none'; });
        _verifyEl.innerHTML  = '<span class="verify-loading">Verifying requirements…</span>';

        try {
            const wsCtx = await WorksheetContext.gather(['sheet']);
            const res   = await LLMClient.rubricVerify(_rubric, wsCtx);

            // Build a lookup from the verify response
            const resultMap = {};
            (res.results || []).forEach(r => { resultMap[r.id] = r; });

            const hard = _rubric.hard_requirements || [];
            const soft = _rubric.soft_requirements || [];
            const all  = [...hard, ...soft];

            // Count met only for requirements that exist in the rubric
            const metCount = all.filter(req => resultMap[req.id]?.met).length;
            const total    = all.length;
            const allMet   = metCount === total;

            // ── Clear loading text and render results ─────────────────────
            _verifyEl.innerHTML = '';

            // Score line at the top of the verify area
            const scoreEl = document.createElement('div');
            scoreEl.className = 'verify-score';
            scoreEl.innerHTML =
                `<span class="verify-score-num ${allMet ? 'all-met' : metCount === 0 ? 'none-met' : ''}">`
                + `${metCount}<span class="verify-score-denom">/${total}</span></span>`
                + `<span class="verify-score-label">${allMet ? 'All requirements met 🎉' : 'requirements met'}</span>`;
            _verifyEl.appendChild(scoreEl);

            // ── Read-only rows — same visual style as the gate rubric rows
            //    but with ✓/⚠ on the left instead of drag/move/delete ────
            const _appendVerifyRow = (req, type) => {
                const r   = resultMap[req.id];
                const met = r?.met ?? false;

                const row = document.createElement('div');
                row.className = 'rubric-row';

                const icon = document.createElement('span');
                icon.className   = met ? 'rubric-check' : 'rubric-warn';
                icon.textContent = met ? '✓' : '⚠';

                const badge = document.createElement('span');
                badge.className   = `rubric-badge rubric-${type}`;
                badge.textContent = type === 'hard' ? 'H' : 'S';

                const label = document.createElement('span');
                label.className   = 'rubric-label verify-label-full';
                label.textContent = req.label;

                row.appendChild(icon);
                row.appendChild(badge);
                row.appendChild(label);

                // Click row to toggle reasoning
                if (r?.reasoning) {
                    row.style.cursor = 'pointer';
                    let detail = null;
                    row.addEventListener('click', () => {
                        if (detail) { detail.remove(); detail = null; return; }
                        detail = document.createElement('div');
                        detail.className = 'rubric-detail';
                        detail.style.display = 'block';
                        detail.textContent = r.reasoning
                            + (r.references?.length ? ` (${r.references.join(', ')})` : '');
                        row.after(detail);
                    });
                }

                _verifyEl.appendChild(row);
            };

            if (hard.length) {
                const hl = document.createElement('div');
                hl.className   = 'rubric-section-label';
                hl.textContent = 'Hard Requirements';
                _verifyEl.appendChild(hl);
                hard.forEach(r => _appendVerifyRow(r, 'hard'));
            }
            if (soft.length) {
                const sl = document.createElement('div');
                sl.className   = 'rubric-section-label';
                sl.textContent = 'Soft Requirements';
                _verifyEl.appendChild(sl);
                soft.forEach(r => _appendVerifyRow(r, 'soft'));
            }

        } catch (err) {
            _verifyEl.innerHTML =
                `<span style="color:var(--color-error);font-size:11px">Verification failed: ${err.message}</span>`;
        }

        return gatePromise;
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

    /** Called by StepNavigator.dismiss() to abort any pending gate promise. */
    function rejectGate() {
        _panel?.querySelector('.rubric-loading')?.remove();
        showPanel(false);
        if (_advanceReject) {
            const rj = _advanceReject;
            _advanceResolve = null; _advanceReject = null;
            rj(new Error('dismissed'));
        }
        if (_advanceResolve) {
            const r = _advanceResolve;
            _advanceResolve = null;
            r();  // also resolve plain advance promises
        }
    }

    return { init, reset, setRubric, getRubric, showPanel, showRubricGate, waitForGate, rejectGate, showVerifyResults };
})();
