/**
 * stepNavigator.js
 * All SF1/SF2 features: graph, rubric, ask, edit, affordances, alternatives, chat dimming.
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
    const _altList      = document.getElementById('step-nav-alt-list');
    const _affordances  = document.getElementById('step-nav-affordances');
    const _rubricPanel  = document.getElementById('step-nav-rubric');
    const _rubricHard   = document.getElementById('rubric-hard-list');
    const _rubricSoft   = document.getElementById('rubric-soft-list');
    const _rubricAdd    = document.getElementById('rubric-add-btn');
    const _rubricVerifyPanel = document.getElementById('rubric-verify-results');
    const _chatPanel    = document.getElementById('chat-panel');
    const _qaList       = document.getElementById('step-nav-qa-list');

    // ── State ─────────────────────────────────────────────────────────────────
    let _segments       = [];
    let _completedUpTo  = -1;
    let _currentIndex   = 0;
    let _isRunning      = false;
    let _advanceResolve = null;
    let _askHistory     = [];
    let _selectedAltId  = null;
    let _rubric         = { hard_requirements: [], soft_requirements: [] };
    let _activePanel    = null; // 'ask' | 'edit' | 'rubric' | null
    let _isRubricGate   = false;
    let _isVerifyGate   = false; // true while showing verification results screen
    let _failedIndices  = new Set(); // indices of segments that errored

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

    function loadSegments(segments) {
        _segments       = segments;
        _completedUpTo  = -1;
        _currentIndex   = 0;
        _isRunning      = false;
        _advanceResolve = null;
        _selectedAltId  = null;
        _askHistory     = [];
        _failedIndices  = new Set();
        _isVerifyGate   = false;
        _renderGraph();
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
        _renderGraph();
    }

    async function waitForNext(index) {
        _isRunning     = false;
        _completedUpTo = Math.max(_completedUpTo, index);
        _currentIndex  = index;
        _selectedAltId = null;
        _askHistory    = [];
        _overlay.classList.remove('running');
        _overlay.classList.add('visible');
        _chatPanel.classList.add('nav-active');

        const promise = new Promise(resolve => { _advanceResolve = resolve; });
        _render();
        _renderGraph();
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
        _selectedAltId = null;
        _askHistory    = [];
        _overlay.classList.remove('running');
        _overlay.classList.add('visible', 'failed');
        _chatPanel.classList.add('nav-active');

        // Temporarily show error in explanation area
        const seg = _segments[index];
        if (seg) seg._errorMsg = errorMsg;

        const promise = new Promise(resolve => { _advanceResolve = resolve; });
        _render();
        _renderGraph();
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
        _selectedAltId = null;
        _askHistory    = [];
        _closePanel();
        // Reflect failed state on overlay if navigating to a failed step
        if (_failedIndices.has(targetIndex)) {
            _overlay.classList.add('failed');
        } else {
            _overlay.classList.remove('failed');
        }
        _render();
        _renderGraph();
        await _focusRanges(targetIndex);
    }

    // Navigate to a node by segment id (from graph click)
    async function _navigateById(segId) {
        const idx = _segments.findIndex(s => s.id === segId);
        if (idx < 0) return;
        // Allow visiting completed or failed steps
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
        if (name === 'edit') _renderEditPanel();
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

    // ── Edit / Alternatives feature ───────────────────────────────────────────

    function _renderEditPanel() {
        const seg  = _segments[_currentIndex];
        const alts = seg?.alternatives || [];
        _altList.innerHTML = '';
        alts.forEach(alt => {
            const row = document.createElement('div');
            row.className = 'alt-row' + (alt.id === (_selectedAltId || alts[0]?.id) ? ' selected' : '');
            row.innerHTML = `
                <div class="alt-row-header">
                    <span class="alt-label">${alt.label}</span>
                    <span class="alt-prob">${Math.round(alt.probability * 100)}%</span>
                </div>`;
            row.onclick = () => _selectAlt(alt.id);
            _altList.appendChild(row);
        });
        _renderAffordances();
    }

    function _selectAlt(altId) {
        _selectedAltId = altId;
        _renderEditPanel();
    }

    function _renderAffordances() {
        const seg = _segments[_currentIndex];
        const affs = seg?.affordances || [];
        _affordances.innerHTML = '';
        if (affs.length === 0) {
            const empty = document.createElement('div');
            empty.style.cssText = 'font-size:11px;color:rgba(255,255,255,0.4);padding:4px 0';
            empty.textContent = 'No dynamic controls for this step.';
            _affordances.appendChild(empty);
            return;
        }
        affs.forEach(aff => {
            const row = document.createElement('div');
            row.className = 'aff-row';
            let control;
            if (aff.type === 'dropdown' && aff.options?.length) {
                control = document.createElement('select');
                control.className = 'aff-control';
                aff.options.forEach(o => {
                    const opt = document.createElement('option');
                    opt.value = o; opt.textContent = o;
                    if (o === aff.value) opt.selected = true;
                    control.appendChild(opt);
                });
            } else if (aff.type === 'color') {
                control = document.createElement('input');
                control.type = 'color';
                control.className = 'aff-control aff-color';
                control.value = aff.value || '#000000';
            } else if (aff.type === 'number') {
                control = document.createElement('input');
                control.type = 'number';
                control.className = 'aff-control aff-number';
                control.value = aff.value || '0';
            } else if (aff.type === 'toggle') {
                control = document.createElement('input');
                control.type = 'checkbox';
                control.className = 'aff-control';
                control.checked = aff.value === 'true';
            } else {
                control = document.createElement('input');
                control.type = 'text';
                control.className = 'aff-control aff-number';
                control.value = aff.value || '';
            }
            control.dataset.affId = aff.id;

            // Live apply: re-run the current segment's code with the updated affordance value
            control.addEventListener('change', () => {
                aff.value = control.type === 'checkbox' ? String(control.checked) : control.value;
                _applyAffordance(seg, aff);
            });

            const label = document.createElement('span');
            label.className = 'aff-label';
            label.textContent = aff.label;
            row.appendChild(label);
            row.appendChild(control);
            _affordances.appendChild(row);
        });
    }

    /**
     * Inject the updated affordance value into the segment code (via placeholder comment
     * or by re-running the code with a find-replace on the old value) and run it live.
     */
    async function _applyAffordance(seg, aff) {
        // Replace the placeholder comment /* AFFORDANCE:aff-id */ with the value,
        // or fall back to a simple string replace of the old default in the code.
        let code = seg.code;
        const placeholder = `/* AFFORDANCE:${aff.id} */`;
        if (code.includes(placeholder)) {
            code = code.replace(placeholder, JSON.stringify(aff.value));
        } else {
            // Heuristic: replace the first occurrence of the previous rendered value
            // We store the live value on the aff object so we can swap it
            // Build a snippet that directly applies based on affordance type
            code = _buildAffordanceSnippet(seg, aff);
        }
        if (!code) return;
        try {
            const fn = new (Object.getPrototypeOf(async function(){}).constructor)(code);
            await fn();
        } catch(err) {
            console.warn('[StepNavigator] affordance apply error:', err.message);
        }
    }

    /**
     * Build a minimal targeted Office.js snippet for common affordance patterns
     * without needing to re-run the full segment code.
     */
    function _buildAffordanceSnippet(seg, aff) {
        const ranges = (seg.sheet_context || []).join(', ') || 'A1';
        const val    = aff.value;
        const label  = aff.label.toLowerCase();

        // Color affordances
        if (aff.type === 'color') {
            if (label.includes('background') || label.includes('fill')) {
                return `await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("${ranges}").format.fill.color = ${JSON.stringify(val)}; await ctx.sync(); });`;
            }
            if (label.includes('font') || label.includes('text')) {
                return `await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("${ranges}").format.font.color = ${JSON.stringify(val)}; await ctx.sync(); });`;
            }
            // Default: try fill
            return `await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("${ranges}").format.fill.color = ${JSON.stringify(val)}; await ctx.sync(); });`;
        }

        // Number format affordances
        if (aff.type === 'dropdown' && (label.includes('format') || label.includes('number'))) {
            const cells = (seg.sheet_context || ['A1'])[0];
            return `await Excel.run(async (ctx) => {
                const r = ctx.workbook.worksheets.getActiveWorksheet().getRange("${cells}");
                r.load("rowCount,columnCount"); await ctx.sync();
                const fmt = Array.from({length:r.rowCount}, () => Array.from({length:r.columnCount}, () => ${JSON.stringify(val)}));
                r.numberFormat = fmt; await ctx.sync();
            });`;
        }

        // Font size
        if (aff.type === 'number' && label.includes('size')) {
            return `await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("${ranges}").format.font.size = ${Number(val)}; await ctx.sync(); });`;
        }

        // Threshold for conditional coloring — re-run full segment with updated code
        if (aff.type === 'number' && label.includes('threshold')) {
            // Patch the numeric literal in the segment code
            return seg.code.replace(/>=\s*\d+/g, `>= ${Number(val)}`);
        }

        // Alignment
        if (aff.type === 'dropdown' && label.includes('alignment')) {
            return `await Excel.run(async (ctx) => { ctx.workbook.worksheets.getActiveWorksheet().getRange("${ranges}").format.horizontalAlignment = ${JSON.stringify(val)}; await ctx.sync(); });`;
        }

        // Label/text values — re-run full segment substituting first string literal
        if (aff.type === 'dropdown' && label.includes('label')) {
            return seg.code.replace(/"TOTAL"|"SUM"|"AVERAGE"|"GRAND TOTAL"/, JSON.stringify(val));
        }

        // Row color (even/odd stripe)
        if (aff.type === 'color' && (label.includes('even') || label.includes('odd'))) {
            const isEven = label.includes('even');
            return `await Excel.run(async (ctx) => {
                const s = ctx.workbook.worksheets.getActiveWorksheet();
                for (let i = 2; i <= 7; i++) {
                    if (i % 2 === ${isEven ? 0 : 1}) s.getRange("A"+i+":E"+i).format.fill.color = ${JSON.stringify(val)};
                }
                await ctx.sync();
            });`;
        }

        // Fallback: re-run full segment code as-is (user changed something, attempt full re-apply)
        return seg.code;
    }

    async function _onEditSend() {
        const msg    = _editFeedback.value.trim();
        const seg    = _segments[_currentIndex];
        const altId  = _selectedAltId;
        _editSend.disabled = true;
        try {
            const wsCtx     = await WorksheetContext.gather(['sheet']);
            const updated   = await LLMClient.edit(msg, wsCtx, seg, altId);
            _segments[_currentIndex] = { ..._segments[_currentIndex], ...updated };
            _editFeedback.value = '';
            _render();
            _renderEditPanel();
            _renderGraph();
        } catch(err) {
            console.error('[StepNavigator] edit error:', err);
        } finally {
            _editSend.disabled = false;
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

    // ── Dependency Graph ──────────────────────────────────────────────────────

    function _renderGraph() {
        if (!_graphEl) return;
        _graphEl.innerHTML = '';
        if (_segments.length === 0) return;

        const SVG_NS = 'http://www.w3.org/2000/svg';
        const NODE_R = 7;
        const GAP    = 34;
        const H      = 36;
        const total  = _segments.length;
        const W      = Math.max(GAP * (total - 1) + NODE_R * 2 + 20, 100);

        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
        svg.setAttribute('width', '100%');
        svg.setAttribute('height', H);

        // Node x positions
        const cx = i => NODE_R + 10 + i * GAP;
        const cy = H / 2;

        // Draw predecessor edges
        _segments.forEach((seg, i) => {
            (seg.predecessors || []).forEach(predId => {
                const pi = _segments.findIndex(s => s.id === predId);
                if (pi < 0) return;
                const line = document.createElementNS(SVG_NS, 'line');
                line.setAttribute('x1', cx(pi)); line.setAttribute('y1', cy);
                line.setAttribute('x2', cx(i));  line.setAttribute('y2', cy);
                const done = pi <= _completedUpTo && i <= _completedUpTo;
                line.setAttribute('stroke', done ? 'rgba(79,142,247,0.7)' : 'rgba(79,142,247,0.2)');
                line.setAttribute('stroke-width', '2');
                line.setAttribute('stroke-dasharray', done ? 'none' : '4,3');
                svg.appendChild(line);
            });
        });

        // Draw alternative edges (dotted, low opacity)
        // _segments.forEach((seg, i) => {
        //     (seg.alternatives || []).slice(1).forEach((alt, ai) => {
        //         if (i > 0) {
        //             const line = document.createElementNS(SVG_NS, 'line');
        //             line.setAttribute('x1', cx(i-1)); line.setAttribute('y1', cy - 8);
        //             line.setAttribute('x2', cx(i));   line.setAttribute('y2', cy - 8);
        //             line.setAttribute('stroke', 'rgba(245,166,35,0.35)');
        //             line.setAttribute('stroke-width', '1.5');
        //             line.setAttribute('stroke-dasharray', '3,3');
        //             svg.appendChild(line);
        //         }
        //     });
        // });

        // Draw nodes
        _segments.forEach((seg, i) => {
            const g = document.createElementNS(SVG_NS, 'g');
            g.style.cursor = 'pointer';
            g.addEventListener('click', () => _navigateById(seg.id));

            const circle = document.createElementNS(SVG_NS, 'circle');
            circle.setAttribute('cx', cx(i));
            circle.setAttribute('cy', cy);
            circle.setAttribute('r', i === _currentIndex ? NODE_R + 2 : NODE_R);

            let fill = 'rgba(79,142,247,0.15)';
            let stroke = 'rgba(79,142,247,0.4)';
            if (_failedIndices.has(i))   { fill = 'rgba(245,100,60,0.8)'; stroke = '#f5643c'; }
            else if (i < _completedUpTo) { fill = 'rgba(62,207,142,0.7)'; stroke = '#3ecf8e'; }
            else if (i === _completedUpTo) { fill = '#3ecf8e'; stroke = '#3ecf8e'; }
            else if (i > _completedUpTo) { fill = 'rgba(79,142,247,0.1)'; stroke = 'rgba(79,142,247,0.25)'; }
            if (i === _currentIndex) { stroke = '#fff'; }

            circle.setAttribute('fill', fill);
            circle.setAttribute('stroke', stroke);
            circle.setAttribute('stroke-width', i === _currentIndex ? '2.5' : '1.5');

            // Running pulse
            if (i === _currentIndex && _isRunning) {
                const anim = document.createElementNS(SVG_NS, 'animate');
                anim.setAttribute('attributeName', 'r');
                anim.setAttribute('values', `${NODE_R};${NODE_R+4};${NODE_R}`);
                anim.setAttribute('dur', '1s');
                anim.setAttribute('repeatCount', 'indefinite');
                circle.appendChild(anim);
            }

            const title = document.createElementNS(SVG_NS, 'title');
            title.textContent = `${i+1}. ${seg.description}`;

            g.appendChild(circle);
            g.appendChild(title);
            svg.appendChild(g);
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
        // if (isFailed && seg._errorMsg) {
        //     // _expl.innerHTML = `<span style="opacity:0.75">${seg.explanation || ''}</span>` +
        //     //     `<div class="step-error-msg">⚠ ${seg._errorMsg}</div>`;
        // } else {
        _expl.textContent = seg.explanation || '';
        // }
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
