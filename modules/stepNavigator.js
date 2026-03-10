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
    let _askHistory     = [];   // per-step ask history
    let _selectedAltId  = null; // selected alternative for current step
    let _rubric         = { hard_requirements: [], soft_requirements: [] };
    let _activePanel    = null; // 'ask' | 'edit' | 'rubric' | null

    // ── Public ────────────────────────────────────────────────────────────────

    function init() {
        _btnPrev.addEventListener('click', _onPrev);
        _btnNext.addEventListener('click', _onNext);
        _btnClose.addEventListener('click', dismiss);
        _btnAsk.addEventListener('click', () => _togglePanel('ask'));
        _btnEdit.addEventListener('click', () => _togglePanel('edit'));
        _btnVerify.addEventListener('click', _onVerify);

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
        _renderGraph();
    }

    function setRubric(rubric) {
        _rubric = rubric;
        _renderRubric();
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

    function dismiss() {
        _overlay.classList.remove('visible', 'running');
        _chatPanel.classList.remove('nav-active');
        _closePanel();
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    function _onNext() {
        if (_isRunning) return;
        if (_currentIndex < _completedUpTo) { _navigate(_currentIndex + 1); return; }
        if (_advanceResolve) {
            const resolve = _advanceResolve;
            _advanceResolve = null;
            if (_currentIndex >= _segments.length - 1) dismiss();
            resolve();
        }
    }

    function _onPrev() {
        if (_isRunning || _currentIndex <= 0) return;
        _navigate(_currentIndex - 1);
    }

    async function _navigate(targetIndex) {
        if (targetIndex < 0 || targetIndex > _completedUpTo) return;
        _currentIndex  = targetIndex;
        _selectedAltId = null;
        _askHistory    = [];
        _closePanel();
        _render();
        _renderGraph();
        await _focusRanges(targetIndex);
    }

    // Navigate to a node by segment id (from graph click)
    async function _navigateById(segId) {
        const idx = _segments.findIndex(s => s.id === segId);
        if (idx < 0 || idx > _completedUpTo) return;
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
        affs.forEach(aff => {
            const row = document.createElement('div');
            row.className = 'aff-row';
            let control = '';
            if (aff.type === 'dropdown' && aff.options?.length) {
                control = `<select class="aff-control" data-aff="${aff.id}">
                    ${aff.options.map(o => `<option ${o===aff.value?'selected':''}>${o}</option>`).join('')}
                </select>`;
            } else if (aff.type === 'color') {
                control = `<input type="color" class="aff-control aff-color" data-aff="${aff.id}" value="${aff.value}">`;
            } else if (aff.type === 'number') {
                control = `<input type="number" class="aff-control aff-number" data-aff="${aff.id}" value="${aff.value}">`;
            } else if (aff.type === 'toggle') {
                control = `<input type="checkbox" class="aff-control" data-aff="${aff.id}" ${aff.value==='true'?'checked':''}>`;
            }
            row.innerHTML = `<span class="aff-label">${aff.label}</span>${control}`;
            _affordances.appendChild(row);
        });
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

    function _renderRubric() {
        _rubricHard.innerHTML = '';
        _rubricSoft.innerHTML = '';
        (_rubric.hard_requirements || []).forEach(r => _appendRubricRow(r, 'hard'));
        (_rubric.soft_requirements || []).forEach(r => _appendRubricRow(r, 'soft'));
    }

    function _appendRubricRow(req, type, verifyResult = null) {
        const list = type === 'hard' ? _rubricHard : _rubricSoft;
        const row  = document.createElement('div');
        row.className = 'rubric-row';
        row.dataset.id   = req.id;
        row.dataset.type = type;

        let statusIcon = '';
        if (verifyResult !== null) {
            statusIcon = verifyResult.met
                ? `<span class="rubric-check" title="${verifyResult.reasoning}">✓</span>`
                : `<span class="rubric-warn" title="${verifyResult.reasoning}">⚠</span>`;
        }

        row.innerHTML = `
            <span class="rubric-badge rubric-${type}">${type === 'hard' ? 'H' : 'S'}</span>
            <span class="rubric-label">${req.label}</span>
            ${statusIcon}
            <button class="rubric-move" title="Move to ${type==='hard'?'soft':'hard'}">⇄</button>
            <button class="rubric-del" title="Remove">✕</button>`;

        row.querySelector('.rubric-move').onclick = () => _moveRubricItem(req.id, type);
        row.querySelector('.rubric-del').onclick  = () => _deleteRubricItem(req.id, type);
        list.appendChild(row);
    }

    function _moveRubricItem(id, fromType) {
        const fromList = fromType === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const toList   = fromType === 'hard' ? _rubric.soft_requirements : _rubric.hard_requirements;
        const idx = fromList.findIndex(r => r.id === id);
        if (idx < 0) return;
        const [item] = fromList.splice(idx, 1);
        toList.push(item);
        _renderRubric();
    }

    function _deleteRubricItem(id, type) {
        const list = type === 'hard' ? _rubric.hard_requirements : _rubric.soft_requirements;
        const idx  = list.findIndex(r => r.id === id);
        if (idx >= 0) list.splice(idx, 1);
        _renderRubric();
    }

    function _onRubricAdd() {
        const label = prompt('New requirement:');
        if (!label) return;
        const id = 's' + Date.now();
        _rubric.soft_requirements.push({ id, label, checked: false });
        _renderRubric();
    }

    async function _onVerify() {
        _rubricVerifyPanel.innerHTML = '<span style="color:var(--color-text-muted);font-size:11px">Verifying…</span>';
        _rubricPanel.style.display = 'block';
        _activePanel = 'rubric';
        try {
            const wsCtx = await WorksheetContext.gather(['sheet']);
            const res   = await LLMClient.rubricVerify(_rubric, wsCtx);
            _rubricVerifyPanel.innerHTML = '';
            (res.results || []).forEach(r => {
                const item = [...(_rubric.hard_requirements), ...(_rubric.soft_requirements)].find(x => x.id === r.id);
                const type = _rubric.hard_requirements.find(x => x.id === r.id) ? 'hard' : 'soft';
                if (item) _appendRubricRow(item, type, r);
            });
        } catch(err) {
            _rubricVerifyPanel.innerHTML = `<span style="color:var(--color-error)">${err.message}</span>`;
        }
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
        _segments.forEach((seg, i) => {
            (seg.alternatives || []).slice(1).forEach((alt, ai) => {
                if (i > 0) {
                    const line = document.createElementNS(SVG_NS, 'line');
                    line.setAttribute('x1', cx(i-1)); line.setAttribute('y1', cy - 8);
                    line.setAttribute('x2', cx(i));   line.setAttribute('y2', cy - 8);
                    line.setAttribute('stroke', 'rgba(245,166,35,0.35)');
                    line.setAttribute('stroke-width', '1.5');
                    line.setAttribute('stroke-dasharray', '3,3');
                    svg.appendChild(line);
                }
            });
        });

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
            if (i < _completedUpTo) { fill = 'rgba(62,207,142,0.7)'; stroke = '#3ecf8e'; }
            if (i === _completedUpTo) { fill = '#3ecf8e'; stroke = '#3ecf8e'; }
            if (i === _currentIndex) { stroke = '#fff'; }
            if (i > _completedUpTo) { fill = 'rgba(79,142,247,0.1)'; stroke = 'rgba(79,142,247,0.25)'; }

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

        _badge.textContent = _isRunning
            ? `Running ${idx+1}/${total}…`
            : `Step ${idx+1} of ${total}`;

        _ranges.innerHTML = '';
        (seg.sheet_context || []).forEach(addr => {
            const c = document.createElement('span');
            c.className = 'range-chip'; c.textContent = addr;
            _ranges.appendChild(c);
        });

        _desc.textContent = seg.description;
        _expl.textContent = seg.explanation || '';
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
            _btnNext.textContent = '…'; _btnNext.disabled = true;
        } else if (isLatest && awaitingAdvance) {
            _btnNext.textContent = isLast ? '✓' : '→'; _btnNext.disabled = false;
        } else {
            _btnNext.textContent = '→'; _btnNext.disabled = idx >= _completedUpTo;
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

    return { init, loadSegments, setRubric, markRunning, waitForNext, dismiss };
})();
