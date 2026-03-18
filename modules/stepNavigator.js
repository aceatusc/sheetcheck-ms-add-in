/**
 * stepNavigator.js — Pure renderer.
 *
 * Responsibilities:
 *   - Render the step card (description, explanation, ranges, Q&A)
 *   - Render the git-style DAG graph (nodes carry state; edges are lines)
 *   - Delegate all navigation to DagRunner; all rubric to RubricManager
 *
 * Public API:
 *   init()
 *   load(chain)                          — called by DagRunner.start()
 *   markRunning()                        — called by ExecutionEngine
 *   markFailed(errorMsg) → Promise       — called by ExecutionEngine
 *   waitForNext()        → Promise       — called by ExecutionEngine; resolves on →
 *   dismiss()
 *   refreshGraph()                       — re-render graph after DagRunner mutation
 *
 * Navigation flow:
 *   → / ← buttons  → DagRunner.stepForward / stepBack  → refreshGraph()
 *   Node click      → DagRunner.navigateTo              → refreshGraph()
 */
const StepNavigator = (() => {

    // ── DOM ───────────────────────────────────────────────────────────────────
    const _overlay      = document.getElementById('step-navigator');
    const _chatPanel    = document.getElementById('chat-panel');
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
    const _graphEl      = document.getElementById('step-nav-graph');
    const _askPanel     = document.getElementById('step-nav-ask-panel');
    const _askInput     = document.getElementById('step-nav-ask-input');
    const _askSend      = document.getElementById('step-nav-ask-send');
    const _askThread    = document.getElementById('step-nav-ask-thread');
    const _askChips     = document.getElementById('step-nav-ask-chips');
    const _editPanel    = document.getElementById('step-nav-edit-panel');
    const _editParams   = document.getElementById('step-nav-edit-params');
    const _editChips    = document.getElementById('step-nav-edit-chips');
    const _editFeedback = document.getElementById('step-nav-edit-feedback');
    const _editSend     = document.getElementById('step-nav-edit-send');
    const _qaList       = document.getElementById('step-nav-qa-list');

    // ── State ─────────────────────────────────────────────────────────────────
    let _chainId        = null;
    let _isRunning      = false;
    let _advanceResolve = null;
    let _dismissed      = false;  // set on dismiss; checked by executionEngine
    let _askHistory     = [];
    let _activePanel    = null;  // 'ask' | 'edit' | 'rubric' | null
    let _gateCallback   = null;  // one-shot: fires on the next → click
    let _lastNodeId     = null;  // detect step changes to close panels

    // ── Init ──────────────────────────────────────────────────────────────────

    function init() {
        _btnPrev.addEventListener('click',  _onPrev);
        _btnNext.addEventListener('click',  _onNext);
        _btnClose.addEventListener('click', dismiss);
        _btnAsk.addEventListener('click',   () => _togglePanel('ask'));
        _btnEdit.addEventListener('click',  () => _togglePanel('edit'));

        _askSend.addEventListener('click', _onAskSend);
        _askInput.addEventListener('keydown', e => {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); _onAskSend(); }
        });
        _editSend.addEventListener('click', _onEditSend);
        _editFeedback.addEventListener('keydown', e => {
            if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); _onEditSend(); }
        });
    }

    // ── Load ──────────────────────────────────────────────────────────────────

    /** Called by DagRunner.start() with the chain handle. */
    function load(chain) {
        _chainId        = chain.chainId;
        _isRunning      = false;
        _advanceResolve = null;
        _askHistory     = [];
        _activePanel    = null;
        _gateCallback   = null;
        _dismissed      = false;
        _lastNodeId     = null;
        _closePanel();
        _render();
    }

    // ── ExecutionEngine interface ─────────────────────────────────────────────

    function markRunning() {
        _isRunning = true;
        _overlay.classList.add('running', 'visible');
        _chatPanel.classList.add('nav-active');
        _render();
    }

    async function markFailed(errorMsg) {
        _isRunning = false;
        _overlay.classList.remove('running');
        _overlay.classList.add('visible', 'failed');
        _chatPanel.classList.add('nav-active');

        const chain = DagRunner.getChain(_chainId);
        const seg   = _currentSegment(chain);
        if (seg) seg._errorMsg = errorMsg;

        const promise = new Promise(resolve => { _advanceResolve = resolve; });
        _render();
        return promise;
    }

    async function waitForNext() {
        _isRunning = false;
        _overlay.classList.remove('running');
        _overlay.classList.add('visible');
        _chatPanel.classList.add('nav-active');
        _askHistory = [];

        const promise = new Promise(resolve => { _advanceResolve = resolve; });
        _render();
        await _focusRanges();
        return promise;
    }

    function dismiss() {
        _dismissed = true;   // signals executionEngine to stop looping
        _isRunning = false;
        _overlay.classList.remove('visible', 'running', 'failed', 'rubric-gate', 'verify-gate');
        _chatPanel.classList.remove('nav-active');
        _closePanel();
        _gateCallback = null;
        // Resolve the pending promise so awaiting callers unblock,
        // but executionEngine checks _dismissed before continuing.
        _resolve();
    }

    function _render() { _renderCard(); _renderGraph(); }

    /** Re-render after DagRunner mutates chain state (public). */
    function refreshGraph() { _render(); }

    // ── Navigation ────────────────────────────────────────────────────────────

    function _onNext() {
        if (_isRunning) return;

        // Gate mode: rubric-gate or verify-gate — delegate entirely to the registered callback
        if (_gateCallback) {
            const cb = _gateCallback;
            _gateCallback = null;
            cb();
            return;
        }

        if (_advanceResolve) {
            _overlay.classList.remove('failed');
            _resolve();
        }
    }

    /**
     * Register a one-shot callback for the next → click.
     * Called by RubricManager.showRubricGate() and showVerifyResults().
     * @param {'rubric'|'verify'} _mode  (informational, unused internally)
     * @param {Function}          cb     called when → is clicked
     */
    function setGateMode(cb) {
        _gateCallback = cb;
    }

    /** Called by RubricManager after a gate resolves to clean overlay state. */
    function dismissGate(type) {
        _overlay.classList.remove(`${type}-gate`, 'visible');
        _chatPanel.classList.remove('nav-active');
    }

    function _onPrev() {
        if (_isRunning) return;
        DagRunner.stepBack(_chainId)
            .then(() => {
                _overlay.classList.remove('failed');
                _render();
            })
            .catch(err => {
                ExecutionEngine.log('err', `✗ Step back failed: ${err.message}`);
                _render();
            });
    }

    async function _onNodeClick(nodeId) {
        if (_isRunning) return;
        try {
            await DagRunner.navigateTo(_chainId, nodeId);
            _overlay.classList.remove('failed');
            _render();
            await _focusRanges();
        } catch (err) {
            ExecutionEngine.log('err', `✗ Navigate failed: ${err.message}`);
            _render();
        }
    }

    function _resolve() {
        if (_advanceResolve) { const r = _advanceResolve; _advanceResolve = null; r(); }
    }

    // ── Card render ───────────────────────────────────────────────────────────

    function _renderCard() {
        const chain = DagRunner.getChain(_chainId);
        if (!chain) return;

        // Close open panel whenever the user moves to a different step
        if (chain.currentNodeId !== _lastNodeId) {
            _closePanel();
            _lastNodeId = chain.currentNodeId;
        }

        const seg   = _currentSegment(chain);   // segment that PRODUCED this state
        const next  = _nextSegment(chain);       // segment ABOUT to run (for running badge)
        const total = chain.segments.length;
        const idx   = _currentIndex(chain);      // 0-based position of currentNodeId

        const atRoot   = DagStore.edgesTo(chain.currentNodeId).length === 0;
        const atLeaf   = DagStore.edgesFrom(chain.currentNodeId).length === 0;
        const isFailed = !!(seg?._errorMsg);

        // Badge: "Running step N…" / "Step N of M applied" / "Ready — N steps"
        if (_isRunning) {
            _badge.textContent = `Applying step ${idx + 1}…`;
        } else if (atRoot) {
            _badge.textContent = `Ready — ${total} step${total !== 1 ? 's' : ''}`;
        } else if (isFailed) {
            _badge.textContent = `✗ Step ${idx} of ${total} — failed`;
        } else {
            _badge.textContent = `Step ${idx} of ${total} applied`;
        }

        // Ranges, description, explanation from current node's incoming segment
        const displaySeg = _isRunning ? (next || seg) : seg;
        _ranges.innerHTML = '';
        (displaySeg?.sheet_context || []).forEach(addr => {
            const c = document.createElement('span');
            c.className   = 'range-chip';
            c.textContent = addr;
            _ranges.appendChild(c);
        });

        if (atRoot && !_isRunning) {
            _desc.textContent = 'Original sheet — no changes applied yet.';
            _expl.textContent = next ? `Next: ${next.description}` : '';
        } else if (displaySeg) {
            _desc.textContent = displaySeg.description || '';
            if (isFailed && displaySeg._errorMsg) {
                _expl.innerHTML = `<span style="opacity:0.75">${displaySeg.explanation || ''}</span>`
                    + `<div class="step-error-msg">⚠ ${displaySeg._errorMsg}</div>`;
            } else {
                _expl.textContent = displaySeg.explanation || '';
            }
        }

        // Counter: shows which node we're on. Root = 0/N (before step 1).
        _counter.textContent = `${idx}/${total}`;

        // Q&A from the incoming segment (what was just applied)
        _qaList.innerHTML = '';
        (seg?.qa_pairs || []).forEach(pair => {
            const item = document.createElement('details');
            item.className = 'qa-item';
            item.innerHTML = `<summary class="qa-q">${pair.q}</summary><p class="qa-a">${pair.a}</p>`;
            _qaList.appendChild(item);
        });

        _btnPrev.disabled = _isRunning || atRoot;

        if (_isRunning) {
            _btnNext.textContent = '…';
            _btnNext.disabled    = true;
        } else {
            _btnNext.textContent = atLeaf ? '✓' : '→';
            _btnNext.disabled    = !_advanceResolve && !isFailed;
        }

        _btnEdit.disabled = _isRunning || atRoot;
        _btnAsk.disabled  = _isRunning || atRoot;
    }

    // ── Graph render — git-style SVG ──────────────────────────────────────────
    //
    // Layout:
    //   Columns are evenly spaced horizontally (one column per main-chain node).
    //   Row 0 = main chain (horizontal).
    //   Rows 1, 2, … = edit branches, laid out diagonally (git-branch style).
    //   The fork connector is a diagonal line: (forkCol, row 0) → (forkCol+1, branchRow).
    //
    // Node states:
    //   current   — white filled circle with glow ring
    //   visited   — green filled circle
    //   unvisited — dim ghost circle

    const SVG_NS = 'http://www.w3.org/2000/svg';
    const NODE_R = 6;
    const COL_W  = 34;
    const ROW_H  = 34;  // equal to COL_W → 45° diagonal connectors
    const PAD_X  = 16;
    const PAD_Y  = 16;

    const NODE_STYLE = {
        current:   { fill: 'rgba(255,255,255,0.18)', stroke: '#ffffff',              strokeW: 2,   r: NODE_R },
        visited:   { fill: 'rgba(62,207,142,0.85)',  stroke: '#3ecf8e',              strokeW: 1.5, r: NODE_R },
        unvisited: { fill: 'rgba(255,255,255,0.07)', stroke: 'rgba(255,255,255,0.25)', strokeW: 1, r: NODE_R },
    };

    function _px(col) { return PAD_X + col * COL_W; }
    function _py(row) { return PAD_Y + row * ROW_H; }

    function _renderGraph() {
        if (!_graphEl || !_chainId) return;
        _graphEl.innerHTML = '';

        const graph = DagRunner.buildRenderGraph(_chainId);
        if (!graph.nodes.length) return;

        const svgW = PAD_X * 2 + graph.cols * COL_W;
        // Height: pad top + (rows-1 gaps) + pad bottom — no extra trailing row
        const svgH = PAD_Y * 2 + Math.max(0, graph.rows - 1) * ROW_H;

        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.setAttribute('viewBox', `0 0 ${svgW} ${svgH}`);
        // Let CSS width constrain the element; height scales proportionally
        svg.setAttribute('width',  '100%');
        svg.style.overflow = 'visible';

        // ── Edges (drawn first, behind nodes) ────────────────────────────────
        graph.edges.forEach(e => {
            const x1 = _px(e.fromCol), y1 = _py(e.fromRow);
            const x2 = _px(e.toCol),   y2 = _py(e.toRow);

            // Diagonal fork connector for branch edges that change row
            if (e.isBranch && e.fromRow !== e.toRow) {
                const isActiveFork = e.toId === graph.currentNodeId;
                const forkStroke = e.failed      ? '#f5643c'
                                 : isActiveFork  ? '#ffffff'
                                 : e.executed    ? '#3ecf8e'
                                 : 'rgba(255,255,255,0.22)';
                const fork = document.createElementNS(SVG_NS, 'line');
                fork.setAttribute('x1', x1); fork.setAttribute('y1', y1);
                fork.setAttribute('x2', x2); fork.setAttribute('y2', y2);
                fork.setAttribute('stroke',       forkStroke);
                fork.setAttribute('stroke-width', e.executed || e.failed ? '2' : '1.5');
                if (!e.executed && !e.failed) fork.setAttribute('stroke-dasharray', '3 2');
                svg.appendChild(fork);
                return;
            }

            // Horizontal edge colour priority:
            //   failed → red | active (leads to current node) → white
            //   executed → green | branch unexecuted → dim white | main unexecuted → dim blue
            const isActive = e.toId === graph.currentNodeId;
            let stroke, width, dash;
            if (e.failed) {
                stroke = '#f5643c'; width = '2.5'; dash = 'none';
            } else if (isActive) {
                stroke = '#ffffff'; width = '2.5'; dash = 'none';
            } else if (e.executed) {
                stroke = '#3ecf8e'; width = '2'; dash = 'none';
            } else if (e.isBranch) {
                stroke = 'rgba(255,255,255,0.15)'; width = '1.5'; dash = '3 2';
            } else {
                stroke = 'rgba(79,142,247,0.25)'; width = '1.5'; dash = '4 3';
            }

            const line = document.createElementNS(SVG_NS, 'line');
            line.setAttribute('x1', x1); line.setAttribute('y1', y1);
            line.setAttribute('x2', x2); line.setAttribute('y2', y2);
            line.setAttribute('stroke',       stroke);
            line.setAttribute('stroke-width', width);
            if (dash !== 'none') line.setAttribute('stroke-dasharray', dash);
            svg.appendChild(line);
        });

        // ── Nodes (drawn on top of edges) ─────────────────────────────────────
        graph.nodes.forEach(n => {
            const x  = _px(n.col);
            const y  = _py(n.row);
            const st = NODE_STYLE[n.state] || NODE_STYLE.unvisited;

            const dot = document.createElementNS(SVG_NS, 'circle');
            dot.setAttribute('cx', x);          dot.setAttribute('cy', y);
            dot.setAttribute('r',  st.r);
            dot.setAttribute('fill',         st.fill);
            dot.setAttribute('stroke',       st.stroke);
            dot.setAttribute('stroke-width', st.strokeW);
            dot.style.cursor = 'pointer';
            dot.addEventListener('click', () => _onNodeClick(n.id));

            const title = document.createElementNS(SVG_NS, 'title');
            title.textContent = n.label || n.id;
            dot.appendChild(title);

            svg.appendChild(dot);
        });

        _graphEl.style.overflowX = 'auto';
        _graphEl.appendChild(svg);
    }

    // ── Panels ────────────────────────────────────────────────────────────────

    function _togglePanel(name) {
        if (_activePanel === name) { _closePanel(); return; }
        _activePanel = name;
        _askPanel.style.display  = name === 'ask'  ? 'flex' : 'none';
        _editPanel.style.display = name === 'edit' ? 'flex' : 'none';
        RubricManager.showPanel(name === 'rubric');
        _btnAsk.classList.toggle('active',  name === 'ask');
        _btnEdit.classList.toggle('active', name === 'edit');
        if (name === 'edit') _populateEditPanel();
    }

    function _closePanel() {
        _activePanel = null;
        _askPanel.style.display  = 'none';
        _editPanel.style.display = 'none';
        RubricManager.showPanel(false);
        _btnAsk.classList.remove('active');
        _btnEdit.classList.remove('active');
    }

    // ── Ask panel ─────────────────────────────────────────────────────────────

    async function _onAskSend() {
        const msg = _askInput.value.trim();
        const chain0 = DagRunner.getChain(_chainId);
        const atLeaf0 = DagStore.edgesFrom(chain0.currentNodeId).length === 0;
        const seg0 = atLeaf0 ? _currentSegment(chain0) : _nextSegment(chain0);
        const hasParamChange = !!_collectParamChanges(seg0);
        if (!msg && !hasParamChange) return;
        _askInput.value    = '';
        _askInput.disabled = true;
        _askSend.disabled  = true;
        _appendAskBubble('user', msg);

        const chain = DagRunner.getChain(_chainId);
        const seg   = _currentSegment(chain);
        try {
            const wsCtx = await WorksheetContext.gather(['sheet']);
            const res   = await LLMClient.ask(msg, wsCtx,
                { description: seg?.description, explanation: seg?.explanation },
                _askHistory);
            _askHistory.push({ q: msg, a: res.answer });
            _appendAskBubble('agent', res.answer);
            _renderAskChips(res.follow_up_questions || []);
        } catch (err) {
            _appendAskBubble('agent', `⚠️ ${err.message}`);
        } finally {
            _askInput.disabled = false;
            _askSend.disabled  = false;
        }
    }

    function _appendAskBubble(role, text) {
        const d = document.createElement('div');
        d.className   = `ask-bubble ask-${role}`;
        d.textContent = text;
        _askThread.appendChild(d);
        _askThread.scrollTop = _askThread.scrollHeight;
    }

    function _renderAskChips(questions) {
        _askChips.innerHTML = '';
        questions.forEach(q => {
            const btn = document.createElement('button');
            btn.className   = 'ask-chip';
            btn.textContent = q;
            btn.onclick     = () => { _askInput.value = q; _onAskSend(); };
            _askChips.appendChild(btn);
        });
    }

    // ── Edit panel ────────────────────────────────────────────────────────────

    /** Populate suggestion chips and parameter inputs for the displayed segment. */
    function _populateEditPanel() {
        const chain = DagRunner.getChain(_chainId);
        const seg = _displayedSeg(chain);

        // ── Suggestion chips ───────────────────────────────────────────────
        _editChips.innerHTML = '';
        (seg?.edit_suggestions || []).forEach(suggestion => {
            const btn = document.createElement('button');
            btn.className   = 'edit-chip';
            btn.textContent = suggestion;
            btn.onclick     = () => { _editFeedback.value = suggestion; _editFeedback.focus(); };
            _editChips.appendChild(btn);
        });

        // ── Parameter controls ─────────────────────────────────────────────
        _editParams.innerHTML = '';
        const params = seg?.parameters || [];
        if (!params.length) return;

        const grid = document.createElement('div');
        grid.className = 'edit-params-grid';
        params.forEach((p, i) => {
            const row = document.createElement('div');
            row.className = 'edit-param-row';

            const lbl = document.createElement('label');
            lbl.className   = 'edit-param-label';
            lbl.textContent = p.label;

            let ctrl;
            if (p.type === 'select') {
                ctrl = document.createElement('select');
                ctrl.className = 'edit-param-select';
                (p.options || [p.value]).forEach(opt => {
                    const o = document.createElement('option');
                    o.value = opt; o.textContent = opt;
                    if (opt === String(p.value)) o.selected = true;
                    ctrl.appendChild(o);
                });
                ctrl.addEventListener('change', () => _applyParam(seg, i, ctrl.value, ctrl));
            } else if (p.type === 'color') {
                // Pair: color swatch picker + text input for hex
                const wrap = document.createElement('div');
                wrap.className = 'edit-param-color-wrap';

                const swatch = document.createElement('input');
                swatch.type  = 'color';
                swatch.value = p.value;
                swatch.className = 'edit-param-swatch';

                const hex = document.createElement('input');
                hex.type      = 'text';
                hex.value     = p.value;
                hex.className = 'edit-param-input';
                hex.style.width = '68px';

                swatch.addEventListener('input', () => {
                    hex.value = swatch.value;
                    _applyParam(seg, i, swatch.value, hex);
                });
                hex.addEventListener('input', () => {
                    if (/^#[0-9a-fA-F]{6}$/.test(hex.value)) {
                        swatch.value = hex.value;
                        _applyParam(seg, i, hex.value, hex);
                    }
                });

                wrap.appendChild(swatch);
                wrap.appendChild(hex);
                ctrl = wrap;
            } else {
                ctrl = document.createElement('input');
                ctrl.className   = 'edit-param-input';
                ctrl.type        = p.type === 'number' ? 'number' : 'text';
                ctrl.value       = p.value;
                ctrl.addEventListener('input', () => _applyParam(seg, i, ctrl.value, ctrl));
            }
            if (ctrl.dataset !== undefined) {
                ctrl.dataset.key = p.key;
                ctrl.dataset.idx = i;
            }

            row.appendChild(lbl);
            row.appendChild(ctrl);
            grid.appendChild(row);
        });
        _editParams.appendChild(grid);
    }

    /** Collect current parameter input values and build an auto-message prefix. */
    function _collectParamChanges(seg) {
        const params = seg?.parameters || [];
        if (!params.length) return '';
        const changes = [];
        params.forEach(p => {
            const inp = _editParams.querySelector(`[data-key="${p.key}"]`);
            if (!inp) return;
            const newVal = p.type === 'number' ? Number(inp.value) : inp.value;
            if (String(newVal) !== String(p.value)) {
                changes.push(`set ${p.label} to ${newVal}`);
            }
        });
        return changes.length ? changes.join(', ') + '. ' : '';
    }

    /**
     * Patch a single parameter value directly into seg.code and re-run it.
     * No LLM call — instant. Updates seg.parameters[idx].value so subsequent
     * opens of the Edit panel show the new value.
     */
    async function _applyParam(seg, idx, rawValue, btn) {
        if (!seg?.code) return;
        const p      = seg.parameters[idx];
        const oldVal = p.value;
        const newVal = p.type === 'number' ? Number(rawValue) : String(rawValue);
        if (String(newVal) === String(oldVal)) return;  // no change

        try {
            // Patch: replace the old literal with the new one in the code string.
            // Numbers: match the bare number; text/colors: match the quoted string.
            let patchedCode;
            if (p.type === 'number') {
                // Replace exact number literal (whole-word boundary)
                patchedCode = seg.code.replace(
                    new RegExp(`(?<![\d.])${_escapeRegex(String(oldVal))}(?![\d.])`, 'g'),
                    String(newVal)
                );
            } else {
                // Replace quoted string literal — try both quote styles
                const esc = _escapeRegex(String(oldVal));
                patchedCode = seg.code
                    .replace(new RegExp('"'  + esc + '"',  'g'), '"'  + newVal + '"')
                    .replace(new RegExp("'"  + esc + "'",  'g'), "'"  + newVal + "'");
            }

            // Run the patched code
            const fn = new (Object.getPrototypeOf(async function(){}).constructor)(patchedCode);
            await fn();

            // Persist the patch into the segment so it's used from now on
            seg.code          = patchedCode;
            seg.parameters[idx].value = newVal;

            inp.style.borderColor = 'rgba(62,207,142,0.8)';
            setTimeout(() => { inp.style.borderColor = ''; }, 800);
            ExecutionEngine.log('ok', `✓ Param "${p.label}" → ${newVal}`);
        } catch (err) {
            inp.style.borderColor = 'rgba(245,100,60,0.8)';
            setTimeout(() => { inp.style.borderColor = ''; }, 800);
        } finally {}
    }

    function _escapeRegex(str) {
        return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    async function _onEditSend() {
        const msg = _editFeedback.value.trim();
        if (!msg) return;

        const chain   = DagRunner.getChain(_chainId);
        const atLeaf  = DagStore.edgesFrom(chain.currentNodeId).length === 0;
        const seg     = _displayedSeg(chain);
        const fromIdx = atLeaf
            ? Math.max(0, _currentIndex(chain) - 1)
            : _currentIndex(chain);
        const remaining = chain.segments.slice(fromIdx + 1);

        _editSend.disabled    = true;
        _editSend.textContent = '…';
        _editSend.classList.add('loading');

        try {
            const paramPrefix = _collectParamChanges(seg);
            const fullMsg  = paramPrefix + (msg || 'Apply the parameter changes above.');
            const wsCtx    = await WorksheetContext.gather(['sheet']);
            const newChain = await LLMClient.edit(fullMsg, wsCtx, seg, remaining);
            _editFeedback.value = '';

            DagRunner.applyEdit(_chainId, fromIdx, chain.taskLabel, newChain);
            _editSend.textContent = '✓ Ready';

            // Don't call stepForward here — ExecutionEngine is already running
            // and owns the step loop. Resolving _advanceResolve lets it advance
            // naturally from the new fork position, running the edited step next.
            _render();
            _resolve();
        } catch (err) {
            _editSend.textContent = '⚠ Error';
            ExecutionEngine.log('err', `✗ Edit LLM error: ${err.message}`);
        } finally {
            _editSend.classList.remove('loading');
            setTimeout(() => {
                _editSend.textContent = 'Apply Edit';
                _editSend.disabled    = false;
            }, 1800);
        }
    }

    // ── Range focus ───────────────────────────────────────────────────────────

    async function _focusRanges() {
        const chain    = DagRunner.getChain(_chainId);
        // Focus the ranges of the segment just applied (incoming edge)
        const seg      = _currentSegment(chain) || _nextSegment(chain);
        const contexts = seg?.sheet_context;
        if (!contexts?.length) return;
        try {
            await Excel.run(async ctx => {
                ctx.workbook.worksheets.getActiveWorksheet()
                    .getRange(contexts.join(', ')).select();
                await ctx.sync();
            });
        } catch (err) {
            console.warn('[StepNavigator] focusRanges:', err.message);
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /**
     * The card shows the segment of the edge that LED INTO the current node —
     * i.e. the change that produced the sheet state the user is looking at.
     *
     * At the root node there is no incoming edge: return null so _renderCard
     * can show a "ready to begin" state instead.
     *
     * This matches the commit model: current node = committed state,
     * incoming edge = the change that was agreed to.
     */
    function _currentSegment(chain) {
        if (!chain) return null;
        const incoming = DagStore.edgesTo(chain.currentNodeId);
        if (!incoming.length) return null;   // root node — no step applied yet
        // Prefer the executed incoming edge; fall back to any incoming edge
        const e = incoming.find(ed => ed.executed) || incoming[0];
        return e?.segment || null;
    }

    /**
     * Segment for the NEXT step — the outgoing edge from currentNodeId.
     * Used by markRunning to show what is about to be applied.
     */
    function _nextSegment(chain) {
        if (!chain) return null;
        const outgoing   = DagStore.edgesFrom(chain.currentNodeId);
        const mainEdgeSet = new Set(chain.edgeIds);
        if (!outgoing.length) return null;
        return (outgoing.find(e => mainEdgeSet.has(e.id)) || outgoing[0]).segment || null;
    }

    /**
     * The segment the user is currently looking at — always _currentSegment
     * (the last applied step). This is what Edit and Ask should target.
     */
    function _displayedSeg(chain) {
        return _currentSegment(chain) || _nextSegment(chain);
    }

    /** 0-based index of currentNodeId in the chain's nodeIds list. */
    function _currentIndex(chain) {
        if (!chain) return 0;
        return Math.max(0, chain.nodeIds.indexOf(chain.currentNodeId));
    }

    return { init, load, markRunning, markFailed, waitForNext, dismiss, refreshGraph, setGateMode, dismissGate, isDismissed: () => _dismissed };
})();
