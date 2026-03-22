/**
 * aspectManager.js
 *
 * Standalone "Verify" panel — accessible at any time via the Verify button
 * in the chat header. Completely independent of the code generation pipeline.
 *
 * Features:
 *   - Flat list of "aspects" (no hard/soft distinction)
 *   - "Populate" button → calls /rubric/scaffold with full sheet + chat context
 *   - "Verify" button   → calls /rubric/verify and shows results with clickable
 *                         cell references (clicking focuses that range in Excel)
 *   - Add / edit / delete aspects manually at any time
 *
 * DOM: injects its own overlay element into <body> on init().
 */
const AspectManager = (() => {

    // ── State ─────────────────────────────────────────────────────────────────
    let _aspects   = [];   // [{ id, label }]
    let _overlay   = null;
    let _listEl    = null;
    let _resultsEl = null;
    let _populateBtn = null;
    let _verifyBtn   = null;

    // ── Init ──────────────────────────────────────────────────────────────────

    function init() {
        _buildOverlay();
    }

    function _buildOverlay() {
        _overlay = document.createElement('div');
        _overlay.id        = 'aspect-overlay';
        _overlay.className = 'aspect-overlay';
        _overlay.innerHTML = `
            <div class="aspect-header">
                <span class="aspect-title">✦ Verify</span>
                <button class="aspect-close" id="aspect-close-btn">✕</button>
            </div>
            <p class="aspect-subtitle">
                Review important aspects of your sheet. Populate based on context or add your own, then run Verify to check each one.
            </p>
            <div class="aspect-toolbar">
                <button class="aspect-btn aspect-btn--populate" id="aspect-populate-btn">
                    <span class="aspect-btn-icon">⟳</span> Populate
                </button>
                <button class="aspect-btn aspect-btn--verify" id="aspect-verify-btn">
                    <span class="aspect-btn-icon">✓</span> Verify
                </button>
                <button class="aspect-btn aspect-btn--add" id="aspect-add-btn">
                    <span class="aspect-btn-icon">+</span> Add
                </button>
            </div>
            <div id="aspect-list" class="aspect-list"></div>
            <div id="aspect-results" class="aspect-results"></div>
        `;

        document.body.appendChild(_overlay);

        _listEl      = _overlay.querySelector('#aspect-list');
        _resultsEl   = _overlay.querySelector('#aspect-results');
        _populateBtn = _overlay.querySelector('#aspect-populate-btn');
        _verifyBtn   = _overlay.querySelector('#aspect-verify-btn');

        _overlay.querySelector('#aspect-close-btn').addEventListener('click', close);
        _populateBtn.addEventListener('click', _onPopulate);
        _verifyBtn.addEventListener('click', _onVerify);
        _overlay.querySelector('#aspect-add-btn').addEventListener('click', _onAdd);

        // Click outside to close
        _overlay.addEventListener('click', e => {
            if (e.target === _overlay) close();
        });
    }

    // ── Public ────────────────────────────────────────────────────────────────

    function open() {
        _overlay.classList.add('visible');
        document.getElementById('chat-panel')?.classList.add('nav-active');
        _render();
    }

    function close() {
        _overlay.classList.remove('visible');
        document.getElementById('chat-panel')?.classList.remove('nav-active');
    }

    // ── Populate ──────────────────────────────────────────────────────────────

    async function _onPopulate() {
        _populateBtn.disabled    = true;
        _populateBtn.textContent = '⟳ Populating…';
        _resultsEl.innerHTML     = '';

        try {
            const wsCtx = await WorksheetContext.gather(['selection', 'sheet', 'styles', 'charts']);
            // Use last chat message as the user_message context; fall back to empty string
            const history = ChatHistory.get();
            const lastMsg = history[history.length - 1] || '';
            const res = await LLMClient.rubricScaffold(lastMsg, wsCtx, history);

            // res is { aspects: [{id, label}, ...] }
            if (Array.isArray(res.aspects) && res.aspects.length) {
                _aspects = res.aspects;
            } else {
                // Fallback: old hard/soft shape from server if not yet updated
                const all = [
                    ...(res.hard_requirements || []),
                    ...(res.soft_requirements || []),
                ];
                _aspects = all.map(r => ({ id: r.id, label: r.label }));
            }
            _render();
        } catch (err) {
            _listEl.innerHTML = `<span class="aspect-error">Populate failed: ${err.message}</span>`;
        } finally {
            _populateBtn.disabled    = false;
            _populateBtn.innerHTML   = '<span class="aspect-btn-icon">⟳</span> Populate';
        }
    }

    // ── Verify ────────────────────────────────────────────────────────────────

    async function _onVerify() {
        if (!_aspects.length) {
            _resultsEl.innerHTML = '<span class="aspect-error">Add some aspects first, then verify.</span>';
            return;
        }

        _verifyBtn.disabled    = true;
        _verifyBtn.textContent = '✓ Verifying…';
        _resultsEl.innerHTML   = '<span class="aspect-loading">Checking aspects against your sheet…</span>';

        try {
            const wsCtx = await WorksheetContext.gather(['sheet']);
            const payload = { aspects: _aspects };
            const res = await LLMClient.rubricVerify(payload, wsCtx, ChatHistory.get());

            // Build lookup by id
            const byId = {};
            (res.results || []).forEach(r => { byId[r.id] = r; });

            _resultsEl.innerHTML = '';

            // Score line
            const met   = _aspects.filter(a => byId[a.id]?.met).length;
            const total = _aspects.length;
            const allMet = met === total;

            const score = document.createElement('div');
            score.className = 'aspect-score';
            score.innerHTML =
                `<span class="aspect-score-num ${allMet ? 'all-met' : met === 0 ? 'none-met' : ''}">`
                + `${met}<span class="aspect-score-denom">/${total}</span></span>`
                + `<span class="aspect-score-label">${allMet ? 'All aspects satisfied 🎉' : 'aspects satisfied'}</span>`;
            _resultsEl.appendChild(score);

            // Result rows
            _aspects.forEach(aspect => {
                const r   = byId[aspect.id];
                const met = r?.met ?? false;

                const row = document.createElement('div');
                row.className = `aspect-result-row ${met ? 'met' : 'unmet'}`;

                const icon = document.createElement('span');
                icon.className   = 'aspect-result-icon';
                icon.textContent = met ? '✓' : '⚠';

                const body = document.createElement('div');
                body.className = 'aspect-result-body';

                const label = document.createElement('div');
                label.className   = 'aspect-result-label';
                label.textContent = aspect.label;

                body.appendChild(label);

                if (r?.reasoning) {
                    const reason = document.createElement('div');
                    reason.className   = 'aspect-result-reason';
                    reason.textContent = r.reasoning;
                    body.appendChild(reason);
                }

                // Clickable cell references
                if (r?.references?.length) {
                    const refs = document.createElement('div');
                    refs.className = 'aspect-result-refs';
                    r.references.forEach(addr => {
                        const chip = document.createElement('button');
                        chip.className   = 'aspect-ref-chip';
                        chip.textContent = addr;
                        chip.title       = `Focus ${addr} in sheet`;
                        chip.addEventListener('click', () => _focusRange(addr));
                        refs.appendChild(chip);
                    });
                    body.appendChild(refs);
                }

                row.appendChild(icon);
                row.appendChild(body);
                _resultsEl.appendChild(row);
            });

        } catch (err) {
            _resultsEl.innerHTML = `<span class="aspect-error">Verification failed: ${err.message}</span>`;
        } finally {
            _verifyBtn.disabled  = false;
            _verifyBtn.innerHTML = '<span class="aspect-btn-icon">✓</span> Verify';
        }
    }

    // ── Focus range in Excel ──────────────────────────────────────────────────

    async function _focusRange(address) {
        try {
            await Excel.run(async ctx => {
                ctx.workbook.worksheets.getActiveWorksheet()
                    .getRange(address).select();
                await ctx.sync();
            });
        } catch (err) {
            console.warn('[AspectManager] focusRange failed:', err.message);
        }
    }

    // ── Add aspect ────────────────────────────────────────────────────────────

    function _onAdd() {
        const id = 'a' + Date.now();
        _aspects.push({ id, label: '' });
        _render();
        // Focus the new row's input
        requestAnimationFrame(() => {
            const input = _listEl.querySelector(`[data-id="${id}"] .aspect-label-input`);
            if (input) input.focus();
        });
    }

    // ── Render ────────────────────────────────────────────────────────────────

    function _render() {
        _listEl.innerHTML = '';

        if (!_aspects.length) {
            const hint = document.createElement('div');
            hint.className   = 'aspect-empty';
            hint.textContent = 'No aspects yet — click Populate or Add.';
            _listEl.appendChild(hint);
            return;
        }

        _aspects.forEach((aspect, idx) => {
            const row = document.createElement('div');
            row.className    = 'aspect-row';
            row.dataset.id   = aspect.id;

            const num = document.createElement('span');
            num.className   = 'aspect-num';
            num.textContent = idx + 1;

            const input = document.createElement('input');
            input.type        = 'text';
            input.className   = 'aspect-label-input';
            input.value       = aspect.label;
            input.placeholder = 'Describe an aspect to check…';
            input.addEventListener('input', () => {
                const a = _aspects.find(x => x.id === aspect.id);
                if (a) a.label = input.value;
            });
            input.addEventListener('keydown', e => {
                if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
            });

            const del = document.createElement('button');
            del.className   = 'aspect-del';
            del.textContent = '✕';
            del.title       = 'Remove';
            del.addEventListener('click', () => {
                _aspects = _aspects.filter(x => x.id !== aspect.id);
                _render();
            });

            row.appendChild(num);
            row.appendChild(input);
            row.appendChild(del);
            _listEl.appendChild(row);
        });
    }

    return { init, open, close };
})();
