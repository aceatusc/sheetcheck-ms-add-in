/**
 * aspectManager.js
 *
 * Standalone "Verify" panel. DOM is declared in index.html.
 *
 * After Verify runs, results appear inline under each aspect row:
 *   ✓/⚠  reasoning text
 *   [Sheet1!A1:F1] clickable range chips
 * No separate results section — saves space and avoids repetition.
 */
const AspectManager = (() => {

    // ── State ─────────────────────────────────────────────────────────────────
    let _aspects = [];   // [{ id, label, result?, loading? }]
    let _closeResolve = null;  // resolves waitForClose()

    // ── DOM refs (elements declared in index.html) ────────────────────────────
    let _overlay, _listEl, _populateBtn;

    // ── Init ──────────────────────────────────────────────────────────────────

    function init() {
        _overlay     = document.getElementById('aspect-overlay');
        _listEl      = document.getElementById('aspect-list');
        _populateBtn = document.getElementById('aspect-populate-btn');

        document.getElementById('aspect-confirm-btn').addEventListener('click', close);
        _populateBtn.addEventListener('click', _onPopulate);
        // _verifyBtn.addEventListener('click', _onVerify);
        document.getElementById('aspect-add-btn').addEventListener('click', _onAdd);
    }

    // ── Public ────────────────────────────────────────────────────────────────

    function open() {
        _overlay.classList.add('visible');
        _overlay.classList.remove('closing');
        document.getElementById('chat-panel')?.classList.add('nav-active');
        _render();
        if (!_aspects.length) _onPopulate();
    }

    /**
     * Returns a promise that resolves when the panel is closed.
     * ChatManager awaits this before proceeding with the LLM call.
     */
    function waitForClose() {
        return new Promise(resolve => { _closeResolve = resolve; });
    }

    /**
     * Populate aspects in the background without opening the panel.
     */
    async function populate() {
        await _onPopulate();
    }

    function close() {
        // Animate the panel shrinking toward the verify button
        const btn = document.getElementById('verify-btn');
        if (btn) {
            const br = btn.getBoundingClientRect();
            const or = _overlay.getBoundingClientRect();
            const tx = (br.left + br.width  / 2) - (or.left + or.width  / 2);
            const ty = (br.top  + br.height / 2) - (or.top  + or.height / 2);
            _overlay.style.setProperty('--close-tx', `${tx}px`);
            _overlay.style.setProperty('--close-ty', `${ty}px`);
        }
        _overlay.classList.add('closing');
        _overlay.classList.remove('visible');
        document.getElementById('chat-panel')?.classList.remove('nav-active');

        // Resolve any pending waitForClose promise after the animation finishes
        setTimeout(() => {
            _overlay.classList.remove('closing');
            if (_closeResolve) { _closeResolve(); _closeResolve = null; }
        }, 320);
    }

    // ── Populate ──────────────────────────────────────────────────────────────

    async function _onPopulate() {
        _populateBtn.disabled  = true;
        _populateBtn.innerHTML = '<span class="aspect-btn-icon">✦</span> Generating with AI…';
        _aspects = _aspects.map(a => ({ id: a.id, label: a.label }));

        try {
            // Include all sheets so aspects can reference cross-sheet data
            const wsCtx   = await WorksheetContext.gather(['selection', 'sheet', 'styles', 'charts']);
            const history = ChatHistory.get();   // full conversation history
            const lastMsg = history[history.length - 1] || '';
            const res     = await LLMClient.rubricScaffold(lastMsg, wsCtx, history);

            if (Array.isArray(res.aspects) && res.aspects.length) {
                _aspects = res.aspects.map(a => ({ id: a.id, label: a.label }));
            } else {
                const all = [...(res.hard_requirements || []), ...(res.soft_requirements || [])];
                _aspects  = all.map(r => ({ id: r.id, label: r.label }));
            }
            _render();
            _populateBtn.innerHTML = '<span class="aspect-btn-icon">✦</span> Regenerate with AI';
        } catch (err) {
            _listEl.innerHTML = `<span class="aspect-error">Populate failed: ${err.message}</span>`;
            _populateBtn.innerHTML = '<span class="aspect-btn-icon">✦</span> Regenerate with AI';
        } finally {
            _populateBtn.disabled = false;
        }
    }

    // ── Verify ────────────────────────────────────────────────────────────────

    async function _onVerify() {
        if (!_aspects.length) {
            _listEl.innerHTML = '<span class="aspect-error">Add some aspects first, then verify.</span>';
            return;
        }

        // _verifyBtn.disabled  = true;
        // _verifyBtn.innerHTML = '<span class="aspect-btn-icon">✓</span> Verifying…';

        _aspects = _aspects.map(a => ({ id: a.id, label: a.label, loading: true }));
        _render();

        try {
            // Gather ALL sheets so the LLM can verify cross-sheet references
            // (e.g. dropdown sources on a Data sheet, VLOOKUP targets, etc.)
            const wsCtx   = await WorksheetContext.gather(['sheet']);
            const payload = { aspects: _aspects.map(a => ({ id: a.id, label: a.label })) };
            // Pass the FULL chat history so the LLM knows everything that was
            // requested and applied — not just the last message.
            const res     = await LLMClient.rubricVerify(payload, wsCtx, ChatHistory.get());

            const byId = {};
            (res.results || []).forEach(r => { byId[r.id] = r; });

            _aspects = _aspects.map(a => ({
                id:     a.id,
                label:  a.label,
                result: byId[a.id] || null,
            }));
            _render();
        } catch (err) {
            _aspects = _aspects.map(a => ({ id: a.id, label: a.label }));
            _render();
            const errEl = document.createElement('span');
            errEl.className   = 'aspect-error';
            errEl.textContent = `Verification failed: ${err.message}`;
            _listEl.appendChild(errEl);
        } finally {
            // _verifyBtn.disabled  = false;
            // _verifyBtn.innerHTML = '<span class="aspect-btn-icon">✓</span> Verify';
        }
    }

    // ── Add aspect ────────────────────────────────────────────────────────────

    function _onAdd() {
        const id = 'a' + Date.now();
        _aspects.push({ id, label: '' });
        _render();
        requestAnimationFrame(() => {
            const ta = _listEl.querySelector(`[data-id="${id}"] .aspect-label-input`);
            if (ta) ta.focus();
        });
    }

    // ── Render ────────────────────────────────────────────────────────────────

    function _render() {
        _listEl.innerHTML = '';

        if (!_aspects.length) {
            const hint = document.createElement('div');
            hint.className   = 'aspect-empty';
            hint.textContent = 'No aspects yet — wait for auto-population or Add.';
            _listEl.appendChild(hint);
            return;
        }

        _aspects.forEach((aspect, idx) => {
            const row = document.createElement('div');
            row.className  = 'aspect-row';
            row.dataset.id = aspect.id;

            // Left column: result icon (after verify) or index number
            const left = document.createElement('div');
            left.className = 'aspect-row-left';
            if (aspect.result) {
                const icon = document.createElement('span');
                icon.className   = aspect.result.met ? 'aspect-result-icon met' : 'aspect-result-icon unmet';
                icon.textContent = aspect.result.met ? '✓' : '⚠';
                left.appendChild(icon);
            } else {
                const num = document.createElement('span');
                num.className   = 'aspect-num';
                num.textContent = idx + 1;
                left.appendChild(num);
            }

            // Body: textarea + inline result detail
            const body = document.createElement('div');
            body.className = 'aspect-row-body';

            const ta = document.createElement('textarea');
            ta.className   = 'aspect-label-input';
            ta.value       = aspect.label;
            ta.placeholder = 'Describe an aspect to check…';
            ta.rows        = 1;
            const _resize  = () => { ta.style.height = 'auto'; ta.style.height = ta.scrollHeight + 'px'; };
            ta.addEventListener('input', () => {
                const a = _aspects.find(x => x.id === aspect.id);
                if (a) a.label = ta.value;
                _resize();
            });
            ta.addEventListener('keydown', e => {
                if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); ta.blur(); }
            });
            requestAnimationFrame(_resize);
            body.appendChild(ta);

            // Loading state
            if (aspect.loading) {
                const loading = document.createElement('div');
                loading.className   = 'aspect-inline-loading';
                loading.textContent = 'Checking…';
                body.appendChild(loading);
            }

            // Result detail: reasoning + clickable refs
            if (aspect.result) {
                const detail = document.createElement('div');
                detail.className = 'aspect-inline-detail';

                if (aspect.result.reasoning) {
                    const reason = document.createElement('span');
                    reason.className   = 'aspect-inline-reason';
                    reason.textContent = aspect.result.reasoning;
                    detail.appendChild(reason);
                }

                if (aspect.result.references?.length) {
                    const refs = document.createElement('div');
                    refs.className = 'aspect-result-refs';
                    aspect.result.references.forEach(addr => {
                        const chip = document.createElement('button');
                        chip.className   = 'aspect-ref-chip';
                        chip.textContent = addr;
                        chip.title       = `Focus ${addr}`;
                        chip.addEventListener('click', () => _focusRange(addr));
                        refs.appendChild(chip);
                    });
                    detail.appendChild(refs);
                }

                body.appendChild(detail);
            }

            // Delete button
            const del = document.createElement('button');
            del.className   = 'aspect-del';
            del.textContent = '✕';
            del.title       = 'Remove';
            del.addEventListener('click', () => {
                _aspects = _aspects.filter(x => x.id !== aspect.id);
                _render();
            });

            row.appendChild(left);
            row.appendChild(body);
            row.appendChild(del);
            _listEl.appendChild(row);
        });
    }

    // ── Focus range in Excel ──────────────────────────────────────────────────

    async function _focusRange(address) {
        const clean = address.includes('!') ? address.split('!')[1] : address;
        try {
            await Excel.run(async ctx => {
                ctx.workbook.worksheets.getActiveWorksheet().getRange(clean).select();
                await ctx.sync();
            });
        } catch (err) {
            console.warn('[AspectManager] focusRange failed:', err.message);
        }
    }

    return { init, open, close, populate, waitForClose };
})();
