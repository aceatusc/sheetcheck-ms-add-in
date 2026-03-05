/**
 * stepNavigator.js
 *
 * Manages the floating step-by-step explanation overlay.
 * After each segment runs, shows a card and WAITS for the user
 * to click → (or "Done" on the last step) before execution continues.
 *
 * Public API:
 *   StepNavigator.init()
 *   StepNavigator.loadSegments(segments)      — call once before execution starts
 *   StepNavigator.markRunning(index)          — call while a segment is executing
 *   StepNavigator.waitForNext(index)          — call after segment completes;
 *                                               resolves when user clicks →
 *   StepNavigator.dismiss()
 */
const StepNavigator = (() => {

    // ── DOM refs ──────────────────────────────────────────────────────────────
    const _overlay  = document.getElementById('step-navigator');
    const _badge    = document.getElementById('step-nav-badge');
    const _ranges   = document.getElementById('step-nav-ranges');
    const _desc     = document.getElementById('step-nav-description');
    const _expl     = document.getElementById('step-nav-explanation');
    const _counter  = document.getElementById('step-nav-counter');
    const _btnPrev  = document.getElementById('step-nav-prev');
    const _btnNext  = document.getElementById('step-nav-next');
    const _btnEdit  = document.getElementById('step-nav-edit');
    const _btnClose = document.getElementById('step-nav-close');

    // ── State ─────────────────────────────────────────────────────────────────
    let _segments      = [];
    let _completedUpTo = -1;   // highest index fully run
    let _currentIndex  = 0;    // step currently shown in the card
    let _isRunning     = false;

    // Resolve fn for the current waitForNext() promise
    let _advanceResolve = null;

    // ── Public ────────────────────────────────────────────────────────────────

    function init() {
        _btnPrev.addEventListener('click', _onPrev);
        _btnNext.addEventListener('click', _onNext);
        _btnClose.addEventListener('click', dismiss);
        _btnEdit.addEventListener('click', () => {
            // PLACEHOLDER: open an edit dialog for the current segment's code
            console.log('[StepNavigator] Edit clicked for segment:', _segments[_currentIndex]?.id);
        });
    }

    /** Load segment list before execution starts. */
    function loadSegments(segments) {
        _segments      = segments;
        _completedUpTo = -1;
        _currentIndex  = 0;
        _isRunning     = false;
        _advanceResolve = null;
    }

    /** Show the card in "running" (blue pulse) state while code executes. */
    function markRunning(index) {
        _isRunning    = true;
        _currentIndex = index;
        _render();
        _overlay.classList.add('running');
        _overlay.classList.add('visible');
    }

    /**
     * Switch card to "complete" (green) state, focus ranges, then
     * block until the user clicks →.
     * @returns {Promise<void>} resolves when user is ready for next step
     */
    async function waitForNext(index) {
        _isRunning     = false;
        _completedUpTo = Math.max(_completedUpTo, index);
        _currentIndex  = index;

        _overlay.classList.remove('running');
        _render();
        _overlay.classList.add('visible');

        await _focusRanges(index);

        // If this is the last step, → resolves immediately after user clicks Done
        return new Promise(resolve => {
            _advanceResolve = resolve;
        });
    }

    /** Hide the overlay and resolve any pending wait (e.g. user closed mid-run). */
    function dismiss() {
        _overlay.classList.remove('visible', 'running');
        if (_advanceResolve) {
            _advanceResolve();
            _advanceResolve = null;
        }
    }

    // ── Private ───────────────────────────────────────────────────────────────

    function _onNext() {
        const isLastStep = _currentIndex >= _segments.length - 1;

        if (_currentIndex === _completedUpTo) {
            // We're on the latest completed step → advance execution
            if (_advanceResolve) {
                const res = _advanceResolve;
                _advanceResolve = null;
                if (isLastStep) dismiss();
                res();
            }
        } else {
            // We're reviewing an earlier step → just navigate forward
            _navigate(_currentIndex + 1);
        }
    }

    function _onPrev() {
        if (_currentIndex > 0) _navigate(_currentIndex - 1);
    }

    /** Navigate card to a different already-completed step (no execution effect). */
    async function _navigate(targetIndex) {
        if (targetIndex < 0 || targetIndex > _completedUpTo) return;
        _currentIndex = targetIndex;
        _render();
        await _focusRanges(targetIndex);
    }

    /** Re-render card content for _currentIndex. */
    function _render() {
        const seg   = _segments[_currentIndex];
        const total = _segments.length;
        const idx   = _currentIndex;
        if (!seg) return;

        const isLastStep     = idx >= total - 1;
        const isLatestStep   = idx === _completedUpTo;
        const awaitingAdvance = !!_advanceResolve;

        // Badge
        _badge.textContent = _isRunning
            ? `Running step ${idx + 1} of ${total}…`
            : `Step ${idx + 1} of ${total}`;

        // Range chips
        _ranges.innerHTML = '';
        (seg.sheet_context || []).forEach(addr => {
            const chip = document.createElement('span');
            chip.className   = 'range-chip';
            chip.textContent = addr;
            _ranges.appendChild(chip);
        });

        // Text
        _desc.textContent = seg.description;
        _expl.textContent = seg.explanation || '';

        // Counter
        _counter.textContent = `${idx + 1}/${total}`;

        // ← always disabled on first step; disabled while running
        _btnPrev.disabled = idx <= 0 || _isRunning;

        // → label and state:
        //   • "Running…" — code is executing
        //   • "Done"     — last step, waiting for user confirm
        //   • "Next →"   — more steps remaining, waiting for user
        //   • "→"        — reviewing an earlier step (just navigates)
        if (_isRunning) {
            _btnNext.textContent = '…';
            _btnNext.disabled    = true;
        } else if (isLatestStep && awaitingAdvance) {
            _btnNext.textContent = isLastStep ? 'Done ✓' : 'Next →';
            _btnNext.disabled    = false;
        } else {
            _btnNext.textContent = '→';
            _btnNext.disabled    = idx >= _completedUpTo;
        }
    }

    /** Select the sheet ranges for this step in Excel so it scrolls + highlights. */
    async function _focusRanges(index) {
        const seg      = _segments[index];
        const contexts = seg?.sheet_context;
        if (!contexts || contexts.length === 0) return;

        try {
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();
                sheet.getRange(contexts.join(', ')).select();
                await ctx.sync();
            });
        } catch (err) {
            console.warn('[StepNavigator] Could not focus range:', err.message);
        }
    }

    return { init, loadSegments, markRunning, waitForNext, dismiss };
})();
