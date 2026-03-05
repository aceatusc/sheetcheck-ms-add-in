/**
 * stepNavigator.js
 *
 * Shows a floating card after each segment runs and waits for the user
 * to click "Next" before execution continues.
 *
 * Public API:
 *   StepNavigator.init()
 *   StepNavigator.loadSegments(segments)   — call once before execution starts
 *   StepNavigator.markRunning(index)       — call while a segment is executing
 *   StepNavigator.waitForNext(index)       — resolves when user clicks Next
 *   StepNavigator.dismiss()               — hides card without affecting execution
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
    let _segments       = [];
    let _completedUpTo  = -1;  // highest index fully executed
    let _currentIndex   = 0;   // step shown in the card right now
    let _isRunning      = false;
    let _advanceResolve = null; // set while waiting for user to click Next

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

    /** Load the full segment list before execution starts. */
    function loadSegments(segments) {
        _segments       = segments;
        _completedUpTo  = -1;
        _currentIndex   = 0;
        _isRunning      = false;
        _advanceResolve = null;
    }

    /** Show the card in blue "running" state while a segment's code executes. */
    function markRunning(index) {
        _isRunning    = true;
        _currentIndex = index;
        _overlay.classList.add('running', 'visible');
        _render();
    }

    /**
     * Switch to green "done" state and block until the user clicks Next.
     * @returns {Promise<void>}
     */
    async function waitForNext(index) {
        _isRunning     = false;
        _completedUpTo = Math.max(_completedUpTo, index);
        _currentIndex  = index;
        _overlay.classList.remove('running');
        _overlay.classList.add('visible');

        // !! Set the resolve BEFORE rendering so _render sees awaitingAdvance=true
        const promise = new Promise(resolve => { _advanceResolve = resolve; });

        _render();
        await _focusRanges(index);

        return promise;
    }

    /** Hide the card. Does NOT resolve a pending waitForNext. */
    function dismiss() {
        _overlay.classList.remove('visible', 'running');
    }

    // ── Private ───────────────────────────────────────────────────────────────

    function _onNext() {
        if (_isRunning) return;

        if (_currentIndex < _completedUpTo) {
            // Reviewing an earlier step — just move the card forward
            _navigate(_currentIndex + 1);
            return;
        }

        // On the latest completed step — advance execution
        if (_advanceResolve) {
            const resolve   = _advanceResolve;
            _advanceResolve = null;

            const isLast = _currentIndex >= _segments.length - 1;
            if (isLast) dismiss();

            resolve(); // unblocks executionEngine
        }
    }

    function _onPrev() {
        if (_isRunning || _currentIndex <= 0) return;
        _navigate(_currentIndex - 1);
    }

    async function _navigate(targetIndex) {
        if (targetIndex < 0 || targetIndex > _completedUpTo) return;
        _currentIndex = targetIndex;
        _render();
        await _focusRanges(targetIndex);
    }

    function _render() {
        const seg   = _segments[_currentIndex];
        const total = _segments.length;
        const idx   = _currentIndex;
        if (!seg) return;

        const isLast           = idx >= total - 1;
        const isLatestComplete = idx === _completedUpTo;
        const awaitingAdvance  = _advanceResolve !== null;

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

        // ← disabled while running or on first step
        _btnPrev.disabled = _isRunning || idx <= 0;

        // → behaviour:
        if (_isRunning) {
            _btnNext.textContent = '…';
            _btnNext.disabled    = true;
        } else if (isLatestComplete && awaitingAdvance) {
            // Ready to advance execution
            _btnNext.textContent = isLast ? 'Done ✓' : 'Next →';
            _btnNext.disabled    = false;
        } else {
            // Reviewing a past step — just a card navigation arrow
            _btnNext.textContent = '→';
            _btnNext.disabled    = idx >= _completedUpTo;
        }
    }

    async function _focusRanges(index) {
        const contexts = _segments[index]?.sheet_context;
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
