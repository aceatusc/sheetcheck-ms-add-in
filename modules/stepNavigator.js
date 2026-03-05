/**
 * stepNavigator.js
 *
 * Manages the floating step-by-step explanation overlay.
 * After each segment runs, shows a tooltip-style card with:
 *   - the step's description & explanation
 *   - which sheet ranges were affected (as chips)
 *   - prev / next navigation to move between completed steps
 *   - on navigation: selects & scrolls to the relevant range in Excel
 *
 * Public API:
 *   StepNavigator.init()
 *   StepNavigator.loadSegments(segments)   — call once before execution starts
 *   StepNavigator.markComplete(index)      — call after each segment finishes
 *   StepNavigator.markRunning(index)       — call while a segment is executing
 *   StepNavigator.dismiss()
 */
const StepNavigator = (() => {

    // ── DOM refs ──────────────────────────────────────────────────────────────
    const _overlay    = document.getElementById('step-navigator');
    const _badge      = document.getElementById('step-nav-badge');
    const _ranges     = document.getElementById('step-nav-ranges');
    const _desc       = document.getElementById('step-nav-description');
    const _expl       = document.getElementById('step-nav-explanation');
    const _counter    = document.getElementById('step-nav-counter');
    const _btnPrev    = document.getElementById('step-nav-prev');
    const _btnNext    = document.getElementById('step-nav-next');
    const _btnEdit    = document.getElementById('step-nav-edit');
    const _btnClose   = document.getElementById('step-nav-close');

    // ── State ─────────────────────────────────────────────────────────────────
    let _segments      = [];   // full CodeSegment[]
    let _completedUpTo = -1;   // highest index that has finished running
    let _currentIndex  = 0;    // which step is currently displayed
    let _isRunning     = false;

    // ── Public ────────────────────────────────────────────────────────────────

    function init() {
        _btnPrev.addEventListener('click', () => _navigate(_currentIndex - 1));
        _btnNext.addEventListener('click', () => _navigate(_currentIndex + 1));
        _btnClose.addEventListener('click', dismiss);
        _btnEdit.addEventListener('click', () => {
            // PLACEHOLDER: open an edit dialog for the current segment's code
            console.log('[StepNavigator] Edit clicked for segment:', _segments[_currentIndex]?.id);
        });
    }

    /**
     * Load the full segment list before execution starts.
     * @param {CodeSegment[]} segments
     */
    function loadSegments(segments) {
        _segments      = segments;
        _completedUpTo = -1;
        _currentIndex  = 0;
        _isRunning     = false;
    }

    /**
     * Call while segment at `index` is actively running.
     * Shows the overlay in "running" state (blue pulse).
     */
    function markRunning(index) {
        _isRunning    = true;
        _currentIndex = index;
        _render();
        _overlay.classList.add('running');
        _show();
    }

    /**
     * Call after segment at `index` has successfully completed.
     * Switches to "complete" state (green) and focuses the range in Excel.
     */
    async function markComplete(index) {
        _isRunning     = false;
        _completedUpTo = Math.max(_completedUpTo, index);
        _currentIndex  = index;
        _overlay.classList.remove('running');
        _render();
        _show();
        await _focusRanges(index);
    }

    /** Hide the overlay. */
    function dismiss() {
        _overlay.classList.remove('visible', 'running');
    }

    // ── Private ───────────────────────────────────────────────────────────────

    function _show() {
        _overlay.classList.add('visible');
    }

    /** Navigate to a different (already-completed) step. */
    async function _navigate(targetIndex) {
        if (targetIndex < 0 || targetIndex > _completedUpTo) return;
        _currentIndex = targetIndex;
        _render();
        await _focusRanges(targetIndex);
    }

    /** Re-render all overlay content for _currentIndex. */
    function _render() {
        const seg   = _segments[_currentIndex];
        const total = _segments.length;
        const idx   = _currentIndex;

        if (!seg) return;

        // Badge
        _badge.textContent = _isRunning
            ? `Running step ${idx + 1} of ${total}…`
            : `Step ${idx + 1} of ${total}`;

        // Range chips
        _ranges.innerHTML = '';
        const contexts = seg.sheet_context || [];
        contexts.forEach(addr => {
            const chip = document.createElement('span');
            chip.className   = 'range-chip';
            chip.textContent = addr;
            _ranges.appendChild(chip);
        });

        // Text content
        _desc.textContent = seg.description;
        _expl.textContent = seg.explanation || '';

        // Counter
        _counter.textContent = `${idx + 1}/${total}`;

        // Arrow states
        _btnPrev.disabled = idx <= 0;
        _btnNext.disabled = idx >= _completedUpTo; // can't go forward past completed
    }

    /**
     * Select and scroll to the sheet_context ranges for the given segment index.
     * Uses the first range to scroll, then selects all as a union if multiple.
     */
    async function _focusRanges(index) {
        const seg      = _segments[index];
        const contexts = seg?.sheet_context;
        if (!contexts || contexts.length === 0) return;

        try {
            await Excel.run(async (ctx) => {
                const sheet = ctx.workbook.worksheets.getActiveWorksheet();

                // Build a union range string (e.g. "A1:E1, B2:D7")
                const unionAddress = contexts.join(', ');
                const range = sheet.getRange(unionAddress);

                // Select so Excel scrolls to it and highlights it
                range.select();

                await ctx.sync();
            });
        } catch (err) {
            // Non-fatal — range focus is best-effort
            console.warn('[StepNavigator] Could not focus range:', err.message);
        }
    }

    return { init, loadSegments, markRunning, markComplete, dismiss };
})();
