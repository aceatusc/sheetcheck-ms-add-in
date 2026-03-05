/**
 * executionEngine.js
 * Runs code segments sequentially with pauses between each.
 * Updates the execution log panel and progress bar.
 */
const ExecutionEngine = (() => {

    const _logEl         = document.getElementById('execution-log');
    const _statusDot     = document.getElementById('execution-status-dot');
    const _barFill       = document.getElementById('segment-bar-fill');
    const _barLabel      = document.getElementById('segment-label');
    const _progressStrip = document.getElementById('segment-progress');
    const _panel         = document.getElementById('execution-panel');

    /**
     * Run an array of CodeSegments in order, pausing between each.
     * @param {CodeSegment[]} segments
     */
    async function run(segments) {
        if (!segments || segments.length === 0) return;

        _openPanel();
        _setStatus('running');
        _progressStrip.classList.add('visible');
        _log('info', `Starting execution: ${segments.length} segment(s)`);

        let completed = 0;

        for (const seg of segments) {
            _updateProgress(completed, segments.length);
            _log('info', `▶ ${seg.description}`);

            try {
                // Execute the segment's code string in an async context.
                // PLACEHOLDER: consider sandboxing / validation before eval.
                const fn = _makeAsyncFn(seg.code);
                await fn();

                completed++;
                _updateProgress(completed, segments.length);
                _log('ok', `✓ ${seg.description}`);

            } catch (err) {
                _log('err', `✗ ${seg.description}: ${err.message}`);
                _setStatus('error');
                console.error('[ExecutionEngine] Segment error:', err);
                break; // PLACEHOLDER: add configurable error-recovery strategy
            }

            // Pause so user can observe the change
            if (seg.pauseAfterMs > 0) {
                await _sleep(seg.pauseAfterMs);
            }
        }

        if (completed === segments.length) {
            _setStatus('success');
            _log('ok', 'All segments complete.');
        }
    }

    // --- Private helpers ---

    function _sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /** Wrap a code string in an async function. */
    function _makeAsyncFn(code) {
        // PLACEHOLDER: add input sanitization / CSP-safe alternative
        return new (Object.getPrototypeOf(async function () {}).constructor)(code);
    }

    function _setStatus(state) {
        _statusDot.className = '';
        if (state) _statusDot.classList.add(state);
    }

    function _updateProgress(done, total) {
        const pct = total > 0 ? Math.round((done / total) * 100) : 0;
        _barFill.style.width  = pct + '%';
        _barLabel.textContent = `${done} / ${total}`;
    }

    function _log(type, msg) {
        const time = new Date().toLocaleTimeString('en-US', { hour12: false });
        const line = document.createElement('div');
        line.className = `log-line log-${type}`;
        line.innerHTML = `<span class="log-time">${time}</span><span class="log-msg">${msg}</span>`;
        _logEl.appendChild(line);
        _logEl.scrollTop = _logEl.scrollHeight;
    }

    function _openPanel() {
        _panel.classList.add('open');
        document.getElementById('execution-toggle-icon').textContent = '▼';
    }

    return { run };
})();
