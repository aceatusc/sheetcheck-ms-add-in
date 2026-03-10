/**
 * executionEngine.js
 * Runs code segments one at a time, pausing between each for the user
 * to review the change and click → in the StepNavigator overlay.
 */
const ExecutionEngine = (() => {

    const _logEl         = document.getElementById('execution-log');
    const _statusDot     = document.getElementById('execution-status-dot');
    const _barFill       = document.getElementById('segment-bar-fill');
    const _barLabel      = document.getElementById('segment-label');
    const _progressStrip = document.getElementById('segment-progress');
    const _panel         = document.getElementById('execution-panel');

    /**
     * Run segments sequentially, waiting for user confirmation after each one.
     * @param {CodeSegment[]} segments
     */
    async function run(segments) {
        if (!segments || segments.length === 0) return;

        _setStatus('running');
        _progressStrip.classList.add('visible');
        _log('info', `Starting execution: ${segments.length} segment(s)`);

        let completed = 0;

        for (let i = 0; i < segments.length; i++) {
            const seg = segments[i];

            _updateProgress(completed, segments.length);
            _log('info', `▶ ${seg.description}`);

            // Show overlay in "running" state while code executes
            StepNavigator.markRunning(i);

            try {
                const fn = _makeAsyncFn(seg.code);
                await fn();

                completed++;
                _updateProgress(completed, segments.length);
                _log('ok', `✓ ${seg.description}`);

            } catch (err) {
                _log('err', `✗ ${seg.description}: ${err.message} [${seg.code}]`);
                _setStatus('error');
                console.error('[ExecutionEngine] Segment error:', err);
                break;
            }

            // Hand control to the user — resolves only when they click →
            await StepNavigator.waitForNext(i);
        }

        if (completed === segments.length) {
            _setStatus('success');
            _log('ok', 'All segments complete.');
            // Trigger rubric verification automatically at the end
            await StepNavigator.showVerifyResults();
        }
    }

    // ── Private helpers ───────────────────────────────────────────────────────

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
