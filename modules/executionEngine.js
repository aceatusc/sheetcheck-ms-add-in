/**
 * executionEngine.js
 * Runs code segments one at a time, pausing between each for the user
 * to review the change and click → in the StepNavigator overlay.
 *
 * Accepts optional callbacks for DAG tracking:
 *   { onStepDone(index), onStepFailed(index) }
 */
const ExecutionEngine = (() => {

    const STORAGE_KEY = 'sheetcheck_dag';
    const _logEl         = document.getElementById('execution-log');
    const _statusDot     = document.getElementById('execution-status-dot');
    const _barFill       = document.getElementById('segment-bar-fill');
    const _barLabel      = document.getElementById('segment-label');
    const _progressStrip = document.getElementById('segment-progress');
    const _panel         = document.getElementById('execution-panel');

    /**
     * Run segments sequentially, waiting for user confirmation after each one.
     * @param {CodeSegment[]} segments
     * @param {{ onStepDone?: Function, onStepFailed?: Function }} [callbacks]
     */
    async function run(segments, callbacks = {}) {
        if (!segments || segments.length === 0) return;

        _setStatus('running');
        _progressStrip.classList.add('visible');
        _log('info', `Starting execution: ${segments.length} segment(s)`);

        for (let i = 0; i < segments.length; i++) {
            const seg = segments[i];

            _updateProgress(i, segments.length);
            _log('info', `▶ ${seg.description}`);
            _log('info', localStorage.getItem(STORAGE_KEY))

            StepNavigator.markRunning(i);

            let stepOk = false;
            try {
                const fn = _makeAsyncFn(seg.code);
                await fn();
                stepOk = true;
                _updateProgress(i + 1, segments.length);
                _log('ok', `✓ ${seg.description}`);
                callbacks.onStepDone?.(i);
            } catch (err) {
                _log('err', `✗ ${seg.description}: ${err.message}`);
                _setStatus('error');
                console.error('[ExecutionEngine] Segment error:', err);
                callbacks.onStepFailed?.(i);
                await StepNavigator.markFailed(i, err.message);
                continue;
            }

            if (stepOk) {
                await StepNavigator.waitForNext(i);
            }
        }

        if (_statusDot.classList.contains('error')) {
            _log('info', 'Execution complete with errors — review failed steps above.');
        } else {
            _setStatus('success');
            _log('ok', 'All segments complete.');
        }
        await StepNavigator.showVerifyResults();
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    function _makeAsyncFn(code) {
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

    return { run };
})();
