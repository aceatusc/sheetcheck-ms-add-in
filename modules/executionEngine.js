/**
 * executionEngine.js
 * Runs a chain sequentially by calling DagRunner.stepForward() for each step,
 * pausing between each for the user to review the change and click →.
 */
const ExecutionEngine = (() => {

    const _logEl         = document.getElementById('execution-log');
    const _statusDot     = document.getElementById('execution-status-dot');
    const _barFill       = document.getElementById('segment-bar-fill');
    const _barLabel      = document.getElementById('segment-label');
    const _progressStrip = document.getElementById('segment-progress');

    /**
     * Run all steps in a chain to completion.
     * @param {string} chainId
     */
    async function run(chainId) {
        const chain = DagRunner.getChain(chainId);
        if (!chain?.segments?.length) return;

        _setStatus('running');
        _progressStrip.classList.add('visible');
        _log('info', `Starting execution: ${chain.segments.length} segment(s)`);

        const total = chain.segments.length;

        while (true) {
            // Fetch current chain state (may have been mutated by an edit)
            const current = DagRunner.getChain(chainId);
            if (!current) break;

            // Check whether there's a next step to run
            const outgoing = DagStore.edgesFrom(current.currentNodeId);
            if (!outgoing.length) break;  // reached a leaf — all done

            const stepIdx = current.nodeIds.indexOf(current.currentNodeId);
            _updateProgress(stepIdx, current.segments.length);

            StepNavigator.markRunning();

            let stepOk = false;
            try {
                const result = await DagRunner.stepForward(chainId);
                if (!result) break;  // no outgoing edge (leaf)

                stepOk = true;
                _updateProgress(stepIdx + 1, DagRunner.getChain(chainId).segments.length);
                _log('ok', `✓ ${result.segment.description}`);
            } catch (err) {
                _log('err', `✗ ${err.message}`);
                _setStatus('error');
                console.error('[ExecutionEngine]', err);
                await StepNavigator.markFailed(err.message);
                // Advance past the failed step and continue
                continue;
            }

            if (stepOk) {
                await StepNavigator.waitForNext();
            }
        }

        if (_statusDot.classList.contains('error')) {
            _log('info', 'Execution complete with errors — review failed steps above.');
        } else {
            _setStatus('success');
            _log('ok', 'All segments complete.');
        }

        await RubricManager.showVerifyResults();
    }

    // ── Private helpers ───────────────────────────────────────────────────────

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

    return { run, log: _log };
})();
