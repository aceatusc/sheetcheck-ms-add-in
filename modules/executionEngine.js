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
            if (StepNavigator.isDismissed()) break;  // user closed the navigator

            const current = DagRunner.getChain(chainId);
            if (!current) break;

            const outgoing = DagStore.edgesFrom(current.currentNodeId);
            if (!outgoing.length) break;

            const stepIdx = current.nodeIds.indexOf(current.currentNodeId);
            _updateProgress(stepIdx, current.segments.length);

            StepNavigator.markRunning();
            if (StepNavigator.isDismissed()) break;

            let stepOk = false;
            try {
                const result = await DagRunner.stepForward(chainId);
                if (!result) break;

                stepOk = true;
                _updateProgress(stepIdx + 1, DagRunner.getChain(chainId).segments.length);
                _log('ok', `✓ ${result.segment.description}`);
            } catch (err) {
                // Mark the edge red and log — but do NOT pause execution.
                // The red edge in the graph is enough; the loop continues to the next step.
                _log('err', `✗ ${err.message}`);
                _setStatus('error');
                const failedEdges = DagStore.edgesFrom(current.currentNodeId);
                if (failedEdges[0]) DagStore.markEdge(failedEdges[0].id, { executed: true, failed: true });
                if (StepNavigator.isDismissed()) break;
                // Advance currentNodeId past the failed step so the loop continues
                DagRunner.advancePastFailed(chainId);
                StepNavigator.refreshGraph();
                continue;
            }

            if (stepOk) {
                await StepNavigator.waitForNext();
                if (StepNavigator.isDismissed()) break;
            }
        }

        if (StepNavigator.isDismissed()) {
            _log('info', 'Execution stopped — navigator dismissed.');
            _setStatus('');
            return;  // don't show verify results
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
