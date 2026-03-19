/**
 * executionEngine.js
 * Runs a chain sequentially by calling DagRunner.stepForward() for each step,
 * pausing between each for the user to review the change and click →.
 */
const ExecutionEngine = (() => {

    const _logEl         = document.getElementById('execution-log');
    const _statusDot     = document.getElementById('execution-status-dot');

    /**
     * Run all steps in a chain to completion.
     * @param {string} chainId
     */
    async function run(chainId) {
        const chain = DagRunner.getChain(chainId);
        if (!chain?.segments?.length) return;

        _setStatus('running');
        _log('info', `Starting execution: ${chain.segments.length} segment(s)`);

        while (true) {
            if (StepNavigator.isDismissed()) break;  // user closed the navigator

            const current = DagRunner.getChain(chainId);
            if (!current) break;

            const outgoing = (current.store || DagStore).edgesFrom(current.currentNodeId);
            if (!outgoing.length) break;

            if (StepNavigator.isDismissed()) break;

            let stepOk = false;
            try {
                const result = await DagRunner.stepForward(chainId);
                if (!result) break;

                stepOk = true;
                _log('ok', `✓ ${result.segment.description}`);
            } catch (err) {
                // Mark the edge red and log — but do NOT pause execution.
                // The red edge in the graph is enough; the loop continues to the next step.
                _log('err', `✗ ${err.message}`);
                _setStatus('error');
                const failedEdges = (current.store || DagStore).edgesFrom(current.currentNodeId);
                if (failedEdges[0]) (current.store || DagStore).markEdge(failedEdges[0].id, { executed: true, failed: true });
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

        // ── Review loop ────────────────────────────────────────────────────
        // Identical to the execution loop above — waitForNext keeps
        // _advanceResolve active so → / ← / graph-click work normally.
        // stepForward re-runs the code on each → click, same as first time.
        while (!StepNavigator.isDismissed()) {
            const rc = DagRunner.getChain(chainId);
            if (!rc) break;
            const rcOut = (rc.store || DagStore).edgesFrom(rc.currentNodeId);
            if (!rcOut.length) {
                await StepNavigator.waitForNext();
                break;
            }
            await StepNavigator.waitForNext();
            if (StepNavigator.isDismissed()) break;
            try {
                const r = await DagRunner.stepForward(chainId);
                if (!r) break;
                _log('ok', `↺ ${r.segment.description}`);
                StepNavigator.refreshGraph();
            } catch (err) {
                _log('err', `✗ ${err.message}`);
                DagRunner.advancePastFailed(chainId);
                StepNavigator.refreshGraph();
            }
        }
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    function _setStatus(state) {
        _statusDot.className = '';
        if (state) _statusDot.classList.add(state);
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
