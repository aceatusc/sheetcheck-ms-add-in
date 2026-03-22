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
                _log('err', `✗ ${err.message}`);
                _setStatus('error');
                const failedEdges = (current.store || DagStore).edgesFrom(current.currentNodeId);
                if (failedEdges[0]) (current.store || DagStore).markEdge(failedEdges[0].id, { executed: true, failed: true });
                // Pause on the failed node — StepNavigator shows the error message
                // and the segment's manual_steps guidance so the user can do it by hand.
                // Execution resumes (advances past the failure) only after they click →.
                await StepNavigator.markFailed(err.message);
                if (StepNavigator.isDismissed()) break;
                DagRunner.advancePastFailed(chainId);
                StepNavigator.refreshGraph();
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

        // ── Review loop ────────────────────────────────────────────────────
        // Lets the user navigate back/forward through completed steps.
        // ✓ button calls dismiss() which sets _dismissed and resolves the
        // pending waitForNext promise — we check isDismissed() immediately
        // on the next line so we exit without doing any extra work.
        while (true) {
            await StepNavigator.waitForNext();
            if (StepNavigator.isDismissed()) break;  // ✓ was clicked — exit cleanly

            const rc = DagRunner.getChain(chainId);
            if (!rc) break;
            const rcOut = (rc.store || DagStore).edgesFrom(rc.currentNodeId);
            if (!rcOut.length) break;  // still at leaf after navigation — stop

            try {
                const r = await DagRunner.stepForward(chainId);
                if (!r) break;
                _log('ok', `↺ ${r.segment.description}`);
                StepNavigator.refreshGraph();
            } catch (err) {
                _log('err', `✗ ${err.message}`);
                await StepNavigator.markFailed(err.message);
                if (StepNavigator.isDismissed()) break;
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
