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
                // Step navigator suppressed — failed steps are silently skipped.
                DagRunner.advancePastFailed(chainId);
            }

            // Step pause removed — all segments apply immediately without user confirmation.
            // if (stepOk) {
            //     await StepNavigator.waitForNext();
            //     if (StepNavigator.isDismissed()) break;
            // }
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

        // Post-completion review loop removed — navigator is not shown in auto-apply mode.
        // if (true) { await StepNavigator.waitForNext(); ... }
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
