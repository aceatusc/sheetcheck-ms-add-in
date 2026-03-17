/**
 * dagRunner.js
 * Owns the "active chain" of segments stored in DagStore and coordinates execution.
 */
const DagRunner = (() => {

    let _chainsById = {};
    let _activeChainId = null;

    function _uid() {
        return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    }

    /**
     * Stores a new chain in the DAG and returns a chain handle.
     * @param {string} taskLabel
     * @param {object[]} segments
     */
    function prepareChain(taskLabel, segments) {
        const chain = DagStore.addChain(taskLabel, segments);
        const chainId = _uid();
        _chainsById[chainId] = {
            chainId,
            taskLabel,
            segments,
            ...chain, // rootNodeId, nodeIds, edgeIds
        };
        return _chainsById[chainId];
    }

    function getChain(chainId) {
        return _chainsById[chainId] || null;
    }

    function getActiveChain() {
        return _activeChainId ? getChain(_activeChainId) : null;
    }

    /**
     * Start executing a prepared chain, pausing between steps.
     * @param {string} chainId
     */
    async function start(chainId) {
        const chain = getChain(chainId);
        if (!chain) throw new Error('Unknown chain.');
        _activeChainId = chainId;

        StepNavigator.loadSegments(chain.segments, {
            chainId,
            rootNodeId: chain.rootNodeId,
            nodeIds: chain.nodeIds,
            edgeIds: chain.edgeIds,
            taskLabel: chain.taskLabel,
        });

        await ExecutionEngine.run(chain.segments, {
            onStepDone: (i) => {
                const edgeId = chain.edgeIds?.[i];
                if (edgeId) DagStore.markEdgeExecuted(edgeId, false);
            },
            onStepFailed: (i) => {
                const edgeId = chain.edgeIds?.[i];
                if (edgeId) DagStore.markEdgeExecuted(edgeId, true);
            },
        });
    }

    /**
     * Branch from a given index (node) and replace the active chain tail.
     * Returns the updated chain object.
     *
     * @param {string} chainId
     * @param {number} fromIndex
     * @param {string} taskLabel
     * @param {object[]} newSegmentsFromHere
     */
    function applyEdit(chainId, fromIndex, taskLabel, newSegmentsFromHere) {
        const chain = getChain(chainId);
        if (!chain) throw new Error('Unknown chain.');
        const fromNodeId = chain.nodeIds?.[fromIndex];
        if (!fromNodeId) throw new Error('Invalid edit point.');

        const branched = DagStore.branchFrom(fromNodeId, taskLabel, newSegmentsFromHere);
        const prefix = chain.segments.slice(0, fromIndex);

        const updated = {
            ...chain,
            taskLabel,
            segments: [...prefix, ...newSegmentsFromHere],
            rootNodeId: chain.rootNodeId,
            nodeIds: [...chain.nodeIds.slice(0, fromIndex + 1), ...branched.nodeIds.slice(1)],
            edgeIds: [...chain.edgeIds.slice(0, fromIndex), ...branched.edgeIds],
        };

        _chainsById[chainId] = updated;
        return updated;
    }

    return { prepareChain, start, applyEdit, getChain, getActiveChain };
})();

