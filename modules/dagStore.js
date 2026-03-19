/**
 * dagStore.js — Per-run isolated DAG store.
 *
 * Instead of a single global DAG, each run gets its own DagStore instance
 * created via DagStore.create(). This prevents nodes/edges from different
 * runs (or different worksheets) leaking into each other.
 *
 * Model:
 *   Node { id, label, taskLabel, isRoot, snapshot? }
 *   Edge { id, from, to, segment, executed, failed }
 */
const DagStore = (() => {

    function _uid() {
        return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    }

    /**
     * Create a fresh, isolated DAG store for one run.
     * Returns a store object with the same API as the old singleton.
     */
    function create() {
        const nodes = [];
        const edges = [];

        function addChain(taskLabel, segments, fromNodeId = null) {
            let rootId = fromNodeId;
            if (!rootId) {
                rootId = _uid();
                nodes.push({ id: rootId, label: 'Start', taskLabel, isRoot: true });
            }
            const nodeIds = [rootId];
            const edgeIds = [];
            let prev = rootId;
            segments.forEach(seg => {
                const nid = _uid();
                const eid = _uid();
                nodes.push({ id: nid, label: seg.description, taskLabel, isRoot: false });
                edges.push({ id: eid, from: prev, to: nid, segment: seg, executed: false, failed: false });
                nodeIds.push(nid);
                edgeIds.push(eid);
                prev = nid;
            });
            return { rootNodeId: rootId, nodeIds, edgeIds };
        }

        function markEdge(edgeId, { executed = true, failed = false } = {}) {
            const e = edges.find(e => e.id === edgeId);
            if (e) { e.executed = executed; e.failed = failed; }
        }

        function setNodeSnapshot(nodeId, snapshot) {
            const n = nodes.find(n => n.id === nodeId);
            if (n) n.snapshot = snapshot;
        }

        function getNodeSnapshot(nodeId) {
            const n = nodes.find(n => n.id === nodeId);
            return n?.snapshot || null;
        }

        function getEdge(id)      { return edges.find(e => e.id === id) || null; }
        function getNode(id)      { return nodes.find(n => n.id === id) || null; }
        function getAll()         { return { nodes, edges }; }
        function edgesFrom(nid)   { return edges.filter(e => e.from === nid); }
        function edgesTo(nid)     { return edges.filter(e => e.to   === nid); }

        return { addChain, markEdge, setNodeSnapshot, getNodeSnapshot, getEdge, getNode, getAll, edgesFrom, edgesTo };
    }

    // Backwards-compat singleton init for app.js (no-op now)
    function init() {}

    return { init, create };
})();
