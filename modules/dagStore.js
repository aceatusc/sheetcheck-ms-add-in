/**
 * dagStore.js — Minimal in-memory DAG.
 *
 * Model:
 *   Node { id, label, taskLabel, isRoot }
 *   Edge { id, from, to, segment, executed, failed }
 *
 * Nodes are worksheet states. Edges are code-segment transitions.
 * All path-finding and traversal logic lives in dagRunner.js.
 */
const DagStore = (() => {

    const GLOBAL_KEY = '__sheetcheckDag';
    let _dag = null;

    function _uid() {
        return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    }

    function init({ reset = false } = {}) {
        if (typeof window === 'undefined') {
            _dag = { nodes: [], edges: [] };
            return;
        }
        if (reset || !window[GLOBAL_KEY]) {
            window[GLOBAL_KEY] = { nodes: [], edges: [] };
        }
        _dag = window[GLOBAL_KEY];
    }

    function _ensure() { if (!_dag) init(); }

    // ── Write ─────────────────────────────────────────────────────────────────

    /**
     * Add a linear chain of segments to the DAG.
     * fromNodeId — branch point node (null = create a new root).
     * Returns { rootNodeId, nodeIds[], edgeIds[] }
     */
    function addChain(taskLabel, segments, fromNodeId = null) {
        _ensure();
        const now = Date.now();

        let rootId = fromNodeId;
        if (!rootId) {
            rootId = _uid();
            _dag.nodes.push({ id: rootId, label: 'Start', taskLabel, isRoot: true });
        }

        const nodeIds = [rootId];
        const edgeIds = [];
        let prev = rootId;

        segments.forEach((seg, i) => {
            const nid = _uid();
            const eid = _uid();
            _dag.nodes.push({ id: nid, label: seg.description, taskLabel, isRoot: false });
            _dag.edges.push({ id: eid, from: prev, to: nid, segment: seg, executed: false, failed: false });
            nodeIds.push(nid);
            edgeIds.push(eid);
            prev = nid;
        });

        return { rootNodeId: rootId, nodeIds, edgeIds };
    }

    /** Mark an edge executed (or failed). */
    function markEdge(edgeId, { executed = true, failed = false } = {}) {
        _ensure();
        const e = _dag.edges.find(e => e.id === edgeId);
        if (e) { e.executed = executed; e.failed = failed; }
    }

    /**
     * Attach a worksheet snapshot to a node (stored on the node object itself).
     * Called by DagRunner before each stepForward so the node records the
     * worksheet state that existed when the user was sitting on it.
     */
    function setNodeSnapshot(nodeId, snapshot) {
        _ensure();
        const n = _dag.nodes.find(n => n.id === nodeId);
        if (n) n.snapshot = snapshot;
    }

    function getNodeSnapshot(nodeId) {
        _ensure();
        const n = _dag.nodes.find(n => n.id === nodeId);
        return n?.snapshot || null;
    }

    // ── Read ──────────────────────────────────────────────────────────────────

    function getEdge(id)  { _ensure(); return _dag.edges.find(e => e.id === id) || null; }
    function getNode(id)  { _ensure(); return _dag.nodes.find(n => n.id === id) || null; }
    function getAll()     { _ensure(); return _dag; }

    /** All edges leaving a node. */
    function edgesFrom(nodeId) { _ensure(); return _dag.edges.filter(e => e.from === nodeId); }
    /** All edges entering a node. */
    function edgesTo(nodeId)   { _ensure(); return _dag.edges.filter(e => e.to   === nodeId); }

    function clear() { _ensure(); _dag.nodes.length = 0; _dag.edges.length = 0; }

    return { init, addChain, markEdge, setNodeSnapshot, getNodeSnapshot, getEdge, getNode, getAll, edgesFrom, edgesTo, clear };
})();
