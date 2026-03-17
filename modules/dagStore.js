/**
 * dagStore.js — Persistent DAG of worksheet states and code segments.
 *
 * Model:
 *   Node  — a worksheet state snapshot  { id, label, ts, taskLabel }
 *   Edge  — a code segment transition   { id, from, to, segment, executed, failed }
 *
 * The graph is directed and acyclic. Each new /code run produces a linear
 * chain: root → n1 → n2 → … → nK. An edit at step i branches from node i,
 * creating a new chain parallel to the original tail.
 *
 * Stored in localStorage under key "sheetcheck_dag".
 */
const DagStore = (() => {

    const STORAGE_KEY = 'sheetcheck_dag';

    // In-memory mirror of what's in localStorage
    let _dag = _load();

    // ── Persistence ───────────────────────────────────────────────────────────

    function _load() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (raw) return JSON.parse(raw);
        } catch(e) { console.warn('[DagStore] load error', e); }
        return { nodes: [], edges: [] };
    }

    function _save() {
        try { localStorage.setItem(STORAGE_KEY, JSON.stringify(_dag)); }
        catch(e) { console.warn('[DagStore] save error', e); }
    }

    function _uid() {
        return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    }

    // ── Public API ────────────────────────────────────────────────────────────

    /**
     * Start a brand-new linear chain from a fresh task.
     * Returns { rootNodeId, nodeIds[], edgeIds[] }
     *
     * @param {string}   taskLabel   — the user's message
     * @param {object[]} segments    — array of code segments from /code
     * @param {string}   [fromNodeId] — if set, branch from an existing node
     *                                  (used when Edit generates a new chain)
     */
    function addChain(taskLabel, segments, fromNodeId = null) {
        const now = Date.now();

        // Root node — either a fresh one or reuse the branch point
        let rootId = fromNodeId;
        if (!rootId) {
            rootId = _uid();
            _dag.nodes.push({ id: rootId, label: 'Start', ts: now, taskLabel, isRoot: true });
        }

        const nodeIds = [rootId];
        const edgeIds = [];

        let prevNodeId = rootId;
        segments.forEach((seg, i) => {
            const nodeId = _uid();
            const edgeId = _uid();
            _dag.nodes.push({
                id: nodeId,
                label: seg.description,
                ts: now + i + 1,
                taskLabel,
                isRoot: false,
            });
            _dag.edges.push({
                id: edgeId,
                from: prevNodeId,
                to: nodeId,
                segment: seg,
                executed: false,
                failed: false,
            });
            nodeIds.push(nodeId);
            edgeIds.push(edgeId);
            prevNodeId = nodeId;
        });

        _save();
        return { rootNodeId: rootId, nodeIds, edgeIds };
    }

    /**
     * Mark an edge as executed (or failed).
     * @param {string}  edgeId
     * @param {boolean} failed
     */
    function markEdgeExecuted(edgeId, failed = false) {
        const e = _dag.edges.find(e => e.id === edgeId);
        if (e) { e.executed = true; e.failed = failed; _save(); }
    }

    /**
     * Replace tail edges from a given node (for Edit flow).
     * Removes all edges that are descendants of fromNodeId
     * (does not remove nodes — they stay as abandoned states),
     * then adds the new chain.
     */
    function branchFrom(fromNodeId, taskLabel, newSegments) {
        // Just add a new chain branching from the existing node
        return addChain(taskLabel, newSegments, fromNodeId);
    }

    /** Full graph snapshot — returns deep copy. */
    function getGraph() {
        return JSON.parse(JSON.stringify(_dag));
    }

    /**
     * Find a path (array of edges) between two node ids using BFS.
     * Returns null if no path exists.
     * Each step is { edge, direction } where direction = 'forward' | 'backward'.
     */
    function findPath(fromNodeId, toNodeId) {
        if (fromNodeId === toNodeId) return [];

        // Build adjacency: forward and backward edges
        const fwd = {}; // nodeId → [edge]
        const bwd = {}; // nodeId → [edge]
        _dag.edges.forEach(e => {
            if (!fwd[e.from]) fwd[e.from] = [];
            fwd[e.from].push(e);
            if (!bwd[e.to]) bwd[e.to] = [];
            bwd[e.to].push(e);
        });

        // BFS over (nodeId, path_so_far)
        const visited = new Set([fromNodeId]);
        const queue = [{ node: fromNodeId, path: [] }];

        while (queue.length) {
            const { node, path } = queue.shift();

            // Try forward edges
            for (const edge of (fwd[node] || [])) {
                if (visited.has(edge.to)) continue;
                const newPath = [...path, { edge, direction: 'forward' }];
                if (edge.to === toNodeId) return newPath;
                visited.add(edge.to);
                queue.push({ node: edge.to, path: newPath });
            }

            // Try backward edges
            for (const edge of (bwd[node] || [])) {
                if (visited.has(edge.from)) continue;
                const newPath = [...path, { edge, direction: 'backward' }];
                if (edge.from === toNodeId) return newPath;
                visited.add(edge.from);
                queue.push({ node: edge.from, path: newPath });
            }
        }
        return null; // unreachable
    }

    /** Get the edge that leads INTO a given node (most recently executed one). */
    function getIncomingEdge(nodeId) {
        // Among all edges pointing to nodeId, prefer executed ones
        const incoming = _dag.edges.filter(e => e.to === nodeId);
        return incoming.find(e => e.executed) || incoming[0] || null;
    }

    /** Get a node by id. */
    function getNode(nodeId) {
        return _dag.nodes.find(n => n.id === nodeId) || null;
    }

    /** Get all nodes + edges. */
    function getAll() { return _dag; }

    /** Clear the entire DAG (for testing / reset). */
    function clear() {
        _dag = { nodes: [], edges: [] };
        _save();
    }

    return { addChain, branchFrom, markEdgeExecuted, getGraph, findPath,
             getIncomingEdge, getNode, getAll, clear };
})();
