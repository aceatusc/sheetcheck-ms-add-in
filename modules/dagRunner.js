/**
 * dagRunner.js
 *
 * Single source of truth for:
 *   - Active chain registry (chainId → chain metadata)
 *   - Current node position per chain
 *   - Forward / backward / arbitrary navigation (with undo/redo execution)
 *   - Edit branching
 *   - Graph render descriptor (git-style layout data for StepNavigator)
 *
 * Navigation model
 * ────────────────
 * A "chain" is a linear path through the DAG: rootNode → n1 → n2 → … → nK
 * The user sits ON a node. The outgoing edge from that node holds the `code`
 * segment that runs when they advance.
 *
 *   stepForward()    — run outgoing edge `code`,    move to its `to` node
 *   stepBack()       — restore snapshot of source node, retreat currentNodeId
 *   navigateTo(id)   — BFS shortest path; backward hops restore snapshots,
 *                      forward hops re-run code and refresh snapshots
 *
 * Graph render descriptor
 * ───────────────────────
 * buildRenderGraph(chainId) returns:
 *   {
 *     nodes: [{ id, label, state, col, row }],  // state: 'current'|'visited'|'unvisited'
 *     edges: [{ id, fromId, toId, label,
 *               isBranch, branchRow,            // branch rows count up from 1
 *               fromCol, fromRow, toCol, toRow }],
 *     currentNodeId,
 *     cols, rows                                 // grid dimensions
 *   }
 *
 * Branches are laid out diagonally (git style): the fork node shares its
 * column with the main path; each branch occupies its own row below row 0.
 */
const DagRunner = (() => {

    // chainId → { chainId, taskLabel, segments, rootNodeId, nodeIds[], edgeIds[], currentNodeId }
    let _chains = {};

    function _uid() {
        return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    }

    // ── Chain management ──────────────────────────────────────────────────────

    /**
     * Register a new chain from segments returned by /code.
     * Returns the chain handle (pass chainId to start()).
     */
    function prepareChain(taskLabel, segments) {
        const dag = DagStore.addChain(taskLabel, segments);
        const chainId = _uid();
        const chain = {
            chainId,
            taskLabel,
            segments,
            rootNodeId:  dag.rootNodeId,
            nodeIds:     dag.nodeIds,
            edgeIds:     dag.edgeIds,
            currentNodeId: dag.rootNodeId,   // user starts at root (before any step)
        };
        _chains[chainId] = chain;
        return chain;
    }

    function getChain(chainId)  { return _chains[chainId] || null; }

    // ── Execution start ───────────────────────────────────────────────────────

    /**
     * Begin executing a prepared chain.
     * Hands control to ExecutionEngine which calls stepForward() per step.
     */
    async function start(chainId) {
        const chain = getChain(chainId);
        if (!chain) throw new Error('Unknown chain.');
        StepNavigator.load(chain);
        await ExecutionEngine.run(chainId);
    }

    // ── Navigation ────────────────────────────────────────────────────────────

    /**
     * Advance one step: run the outgoing edge's `code`, mark it executed,
     * move currentNodeId forward.
     * Returns { edge, segment } or null if already at a leaf.
     */
    async function stepForward(chainId) {
        const chain = getChain(chainId);
        if (!chain) throw new Error('Unknown chain.');

        const outgoing = DagStore.edgesFrom(chain.currentNodeId);
        if (!outgoing.length) return null;   // already at end of this branch

        // Follow the main-chain edge if multiple outgoing (edit branches)
        const edge = _mainEdge(chain, outgoing);
        if (!edge) return null;

        // Snapshot the sheet BEFORE running code (full used range).
        const snapshot = await WorksheetSnapshot.capture();
        DagStore.setNodeSnapshot(chain.currentNodeId, snapshot);

        await _runCode(edge.segment.code);
        DagStore.markEdge(edge.id, { executed: true, failed: false });
        chain.currentNodeId = edge.to;
        return { edge, segment: edge.segment };
    }

    /**
     * Restore the snapshot stored on the source node (edge.from),
     * un-mark the edge, and retreat currentNodeId.
     * Returns { edge, segment } or null if already at root.
     */
    async function stepBack(chainId) {
        const chain = getChain(chainId);
        if (!chain) throw new Error('Unknown chain.');

        const incoming = DagStore.edgesTo(chain.currentNodeId);
        if (!incoming.length) return null;   // at root

        // Prefer the most recently executed incoming edge
        const edge = incoming.find(e => e.executed) || incoming[0];
        const desc = edge.segment?.description || edge.id;

        _log('info', `↩ Restoring snapshot: ${desc}`);
        const snapshot = DagStore.getNodeSnapshot(edge.from);
        if (!snapshot) {
            _log('err', `✗ No snapshot for node before "${desc}". Was it ever visited?`);
            throw new Error(`No snapshot available for "${desc}".`);
        }
        try {
            await WorksheetSnapshot.restore(snapshot);
        } catch (err) {
            _log('err', `✗ Undo restore failed (${desc}): ${err.message}`);
            throw err;
        }
        DagStore.markEdge(edge.id, { executed: false, failed: false });
        chain.currentNodeId = edge.from;
        _log('ok', `✓ Undone: ${desc}`);
        return { edge, segment: edge.segment };
    }

    /**
     * Navigate instantly to an arbitrary node, running undo/redo along the
     * BFS shortest path. Silently executes all intermediate edges.
     * Returns the final node id.
     */
    async function navigateTo(chainId, targetNodeId) {
        const chain = getChain(chainId);
        if (!chain) throw new Error('Unknown chain.');
        if (chain.currentNodeId === targetNodeId) return targetNodeId;

        const path = _bfsPath(chain.currentNodeId, targetNodeId);
        if (!path) throw new Error('No path to target node.');

        for (const hop of path) {
            if (hop.direction === 'forward') {
                await stepForward(chainId);
            } else {
                await stepBack(chainId);
            }
        }

        return chain.currentNodeId;
    }

    // ── Edit branching ────────────────────────────────────────────────────────

    /**
     * Branch from fromIndex: add a new chain off the fork node.
     *
     * The ORIGINAL chain's nodeIds/edgeIds are preserved unchanged as the
     * permanent "main" path shown in row 0 of the graph. The edit creates a
     * new branch that forks from the same node, visible as a diagonal row below.
     *
     * The chain's "active" execution path (segments/nodeIds/edgeIds) is switched
     * to follow the branch so ExecutionEngine continues on the new path.
     * originalNodeIds / originalEdgeIds always hold the first-generation path.
     *
     * Returns the updated chain.
     */
    function applyEdit(chainId, fromIndex, taskLabel, newSegments) {
        const chain = getChain(chainId);
        if (!chain) throw new Error('Unknown chain.');

        const forkNodeId = chain.nodeIds[fromIndex];
        if (!forkNodeId) throw new Error('Invalid edit index.');

        // Preserve original path on first edit; keep it stable on subsequent edits
        const originalNodeIds = chain.originalNodeIds || chain.nodeIds;
        const originalEdgeIds = chain.originalEdgeIds || chain.edgeIds;

        const branched = DagStore.addChain(taskLabel, newSegments, forkNodeId);

        // Save the previous active path as a historical branch so
        // buildRenderGraph can keep it on its own fixed row.
        const prevHistory = chain.branchHistory || [];
        const branchHistory = [
            ...prevHistory,
            { nodeIds: chain.nodeIds, edgeIds: chain.edgeIds },
        ];

        const updated = {
            ...chain,
            taskLabel,
            segments:    [...chain.segments.slice(0, fromIndex), ...newSegments],
            nodeIds:     [...chain.nodeIds.slice(0, fromIndex + 1), ...branched.nodeIds.slice(1)],
            edgeIds:     [...chain.edgeIds.slice(0, fromIndex),     ...branched.edgeIds],
            originalNodeIds,
            originalEdgeIds,
            branchHistory,   // all previous active paths in creation order
            currentNodeId: forkNodeId,
        };

        _chains[chainId] = updated;
        return updated;
    }

    // ── Graph render descriptor ───────────────────────────────────────────────

    /**
     * Build a layout-ready graph descriptor for StepNavigator to render.
     *
     * Layout rules (git-style):
     *   - Main chain occupies row 0, columns 0..N
     *   - Each edit branch forks from its branch-point column and descends
     *     one row per branch (row 1, 2, …), going diagonally right
     *   - Nodes carry their state: 'current' | 'visited' | 'unvisited'
     */
    function buildRenderGraph(chainId) {
        const chain = getChain(chainId);
        if (!chain) return { nodes: [], edges: [], currentNodeId: null, cols: 0, rows: 1 };

        const dag      = DagStore.getAll();
        const allEdges = dag.edges;

        // Fast lookup: nodeId → outgoing edges
        const edgesFrom = {};
        allEdges.forEach(e => {
            (edgesFrom[e.from] = edgesFrom[e.from] || []).push(e);
        });

        // ── Layout model ───────────────────────────────────────────────────
        //
        // We track a stack of "active paths" — each path is a linear sequence
        // of edges that runs horizontally on its own row. The original chain is
        // path 0 (row 0). Each edit creates a new path that forks from some node
        // on an existing path.
        //
        // Key invariant: once a path is assigned a row, ALL its nodes stay on
        // that row at their assigned columns. A fork never re-routes a path —
        // it always creates a NEW path on a new row below.
        //
        // We build paths by collecting the sequence of edge-chains from the DAG:
        //   - Start with the original chain edges as path 0
        //   - For every edit, chain.edgeIds contains the NEW active edges
        //     (prefix of original + new branch edges). Collect all edge IDs
        //     that appear in any version of chain.edgeIds across edits.
        //   - Any edge not on any "active path" is an abandoned old branch.

        // Collect all distinct linear paths in order of creation.
        // Path 0 = original chain. Path 1 = first edit's new edges. Etc.
        // We detect paths by finding all unique sequences of edges that share
        // a common prefix with the original chain.

        const originalNodeIds = chain.originalNodeIds || chain.nodeIds;
        const originalEdgeIds = chain.originalEdgeIds || chain.edgeIds;
        const originalEdgeSet = new Set(originalEdgeIds);

        // The "current active" edges are chain.edgeIds (may differ after edits)
        const activeEdgeSet = new Set(chain.edgeIds);

        // layout: nodeId → { col, row }
        const layout = {};

        // ── 1. Assign original chain to row 0 ────────────────────────────
        originalNodeIds.forEach((nid, i) => { layout[nid] = { col: i, row: 0 }; });

        // ── 2. Assign each historical branch path to its own fixed row ────
        //
        // branchHistory is an array of { nodeIds, edgeIds } saved in applyEdit,
        // one entry per edit in creation order. Each represents the full active
        // path at the time of that edit. We assign them rows 1, 2, 3… in order.
        //
        // For each historical path, only nodes NOT already placed (i.e. not on
        // the original chain or a previously-placed path) get a new layout entry.
        // Nodes shared with the original chain (the prefix up to the fork point)
        // already have layout from step 1 and are skipped.
        //
        // This guarantees: every historical path stays permanently on its
        // assigned row. The current active path gets the next available row.

        const branchHistory = chain.branchHistory || [];
        let nextRow = 1;

        // Assign historical paths — each forks exactly one row below its parent
        branchHistory.forEach(({ nodeIds: pathNodeIds }) => {
            let col       = -1;
            let parentRow = 0;
            pathNodeIds.forEach(nid => {
                if (layout[nid]) {
                    col       = layout[nid].col;
                    parentRow = layout[nid].row;
                    return;
                }
                col++;
                const row = parentRow + 1;
                layout[nid] = { col, row };
                parentRow   = row;
            });
        });

        // ── 3. Assign current active path nodes not yet placed ────────────
        // Walk chain.nodeIds. For each node not yet in layout, find the last
        // placed node's row and place this one exactly one row below it.
        // This ensures every branch connector spans exactly ROW_H regardless
        // of how many sibling branches exist at the same fork point.
        {
            let col       = -1;
            let parentRow = 0;   // row of the last placed (shared prefix) node
            chain.nodeIds.forEach(nid => {
                if (layout[nid]) {
                    col       = layout[nid].col;
                    parentRow = layout[nid].row;
                    return;
                }
                // First unplaced node: fork one row below the last placed node
                col++;
                const row = parentRow + 1;
                layout[nid] = { col, row };
                parentRow   = row;  // subsequent unplaced nodes continue on same row
            });
        }

        // ── 3. Compute grid dimensions ─────────────────────────────────────
        let maxCol = originalNodeIds.length - 1;
        let maxRow = 0;
        Object.values(layout).forEach(({ col, row }) => {
            if (col > maxCol) maxCol = col;
            if (row > maxRow) maxRow = row;
        });
        const totalCols = maxCol + 1;
        const totalRows = maxRow + 1;

        // ── 4. Build node descriptors ──────────────────────────────────────
        const visitedNodeIds = _visitedNodes(chain);
        const currentNodeId  = chain.currentNodeId;
        const nodesOut = [];
        const seenNodes = new Set();

        Object.entries(layout).forEach(([nid, pos]) => {
            if (seenNodes.has(nid)) return;
            seenNodes.add(nid);
            const n = DagStore.getNode(nid);
            nodesOut.push({
                id:    nid,
                label: n?.label || '',
                state: nid === currentNodeId ? 'current'
                     : visitedNodeIds.has(nid) ? 'visited'
                     : 'unvisited',
                col:   pos.col,
                row:   pos.row,
            });
        });

        // ── 5. Build edge descriptors ──────────────────────────────────────
        const edgesOut = [];
        const seenEdges = new Set();

        allEdges.forEach(e => {
            if (seenEdges.has(e.id)) return;
            const fp = layout[e.from];
            const tp = layout[e.to];
            if (!fp || !tp) return;
            seenEdges.add(e.id);
            edgesOut.push({
                id:       e.id,
                fromId:   e.from,
                toId:     e.to,
                label:    e.segment?.description || '',
                isBranch: !originalEdgeSet.has(e.id),
                fromCol:  fp.col, fromRow: fp.row,
                toCol:    tp.col, toRow:   tp.row,
                executed: e.executed,
                failed:   e.failed,
            });
        });

        return {
            nodes:         nodesOut,
            edges:         edgesOut,
            currentNodeId,
            cols:          totalCols,
            rows:          totalRows,
        };
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    /** BFS over the full DAG; returns [{edge, direction}] or null. */
    function _bfsPath(fromNodeId, toNodeId) {
        if (fromNodeId === toNodeId) return [];
        const dag = DagStore.getAll();

        const fwd = {}; // nodeId → edges[]
        const bwd = {};
        dag.edges.forEach(e => {
            (fwd[e.from] = fwd[e.from] || []).push(e);
            (bwd[e.to]   = bwd[e.to]   || []).push(e);
        });

        const visited = new Set([fromNodeId]);
        const queue   = [{ node: fromNodeId, path: [] }];

        while (queue.length) {
            const { node, path } = queue.shift();
            for (const e of (fwd[node] || [])) {
                if (visited.has(e.to)) continue;
                const p = [...path, { edge: e, direction: 'forward' }];
                if (e.to === toNodeId) return p;
                visited.add(e.to);
                queue.push({ node: e.to, path: p });
            }
            for (const e of (bwd[node] || [])) {
                if (visited.has(e.from)) continue;
                const p = [...path, { edge: e, direction: 'backward' }];
                if (e.from === toNodeId) return p;
                visited.add(e.from);
                queue.push({ node: e.from, path: p });
            }
        }
        return null;
    }

    /**
     * Among a list of outgoing edges, prefer the one on the current main chain.
     * Falls back to first edge if none matches (shouldn't happen on a valid chain).
     */
    function _mainEdge(chain, candidates) {
        const mainSet = new Set(chain.edgeIds);
        return candidates.find(e => mainSet.has(e.id)) || candidates[0] || null;
    }

    /**
     * Set of nodeIds whose state is "committed" — the user has seen and agreed
     * to the sheet at this node.
     * Root is always committed (it's the baseline before any agent change).
     * Every node reached by an executed edge is also committed.
     */
    function _visitedNodes(chain) {
        const visited = new Set();
        const dag = DagStore.getAll();
        // Root is always committed — it's the sheet state before anything ran
        visited.add(chain.rootNodeId);
        dag.edges.forEach(e => { if (e.executed) visited.add(e.to); });
        return visited;
    }

    /** Write a line to the visible Execution Log. Safe to call before ExecutionEngine exists. */
    function _log(type, msg) {
        try { ExecutionEngine.log(type, msg); } catch (_) { console.log(`[DagRunner][${type}] ${msg}`); }
    }

    /**
     * Compile and run a code string as an async function.
     * Throws on error (caller is responsible for catching and logging).
     */
    async function _runCode(code) {
        if (!code) return;
        const fn = new (Object.getPrototypeOf(async function(){}).constructor)(code);
        await fn();
    }

    return {
        prepareChain,
        start,
        getChain,
        stepForward,
        stepBack,
        navigateTo,
        applyEdit,
        buildRenderGraph,
    };
})();
