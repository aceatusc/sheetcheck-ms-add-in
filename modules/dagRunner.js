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
    let _activeChainId = null;

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
    function getActiveChain()   { return _activeChainId ? getChain(_activeChainId) : null; }

    // ── Execution start ───────────────────────────────────────────────────────

    /**
     * Begin executing a prepared chain.
     * Hands control to ExecutionEngine which calls stepForward() per step.
     */
    async function start(chainId) {
        const chain = getChain(chainId);
        if (!chain) throw new Error('Unknown chain.');
        _activeChainId = chainId;

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

        // Snapshot the sheet BEFORE running code. Scoped to the segment's
        // sheet_context ranges for speed; falls back to full used range.
        const snapshot = await WorksheetSnapshot.capture(edge.segment.sheet_context);
        DagStore.setNodeSnapshot(chain.currentNodeId, snapshot);

        await _runCode(edge.segment.code);
        DagStore.markEdge(edge.id, { executed: true, failed: false });
        chain.currentNodeId = edge.to;
        // Capture the post-step state onto the destination node so it
        // can be used as the "before" snapshot for the *next* forward step.
        const snapAfter = await WorksheetSnapshot.capture(edge.segment.sheet_context);
        DagStore.setNodeSnapshot(chain.currentNodeId, snapAfter);
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
            const desc = hop.edge.segment?.description || hop.edge.id;
            if (hop.direction === 'forward') {
                _log('info', `▶ Navigate forward: ${desc}`);
                const snapFwd = await WorksheetSnapshot.capture(hop.edge.segment.sheet_context);
                DagStore.setNodeSnapshot(chain.currentNodeId, snapFwd);
                try {
                    await _runCode(hop.edge.segment.code);
                } catch (err) {
                    _log('err', `✗ Navigate forward failed (${desc}): ${err.message}`);
                    throw err;
                }
                DagStore.markEdge(hop.edge.id, { executed: true, failed: false });
                chain.currentNodeId = hop.edge.to;
                const snapFwdAfter = await WorksheetSnapshot.capture(hop.edge.segment.sheet_context);
                DagStore.setNodeSnapshot(chain.currentNodeId, snapFwdAfter);
                _log('ok', `✓ ${desc}`);
            } else {
                _log('info', `↩ Restoring snapshot: ${desc}`);
                const snapBack = DagStore.getNodeSnapshot(hop.edge.from);
                if (!snapBack) {
                    _log('err', `✗ No snapshot for node before "${desc}".`);
                    throw new Error(`No snapshot for "${desc}".`);
                }
                try {
                    await WorksheetSnapshot.restore(snapBack);
                } catch (err) {
                    _log('err', `✗ Snapshot restore failed (${desc}): ${err.message}`);
                    throw err;
                }
                DagStore.markEdge(hop.edge.id, { executed: false, failed: false });
                chain.currentNodeId = hop.edge.from;
                _log('ok', `↩ Restored: ${desc}`);
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

        const updated = {
            ...chain,
            taskLabel,
            // Active execution path = prefix up to fork + new branch
            segments:    [...chain.segments.slice(0, fromIndex), ...newSegments],
            nodeIds:     [...chain.nodeIds.slice(0, fromIndex + 1), ...branched.nodeIds.slice(1)],
            edgeIds:     [...chain.edgeIds.slice(0, fromIndex),     ...branched.edgeIds],
            // Original path stays fixed for graph rendering
            originalNodeIds,
            originalEdgeIds,
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

        // The "main" path is always the original chain — it never moves.
        // After an edit, chain.originalNodeIds holds the first-generation path.
        const mainNodeIds = chain.originalNodeIds || chain.nodeIds;
        const mainEdgeIds = chain.originalEdgeIds || chain.edgeIds;
        const mainEdgeSet = new Set(mainEdgeIds);

        // ── 1. Assign columns to main-chain nodes (row 0) ─────────────────
        const layout = {};
        mainNodeIds.forEach((nid, i) => { layout[nid] = { col: i, row: 0 }; });

        // ── 2. Walk ALL non-main edges as branches ────────────────────────
        // Each distinct outgoing non-main edge from a main node starts a new row.
        let branchRow = 0;
        const branchRowMap = {};   // edgeId → branchRow

        mainNodeIds.forEach((forkNodeId, forkCol) => {
            const outgoing = allEdges.filter(e => e.from === forkNodeId && !mainEdgeSet.has(e.id));
            outgoing.forEach(startEdge => {
                branchRow++;
                let col = forkCol;  // diagonal: branch starts at fork column
                let cur = startEdge;
                const visited = new Set();
                while (cur && !visited.has(cur.id)) {
                    visited.add(cur.id);
                    branchRowMap[cur.id] = branchRow;
                    col++;
                    if (!layout[cur.to]) layout[cur.to] = { col, row: branchRow };
                    const nextOuts = allEdges.filter(e => e.from === cur.to && !mainEdgeSet.has(e.id));
                    cur = nextOuts.length === 1 ? nextOuts[0] : null;
                }
            });
        });

        const totalRows = branchRow + 1;
        const totalCols = mainNodeIds.length;

        // ── 3. Build node descriptors ──────────────────────────────────────
        const visitedNodeIds = _visitedNodes(chain);
        const currentNodeId  = chain.currentNodeId;

        const nodesOut = [];
        const seenNodes = new Set();

        // Main chain nodes
        mainNodeIds.forEach(nid => {
            if (seenNodes.has(nid)) return;
            seenNodes.add(nid);
            const n = DagStore.getNode(nid);
            const pos = layout[nid] || { col: 0, row: 0 };
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

        // Branch nodes (destination nodes of branch edges)
        Object.keys(branchRowMap).forEach(eid => {
            const edge = DagStore.getEdge(eid);
            if (!edge) return;
            [edge.from, edge.to].forEach(nid => {
                if (seenNodes.has(nid)) return;
                seenNodes.add(nid);
                const n   = DagStore.getNode(nid);
                const pos = layout[nid] || { col: 0, row: 0 };
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
        });

        // ── 4. Build edge descriptors ──────────────────────────────────────
        const edgesOut = [];

        // Main chain edges
        mainEdgeIds.forEach(eid => {
            const e = DagStore.getEdge(eid);
            if (!e) return;
            const fp = layout[e.from] || { col: 0, row: 0 };
            const tp = layout[e.to]   || { col: 1, row: 0 };
            edgesOut.push({
                id:        eid,
                fromId:    e.from,
                toId:      e.to,
                label:     e.segment?.description || '',
                isBranch:  false,
                branchRow: 0,
                fromCol:   fp.col, fromRow: fp.row,
                toCol:     tp.col, toRow:   tp.row,
                executed:  e.executed,
                failed:    e.failed,
            });
        });

        // Branch edges
        Object.entries(branchRowMap).forEach(([eid, bRow]) => {
            const e = DagStore.getEdge(eid);
            if (!e) return;
            const fp = layout[e.from] || { col: 0, row: 0 };
            const tp = layout[e.to]   || { col: 1, row: bRow };
            edgesOut.push({
                id:        eid,
                fromId:    e.from,
                toId:      e.to,
                label:     e.segment?.description || '',
                isBranch:  true,
                branchRow: bRow,
                fromCol:   fp.col, fromRow: fp.row,
                toCol:     tp.col, toRow:   tp.row,
                executed:  e.executed,
                failed:    e.failed,
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

    /** Set of all nodeIds that have at least one executed incoming edge. */
    function _visitedNodes(chain) {
        const visited = new Set();
        const dag = DagStore.getAll();
        dag.edges.forEach(e => { if (e.executed) visited.add(e.to); });
        // Root is always "visited" once any step has been taken
        if (visited.size > 0) visited.add(chain.rootNodeId);
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
        getActiveChain,
        stepForward,
        stepBack,
        navigateTo,
        applyEdit,
        buildRenderGraph,
    };
})();
