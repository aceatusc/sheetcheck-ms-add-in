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

        // Build fast lookup: nodeId → outgoing edges
        const edgesFrom = {};
        allEdges.forEach(e => {
            (edgesFrom[e.from] = edgesFrom[e.from] || []).push(e);
        });

        // The original (first-generation) chain always occupies row 0.
        const mainNodeIds = chain.originalNodeIds || chain.nodeIds;
        const mainEdgeIds = chain.originalEdgeIds || chain.edgeIds;
        const mainEdgeSet = new Set(mainEdgeIds);

        // ── 1. Assign layout to ALL nodes via BFS from the DAG root ───────
        //
        // Rules:
        //   - Main-chain node i → { col: i, row: 0 }
        //   - A non-main edge forking from a node at (col, row) places its
        //     destination at (col+1, nextAvailableRow) — one row per branch,
        //     counting depth-first so nested forks get increasing row numbers.
        //   - Subsequent edges continuing a branch keep the same row and
        //     increment col normally.
        //
        // We use a queue of { nodeId, col, row } and assign layout on first visit.

        const layout = {};   // nodeId → { col, row }
        const edgeRow = {};  // edgeId → row it belongs to (for edge descriptors)

        // Seed main chain
        mainNodeIds.forEach((nid, i) => { layout[nid] = { col: i, row: 0 }; });

        // Track the next available branch row
        let nextRow = 1;

        // BFS queue: process nodes in column order so parent branches are
        // laid out before child branches that fork from them.
        // We iterate main-chain nodes in order, then discover branches from each.
        const processQueue = [...mainNodeIds];
        const processedNodes = new Set(mainNodeIds);

        // Assign main edge rows
        mainEdgeIds.forEach(eid => { edgeRow[eid] = 0; });

        let head = 0;
        while (head < processQueue.length) {
            const nodeId = processQueue[head++];
            const pos    = layout[nodeId];
            if (!pos) continue;

            const outgoing = edgesFrom[nodeId] || [];
            outgoing
                .filter(e => !mainEdgeSet.has(e.id))
                .forEach(e => {
                    // This edge forks from an existing node — it starts a new row
                    const forkRow = nextRow++;
                    edgeRow[e.id] = forkRow;

                    // Walk this branch linearly, assigning the same row
                    let col = pos.col;
                    let cur = e;
                    const visited = new Set();
                    while (cur && !visited.has(cur.id)) {
                        visited.add(cur.id);
                        edgeRow[cur.id] = edgeRow[cur.id] ?? forkRow; // first assignment wins
                        col++;
                        if (!layout[cur.to]) {
                            layout[cur.to] = { col, row: forkRow };
                        }
                        if (!processedNodes.has(cur.to)) {
                            processedNodes.add(cur.to);
                            processQueue.push(cur.to);
                        }
                        // Continue along the single non-branching successor on this row
                        const nexts = (edgesFrom[cur.to] || []).filter(ne => !mainEdgeSet.has(ne.id));
                        // Only follow if exactly one continues straight (no new forks yet)
                        cur = nexts.length === 1 && !edgeRow[nexts[0].id] ? nexts[0] : null;
                    }
                });
        }

        // ── 2. Compute grid dimensions from layout ─────────────────────────
        let maxCol = mainNodeIds.length - 1;
        let maxRow = 0;
        Object.values(layout).forEach(({ col, row }) => {
            if (col > maxCol) maxCol = col;
            if (row > maxRow) maxRow = row;
        });
        const totalCols = maxCol + 1;
        const totalRows = maxRow + 1;

        // ── 3. Build node descriptors ──────────────────────────────────────
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

        // ── 4. Build edge descriptors ──────────────────────────────────────
        const edgesOut = [];
        const seenEdges = new Set();

        // All edges that have a layout position for both endpoints
        allEdges.forEach(e => {
            if (seenEdges.has(e.id)) return;
            const fp = layout[e.from];
            const tp = layout[e.to];
            if (!fp || !tp) return;
            seenEdges.add(e.id);
            edgesOut.push({
                id:        e.id,
                fromId:    e.from,
                toId:      e.to,
                label:     e.segment?.description || '',
                isBranch:  !mainEdgeSet.has(e.id),
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
