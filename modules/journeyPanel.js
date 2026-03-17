/**
 * journeyPanel.js — Session-wide visual DAG showing all worksheet states
 * (nodes) and code segment transitions (edges) across every task.
 *
 * Layout: horizontal left-to-right, branches go downward.
 * Nodes are clickable — clicking navigates the worksheet to that state
 * by applying undo_code (backward) or code (forward) silently.
 *
 * No edit feature — read-only exploration only.
 */
const JourneyPanel = (() => {

    // ── DOM ───────────────────────────────────────────────────────────────────
    let _panel       = null;  // created lazily
    let _svgWrap     = null;
    let _currentNodeId = null; // which node the worksheet is currently at

    // ── Layout constants ──────────────────────────────────────────────────────
    const NODE_R  = 9;
    const COL_W   = 80;   // horizontal spacing between columns
    const ROW_H   = 44;   // vertical spacing between branches
    const PAD_X   = 24;
    const PAD_Y   = 28;
    const SVG_NS  = 'http://www.w3.org/2000/svg';

    // ── Public ────────────────────────────────────────────────────────────────

    function init() {
        _buildPanel();
    }

    /**
     * Show the panel and re-render the DAG.
     * @param {string} currentNodeId — which node the sheet is currently at
     */
    function show(currentNodeId) {
        _currentNodeId = currentNodeId;
        _panel.classList.add('visible');
        _render();
    }

    function hide() {
        _panel.classList.remove('visible');
    }

    function toggle(currentNodeId) {
        if (_panel.classList.contains('visible')) { hide(); }
        else { show(currentNodeId); }
    }

    /** Call after any DAG change to refresh if open. */
    function refresh(currentNodeId) {
        if (currentNodeId !== undefined) _currentNodeId = currentNodeId;
        if (_panel.classList.contains('visible')) _render();
    }

    // ── Panel construction ────────────────────────────────────────────────────

    function _buildPanel() {
        _panel = document.createElement('div');
        _panel.id = 'journey-panel';
        _panel.innerHTML = `
            <div id="journey-header">
                <span id="journey-title">📍 Journey</span>
                <button id="journey-close">✕</button>
            </div>
            <div id="journey-hint">Click any node to navigate your worksheet to that state.</div>
            <div id="journey-svg-wrap"></div>
            <div id="journey-tooltip"></div>`;
        document.body.appendChild(_panel);
        _svgWrap = _panel.querySelector('#journey-svg-wrap');
        _panel.querySelector('#journey-close').addEventListener('click', hide);
    }

    // ── Layout engine ─────────────────────────────────────────────────────────

    /**
     * Assign (col, row) to each node.
     * Algorithm:
     *   - BFS from all root nodes left→right assigns columns (depth).
     *   - When a node has multiple outgoing edges, each child beyond the first
     *     gets a new row below the parent's row (branching downward).
     */
    function _layoutNodes(nodes, edges) {
        const fwd = {}; // nodeId → [toNodeId]
        edges.forEach(e => {
            if (!fwd[e.from]) fwd[e.from] = [];
            fwd[e.from].push(e.to);
        });
        const bwd = {};
        edges.forEach(e => {
            if (!bwd[e.to]) bwd[e.to] = [];
            bwd[e.to].push(e.from);
        });

        // Find roots (nodes with no incoming edges)
        const allToIds = new Set(edges.map(e => e.to));
        const roots = nodes.filter(n => !allToIds.has(n.id));

        const pos = {};   // nodeId → {col, row}
        let nextRow = 0;

        function bfsFrom(rootId, startRow) {
            const queue = [{ id: rootId, col: 0, row: startRow }];
            const seen  = new Set();
            while (queue.length) {
                const { id, col, row } = queue.shift();
                if (seen.has(id)) continue;
                seen.add(id);
                if (!pos[id]) pos[id] = { col, row };
                const children = fwd[id] || [];
                children.forEach((childId, ci) => {
                    if (seen.has(childId)) return;
                    const childRow = ci === 0 ? row : nextRow++;
                    queue.push({ id: childId, col: col + 1, row: childRow });
                });
            }
        }

        roots.forEach(r => {
            bfsFrom(r.id, nextRow);
            nextRow++;
        });

        // Any remaining nodes not reached (disconnected)
        nodes.forEach(n => {
            if (!pos[n.id]) { pos[n.id] = { col: 0, row: nextRow++ }; }
        });

        return pos;
    }

    // ── Rendering ─────────────────────────────────────────────────────────────

    function _render() {
        _svgWrap.innerHTML = '';
        const { nodes, edges } = DagStore.getAll();

        if (nodes.length === 0) {
            _svgWrap.innerHTML = '<div style="color:rgba(255,255,255,0.4);font-size:11px;padding:20px 8px;text-align:center">No steps yet. Send a message to get started.</div>';
            return;
        }

        const pos = _layoutNodes(nodes, edges);

        // SVG size
        const maxCol = Math.max(...Object.values(pos).map(p => p.col));
        const maxRow = Math.max(...Object.values(pos).map(p => p.row));
        const svgW   = PAD_X * 2 + (maxCol + 1) * COL_W;
        const svgH   = PAD_Y * 2 + (maxRow + 1) * ROW_H;

        const svg = document.createElementNS(SVG_NS, 'svg');
        svg.setAttribute('viewBox', `0 0 ${svgW} ${svgH}`);
        svg.setAttribute('width',  svgW);
        svg.setAttribute('height', svgH);

        const cx = id => PAD_X + pos[id].col * COL_W;
        const cy = id => PAD_Y + pos[id].row * ROW_H;

        // Group edges by task to colour them distinctly
        const taskLabels = [...new Set(nodes.map(n => n.taskLabel).filter(Boolean))];
        const taskColor = label => {
            const i = taskLabels.indexOf(label);
            const palette = ['#4f8ef7','#3ecf8e','#f5a623','#c084fc','#fb7185','#38bdf8'];
            return palette[i % palette.length];
        };

        // ── Draw edges ────────────────────────────────────────────────────────
        edges.forEach(e => {
            const x1 = cx(e.from), y1 = cy(e.from);
            const x2 = cx(e.to),   y2 = cy(e.to);
            const fromNode = nodes.find(n => n.id === e.from);
            const color = taskColor(fromNode?.taskLabel);

            // Curved path for branches, straight for same row
            let d;
            if (y1 === y2) {
                d = `M${x1} ${y1} L${x2} ${y2}`;
            } else {
                const mx = (x1 + x2) / 2;
                d = `M${x1} ${y1} C${mx} ${y1} ${mx} ${y2} ${x2} ${y2}`;
            }

            const path = document.createElementNS(SVG_NS, 'path');
            path.setAttribute('d', d);
            path.setAttribute('fill', 'none');
            path.setAttribute('stroke', e.executed ? color : 'rgba(255,255,255,0.15)');
            path.setAttribute('stroke-width', e.executed ? '2' : '1.5');
            path.setAttribute('stroke-dasharray', e.executed ? 'none' : '4 3');

            // Arrowhead at midpoint
            if (e.executed) {
                const mx = (x1 + x2) / 2;
                const my = (y1 + y2) / 2;
                const marker = document.createElementNS(SVG_NS, 'polygon');
                const angle  = Math.atan2(y2 - y1, x2 - x1);
                const tip    = [mx + Math.cos(angle) * 5, my + Math.sin(angle) * 5];
                const left   = [mx - Math.cos(angle) * 5 + Math.sin(angle) * 4,
                                 my - Math.sin(angle) * 5 - Math.cos(angle) * 4];
                const right  = [mx - Math.cos(angle) * 5 - Math.sin(angle) * 4,
                                 my - Math.sin(angle) * 5 + Math.cos(angle) * 4];
                marker.setAttribute('points', `${tip} ${left} ${right}`);
                marker.setAttribute('fill', color);
                svg.appendChild(marker);
            }

            // Hover label on edge
            const title = document.createElementNS(SVG_NS, 'title');
            title.textContent = e.segment?.description || '';
            path.appendChild(title);

            svg.appendChild(path);
        });

        // ── Draw nodes ────────────────────────────────────────────────────────
        nodes.forEach(node => {
            const x = cx(node.id);
            const y = cy(node.id);
            const isCurrent = node.id === _currentNodeId;
            const isRoot    = node.isRoot;
            const color     = taskColor(node.taskLabel);

            // Determine execution state
            const inEdge    = edges.find(e => e.to === node.id);
            const isExecuted = isRoot || (inEdge && inEdge.executed);
            const isFailed   = inEdge && inEdge.failed;

            const g = document.createElementNS(SVG_NS, 'g');
            g.style.cursor = 'pointer';

            // Outer ring (current node indicator)
            if (isCurrent) {
                const ring = document.createElementNS(SVG_NS, 'circle');
                ring.setAttribute('cx', x); ring.setAttribute('cy', y);
                ring.setAttribute('r', NODE_R + 4);
                ring.setAttribute('fill', 'none');
                ring.setAttribute('stroke', '#fff');
                ring.setAttribute('stroke-width', '2');
                ring.setAttribute('opacity', '0.6');
                g.appendChild(ring);
            }

            // Main circle
            const circle = document.createElementNS(SVG_NS, 'circle');
            circle.setAttribute('cx', x); circle.setAttribute('cy', y);
            circle.setAttribute('r', NODE_R);
            if (isFailed) {
                circle.setAttribute('fill', 'rgba(245,100,60,0.85)');
                circle.setAttribute('stroke', '#f5643c');
            } else if (isCurrent) {
                circle.setAttribute('fill', color);
                circle.setAttribute('stroke', '#fff');
            } else if (isExecuted) {
                circle.setAttribute('fill', isRoot ? color : color.replace(')', ',0.7)').replace('rgb', 'rgba'));
                circle.setAttribute('stroke', color);
            } else {
                circle.setAttribute('fill', 'rgba(255,255,255,0.05)');
                circle.setAttribute('stroke', 'rgba(255,255,255,0.2)');
            }
            circle.setAttribute('stroke-width', isCurrent ? '2.5' : '1.5');
            g.appendChild(circle);

            // Tick / dot inside
            if (isExecuted && !isCurrent && !isRoot) {
                const tick = document.createElementNS(SVG_NS, 'circle');
                tick.setAttribute('cx', x); tick.setAttribute('cy', y);
                tick.setAttribute('r', 3);
                tick.setAttribute('fill', '#fff');
                tick.setAttribute('opacity', '0.9');
                g.appendChild(tick);
            }

            // Label below node
            const text = document.createElementNS(SVG_NS, 'text');
            text.setAttribute('x', x);
            text.setAttribute('y', y + NODE_R + 11);
            text.setAttribute('text-anchor', 'middle');
            text.setAttribute('font-size', '8');
            text.setAttribute('fill', isCurrent ? '#fff' : 'rgba(255,255,255,0.55)');
            text.setAttribute('font-family', 'Segoe UI, system-ui, sans-serif');
            const labelWords = (isRoot ? (node.taskLabel || 'Start') : node.label).split(' ');
            text.textContent = labelWords.slice(0, 3).join(' ') + (labelWords.length > 3 ? '…' : '');
            g.appendChild(text);

            // Tooltip via title
            const title = document.createElementNS(SVG_NS, 'title');
            title.textContent = (isRoot ? '📌 ' : '') + node.label +
                (isCurrent ? ' (current)' : '') +
                (node.taskLabel ? `\nTask: ${node.taskLabel}` : '');
            g.appendChild(title);

            // Click → navigate
            g.addEventListener('click', () => _onNodeClick(node));
            g.addEventListener('mouseenter', () => _showTooltip(node, x, y));
            g.addEventListener('mouseleave', () => _hideTooltip());

            svg.appendChild(g);
        });

        _svgWrap.appendChild(svg);
    }

    // ── Node click navigation ─────────────────────────────────────────────────

    async function _onNodeClick(node) {
        if (!_currentNodeId) return;
        if (node.id === _currentNodeId) return; // already here

        const path = DagStore.findPath(_currentNodeId, node.id);
        if (!path) {
            _showBanner('⚠ No path found to this state.');
            return;
        }
        if (path.length === 0) return;

        _showBanner('Navigating… please wait');

        try {
            for (const step of path) {
                const code = step.direction === 'forward'
                    ? step.edge.segment?.code
                    : step.edge.segment?.undo_code;
                if (code) {
                    const fn = new (Object.getPrototypeOf(async function(){}).constructor)(code);
                    await fn();
                }
            }
            _currentNodeId = node.id;
            _hideBanner();
            _render(); // refresh to show new current node

            // Also show the incoming edge's segment info if StepNavigator is open
            const inEdge = DagStore.getIncomingEdge(node.id);
            if (inEdge?.segment) {
                JourneyPanel.onNavigate?.(inEdge.segment, node);
            }
        } catch(err) {
            _showBanner(`⚠ Navigation failed: ${err.message}`);
            console.error('[JourneyPanel] navigation error', err);
        }
    }

    // ── Tooltip ───────────────────────────────────────────────────────────────

    function _showTooltip(node, svgX, svgY) {
        const tooltip = _panel.querySelector('#journey-tooltip');
        const inEdge  = DagStore.getIncomingEdge(node.id);
        const lines   = [
            node.isRoot ? '📌 Start state' : node.label,
            inEdge ? `← ${inEdge.segment?.description || 'edge'}` : '',
            node.taskLabel ? `Task: "${node.taskLabel}"` : '',
        ].filter(Boolean);
        tooltip.textContent = lines.join('\n');
        tooltip.style.display = 'block';
    }

    function _hideTooltip() {
        const tooltip = _panel.querySelector('#journey-tooltip');
        tooltip.style.display = 'none';
    }

    // ── Status banner ─────────────────────────────────────────────────────────

    function _showBanner(msg) {
        let banner = _panel.querySelector('#journey-banner');
        if (!banner) {
            banner = document.createElement('div');
            banner.id = 'journey-banner';
            _panel.querySelector('#journey-header').after(banner);
        }
        banner.textContent = msg;
        banner.style.display = 'block';
    }

    function _hideBanner() {
        const b = _panel.querySelector('#journey-banner');
        if (b) b.style.display = 'none';
    }

    return { init, show, hide, toggle, refresh };
})();
