document.addEventListener("DOMContentLoaded", () => {
    document.body.classList.add("js-ready");
    setupDeleteConfirmations();
    setupScrollPreservation();
    setupDocumentSelection();
    setupSegmentPanelSelection();
    setupNetworkExplorer();
    restoreDocumentScroll();
});

function setupDeleteConfirmations() {
    document.addEventListener("submit", (event) => {
        const form = event.target;
        if (!(form instanceof HTMLFormElement) || !form.classList.contains("confirm-delete")) {
            return;
        }
        const message = form.dataset.confirm || "Delete this item?";
        if (!window.confirm(message)) {
            event.preventDefault();
        }
    });
}

function setupScrollPreservation() {
    document.addEventListener("submit", (event) => {
        const form = event.target;
        if (!(form instanceof HTMLFormElement) || !isDocumentPageForm(form)) {
            return;
        }

        let input = form.querySelector('input[name="scroll_y"]');
        if (!input) {
            input = document.createElement("input");
            input.type = "hidden";
            input.name = "scroll_y";
            form.appendChild(input);
        }
        input.value = String(Math.max(0, Math.round(window.scrollY)));

        const textPanel = document.getElementById("document-text-panel");
        const documentScrollY = textPanel ? String(Math.max(0, Math.round(textPanel.scrollTop))) : "";
        let documentScrollInput = form.querySelector('input[name="document_scroll_y"]');
        if (!documentScrollInput) {
            documentScrollInput = document.createElement("input");
            documentScrollInput.type = "hidden";
            documentScrollInput.name = "document_scroll_y";
            form.appendChild(documentScrollInput);
        }
        documentScrollInput.value = documentScrollY;

        const nextUrlInput = form.querySelector('input[name="next_url"]');
        if (nextUrlInput && nextUrlInput.value.startsWith("/documents/")) {
            nextUrlInput.value = withScrollPosition(nextUrlInput.value, input.value, documentScrollY);
        }
    });
}

function isDocumentPageForm(form) {
    return Boolean(document.getElementById("document-text-panel")) && (
        form.classList.contains("preserve-scroll") || form.closest(".segment-card") || form.closest("#selection-helper")
    );
}

function withScrollPosition(url, scrollY, documentScrollY) {
    const parsedUrl = new URL(url, window.location.origin);
    parsedUrl.searchParams.set("scroll_y", scrollY);
    if (documentScrollY !== "") {
        parsedUrl.searchParams.set("document_scroll_y", documentScrollY);
    }
    return parsedUrl.pathname + parsedUrl.search + parsedUrl.hash;
}

function restoreDocumentScroll() {
    if (!document.getElementById("document-text-panel")) {
        return;
    }
    const url = new URL(window.location.href);
    const scrollY = parseInt(url.searchParams.get("scroll_y") || "", 10);
    const documentScrollY = parseInt(url.searchParams.get("document_scroll_y") || "", 10);
    if ((!Number.isFinite(scrollY) || scrollY < 0) && (!Number.isFinite(documentScrollY) || documentScrollY < 0)) {
        return;
    }
    window.requestAnimationFrame(() => {
        if (Number.isFinite(scrollY) && scrollY >= 0) {
            window.scrollTo(0, scrollY);
        }
        if (Number.isFinite(documentScrollY) && documentScrollY >= 0) {
            document.getElementById("document-text-panel").scrollTop = documentScrollY;
        }
        url.searchParams.delete("scroll_y");
        url.searchParams.delete("document_scroll_y");
        window.history.replaceState({}, "", url.pathname + url.search + url.hash);
    });
}

function setupDocumentSelection() {
    const textPanel = document.getElementById("document-text-panel");
    const preview = document.getElementById("selection-preview");
    const selectedTextInput = document.getElementById("selected-text-input");
    const startOffsetInput = document.getElementById("start-offset-input");
    const endOffsetInput = document.getElementById("end-offset-input");
    const createButton = document.getElementById("create-segment-button");
    const helper = document.getElementById("selection-helper");
    const clearButton = document.getElementById("clear-selection-button");

    if (!textPanel || !preview || !selectedTextInput || !startOffsetInput || !endOffsetInput || !createButton || !helper) {
        return;
    }

    let isInteractingWithSelectionHelper = false;
    document.addEventListener("pointerdown", (event) => {
        isInteractingWithSelectionHelper = helper.contains(event.target);
    }, true);
    helper.addEventListener("focusin", () => {
        isInteractingWithSelectionHelper = true;
    });

    if (clearButton) {
        clearButton.addEventListener("click", () => {
            window.getSelection().removeAllRanges();
            clearSegmentSelection(preview, selectedTextInput, startOffsetInput, endOffsetInput, createButton, helper);
            showEmptySegmentPanel();
        });
    }

    document.addEventListener("selectionchange", () => {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0 || selection.isCollapsed) {
            if (isInteractingWithSelectionHelper && selectedTextInput.value.trim()) {
                return;
            }
            clearSegmentSelection(preview, selectedTextInput, startOffsetInput, endOffsetInput, createButton, helper);
            return;
        }

        const range = selection.getRangeAt(0);
        if (!textPanel.contains(range.commonAncestorContainer)) {
            if (isInteractingWithSelectionHelper && selectedTextInput.value.trim()) {
                return;
            }
            clearSegmentSelection(preview, selectedTextInput, startOffsetInput, endOffsetInput, createButton, helper);
            return;
        }

        const selectedText = selection.toString();
        if (!selectedText.trim()) {
            clearSegmentSelection(preview, selectedTextInput, startOffsetInput, endOffsetInput, createButton, helper);
            return;
        }

        const offsets = getOffsetsWithinContainer(textPanel, range);
        if (!offsets || offsets.end <= offsets.start) {
            clearSegmentSelection(preview, selectedTextInput, startOffsetInput, endOffsetInput, createButton, helper);
            return;
        }

        preview.textContent = selectedText;
        preview.classList.remove("empty");
        selectedTextInput.value = selectedText;
        startOffsetInput.value = String(offsets.start);
        endOffsetInput.value = String(offsets.end);
        createButton.disabled = false;
        helper.classList.remove("hidden");
        showCreateSegmentPanel();
        document.body.classList.add("selection-bar-visible");
    });
}

function setupSegmentPanelSelection() {
    const textPanel = document.getElementById("document-text-panel");
    const picker = document.getElementById("segment-picker");
    if (!textPanel) {
        return;
    }

    if (picker) {
        picker.addEventListener("change", () => {
            if (picker.value) {
                showSegmentPanel(picker.value);
                scrollSegmentHighlightIntoView(picker.value);
            } else {
                showEmptySegmentPanel();
            }
        });
    }

    textPanel.addEventListener("click", (event) => {
        const highlight = event.target.closest(".segment-highlight");
        if (!highlight || !textPanel.contains(highlight)) {
            return;
        }
        const segmentId = highlight.dataset.segmentId;
        if (segmentId) {
            window.getSelection().removeAllRanges();
            showSegmentPanel(segmentId);
        }
    });
}

function showCreateSegmentPanel() {
    hideSegmentPanels();
    const helper = document.getElementById("selection-helper");
    const picker = document.getElementById("segment-picker");
    if (helper) {
        helper.classList.remove("hidden");
    }
    if (picker) {
        picker.value = "";
    }
}

function showSegmentPanel(segmentId) {
    hideSegmentPanels();
    const escapedSegmentId = escapeSelectorValue(segmentId);
    const panel = document.querySelector(`[data-segment-panel="${escapedSegmentId}"]`);
    const picker = document.getElementById("segment-picker");
    if (panel) {
        panel.classList.remove("hidden");
    }
    if (picker) {
        picker.value = segmentId;
    }
    document.querySelectorAll(".segment-highlight.is-active").forEach((highlight) => {
        highlight.classList.remove("is-active");
    });
    document.querySelectorAll(`.segment-highlight[data-segment-id="${escapedSegmentId}"]`).forEach((highlight) => {
        highlight.classList.add("is-active");
    });
    document.body.classList.remove("selection-bar-visible");
}

function showEmptySegmentPanel() {
    hideSegmentPanels();
    const emptyPanel = document.getElementById("segment-panel-empty");
    if (emptyPanel) {
        emptyPanel.classList.remove("hidden");
    }
    document.querySelectorAll(".segment-highlight.is-active").forEach((highlight) => {
        highlight.classList.remove("is-active");
    });
}

function hideSegmentPanels() {
    document.querySelectorAll("[data-segment-panel]").forEach((panel) => {
        panel.classList.add("hidden");
    });
    const emptyPanel = document.getElementById("segment-panel-empty");
    if (emptyPanel) {
        emptyPanel.classList.add("hidden");
    }
    const helper = document.getElementById("selection-helper");
    if (helper) {
        helper.classList.add("hidden");
    }
}

function scrollSegmentHighlightIntoView(segmentId) {
    const highlight = document.querySelector(`.segment-highlight[data-segment-id="${escapeSelectorValue(segmentId)}"]`);
    if (highlight) {
        highlight.scrollIntoView({ block: "center", behavior: "smooth" });
    }
}

function escapeSelectorValue(value) {
    if (window.CSS && typeof window.CSS.escape === "function") {
        return window.CSS.escape(value);
    }
    return String(value).replace(/["\\]/g, "\\$&");
}

function getOffsetsWithinContainer(container, range) {
    const beforeSelection = document.createRange();
    beforeSelection.selectNodeContents(container);

    try {
        beforeSelection.setEnd(range.startContainer, range.startOffset);
    } catch (error) {
        return null;
    }

    return {
        start: beforeSelection.toString().length,
        end: beforeSelection.toString().length + range.toString().length,
    };
}

function clearSegmentSelection(preview, selectedTextInput, startOffsetInput, endOffsetInput, createButton, helper) {
    preview.textContent = "No text selected.";
    preview.classList.add("empty");
    selectedTextInput.value = "";
    startOffsetInput.value = "";
    endOffsetInput.value = "";
    createButton.disabled = true;
    if (helper) {
        helper.classList.add("hidden");
    }
    document.body.classList.remove("selection-bar-visible");
}

function setupNetworkExplorer() {
    const svg = document.getElementById("cooccurrence-network");
    if (!svg) {
        return;
    }

    const summary = document.getElementById("network-summary");
    const empty = document.getElementById("network-empty");
    const detail = document.getElementById("network-detail-panel");
    const tooltip = document.getElementById("network-tooltip");
    const jsonDownload = document.getElementById("network-json-download");
    const csvDownload = document.getElementById("network-csv-download");
    const params = new URLSearchParams(window.location.search);
    const query = params.toString();
    const suffix = query ? `?${query}` : "";

    if (jsonDownload) {
        jsonDownload.href = `/exports/cooccurrence-network.json${suffix}`;
    }
    if (csvDownload) {
        csvDownload.href = `/exports/cooccurrence-edges.csv${suffix}`;
    }

    fetch(`/network/data${suffix}`)
        .then((response) => {
            if (!response.ok) {
                throw new Error("Network data request failed.");
            }
            return response.json();
        })
        .then((graph) => renderNetworkGraph(graph, svg, summary, empty, detail, tooltip))
        .catch((error) => {
            if (summary) {
                summary.textContent = error.message;
            }
        });
}

function renderNetworkGraph(graph, svg, summary, empty, detail, tooltip) {
    const nodes = graph.nodes || [];
    const edges = graph.edges || [];
    svg.textContent = "";
    if (summary) {
        summary.textContent = `${nodes.length} nodes · ${edges.length} edges · min weight ${graph.meta.min_weight}`;
    }
    if (!nodes.length || !edges.length) {
        if (empty) {
            empty.classList.remove("hidden");
        }
        return;
    }
    if (empty) {
        empty.classList.add("hidden");
    }

    const width = Math.max(760, svg.clientWidth || 960);
    const height = Math.max(520, svg.clientHeight || 620);
    svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
    const nodeMap = new Map(nodes.map((node) => [node.id, node]));
    const positions = computeNetworkPositions(nodes, edges, width, height, graph.meta.layout || "columns");
    const edgeLayer = makeSvgElement("g", { class: "network-edge-layer" });
    const nodeLayer = makeSvgElement("g", { class: "network-node-layer" });
    svg.append(edgeLayer, nodeLayer);

    const edgeElements = [];
    edges.forEach((edge) => {
        const source = positions.get(edge.source);
        const target = positions.get(edge.target);
        if (!source || !target) {
            return;
        }
        const line = makeSvgElement("line", {
            class: "network-edge",
            x1: source.x,
            y1: source.y,
            x2: target.x,
            y2: target.y,
            "stroke-width": Math.min(8, 1 + edge.weight),
            "data-source": edge.source,
            "data-target": edge.target,
        });
        line.addEventListener("mouseenter", (event) => {
            highlightNetwork(svg, edge.source, edge.target);
            showNetworkTooltip(tooltip, event, `Co-occurs in ${edge.weight} segment${edge.weight === 1 ? "" : "s"}`);
        });
        line.addEventListener("mousemove", (event) => moveNetworkTooltip(tooltip, event));
        line.addEventListener("mouseleave", () => {
            clearNetworkHighlight(svg);
            hideNetworkTooltip(tooltip);
        });
        line.addEventListener("click", () => showEdgeDetail(edge, nodeMap, detail));
        edgeLayer.appendChild(line);
        edgeElements.push({ edge, element: line });
    });

    nodes.forEach((node) => {
        const position = positions.get(node.id);
        if (!position) {
            return;
        }
        const group = makeSvgElement("g", {
            class: `network-node node-${node.type}`,
            transform: `translate(${position.x}, ${position.y})`,
            "data-node-id": node.id,
        });
        const radius = networkNodeRadius(node);
        const circle = makeSvgElement("circle", {
            r: radius,
            fill: networkNodeColor(node),
        });
        const label = makeSvgElement("text", {
            y: radius + 13,
            "text-anchor": "middle",
        });
        label.textContent = truncateLabel(node.label, 24);
        group.append(circle, label);
        group.addEventListener("mouseenter", (event) => {
            highlightNetwork(svg, node.id, null);
            showNetworkTooltip(tooltip, event, `${node.label}<br>${node.type} · ${node.subtype || "none"} · ${node.count}`);
        });
        group.addEventListener("mousemove", (event) => moveNetworkTooltip(tooltip, event));
        group.addEventListener("mouseleave", () => {
            clearNetworkHighlight(svg);
            hideNetworkTooltip(tooltip);
        });
        group.addEventListener("click", () => showNodeDetail(node, edges, nodeMap, detail));
        nodeLayer.appendChild(group);
    });
}

function computeNetworkPositions(nodes, edges, width, height, layout) {
    if (layout === "circle") {
        return computeCirclePositions(nodes, width, height);
    }
    if (layout === "force") {
        return computeSimpleForcePositions(nodes, edges, width, height);
    }
    return computeColumnPositions(nodes, width, height);
}

function computeColumnPositions(nodes, width, height) {
    const columns = ["code", "discourse_marker", "actor", "discourse_feature"];
    const grouped = new Map(columns.map((column) => [column, []]));
    nodes.forEach((node) => {
        const bucket = grouped.get(node.type) || grouped.get("discourse_feature");
        bucket.push(node);
    });
    const positions = new Map();
    columns.forEach((column, columnIndex) => {
        const bucket = grouped.get(column);
        const x = 80 + columnIndex * ((width - 160) / Math.max(1, columns.length - 1));
        bucket.sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
        bucket.forEach((node, rowIndex) => {
            const y = 70 + rowIndex * ((height - 140) / Math.max(1, bucket.length - 1));
            positions.set(node.id, { x, y });
        });
    });
    return positions;
}

function computeCirclePositions(nodes, width, height) {
    const positions = new Map();
    const cx = width / 2;
    const cy = height / 2;
    const radius = Math.min(width, height) * 0.38;
    nodes.forEach((node, index) => {
        const angle = (Math.PI * 2 * index) / Math.max(1, nodes.length);
        positions.set(node.id, {
            x: cx + Math.cos(angle) * radius,
            y: cy + Math.sin(angle) * radius,
        });
    });
    return positions;
}

function computeSimpleForcePositions(nodes, edges, width, height) {
    const positions = computeCirclePositions(nodes, width, height);
    const nodeIds = nodes.map((node) => node.id);
    for (let iteration = 0; iteration < 90; iteration += 1) {
        const forces = new Map(nodeIds.map((id) => [id, { x: 0, y: 0 }]));
        for (let i = 0; i < nodeIds.length; i += 1) {
            for (let j = i + 1; j < nodeIds.length; j += 1) {
                applyRepulsion(positions, forces, nodeIds[i], nodeIds[j]);
            }
        }
        edges.forEach((edge) => applyAttraction(positions, forces, edge.source, edge.target, edge.weight));
        nodeIds.forEach((id) => {
            const position = positions.get(id);
            const force = forces.get(id);
            position.x = clamp(position.x + force.x * 0.04, 40, width - 40);
            position.y = clamp(position.y + force.y * 0.04, 40, height - 40);
        });
    }
    return positions;
}

function applyRepulsion(positions, forces, a, b) {
    const pa = positions.get(a);
    const pb = positions.get(b);
    const dx = pa.x - pb.x || 1;
    const dy = pa.y - pb.y || 1;
    const distanceSq = Math.max(80, dx * dx + dy * dy);
    const strength = 9000 / distanceSq;
    forces.get(a).x += dx * strength;
    forces.get(a).y += dy * strength;
    forces.get(b).x -= dx * strength;
    forces.get(b).y -= dy * strength;
}

function applyAttraction(positions, forces, a, b, weight) {
    const pa = positions.get(a);
    const pb = positions.get(b);
    if (!pa || !pb) {
        return;
    }
    const dx = pb.x - pa.x;
    const dy = pb.y - pa.y;
    const strength = Math.min(0.35, 0.03 * weight);
    forces.get(a).x += dx * strength;
    forces.get(a).y += dy * strength;
    forces.get(b).x -= dx * strength;
    forces.get(b).y -= dy * strength;
}

function showNodeDetail(node, edges, nodeMap, detail) {
    if (!detail) {
        return;
    }
    const connected = edges
        .filter((edge) => edge.source === node.id || edge.target === node.id)
        .sort((a, b) => b.weight - a.weight)
        .slice(0, 12)
        .map((edge) => {
            const otherId = edge.source === node.id ? edge.target : edge.source;
            const other = nodeMap.get(otherId);
            return `<li><span>${escapeHtml(other ? other.label : otherId)}</span><span>${edge.weight}</span></li>`;
        })
        .join("");
    detail.innerHTML = `
        <h2>${escapeHtml(node.label)}</h2>
        <dl class="metadata-list compact-metadata">
            <div><dt>Type</dt><dd>${escapeHtml(node.type)}</dd></div>
            <div><dt>Subtype</dt><dd>${escapeHtml(node.subtype || "None")}</dd></div>
            <div><dt>Count</dt><dd>${node.count}</dd></div>
        </dl>
        <h3>Connected nodes</h3>
        <ul class="compact-list">${connected || "<li><span>No visible connections.</span></li>"}</ul>
    `;
}

function showEdgeDetail(edge, nodeMap, detail) {
    if (!detail) {
        return;
    }
    const source = nodeMap.get(edge.source);
    const target = nodeMap.get(edge.target);
    const previews = (edge.segment_previews || []).map((segment) => `
        <li>
            <a href="/documents/${segment.document_id}#segment-${segment.segment_id}">${escapeHtml(segment.segment_title)}</a>
            <span>${escapeHtml(segment.document_title)}</span>
            <p>${escapeHtml(segment.text_preview)}</p>
        </li>
    `).join("");
    detail.innerHTML = `
        <h2>Co-occurrence</h2>
        <p><strong>${escapeHtml(source ? source.label : edge.source)}</strong> ↔ <strong>${escapeHtml(target ? target.label : edge.target)}</strong></p>
        <dl class="metadata-list compact-metadata">
            <div><dt>Weight</dt><dd>${edge.weight}</dd></div>
            <div><dt>Segments</dt><dd>${edge.segment_count || edge.segments.length}</dd></div>
            <div><dt>Relation</dt><dd>${escapeHtml(edge.relation)}</dd></div>
        </dl>
        <h3>Segment previews</h3>
        <ul class="segment-preview-list">${previews || "<li>No previews available.</li>"}</ul>
    `;
}

function highlightNetwork(svg, sourceId, targetId) {
    const highlightedNodeIds = new Set([sourceId]);
    if (targetId) {
        highlightedNodeIds.add(targetId);
    }
    svg.querySelectorAll(".network-node, .network-edge").forEach((element) => {
        element.classList.add("is-dimmed");
        element.classList.remove("is-highlighted");
    });
    svg.querySelectorAll(".network-edge").forEach((edge) => {
        const source = edge.getAttribute("data-source");
        const target = edge.getAttribute("data-target");
        if (source === sourceId || target === sourceId || (targetId && source === targetId) || (targetId && target === targetId)) {
            highlightedNodeIds.add(source);
            highlightedNodeIds.add(target);
            edge.classList.remove("is-dimmed");
            edge.classList.add("is-highlighted");
        }
    });
    svg.querySelectorAll(".network-node").forEach((node) => {
        if (highlightedNodeIds.has(node.dataset.nodeId)) {
            node.classList.remove("is-dimmed");
            node.classList.add("is-highlighted");
        }
    });
}

function clearNetworkHighlight(svg) {
    svg.querySelectorAll(".network-node, .network-edge").forEach((element) => {
        element.classList.remove("is-dimmed", "is-highlighted");
    });
}

function showNetworkTooltip(tooltip, event, html) {
    if (!tooltip) {
        return;
    }
    tooltip.innerHTML = html;
    tooltip.classList.remove("hidden");
    moveNetworkTooltip(tooltip, event);
}

function moveNetworkTooltip(tooltip, event) {
    if (!tooltip) {
        return;
    }
    tooltip.style.left = `${event.offsetX + 16}px`;
    tooltip.style.top = `${event.offsetY + 16}px`;
}

function hideNetworkTooltip(tooltip) {
    if (tooltip) {
        tooltip.classList.add("hidden");
    }
}

function makeSvgElement(tagName, attributes) {
    const element = document.createElementNS("http://www.w3.org/2000/svg", tagName);
    Object.entries(attributes || {}).forEach(([key, value]) => {
        element.setAttribute(key, value);
    });
    return element;
}

function networkNodeRadius(node) {
    return Math.min(28, 9 + Math.sqrt(Math.max(1, node.count)) * 4);
}

function networkNodeColor(node) {
    if (node.color) {
        return node.color;
    }
    if (node.type === "code") {
        return node.subtype === "category" ? "#6d4bb2" : node.subtype === "axial" ? "#4a6fd0" : "#245aa5";
    }
    if (node.type === "discourse_marker") {
        return "#2c8f8a";
    }
    if (node.type === "actor") {
        return "#b36b1f";
    }
    return "#4b9f68";
}

function truncateLabel(label, length) {
    return label.length > length ? `${label.slice(0, length - 1)}...` : label;
}

function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
