document.addEventListener("DOMContentLoaded", () => {
    document.body.classList.add("js-ready");
    setupDeleteConfirmations();
    setupScrollPreservation();
    setupDocumentSelection();
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

    if (clearButton) {
        clearButton.addEventListener("click", () => {
            window.getSelection().removeAllRanges();
            clearSegmentSelection(preview, selectedTextInput, startOffsetInput, endOffsetInput, createButton, helper);
        });
    }

    document.addEventListener("selectionchange", () => {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0 || selection.isCollapsed) {
            clearSegmentSelection(preview, selectedTextInput, startOffsetInput, endOffsetInput, createButton, helper);
            return;
        }

        const range = selection.getRangeAt(0);
        if (!textPanel.contains(range.commonAncestorContainer)) {
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
        document.body.classList.add("selection-bar-visible");
    });
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
