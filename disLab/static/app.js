document.addEventListener("DOMContentLoaded", () => {
    document.body.classList.add("js-ready");
    setupDocumentSelection();
});

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
            return;
        }

        const range = selection.getRangeAt(0);
        if (!textPanel.contains(range.commonAncestorContainer)) {
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
}
