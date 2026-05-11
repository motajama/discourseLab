# discourseLab

Phase: 5 — Memos and basic codebook

discourseLab is a local lightweight qualitative analysis workspace for discourse analysis, critical discourse analysis, and grounded theory.

## Installation

```bash
cd disLab
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

The application runs at:

```text
http://127.0.0.1:5000
```

## Dependencies

- Flask
- python-docx

## Verification

1. Start the app with `python app.py`.
2. Open `http://127.0.0.1:5000`.
3. Go to Memos.
4. Create a project memo.
5. Create a memo linked to a document.
6. Create a memo linked to a segment.
7. Create a memo linked to a code.
8. Edit a memo.
9. Filter memos by type and status.
10. Delete a memo.
11. Go to Codes.
12. Open a code detail page.
13. Edit a code and fill in definition, include_when, exclude_when, example, and analytical_note.
14. Confirm the codebook completeness indicator changes.
15. Go to Exports.
16. Download codebook Markdown.
17. Confirm dashboard memo count updates.
18. Confirm audit log shows memo actions.
19. Restart the app and confirm memos and codebook fields persist.

## Current Phase 5

Phase 5 adds memos and a basic codebook:

- Project, document, segment, code, methodological, theoretical, reflexive, comparison, and negative case memos
- Memo statuses: draft, important, use in article, archived
- Memo creation, editing, deletion, filtering, and linked entity labels
- Document-level and segment-level memos in the document reading view
- Code detail pages with linked memos and assigned segments
- Codebook fields for open codes
- Codebook completeness indicators
- Markdown codebook export at `/exports/codebook.md`

This phase does not implement axial coding, category coding, GT workspace functionality, CDA markers, actor maps, voice/silence reports, relation model building, PNG/SVG/TikZ model export, coded segments CSV export, full project JSON export, AI features, authentication, or cloud functionality.

## Planned Later Phases

- Phase 6: Add grounded theory and CDA workspaces
- Later: Add axial coding, category coding, model building, and export packages
