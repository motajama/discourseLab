# discourseLab

Phase: 3 — Segment creation

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
3. Import a TXT or DOCX document if none exists.
4. Open a document.
5. Select a passage inside the document text panel.
6. Confirm the floating segment helper appears.
7. Add an optional segment name and note.
8. Click `Create segment`.
9. Confirm the segment appears in the segment list.
10. Confirm the segment is highlighted in the document text.
11. Delete the segment.
12. Confirm it disappears from the list and highlight.
13. Confirm dashboard segment count updates.
14. Confirm audit log shows `create_segment` and `delete_segment` actions.
15. Restart the app and confirm saved segments persist.

## Project Structure

```text
discourseLab/
  disLab/
    app.py
    requirements.txt
    README.md
    schema.sql
    instance/
      disLab.sqlite
    templates/
      base.html
      dashboard.html
      document_view.html
      documents.html
      codes.html
      memos.html
      gt_workspace.html
      cda_workspace.html
      exports.html
      error.html
    static/
      style.css
      app.js
    uploads/
    exports/
```

## Current Phase 3

Phase 3 adds segment creation and segment management:

- Document reading view for close analysis
- Text selection capture inside the document text panel
- Offset calculation relative to extracted plain text
- Segment creation with optional names and notes
- Floating segment helper so segments can be saved without scrolling
- Segment list on the document detail page
- Segment deletion through POST
- Neutral visual highlighting for saved segments
- Audit log entries for segment creation and deletion
- Dashboard segment count and latest segments panel
- `/health` route for a simple local health check

This phase does not implement open code creation, code assignment, code colors, axial coding, category coding, codebook management, memo editing, CDA markers, actor maps, voice/silence reports, model building, export packages, AI features, authentication, or cloud functionality.

## Planned Later Phases

- Phase 4: Add open coding and color highlighting
- Phase 5: Add memos and codebook
- Phase 6: Add grounded theory and CDA workspaces
- Later: Export codebooks, coded segments, memos, research models, Markdown, CSV, JSON, PNG/SVG, and LaTeX/TikZ
