# discourseLab

Phase: 4 — Open coding

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
3. Go to Codes.
4. Create a new open code with a custom color.
5. Edit the code.
6. Confirm the code appears in the open codes table.
7. Open an imported document.
8. Create a segment if none exists.
9. Assign the open code to a segment.
10. Confirm the segment shows the code badge.
11. Confirm the segment highlight uses the code color.
12. Remove the code from the segment.
13. Confirm the segment returns to neutral highlight.
14. Delete the code.
15. Confirm related segment assignments are removed.
16. Confirm dashboard counts update.
17. Confirm audit log shows code and assignment actions.
18. Restart the app and confirm codes and assignments persist.

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
      code_edit.html
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

## Current Phase 4

Phase 4 adds open coding and code assignment:

- Open code creation, editing, and deletion
- Optional code descriptions and colors
- Open code usage counts
- Segment-code assignment and removal
- Colored code badges in the document reading view
- Segment highlights using the first assigned open code color
- Neutral highlights for uncoded segments
- Coded segment counts in the document list and dashboard
- Latest open codes and latest coded segments on the dashboard
- Audit log entries for code and assignment actions

This phase does not implement axial coding, category coding, full codebook editing, include/exclude rules UI, memo editing, CDA markers, actor maps, voice/silence reports, model building, export packages, AI features, authentication, or cloud functionality.

## Planned Later Phases

- Phase 5: Add memos and codebook
- Phase 6: Add grounded theory and CDA workspaces
- Later: Add axial coding, category coding, model building, and export packages
