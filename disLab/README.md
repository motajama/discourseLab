# discourseLab

Phase: 2 — Document import

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

1. Open `http://127.0.0.1:5000`.
2. Confirm the header says `discourseLab`.
3. Go to Documents.
4. Upload a `.txt` file.
5. Upload a `.docx` file.
6. Confirm both appear in the document list.
7. Open each document and confirm extracted text is visible.
8. Confirm the dashboard document count increased.
9. Confirm the audit log shows import actions.
10. Delete a document and confirm it disappears.
11. Restart the app and confirm remaining documents are still present.

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

## Current Phase 2

Phase 2 adds document import and document management:

- TXT import with UTF-8 and latin-1 fallback
- DOCX import with paragraph text extraction through `python-docx`
- Safe local upload filenames
- 16 MB upload size limit
- Document list with text length and segment count
- Document detail view with escaped extracted plain text
- Document deletion through POST
- Audit log entries for document import and deletion
- Dashboard document count and latest documents panel
- `/health` route for a simple local health check

Document deletion removes database rows for the document and any associated future segments or segment-code links. The original uploaded source file is intentionally left in `uploads/` for now.

This phase does not implement segment selection, text coding, color highlighting, memo authoring, CDA tools, GT tools, model building, exports, AI features, authentication, or cloud functionality.

## Planned Later Phases

- Phase 3: Display documents and create text segments
- Phase 4: Add open coding and color highlighting
- Phase 5: Add memos and codebook
- Phase 6: Add grounded theory and CDA workspaces
- Later: Export codebooks, coded segments, memos, research models, Markdown, CSV, JSON, PNG/SVG, and LaTeX/TikZ
