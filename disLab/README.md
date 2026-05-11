# discourseLab

Phase: 1 — Project foundation

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

## Verification

1. Open `http://127.0.0.1:5000`.
2. Confirm the header says `discourseLab`.
3. Confirm the dashboard shows `Demo Project`.
4. Confirm count cards appear for documents, codes, segments, memos, research questions, and relations.
5. Open `/health` and confirm the JSON response.
6. Restart the app and confirm existing database data is not deleted.

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
      documents.html
      codes.html
      memos.html
      gt_workspace.html
      cda_workspace.html
      exports.html
    static/
      style.css
      app.js
    uploads/
    exports/
```

## Current Phase 1

Phase 1 strengthens the application foundation before document import:

- Consistent `discourseLab` identity in the UI and documentation
- Reusable SQLite helper functions
- Safe first-start database initialization
- Active project loading with default project creation if needed
- Dashboard count cards for core analysis entities
- Project metadata panel
- Next steps panel
- Audit log helper and dashboard preview
- `/health` route for a simple local health check

This phase does not implement document upload, document parsing, segment creation, coding UI, AI features, authentication, or cloud functionality.

## Planned Later Phases

- Phase 2: Import TXT and DOCX documents
- Phase 3: Display documents and create text segments
- Phase 4: Add open coding and color highlighting
- Phase 5: Add memos and codebook
- Phase 6: Add grounded theory and CDA workspaces
- Later: Export codebooks, coded segments, memos, research models, Markdown, CSV, JSON, PNG/SVG, and LaTeX/TikZ
