# discourseLab

discourseLab is a local lightweight CAQDAS application skeleton for qualitative text analysis, with a planned focus on discourse analysis, critical discourse analysis, and grounded theory.

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

## Project Structure

```text
discourseLab/
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

## Current Phase 0

Phase 0 creates the basic local technical skeleton only:

- Python and Flask backend
- SQLite database
- Plain HTML, CSS, and vanilla JavaScript frontend
- Local browser interface with top bar, sidebar, dashboard, and placeholder workspaces
- First-start database creation from `schema.sql`
- One default project named `Demo Project`

This phase does not implement document import, segment highlighting, coding workflows, memos, AI features, exports, authentication, or cloud functionality.

## Planned Later Phases

- TXT and DOCX document import
- Color-based segment highlighting
- Open coding
- Axial coding
- Category coding
- Memo writing and entity-linked memos
- CDA helper workspace
- Grounded theory workspace
- Codebook, segment, memo, and report exports
