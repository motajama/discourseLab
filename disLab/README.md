# discourseLab

Phase: 8 — Exports and research outputs

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

1. Start the app with `python app.py`.
2. Open `http://127.0.0.1:5000`.
3. Go to Exports.
4. Download Codebook Markdown.
5. Download Coded Segments CSV.
6. Download Coded Segments Markdown.
7. Download Memos Markdown.
8. Download GT Hierarchy Markdown.
9. Download CDA Features CSV.
10. Download Voice/Silence CSV.
11. Download Project Summary Markdown.
12. Download Full Project JSON.
13. Download Complete Research Package ZIP.
14. Open the ZIP and confirm it contains all expected files.
15. Confirm exports only include the active project.
16. Restart the app and confirm data is unchanged.

## Current Phase 8

Phase 8 adds robust research exports:

- Codebook Markdown
- Coded segments CSV
- Coded segments Markdown
- Memos Markdown
- GT hierarchy Markdown
- CDA features CSV
- Actor voice/silence CSV
- Project summary Markdown
- Full active-project JSON
- Complete research package ZIP generated in memory

Exports are generated on demand, use UTF-8, use active project data only, and do not modify the database.

This phase does not implement advanced network model building, interactive graph editing, PNG/SVG/TikZ graph export, inter-coder mode, blind recoding mode, AI helpers, authentication, cloud sync, or source-document bundling in the ZIP package.

## Planned Later Phases

- Model building and relation visualization
- Visual graph exports
- Inter-coder workflows
