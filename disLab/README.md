# discourseLab

Phase: 12.5a — Safety and workflow hotfixes

discourseLab is a local lightweight qualitative analysis workspace for discourse analysis, critical discourse analysis, grounded theory, and generic qualitative coding.

Phase 12.5a keeps the local Flask/SQLite app structure and adds small safety and workflow fixes: CDA marker editing, actor editing, delete confirmations, document-view scroll preservation after coding actions, a non-overlapping bottom segment selection bar, and a basic project backup ZIP export.

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

## Methodology Helper Library

The methodology helper library is stored locally in JSON files under `methodology/`.

It provides:

- local JSON-based methodology guidance
- no AI and no internet access
- APA-style methodological sources
- project protocol notes
- methodology-aware helper panels
- Markdown export of the methodological protocol

Available libraries:

- generic qualitative coding
- grounded theory
- Fairclough-oriented CDA
- van Dijk-oriented CDA
- Wodak/DHA-oriented CDA
- mixed GT/CDA protocol guidance

## Verification

1. Start the app with `python app.py`.
2. Edit a CDA marker.
3. Edit an actor.
4. Try deleting a segment, document, or code and confirm the warning appears.
5. Open a long document, scroll down, create a segment, and confirm scroll position is preserved.
6. Confirm the segment creation UI appears as a bottom bar and does not cover selected text.
7. Download `/exports/project-backup.zip`.
8. Restart the app and confirm existing data is still present.

## Current Phase 12.5a

Phase 12.5a adds:

- Editable CDA markers
- Editable actors
- Confirmation dialogs for destructive forms
- Scroll preservation on the document detail page after coding and annotation actions
- Fixed bottom segment selection bar replacing the floating popup
- Basic project backup ZIP export
- `/health` reports `phase: "12.5a"`

This phase does not implement the Atari/TOS/GEM skin, dashboard redesign, navigation redesign, AI, cloud sync, authentication, React, Vue, Node, npm, or external CDN usage.
