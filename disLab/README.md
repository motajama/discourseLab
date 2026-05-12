# discourseLab

Phase: 12 — Methodology helper library and project protocol

discourseLab is a local lightweight qualitative analysis workspace for discourse analysis, critical discourse analysis, grounded theory, and generic qualitative coding.

Phase 12 adds a local methodology helper library and project-specific methodological protocol notes. The helper is not an AI chatbot and does not use network access.

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
2. Open a project.
3. Open Methodology.
4. Confirm relevant methodology libraries are shown based on project methodology mode.
5. Open Grounded Theory library if project mode is GT or mixed.
6. Open CDA libraries if project mode is CDA or mixed.
7. Create a protocol note.
8. Create a coding rule.
9. Create a reflexive note.
10. Link a methodology note to a code or segment.
11. Edit a methodology note.
12. Filter methodology notes.
13. Delete a methodology note.
14. Open GT, CDA, Model, and document pages and confirm methodology helper panels use library content.
15. Download methodology protocol Markdown.
16. Download complete research package ZIP and confirm `methodology_protocol.md` is included.
17. Download project JSON and confirm `methodology_notes` are included.
18. Restart the app and confirm methodology notes persist.

## Current Phase 12

Phase 12 adds:

- Methodology page
- Methodology library browsing
- Project-specific methodology notes
- Methodology note create/edit/delete/filter workflow
- Links from methodology notes to project entities
- Methodological protocol Markdown export
- Methodological protocol in package ZIP
- Methodology notes in project JSON
- Methodology-aware helper panels

This phase does not implement an AI methodology assistant, online bibliography lookup, Zotero integration, inter-coder mode, blind recoding mode, advanced protocol validation, cloud sync, authentication, or multi-user permissions.

