# discourseLab

Version: 0.1.0-beta

Phase: beta-prep

Release label: local testing release

discourseLab is a local lightweight qualitative analysis workspace for discourse analysis, critical discourse analysis, grounded theory, and generic qualitative coding.

The current UI pass keeps the local Flask/SQLite app structure and adapts discourseLab to a flat Atari/TOS/GEM-inspired workstation skin using Libertinus typography from Google Fonts. Document and segment source text remains monospace.

For release testing from the repository root, see `TESTING.md`, `RELEASE_CHECKLIST.md`, `CHANGELOG.md`, and `RELEASE_NOTES.md`.

Keep backups of important data. Use the project backup ZIP regularly. Restore from backup is not implemented yet.

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
2. Confirm the fixed top navigation uses flat bordered buttons and the active item is clear.
3. Confirm dashboard, document, exports, methodology, model, and network pages use rectangular Atari-style panels.
4. Confirm document text, segment text, selected text previews, and coded excerpts use monospace.
5. Confirm document coding still supports text selection, segment creation, assignment, and scroll preservation.
6. Confirm delete confirmations still appear.
7. Confirm `/health` reports `version: "0.1.0-beta"` and `phase: "beta-prep"`.
8. Restart the app and confirm existing data is still present.

## Current UI Skin

This pass adds:

- Libertinus Serif and Libertinus Sans from Google Fonts in the base template
- CSS font variables for body, headings, UI controls, and monospace excerpts
- Atari/TOS/GEM-style color, border, and hard-shadow tokens
- Consistent styling for navigation, panels, cards, buttons, forms, tables, badges, tooltips, document coding, exports, model, and network pages

If offline, the browser falls back to local serif and sans-serif fonts. No font files are bundled in the repository.

This phase does not implement analytical features, schema changes, AI suggestions, cloud sync, authentication, React, Vue, Node, npm, or a frontend build system.
