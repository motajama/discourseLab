# discourseLab

Phase: 12.5d - Atari skin, methodology tooltips, and codebook additions

discourseLab is a local lightweight qualitative analysis workspace for discourse analysis, critical discourse analysis, grounded theory, and generic qualitative coding.

Phase 12.5d keeps the local Flask/SQLite app structure and adds a compact Atari ST / TOS / GEM-inspired skin, local methodology tooltips, compact helper panels, and CDA marker/actor sections in the codebook export.

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
7. Confirm marker type, actor type, feature type, methodology mode, and relation strength tooltips appear.
8. Download `/exports/codebook.md` and confirm CDA markers and actors are included.
9. Download `/exports/package.zip` and confirm its codebook includes the same sections.
10. Confirm `/health` reports `phase: "12.5d"`.
11. Restart the app and confirm existing data is still present.

## Current Phase 12.5d

Phase 12.5d adds:

- Flat Atari/TOS/GEM-inspired local CSS skin
- Local hover/focus methodology tooltips
- Compact or collapsed methodology helper panels
- Explanations for methodology modes, CDA markers, actors, feature types, GT terms, relation types, and strengths
- CDA marker and actor sections in codebook exports
- `/health` reports `phase: "12.5d"`

This phase does not implement the Atari/TOS/GEM skin, dashboard redesign, navigation redesign, AI, cloud sync, authentication, React, Vue, Node, npm, or external CDN usage.
