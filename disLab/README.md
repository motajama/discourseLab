# discourseLab

Phase: 6 — Grounded Theory workspace

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
3. Create or confirm at least two open codes exist.
4. Go to GT Workspace.
5. Create an axial code.
6. Create a category.
7. Assign an open code to the axial code.
8. Assign the axial code to the category.
9. Open the code detail page and confirm hierarchy appears.
10. Edit GT fields for the axial code.
11. Edit GT fields for the category.
12. Go to `/gt/compare`.
13. Compare two open codes.
14. Create a comparison memo.
15. Confirm dashboard counts update.
16. Confirm codebook Markdown export includes open, axial, and category codes.
17. Restart the app and confirm GT hierarchy persists.

## Current Phase 6

Phase 6 adds the Grounded Theory workspace:

- Three-column Open Codes, Axial Codes, Categories board
- Axial code and category creation
- Open code to axial code assignment
- Axial code to category assignment
- GT-specific fields for axial and category codes
- Code detail hierarchy views
- Constant comparison screen for open codes
- Comparison memo creation
- GT structure preview on the dashboard
- Codebook Markdown export including hierarchy and GT fields

This phase does not implement CDA markers, actor maps, voice/silence reports, modality tracking, metaphor tracking, presupposition tracking, relation model building, PNG/SVG/TikZ model export, full coded segments CSV export, full project JSON export, AI features, authentication, or cloud functionality.

## Planned Later Phases

- CDA workspace behavior
- Model building and relation visualization
- Additional export formats
