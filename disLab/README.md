# discourseLab

Phase: 10 — Relations and analytical model builder

discourseLab is a local lightweight qualitative analysis workspace for discourse analysis, critical discourse analysis, grounded theory, and generic qualitative coding.

The analytical model builder connects project entities such as segments, codes, categories, actors, discourse features, memos, and research questions into explicit analytical relations.

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
3. Ensure you have at least:
   - one segment,
   - one code,
   - one memo,
   - optionally one actor or discourse feature.
4. Open Model.
5. Create a relation between a segment and a code.
6. Create a relation between an actor and a code or discourse feature if CDA data exists.
7. Create a relation between an open, axial, or category code if GT data exists.
8. Add memo and evidence note to the relation.
9. Edit the relation.
10. Open the entity-centered view for one related entity.
11. Filter relations by type and strength.
12. Create a research question from the Model page.
13. Link a relation to the research question.
14. Download `/exports/model.md`.
15. Download `/exports/model.json`.
16. Download the complete research package ZIP and confirm `analytical_model.md` and `analytical_model.json` are included.
17. Restart the app and confirm relations persist.
18. Confirm data remains scoped to the active project.

## Current Phase 10

Phase 10 adds relations and analytical model building:

- Relation creation, editing, deletion, filtering, and browsing
- Relations between documents, segments, codes, memos, research questions, CDA markers, actors, and discourse features
- Relation title, type, strength, analytical memo, and evidence note
- Entity-centered relation maps
- Minimal research question creation from the Model page
- Analytical model Markdown export
- Analytical model JSON export
- Analytical model files included in the complete research package ZIP
- Compact relation integrations on code detail, document segment cards, memos, and CDA actor lists

This phase does not implement graphical network visualization, drag-and-drop graph editing, PNG/SVG/TikZ/Mermaid/DOT export, inter-coder mode, blind recoding mode, AI helpers, authentication, cloud sync, or multi-user permissions.

## Planned Later Phases

- Graphical model visualization
- Visual graph exports
- Inter-coder workflows
