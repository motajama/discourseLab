# discourseLab

Phase: 11.5 — Readable visual model exports

discourseLab is a local lightweight qualitative analysis workspace for discourse analysis, critical discourse analysis, grounded theory, and generic qualitative coding.

Phase 11.5 improves generated visual model exports so analytical relations are readable instead of exported as one undifferentiated graph.

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

## Visual Model Exports

Visual exports are scoped to the active project and generated from saved analytical relations. They do not modify the database.

Model modes:

- `full`: all relations, including weak and uncertain, capped at 100 by default.
- `simplified`: default readable mode; hides weak and uncertain relations by default.
- `argument`: research questions, categories, axial codes, memos, strong relations, and argument-building relation types.
- `evidence`: documents, segments, memos, evidence, examples, negative cases, support, and contradiction.
- `gt`: code-only grounded theory relations.
- `cda`: actors, discourse markers, discourse features, and CDA relation types.

Visual encodings:

- Node colors = entity type.
- Edge color = relation family.
- Edge thickness/dash = relation strength.
- Edge labels = shortened relation type.
- SVG tooltips = full metadata for nodes and edges.

Formats:

- Mermaid `.mmd`
- Graphviz `.dot`
- TikZ `.tikz`
- SVG `.svg`

Graphviz example:

```bash
dot -Tpng discourseLab_analytical_model.dot -o model.png
```

TikZ exports are snippets to include inside a LaTeX document with TikZ enabled.

## Verification

1. Start the app with `python app.py`.
2. Open a project with several analytical relations.
3. Go to Model.
4. Download simplified SVG.
5. Confirm it has:
   - legend
   - different node colors
   - different edge styles by strength
   - different edge colors by relation family
   - readable wrapped node labels
6. Download argument SVG.
7. Confirm it shows fewer and more analytically relevant relations.
8. Download full SVG.
9. Confirm it still works even if dense.
10. Download GT mode and CDA mode visual exports.
11. Confirm filters do not crash with empty result sets.
12. Confirm package ZIP includes visual exports by model mode.
13. Confirm no project data is modified by exports.

## Current Phase 11.5

Phase 11.5 adds:

- model modes for visual exports
- relation strength visual encoding
- relation family visual encoding
- node type styling
- layered SVG layout
- SVG legend
- SVG node and edge tooltips
- wrapped SVG node labels
- mode-specific Mermaid, DOT, TikZ, and SVG files in the research package ZIP

This phase does not implement a drag-and-drop graph editor, browser-based interactive network editor, heavy graph layout libraries, AI helpers, inter-coder mode, cloud sync, authentication, or multi-user permissions.

