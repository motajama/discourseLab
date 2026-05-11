# discourseLab

Phase: 9 — Project modes and project management

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
3. Confirm `/health` returns JSON with `app = discourseLab`, `phase = 9`, `active_project_id`, and `methodology_mode`.
4. Go to Projects.
5. Create a new project and choose a methodology mode.
6. Confirm the new project becomes active.
7. Switch between projects with Open.
8. Confirm document, code, memo, GT, CDA, and export data is scoped to the active project.
9. Edit project metadata, research goal, principal investigator, and methodology mode.
10. Confirm changing methodology mode hides or shows GT/CDA workspaces without deleting GT or CDA data.
11. In Generic mode, confirm Documents, Codes, Memos, and Exports work while GT/CDA are disabled.
12. In GT mode, confirm GT works and CDA is disabled.
13. In CDA mode, confirm CDA works and GT is disabled.
14. In Mixed mode, confirm both GT and CDA work.
15. Soft-delete a project and confirm it disappears from Projects without being physically destroyed.
16. If the active project is deleted, confirm another project becomes active or a default project is created.
17. Confirm Phase 8 exports still download and the ZIP package follows the active methodology mode.
18. Restart the app and confirm existing SQLite data remains.

## Current Phase 9

Phase 9 adds project management and methodology modes:

- Project list, create, edit, open, and soft-delete routes
- Session-backed active project selection
- Safe project table migration for:
  - methodology_mode
  - status
  - last_opened_at
  - research_goal
  - principal_investigator
- Existing projects default to `methodology_mode = mixed`
- Dashboard project metadata and methodology mode
- Mode-aware sidebar, document view, codes page, exports page, and research package ZIP
- Disabled notices for GT/CDA workspaces when unsupported by the active project mode
- Health JSON includes active project and methodology mode

This phase does not implement advanced network model building, interactive graph editing, PNG/SVG/TikZ graph export, inter-coder mode, blind recoding mode, AI helpers, authentication, or cloud sync.

## Planned Later Phases

- Model building and relation visualization
- Visual graph exports
- Inter-coder workflows
