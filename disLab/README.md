# discourseLab

Phase: 7 — CDA workspace

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
3. Go to CDA Workspace.
4. Create a CDA marker, for example:
   - name: Metaphor
   - marker_type: metaphor
5. Create an actor, for example:
   - name: Refugees
   - actor_type: vulnerable_group
6. Open a document with at least one segment.
7. Assign the CDA marker to a segment.
8. Assign the actor to a segment with `relation_type = is_spoken_about`.
9. Add a discourse feature:
   - feature_type = metaphor
   - value = "wave"
   - interpretation = "Frames migration as natural force."
10. Confirm the segment card shows CDA marker, actor annotation, and feature.
11. Confirm the document text highlight uses CDA marker color when no open code is assigned.
12. Go to `/cda/features` and confirm the feature appears.
13. Go to `/cda/voice-silence` and confirm actor counts appear.
14. Confirm dashboard CDA counts update.
15. Confirm audit log shows CDA actions.
16. Restart the app and confirm CDA data persists.

## Current Phase 7

Phase 7 adds the first working Critical Discourse Analysis workspace:

- CDA marker creation and deletion
- Actor creation and deletion
- CDA marker assignment to existing segments
- Actor relation annotations on segments
- Discourse feature annotations for metaphor, modality, evaluation, presupposition, legitimation, intertextuality, framing, nominalization, passivization, agency, ideology, power relation, and other
- CDA feature overview with feature type and document filters
- Voice/silence actor report
- CDA prompts helper panel
- CDA counts and preview on the dashboard
- CDA information shown in document segment cards
- CDA marker fallback highlighting when a segment has no open code

This phase does not implement advanced network model building, relation graph editing, PNG/SVG/TikZ export, CDA CSV exports, full project JSON export, inter-coder mode, blind recoding mode, AI helpers, authentication, or cloud functionality.

## Planned Later Phases

- Model building and relation visualization
- Additional export formats
- Inter-coder workflows
