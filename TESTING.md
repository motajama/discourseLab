# Testing discourseLab 0.1.0-beta

discourseLab 0.1.0-beta is a local testing release. Use non-sensitive test material first and keep backups of important data.

## Installation

```bash
cd disLab
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

Open:

```text
http://127.0.0.1:5000
```

On Windows, use `py -m venv .venv` and activate the virtual environment from PowerShell.

## Generic workflow test

1. Create a project.
2. Switch to the project.
3. Edit project settings.
4. Import a TXT document.
5. Import a DOCX document.
6. Open a document.
7. Select text and create a segment.
8. Confirm scroll position is preserved after segment actions.
9. Create an open code.
10. Assign the code to a segment.
11. Create a memo.
12. Export coded segments and memos.
13. Restart the app and confirm data persists.

## GT workflow test

1. Set or open a project that supports Grounded Theory.
2. Create an open code.
3. Create an axial code.
4. Create a category.
5. Assign the open code to the axial code.
6. Assign the axial code to the category.
7. Add GT notes if available.
8. Export `gt_hierarchy.md`.

## CDA workflow test

1. Set or open a project that supports CDA.
2. Create a CDA marker.
3. Edit the CDA marker.
4. Create an actor.
5. Edit the actor.
6. Assign the marker to a segment.
7. Assign the actor to a segment.
8. Create a discourse feature.
9. Inspect the voice/silence report.
10. Export CDA feature and voice/silence CSV files.

## Mixed workflow test

1. Create or open a mixed GT + CDA project.
2. Create segments.
3. Assign open codes.
4. Add CDA markers and actors to the same segments.
5. Create GT hierarchy items.
6. Create discourse features.
7. Export the codebook and methodology protocol.
8. Confirm GT and CDA materials both appear in exports.

## Model and network test

1. Create an analytical relation in the Model page.
2. Export model markdown and JSON.
3. Export visual model files if available.
4. Assign at least two codes, markers, actors, or features to one segment.
5. Open the Co-occurrence Network.
6. Confirm nodes and edges appear.
7. Click a node and an edge.
8. Export co-occurrence network JSON.

## Export and backup test

1. Export `codebook.md`.
2. Export `coded_segments.csv`.
3. Export `coded_segments.md`.
4. Export `memos.md`.
5. Export `gt_hierarchy.md`.
6. Export `cda_features.csv`.
7. Export `voice_silence.csv`.
8. Export `methodology_protocol.md`.
9. Export `analytical_model.md`.
10. Export `project.json`.
11. Export `package.zip`.
12. Export `project-backup.zip`.
13. Restart the app and confirm the project is still present.

## Integrity check test

1. Open `/admin/integrity`.
2. Run the check.
3. Confirm the page does not crash.
4. Record any warnings or orphaned-record reports in the release checklist.

## Bug report template

```text
Title:

Version: discourseLab 0.1.0-beta
Operating system:
Browser:
Python version:

What I was trying to do:

Steps to reproduce:
1.
2.
3.

What happened:

What I expected:

Error message or screenshot:

Does it still happen after restarting the app?

Does it involve sensitive research data? If yes, do not attach the data.
```

## Known limitations

- Restore from backup is not implemented yet.
- PDF import is not implemented.
- OCR is not implemented.
- Full-text search is not implemented yet.
- Inter-coder mode is not implemented.
- Cloud sync is not implemented.
- Authentication is not implemented.
- SVG and network visualizations are exploratory.
- DOCX import extracts text, not full formatting.
- The app is a local testing release.
