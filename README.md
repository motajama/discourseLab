# discourseLab

discourseLab is a local, lightweight qualitative analysis workspace for discourse analysis, critical discourse analysis, grounded theory, and generic qualitative coding.

It runs on your own computer in a web browser. It does not require an account, cloud service, npm, Node, React, or a frontend build system. The UI attempts to load Libertinus from Google Fonts when online and falls back to local fonts when offline.

Version: **0.1.0-beta**

Current phase: **beta-prep**

Release label: **local testing release**

The analytical model builder connects project entities such as segments, codes, categories, actors, discourse features, memos, and research questions into explicit analytical relations.

The current UI pass adapts discourseLab to a flat Atari/TOS/GEM-inspired local workstation skin using Libertinus typography from Google Fonts. Document and segment source text remains monospace for careful reading and coding. It does not add analytical features, schema changes, AI, cloud sync, authentication, Node, npm, React, Vue, an external CSS framework, or a frontend build system.

For release testing, read:

- [TESTING.md](TESTING.md)
- [RELEASE_CHECKLIST.md](RELEASE_CHECKLIST.md)
- [CHANGELOG.md](CHANGELOG.md)
- [RELEASE_NOTES.md](RELEASE_NOTES.md)

Keep backups of important data. Use the project backup ZIP regularly from the Exports page. Restore from backup is not implemented yet.

## Atari UI Verification

1. Start the app with `python app.py`.
2. Confirm the fixed top navigation uses flat Atari-style bordered buttons and the active item is clear.
3. Confirm dashboards, cards, forms, tables, exports, model, methodology, and network pages use strong rectangular borders and hard shadows.
4. Confirm the Google Fonts Libertinus link is present in the base template.
5. Confirm main UI text uses Libertinus with local fallbacks.
6. Confirm document text, segment text, selected text previews, and coded excerpts use monospace.
7. Confirm document coding still supports text selection, segment creation, assignment, and scroll preservation.
8. Confirm delete confirmations still appear.
9. Confirm no external CSS framework or frontend build system was introduced.
10. Confirm `/health` returns `"version": "0.1.0-beta"` and `"phase": "beta-prep"`.

## What You Need

- A computer running Linux, macOS, or Windows
- Python 3.10 or newer
- A web browser such as Firefox, Chrome, Edge, or Safari

## 1. Install Python

### Linux

Most Linux systems already include Python. Check with:

```bash
python3 --version
```

If Python is missing, install it with your package manager.

Ubuntu or Debian:

```bash
sudo apt update
sudo apt install python3 python3-venv python3-pip
```

Fedora:

```bash
sudo dnf install python3 python3-pip
```

### macOS

Check whether Python is installed:

```bash
python3 --version
```

If Python is missing, install it from:

```text
https://www.python.org/downloads/
```

You can also install it with Homebrew:

```bash
brew install python
```

### Windows

Install Python from:

```text
https://www.python.org/downloads/windows/
```

During installation, check:

```text
Add python.exe to PATH
```

Then open PowerShell and check:

```powershell
py --version
```

## 2. Get discourseLab

If you use Git:

```bash
git clone https://github.com/motajama/discourseLab.git
cd discourseLab
```

If you downloaded a ZIP file:

1. Unzip the file.
2. Open a terminal or PowerShell window.
3. Move into the unzipped `discourseLab` folder.

Example:

```bash
cd Downloads/discourseLab
```

## 3. Create a Virtual Environment

A virtual environment keeps discourseLab's Python packages separate from the rest of your computer.

### Linux and macOS

```bash
cd disLab
python3 -m venv .venv
source .venv/bin/activate
```

### Windows PowerShell

```powershell
cd disLab
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, run:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Then try again:

```powershell
.\.venv\Scripts\Activate.ps1
```

### Windows Command Prompt

```cmd
cd disLab
py -m venv .venv
.\.venv\Scripts\activate.bat
```

## 4. Install Required Packages

Run this from inside the `disLab` folder, with the virtual environment activated.

### Linux and macOS

```bash
pip install -r requirements.txt
```

### Windows

```powershell
pip install -r requirements.txt
```

## 5. Start discourseLab

Run this from inside the `disLab` folder.

### Linux and macOS

```bash
python app.py
```

If your system uses `python3` instead of `python`, run:

```bash
python3 app.py
```

### Windows

```powershell
py app.py
```

You should see a message similar to:

```text
Running on http://127.0.0.1:5000
```

## 6. Open the App

Open your web browser and go to:

```text
http://127.0.0.1:5000
```

The app runs locally on your computer. Keep the terminal or PowerShell window open while using discourseLab.

## 7. Stop the App

In the terminal or PowerShell window where discourseLab is running, press:

```text
Ctrl + C
```

## 8. Start It Again Later

### Linux and macOS

```bash
cd discourseLab/disLab
source .venv/bin/activate
python app.py
```

### Windows PowerShell

```powershell
cd discourseLab\disLab
.\.venv\Scripts\Activate.ps1
py app.py
```

### Windows Command Prompt

```cmd
cd discourseLab\disLab
.\.venv\Scripts\activate.bat
py app.py
```

## 9. First Things To Try

1. Open `http://127.0.0.1:5000`.
2. Go to **Projects**.
3. Create a new project and choose a methodology mode: Generic, Grounded Theory, CDA, or Mixed.
4. Open the project and confirm the sidebar changes for the selected methodology mode.
5. Go to **Documents**.
6. Import a `.txt` or `.docx` file.
7. Open the document.
8. Select text and create a segment.
9. Create an open code and assign it to a segment.
10. If the project supports CDA, create CDA markers, actors, and discourse features.
11. If the project supports GT, use the GT workspace for axial codes, categories, and comparison.
12. Ensure you have at least one segment, one code, one memo, and optionally one actor or discourse feature.
13. Open **Model**.
14. Create a relation between a segment and a code.
15. Create a relation between an actor and a code or discourse feature if CDA data exists.
16. Create a relation between an open, axial, or category code if GT data exists.
17. Add memo and evidence notes to the relation.
18. Edit the relation.
19. Open the entity-centered view for one related entity.
20. Filter relations by type and strength.
21. Create a research question from the Model page.
22. Link a relation to the research question.
23. Download `/exports/model.md`.
24. Download `/exports/model.json`.
25. Open **Methodology**.
26. Confirm relevant methodology libraries are shown based on the project methodology mode.
27. Open Grounded Theory library if the project mode is GT or mixed.
28. Open CDA libraries if the project mode is CDA or mixed.
29. Create a protocol note, a coding rule, and a reflexive note.
30. Link a methodology note to a code or segment.
31. Edit and filter methodology notes.
32. Delete a methodology note.
33. Open GT, CDA, Model, and document pages and confirm methodology helper panels are compact or collapsed and use library content.
34. Open **Network** and confirm empirical co-occurrence between segment assignments is shown.
35. Download methodology protocol Markdown.
36. Download the complete research package ZIP and confirm `methodology_protocol.md` and `cooccurrence_network.json` are included.
37. Download project JSON and confirm `methodology_notes` are included.
38. Restart the app and confirm methodology notes persist.

## Methodology Helper Library

The methodology helper library is local JSON guidance stored in `disLab/methodology`. It does not call AI services, does not require internet access, and does not look up bibliography online.

It provides:

- methodology-aware guidance for generic qualitative coding, Grounded Theory, CDA, and mixed GT/CDA work
- structured phases, concepts, prompts, and recommended actions
- APA-style methodological sources
- project-specific protocol notes and decision logs
- helper panels on GT, CDA, Model, and document pages
- Markdown export of the project methodological protocol

## Visual Model Exports

Model modes:

- `simplified`: default readable mode; hides weak and uncertain relations unless requested.
- `argument`: focuses on research questions, categories, axial codes, memos, strong relations, and argument-building relation types.
- `evidence`: focuses on documents, segments, memos, evidence, examples, negative cases, support, and contradiction.
- `gt`: focuses on code-only grounded theory relations.
- `cda`: focuses on actors, discourse markers, discourse features, and CDA relation types.
- `full`: shows all relations and may be visually dense.

Visual encodings:

- Node colors show entity type.
- Edge color shows relation family.
- Edge thickness and dash pattern show relation strength.
- Edge labels use shortened relation names.
- SVG tooltips include full node and edge metadata.

- Mermaid `.mmd`: paste into Mermaid-compatible Markdown tools, GitHub, Obsidian, or Mermaid editors.
- Graphviz `.dot`: render with Graphviz:

```bash
dot -Tpng discourseLab_analytical_model.dot -o model.png
```

- TikZ `.tikz`: a snippet to include inside a LaTeX document with TikZ enabled.
- SVG `.svg`: open directly in a browser or vector editor.

## Co-occurrence Network

The Co-occurrence Network is generated from segment assignments. It connects codes, CDA markers, actors, and discourse features when they appear in the same segment.

This network shows empirical co-presence. It is not the same as the manually curated Analytical Model, and it does not imply causality or replace interpretation. Use it to explore which coding and annotation items cluster together, then use the Model page to create explicit analytical relations.

Available outputs:

- `/network`: interactive in-app SVG network explorer.
- `/network/data`: JSON data endpoint for the current active project.
- `/exports/cooccurrence-network.json`: downloadable graph JSON.
- `/exports/cooccurrence-edges.csv`: downloadable edge list.
- Complete research package ZIP includes `cooccurrence_network.json` and `cooccurrence_edges.csv`.

## Atari UI Skin With Libertinus Typography

discourseLab uses a flat Atari ST / TOS / GEM-inspired local workstation style: gray workspace background, rectangular panels, strong dark borders, hard pixel-like shadows, compact controls, and no external CSS framework.

The main UI uses Libertinus from Google Fonts. If offline, the browser falls back to local serif and sans-serif fonts. Document and segment text remains monospace so source excerpts are easy to inspect and code.

## Phase 12.5d Notes

- The visual skin uses only local CSS: light gray workspace, crisp rectangular panels, thin dark borders, flat compact controls, and restrained blue accents.
- Tooltips use local HTML/CSS with `title` fallback and keyboard focus support.
- Explanations cover methodology modes, CDA marker types, actor types, actor relation types, discourse feature types, GT terms, model relation types, and relation strengths.
- Codebook exports now include `## CDA Markers` and `## Actors`; the research package ZIP includes the updated codebook automatically.

## Data Storage

discourseLab stores its local SQLite database here:

```text
disLab/instance/disLab.sqlite
```

Imported source files are stored here:

```text
disLab/uploads/
```

Do not delete these files unless you intentionally want to remove your local project data.

## 0.1.0-beta Testing Notes

This is a local testing release. Use non-sensitive test material first, export the complete research package when testing major workflows, and create project backup ZIP files regularly.

Known limitations:

- Restore from backup is not implemented yet.
- PDF import and OCR are not implemented.
- Full-text search is not implemented yet.
- Inter-coder mode is not implemented.
- Cloud sync and authentication are not implemented.
- DOCX import extracts text, not full formatting.
- SVG and network visualizations are exploratory.

## Troubleshooting

### `python: command not found`

Try:

```bash
python3 --version
```

On Windows, try:

```powershell
py --version
```

### `No module named flask` or `No module named docx`

Activate the virtual environment and reinstall requirements:

```bash
pip install -r requirements.txt
```

### The browser cannot open the app

Make sure the terminal still shows that discourseLab is running. Then open:

```text
http://127.0.0.1:5000
```

### Port 5000 is already in use

Another app may already be using port 5000. Stop the other app, or run discourseLab on another port:

```bash
flask --app app run --port 5001
```

Then open:

```text
http://127.0.0.1:5001
```

## Project Notes

- App name: discourseLab
- Version: 0.1.0-beta
- Current phase: beta-prep
- Release label: local testing release
- Technology: Python, Flask, SQLite, HTML, CSS, vanilla JavaScript
- Runs locally in your browser
- No authentication
- No cloud service
- No AI features

## Credits

- Jan Motal: project author, concept, research workflow, and product direction.
- OpenAI Codex: implementation assistance and coding support.
