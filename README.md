# discourseLab

discourseLab is a local, lightweight qualitative analysis workspace for discourse analysis, critical discourse analysis, and grounded theory.

It runs on your own computer in a web browser. It does not require an account, cloud service, npm, Node, React, or external web services.

Current phase: **Phase 8 — Exports and research outputs**

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
git clone https://github.com/YOUR-USERNAME/discourseLab.git
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
2. Go to **Documents**.
3. Import a `.txt` or `.docx` file.
4. Open the document.
5. Select text and create a segment.
6. Create an open code and assign it to a segment.
7. Go to **CDA Workspace**.
8. Create a CDA marker, for example `Metaphor`.
9. Create an actor, for example `Refugees`.
10. Return to the document and assign CDA markers, actors, and discourse features to segments.
11. Open `/cda/features` to view discourse features.
12. Open `/cda/voice-silence` to view actor voice and silence counts.
13. Go to **Exports**.
14. Download the codebook, coded segments, memos, CDA reports, project summary, JSON archive, or complete ZIP package.

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
- Current phase: Phase 8 — Exports and research outputs
- Technology: Python, Flask, SQLite, HTML, CSS, vanilla JavaScript
- Runs locally in your browser
- No authentication
- No cloud service
- No AI features
