# Deploying the CAQM Reconciliation App

This folder is ready to push to GitHub and deploy on Streamlit Community Cloud.

## What each file is for

| File | Purpose |
|---|---|
| `requirements.txt` | Python packages Streamlit Cloud installs automatically (streamlit, openpyxl). |
| `runtime.txt` | Pins the Python version (3.11) so the deploy is reproducible. |
| `packages.txt` | System-level packages (empty here, but Streamlit looks for it). |
| `.streamlit/config.toml` | Raises the file-upload limit to 50 MB and sets a light theme. |
| `.gitignore` | Keeps generated output and caches out of the repo. |
| `LICENSE` | MIT licence (change if your organisation needs a different one). |

## One-time push to GitHub (from VS Code terminal, inside this folder)

```bash
git init
git add .
git commit -m "CAQM reconciliation app"
git remote add origin https://github.com/YOUR_USERNAME/caqm-reconciliation.git
git branch -M main
git push -u origin main
```

If Git asks for a password and rejects it: GitHub needs a Personal Access
Token. GitHub -> Settings -> Developer settings -> Personal access tokens ->
Tokens (classic) -> Generate new token -> tick `repo` -> use the token as the
password.

## Deploy live (free)

1. Go to https://share.streamlit.io and sign in with GitHub.
2. New app -> Repository: `YOUR_USERNAME/caqm-reconciliation`, Branch: `main`,
   Main file path: `app.py`.
3. Deploy. You get a public URL like `https://....streamlit.app`.

Every `git push` after that redeploys the live app automatically.

## Important

- The `templates/` folder MUST be committed (Module 3 needs it). After pushing,
  confirm the four .xlsx files show under `templates/` on GitHub.
- Uploaded workbooks are processed in-session only; nothing is written back to
  the repo. Keep the repo Private if the data is sensitive.
