# Excel Mock Interviewer — PoC (Rule-based)

This repository contains a runnable Proof-of-Concept of an Excel Mock Interviewer implemented as a Streamlit app.

## Contents
- `app.py` — main Streamlit app (runs locally).
- `DESIGN.md` — design document & approach.
- `samples/` — example transcript JSON files.

## Requirements
- Python 3.9+
- pip

## Install & run locally
```bash
python -m venv venv
# macOS / Linux
source venv/bin/activate
# Windows (PowerShell)
# .\venv\Scripts\Activate.ps1

pip install -r requirements.txt
# or:
pip install streamlit pandas openpyxl python-dateutil

streamlit run app.py
