# Email Migration Script

This repository contains a script that migrates Outlook emails to Asana tasks.
It was originally designed to run within GitHub Actions but can also be run
locally.

## Local Setup

1. Install Python 3.10 or later.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Copy `.env.example` to `.env` and fill in the required values. Alternatively
   you can specify a custom file via the `ENV_FILE` environment variable.
4. Run the script:
   ```bash
   python main.py
   ```

`asana_outlook_integration_script.py` automatically loads environment variables
from `.env` (or the path specified by `ENV_FILE`).
