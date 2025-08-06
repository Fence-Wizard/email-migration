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

## Email Analytics

`email_analytics.py` retrieves messages from an Outlook Inbox using the Microsoft Graph API and performs basic analytics such as sender counts, sentiment analysis, and topic modeling.

1. Ensure the dependencies in `requirements.txt` are installed.
2. Populate `.env` (or `config.toml`) with your Azure AD details. Add `AZ_AUTH_MODE`
   or `auth_mode` to choose between:
   - `app` &mdash; client credential flow using `AZ_CLIENT_SECRET`.
   - `delegated` &mdash; username/password flow using `AZ_USERNAME` and `AZ_PASSWORD`.
   At minimum the script expects `AZ_CLIENT_ID` and `AZ_TENANT_ID`, plus the
   credentials required for the selected mode.
3. Run the script:
   ```bash
   python email_analytics.py
   ```

The script saves a summary CSV (`tmyers_inbox_summary.csv`) for further analysis.
