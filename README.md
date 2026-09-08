# CIDA Script Monitoring

Contains `ScriptImportStart.py` (shared helper functions, imported at the
start of automation scripts) and `ScriptMonitoringEnd.py` (logs script
run results to Asana and/or SQL Server, called at the end of automation
scripts).

## Setup

1. Copy `.env.example` to `.env`:
   ```
   cp .env.example .env
   ```
2. Fill in your own values in `.env`:
   - `CIDA_ASANA_ACCESS_TOKEN` — your personal Asana access token
   - `CIDA_SQL_USERNAME` / `CIDA_SQL_PASSWORD` — credentials for the
     `CIDA-SQL-T-01` / `cida_montioring` SQL Server database
3. `.env` is git-ignored and will never be pushed. Each person running
   this script needs their own `.env` with their own credentials.

If `python-dotenv` is installed (`pip install python-dotenv`), `.env` is
loaded automatically. Otherwise, set the variables as real environment
variables (e.g. via System Properties on Windows, or your shell profile).

## Notes

- Never commit real values to `.env.example`, `.gitignore`, or anywhere
  else in this repo — only placeholders.
- If a credential is ever accidentally committed, removing it in a later
  commit is not enough — it still exists in git history. Rotate/change
  the credential and consider scrubbing history.
