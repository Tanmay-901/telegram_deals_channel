# clasp workflow

This repo is configured for Apps Script deployment with **clasp**.

## One-time setup

1. Install clasp:
   - `npm i -g @google/clasp`

2. Authenticate:
   - `clasp login`

3. Initialize local config mapping (optional if already correct):
   - `clasp clone <scriptId> .` 

## Common commands

- Check status:
  - `clasp status`

- Push code changes:
  - `clasp push`

## Web App deployment note

If you rely on the Apps Script UI to update the Web App version, do:

1. Edit `MAIN_telegram_deals_channel.gs`.
2. Run `clasp push`.
3. Go to Apps Script UI → Deploy → Manage deployments → Update the Web App version.
