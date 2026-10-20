# Engineering guardrails for this repository

- Respect the user's stated priorities in `HANDOFF_PROMPT.md` and read `README.md` first.
- This is an **existing live prototype**, not a greenfield rebuild. No sheet resets, no destructive data edits, no changing triggers or production Script Properties without approval.
- Treat `billing/v4.5.2_LAST_CONFIRMED_Billing.gs` as the last user-verified functional baseline; `billing/v4.5.3_CANDIDATE_Billing.gs` is unverified proposed code.
- Maintain exactly two Apps Script project roles: billing reads Gmail and calls OpenAI; personal hosts Estimate and Buildertrend export. Same existing spreadsheet ID. Preserve the 12-tab schema where possible.
- Keep Google Sheets `Documents`, `Cost Lines`, `Price History`, `Source Ledger`, `Projects`, `Label Rules`, `Review Queue`, and source URLs compatible unless a separately approved migration is essential.
- 75% similarity is a candidate correlation threshold, **not** a duplicate verdict. Exact financial identity must be verified before suppressing or promoting charges.
- Capture every useful price observation, but distinguish informational/reference pricing from booked actual expenses. Do not double book invoices/receipts/credits.
- Protect API credentials and personally identifying Gmail content. Never hardcode keys or upload real inbox files to tests.
- Add deterministic tests for modified behavior. Report test output honestly: offline mocks do not validate live Gmail/Google Sheets/OpenAI integration.
- Keep changes small and traceable; preserve source links and idempotent processing; do not overwrite an entire row just to update a status.
