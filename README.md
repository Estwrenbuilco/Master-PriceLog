# CCI v4.5 — Codex continuation handoff

This is a **snapshot of the existing two-script project**, not a request to deploy or rebuild Google Sheets.

## Which version is live?
- **v4.5.2 billing**: last user-confirmed operational script. Email backfill, source links, and a 30-line rate-eligibility recheck are confirmed working.
- **v4.5.3 billing**: candidate code prepared to re-analyze supporting/held documents, not yet verified in live accounts.
- **v4.5 personal**: companion Google Sheets-bound script for workbook initialization, optional estimate, Buildertrend export. Do not run reset functions.

## Contents
- `billing/v4.5.2_LAST_CONFIRMED_Billing.gs`: working reference baseline.
- `billing/v4.5.3_CANDIDATE_Billing.gs`: proposed update; audit/diff and test before using.
- `personal/v4.5_Personal.gs`: existing companion project.
- `templates/v4.5_Lean_Master_SCHEMA.xlsx`: **reference workbook template only**, not live transaction data.
- `reference/Buildertrend_Cost_Codes.xlsx`: uploaded Buildertrend cost codes.
- `docs/v4.5.3_Document_Review_Guide.md`: proposed installation and tests; not a confirmed live success.
- `docs/v4.5.3_Test_Summary.md`: earlier mock-based test claims; independently verify.
- `HANDOFF_PROMPT.md`: complete task details; paste into Codex as initial instruction.
- `AGENTS.md`: persistent engineering and safety rules for Codex.

## Before working
1. Unzip this archive in a folder or private Git repository.
2. Open it as a project in Codex. Do not add real API keys or sensitive invoice emails to the repository.
3. Start Codex with the content of `HANDOFF_PROMPT.md`.
4. Have Codex audit the codebase, write tests, and propose a minimal change. No live deployment until the user explicitly approves.

## Known issues/open work
- Two previously reported AI analysis errors; **exact log messages were not provided yet**.
- Overclassification of priced documents as `SUPPORTING`, potentially hiding usable historical rates.
- Need strong line-by-line document comparison for duplicates/revisions without discarding legitimately different costs.
- Gmail search-offset historical pagination can miss messages as results shift; ledger remains essential for dedupe.
- Some Excel/CSV attachment pathways may be unsupported by model/file-input method; confirm in code and tests.
