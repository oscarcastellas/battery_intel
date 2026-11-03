# CLAUDE.md

Guidance for Claude Code sessions in this repository.

## Project

A private intelligence and deal system for US lithium-ion (EV) battery resource recovery. Current scope: a complete map of US **pre-processors and refiners**, used for **feedstock placement** (brokering scrap, packs, modules, cells and black mass). Read `README.md`, `docs/plan.md`, `docs/research-method.md`, `docs/data-quality.md` and `docs/decisions.md` before proposing changes.

## Working agreements

- The owner has limited weekly time and is building SQL and Python skills. Keep work small, finished and runnable; prefer readable, explicit code with short comments.
- `templates/build_template.py` (`TABLES`, `LISTS`, `ID_PREFIX`) is the single source of truth for tabs, columns and dropdowns. `sql/0001_schema.sql` must match it (enforced by `tests/test_schema_matches_workbook.py`). Change both together and regenerate the template; never hand-edit the template's structure.
- Run `python -m unittest discover -s tests` after any change (standard library + openpyxl only).
- Desk research is loaded only via `tools/import_research.py` from JSON in `research/FORMAT.md`; keep the raw JSON in `research/raw/`.
- Free and public sources only for now. Check terms before automating access; no scraping where terms forbid it; no paid databases without asking.
- openpyxl gotcha: `ws.cell(row, col, value=None)` does not clear a cell; use `ws.cell(row, col).value = None`.

## Data rules (non-negotiable)

- Every fact row has a `source_id` (a real, opened source with URL and retrieval date). No source, no row.
- Blank means unknown. Never fill in likely values (addresses, ZIPs, coordinates, capacities).
- Capacity only with unit, measure and basis; never converted or averaged; conflicting figures are separate rows with their own sources.
- Status always has `status_as_of`; actively check for pauses, bankruptcies, sales and closures.
- Only the facility itself (quote, call, email) can make `confirmed_by_party = yes`. Desk research is always `no`.
- **No deal moves to `proposal_sent` or beyond unless `tools/check_workbook.py` reports a clean deal gate.**
- Business contacts only (work details), recorded with how they were obtained.
- Rows with IDs starting `EX-` are template examples. Anything synthetic is labelled `[SYNTH]` and never mixed with real data.
- `docs/regulatory-notes.md` lists open legal questions. Don't state them as facts until answered from official sources.
