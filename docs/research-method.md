# Research method: US lithium-ion battery pre-processors and refiners

Written before the first research pass (2026-11-03), so the rules don't bend to fit the results.

## Universe (what counts)
A **site in the US** (50 states, DC, territories) whose operator states that it carries out, or plans to carry out, at least one of:
- **pre-processing** of lithium-ion batteries: discharge, dismantling, or mechanical shredding/separation to black mass;
- **refining** of lithium-ion battery materials: hydrometallurgical, pyrometallurgical or direct recycling;
- **precursor production** from recycled material, only when on the same site as one of the above.

Any status counts, from `announced` (2018 or later) to `closed`: history matters for judging counterparties.

**Out of scope:** lead-acid recycling; pure collectors and logistics (unless they also pre-process); foreign sites of US companies; NiMH (recorded in notes only).

## Sources, in order of strength
1. The company: website, press releases, investor material.
2. Official records: SEC filings (EDGAR), US DOE announcements (grants, Loan Programs Office), state regulators (permits), EPA ECHO / RCRAInfo, court filings (bankruptcy).
3. Reputable trade and news press: used to find leads, and cited only when nothing primary exists (`source_type = news`).

## Rules
1. **Every row cites a source** that was actually opened, with its URL and the date retrieved. No source, no row.
2. **Unknown stays blank.** No guessed addresses, ZIPs, coordinates or capacities.
3. **Capacity exactly as stated,** with unit (t/yr, GWh/yr...), what it measures (batteries in, black mass in or out, metal out, product out) and basis (stated nameplate, planned, reported throughput). Never converted or averaged. When sources disagree, each figure gets its own row with its own source. A figure without unit, measure and basis goes into notes, not the capacity column.
4. **Status is always dated** (`status_as_of` = date of the newest source supporting it), and pauses, cancellations, bankruptcies, sales and closures are actively searched for (2023–2026).
5. **One row per site and process stage.** A site that shreds and refines has two rows.
6. **Desk research never confirms acceptance:** `confirmed_by_party` is always `no` until the company itself confirms.
7. No contacting companies, no paid databases, no scraping where a site's terms forbid it.

## How the first pass was run
- Three parallel research passes:
  - (A) the largest operators;
  - (B) mid-size and specialist operators;
  - (C) discovery of operators not on either list, plus any public lists of US facilities to cross-check completeness.
- Output in a fixed JSON format ([research/FORMAT.md](../research/FORMAT.md)), kept in `research/raw/`.
- `tools/import_research.py` loads the JSON into the workbook:
  - it assigns IDs and merges duplicates (same URL, same company name or website);
  - **it drops or blanks anything that breaks the rules, and lists every such case in the import report;**
  - the researchers' unresolved leads and conflicts also go in that report.
- `tools/check_workbook.py` must report zero errors before the seed is used.

## Completeness check
Public lists of US recycling facilities found during discovery are recorded as sources. Operators they name that aren't in the workbook are listed in the seed report, with the reason (out of scope, couldn't verify, or added).
