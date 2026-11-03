# Plan

**Goal:** a complete, accurate map of US lithium-ion battery pre-processors and refiners, then the first feedstock placement deal.

| Step | Task | Done when |
|---|---|---|
| **1** (done) | Framework: workbook, checker, importer, tests, schema; research method written | Tests pass |
| **2** (done) | First desk-research pass → `research/battery_intel_seed_2026-11.xlsx` + report | Checker 0 errors; report lists gaps and conflicts |
| **3** (done) | First spot-check: 10 records re-verified, 0 wrong ([research/audit_2026-11-03.md](../research/audit_2026-11-03.md)) | Audit logged |
| **4** | **Regulatory checklist** (`docs/regulatory-notes.md`), starting with black mass export controls (item 0) | Every item answered with a source |
| **5** | **Call the top 10 sites** listed in [research/seed_report.md](../research/seed_report.md): which feed forms and chemistries, conditions, minimum volumes, how they price (payables, fees). Record contacts, interactions, `confirmed_by_party`, prices; update stale statuses | 10 sites confirmed; first price observations |
| **5b** | **Completeness check** against the NREL/NAATBatt supply-chain database and Argonne's facility map; resolve the leads in the seed report | Differences documented |
| **6** | **Map feedstock holders** (cell and pack plants' manufacturing scrap, OEM service and dealer networks, fleets, ESS operators) as organizations and sites | First list of holders by region |
| **7** | **First feed lot → deal**, through the deal gate | One deal at proposal_sent with a clean gate |

Work in your own copy of the seed workbook (e.g. `data/battery_intel.xlsx`, which git ignores) and run the checker after every session.
