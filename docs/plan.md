# Plan

**Goal:** a complete, accurate map of US lithium-ion battery pre-processors and refiners, then the first feedstock placement deal.

| Step | Task | Done when |
|---|---|---|
| **1** (done) | Framework: workbook, checker, importer, tests, schema; research method written | Tests pass |
| **2** (done) | First desk-research pass → `research/battery_intel_seed_2026-11.xlsx` + report | Checker 0 errors; report lists gaps and conflicts |
| **3** | **Review the seed:** spot-check 10 records against their sources (`--audit-sample 10`); fix anything wrong; work through the "unresolved" list | Audit logged; error rate known |
| **4** | **Regulatory checklist** (`docs/regulatory-notes.md`) | Every item answered with a source |
| **5** | **Call the top sites** from the seed report (operating, taking third-party feed): which feed forms and chemistries, conditions, minimum volumes, how they price (payables, fees). Record contacts, interactions, `confirmed_by_party`, prices | 10 sites confirmed; first price observations |
| **6** | **Map feedstock holders** (cell and pack plants' manufacturing scrap, OEM service and dealer networks, fleets, ESS operators) as organizations and sites | First list of holders by region |
| **7** | **First feed lot → deal**, through the deal gate | One deal at proposal_sent with a clean gate |

Work in your own copy of the seed workbook (e.g. `data/battery_intel.xlsx`, which git ignores) and run the checker after every session.
