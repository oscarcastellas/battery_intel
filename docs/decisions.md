# Decision log

Newest first. To change a decision, add a new entry that supersedes it.

## B-004 · Desk research is loaded only through the importer (2026-11-03)
Research arrives as JSON in a fixed format and is loaded by `tools/import_research.py`, which never guesses: invalid values are blanked, unsourced rows are dropped, capacity without unit, measure and basis is moved to notes, and every such case is listed in the import report. Desk research always sets `confirmed_by_party = no`.

## B-003 · Capacity is never stored without unit, measure and basis (2026-11-03)
Battery recycling capacity is quoted in many ways (tonnes of batteries in, black mass out, metal out, GWh). Mixing them is the biggest accuracy risk, so a capacity number must carry all three, conflicting figures get separate rows, and totals only add like-for-like.

## B-002 · Status is always dated (2026-11-03)
The market moves fast (pauses, bankruptcies, sales, restarts), so every process row has `status_as_of`, warns after 90 days, and a deal needs the receiving site confirmed operating within 30 days.

## B-001 · New repo for US EV battery resource recovery; feedstock placement first (2026-11-03)
Separate from `waste_model` (Victoria contaminated soil, now paused). First goal: a complete map of US pre-processors and refiners, as the base for brokering battery scrap, end-of-life packs and black mass. The framework (workbook generator, checker, deal gate, tests, evidence rules) is adapted from waste_model.
