# Seed report: US lithium-ion battery pre-processors and refiners (first pass, 2026-11-03)

Workbook: [battery_intel_seed_2026-11.xlsx](battery_intel_seed_2026-11.xlsx) · import log: [battery_intel_seed_2026-11.md](battery_intel_seed_2026-11.md) · audit: [audit_2026-11-03.md](audit_2026-11-03.md) · method: [../docs/research-method.md](../docs/research-method.md)

## What's in it
| | Count |
|---|---|
| Organizations | 25 |
| Sites | 44 |
| Process rows (stage × status × capacity) | 61 |
| Acceptance rows (feed form × chemistry) | 43, all `confirmed_by_party = no` |
| Outputs / events / permits & EPA IDs | 48 / 75 / 34 |
| Sources | 175, of which **146 opened directly** (83%); 29 rest on search snippets and say so |

**Checker:** 0 errors. 38 warnings, all "status older than 90 days" (these are the sites to re-confirm first).
**Spot-check:** 10 random records re-verified: **10 supported, 0 wrong**. Two process-stage labels are researcher classifications, not stated by the source (Redwood South Carolina; ABTC Nevada), and are flagged in row notes.

**Every fact is desk research.** Nothing here has been confirmed by the companies. Statuses are as of the date shown, and several are old.

## The map at a glance (status as of the source date)

### Operating or starting up: potential outlets for feedstock
| Site | Company | What it does (per sources) | Status, as of | Capacity as stated | Published acceptance |
|---|---|---|---|---|---|
| McCarran, NV (Tahoe campus) | Redwood Materials | Pyro pre-treatment + recovery; cathode status conflicting | operating, 2026-04 | >40,000 t/yr batteries in (nameplate); >20 GWh/yr received and 60,000 t of minerals produced in 2025 (reported) | cells, modules, EV packs, consumer batteries, manufacturing scrap |
| Ridgeville, SC (Carolina campus) | Redwood Materials | Recovery (stage not stated by Redwood) | operating since 2025-11-06 | 20,000 t/yr | not published |
| McCarran, NV (2500 Peru Dr) | American Battery Technology | Shredding + hydromet; sells black mass (export licence up to $100M) | operating, 2026-10-01 (8-K) | ~20,000 t/yr batteries in | scrap, damaged/defective, all chemistries |
| Lancaster, OH | Cirba Solutions | Dismantling + shredding; hydromet circuit under construction | operating, 2026-09 | hydromet 15,000 t/yr black mass in (planned) | EV packs, manufacturing scrap |
| Wixom MI · Mesa AZ · Baltimore OH | Cirba Solutions | Collection/sorting (from EPA registrations; Li-ion processing role unverified) | operating, 2025–26 | n/a | n/a |
| Bartlesville, OK | Blue Whale Materials | Discharge + shredding to black mass | discharge operating; full line commissioning since 2026-09 | 20,000 t/yr batteries in | scrap, cells, modules, EV packs, consumer |
| Covington, GA | R3 Lithium (ex-Ascend Elements) | Shredding + lithium carbonate | commissioning, 2026-09 (bought out of Ascend's Chapter 11, July 2026) | shredding 30,000 t/yr in; Li₂CO₃ 2,500 t/yr out (Ascend said 3,000) | not published |
| Chester, SC | Princeton NuEnergy | Direct recycling of manufacturing scrap | operating, 2025-08 (stale) | 5,000 t/yr (measure not stated) | **manufacturing scrap, NMC and LFP** |
| McKinney, TX | Princeton NuEnergy (at Wistron GreenTech) | Direct-recycling pilot | operating, 2025-08 (stale) | 500 t/yr | not specified |
| Fairfield, OH | Nth Cycle | Refines black mass/scrap to MHP | operating, 2026-08 | 900 t/yr MHP out (3,100 t/yr scrap in per company) | black mass |
| Atoka, OK | Green Li-ion | Hydromet + pCAM | operating, 2025-09 (stale) | 2 t/day pCAM out | black mass |
| Carrollton, TX | Momentum Technologies | Demonstration hydromet | operating, 2026-10 | 200 t/yr (measure not stated) | black mass |
| McCarran, NV (160 Denmark Dr) | Aqua Metals | Pilot hydromet | operating, 2026-07 | 75 t/yr black mass in | black mass (NMC, LFP) |
| Canal Winchester, OH | Omega Harvested Metallurgical | Mechanical separation | "operational" in a 2023 list; **current status unknown** | not stated | all (claimed) |

### Paused, sold, cancelled or closed
- **Li-Cycle → Glencore Battery Recycling** (bought out of insolvency, August 2025):
  - spokes at Rochester NY, Gilbert AZ and Northport AL are **paused**; an EPA record calls Northport a "former Li-Cycle facility", so it may be closed;
  - the Rochester Hub is **paused**;
  - the Ohio spoke at Ultium Warren was probably never built.
- **Ascend Elements** (Chapter 11, April 2026): Apex 1, Hopkinsville KY, was never completed; it was sold to the contractor joint venture Turner-Kokosing in June 2026.
- **Aqua Metals Sierra ARC:** cancelled; the property was sold in June 2025.
- **SungEel, Toccoa GA:** cancelled in April 2025.
- **Critical Mineral Recovery, Fredericktown MO:** destroyed by fire in October 2024; rebuild plans conflict.
- **Ecobat, Casa Grande AZ:** the lithium-ion business was sold in September 2026 to an **undisclosed buyer**, so the current operator is unknown.

### Announced or under development
- ABTC second facility (South Carolina), 100,000 t/yr, no site named.
- Cirba Solutions South Carolina, 60,000 t/yr; its $200M DOE grant was on DOE's October 2025 revocation list.
- Nth Cycle "Project SHIELD", 24,000 t/yr of black mass in the Southeast; Glencore and Trafigura agreements; start as early as 2029.
- Princeton NuEnergy, Commerce GA, 3,000 t/yr.
- Clarios, Florence SC, 20,000 t/yr of cathode material.
- Li Industries, Kettering OH, LFP direct recycling.
- GMBI (LG Energy Solution / Toyota Tsusho), Winston-Salem NC, 13,500 t/yr of production scrap.
- ACE Green, Texas: lithium-ion plans unclear. Its merger with Athena was not closed at the 2026-09-28 10-Q.
- Re-New-Able Technologies, Illinois (location undisclosed).
- Envergia, Detroit: its website is suspended, so the project has probably stalled.
- Tesla, Gigafactory Nevada: in-house recycling; process not described in any source.

## What this suggests for feedstock placement (hypotheses to test, not conclusions)
1. **Shredding capacity is running, but much domestic refining has paused or been cancelled** (the Li-Cycle hub, Apex 1, Sierra ARC, SungEel). Black mass may therefore be harder to place domestically than scrap, which is consistent with ABTC's export licence. **But** a researcher reported a federal directive (from 2026-08-27) requiring black mass to be sold to US buyers for about a year. ABTC's 8-K confirms "recently-enacted federal export controls". This needs verifying first (regulatory item 5), because it changes where black mass can go.
2. **LFP is the hard chemistry to place.** It has little nickel or cobalt value, so it usually carries a processing fee. Princeton NuEnergy is the only site found that publishes acceptance of LFP manufacturing scrap; Aqua Metals' pilot also lists LFP black mass.
3. **New owners and restarts mean the incumbents' supply relationships are in flux.** R3 Lithium (restarting Covington), Blue Whale (commissioning), Glencore (deciding the future of the Li-Cycle sites) and the Ecobat buyer are all likely to be looking for feed.

## Sites to call first (feedstock placement)
Operating or commissioning, with a published or likely appetite for third-party feed. For each, ask: which feed forms and chemistries they take, conditions (state of charge, packaging, damaged batteries), minimum volumes, how they price (payable % or processing fee), and freight terms.

| # | Site | Why |
|---|---|---|
| 1 | Blue Whale Materials, Bartlesville OK | Ramping a new 20,000 t/yr line; takes scrap, cells, modules and packs |
| 2 | R3 Lithium, Covington GA | Restarting 30,000 t/yr shredding under a new owner |
| 3 | American Battery Technology, McCarran NV | Operating; takes damaged/defective batteries and scrap; sells black mass |
| 4 | Cirba Solutions, Lancaster OH (and its collection sites) | Long-running dismantling and shredding; national collection network |
| 5 | Redwood Materials, McCarran NV / Ridgeville SC | Largest operator; publishes acceptance of cells, modules and consumer batteries |
| 6 | Princeton NuEnergy, Chester SC | Only published outlet for **LFP** manufacturing scrap |
| 7 | Nth Cycle, Fairfield OH | Buys black mass and scrap for MHP |
| 8 | Green Li-ion, Atoka OK | Takes black mass; status needs re-confirming |
| 9 | Glencore Battery Recycling (ex-Li-Cycle spokes) | Paused sites: ask whether restart or third-party processing is planned |
| 10 | Ecobat Casa Grande AZ, new owner | Identify the buyer; an existing permitted lithium-ion pre-processing site |

## Gaps, conflicts and judgement calls
- **Stale statuses (38):** most "operating" claims for Princeton NuEnergy, Green Li-ion and Omega date from 2023–2025. Re-confirm by call before relying on them.
- **Capacity measure not stated:** Princeton NuEnergy Chester (5,000 t/yr) and Commerce, Momentum (200 t/yr). Ecobat's 10,000 t/yr is given for Casa Grande alone in one source and for three plants in another.
- **Conflicts kept as separate rows or notes, never averaged:**
  - Covington lithium line, 3,000 vs 2,500 t/yr;
  - Covington sale price, $3M vs $3.5M;
  - Redwood cathode production, "at scale" vs delayed;
  - Critical Mineral Recovery rebuild, same site vs elsewhere;
  - Trafigura–Nth Cycle deal, "binding agreement" vs "term sheet".
- **Process stages classified by researchers, not sources:** Redwood South Carolina and ABTC Nevada (both "hydrometallurgical"), Tesla (placeholder). These are flagged in row notes.
- **Dates known only to the month:** some event and status dates use the first or last day of the month and say so in notes ("DAY UNKNOWN", "approximate").
- **EPA IDs not found for some operating sites:** Nth Cycle, Princeton NuEnergy, Green Li-ion, Redwood South Carolina, ABTC 2500 Peru Dr. **Lead:** EPA's facility database lists Aqua Metals ID NVR000092049 (large-quantity generator) at 2500 Peru Dr, ABTC's site since 2023. Check with Nevada DEP before assigning it.

## Leads not yet verified (not in the workbook)
- **INMETCO, Ellwood City PA:** a 2023 list gives 6,000 t/yr of lithium-ion capacity; the owners later filed Chapter 7.
- **Green Metals (Toyota), Liberty NC** (2024): process unknown; relationship to GMBI unknown.
- **LOHUM / ReElement / American Metals, Marion IN** (an MOU only); **ReElement, Noblesville IN**.
- **R3 Lithium's second EPA ID** at Lithonia GA.
- **Redivivus, Los Angeles hub.**
- **ExPost Technology** (mobile pre-processing).
- **Aqua Metals' "Headwaters" site** (state not disclosed).
- **Recycling roles of Cirba's Wixom, Mesa and Baltimore sites.**

## Completeness check
- **Montana Legislature table** (2023, "EV battery recycling plants in the United States v4"): every company listed is covered, except INMETCO (above) and Li Industries' Blacksburg R&D centre, which is recorded as a site with no process.
- **DOE grant rounds:** the round-2 (2024) and round-3 (2026) recycling selectees are all covered.
- **Not extracted:** the NREL/NAATBatt supply-chain database (interactive) and Argonne's 2024 map. These are the next completeness check.
- **Checked and out of scope:** Comstock (sold its lithium-ion plant to ABTC; solar panel recycling now), Heritage (merged into Cirba), BRME, OnTo (R&D only), EVelution, and European operators with no US plant (Cylib, tozero, Altilium, Stena, Hydrovolt, Fortum, BASF, Umicore).

## Next steps
See [../docs/plan.md](../docs/plan.md):
1. Work through the regulatory checklist, starting with **black mass export controls**.
2. Call the top 10 sites.
3. Re-confirm stale statuses.
4. Run the NREL/NAATBatt and Argonne completeness check.
5. Map feedstock holders.
