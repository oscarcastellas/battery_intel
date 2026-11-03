# Research output format (JSON), one file per agent

Top-level object:
{
  "sources":   [ {"key": "S1", "source_type": <source_type>, "title": str, "url": str, "publisher": str,
                  "date_published": "YYYY-MM-DD" or null, "date_retrieved": "YYYY-MM-DD", "notes": str} ],
  "organizations": [ {"key": "redwood", "name": str, "website": str|null, "hq_state": "NV"|null,
                  "ownership": "private"|"public"|"subsidiary"|"unknown", "parent": str|null, "sec_cik": str|null,
                  "primary_type": <org_type>, "source": "S1", "notes": str} ],
  "sites": [ {"key": "redwood_mccarran", "org": "redwood", "site_name": str, "address": str|null, "city": str|null,
                  "state": "NV", "zip": "89434"|null, "county": str|null, "latitude": float|null, "longitude": float|null,
                  "epa_id": str|null, "source": "S2", "notes": str} ],
  "site_process": [ {"site": "redwood_mccarran", "process_stage": <process_stage>, "status": <status>,
                  "status_as_of": "YYYY-MM-DD", "capacity": number|null, "capacity_unit": <capacity_unit>|null,
                  "capacity_measure": <capacity_measure>|null, "capacity_basis": <capacity_basis>|null,
                  "expected_start": "YYYY-MM-DD"|null, "source": "S3", "notes": str} ],
  "accepts": [ {"site": key, "feed_form": <feed_form>, "chemistry": <chemistry>, "acceptance_conditions": str|null,
                  "basis": "company_website"|"report"|"other", "source": key, "notes": str} ],
  "outputs": [ {"site": key, "product": <product>, "source": key, "notes": str} ],
  "events":  [ {"org": key, "counterparty": org name or key|null, "site": key|null, "event_type": <event_type>,
                  "event_date": "YYYY-MM-DD", "amount": number|null, "currency": "USD"|null, "summary": str, "source": key} ],
  "permits": [ {"site": key, "regulator": str, "permit_type": <permit_type>, "number": str, "status": str|null, "source": key} ],
  "unresolved": [ str ]   // leads you could not verify, conflicts, gaps
}

Allowed values (use EXACTLY these strings):
- source_type: company_website, press_release, sec_filing, doe_announcement, state_permit, epa_echo, court_filing, news, government_report, other
- org_type: preprocessor, refiner, integrated, oem, cell_manufacturer, collector, logistics, other
- process_stage: collection_sorting, discharge, dismantling, mechanical_preprocessing, hydrometallurgical, pyrometallurgical, direct_recycling, precursor_production
- status: announced, permitting, under_construction, commissioning, operating, paused, cancelled, closed, unknown
- capacity_unit: t/yr, kt/yr, GWh/yr, lb/yr, other
- capacity_measure: input_batteries, input_black_mass, output_black_mass, output_metal_content, output_product, unspecified
- capacity_basis: nameplate_stated, planned, reported_throughput, unspecified
- feed_form: manufacturing_scrap, cells, modules, ev_packs, ess_packs, consumer_batteries, black_mass, damaged_defective_recalled, unspecified
- chemistry: NMC, NCA, LFP, LCO, LMO, mixed, all, unspecified
- product: black_mass, MHP, nickel_sulfate, cobalt_sulfate, manganese_sulfate, lithium_carbonate, lithium_hydroxide, copper, aluminum, graphite, alloy_matte, pCAM, CAM, other
- event_type: supply_agreement, offtake_agreement, grant, loan, investment, acquisition, bankruptcy, pause, closure, restart, permit_issued, enforcement, other
- permit_type: rcra_epa_id, rcra_permit, universal_waste_handler, air_permit, state_solid_waste, other

Rules (non-negotiable):
1. Every row has a "source" key pointing to a sources entry with a real URL you actually opened (or a search result whose snippet you quote in notes). No source, no row.
2. Never guess. Unknown = null. Don't infer a ZIP, coordinates or capacity. Coordinates only if a source gives them or the exact street address is stated (then you may leave lat/long null; we geocode later).
3. Capacity exactly as stated, with unit, what it measures (batteries in? black mass out? metal out?) and basis. If two sources disagree, record two site_process rows with separate sources and explain in notes. Never average or convert.
4. Status always with status_as_of = the date of the newest source supporting it. Prefer primary sources (company, SEC, DOE, state regulator) over press. Look actively for pauses, cancellations, bankruptcies, sales, closures (2023-2026).
5. One site_process row per (site, process_stage). A site that shreds and refines gets two rows.
6. Only US sites (50 states + DC + territories). Foreign sites of US companies: skip (mention in unresolved if relevant).
7. Scope: lithium-ion battery pre-processing (discharge, dismantling, shredding to black mass) and refining (hydro, pyro, direct recycling), plus precursor production only where it's on the same site. Collectors/logistics only as organizations if they own a pre-processing site. Lead-acid: out of scope.
8. Write the JSON to the file path given in your task, valid JSON (no comments), then reply with a short summary: counts, biggest uncertainties.
