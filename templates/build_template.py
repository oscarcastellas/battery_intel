"""Build the Excel collection workbook: templates/battery_intel_v1.xlsx

Run:  python templates/build_template.py [output_path]

TABLES below is the single source of truth for the workbook's tabs and columns.
sql/0001_schema.sql uses the same table and column names (tests enforce it),
and tools/check_workbook.py reads TABLES to know what to check.

Adapted from the waste_model framework (github.com/oscarcastellas/waste_model).
"""

import sys
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).with_name("battery_intel_v1.xlsx")
MAX_ROW = 2000  # rows prepared with dropdowns and lookup formulas

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")     # light yellow: you type here
AUTO_FILL = PatternFill("solid", fgColor="E7E6E6")      # grey: calculated, don't type
EXAMPLE_FONT = Font(name=FONT, italic=True, color="808080")
THIN = Side(style="thin", color="BFBFBF")

US_STATES = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "DC", "FL", "GA", "HI", "ID", "IL", "IN", "IA",
             "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM",
             "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA",
             "WV", "WI", "WY", "PR", "GU", "VI", "AS", "MP"]

# ---------------------------------------------------------------------------
# Dropdown lists (sheet "lists"). Add values there if you need more.
# ---------------------------------------------------------------------------
LISTS = {
    "state": US_STATES,
    "yes_no": ["yes", "no"],
    "yes_no_unknown": ["yes", "no", "unknown"],
    "source_type": ["company_website", "press_release", "sec_filing", "doe_announcement", "state_permit",
                    "epa_echo", "court_filing", "news", "government_report", "phone_call", "email", "meeting", "other"],
    "ownership": ["private", "public", "subsidiary", "joint_venture", "unknown"],
    "org_type": ["preprocessor", "refiner", "integrated", "oem", "cell_manufacturer", "collector", "logistics",
                 "transporter", "other"],
    "process_stage": ["collection_sorting", "discharge", "dismantling", "mechanical_preprocessing",
                      "hydrometallurgical", "pyrometallurgical", "direct_recycling", "precursor_production"],
    "status": ["announced", "permitting", "under_construction", "commissioning", "operating", "paused",
               "cancelled", "closed", "unknown"],
    "capacity_unit": ["t/yr", "kt/yr", "GWh/yr", "lb/yr", "other"],
    "capacity_measure": ["input_batteries", "input_black_mass", "output_black_mass", "output_metal_content",
                         "output_product", "unspecified"],
    "capacity_basis": ["nameplate_stated", "planned", "reported_throughput", "unspecified"],
    "feed_form": ["manufacturing_scrap", "cells", "modules", "ev_packs", "ess_packs", "consumer_batteries",
                  "black_mass", "damaged_defective_recalled", "unspecified"],
    "chemistry": ["NMC", "NCA", "LFP", "LCO", "LMO", "mixed", "all", "unspecified"],
    "product": ["black_mass", "MHP", "nickel_sulfate", "cobalt_sulfate", "manganese_sulfate", "lithium_carbonate",
                "lithium_hydroxide", "copper", "aluminum", "graphite", "alloy_matte", "pCAM", "CAM", "other"],
    "event_type": ["supply_agreement", "offtake_agreement", "grant", "loan", "investment", "acquisition",
                   "bankruptcy", "pause", "closure", "restart", "permit_issued", "enforcement", "other"],
    "permit_type": ["rcra_epa_id", "rcra_permit", "universal_waste_handler", "air_permit", "state_solid_waste", "other"],
    "accepts_basis": ["company_website", "report", "quote", "phone_call", "email", "other"],
    "contact_obtained": ["company_website", "direct_from_person", "business_card", "referral", "public_register", "other"],
    "channel": ["phone", "email", "meeting", "site_visit", "other"],
    "price_type": ["purchase_price", "processing_fee", "payable_percent"],
    "price_unit": ["$/kg", "$/lb", "$/t", "% of metal value"],
    "metal": ["Ni", "Co", "Cu", "Li", "Mn", "n/a"],
    "price_basis": ["quote", "published_price_list", "phone_call", "email", "report"],
    "lot_unit": ["kg", "t", "lb", "packs", "modules", "cells"],
    "lot_condition": ["intact", "damaged_defective", "recalled", "unknown"],
    "lot_basis": ["direct_contact", "referral", "tender", "marketplace"],
    "lot_status": ["lead", "qualified", "quoting", "placed", "lost", "withdrawn"],
    "deal_stage": ["enquiry", "quote_requested", "quoted", "proposal_sent", "accepted", "scheduled",
                   "completed", "invoiced", "paid", "lost"],
    "audit_result": ["correct", "incorrect", "changed", "could_not_verify"],
}

# ID prefix for each tab that has its own ID in column A (used by the checker).
ID_PREFIX = {
    "source": "SRC-", "organization": "ORG-", "site": "SITE-", "event": "EVT-", "contact": "CON-",
    "price_observation": "PRC-", "feed_lot": "LOT-", "deal": "DEAL-", "interaction": "INT-", "audit_log": "AUD-",
}

# ---------------------------------------------------------------------------
# Tables. Each column: (name, width, kind, help text)
#   kind: "id", "text", "number", "date", "lat", "lon", "zip", "cik",
#         "list:<list name>", "ref:<sheet>" (an ID from column A of that sheet),
#         "auto:<key column>><sheet>.<column>" (looked up automatically; never typed)
# ---------------------------------------------------------------------------
TABLES = {
    "source": [
        ("source_id", 13, "id", "Your ID, e.g. SRC-0001. Every fact in other tabs points to one of these."),
        ("source_type", 18, "list:source_type", "What kind of source this is."),
        ("title", 40, "text", "Short description of the page or document."),
        ("url_or_reference", 45, "text", "Link, or a reference such as 'call with plant manager'."),
        ("publisher", 22, "text", "Who published it."),
        ("date_published", 14, "date", "When the source published it, if stated. Blank if not stated."),
        ("date_retrieved", 14, "date", "When you looked at it. Required."),
        ("notes", 40, "text", "What it does or doesn't say; limitations."),
    ],
    "organization": [
        ("org_id", 13, "id", "Your ID, e.g. ORG-0001."),
        ("name", 34, "text", "Company name as it presents itself."),
        ("website", 28, "text", "Main website (used to catch duplicates)."),
        ("hq_state", 8, "list:state", ""),
        ("ownership", 13, "list:ownership", ""),
        ("parent_org_id", 13, "ref:organization", "Parent company, if a subsidiary."),
        ("sec_cik", 12, "cik", "SEC CIK number for public companies (up to 10 digits)."),
        ("primary_type", 14, "list:org_type", "Main role."),
        ("other_types", 20, "text", "e.g. 'collector, logistics'."),
        ("usdot_number", 12, "text", "Transporters: USDOT number (FMCSA)."),
        ("source_id", 13, "ref:source", "Where the organization details came from."),
        ("date_checked", 13, "date", "When you last confirmed these details."),
        ("notes", 40, "text", ""),
    ],
    "site": [
        ("site_id", 13, "id", "Your ID, e.g. SITE-0001. One row per physical location."),
        ("org_id", 13, "ref:organization", "Operator of the site."),
        ("org_name", 28, "auto:org_id>organization.name", "Filled automatically."),
        ("site_name", 28, "text", "e.g. 'McCarran campus'."),
        ("address", 32, "text", "Street address as published."),
        ("city", 16, "text", ""),
        ("state", 7, "list:state", ""),
        ("zip", 8, "zip", "5-digit ZIP (or ZIP+4)."),
        ("county", 16, "text", ""),
        ("latitude", 11, "lat", "Decimal degrees, e.g. 39.5296."),
        ("longitude", 11, "lon", "Decimal degrees, negative, e.g. -119.8138."),
        ("epa_id", 14, "text", "RCRA handler EPA ID number if found (e.g. on EPA ECHO). Strong duplicate guard."),
        ("source_id", 13, "ref:source", "Where the site details came from."),
        ("date_checked", 13, "date", "When you last confirmed these details."),
        ("notes", 40, "text", ""),
    ],
    "site_process": [
        ("site_id", 13, "ref:site", "The site."),
        ("site_name", 26, "auto:site_id>site.site_name", "Filled automatically."),
        ("process_stage", 22, "list:process_stage", "One row per stage at a site."),
        ("status", 16, "list:status", "Status of this stage at this site."),
        ("status_as_of", 12, "date", "Date of the newest source supporting the status. Required."),
        ("capacity", 11, "number", "Exactly as stated. Blank if not stated."),
        ("capacity_unit", 10, "list:capacity_unit", "Required with capacity."),
        ("capacity_measure", 18, "list:capacity_measure", "What the capacity counts (batteries in, black mass out...). Required with capacity."),
        ("capacity_basis", 16, "list:capacity_basis", "Stated nameplate, planned, or reported throughput. Required with capacity."),
        ("expected_start", 12, "date", "For non-operating stages, if stated."),
        ("source_id", 13, "ref:source", "Required."),
        ("date_checked", 13, "date", ""),
        ("notes", 40, "text", "Conflicting figures: keep each as its own row with its own source."),
    ],
    "facility_accepts": [
        ("site_id", 13, "ref:site", ""),
        ("site_name", 26, "auto:site_id>site.site_name", "Filled automatically."),
        ("feed_form", 22, "list:feed_form", ""),
        ("chemistry", 11, "list:chemistry", ""),
        ("acceptance_conditions", 36, "text", "State of charge, packaging, damaged batteries, minimum volumes..."),
        ("basis", 15, "list:accepts_basis", "company_website = a claim; only a quote/call/email from them is confirmation."),
        ("confirmed_by_party", 12, "list:yes_no", "'yes' only if the facility itself confirmed it takes this feed and chemistry."),
        ("source_id", 13, "ref:source", "Required."),
        ("date_checked", 13, "date", ""),
        ("notes", 36, "text", ""),
    ],
    "site_output": [
        ("site_id", 13, "ref:site", ""),
        ("site_name", 26, "auto:site_id>site.site_name", "Filled automatically."),
        ("product", 18, "list:product", ""),
        ("source_id", 13, "ref:source", "Required."),
        ("notes", 40, "text", ""),
    ],
    "event": [
        ("event_id", 13, "id", "Your ID, e.g. EVT-0001."),
        ("org_id", 13, "ref:organization", "Main organization."),
        ("org_name", 26, "auto:org_id>organization.name", "Filled automatically."),
        ("counterparty_org_id", 13, "ref:organization", "Other party, if in the directory."),
        ("counterparty_name", 24, "text", "Other party's name if not in the directory (e.g. an OEM, DOE)."),
        ("site_id", 13, "ref:site", "Site concerned, if any."),
        ("event_type", 18, "list:event_type", ""),
        ("event_date", 12, "date", "Date it happened or was announced."),
        ("amount", 14, "number", "As stated."),
        ("currency", 8, "text", "e.g. USD. Required with amount."),
        ("summary", 50, "text", "What happened, in one or two sentences."),
        ("source_id", 13, "ref:source", "Required."),
    ],
    "permit": [
        ("site_id", 13, "ref:site", ""),
        ("site_name", 26, "auto:site_id>site.site_name", "Filled automatically."),
        ("regulator", 18, "text", "e.g. US EPA Region 9, Nevada DEP."),
        ("permit_type", 20, "list:permit_type", ""),
        ("number", 18, "text", "Exactly as shown."),
        ("status", 12, "text", "As shown by the regulator."),
        ("date_checked", 13, "date", "When you last checked the regulator's record. Required."),
        ("source_id", 13, "ref:source", "Required."),
        ("notes", 36, "text", ""),
    ],
    "contact": [
        ("contact_id", 13, "id", "Your ID, e.g. CON-0001. Business contacts only."),
        ("org_id", 13, "ref:organization", ""),
        ("org_name", 26, "auto:org_id>organization.name", "Filled automatically."),
        ("site_id", 13, "ref:site", "Optional."),
        ("name", 22, "text", ""),
        ("role", 22, "text", "e.g. 'Feedstock procurement manager'."),
        ("work_phone", 16, "text", "Work number only."),
        ("work_email", 28, "text", "Work email only."),
        ("how_obtained", 18, "list:contact_obtained", "Where you got the details (privacy record)."),
        ("source_id", 13, "ref:source", ""),
        ("date_checked", 13, "date", ""),
        ("notes", 36, "text", ""),
    ],
    "price_observation": [
        ("price_id", 13, "id", "Your ID, e.g. PRC-0001. One row per price you see or are quoted."),
        ("site_id", 13, "ref:site", "Receiving site (or transporter depot for freight)."),
        ("site_name", 26, "auto:site_id>site.site_name", "Filled automatically."),
        ("feed_form", 20, "list:feed_form", ""),
        ("chemistry", 11, "list:chemistry", ""),
        ("price_type", 16, "list:price_type", "purchase_price = they pay you; processing_fee = you pay them; payable_percent = % of metal value."),
        ("amount", 10, "number", "As quoted (always positive; direction comes from price_type)."),
        ("price_unit", 14, "list:price_unit", ""),
        ("metal", 8, "list:metal", "For payable_percent: which metal. Otherwise n/a."),
        ("includes_freight", 12, "list:yes_no_unknown", ""),
        ("valid_from", 12, "date", "Date of the quote."),
        ("valid_until", 12, "date", "Expiry, if stated."),
        ("basis", 16, "list:price_basis", ""),
        ("source_id", 13, "ref:source", "Required."),
        ("notes", 36, "text", "Metal price reference, minimum volumes, conditions."),
    ],
    "feed_lot": [
        ("lot_id", 13, "id", "Your ID, e.g. LOT-0001. Feedstock someone wants placed."),
        ("holder_org_id", 13, "ref:organization", "Who holds the material."),
        ("holder_name", 26, "auto:holder_org_id>organization.name", "Filled automatically."),
        ("site_id", 13, "ref:site", "Where the material is."),
        ("feed_form", 20, "list:feed_form", ""),
        ("chemistry", 11, "list:chemistry", "Required (not unspecified) before quoting."),
        ("quantity", 10, "number", "As stated."),
        ("unit", 8, "list:lot_unit", ""),
        ("condition", 16, "list:lot_condition", "Damaged/defective/recalled changes shipping rules and who can take it."),
        ("state_of_charge", 14, "text", "As stated, e.g. '<30% SOC'."),
        ("packaging", 22, "text", "As stated."),
        ("window_start", 12, "date", ""),
        ("window_end", 12, "date", ""),
        ("basis", 14, "list:lot_basis", ""),
        ("status", 11, "list:lot_status", ""),
        ("source_id", 13, "ref:source", "Required."),
        ("date_checked", 13, "date", ""),
        ("notes", 36, "text", ""),
    ],
    "deal": [
        ("deal_id", 13, "id", "Your ID, e.g. DEAL-0001."),
        ("lot_id", 13, "ref:feed_lot", ""),
        ("receiving_site_id", 13, "ref:site", "Processor receiving the material."),
        ("receiving_name", 26, "auto:receiving_site_id>site.site_name", "Filled automatically."),
        ("transporter_org_id", 13, "ref:organization", "Hazmat-capable carrier."),
        ("transporter_name", 24, "auto:transporter_org_id>organization.name", "Filled automatically."),
        ("price_id", 13, "ref:price_observation", "Processor price used."),
        ("transport_price_id", 13, "ref:price_observation", "Freight price used (optional)."),
        ("stage", 14, "list:deal_stage", "From proposal_sent onward, the deal gate must pass."),
        ("stage_date", 12, "date", ""),
        ("fee_or_margin", 12, "number", "Your fee or margin, $."),
        ("chemistry_condition_verified", 14, "list:yes_no", "Chemistry, condition and state of charge verified with the holder."),
        ("receiving_site_confirmed", 14, "list:yes_no", "Processor confirmed it will take this lot."),
        ("transporter_hazmat_verified", 14, "list:yes_no", "Carrier's USDOT and hazmat capability checked."),
        ("packaging_compliant", 14, "list:yes_no", "Packaging meets the shipping rules for this condition."),
        ("outcome_reason", 28, "text", ""),
        ("notes", 36, "text", ""),
    ],
    "interaction": [
        ("interaction_id", 13, "id", "Your ID, e.g. INT-0001."),
        ("interaction_date", 12, "date", ""),
        ("contact_id", 13, "ref:contact", ""),
        ("contact_name", 22, "auto:contact_id>contact.name", "Filled automatically."),
        ("org_id", 13, "ref:organization", ""),
        ("channel", 11, "list:channel", ""),
        ("summary", 50, "text", "What was said; facts learned (add them to the right tab with this as their source)."),
        ("outcome", 28, "text", ""),
        ("follow_up_date", 13, "date", ""),
        ("lot_id", 13, "ref:feed_lot", "Optional."),
        ("deal_id", 13, "ref:deal", "Optional."),
    ],
    "audit_log": [
        ("audit_id", 13, "id", "Your ID, e.g. AUD-0001."),
        ("audit_date", 12, "date", ""),
        ("tab", 18, "text", "Which tab the checked record is in."),
        ("record_id", 18, "text", "The record's ID (or site_id + stage for rows without an ID)."),
        ("field", 18, "text", "Which value you re-checked."),
        ("result", 16, "list:audit_result", ""),
        ("notes", 40, "text", "What the correct value is, and what you fixed."),
    ],
}

# One clearly-marked example row per table, one value per column (auto columns: None).
# IDs start with EX-; these rows are never imported and the checker skips them.
EXAMPLES = {
    "source": ["EX-SRC-1", "company_website", "(example) Company facility page", "https://... (example)",
               "Example Recycling Inc", None, "2026-11-03", "Example row - delete or keep"],
    "organization": ["EX-ORG-1", "(example) Example Recycling Inc", "example.com", "NV", "private", None, None,
                     "integrated", None, None, "EX-SRC-1", "2026-11-03", "Example row"],
    "site": ["EX-SITE-1", "EX-ORG-1", None, "(example) Reno plant", "1 Example Way", "Reno", "NV", "89501",
             "Washoe", 39.5296, -119.8138, None, "EX-SRC-1", "2026-11-03", "Example row"],
    "site_process": ["EX-SITE-1", None, "mechanical_preprocessing", "operating", "2026-11-03", 10000, "t/yr",
                     "input_batteries", "nameplate_stated", None, "EX-SRC-1", "2026-11-03", "Example row"],
    "facility_accepts": ["EX-SITE-1", None, "manufacturing_scrap", "NMC", "Discharged; palletised", "company_website",
                         "no", "EX-SRC-1", "2026-11-03", "Example row"],
    "site_output": ["EX-SITE-1", None, "black_mass", "EX-SRC-1", "Example row"],
    "event": ["EX-EVT-1", "EX-ORG-1", None, None, "(example) US DOE", "EX-SITE-1", "grant", "2026-11-03", 1000000,
              "USD", "(example) Grant for recycling capacity", "EX-SRC-1"],
    "permit": ["EX-SITE-1", None, "US EPA", "rcra_epa_id", "XXR000000000", "active", "2026-11-03", "EX-SRC-1",
               "Example row"],
    "contact": ["EX-CON-1", "EX-ORG-1", None, "EX-SITE-1", "(example) Sam Lee", "Feedstock manager", None,
                "sam@example.com", "company_website", "EX-SRC-1", "2026-11-03", "Example row"],
    "price_observation": ["EX-PRC-1", "EX-SITE-1", None, "manufacturing_scrap", "NMC", "payable_percent", 0,
                          "% of metal value", "Ni", "no", "2026-11-03", None, "quote", "EX-SRC-1",
                          "Example row (amount 0 = placeholder)"],
    "feed_lot": ["EX-LOT-1", "EX-ORG-1", None, "EX-SITE-1", "manufacturing_scrap", "NMC", 20, "t", "intact",
                 None, "Gaylords", "2026-12-01", "2027-01-31", "direct_contact", "lead", "EX-SRC-1", "2026-11-03",
                 "Example row"],
    "deal": ["EX-DEAL-1", "EX-LOT-1", "EX-SITE-1", None, "EX-ORG-1", None, "EX-PRC-1", None, "enquiry",
             "2026-11-03", None, "no", "no", "no", "no", None, "Example row"],
    "interaction": ["EX-INT-1", "2026-11-03", "EX-CON-1", None, "EX-ORG-1", "phone",
                    "(example) Asked which chemistries and forms they accept", "Will send spec sheet",
                    "2026-11-10", None, None],
    "audit_log": ["EX-AUD-1", "2026-11-03", "site_process", "EX-SITE-1", "status", "correct", "Example row"],
}

README_LINES = [
    ("Battery Intel: US lithium-ion battery recycling map, workbook v1", "title"),
    ("Purpose: a complete, accurate map of US lithium-ion battery pre-processors and refiners (what each site does, "
     "its status and capacity, what feed it accepts), as the base for feedstock placement deals. Every fact has a source.", None),
    ("", None),
    ("Tabs", "h"),
    ("source: every page, document, call or email used.  organization / site: companies and their physical sites.", None),
    ("site_process: what each site does (discharge, shredding to black mass, hydro/pyro refining...), its dated status and capacity.", None),
    ("facility_accepts / site_output: what feed each site takes and what it produces.  event: agreements, grants, loans, pauses, closures.", None),
    ("permit: EPA ID numbers and permits.  contact / interaction: people and conversations (business details only).", None),
    ("price_observation / feed_lot / deal: prices and payables, feedstock to place, and placement deals.  audit_log: monthly spot-checks.", None),
    ("", None),
    ("Check the data", "h"),
    ("Run:  python tools/check_workbook.py <this file>   It lists errors (must fix) and warnings. "
     "A deal cannot move to proposal_sent or beyond until the deal gate passes.", None),
    ("", None),
    ("Legend", "h"),
    ("Yellow cells: you type here (many have dropdowns).", "input"),
    ("Grey cells: filled automatically. Don't type in them. 'NOT FOUND' means the ID you typed doesn't exist yet.", "auto"),
    ("Grey italic rows with IDs starting EX-: examples showing the expected format. Never imported.", "example"),
    ("Hover over any column header to see what goes in it.", None),
    ("", None),
    ("Rules", "h"),
    ("1. Every fact points to a source. No source, no row.  2. Blank means unknown; never fill in likely values.", None),
    ("3. Status always has status_as_of. This market changes fast: plants pause, get sold or close.", None),
    ("4. Capacity is recorded exactly as stated, with unit, what it measures and its basis. Never convert or average; "
     "conflicting figures get separate rows with their own sources.", None),
    ("5. A company website shows what a facility claims to accept. Only the facility itself (quote, call, email) "
     "makes confirmed_by_party = yes.", None),
    ("", None),
    ("Template generated by templates/build_template.py. Column names match sql/0001_schema.sql.", None),
]


def column_letter(table, column_name):
    names = [c[0] for c in TABLES[table]]
    return get_column_letter(names.index(column_name) + 1)


def auto_formula(table, kind, row):
    """'auto:org_id>organization.name' -> INDEX/MATCH lookup on this row."""
    key_col, target = kind[len("auto:"):].split(">")
    sheet, value_col = target.split(".")
    k = column_letter(table, key_col)
    v = column_letter(sheet, value_col)
    return f'=IF({k}{row}="","",IFERROR(INDEX({sheet}!${v}:${v},MATCH({k}{row},{sheet}!$A:$A,0)),"NOT FOUND"))'


def style_header(cell, help_text):
    cell.font = Font(name=FONT, bold=True, color="FFFFFF")
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    if help_text:
        cell.comment = Comment(help_text, "template")


def add_validation(ws, letter, dv):
    ws.add_data_validation(dv)
    dv.add(f"{letter}2:{letter}{MAX_ROW}")


def list_validation(formula, prompt=None):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True, showDropDown=False)
    dv.error = "Pick a value from the list (or add it to the lists tab first)."
    dv.errorTitle = "Not in list"
    if prompt:
        dv.prompt = prompt
    return dv


def simple_validation(kind, letter):
    if kind == "date":
        dv = DataValidation(type="date", operator="greaterThan", formula1="DATE(1990,1,1)", allow_blank=True)
        dv.error = "Enter a date as yyyy-mm-dd."
    elif kind == "lat":
        dv = DataValidation(type="decimal", operator="between", formula1="17", formula2="72", allow_blank=True)
        dv.error = "US latitudes are between 17 and 72."
    elif kind == "lon":
        dv = DataValidation(type="decimal", operator="between", formula1="-180", formula2="-64", allow_blank=True)
        dv.error = "US longitudes are negative, between -180 and -64."
    elif kind == "number":
        dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
        dv.error = "Enter a number (0 or more). Leave blank if unknown."
    else:
        return None
    return dv


def write_row(ws, table, row, values, font=None):
    """Write values into input columns only (auto columns keep their formulas)."""
    for c_idx, (v, (_, _, kind, _)) in enumerate(zip(values, TABLES[table]), start=1):
        if kind.startswith("auto:"):
            continue
        if kind == "date" and isinstance(v, str):
            v = date.fromisoformat(v)
        cell = ws.cell(row=row, column=c_idx, value=v)
        if font:
            cell.font = font


def build(out=OUT):
    wb = Workbook()

    ws = wb.active
    ws.title = "README"
    ws.column_dimensions["A"].width = 140
    styles = {"title": Font(name=FONT, size=14, bold=True), "h": Font(name=FONT, size=11, bold=True),
              "example": EXAMPLE_FONT}
    for i, (text, style) in enumerate(README_LINES, start=1):
        c = ws.cell(row=i, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.font = styles.get(style, Font(name=FONT, size=10))
        if style == "input":
            c.fill = INPUT_FILL
        elif style == "auto":
            c.fill = AUTO_FILL

    lists_ws = wb.create_sheet("lists")
    list_ranges = {}
    for col_idx, (name, values) in enumerate(LISTS.items(), start=1):
        letter = get_column_letter(col_idx)
        style_header(lists_ws.cell(row=1, column=col_idx, value=name),
                     "Dropdown values. Add new values at the bottom of the column (up to row 80).")
        for r, v in enumerate(values, start=2):
            lists_ws.cell(row=r, column=col_idx, value=v).font = Font(name=FONT)
        lists_ws.column_dimensions[letter].width = max(12, max(len(v) for v in values) + 2)
        list_ranges[name] = f"lists!${letter}$2:${letter}$80"
    lists_ws.freeze_panes = "A2"

    for table, columns in TABLES.items():
        ws = wb.create_sheet(table, index=len(wb.sheetnames) - 1)  # keep 'lists' last
        for c_idx, (name, width, kind, help_text) in enumerate(columns, start=1):
            letter = get_column_letter(c_idx)
            style_header(ws.cell(row=1, column=c_idx, value=name), help_text)
            ws.column_dimensions[letter].width = width

            if kind.startswith("auto:"):
                for r in range(2, MAX_ROW + 1):
                    cell = ws.cell(row=r, column=c_idx, value=auto_formula(table, kind, r))
                    cell.fill = AUTO_FILL
                    cell.font = Font(name=FONT, color="595959")
                continue

            for r in range(2, MAX_ROW + 1):
                cell = ws.cell(row=r, column=c_idx)
                cell.fill = INPUT_FILL
                cell.font = Font(name=FONT)
                cell.border = Border(bottom=THIN)
                if kind == "date":
                    cell.number_format = "yyyy-mm-dd"
                elif kind in ("id", "text", "zip", "cik"):
                    cell.number_format = "@"   # keeps leading zeros in ZIPs and CIKs
                elif kind in ("lat", "lon"):
                    cell.number_format = "0.000000"

            if kind.startswith("list:"):
                add_validation(ws, letter, list_validation(list_ranges[kind[5:]]))
            elif kind.startswith("ref:"):
                ref_sheet = kind[4:]
                add_validation(ws, letter, list_validation(f"{ref_sheet}!$A$2:$A${MAX_ROW}",
                                                           prompt=f"Pick an ID from the {ref_sheet} tab."))
            else:
                dv = simple_validation(kind, letter)
                if dv:
                    add_validation(ws, letter, dv)

        if table in EXAMPLES:
            write_row(ws, table, 2, EXAMPLES[table], font=EXAMPLE_FONT)
        ws.freeze_panes = "B2"
        ws.row_dimensions[1].height = 30

    wb.save(out)
    print(f"Wrote {out}")
    return out


if __name__ == "__main__":
    build(Path(sys.argv[1]) if len(sys.argv) > 1 else OUT)
