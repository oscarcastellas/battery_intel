"""Shared helpers for the tests: build a small workbook and fill it with [SYNTH] data."""

import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "templates"))
sys.path.insert(0, str(ROOT / "tools"))

import build_template as bt  # noqa: E402

TODAY = date(2026, 11, 3)
_template = None


def blank_workbook():
    """Path to a fresh copy of a small (60-row) template, built once per test run."""
    global _template
    tmp = Path(tempfile.mkdtemp())
    if _template is None:
        bt.MAX_ROW = 60
        _template = tmp / "template.xlsx"
        bt.build(_template)
    path = tmp / "wb.xlsx"
    shutil.copy(_template, path)
    return path


def put(ws, table, **values):
    """Write a row into the first empty row of a tab, by column name."""
    names = [c[0] for c in bt.TABLES[table]]
    row = 2
    while ws.cell(row=row, column=1).value not in (None, ""):
        row += 1
    for name, v in values.items():
        ws.cell(row=row, column=names.index(name) + 1, value=v)
    return row


def good_scenario(path, **overrides):
    """A complete, valid placement of NMC manufacturing scrap; deal at proposal_sent.
    `overrides` maps 'tab.column' -> value for the FIRST row of that tab, to inject one fault."""
    wb = load_workbook(path)
    d = date
    rows = [
        ("source", dict(source_id="SRC-0001", source_type="phone_call", title="[SYNTH] call",
                        url_or_reference="call with procurement", date_retrieved=d(2026, 10, 30))),
        ("organization", dict(org_id="ORG-0001", name="[SYNTH] Recycler Co", website="https://recycler.example",
                              hq_state="NV", primary_type="integrated", source_id="SRC-0001", date_checked=d(2026, 10, 30))),
        ("organization", dict(org_id="ORG-0002", name="[SYNTH] Hazmat Haulers", primary_type="transporter",
                              usdot_number="1234567", source_id="SRC-0001", date_checked=d(2026, 10, 30))),
        ("organization", dict(org_id="ORG-0003", name="[SYNTH] Cell Maker", primary_type="cell_manufacturer",
                              source_id="SRC-0001", date_checked=d(2026, 10, 30))),
        ("site", dict(site_id="SITE-0001", org_id="ORG-0001", site_name="[SYNTH] plant", city="Reno", state="NV",
                      zip="89501", latitude=39.53, longitude=-119.81, source_id="SRC-0001", date_checked=d(2026, 10, 30))),
        ("site", dict(site_id="SITE-0002", org_id="ORG-0003", site_name="[SYNTH] cell plant", state="GA",
                      zip="30014", latitude=33.60, longitude=-83.85, source_id="SRC-0001", date_checked=d(2026, 10, 30))),
        ("site_process", dict(site_id="SITE-0001", process_stage="mechanical_preprocessing", status="operating",
                              status_as_of=d(2026, 10, 30), capacity=20000, capacity_unit="t/yr",
                              capacity_measure="input_batteries", capacity_basis="nameplate_stated",
                              source_id="SRC-0001", date_checked=d(2026, 10, 30))),
        ("facility_accepts", dict(site_id="SITE-0001", feed_form="manufacturing_scrap", chemistry="NMC",
                                  basis="phone_call", confirmed_by_party="yes", source_id="SRC-0001",
                                  date_checked=d(2026, 10, 30))),
        ("site_output", dict(site_id="SITE-0001", product="black_mass", source_id="SRC-0001")),
        ("event", dict(event_id="EVT-0001", org_id="ORG-0001", counterparty_name="[SYNTH] OEM", site_id="SITE-0001",
                       event_type="supply_agreement", event_date=d(2026, 6, 1), summary="[SYNTH] supply deal",
                       source_id="SRC-0001")),
        ("permit", dict(site_id="SITE-0001", regulator="US EPA", permit_type="rcra_epa_id", number="NVR000000001",
                        date_checked=d(2026, 10, 30), source_id="SRC-0001")),
        ("price_observation", dict(price_id="PRC-0001", site_id="SITE-0001", feed_form="manufacturing_scrap",
                                   chemistry="NMC", price_type="payable_percent", amount=70,
                                   price_unit="% of metal value", metal="Ni", includes_freight="no",
                                   valid_from=d(2026, 10, 30), basis="quote", source_id="SRC-0001")),
        ("feed_lot", dict(lot_id="LOT-0001", holder_org_id="ORG-0003", site_id="SITE-0002",
                          feed_form="manufacturing_scrap", chemistry="NMC", quantity=20, unit="t", condition="intact",
                          basis="direct_contact", status="quoting", source_id="SRC-0001", date_checked=d(2026, 10, 30))),
        ("deal", dict(deal_id="DEAL-0001", lot_id="LOT-0001", receiving_site_id="SITE-0001",
                      transporter_org_id="ORG-0002", price_id="PRC-0001", stage="proposal_sent",
                      stage_date=d(2026, 11, 1), chemistry_condition_verified="yes", receiving_site_confirmed="yes",
                      transporter_hazmat_verified="yes", packaging_compliant="yes")),
    ]
    seen = set()
    for table, values in rows:
        if table not in seen:
            for key, v in overrides.items():
                t, col = key.split(".")
                if t == table:
                    values[col] = v
            seen.add(table)
        put(wb[table], table, **values)
    wb.save(path)
    return path
