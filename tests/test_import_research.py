"""tools/import_research.py must load research JSON faithfully and never guess."""

import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from helpers import TODAY, blank_workbook  # noqa: F401  (sets small MAX_ROW and import paths)
import check_workbook as cw
import import_research as ir

FILE_A = {
    "sources": [{"key": "S1", "source_type": "company_website", "title": "[SYNTH] facility page",
                 "url": "https://synth.example/plant", "publisher": "Synth", "date_published": None,
                 "date_retrieved": "2026-11-02", "notes": ""}],
    "organizations": [{"key": "synth", "name": "[SYNTH] Recycling Inc", "website": "https://www.synth.example",
                       "hq_state": "NV", "ownership": "private", "parent": None, "sec_cik": None,
                       "primary_type": "integrated", "source": "S1", "notes": ""}],
    "sites": [{"key": "plant", "org": "synth", "site_name": "[SYNTH] plant", "address": None, "city": "Reno",
               "state": "NV", "zip": "89501", "county": None, "latitude": None, "longitude": None,
               "epa_id": None, "source": "S1", "notes": ""}],
    "site_process": [
        {"site": "plant", "process_stage": "mechanical_preprocessing", "status": "operating",
         "status_as_of": "2026-11-02", "capacity": 10000, "capacity_unit": "t/yr",
         "capacity_measure": "input_batteries", "capacity_basis": "nameplate_stated", "expected_start": None,
         "source": "S1", "notes": ""},
        {"site": "plant", "process_stage": "hydrometallurgical", "status": "announced",
         "status_as_of": "2026-11-02", "capacity": 5000, "capacity_unit": "t/yr", "capacity_measure": None,
         "capacity_basis": None, "expected_start": None, "source": "S1", "notes": ""},
        {"site": "plant", "process_stage": "smelting", "status": "operating", "status_as_of": "2026-11-02",
         "capacity": None, "capacity_unit": None, "capacity_measure": None, "capacity_basis": None,
         "expected_start": None, "source": "S1", "notes": ""},
    ],
    "accepts": [{"site": "plant", "feed_form": "manufacturing_scrap", "chemistry": "NMC",
                 "acceptance_conditions": None, "basis": "company_website", "source": "S1", "notes": ""}],
    "outputs": [{"site": "plant", "product": "black_mass", "source": "S1", "notes": ""}],
    "events": [{"org": "synth", "counterparty": "US DOE", "site": "plant", "event_type": "grant",
                "event_date": "2025-01-01", "amount": 1000000, "currency": "USD", "summary": "[SYNTH] grant",
                "source": "S1"},
               {"org": "synth", "counterparty": None, "site": None, "event_type": "pause",
                "event_date": "2026-01-01", "amount": None, "currency": None, "summary": "no source",
                "source": "S9"}],
    "permits": [],
    "unresolved": ["[SYNTH] lead X could not be verified"],
}
FILE_B = {  # same organization (by website) and same source URL as file A
    "sources": [{"key": "Q1", "source_type": "company_website", "title": "dup", "url": "https://synth.example/plant",
                 "publisher": "Synth", "date_published": None, "date_retrieved": "2026-11-02", "notes": ""}],
    "organizations": [{"key": "s2", "name": "Synth Recycling", "website": "synth.example", "hq_state": "NV",
                       "ownership": "private", "parent": None, "sec_cik": None, "primary_type": "integrated",
                       "source": "Q1", "notes": ""}],
    "sites": [], "site_process": [], "accepts": [], "outputs": [], "events": [], "permits": [], "unresolved": [],
}


class TestImport(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.files = []
        for name, data in (("a.json", FILE_A), ("b.json", FILE_B)):
            p = self.tmp / name
            p.write_text(json.dumps(data))
            self.files.append(str(p))
        self.out = self.tmp / "seed.xlsx"
        imp = ir.Importer()
        for f in self.files:
            imp.load(f)
        imp.resolve_parents()
        imp.write(self.out)
        imp.report(self.out, self.files)
        self.imp = imp

    def test_merges_same_source_and_organization(self):
        self.assertEqual(len(self.imp.rows["source"]), 1)
        self.assertEqual(len(self.imp.rows["organization"]), 1)

    def test_never_guesses(self):
        report = self.out.with_suffix(".md").read_text()
        self.assertIn("'smelting' is not an allowed process_stage", report)
        self.assertIn("moved to notes", report)            # capacity without measure/basis
        self.assertIn("no valid source; row skipped", report)
        self.assertIn("lead X could not be verified", report)
        hydro = [r for r in self.imp.rows["site_process"] if r["process_stage"] == "hydrometallurgical"][0]
        self.assertIsNone(hydro["capacity"])
        self.assertIn("5000", hydro["notes"])

    def test_desk_research_is_never_confirmation(self):
        self.assertTrue(all(a["confirmed_by_party"] == "no" for a in self.imp.rows["facility_accepts"]))

    def test_result_passes_checker(self):
        wb, _ = cw.check(self.out, TODAY)
        errors = [str(i) for i in wb.issues if i.level == "ERROR"]
        self.assertEqual(errors, [], "\n".join(errors))
        loaded = load_workbook(self.out)
        self.assertEqual(loaded["site"]["H3"].value, "89501")   # ZIP kept as text


if __name__ == "__main__":
    unittest.main()
