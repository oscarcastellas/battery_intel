"""Tests for tools/check_workbook.py. Run:  python -m unittest discover -s tests"""

import unittest
from datetime import date

from openpyxl import load_workbook

from helpers import TODAY, blank_workbook, good_scenario, put

import check_workbook as cw


def run(path):
    wb, _ = cw.check(path, TODAY)
    return [str(i) for i in wb.issues if i.level == "ERROR"], [str(i) for i in wb.issues if i.level == "WARNING"]


def assert_flagged(test, messages, fragment):
    test.assertTrue(any(fragment in m for m in messages), f"expected '{fragment}' in:\n" + "\n".join(messages))


class TestBaseline(unittest.TestCase):
    def test_blank_template_is_clean(self):
        errors, warnings = run(blank_workbook())
        self.assertEqual((errors, warnings), ([], []))

    def test_complete_placement_passes(self):
        errors, warnings = run(good_scenario(blank_workbook()))
        self.assertEqual(errors, [], "\n".join(errors))
        self.assertEqual(warnings, [], "\n".join(warnings))


class TestFaultsAreCaught(unittest.TestCase):
    def check(self, fragment, warning=False, **override):
        errors, warnings = run(good_scenario(blank_workbook(), **override))
        assert_flagged(self, warnings if warning else errors, fragment)

    def test_bad_zip(self):
        self.check("not a 5-digit ZIP", **{"site.zip": "8950"})

    def test_positive_longitude(self):
        self.check("outside the US", **{"site.longitude": 119.81})

    def test_capacity_without_measure(self):
        self.check("capacity without capacity_measure", **{"site_process.capacity_measure": None})

    def test_stale_status_warns_and_blocks_deal(self):
        self.check("over 90 days ago", warning=True, **{"site_process.status_as_of": date(2026, 6, 1)})
        self.check("not confirmed operating", **{"site_process.status_as_of": date(2026, 6, 1)})

    def test_paused_site_blocks_deal(self):
        self.check("not confirmed operating", **{"site_process.status": "paused"})

    def test_future_status_date(self):
        self.check("is in the future", **{"site_process.status_as_of": date(2027, 1, 1)})

    def test_unconfirmed_acceptance_blocks_deal(self):
        self.check("has not confirmed it accepts", **{"facility_accepts.confirmed_by_party": "no"})

    def test_wrong_chemistry_blocks_deal(self):
        self.check("has not confirmed it accepts", **{"facility_accepts.chemistry": "LFP"})

    def test_damaged_lot_needs_damaged_acceptance(self):
        self.check("damaged_defective_recalled", **{"feed_lot.condition": "damaged_defective"})

    def test_transporter_without_usdot(self):
        path = good_scenario(blank_workbook())
        wb = load_workbook(path)
        ws = wb["organization"]
        col = [c[0] for c in cw.bt.TABLES["organization"]].index("usdot_number") + 1
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=1).value == "ORG-0002":
                ws.cell(row=r, column=col).value = None   # cell(value=None) would not clear it
        wb.save(path)
        errors, _ = run(path)
        assert_flagged(self, errors, "no USDOT number")

    def test_payable_without_metal(self):
        self.check("needs unit '% of metal value' and a metal", **{"price_observation.metal": "n/a"})

    def test_stale_price_blocks_deal(self):
        self.check("no current price", **{"price_observation.valid_from": date(2026, 6, 1)})

    def test_lot_quoting_with_unknown_condition(self):
        self.check("chemistry or condition is unknown", **{"feed_lot.condition": "unknown"})

    def test_broken_reference(self):
        self.check("does not exist in the site tab", **{"site_process.site_id": "SITE-9999"})

    def test_example_reference(self):
        self.check("points to an example row", **{"feed_lot.site_id": "EX-SITE-1"})

    def test_wrong_prefix(self):
        self.check("should start with ORG-", **{"organization.org_id": "ORG0001"})

    def test_event_amount_without_currency(self):
        self.check("amount has no currency", **{"event.amount": 1000})

    def test_value_not_in_dropdown(self):
        self.check("not one of the allowed values", **{"site_process.status": "running"})

    def test_bad_cik(self):
        self.check("SEC CIK", **{"organization.sec_cik": "12AB"})


class TestDuplicates(unittest.TestCase):
    def test_duplicate_epa_id(self):
        path = good_scenario(blank_workbook())
        wb = load_workbook(path)
        put(wb["site"], "site", site_id="SITE-0009", org_id="ORG-0001", state="NV", epa_id="NVR000000001",
            source_id="SRC-0001")
        put(wb["permit"], "permit", site_id="SITE-0009", regulator="US EPA", permit_type="rcra_epa_id",
            number="x", date_checked=TODAY, source_id="SRC-0001")
        wb["site"].cell(row=2 + 1, column=12, value="NVR000000001")  # also on SITE-0001
        wb.save(path)
        errors, _ = run(path)
        assert_flagged(self, errors, "same EPA ID")

    def test_same_website_and_nearby_site(self):
        path = good_scenario(blank_workbook())
        wb = load_workbook(path)
        put(wb["organization"], "organization", org_id="ORG-0009", name="[SYNTH] Something else",
            website="http://www.recycler.example/about", primary_type="other", source_id="SRC-0001")
        put(wb["site"], "site", site_id="SITE-0009", org_id="ORG-0001", state="NV", latitude=39.5301,
            longitude=-119.8101, source_id="SRC-0001")
        wb.save(path)
        _, warnings = run(path)
        assert_flagged(self, warnings, "same website")
        assert_flagged(self, warnings, "possible duplicate site")


if __name__ == "__main__":
    unittest.main()
