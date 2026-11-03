"""Check the battery workbook before its data is used.

Run:  python tools/check_workbook.py research/battery_intel_seed.xlsx
      python tools/check_workbook.py my_copy.xlsx --audit-sample 10

Prints ERRORS (must fix), WARNINGS (check), notes and a coverage summary.
Exit code 1 if there are errors. Rows whose first value starts with EX- are template examples and are skipped.
The checks are listed in docs/data-quality.md. Adapted from the waste_model framework.
"""

import argparse
import math
import random
import re
import sys
from collections import defaultdict
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "templates"))
import build_template as bt  # noqa: E402  (TABLES, LISTS, ID_PREFIX)

# Freshness rules (days)
STATUS_MAX_AGE = 90         # site status: the market moves fast (pauses, sales, closures)
DEAL_STATUS_MAX_AGE = 30    # receiving site's status within 30 days of a proposal
PRICE_MAX_AGE = 60          # prices without an expiry are stale after 60 days
FACT_MAX_AGE = 180          # contacts, acceptance, sites, organizations
PERMIT_MAX_AGE = 180

DUPLICATE_NAME_RATIO = 0.9
DUPLICATE_SITE_METRES = 50
EPA_ID_PATTERN = r"[A-Z]{2}[A-Z0-9]{10}"   # 12 characters, starting with a state code (format to be verified)

GATED_STAGES = ["proposal_sent", "accepted", "scheduled", "completed", "invoiced", "paid"]
INACTIVE = {"paused", "cancelled", "closed"}

REQUIRED = {
    "source": ["source_id", "source_type", "date_retrieved"],
    "organization": ["org_id", "name", "primary_type", "source_id"],
    "site": ["site_id", "org_id", "state", "source_id"],
    "site_process": ["site_id", "process_stage", "status", "status_as_of", "source_id"],
    "facility_accepts": ["site_id", "feed_form", "chemistry", "basis", "source_id"],
    "site_output": ["site_id", "product", "source_id"],
    "event": ["event_id", "org_id", "event_type", "event_date", "summary", "source_id"],
    "permit": ["site_id", "regulator", "permit_type", "number", "date_checked", "source_id"],
    "contact": ["contact_id", "org_id", "name", "how_obtained", "source_id", "date_checked"],
    "price_observation": ["price_id", "site_id", "feed_form", "chemistry", "price_type", "amount", "price_unit",
                          "valid_from", "basis", "source_id"],
    "feed_lot": ["lot_id", "holder_org_id", "feed_form", "chemistry", "condition", "basis", "status", "source_id",
                 "date_checked"],
    "deal": ["deal_id", "lot_id", "stage", "stage_date"],
    "interaction": ["interaction_id", "interaction_date", "channel", "summary"],
    "audit_log": ["audit_id", "audit_date", "tab", "record_id", "result"],
}

LEGAL_SUFFIXES = r"\b(inc|incorporated|llc|ltd|limited|corp|corporation|co|company|the|holdings|usa|us|america)\b"


class Issue:
    def __init__(self, level, tab, row, message):
        self.level, self.tab, self.row, self.message = level, tab, row, message

    def __str__(self):
        where = f"{self.tab} row {self.row}" if self.row else self.tab
        return f"[{self.level}] {where}: {self.message}"


class Workbook:
    """The workbook's input values as dicts, one per real (non-example, non-empty) row."""

    def __init__(self, path):
        self.wb = load_workbook(path)  # formulas not needed: auto columns are ignored
        self.issues = []
        self.rows = {}
        self.examples = {}
        self.allowed = self._read_lists()
        for tab in bt.TABLES:
            self._read_tab(tab)

    def add(self, level, tab, row, message):
        self.issues.append(Issue(level, tab, row, message))

    def _read_lists(self):
        allowed = {name: set(values) for name, values in bt.LISTS.items()}
        if "lists" in self.wb.sheetnames:
            for col in self.wb["lists"].iter_cols(min_row=1, values_only=True):
                name, values = col[0], [v for v in col[1:] if v not in (None, "")]
                if name:
                    allowed[name] = set(str(v).strip() for v in values)
        return allowed

    def _read_tab(self, tab):
        columns = bt.TABLES[tab]
        self.rows[tab], self.examples[tab] = [], set()
        if tab not in self.wb.sheetnames:
            self.add("ERROR", tab, None, "tab is missing")
            return
        ws = self.wb[tab]
        header = [c.value for c in ws[1][:len(columns)]]
        expected = [c[0] for c in columns]
        if header != expected:
            self.add("ERROR", tab, 1, f"column headers changed; expected {expected}")
            return
        for r_idx, values in enumerate(ws.iter_rows(min_row=2, max_col=len(columns), values_only=True), start=2):
            record = {}
            for (name, _, kind, _), v in zip(columns, values):
                if kind.startswith("auto:"):
                    continue
                if isinstance(v, str):
                    v = v.strip() or None
                record[name] = v
            if all(v is None for v in record.values()):
                continue
            first = record[expected[0]]
            if isinstance(first, str) and first.startswith("EX-"):
                self.examples[tab].add(first)
                continue
            self.rows[tab].append((r_idx, record))

    def ids(self, tab):
        key = bt.TABLES[tab][0][0]
        return {rec[key] for _, rec in self.rows[tab] if rec.get(key)}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def as_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return date.fromisoformat(str(v))


def safe_date(v):
    try:
        return as_date(v)
    except ValueError:
        return None  # reported by check_fields


def age_days(d, today):
    return (today - d).days if d else None


def normalise_name(name):
    s = re.sub(r"[^a-z0-9 ]", " ", str(name).lower())
    s = re.sub(LEGAL_SUFFIXES, " ", s)
    return " ".join(s.split())


def domain(url):
    s = re.sub(r"^https?://", "", str(url).lower().strip())
    s = s.split("/")[0]
    return s[4:] if s.startswith("www.") else s


def metres_between(lat1, lon1, lat2, lon2):
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = p2 - p1, math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * 6371000 * math.asin(math.sqrt(a))


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ---------------------------------------------------------------------------
# checks
# ---------------------------------------------------------------------------
def check_fields(wb):
    """Required values, dropdowns, types, ranges; IDs unique with the right prefix; references resolve."""
    all_ids = {tab: wb.ids(tab) for tab in bt.TABLES}
    for tab, columns in bt.TABLES.items():
        key = columns[0][0]
        seen = {}
        for row, rec in wb.rows[tab]:
            for col in REQUIRED.get(tab, []):
                if rec.get(col) is None:
                    wb.add("ERROR", tab, row, f"{col} is required")
            for name, _, kind, _ in columns:
                v = rec.get(name)
                if v is None or kind.startswith("auto:"):
                    continue
                if kind.startswith("list:"):
                    if str(v) not in wb.allowed.get(kind[5:], set()):
                        wb.add("ERROR", tab, row, f"{name} '{v}' is not one of the allowed values")
                elif kind.startswith("ref:"):
                    target = kind[4:]
                    if v in wb.examples.get(target, set()):
                        wb.add("ERROR", tab, row, f"{name} '{v}' points to an example row; use a real ID")
                    elif v not in all_ids[target]:
                        wb.add("ERROR", tab, row, f"{name} '{v}' does not exist in the {target} tab")
                elif kind == "date":
                    try:
                        as_date(v)
                    except ValueError:
                        wb.add("ERROR", tab, row, f"{name} '{v}' is not a date (use yyyy-mm-dd)")
                elif kind == "number":
                    if not is_num(v):
                        wb.add("ERROR", tab, row, f"{name} '{v}' is not a number")
                    elif v < 0:
                        wb.add("ERROR", tab, row, f"{name} is negative")
                elif kind == "lat" and not (is_num(v) and 17 <= v <= 72):
                    wb.add("ERROR", tab, row, f"latitude {v} is outside the US (17 to 72)")
                elif kind == "lon" and not (is_num(v) and -180 <= v <= -64):
                    wb.add("ERROR", tab, row, f"longitude {v} is outside the US (-180 to -64; US longitudes are negative)")
                elif kind == "zip" and not re.fullmatch(r"\d{5}(-\d{4})?", str(v)):
                    wb.add("ERROR", tab, row, f"ZIP '{v}' is not a 5-digit ZIP (check leading zeros)")
                elif kind == "cik" and not re.fullmatch(r"\d{1,10}", str(v)):
                    wb.add("ERROR", tab, row, f"SEC CIK '{v}' should be up to 10 digits")
            if tab in bt.ID_PREFIX and rec.get(key):
                if not str(rec[key]).startswith(bt.ID_PREFIX[tab]):
                    wb.add("ERROR", tab, row, f"{key} '{rec[key]}' should start with {bt.ID_PREFIX[tab]}")
                if rec[key] in seen:
                    wb.add("ERROR", tab, row, f"{key} '{rec[key]}' is used twice (also row {seen[rec[key]]})")
                seen.setdefault(rec[key], row)


def check_duplicates(wb):
    orgs = wb.rows["organization"]
    by_domain = {}
    for row, rec in orgs:
        if rec.get("website"):
            d = domain(rec["website"])
            if d in by_domain:
                wb.add("WARNING", "organization", row, f"same website as row {by_domain[d]}: possible duplicate")
            by_domain.setdefault(d, row)
        if rec.get("parent_org_id") and rec.get("parent_org_id") == rec.get("org_id"):
            wb.add("ERROR", "organization", row, "parent_org_id points to itself")
    names = [(row, normalise_name(rec["name"])) for row, rec in orgs if rec.get("name")]
    for i, (r1, n1) in enumerate(names):
        for r2, n2 in names[i + 1:]:
            if n1 and n2 and SequenceMatcher(None, n1, n2).ratio() >= DUPLICATE_NAME_RATIO:
                wb.add("WARNING", "organization", r2, f"name very similar to row {r1}: possible duplicate")

    sites = wb.rows["site"]
    by_epa = {}
    for row, rec in sites:
        epa = rec.get("epa_id")
        if epa:
            if not re.fullmatch(EPA_ID_PATTERN, str(epa)):
                wb.add("WARNING", "site", row, f"EPA ID '{epa}' doesn't look like a 12-character EPA ID")
            if epa in by_epa:
                wb.add("ERROR", "site", row, f"same EPA ID as row {by_epa[epa]}: the same site entered twice?")
            by_epa.setdefault(epa, row)
    for i, (r1, s1) in enumerate(sites):
        for r2, s2 in sites[i + 1:]:
            if s1.get("address") and s2.get("address") and \
                    normalise_name(s1["address"]) == normalise_name(s2["address"]) and \
                    str(s1.get("city") or "").lower() == str(s2.get("city") or "").lower():
                wb.add("WARNING", "site", r2, f"same address as row {r1}: possible duplicate site")
            elif all(is_num(s.get(k)) for s in (s1, s2) for k in ("latitude", "longitude")):
                m = metres_between(s1["latitude"], s1["longitude"], s2["latitude"], s2["longitude"])
                if m < DUPLICATE_SITE_METRES:
                    wb.add("WARNING", "site", r2, f"within {m:.0f} m of row {r1}: possible duplicate site")


def price_is_current(rec, today):
    vf, vu = safe_date(rec.get("valid_from")), safe_date(rec.get("valid_until"))
    if vu:
        return vu >= today
    return bool(vf) and age_days(vf, today) <= PRICE_MAX_AGE


def check_records(wb, today):
    for row, rec in wb.rows["site"]:
        if rec.get("latitude") is None:
            wb.add("INFO", "site", row, "no coordinates yet: distance matching won't work for this site")
        if (rec.get("latitude") is None) != (rec.get("longitude") is None):
            wb.add("ERROR", "site", row, "latitude and longitude must be given together")

    for tab in ("organization", "site", "contact", "facility_accepts"):
        for row, rec in wb.rows[tab]:
            checked = safe_date(rec.get("date_checked"))
            if checked and age_days(checked, today) > FACT_MAX_AGE:
                wb.add("WARNING", tab, row, f"last checked {checked}, over {FACT_MAX_AGE} days ago")

    processes_by_site = defaultdict(list)
    stage_rows = defaultdict(list)
    for row, rec in wb.rows["site_process"]:
        processes_by_site[rec.get("site_id")].append(rec)
        stage_rows[(rec.get("site_id"), rec.get("process_stage"))].append((row, rec))
        if rec.get("capacity") is not None:
            missing = [c for c in ("capacity_unit", "capacity_measure", "capacity_basis") if rec.get(c) is None]
            if missing:
                wb.add("ERROR", "site_process", row, f"capacity without {', '.join(missing)}")
        as_of = safe_date(rec.get("status_as_of"))
        if as_of and age_days(as_of, today) > STATUS_MAX_AGE and rec.get("status") not in ("cancelled", "closed"):
            wb.add("WARNING", "site_process", row, f"status '{rec.get('status')}' last confirmed {as_of}, "
                   f"over {STATUS_MAX_AGE} days ago")
        if as_of and as_of > today:
            wb.add("ERROR", "site_process", row, f"status_as_of {as_of} is in the future")
    for (site, stage), rows in stage_rows.items():
        sources = {r.get("source_id") for _, r in rows}
        if len(rows) > 1 and len(sources) < len(rows):
            wb.add("WARNING", "site_process", rows[-1][0],
                   f"{len(rows)} rows for {site} / {stage} but not each from a different source")
        elif len(rows) > 1:
            statuses = {r.get("status") for _, r in rows}
            if len(statuses) > 1:
                wb.add("INFO", "site_process", rows[-1][0],
                       f"sources disagree on status for {site} / {stage}: {sorted(statuses)}")

    for row, rec in wb.rows["facility_accepts"]:
        if rec.get("site_id") not in processes_by_site:
            wb.add("WARNING", "facility_accepts", row, "site has no site_process row (what does it do with the feed?)")
        if rec.get("confirmed_by_party") == "yes" and rec.get("basis") in ("company_website", "report", "other"):
            wb.add("WARNING", "facility_accepts", row,
                   "confirmed_by_party is yes but basis is not a direct confirmation (quote, phone_call or email)")

    for row, rec in wb.rows["event"]:
        if rec.get("amount") is not None and not rec.get("currency"):
            wb.add("ERROR", "event", row, "amount has no currency")
        if not rec.get("counterparty_org_id") and not rec.get("counterparty_name") and \
                rec.get("event_type") in ("supply_agreement", "offtake_agreement", "grant", "loan", "acquisition"):
            wb.add("WARNING", "event", row, f"{rec.get('event_type')} with no counterparty recorded")

    for row, rec in wb.rows["permit"]:
        checked = safe_date(rec.get("date_checked"))
        if checked and age_days(checked, today) > PERMIT_MAX_AGE:
            wb.add("WARNING", "permit", row, f"last checked {checked}, over {PERMIT_MAX_AGE} days ago")

    for row, rec in wb.rows["price_observation"]:
        vf, vu = safe_date(rec.get("valid_from")), safe_date(rec.get("valid_until"))
        if vf and vu and vu < vf:
            wb.add("ERROR", "price_observation", row, "valid_until is before valid_from")
        if rec.get("price_type") == "payable_percent":
            if rec.get("price_unit") != "% of metal value" or rec.get("metal") in (None, "n/a"):
                wb.add("ERROR", "price_observation", row, "payable_percent needs unit '% of metal value' and a metal")
            elif is_num(rec.get("amount")) and rec["amount"] > 100:
                wb.add("ERROR", "price_observation", row, "payable over 100%")
        elif rec.get("price_unit") == "% of metal value":
            wb.add("ERROR", "price_observation", row, "'% of metal value' is only for payable_percent")
        if not price_is_current(rec, today):
            wb.add("WARNING", "price_observation", row,
                   f"price is out of date (expired, or older than {PRICE_MAX_AGE} days with no expiry)")

    for row, rec in wb.rows["feed_lot"]:
        if rec.get("quantity") is not None and rec.get("unit") is None:
            wb.add("ERROR", "feed_lot", row, "quantity has no unit")
        ws_, we = safe_date(rec.get("window_start")), safe_date(rec.get("window_end"))
        if ws_ and we and we < ws_:
            wb.add("ERROR", "feed_lot", row, "window_end is before window_start")
        if rec.get("status") in ("quoting", "placed"):
            if rec.get("chemistry") in (None, "unspecified") or rec.get("condition") in (None, "unknown"):
                wb.add("ERROR", "feed_lot", row, f"status is {rec['status']} but chemistry or condition is unknown")

    for row, rec in wb.rows["interaction"]:
        if not rec.get("contact_id") and not rec.get("org_id"):
            wb.add("ERROR", "interaction", row, "needs a contact_id or an org_id")
        fu = safe_date(rec.get("follow_up_date"))
        if fu and fu < today:
            wb.add("INFO", "interaction", row, f"follow-up was due {fu}")


def check_deal_gate(wb, today):
    """A deal at proposal_sent or later needs current, confirmed evidence for every party."""
    lots = {rec["lot_id"]: rec for _, rec in wb.rows["feed_lot"] if rec.get("lot_id")}
    orgs = {rec["org_id"]: rec for _, rec in wb.rows["organization"] if rec.get("org_id")}
    prices = {rec["price_id"]: rec for _, rec in wb.rows["price_observation"] if rec.get("price_id")}

    for row, deal in wb.rows["deal"]:
        if deal.get("stage") not in GATED_STAGES:
            continue
        problems = []
        lot = lots.get(deal.get("lot_id")) or {}
        if not lot:
            problems.append("lot not found")
        if lot.get("chemistry") in (None, "unspecified") or lot.get("condition") in (None, "unknown"):
            problems.append("lot chemistry or condition unknown")

        site = deal.get("receiving_site_id")
        operating = [p for _, p in wb.rows["site_process"] if p.get("site_id") == site and p.get("status") == "operating"
                     and safe_date(p.get("status_as_of")) and age_days(safe_date(p["status_as_of"]), today) <= DEAL_STATUS_MAX_AGE]
        if not operating:
            problems.append(f"receiving site not confirmed operating within {DEAL_STATUS_MAX_AGE} days")
        needed_forms = {lot.get("feed_form")}
        if lot.get("condition") in ("damaged_defective", "recalled"):
            needed_forms.add("damaged_defective_recalled")
        for form in needed_forms:
            ok = any(a.get("site_id") == site and a.get("feed_form") == form
                     and a.get("chemistry") in (lot.get("chemistry"), "all") and a.get("confirmed_by_party") == "yes"
                     for _, a in wb.rows["facility_accepts"])
            if not ok:
                problems.append(f"receiving site has not confirmed it accepts {form} / {lot.get('chemistry')}")

        price = prices.get(deal.get("price_id"))
        if not price or price.get("site_id") != site or not price_is_current(price, today):
            problems.append("no current price from the receiving site")
        transporter = orgs.get(deal.get("transporter_org_id"))
        if not transporter or not transporter.get("usdot_number"):
            problems.append("transporter missing or has no USDOT number recorded")
        for tick in ("chemistry_condition_verified", "receiving_site_confirmed", "transporter_hazmat_verified",
                     "packaging_compliant"):
            if deal.get(tick) != "yes":
                problems.append(f"checklist item {tick} is not 'yes'")
        for p in dict.fromkeys(problems):
            wb.add("ERROR", "deal", row, f"deal gate ({deal.get('stage')}): {p}")


def coverage_summary(wb, today):
    lines = []
    status_count = defaultdict(int)
    for _, p in wb.rows["site_process"]:
        status_count[(p.get("process_stage"), p.get("status"))] += 1
    stages = [s for s in bt.LISTS["process_stage"] if any(k[0] == s for k in status_count)]
    for stage in stages:
        parts = [f"{st} {status_count[(stage, st)]}" for st in bt.LISTS["status"] if status_count[(stage, st)]]
        lines.append(f"  {stage}: " + ", ".join(parts))
    cap = defaultdict(float)
    for _, p in wb.rows["site_process"]:
        if p.get("status") == "operating" and is_num(p.get("capacity")):
            cap[(p["process_stage"], p.get("capacity_unit"), p.get("capacity_measure"))] += p["capacity"]
    for (stage, unit, measure), total in sorted(cap.items(), key=lambda kv: str(kv[0])):
        lines.append(f"  operating capacity, {stage}: {total:,.0f} {unit} ({measure}); sums only like-for-like rows, "
                     "and overlapping sources may double count")
    sites_with_accepts = {a.get("site_id") for _, a in wb.rows["facility_accepts"]}
    confirmed = {a.get("site_id") for _, a in wb.rows["facility_accepts"] if a.get("confirmed_by_party") == "yes"}
    lines.append(f"  sites with published acceptance: {len(sites_with_accepts)}; confirmed by the site: {len(confirmed)}")
    recent = [a for _, a in wb.rows["audit_log"]
              if safe_date(a.get("audit_date")) and age_days(safe_date(a.get("audit_date")), today) <= 90]
    checked = [a for a in recent if a.get("result") in ("correct", "incorrect", "changed")]
    if checked:
        wrong = sum(1 for a in checked if a.get("result") in ("incorrect", "changed"))
        lines.append(f"  audit (last 90 days): {len(checked)} checked, {wrong} wrong or changed "
                     f"({100 * wrong / len(checked):.0f}% error rate)")
    else:
        lines.append("  audit (last 90 days): no spot-checks recorded yet")
    return lines


def audit_sample(wb, n, seed=None):
    rng = random.Random(seed)
    sources = {rec["source_id"]: rec for _, rec in wb.rows["source"] if rec.get("source_id")}
    pool = [(tab, row, rec) for tab in ("organization", "site", "site_process", "facility_accepts", "site_output",
                                        "event", "permit", "price_observation") for row, rec in wb.rows[tab]]
    lines = []
    for tab, row, rec in rng.sample(pool, min(n, len(pool))):
        shown = {k: v for k, v in rec.items() if v is not None and k not in ("notes", "source_id")}
        src = sources.get(rec.get("source_id"), {})
        lines.append(f"  {tab} row {row}: {shown}\n      source {rec.get('source_id')}: {src.get('url_or_reference')}")
    return lines


def check(path, today=None):
    today = today or date.today()
    wb = Workbook(path)
    check_fields(wb)
    check_duplicates(wb)
    check_records(wb, today)
    check_deal_gate(wb, today)
    return wb, coverage_summary(wb, today)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("workbook")
    parser.add_argument("--today", help="pretend today is this date (yyyy-mm-dd)")
    parser.add_argument("--audit-sample", type=int, default=0, help="print N random records to re-verify")
    parser.add_argument("--seed", type=int)
    parser.add_argument("--quiet", action="store_true", help="hide INFO notes")
    args = parser.parse_args()

    today = date.fromisoformat(args.today) if args.today else date.today()
    wb, summary = check(args.workbook, today)
    order = {"ERROR": 0, "WARNING": 1, "INFO": 2}
    issues = sorted(wb.issues, key=lambda i: (order[i.level], i.tab, i.row or 0))
    counts = {lvl: sum(1 for i in issues if i.level == lvl) for lvl in order}

    print(f"Checked {args.workbook} as of {today}")
    print("  rows: " + ", ".join(f"{t} {len(r)}" for t, r in wb.rows.items() if r))
    for issue in issues:
        if not (args.quiet and issue.level == "INFO"):
            print(issue)
    print(f"\n{counts['ERROR']} errors, {counts['WARNING']} warnings, {counts['INFO']} notes")
    print("\nCoverage:")
    print("\n".join(summary))
    if args.audit_sample:
        print("\nAudit sample: re-check these against their sources and record results in audit_log:")
        print("\n".join(audit_sample(wb, args.audit_sample, args.seed)))
    sys.exit(1 if counts["ERROR"] else 0)


if __name__ == "__main__":
    main()
