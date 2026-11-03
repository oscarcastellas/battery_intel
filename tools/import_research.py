"""Load desk-research JSON files into a fresh workbook.

Run:  python tools/import_research.py OUT.xlsx research/raw/group_a.json research/raw/group_b.json ...

Each JSON file follows research/FORMAT.md (organizations, sites, site_process, accepts, outputs,
events, permits, sources, unresolved). This script:
- builds a fresh workbook from the template and fills it;
- assigns IDs (SRC-, ORG-, SITE-, EVT-) and merges the same source URL / organization across files;
- never guesses: a value that isn't allowed, or a row without a source, is left out and listed in the report;
- marks every acceptance row confirmed_by_party = no (desk research is never confirmation);
- writes a Markdown report next to the workbook (same name, .md) listing everything it skipped or changed.
"""

import json
import re
import sys
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "templates"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_template as bt  # noqa: E402
from check_workbook import domain, normalise_name  # noqa: E402


class Importer:
    def __init__(self):
        self.rows = {t: [] for t in bt.TABLES}
        self.problems = []      # things left out or changed
        self.unresolved = []    # leads and conflicts the researchers reported
        self.src_by_url = {}    # url -> SRC id
        self.org_by_name = {}   # normalised name -> ORG id
        self.org_by_domain = {}
        self.site_by_epa = {}   # EPA ID -> site row
        self.site_by_addr = {}  # (normalised address, city, state) -> site row
        self.counters = {p: 0 for p in bt.ID_PREFIX}

    def next_id(self, table):
        self.counters[table] += 1
        return f"{bt.ID_PREFIX[table]}{self.counters[table]:04d}"

    # -- value cleaning ---------------------------------------------------
    def enum(self, value, list_name, where):
        if value in (None, ""):
            return None
        if value in bt.LISTS[list_name]:
            return value
        self.problems.append(f"{where}: '{value}' is not an allowed {list_name}; left blank")
        return None

    def as_date(self, value, where):
        if value in (None, ""):
            return None
        try:
            return date.fromisoformat(str(value)[:10])
        except ValueError:
            self.problems.append(f"{where}: date '{value}' not understood; left blank")
            return None

    def number(self, value, where):
        if value in (None, ""):
            return None
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return value
        self.problems.append(f"{where}: '{value}' is not a number; left blank")
        return None

    # -- loading -------------------------------------------------------------
    def load(self, path):
        data = json.loads(Path(path).read_text())
        tag = Path(path).stem
        src = {}      # file's source key -> SRC id
        src_date = {}  # SRC id -> retrieval date
        orgs, sites = {}, {}

        for s in data.get("sources", []):
            url = (s.get("url") or "").strip()
            if not url:
                self.problems.append(f"{tag} source {s.get('key')}: no URL; rows citing it are skipped")
                continue
            if url in self.src_by_url:
                src[s["key"]] = self.src_by_url[url]
                continue
            sid = self.next_id("source")
            retrieved = self.as_date(s.get("date_retrieved"), f"{tag} source {s.get('key')}")
            if retrieved is None:
                retrieved = date.today()
                self.problems.append(f"{tag} source {s.get('key')}: no retrieval date; set to the import date")
            self.rows["source"].append({
                "source_id": sid, "source_type": self.enum(s.get("source_type"), "source_type", f"{tag} {s.get('key')}") or "other",
                "title": s.get("title"), "url_or_reference": url, "publisher": s.get("publisher"),
                "date_published": self.as_date(s.get("date_published"), f"{tag} source {s.get('key')}"),
                "date_retrieved": retrieved, "notes": s.get("notes")})
            self.src_by_url[url] = sid
            src[s["key"]] = sid
            src_date[sid] = retrieved
        for sid, d in ((r["source_id"], r["date_retrieved"]) for r in self.rows["source"]):
            src_date.setdefault(sid, d)

        def source_of(item, where):
            sid = src.get(item.get("source"))
            if not sid:
                self.problems.append(f"{where}: no valid source; row skipped")
            return sid

        for o in data.get("organizations", []):
            where = f"{tag} organization {o.get('key')}"
            sid = source_of(o, where)
            if not sid:
                continue
            norm = normalise_name(o.get("name", ""))
            dom = domain(o["website"]) if o.get("website") else None
            existing = self.org_by_name.get(norm) or (self.org_by_domain.get(dom) if dom else None)
            if existing:
                # same organization researched twice: fill blank fields only, never overwrite
                row = next(r for r in self.rows["organization"] if r["org_id"] == existing)
                extra = {"website": o.get("website"), "hq_state": self.enum(o.get("hq_state"), "state", where),
                         "sec_cik": re.sub(r"\D", "", str(o["sec_cik"])) or None if o.get("sec_cik") else None}
                for k, v in extra.items():
                    if row.get(k) in (None, "") and v:
                        row[k] = v
                orgs[o["key"]] = existing
                continue
            oid = self.next_id("organization")
            self.rows["organization"].append({
                "org_id": oid, "name": o.get("name"), "website": o.get("website"),
                "hq_state": self.enum(o.get("hq_state"), "state", where),
                "ownership": self.enum(o.get("ownership"), "ownership", where),
                "parent_name": o.get("parent"),  # resolved after all files are loaded
                "sec_cik": re.sub(r"\D", "", str(o["sec_cik"])) or None if o.get("sec_cik") else None,
                "primary_type": self.enum(o.get("primary_type"), "org_type", where) or "other",
                "source_id": sid, "date_checked": src_date.get(sid), "notes": o.get("notes")})
            orgs[o["key"]] = oid
            self.org_by_name[norm] = oid
            if dom:
                self.org_by_domain[dom] = oid

        for s in data.get("sites", []):
            where = f"{tag} site {s.get('key')}"
            sid = source_of(s, where)
            oid = orgs.get(s.get("org"))
            if not sid:
                continue
            if not oid:
                self.problems.append(f"{where}: organization '{s.get('org')}' not found; row skipped")
                continue
            lat, lon = self.number(s.get("latitude"), where), self.number(s.get("longitude"), where)
            if (lat is None) != (lon is None):
                self.problems.append(f"{where}: only one coordinate given; both left blank")
                lat = lon = None
            zip_code = str(s["zip"]).strip() if s.get("zip") else None
            if zip_code and not re.fullmatch(r"\d{5}(-\d{4})?", zip_code):
                self.problems.append(f"{where}: ZIP '{zip_code}' not valid; left blank")
                zip_code = None
            row = {
                "site_id": None, "org_id": oid, "site_name": s.get("site_name"), "address": s.get("address"),
                "city": s.get("city"), "state": self.enum(s.get("state"), "state", where), "zip": zip_code,
                "county": s.get("county"), "latitude": lat, "longitude": lon, "epa_id": s.get("epa_id"),
                "source_id": sid, "date_checked": src_date.get(sid), "notes": s.get("notes")}
            # The same physical site researched twice (e.g. a plant that changed owner) is merged:
            # same EPA ID, or same address, city and state.
            addr_key = (normalise_name(row["address"]), str(row["city"] or "").lower(), row["state"]) \
                if row["address"] else None
            existing = self.site_by_epa.get(row["epa_id"]) if row["epa_id"] else None
            existing = existing or (self.site_by_addr.get(addr_key) if addr_key else None)
            if existing:
                for k, v in row.items():   # fill gaps only; never overwrite what is already there
                    if existing.get(k) in (None, "") and v not in (None, "") and k != "site_id":
                        existing[k] = v
                if row["org_id"] != existing["org_id"]:
                    existing["notes"] = f"{existing.get('notes') or ''} Also researched under another " \
                                        f"organization ({where}); check current ownership.".strip()
                self.problems.append(f"{where}: same site as {existing['site_id']} (EPA ID or address); merged")
                sites[s["key"]] = existing["site_id"]
                continue
            row["site_id"] = self.next_id("site")
            self.rows["site"].append(row)
            if row["epa_id"]:
                self.site_by_epa[row["epa_id"]] = row
            if addr_key:
                self.site_by_addr[addr_key] = row
            sites[s["key"]] = row["site_id"]

        def site_of(item, where):
            site_id = sites.get(item.get("site"))
            if not site_id:
                self.problems.append(f"{where}: site '{item.get('site')}' not found; row skipped")
            return site_id

        for p in data.get("site_process", []):
            where = f"{tag} site_process {p.get('site')}/{p.get('process_stage')}"
            sid, site_id = source_of(p, where), site_of(p, where)
            if not (sid and site_id):
                continue
            row = {
                "site_id": site_id, "process_stage": self.enum(p.get("process_stage"), "process_stage", where),
                "status": self.enum(p.get("status"), "status", where) or "unknown",
                "status_as_of": self.as_date(p.get("status_as_of"), where) or src_date.get(sid),
                "capacity": self.number(p.get("capacity"), where),
                "capacity_unit": self.enum(p.get("capacity_unit"), "capacity_unit", where),
                "capacity_measure": self.enum(p.get("capacity_measure"), "capacity_measure", where),
                "capacity_basis": self.enum(p.get("capacity_basis"), "capacity_basis", where),
                "expected_start": self.as_date(p.get("expected_start"), where),
                "source_id": sid, "date_checked": src_date.get(sid), "notes": p.get("notes")}
            if row["process_stage"] is None:
                self.problems.append(f"{where}: no valid process_stage; row skipped")
                continue
            if row["capacity"] is not None and None in (row["capacity_unit"], row["capacity_measure"], row["capacity_basis"]):
                # keep the figure only in notes: a number without unit/measure/basis is not usable
                row["notes"] = f"{row['notes'] or ''} [stated capacity {row['capacity']} {p.get('capacity_unit') or ''} "\
                               f"moved to notes: unit/measure/basis incomplete]".strip()
                self.problems.append(f"{where}: capacity without full unit/measure/basis; moved to notes")
                row["capacity"] = row["capacity_unit"] = row["capacity_measure"] = row["capacity_basis"] = None
            self.rows["site_process"].append(row)

        for a in data.get("accepts", []):
            where = f"{tag} accepts {a.get('site')}"
            sid, site_id = source_of(a, where), site_of(a, where)
            if not (sid and site_id):
                continue
            self.rows["facility_accepts"].append({
                "site_id": site_id, "feed_form": self.enum(a.get("feed_form"), "feed_form", where) or "unspecified",
                "chemistry": self.enum(a.get("chemistry"), "chemistry", where) or "unspecified",
                "acceptance_conditions": a.get("acceptance_conditions"),
                "basis": self.enum(a.get("basis"), "accepts_basis", where) or "other",
                "confirmed_by_party": "no", "source_id": sid, "date_checked": src_date.get(sid), "notes": a.get("notes")})

        for o in data.get("outputs", []):
            where = f"{tag} output {o.get('site')}"
            sid, site_id = source_of(o, where), site_of(o, where)
            product = self.enum(o.get("product"), "product", where)
            if sid and site_id and product:
                self.rows["site_output"].append({"site_id": site_id, "product": product, "source_id": sid,
                                                 "notes": o.get("notes")})

        for e in data.get("events", []):
            where = f"{tag} event {e.get('org')}/{e.get('event_type')}"
            sid, oid = source_of(e, where), orgs.get(e.get("org"))
            if not (sid and oid):
                if sid:
                    self.problems.append(f"{where}: organization not found; row skipped")
                continue
            event_date = self.as_date(e.get("event_date"), where)
            if event_date is None:
                self.problems.append(f"{where}: no event date given; row skipped (summary: {e.get('summary')})")
                continue
            cp = e.get("counterparty")
            cp_id = orgs.get(cp) or self.org_by_name.get(normalise_name(cp or "")) if cp else None
            amount = self.number(e.get("amount"), where)
            self.rows["event"].append({
                "event_id": self.next_id("event"), "org_id": oid, "counterparty_org_id": cp_id,
                "counterparty_name": None if cp_id else cp, "site_id": sites.get(e.get("site")),
                "event_type": self.enum(e.get("event_type"), "event_type", where) or "other",
                "event_date": event_date, "amount": amount,
                "currency": (e.get("currency") or ("USD" if amount is not None else None)),
                "summary": e.get("summary"), "source_id": sid})

        for p in data.get("permits", []):
            where = f"{tag} permit {p.get('site')}"
            sid, site_id = source_of(p, where), site_of(p, where)
            if not (sid and site_id and p.get("number")):
                continue
            self.rows["permit"].append({
                "site_id": site_id, "regulator": p.get("regulator"),
                "permit_type": self.enum(p.get("permit_type"), "permit_type", where) or "other",
                "number": p.get("number"), "status": p.get("status"), "date_checked": src_date.get(sid),
                "source_id": sid, "notes": None})

        self.unresolved += [f"{tag}: {u}" for u in data.get("unresolved", [])]

    def resolve_parents(self):
        for org in self.rows["organization"]:
            parent = org.pop("parent_name", None)
            if parent:
                pid = self.org_by_name.get(normalise_name(parent))
                if pid and pid != org["org_id"]:
                    org["parent_org_id"] = pid
                else:
                    org["notes"] = f"{org.get('notes') or ''} Parent: {parent}.".strip()

    def write(self, out):
        bt.build(out)
        wb = load_workbook(out)
        for table, rows in self.rows.items():
            ws = wb[table]
            names = [c[0] for c in bt.TABLES[table]]
            start = 3 if table in bt.EXAMPLES else 2   # keep the example row
            for i, rec in enumerate(rows):
                for name, v in rec.items():
                    if name in names and v not in (None, ""):
                        ws.cell(row=start + i, column=names.index(name) + 1, value=v)
        wb.save(out)

    def report(self, out, files):
        counts = ", ".join(f"{t} {len(r)}" for t, r in self.rows.items() if r)
        lines = [f"# Import report: {Path(out).name}", "", f"Imported {date.today()} from: "
                 + ", ".join(Path(f).name for f in files), "", f"**Rows:** {counts}", "",
                 "All acceptance rows are `confirmed_by_party = no` (desk research is not confirmation).", "",
                 f"## Left out or changed ({len(self.problems)})", ""]
        lines += [f"- {p}" for p in self.problems] or ["- none"]
        lines += ["", f"## Unresolved leads and conflicts reported by the researchers ({len(self.unresolved)})", ""]
        lines += [f"- {u}" for u in self.unresolved] or ["- none"]
        Path(out).with_suffix(".md").write_text("\n".join(lines) + "\n")


def main():
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    out, files = Path(sys.argv[1]), sys.argv[2:]
    imp = Importer()
    for f in files:
        imp.load(f)
    imp.resolve_parents()
    imp.write(out)
    imp.report(out, files)
    print(f"Wrote {out} and {out.with_suffix('.md')}")


if __name__ == "__main__":
    main()
