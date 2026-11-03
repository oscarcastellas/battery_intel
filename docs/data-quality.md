# Data quality

The goal: accurate enough to put a real lot of battery feedstock on a truck to the right processor.

## Evidence, strongest first
1. **The facility itself confirms** (quote, call, email). Only this sets `confirmed_by_party = yes`.
2. **Official records:** SEC, DOE, state regulator, EPA, court.
3. **The company's own published material:** a claim, not a confirmation.
4. **Press.**
5. **Inference:** never stored as a fact.

## Freshness
| Fact | Warning after | Before a deal proposal |
|---|---|---|
| Site status (`status_as_of`) | 90 days | operating, confirmed within 30 days |
| Price (no expiry stated) | 60 days | must be current |
| Permits | 180 days | (deal checklist) |
| Acceptance, contacts, sites, organizations | 180 days | facility must have confirmed the feed form + chemistry |

## The checker (`python tools/check_workbook.py <workbook>`)
- **Structure:** headers unchanged; IDs unique with the right prefix; references point to real (non-example) rows.
- **Values:** required fields; dropdown values; dates; non-negative numbers; US coordinates (latitude 17–72, negative longitude); 5-digit ZIPs; SEC CIK up to 10 digits.
- **Duplicates:**
  - error: same EPA ID on two sites;
  - warnings: same website, near-identical names, same address, sites within 50 m.
- **Battery rules:**
  - capacity is an error without unit, measure and basis;
  - status older than 90 days is a warning, and a status date in the future is an error;
  - several rows for the same site and stage must each come from a different source, and disagreements are noted;
  - acceptance at a site with no process row is a warning;
  - `confirmed_by_party = yes` must come from a quote, call or email.
- **Prices:**
  - a payable must be "% of metal value" with a metal, and at most 100%;
  - a purchase price or fee can't use the percentage unit;
  - expired or stale prices are flagged.
- **Feed lots:** a lot can't be quoting or placed while its chemistry or condition is unknown.
- **Deal gate** (proposal_sent or later): the deal is an **error** unless:
  - lot chemistry and condition are known;
  - the receiving site was confirmed operating within 30 days;
  - the site confirmed it accepts this feed form and chemistry, and also damaged/defective batteries if the lot is damaged or recalled;
  - a current price from that site is linked;
  - the transporter has a USDOT number recorded;
  - every checklist item is `yes`.
- **Coverage:** sites by stage and status; operating capacity summed only like-for-like (same unit and measure); published vs confirmed acceptance; audit error rate.

The Postgres schema (`sql/0001_schema.sql`) enforces the core rules too.

## Monthly spot-check
`--audit-sample 10` lists 10 random records with their source links. Re-check each one and record the result in `audit_log`. **Target error rate: under 5%.**

## Tests
`python -m unittest discover -s tests` (standard library + openpyxl).
