-- =============================================================================
-- Battery Intel: schema v1 (matches templates/battery_intel_v1.xlsx)
-- Status: DRAFT, not yet applied. Target: PostgreSQL 15+ (no extensions needed).
--
-- Table and column names are identical to the workbook tabs and headers, and CHECK
-- lists mirror the workbook dropdowns (tests/test_schema_matches_workbook.py enforces both).
-- Columns the workbook fills automatically (org_name, site_name, ...) are not stored.
-- IDs are the text IDs used in the workbook (SRC-0001, ORG-0001, SITE-0001, ...).
-- =============================================================================

CREATE SCHEMA IF NOT EXISTS bi;
SET search_path = bi;

CREATE TABLE source (
    source_id        text PRIMARY KEY CHECK (source_id LIKE 'SRC-%'),
    source_type      text NOT NULL CHECK (source_type IN
                       ('company_website','press_release','sec_filing','doe_announcement','state_permit','epa_echo',
                        'court_filing','news','government_report','phone_call','email','meeting','other')),
    title            text,
    url_or_reference text,
    publisher        text,
    date_published   date,
    date_retrieved   date NOT NULL,
    notes            text
);

CREATE TABLE organization (
    org_id        text PRIMARY KEY CHECK (org_id LIKE 'ORG-%'),
    name          text NOT NULL,
    website       text,
    hq_state      text CHECK (hq_state IN ('AL','AK','AZ','AR','CA','CO','CT','DE','DC','FL','GA','HI','ID','IL','IN',
                    'IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND',
                    'OH','OK','OR','PA','RI','SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY','PR','GU','VI','AS','MP')),
    ownership     text CHECK (ownership IN ('private','public','subsidiary','joint_venture','unknown')),
    parent_org_id text REFERENCES organization,
    sec_cik       text CHECK (sec_cik ~ '^[0-9]{1,10}$'),
    primary_type  text NOT NULL CHECK (primary_type IN
                    ('preprocessor','refiner','integrated','oem','cell_manufacturer','collector','logistics','transporter','other')),
    other_types   text,
    usdot_number  text,
    source_id     text NOT NULL REFERENCES source,
    date_checked  date,
    notes         text,
    CHECK (parent_org_id IS NULL OR parent_org_id <> org_id)
);

CREATE TABLE site (
    site_id      text PRIMARY KEY CHECK (site_id LIKE 'SITE-%'),
    org_id       text NOT NULL REFERENCES organization,
    site_name    text,
    address      text,
    city         text,
    state        text NOT NULL CHECK (state IN ('AL','AK','AZ','AR','CA','CO','CT','DE','DC','FL','GA','HI','ID','IL',
                   'IN','IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC',
                   'ND','OH','OK','OR','PA','RI','SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY','PR','GU','VI',
                   'AS','MP')),
    zip          text CHECK (zip ~ '^[0-9]{5}(-[0-9]{4})?$'),
    county       text,
    latitude     numeric(9,6) CHECK (latitude BETWEEN 17 AND 72),
    longitude    numeric(9,6) CHECK (longitude BETWEEN -180 AND -64),
    epa_id       text,
    source_id    text NOT NULL REFERENCES source,
    date_checked date,
    notes        text,
    CHECK ((latitude IS NULL) = (longitude IS NULL))
);
CREATE INDEX ON site (org_id);
CREATE UNIQUE INDEX site_epa_id_uq ON site (epa_id) WHERE epa_id IS NOT NULL;

-- What a site does, its dated status and its capacity. Conflicting sources = separate rows.
CREATE TABLE site_process (
    site_process_id  bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    site_id          text NOT NULL REFERENCES site,
    process_stage    text NOT NULL CHECK (process_stage IN
                       ('collection_sorting','discharge','dismantling','mechanical_preprocessing','hydrometallurgical',
                        'pyrometallurgical','direct_recycling','precursor_production')),
    status           text NOT NULL CHECK (status IN
                       ('announced','permitting','under_construction','commissioning','operating','paused','cancelled',
                        'closed','unknown')),
    status_as_of     date NOT NULL,
    capacity         numeric CHECK (capacity >= 0),
    capacity_unit    text CHECK (capacity_unit IN ('t/yr','kt/yr','GWh/yr','lb/yr','other')),
    capacity_measure text CHECK (capacity_measure IN
                       ('input_batteries','input_black_mass','output_black_mass','output_metal_content','output_product',
                        'unspecified')),
    capacity_basis   text CHECK (capacity_basis IN ('nameplate_stated','planned','reported_throughput','unspecified')),
    expected_start   date,
    source_id        text NOT NULL REFERENCES source,
    date_checked     date,
    notes            text,
    -- a capacity figure is only usable with its unit, what it measures, and its basis
    CHECK (capacity IS NULL OR (capacity_unit IS NOT NULL AND capacity_measure IS NOT NULL AND capacity_basis IS NOT NULL)),
    UNIQUE NULLS NOT DISTINCT (site_id, process_stage, source_id, capacity_measure)
);

CREATE TABLE facility_accepts (
    facility_accepts_id   bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    site_id               text NOT NULL REFERENCES site,
    feed_form             text NOT NULL CHECK (feed_form IN
                            ('manufacturing_scrap','cells','modules','ev_packs','ess_packs','consumer_batteries',
                             'black_mass','damaged_defective_recalled','unspecified')),
    chemistry             text NOT NULL CHECK (chemistry IN ('NMC','NCA','LFP','LCO','LMO','mixed','all','unspecified')),
    acceptance_conditions text,
    basis                 text NOT NULL CHECK (basis IN ('company_website','report','quote','phone_call','email','other')),
    confirmed_by_party    text NOT NULL DEFAULT 'no' CHECK (confirmed_by_party IN ('yes','no')),
    source_id             text NOT NULL REFERENCES source,
    date_checked          date,
    notes                 text,
    -- only the facility itself can confirm
    CHECK (confirmed_by_party = 'no' OR basis IN ('quote','phone_call','email'))
);

CREATE TABLE site_output (
    site_output_id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    site_id        text NOT NULL REFERENCES site,
    product        text NOT NULL CHECK (product IN
                     ('black_mass','MHP','nickel_sulfate','cobalt_sulfate','manganese_sulfate','lithium_carbonate',
                      'lithium_hydroxide','copper','aluminum','graphite','alloy_matte','pCAM','CAM','other')),
    source_id      text NOT NULL REFERENCES source,
    notes          text
);

CREATE TABLE event (
    event_id            text PRIMARY KEY CHECK (event_id LIKE 'EVT-%'),
    org_id              text NOT NULL REFERENCES organization,
    counterparty_org_id text REFERENCES organization,
    counterparty_name   text,
    site_id             text REFERENCES site,
    event_type          text NOT NULL CHECK (event_type IN
                          ('supply_agreement','offtake_agreement','grant','loan','investment','acquisition','bankruptcy',
                           'pause','closure','restart','permit_issued','enforcement','other')),
    event_date          date NOT NULL,
    amount              numeric CHECK (amount >= 0),
    currency            text,
    summary             text NOT NULL,
    source_id           text NOT NULL REFERENCES source,
    CHECK (amount IS NULL OR currency IS NOT NULL)
);

CREATE TABLE permit (
    permit_id    bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    site_id      text NOT NULL REFERENCES site,
    regulator    text NOT NULL,
    permit_type  text NOT NULL CHECK (permit_type IN
                   ('rcra_epa_id','rcra_permit','universal_waste_handler','air_permit','state_solid_waste','other')),
    number       text NOT NULL,
    status       text,
    date_checked date NOT NULL,
    source_id    text NOT NULL REFERENCES source,
    notes        text
);

CREATE TABLE contact (
    contact_id   text PRIMARY KEY CHECK (contact_id LIKE 'CON-%'),
    org_id       text NOT NULL REFERENCES organization,
    site_id      text REFERENCES site,
    name         text NOT NULL,
    role         text,
    work_phone   text,
    work_email   text,
    how_obtained text NOT NULL CHECK (how_obtained IN
                   ('company_website','direct_from_person','business_card','referral','public_register','other')),
    source_id    text REFERENCES source,
    date_checked date NOT NULL,
    notes        text
);

CREATE TABLE price_observation (
    price_id         text PRIMARY KEY CHECK (price_id LIKE 'PRC-%'),
    site_id          text NOT NULL REFERENCES site,
    feed_form        text NOT NULL CHECK (feed_form IN
                       ('manufacturing_scrap','cells','modules','ev_packs','ess_packs','consumer_batteries',
                        'black_mass','damaged_defective_recalled','unspecified')),
    chemistry        text NOT NULL CHECK (chemistry IN ('NMC','NCA','LFP','LCO','LMO','mixed','all','unspecified')),
    price_type       text NOT NULL CHECK (price_type IN ('purchase_price','processing_fee','payable_percent')),
    amount           numeric NOT NULL CHECK (amount >= 0),
    price_unit       text NOT NULL CHECK (price_unit IN ('$/kg','$/lb','$/t','% of metal value')),
    metal            text CHECK (metal IN ('Ni','Co','Cu','Li','Mn','n/a')),
    includes_freight text CHECK (includes_freight IN ('yes','no','unknown')),
    valid_from       date NOT NULL,
    valid_until      date,
    basis            text NOT NULL CHECK (basis IN ('quote','published_price_list','phone_call','email','report')),
    source_id        text NOT NULL REFERENCES source,
    notes            text,
    CHECK (valid_until IS NULL OR valid_until >= valid_from),
    CHECK ((price_type = 'payable_percent') = (price_unit = '% of metal value')),
    CHECK (price_type <> 'payable_percent' OR (metal IS NOT NULL AND metal <> 'n/a' AND amount <= 100))
);

CREATE TABLE feed_lot (
    lot_id          text PRIMARY KEY CHECK (lot_id LIKE 'LOT-%'),
    holder_org_id   text NOT NULL REFERENCES organization,
    site_id         text REFERENCES site,
    feed_form       text NOT NULL CHECK (feed_form IN
                      ('manufacturing_scrap','cells','modules','ev_packs','ess_packs','consumer_batteries',
                       'black_mass','damaged_defective_recalled','unspecified')),
    chemistry       text NOT NULL CHECK (chemistry IN ('NMC','NCA','LFP','LCO','LMO','mixed','all','unspecified')),
    quantity        numeric CHECK (quantity >= 0),
    unit            text CHECK (unit IN ('kg','t','lb','packs','modules','cells')),
    condition       text NOT NULL CHECK (condition IN ('intact','damaged_defective','recalled','unknown')),
    state_of_charge text,
    packaging       text,
    window_start    date,
    window_end      date,
    basis           text NOT NULL CHECK (basis IN ('direct_contact','referral','tender','marketplace')),
    status          text NOT NULL CHECK (status IN ('lead','qualified','quoting','placed','lost','withdrawn')),
    source_id       text NOT NULL REFERENCES source,
    date_checked    date NOT NULL,
    notes           text,
    CHECK (quantity IS NULL OR unit IS NOT NULL),
    CHECK (window_end IS NULL OR window_start IS NULL OR window_end >= window_start),
    CHECK (status NOT IN ('quoting','placed') OR (chemistry <> 'unspecified' AND condition <> 'unknown'))
);

CREATE TABLE deal (
    deal_id                      text PRIMARY KEY CHECK (deal_id LIKE 'DEAL-%'),
    lot_id                       text NOT NULL REFERENCES feed_lot,
    receiving_site_id            text REFERENCES site,
    transporter_org_id           text REFERENCES organization,
    price_id                     text REFERENCES price_observation,
    transport_price_id           text REFERENCES price_observation,
    stage                        text NOT NULL CHECK (stage IN
                                   ('enquiry','quote_requested','quoted','proposal_sent','accepted','scheduled',
                                    'completed','invoiced','paid','lost')),
    stage_date                   date NOT NULL,
    fee_or_margin                numeric,
    chemistry_condition_verified text CHECK (chemistry_condition_verified IN ('yes','no')),
    receiving_site_confirmed     text CHECK (receiving_site_confirmed IN ('yes','no')),
    transporter_hazmat_verified  text CHECK (transporter_hazmat_verified IN ('yes','no')),
    packaging_compliant          text CHECK (packaging_compliant IN ('yes','no')),
    outcome_reason               text,
    notes                        text,
    CHECK (stage IN ('enquiry','quote_requested','quoted','lost') OR
           (chemistry_condition_verified = 'yes' AND receiving_site_confirmed = 'yes'
            AND transporter_hazmat_verified = 'yes' AND packaging_compliant = 'yes'))
);

CREATE TABLE interaction (
    interaction_id   text PRIMARY KEY CHECK (interaction_id LIKE 'INT-%'),
    interaction_date date NOT NULL,
    contact_id       text REFERENCES contact,
    org_id           text REFERENCES organization,
    channel          text NOT NULL CHECK (channel IN ('phone','email','meeting','site_visit','other')),
    summary          text NOT NULL,
    outcome          text,
    follow_up_date   date,
    lot_id           text REFERENCES feed_lot,
    deal_id          text REFERENCES deal,
    CHECK (contact_id IS NOT NULL OR org_id IS NOT NULL)
);

CREATE TABLE audit_log (
    audit_id   text PRIMARY KEY CHECK (audit_id LIKE 'AUD-%'),
    audit_date date NOT NULL,
    tab        text NOT NULL,
    record_id  text NOT NULL,
    field      text,
    result     text NOT NULL CHECK (result IN ('correct','incorrect','changed','could_not_verify')),
    notes      text
);

-- Straight-line distance in km (haversine); road distance is longer.
CREATE FUNCTION km_between(lat1 numeric, lon1 numeric, lat2 numeric, lon2 numeric)
RETURNS numeric LANGUAGE sql IMMUTABLE AS $$
    SELECT round((2 * 6371 * asin(sqrt(
        power(sin(radians(lat2 - lat1) / 2), 2) +
        cos(radians(lat1)) * cos(radians(lat2)) * power(sin(radians(lon2 - lon1) / 2), 2)
    )))::numeric, 1)
$$;
