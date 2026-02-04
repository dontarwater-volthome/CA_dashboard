/***********************
 * HubSpot → Google Sheets (Jobs + Deals + Contacts)
 * Menu: HubSpot → Sync Jobs + Deals
 ***********************/

// ============ CONFIG ============
const HUBSPOT = {
  // Optional inline token fallback (or leave blank to rely on Script Properties)
  TOKEN_INLINE: PropertiesService.getScriptProperties().getProperty('hubspot_token'),
  ENABLE_CA_FILTER: true, // set to false to disable server-side CA filtering
  JOBS_OBJECT_TYPE_ID: '2-41941336',
  JOBS_PROPERTIES: [
    'associated_contact_record_id',
    'associated_deal_record_id',
    'hs_object_id','job_name','job_agreement_date_1','system__size__watts_',
    'street_address','city','state','zip_code','service_area',
    'partner',
    'amount','cashback','payment_method1',
    'installation_status','utility_status',
    'battery_services_cost','dealerfee','adder_amount',
    'fulfillment_partnerfee','additional_services_price','collection_base_amount',
    'override','cp_status_start__submitted__date','volt','lightreach_tesla_adder',
    'materials_request','domestic_content','permit_fees__adder',
    'hs_pipeline_stage','hs_pipeline',
    'hvac_price','roof_price','water_filter_price','estimated_installation_costs',
    'additional_services_amount','is_test',
    'solar_roof_s__size__watts_','system_size__ton_','hvac_quantity',
    'hvac_contract_value__view_only_','actual_stage','project_update','update_date',
    'date_entered__stand_by__stage','sales_team_take',
    'panel___brand__only_view_','panel___model__only_view_','panel_quantity__only_view_',
    'inverter___model__only_view_','inverter___brand__only_view_','inverter_quantity__only_view_',
    'battery___brand__only_view_','battery_quantity__only_view_','battery___model__only_view_',
    'm1_amount___paid','m1_amount','m2_amount','m2_amount___paid','clawbacks___applied'
  ],
  CONTACT_PROPERTIES: [
    'firstname','lastname','full_name','email','phone','language_preference'
  ],
  DEALS_PROPERTIES: [
    'dealname','existing_roof_type','existing_secondary_roof_type',
    'panel_panel_watts','utility_company','roof_size__sq_'
  ],
  FILTER_PIPELINE_LABEL: '', // leave blank to not filter

  FINAL_ORDER: [
    'firstname','lastname','full_name','email','phone','language_preference',
    'hs_object_id','job_name','job_agreement_date_1','system__size__watts_',
    'street_address','city','state','zip_code','service_area',
    'partner',
    'amount','payment_method1','cashback','installation_status','utility_status',
    'panel___brand__only_view_','inverter___brand__only_view_','battery___brand__only_view_',
    'panel_quantity__only_view_','inverter_quantity__only_view_','battery_quantity__only_view_',
    'panel___model__only_view_','inverter___model__only_view_','battery___model__only_view_',
    'panel_panel_watts','battery_services_cost','utility_company','dealerfee',
    'adder_amount','fulfillment_partnerfee','additional_services_price',
    'collection_base_amount','override','cp_status_start__submitted__date',
    'volt','lightreach_tesla_adder','materials_request',
    'existing_roof_type','domestic_content','existing_secondary_roof_type',
    'permit_fees__adder','associated_deal_record_id','hs_pipeline_stage',
    'hs_pipeline','hvac_price','roof_price','water_filter_price',
    'estimated_installation_costs','additional_services_amount','is_test',
    'roof_size__sq_','solar_roof_s__size__watts_',
    'system_size__ton_','hvac_quantity','hvac_contract_value__view_only_',
    'actual_stage','project_update','update_date','date_entered__stand_by__stage',
    'sales_team_take','m1_amount___paid','m1_amount','m2_amount',
    'm2_amount___paid','clawbacks___applied'
  ],
  SHEET_NAME: 'data'
};

// Which columns come from Deals (the rest come from Jobs)
const DEAL_COLS = new Set([
  'utility_company','existing_roof_type',
  'existing_secondary_roof_type','roof_size__sq_','panel_panel_watts'
]);

// Which columns come from Contacts (primary associated to Job)
const CONTACT_COLS = new Set([
  'firstname','lastname','full_name','email','phone','language_preference'
]);

// ============ MENU ============
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('HubSpot')
    .addItem('Sync Jobs + Deals', 'syncHubSpotJobsAndDeals')
    .addItem('Show Stage Summary', 'showSidebar')
    .addToUi();
}

// ============ ENTRYPOINT ============
function syncHubSpotJobsAndDeals() {
  const token = getHubSpotToken_();

  // 1) Pipelines → translate ids to labels
  const pipelineData = fetchPipelineLabels_(token, HUBSPOT.JOBS_OBJECT_TYPE_ID);
  const pipelineMap = pipelineData.pipelineMap;
  const stageMap = pipelineData.stageMap;

  // 2) Pull Jobs (server-side filtered if enabled)
  const jobs = HUBSPOT.ENABLE_CA_FILTER
    ? fetchHubspotDataSearchFiltered_(token, HUBSPOT.JOBS_OBJECT_TYPE_ID, HUBSPOT.JOBS_PROPERTIES, [
        { propertyName: 'state', operator: 'EQ', value: 'CA' },
        { propertyName: 'state', operator: 'EQ', value: 'California' }
      ])
    : fetchHubspotData_(token, HUBSPOT.JOBS_OBJECT_TYPE_ID, HUBSPOT.JOBS_PROPERTIES);

  // 2.5) Pull Deals
  const deals = fetchHubspotData_(token, 'deals', HUBSPOT.DEALS_PROPERTIES);

  // 3) Pull Contact Info (by associated_contact_record_id)
  const jobContactIds = jobs
    .map(j => j.properties?.associated_contact_record_id)
    .filter(Boolean)
    .map(String);

  const contactIds = Array.from(new Set(jobContactIds));
  const contactsMap = fetchContactsByIds_(token, contactIds, HUBSPOT.CONTACT_PROPERTIES);

  // 4) Build Deal map
  const dealsMap = new Map(deals.map(d => [String(d.id), d.properties || {}]));

  // 5) Filter + combine
  const combined = [];
  for (let i = 0; i < jobs.length; i++) {
    const job = jobs[i];

    if (HUBSPOT.FILTER_PIPELINE_LABEL) {
      const jp = job.properties || {};
      const label = pipelineMap.get(jp.hs_pipeline) || jp.hs_pipeline;
      if (label !== HUBSPOT.FILTER_PIPELINE_LABEL) continue;
    }

    const row = { id: job.id };
    const jobProps = job.properties || {};
    const contactId = jobProps.associated_contact_record_id ? String(jobProps.associated_contact_record_id) : null;
    const contactProps = contactId ? (contactsMap.get(contactId) || {}) : {};
    const dealId = jobProps.associated_deal_record_id ? String(jobProps.associated_deal_record_id) : null;
    const dealProps = dealId ? (dealsMap.get(dealId) || {}) : {};

    for (let f = 0; f < HUBSPOT.FINAL_ORDER.length; f++) {
      const col = HUBSPOT.FINAL_ORDER[f];
      let value;

      if (CONTACT_COLS.has(col)) {
        value = contactProps[col] != null ? contactProps[col] : '';
      } else if (DEAL_COLS.has(col)) {
        value = dealProps[col] != null ? dealProps[col] : '';
      } else {
        value = jobProps[col] != null ? jobProps[col] : '';
      }

      if (col === 'full_name' && !value) {
        value = buildFullName_(contactProps.firstname, contactProps.lastname);
      }
      if (col === 'phone' && value) {
        value = normalizePhone_(value);
      }
      if (col === 'job_agreement_date_1' && value !== '' && value != null) {
        value = parseAnyDate_(value);
      }
      if (col === 'state' && value) {
        value = normalizeState_(value);
      }
      if (col === 'update_date' && value !== '' && value != null) {
        value = parseAnyDate_(value);
      }
      if (col === 'date_entered__stand_by__stage' && value !== '' && value != null) {
        value = parseAnyDate_(value);
      }

      if (col === 'hs_pipeline' && value) {
        value = pipelineMap.get(value) || value;
      }
      if (col === 'hs_pipeline_stage' && value) {
        value = stageMap.get(value) || value;
      }

      row[col] = value;
    }

    combined.push(row);
  }

  // 6) Write without wiping other columns
  writeSelective_(HUBSPOT.SHEET_NAME, combined, HUBSPOT.FINAL_ORDER);

  SpreadsheetApp.getUi().alert(`Synced ${combined.length} HubSpot job rows to "${HUBSPOT.SHEET_NAME}".`);
}

// ============ HELPERS ============

function getHubSpotToken_() {
  const props = PropertiesService.getScriptProperties();
  const propUpper = props.getProperty('HUBSPOT_TOKEN');
  if (propUpper && propUpper.trim()) return propUpper.trim();
  const propLower = props.getProperty('hubspot_token');
  if (propLower && propLower.trim()) return propLower.trim();
  if (HUBSPOT.TOKEN_INLINE && HUBSPOT.TOKEN_INLINE.trim()) return HUBSPOT.TOKEN_INLINE.trim();
  throw new Error('No HubSpot token. Set Script Property HUBSPOT_TOKEN or HUBSPOT.TOKEN_INLINE.');
}

function buildFullName_(first, last) {
  const f = first ? String(first).trim() : '';
  const l = last ? String(last).trim() : '';
  return [f, l].filter(Boolean).join(' ');
}

/**
 * Generic fetch for HubSpot objects with pagination.
 */
function fetchHubspotData_(token, objectType, properties) {
  const props = properties.join(',');
  const out = [];
  let after = null;

  do {
    const base = `https://api.hubapi.com/crm/v3/objects/${encodeURIComponent(objectType)}`;
    const url  = `${base}?properties=${encodeURIComponent(props)}&limit=100${after ? `&after=${encodeURIComponent(after)}` : ''}&archived=false`;
    const resp = hubspotFetch_(url, token);
    const data = JSON.parse(resp.getContentText());
    if (data?.results?.length) out.push(...data.results);
    after = data?.paging?.next?.after || null;
  } while (after);

  return out;
}

/**
 * Search API fetch with server-side filters (OR between filter groups).
 */
function fetchHubspotDataSearchFiltered_(token, objectType, properties, filtersOrGroups) {
  const out = [];
  let after = null;

  const filterGroups = (filtersOrGroups || []).map(f => ({
    filters: [f]
  }));

  do {
    const url = `https://api.hubapi.com/crm/v3/objects/${encodeURIComponent(objectType)}/search`;
    const body = {
      properties: properties || [],
      limit: 100,
      after: after,
      filterGroups: filterGroups
    };

    const resp = hubspotFetch_(url, token, 1, {
      method: 'post',
      payload: JSON.stringify(body),
      contentType: 'application/json'
    });

    const data = JSON.parse(resp.getContentText());
    if (data?.results?.length) out.push(...data.results);
    after = data?.paging?.next?.after || null;
  } while (after);

  return out;
}

/**
 * Fetch contacts by id (batch).
 */
function fetchContactsByIds_(token, contactIds, properties) {
  const map = new Map();
  if (!contactIds.length) return map;

  const url = 'https://api.hubapi.com/crm/v3/objects/contacts/batch/read';

  for (const ids of chunk_(contactIds, 100)) {
    const resp = hubspotFetch_(url, token, 1, {
      method: 'post',
      payload: JSON.stringify({
        inputs: ids.map(id => ({ id })),
        properties: properties || []
      }),
      contentType: 'application/json'
    });
    const data = JSON.parse(resp.getContentText());
    for (const c of (data?.results || [])) {
      map.set(String(c.id), c.properties || {});
    }
  }
  return map;
}

function chunk_(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

/**
 * Translate pipeline/stage ids → labels for a given objectTypeId
 */
function fetchPipelineLabels_(token, objectTypeId) {
  const url  = `https://api.hubapi.com/crm/v3/pipelines/${encodeURIComponent(objectTypeId)}`;
  const resp = hubspotFetch_(url, token);
  const data = JSON.parse(resp.getContentText());

  const pipelineMap = new Map();
  const stageMap    = new Map();

  for (const pipe of (data?.results || [])) {
    pipelineMap.set(pipe.id, pipe.label);
    for (const stage of (pipe?.stages || [])) {
      stageMap.set(stage.id, stage.label);
    }
  }
  return { pipelineMap, stageMap };
}

function normalizePhone_(value) {
  if (value == null) return '';
  const raw = String(value).trim();
  if (!raw) return '';

  const hasPlus = raw.startsWith('+');
  const digits = raw.replace(/[^\d]/g, '');

  if (digits.length === 10) {
    return `(${digits.slice(0,3)}) ${digits.slice(3,6)}-${digits.slice(6)}`;
  }
  if (digits.length === 11 && digits.startsWith('1')) {
    const d = digits.slice(1);
    return `(${d.slice(0,3)}) ${d.slice(3,6)}-${d.slice(6)}`;
  }
  if (hasPlus) return `+${digits}`;
  return raw;
}

function normalizeState_(value) {
  if (value == null) return '';
  const v = String(value).trim();
  if (!v) return '';
  const upper = v.toUpperCase();

  const map = {
    'ALABAMA':'AL','ALASKA':'AK','ARIZONA':'AZ','ARKANSAS':'AR','CALIFORNIA':'CA','COLORADO':'CO',
    'CONNECTICUT':'CT','DELAWARE':'DE','FLORIDA':'FL','GEORGIA':'GA','HAWAII':'HI','IDAHO':'ID',
    'ILLINOIS':'IL','INDIANA':'IN','IOWA':'IA','KANSAS':'KS','KENTUCKY':'KY','LOUISIANA':'LA',
    'MAINE':'ME','MARYLAND':'MD','MASSACHUSETTS':'MA','MICHIGAN':'MI','MINNESOTA':'MN',
    'MISSISSIPPI':'MS','MISSOURI':'MO','MONTANA':'MT','NEBRASKA':'NE','NEVADA':'NV',
    'NEW HAMPSHIRE':'NH','NEW JERSEY':'NJ','NEW MEXICO':'NM','NEW YORK':'NY',
    'NORTH CAROLINA':'NC','NORTH DAKOTA':'ND','OHIO':'OH','OKLAHOMA':'OK','OREGON':'OR',
    'PENNSYLVANIA':'PA','RHODE ISLAND':'RI','SOUTH CAROLINA':'SC','SOUTH DAKOTA':'SD',
    'TENNESSEE':'TN','TEXAS':'TX','UTAH':'UT','VERMONT':'VT','VIRGINIA':'VA','WASHINGTON':'WA',
    'WEST VIRGINIA':'WV','WISCONSIN':'WI','WYOMING':'WY','DISTRICT OF COLUMBIA':'DC'
  };

  if (upper.length === 2) return upper;
  return map[upper] || v;
}

function parseAnyDate_(value) {
  const asNum = Number(value);
  if (!Number.isNaN(asNum) && asNum > 0) {
    const ms = asNum < 2e10 ? asNum * 1000 : asNum;
    const d = new Date(ms);
    return isNaN(d) ? value : d;
  }
  const d = new Date(String(value));
  return isNaN(d) ? value : d;
}

/**
 * Ensure headers, clear only our target columns (below header), and write values.
 */
function writeSelective_(sheetName, rows, headers) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  const needed = ['id', ...headers];
  const existingHeader = sheet.getLastRow() >= 1
    ? (sheet.getRange(1,1,1, Math.max(1, sheet.getLastColumn())).getValues()[0] || [])
    : [];

  const headerIndex = new Map();

  needed.forEach(h => {
    let idx = existingHeader.indexOf(h) + 1;
    if (idx <= 0) {
      idx = existingHeader.length + 1;
      sheet.getRange(1, idx).setValue(h);
      existingHeader.push(h);
    }
    headerIndex.set(h, idx);
  });

  const lastRow = Math.max(2, sheet.getLastRow());
  needed.forEach(h => {
    const c = headerIndex.get(h);
    if (lastRow > 1) sheet.getRange(2, c, lastRow - 1, 1).clearContent();
  });

  if (!rows || !rows.length) return;

  for (const h of needed) {
    const col = headerIndex.get(h);
    const colValues = rows.map(r => formatCellValue_(h === 'id' ? r.id : r[h]));
    sheet.getRange(2, col, colValues.length, 1).setValues(colValues.map(v => [v]));

    if (h === 'job_agreement_date_1') {
      sheet.getRange(2, col, colValues.length, 1).setNumberFormat('yyyy-mm-dd;@');
    }
  }
}

function formatCellValue_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  return v == null ? '' : v;
}

/**
 * HubSpot fetch with simple 429 retry/backoff.
 */
function hubspotFetch_(url, token, attempt = 1, options = {}) {
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: options.method || 'get',
      headers: Object.assign(
        { Authorization: `Bearer ${token}` },
        options.headers || {}
      ),
      muteHttpExceptions: true,
      payload: options.payload,
      contentType: options.contentType
    });
    const code = resp.getResponseCode();
    if (code === 429 && attempt <= 5) {
      const retryAfter = parseInt(resp.getAllHeaders()['Retry-After'] || '1', 10);
      Utilities.sleep(Math.max(1, retryAfter) * 1000);
      return hubspotFetch_(url, token, attempt + 1, options);
    }
    if (code >= 200 && code < 300) return resp;
    throw new Error(`HubSpot HTTP ${code}: ${resp.getContentText()}`);
  } catch (err) {
    if (attempt <= 5) {
      Utilities.sleep(1000 * attempt);
      return hubspotFetch_(url, token, attempt + 1, options);
    }
    throw err;
  }
}

// == SIDEBAR DASHBOARD ====
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Pipeline Stage Summary');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getStageSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Active Projects');
  if (!sheet) return { rows: [], total: 0 };
  const lastRow = sheet.getLastRow();
  if (lastRow < 12) return { rows: [], total: 0 };

  const header = sheet.getRange(11, 1, 1, sheet.getLastColumn()).getValues()[0];
  const normalized = header.map(h => String(h).trim().toLowerCase());
  const stageColIndex = normalized.indexOf('stage');
  if (stageColIndex === -1) return { rows: [], total: 0 };

  const dataStartRow = 12;
  const data = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, sheet.getLastColumn()).getValues();

  const map = new Map();
  data.forEach(row => {
    const stage = row[stageColIndex];
    if (!stage) return;
    const key = String(stage).trim();
    if (!key) return;
    map.set(key, (map.get(key) || 0) + 1);
  });

  const order = [
    'Stand by',
    'New Job',
    'Pending Documents',
    'Site Survey',
    'Pending NTP',
    'Engineering',
    'Permitting',
    'Pre-Install Actions Pending',
    'Scheduling',
    'Installation',
    'Issues',
    'Final Inspection Pending',
    'Utility'
  ];

  const rows = order.map(stage => ({
    stage,
    count: map.get(stage) || 0
  }));

  const total = rows.reduce((sum, r) => sum + r.count, 0);

  return { rows, total };
}
// Updated by Codex on 2026-02-04
