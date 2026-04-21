/**************************************************************
 * Project 4 — Automated Credential Generation & Distribution
 **************************************************************/

const SETTINGS = {
  TABS: {
    CONFIG: 'Config',
    CONTENT_BLOCKS: 'Content Blocks',
    ORDER: '4 Assignment Emails',
    P4_EMAIL_ORDER: '4 Assignment Emails',
    LBJ_IMPORT: 'LBJ_Import (latest)',
    MASTER: 'Master_Assignments',
    COUNTY_ROLLUP: 'County_Rollup',
    GENERAL_MM_ROLLUP: 'General_MM_Rollup',
    BRE_MERGE: 'BRE Merge Sheet',
    LOGS: 'Logs',
    README: 'README',
  },

  IMAGE_DEFAULTS: {
    LOGO_MAX_WIDTH: 360,
    SIGNATURE_HEIGHT: 96
  },

  DEFAULT_BRE: { WORKBOOK_ID: '', TAB_NAME: 'BRE Merge Sheet' },

  OUTPUT: {
    VOLUNTEERS_FOLDER_ID: 'PUT_VOLUNTEER_PDFS_FOLDER_ID',
    COUNTIES_FOLDER_ID: 'PUT_COUNTY_OUTPUT_FOLDER_ID',
  },

  MAX_ASSIGNMENTS: 15,

  MASTER_HEADERS: [
    'VAN_ID','Volunteer_Name','Volunteer_Email','Volunteer_Phone','Volunteer Address','County','Assignment_County','Run_Year',
    'Yes_Link','No_Link',
    'Assignment_Count','Doc_Mode','Assignments_JSON','Cred_Hash',
    ...Array.from({length: 15}).flatMap((_,i)=>[
      `A${i+1}_Date`,`A${i+1}_Start`,`A${i+1}_End`,`A${i+1}_LocationName`,`A${i+1}_Address`
    ]),
    'Volunteer_PDF_File_ID','ICS_File_Name',
    'Credential_Sent_On','Credential_Sent_By','Source_LBJ_Sheet_Tag',
    'Needs_Confirmation_Send',
    'Needs_Credential_Send','Errors'
  ],

  COUNTY_HEADERS: [
    'County','Run_Year','Volunteer_Count','VAN_ID_List',
    'County_PDF_File_ID','County_CSV_File_ID','County_Email_Sent_On','Errors'
  ],

  GENERAL_MM_HEADERS: [
    'County','Run_Year','General_MM_Sent_On','Errors'
  ],

  LBJ_FIELDS: {
    VAN_ID:          ['VAN_ID','VanID','VAN ID'],
    FIRST:           ['First Name','FirstName','VolFirstName','Volunteer First Name'],
    LAST:            ['Last Name','LastName','VolLastName','Volunteer Last Name'],
    EMAIL:           ['Email','Email Address','VolEmail','Volunteer Email'],
    PHONE:           ['Phone Number','Phone','Cell Phone','Cell Phone Number','Mobile','VolPhone'],
    COUNTY:          ['County'],
    ASSIGNMENT_COUNTY: ['Assignment County','AssignmentCounty','Assigned County','Assignment Cnty','Assignment'],
    DATE:            ['Date','Assignment Date'],
    START:           ['Start Time','StartTime'],
    END:             ['End Time','EndTime'],
    LOCATION_NAME:   ['Polling Location','Polling Location Name','Location Name','LocationName'],
    LOCATION_ADDR:   ['Polling Location Address','Location Address','LocationAddress'],
    SHEET_TAG:       ['LBJ_Sheet_Tag','Sheet_Tag'],
    VOL_ADDRESS: [
      'Address','Volunteer Address','Mailing Address','Mailing_Address',
      'Address Line 1','Address1','Street Address','Home Address','Residential Address'
    ]
  },

  FORM_LINK: {
    BASE_URL_CONFIG_KEY: 'Form_Base_URL',
    BASE_URL_CELL_KEY:   'Form_Base_URL_Cell',
    YES_PARAM: 'Yes',
    NO_PARAM: 'No',
    PLACEHOLDERS: { first:'VolFirstName', last:'VolLastName', email:'VolEmail', phone:'VolPhone', avail:'AvailGood' },
    ENTRY_IDS:   { first:'entry.576587497', last:'entry.1363204004', email:'entry.957850353', phone:'entry.436579453', avail:'entry.1913347883' }
  },

  SUBJECTS: {
    VOL_M1: 'Assignment Confirmation — Action Requested',
    VOL_M2: 'Your Official Poll Watcher Credential',
    COUNTY_M3: 'County Packet: Poll Watcher Credentials & Roster',
    GEN_MM: 'General Mail Merge'
  },

  ORDER_KEYS: {
    VOL_M1_SINGLE: 'Single Assignment Confirmation Email',
    VOL_M1_MULTI:  'Multiple Assignment Confirmation Email',
    VOL_M2_SINGLE: 'Single Assignment Credential Email',
    VOL_M2_MULTI:  'Multiple Assignment Credential Email',
    COUNTY_EMAIL:  'Email to Counties',
    LETTER_SINGLE: 'Single Assignment Credential Letter',
    LETTER_MULTI:  'Multiple Assignment Credential Letter',
    GENERAL_MM:  'General Mail Merge'
  },

  /**************************************************************
   * ADDED — County volunteer spreadsheet mapping + output tab
   **************************************************************/
  COUNTY_VOL_SHEET: {
    MAP_TAB_NAME: 'County_File_Map',
    MAP_HEADERS: ['County', 'Spreadsheet_ID', 'Spreadsheet_Name', 'Notes'],
    TAB_NAME: 'Assigned Volunteers',
    HEADERS: [
      'VAN_ID',
      'Volunteer_Name',
      'Volunteer_Email',
      'Volunteer_Phone',
      'Volunteer Address',
      'Home_County',
      'Assignment_County',
      'Assignment_Count',
      'Assignment_1',
      'Assignment_2',
      'Assignment_3',
      'Assignment_4',
      'Assignment_5',
      'Assignment_6',
      'Assignment_7',
      'Assignment_8',
      'Assignment_9',
      'Assignment_10',
      'Assignment_11',
      'Assignment_12',
      'Assignment_13',
      'Assignment_14',
      'Assignment_15',
      'Assignments_JSON',
      'Assignments_JSON_Last_Updated'
    ]
  }
};

/** ==============================
 *  UTILITY FUNCTIONS (MUST BE FIRST)
 *  ============================== */

function extractSpreadsheetId_(s) {
  const raw = String(s || '').trim();
  if (!raw) return '';
  if (/^[a-zA-Z0-9-_]{25,}$/.test(raw)) return raw;
  const m = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]{25,})/);
  return m ? m[1] : '';
}

function getChildFolderByName_(parentFolder, childName) {
  const it = parentFolder.getFoldersByName(childName);
  if (!it.hasNext()) throw new Error(`Missing folder "${childName}" inside "${parentFolder.getName()}"`);
  return it.next();
}

function resolveOutputFolderIdsFromElectionRoot_(cfg) {
  const rootId = (cfg.Election_Root_Folder_ID || '').trim();
  if (!rootId) return null;

  try {
    const root = DriveApp.getFolderById(rootId);
    const assignments = getChildFolderByName_(root, 'Assignments');
    const vol = getChildFolderByName_(assignments, '02 Output / Volunteers');
    const cty = getChildFolderByName_(assignments, '03 Output / Counties');

    return {
      volunteersFolderId: vol.getId(),
      countiesFolderId: cty.getId()
    };
  } catch (e) {
    Logger.log('Could not resolve output folders from Election Root: ' + e.message);
    return null;
  }
}

function escapeHtml_(s){ return (s||'').toString().replace(/[&<>"']/g, c=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function stripHtml_(s){ return String(s||'').replace(/<[^>]*>/g,''); }
function icsEscape_(s){ return (s||'').replace(/[,;]/g, '\\$&'); }
function csvEscape_(s){ const v=(s||'').toString(); return /[",\n]/.test(v) ? `"${v.replace(/"/g,'""')}"` : v; }
function sha256_(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str);
  return raw.map(b=>{ const s=(b<0? b+256:b).toString(16); return s.length===1?'0'+s:s; }).join('');
}

function escapeAttr_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function shortenUrlForDisplay_(url) {
  let u = String(url || '').trim();
  if (!u) return '';

  // remove scheme
  u = u.replace(/^https?:\/\//i, '');

  // remove trailing punctuation often present in sentences/bullets
  u = u.replace(/[).,;:!?]+$/g, '');

  // drop query/hash for display only
  u = u.split('#')[0].split('?')[0];

  return u;
}

/**
 * Wraps any visible URLs in the HTML with <a href="...">short text</a>,
 * but only if the HTML does NOT already contain an <a> tag.
 */
function linkifyUrlsIfNoAnchors_(html) {
  const h = String(html || '');
  if (!h.trim()) return h;

  // If author already used <a>, don't interfere.
  if (/<a\b/i.test(h)) return h;

  // Replace URLs that appear as plain text.
  return h.replace(/(https?:\/\/[^\s<]+)([).,;:!?]*)/gi, (m, url, trailing) => {
    const display = shortenUrlForDisplay_(url) || url;
    const safeHref = escapeAttr_(url);
    const safeText = escapeHtml_(display);
    return `<a href="${safeHref}" target="_blank" rel="noopener noreferrer" style="color:#1155cc;text-decoration:underline;">${safeText}</a>${trailing || ''}`;
  });
}


function normalizeUrl_(s) {
  let u = String(s || '').trim();
  if (!u) return '';

  // If it's a Drive file id, convert to a standard open URL (no hardcoding “your” form URL)
  if (/^[A-Za-z0-9_-]{25,}$/.test(u)) {
    return `https://drive.google.com/open?id=${u}`;
  }

  // If it looks like a URL but missing scheme, add https://
  if (!/^https?:\/\//i.test(u)) {
    // common cases: "www.example.com", "drive.google.com/...", etc.
    if (/^(www\.)/i.test(u) || /^[a-z0-9.-]+\.[a-z]{2,}\//i.test(u) || /google\.com/i.test(u)) {
      u = 'https://' + u.replace(/^\/+/, '');
    }
  }

  return u;
}

function wrapIfBareUrl_(html) {
  const h = String(html || '').trim();
  if (!h) return '';

  // If it already has a link, don't touch it.
  if (/<a\b/i.test(h)) return h;

  // If the visible content is JUST a URL, wrap it.
  const txt = stripHtml_(h).trim();

  if (/^https?:\/\/\S+$/i.test(txt)) {
    const url = normalizeUrl_(txt);
    if (!url) return h;
    return `<a href="${escapeAttr_(url)}" target="_blank" rel="noopener noreferrer" style="color:#1155cc;text-decoration:underline;">${escapeHtml_(txt)}</a>`;
  }

  return h;
}


function getTab_(ss, name){ const sh = ss.getSheetByName(name); if (!sh) throw new Error(`Missing sheet: ${name}`); return sh; }
function getOrCreate_(ss, name){ return ss.getSheetByName(name) || ss.insertSheet(name); }
function getOrCreateWithHeader_(ss, name, header) {
  const sh = getOrCreate_(ss, name);
  const vals = sh.getRange(1,1,1, header.length).getValues()[0];
  const hasHeader = (vals.some(v => v && v.toString().trim().length));
  if (!hasHeader) sh.getRange(1,1,1, header.length).setValues([header]);
  return sh;
}
function getHeader_(sh){ return sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0]; }
function getData_(sh){ return sh.getRange(1,1, Math.max(1,sh.getLastRow()), Math.max(1,sh.getLastColumn())).getValues(); }
function rowFromObj_(header, obj){ return header.map(h => (h in obj) ? obj[h] : ''); }
function objFromRow_(header, row){ const o={}; header.forEach((h,i)=>o[h]=row[i]); return o; }
function groupBy_(arr, fn){ const m={}; for (const x of arr){ const k=fn(x); if(!m[k]) m[k]=[]; m[k].push(x); } return m; }
function toast_(msg, sec){ SpreadsheetApp.getActive().toast(msg, 'Project 4', sec||3); }
function log_(cfg, action, payload){
  const ss = getAssignmentsSs_(cfg);
  const sh = getOrCreate_(ss, SETTINGS.TABS.LOGS);
  sh.appendRow([new Date(), action, JSON.stringify(payload||{})]);
}

function getCellByA1_(ss, a1){ if (!a1) return ''; const [tab, cell] = a1.split('!'); const sh=ss.getSheetByName(tab); if (!sh) return ''; return (sh.getRange(cell).getValue()||'').toString(); }

function mapHeaderIdxFlexible_(header, fieldAliases) {
  const idx = {};
  Object.keys(fieldAliases).forEach(key => {
    const aliases = Array.isArray(fieldAliases[key]) ? fieldAliases[key] : [fieldAliases[key]];
    idx[key] = -1;
    for (const name of aliases) {
      const i = header.indexOf(name);
      if (i >= 0) { idx[key] = i; break; }
    }
  });
  return idx;
}

function P4_getEmailOrderTab_(ss, cfg) {
  // Allow overriding via MEC if you ever add a config field later
  const preferred = (cfg && cfg.Email_Order_Tab_Name) ? String(cfg.Email_Order_Tab_Name).trim() : '';

  const candidates = [
    preferred,
    '4 Assignment Emails',
    'Project 4 Email Order',
    'Project 4 Email Order ',      // (sometimes trailing spaces happen)
    '4 Assignment Email Order',
    '4 Assignment Email Order ',
  ].filter(Boolean);

  for (const name of candidates) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }

  // Helpful error that tells you what tabs actually exist
  const names = ss.getSheets().map(s => s.getName()).join(', ');
  throw new Error('Missing sheet: 4 Assignment Emails (tried: ' + candidates.join(' | ') + '). Found tabs: ' + names);
}


function getValIdx_(row, idx) { return (idx >= 0 && idx < row.length) ? (row[idx] || '').toString().trim() : ''; }
function valByAliases_(obj, aliases) {
  const arr = Array.isArray(aliases) ? aliases : [aliases];
  for (const k of arr) {
    if (Object.prototype.hasOwnProperty.call(obj, k) && obj[k] != null) return String(obj[k]).trim();
  }
  return '';
}

function inferSheetTag_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'); }
function dedupeAssignments_(assigns) {
  const seen = new Set(), out = [];
  assigns.forEach(a => {
    const key = [a.date, a.start, a.end, a.locationName, a.address].join('|');
    if (!seen.has(key)) { seen.add(key); out.push(a); }
  });
  return out;
}

function asBool_(v) {
  if (v === true)  return true;
  if (v === false) return false;
  const s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === 't' || s === 'yes' || s === 'y' || s === '1';
}

function parseJsonSafe_(s) {
  try { return JSON.parse(String(s || '{}')); } catch (e) { return {}; }
}

function normalizeDateISO_(v, tz) {
  const d = parseDateFlexible_(v, tz || Session.getScriptTimeZone());
  const y = String(d.y).padStart(4,'0');
  const m = String(d.m).padStart(2,'0');
  const day = String(d.d).padStart(2,'0');
  return `${y}-${m}-${day}`;
}

function normalizeTimeHHMM_(v, tz) {
  const t = parseTimeFlexible_(v, tz || Session.getScriptTimeZone());
  const hh = String(t.hh).padStart(2,'0');
  const mm = String(t.mm).padStart(2,'0');
  return `${hh}:${mm}`;
}

function canonicalizeAssignmentsForHash_(assigns, tz) {
  const seen = new Set();
  const out = [];

  (assigns || []).forEach(a => {
    const norm = {
      d: normalizeDateISO_(a.date, tz),
      s: normalizeTimeHHMM_(a.start, tz),
      e: normalizeTimeHHMM_(a.end, tz),
      n: String(a.locationName || '').trim().replace(/\s+/g,' '),
      a: String(a.address || '').trim().replace(/\s+/g,' ')
    };
    const key = `${norm.d}|${norm.s}|${norm.e}|${norm.n}|${norm.a}`;
    if (!seen.has(key)) { seen.add(key); out.push(norm); }
  });

  out.sort((A,B)=>
    A.d.localeCompare(B.d) ||
    A.s.localeCompare(B.s) ||
    A.n.localeCompare(B.n) ||
    A.a.localeCompare(B.a)
  );
  return JSON.stringify(out);
}

function canonicalizeAssignments_(assignments) {
  if (!Array.isArray(assignments)) return '[]';
  const norm = assignments.map(a => ({
    date:         String(a.date || '').trim(),
    start:        String(a.start || '').trim(),
    end:          String(a.end || '').trim(),
    locationName: String(a.locationName || '').trim(),
    address:      String(a.address || '').trim(),
    county:       String(a.county || '').trim(),
  }))
  .sort((a,b) =>
    a.date.localeCompare(b.date) ||
    a.start.localeCompare(b.start) ||
    a.locationName.localeCompare(b.locationName) ||
    a.address.localeCompare(b.address)
  );
  return JSON.stringify(norm);
}

function assignmentsHashFromJson_(jsonStr) {
  const parsed = parseJsonSafe_(jsonStr);

  if (Array.isArray(parsed)) {
    const items = parsed.map(x => ({
      date: x.date || x.d || '',
      start: x.start || x.s || '',
      end: x.end || x.e || '',
      locationName: x.locationName || x.n || '',
      address: x.address || x.a || ''
    }));
    return sha256_(canonicalizeAssignments_(items));
  }

  const arr = (parsed && (parsed.assignments || parsed.items)) || [];
  return sha256_(canonicalizeAssignments_(arr));
}

function pick_(obj, ...keys) {
  for (const k of keys) {
    if (obj && obj[k] != null && String(obj[k]).trim() !== '') return String(obj[k]).trim();
  }
  return '';
}

function formatPrettyDate_(d, tz) {
  let dt;
  if (Object.prototype.toString.call(d) === '[object Date]') {
    dt = d;
  } else if (/^\d{4}-\d{2}-\d{2}$/.test(String(d))) {
    const [y,m,day] = String(d).split('-').map(Number);
    dt = new Date(y, m - 1, day);
  } else {
    dt = new Date(d);
  }
  return Utilities.formatDate(dt, tz || 'America/New_York', 'MMMM d, yyyy');
}

function parseDateFlexible_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    const parts = Utilities.formatDate(v, tz || Session.getScriptTimeZone(), 'yyyy,MM,dd').split(',');
    return { y: +parts[0], m: +parts[1], d: +parts[2] };
  }
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return { y:+m[1], m:+m[2], d:+m[3] };
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const yy = (+m[3] < 100) ? 2000 + (+m[3]) : +m[3];
    return { y: yy, m: +m[1], d: +m[2] };
  }
  const today = new Date();
  return { y: today.getFullYear(), m: today.getMonth()+1, d: today.getDate() };
}

function parseTimeFlexible_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    const parts = Utilities.formatDate(v, tz || Session.getScriptTimeZone(), 'HH,mm').split(',');
    return { hh: +parts[0], mm: +parts[1] };
  }
  if (typeof v === 'number' && isFinite(v)) {
    const totalMin = Math.round(v * 24 * 60);
    const hh = Math.floor(totalMin / 60) % 24;
    const mm = totalMin % 60;
    return { hh, mm };
  }
  let s = String(v || '').trim();
  s = s.replace(/\u202F|\u00A0/g, ' ');
  let m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return { hh: +m[1], mm: +m[2] };
  m = s.match(/^(\d{1,2})(?::(\d{2}))?\s*([AaPp][Mm])?$/);
  if (m) {
    let hh = +(m[1] || 0);
    let mm = +(m[2] || 0);
    const ap = (m[3] || '').toUpperCase();
    if (ap === 'PM' && hh < 12) hh += 12;
    if (ap === 'AM' && hh === 12) hh = 0;
    return { hh, mm };
  }
  m = s.match(/^(\d{3,4})$/);
  if (m) {
    const digits = m[1].padStart(4, '0');
    return { hh: +digits.slice(0,2), mm: +digits.slice(2,4) };
  }
  return { hh: 8, mm: 0 };
}

function normCountyKey_(s){
  let t = String(s || '').toLowerCase().trim();
  t = t.replace(/^ga[\s\-_]*/, '');
  t = t.replace(/\bcounty\b/g, '');
  t = t.replace(/[^a-z0-9]+/g, ' ');
  return t.trim();
}

function normalizeCountyCore_(s) {
  return String(s||'')
    .trim()
    .replace(/^ga[-_\s]*/i, '')
    .replace(/\s+/g,' ')
    .replace(/[^\w\s-]/g,'')
    .trim();
}

function countySlugVariants_(county) {
  const core = normalizeCountyCore_(county).toLowerCase();
  const underscore = core.replace(/\s+/g,'_');
  const hyphen     = core.replace(/\s+/g,'-');
  return new Set([
    underscore, hyphen,
    `ga_${underscore}`, `ga-${hyphen}`,
    `ga${underscore.startsWith('_')?'':'_'}${underscore}`,
    `ga${hyphen.startsWith('-')?'':'-'}${hyphen}`
  ]);
}

function extractLastName_(fullName) {
  const parts = String(fullName || '').trim().split(/\s+/);
  return parts.length ? parts[parts.length - 1] : 'Unknown';
}

function fileSafe_(s) {
  return String(s || '').replace(/[\\/:*?"<>|]+/g, '').trim();
}

function wrapWithSpacer_(html) {
  if (!html) return '';
  return `<div style="margin:0 0 12px 0;">${html}</div>`;
}

function addInlineImage_(bag, cidName, blob) {
  if (!blob) return '';
  if (bag && bag.__useCid) {
    bag[cidName] = blob;
    return 'cid:' + cidName;
  }
  const mime = blob.getContentType() || 'image/png';
  const b64  = Utilities.base64Encode(blob.getBytes());
  return `data:${mime};base64,${b64}`;
}

function resolveBlobFromRef_(ref) {
  if (!ref) return null;
  const s = String(ref).trim();
  try {
    if (/^[A-Za-z0-9_-]{25,}$/.test(s)) {
      return DriveApp.getFileById(s).getBlob();
    }
    if (/^https?:\/\//i.test(s)) {
      const resp = UrlFetchApp.fetch(s, {muteHttpExceptions:true});
      if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) return resp.getBlob();
    }
  } catch (_) {}
  return null;
}

function tryFindImageLoose_(blocks, key) {
  if (!blocks || !blocks.images) return null;
  const lc = key.toLowerCase();
  for (const k of Object.keys(blocks.images)) {
    if (k.toLowerCase() === lc) return blocks.images[k];
  }
  for (const k of Object.keys(blocks.images)) {
    if (lc.includes(k.toLowerCase()) || k.toLowerCase().includes(lc)) return blocks.images[k];
  }
  return null;
}

function tryFindTextLoose_(textMap, key) {
  if (!textMap) return '';
  const want = (key || '').toString().trim().toLowerCase().replace(/\s+/g,' ');
  if (!want) return '';
  for (const k in textMap) {
    if (!Object.prototype.hasOwnProperty.call(textMap, k)) continue;
    if (k.toLowerCase().trim().replace(/\s+/g,' ') === want) return textMap[k];
  }
  for (const k in textMap) {
    if (!Object.prototype.hasOwnProperty.call(textMap, k)) continue;
    const norm = k.toLowerCase().trim().replace(/\s+/g,' ');
    if (norm.includes(want) || want.includes(norm)) return textMap[k];
  }
  return '';
}

function findLogoBlob_(cfg, blocks) {
  if (blocks && blocks.images) {
    for (const k of Object.keys(blocks.images)) {
      if (String(k).toLowerCase().includes('logo')) return blocks.images[k];
    }
  }

  if (cfg) {
    for (const k of Object.keys(cfg)) {
      const kl = k.toLowerCase();
      if (kl.includes('logo') && (kl.includes('image') || kl.includes('url') || kl.includes('file') || kl.includes('id'))) {
        const blob = resolveBlobFromRef_(cfg[k]);
        if (blob) return blob;
      }
    }
  }

  const aliases = [
    'Logo','Logo Image','Logo Header','Header Logo','Letterhead Logo',
    'Logo_Image','Logo_Image_File_ID','Logo_Image_URL',
    'Logo_Header_File_ID','Logo_Header_URL'
  ];
  for (const alias of aliases) {
    const v = cfg && cfg[alias];
    if (v) {
      const blob = resolveBlobFromRef_(v);
      if (blob) return blob;
    }
  }
  return null;
}

function getImageFromKey_(keyRaw, cfg, blocks) {
  const key = String(keyRaw || '').trim();
  if (!key) return null;

  const fromBlocksExact = (blocks && blocks.images && (blocks.images[key] || blocks.images[keyRaw])) || null;
  if (fromBlocksExact) return fromBlocksExact;

  const norm = key.toLowerCase();
  if (norm.includes('logo')) {
    const b = findLogoBlob_(cfg, blocks);
    if (b) return b;
  }

  const aliasSets = {
    sig1: ['Signature_1_Image','Signature_1_Image_File_ID','Signature_1_Image_URL','Signature_Image','Signature_Image_File_ID','Signature_Image_URL'],
    sig2: ['Signature_2_Image','Signature_2_Image_File_ID','Signature_2_Image_URL']
  };

  if (/signature\s*1/.test(norm) || /signature(?!.*2)/.test(norm)) {
    for (const k of aliasSets.sig1) {
      const v = (cfg && cfg[k]) ? String(cfg[k]).trim() : '';
      if (!v) continue;
      const b = resolveBlobFromRef_(v);
      if (b) return b;
    }
  }
  if (/signature\s*2/.test(norm)) {
    for (const k of aliasSets.sig2) {
      const v = (cfg && cfg[k]) ? String(cfg[k]).trim() : '';
      if (!v) continue;
      const b = resolveBlobFromRef_(v);
      if (b) return b;
    }
  }

  const loose = tryFindImageLoose_(blocks, key);
  return loose || null;
}

function ensureSavedInFolder_(pdfBlob, folder) {
  const name = pdfBlob.getName();
  const files = folder.getFilesByName(name);
  while (files.hasNext()) { const f = files.next(); try { f.setTrashed(true); } catch(e){} }
  const f = folder.createFile(pdfBlob);
  return f.getId();
}

function safeBlobFromFileId_(fileId) {
  if (!fileId) return null;
  try {
    const blob = DriveApp.getFileById(String(fileId).trim()).getBlob();
    if (blob && blob.getBytes && blob.getBytes().length > 0) return blob;
  } catch (e) {}
  return null;
}

function findLatestCountyArtifactsFromFolder_(county, cfg) {
  const folderId = cfg.Counties_Output_Folder_ID || SETTINGS.OUTPUT.COUNTIES_FOLDER_ID;
  const folder = DriveApp.getFolderById(folderId);
  const slugs = countySlugVariants_(county);
  let bestPdf = null, bestPdfTime = 0;
  let bestCsv = null, bestCsvTime = 0;

  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const name = (f.getName()||'').toLowerCase();
    const updated = +f.getLastUpdated();

    if (!/^county[_\s-]/.test(name)) continue;

    let matchesSlug = false;
    for (const s of slugs) {
      if (name.includes(`county_${s}_`) || name.includes(`county-${s}-`) || name.includes(`county ${s} `)) {
        matchesSlug = true; break;
      }
    }
    if (!matchesSlug) continue;

    if (name.endsWith('.pdf')) {
      if (updated > bestPdfTime) { bestPdf = f; bestPdfTime = updated; }
    } else if (name.endsWith('.csv')) {
      if (updated > bestCsvTime) { bestCsv = f; bestCsvTime = updated; }
    }
  }

  return {
    pdfBlob: bestPdf ? bestPdf.getBlob() : null,
    csvBlob: bestCsv ? bestCsv.getBlob() : null
  };
}

function ensureInlineLogoForCid_(html, inlineImages, cfg, blocks) {
  const needs = /cid:logo_header/i.test(String(html));
  const has   = inlineImages && Object.prototype.hasOwnProperty.call(inlineImages, 'logo_header');
  if (needs && !has) {
    const logoBlob = findLogoBlob_(cfg, blocks);
    if (logoBlob) {
      inlineImages = inlineImages || {};
      inlineImages['logo_header'] = logoBlob;
    }
  }
  return inlineImages || {};
}

function cleanInlineImages_(bag) {
  const out = {};
  if (!bag || typeof bag !== 'object') return out;
  Object.keys(bag).forEach(k => {
    if (k.indexOf('__') === 0) return;
    const v = bag[k];
    if (!v) return;
    try {
      const blob = (typeof v.getBlob === 'function') ? v.getBlob() : v;
      if (blob && typeof blob.getBytes === 'function') out[k] = blob;
    } catch (_) {}
  });
  return out;
}

function trashFile_(fileId) {
  try {
    if (Drive.Files && typeof Drive.Files.trash === 'function') {
      Drive.Files.trash(fileId);
      return;
    }
  } catch (e) {}
  try {
    if (Drive.Files && typeof Drive.Files.update === 'function') {
      Drive.Files.update({ trashed: true }, fileId);
    }
  } catch (e) {}
}

function createGoogleDocFromHtml_(html, name, folderId) {
  const blobHtml = Utilities.newBlob(html, 'text/html', `${name}.html`);

  if (Drive.Files && typeof Drive.Files.insert === 'function') {
    const resource = { title: name, mimeType: MimeType.GOOGLE_DOCS, parents: [{ id: folderId }] };
    const gdoc = Drive.Files.insert(resource, blobHtml, {convert: true, supportsTeamDrives: true});
    return gdoc.id;
  }

  const resource = { name, mimeType: 'application/vnd.google-apps.document', parents: [folderId] };
  const gdoc = Drive.Files.create(resource, blobHtml, {supportsAllDrives: true});
  return gdoc.id;
}

function exportGDocToPdf_(fileId, outName) {
  try {
    if (Drive.Files && typeof Drive.Files.export === 'function') {
      const resp = Drive.Files.export(fileId, 'application/pdf');
      const blob = (resp.getBlob ? resp.getBlob() : resp);
      blob.setName(outName);
      return blob;
    }
  } catch (ignored) {}

  const token = ScriptApp.getOAuthToken();
  const url =
    'https://www.googleapis.com/drive/v3/files/' +
    encodeURIComponent(fileId) +
    '/export?mimeType=application/pdf&supportsAllDrives=true&alt=media';

  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Drive export failed (' + code + '): ' + resp.getContentText());
  }

  const blob = resp.getBlob();
  blob.setName(outName);
  return blob;
}

function insertHeaderLogo_(gdocId, cfg, blocks) {
  const logoBlob = findLogoBlob_(cfg, blocks);
  if (!logoBlob) return;

  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const doc = DocumentApp.openById(gdocId);
      let header = doc.getHeader();
      if (!header) header = doc.addHeader();
      header.clear();

      const img = header.appendImage(logoBlob);

      const widthPx = Number(
        (cfg.Logo_Header_Width_Px && String(cfg.Logo_Header_Width_Px).trim()) ||
        (cfg.Logo_Max_Width_Px    && String(cfg.Logo_Max_Width_Px).trim()) ||
        120
      );
      img.setWidth(widthPx);

      const p = img.getParent().asParagraph();
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      p.setSpacingBefore(0);
      p.setSpacingAfter(0);
      p.setLineSpacing(1.0);

      doc.saveAndClose();
      return;
    } catch (e) {
      Utilities.sleep(600);
      if (attempt === 2) throw e;
    }
  }
}

function postProcessDocForPdf_(gdocId, cfg, blocks) {
  const MARK = '###PAGE_BREAK_BETWEEN_LETTERS###';

  const doc  = DocumentApp.openById(gdocId);
  const body = doc.getBody();

  for (let i = body.getNumChildren() - 1; i >= 0; i--) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const p = el.asParagraph();
    const text = (p.getText() || '').trim();

    if (text.indexOf(MARK) !== -1) {
      if (i < body.getNumChildren() - 1) {
        body.insertPageBreak(i + 1);
      }
      body.removeChild(el);
    }
  }

  doc.saveAndClose();
}

function toLocalIcsDateTime_(dateVal, timeVal, tz) {
  const d = parseDateFlexible_(dateVal, tz);
  const t = parseTimeFlexible_(timeVal, tz);
  const y = String(d.y).padStart(4, '0');
  const m = String(d.m).padStart(2, '0');
  const day = String(d.d).padStart(2, '0');
  const hh = String(t.hh).padStart(2, '0');
  const mm = String(t.mm).padStart(2, '0');
  return `${y}${m}${day}T${hh}${mm}00`;
}

function getEmailSs_(cfg) {
  return SpreadsheetApp.openById(cfg.email_sheet_id);
}

function getAssignmentsSs_(cfg) {
  return SpreadsheetApp.openById(cfg.assignments_sheet_id);
}

/**************************************************************
 * ADDED — County volunteer spreadsheet map + writer helpers
 * These are additive only and do not change existing behavior
 * outside of the one call inserted in cmdGenerateCountyPackets.
 **************************************************************/

function getCountyFileMapSheet_(cfg) {
  const assignmentsSs = getAssignmentsSs_(cfg);
  return getOrCreateWithHeader_(
    assignmentsSs,
    SETTINGS.COUNTY_VOL_SHEET.MAP_TAB_NAME,
    SETTINGS.COUNTY_VOL_SHEET.MAP_HEADERS
  );
}

function readCountySpreadsheetMap_(cfg) {
  const sh = getCountyFileMapSheet_(cfg);
  const vals = getData_(sh);
  if (!vals.length) return new Map();

  const header = vals[0] || [];
  const rows = vals.slice(1);

  const idxCounty = header.indexOf('County');
  const idxSpreadsheetId = header.indexOf('Spreadsheet_ID');
  const out = new Map();

  if (idxCounty < 0 || idxSpreadsheetId < 0) return out;

  rows.forEach(r => {
    const county = normalizeCountyCore_(r[idxCounty] || '');
    const spreadsheetId = extractSpreadsheetId_(r[idxSpreadsheetId] || '');
    if (county && spreadsheetId) out.set(county, spreadsheetId);
  });

  return out;
}

function getCountySpreadsheetIdFromMap_(county, cfg) {
  const map = readCountySpreadsheetMap_(cfg);
  return map.get(normalizeCountyCore_(county)) || '';
}

function extractCountyFromVolunteerSpreadsheetName_(fileName, cfg) {
  const s = String(fileName || '').trim();
  if (!s) return '';

  // Expected format:
  // [County] Volunteer Spreadsheet - [Election Key]
  // Example:
  // Peach Volunteer Spreadsheet - 2026_Primary
  // Jeff Davis Volunteer Spreadsheet - 2026_Primary
  const m = s.match(/^(.*?)\s+Volunteer\s+Spreadsheet\s*-\s*(.+)$/i);
  if (!m || !m[1] || !m[2]) return '';

  const county = normalizeCountyCore_(m[1]);
  const fileElectionKey = String(m[2] || '').trim();
  const activeElectionKey = String((cfg && cfg.Election_Key) || '').trim();

  // If we have an active election key, require an exact match
  if (activeElectionKey && fileElectionKey !== activeElectionKey) return '';

  return county;
}

function fileNameMatchesCounty_(fileName, county, cfg) {
  const parsedCounty = extractCountyFromVolunteerSpreadsheetName_(fileName, cfg);
  const countyKey = normalizeCountyCore_(county);
  if (!parsedCounty || !countyKey) return false;
  return parsedCounty === countyKey;
}

function findCountyVolunteerSpreadsheetInFolderTree_(folder, county, depth, maxDepth, cfg) {
  if (!folder || depth > maxDepth) return null;

  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    if (fileNameMatchesCounty_(f.getName(), county, cfg)) {
      return f;
    }
  }

  const childFolders = folder.getFolders();
  while (childFolders.hasNext()) {
    const child = childFolders.next();
    const found = findCountyVolunteerSpreadsheetInFolderTree_(child, county, depth + 1, maxDepth, cfg);
    if (found) return found;
  }

  return null;
}

function backfillCountyFileMap_FromElectionRoot_(cfg) {
  const rootId = (cfg.Election_Root_Folder_ID || '').trim();
  if (!rootId) throw new Error('Missing Election_Root_Folder_ID in config.');

  const sh = getCountyFileMapSheet_(cfg);
  const existing = readCountySpreadsheetMap_(cfg);

  const root = DriveApp.getFolderById(rootId);
  const matches = [];

  (function walk_(folder, depth, maxDepth) {
    if (!folder || depth > maxDepth) return;

    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

      const name = String(f.getName() || '').trim();
      if (!name) continue;

      const county = extractCountyFromVolunteerSpreadsheetName_(name, cfg);
      if (!county) continue;

      matches.push({
        county,
        spreadsheetId: f.getId(),
        spreadsheetName: name
      });
    }

    const folders = folder.getFolders();
    while (folders.hasNext()) {
      walk_(folders.next(), depth + 1, maxDepth);
    }
  })(root, 0, 4);

  const dedup = new Map();
  matches.forEach(m => {
    if (!dedup.has(m.county) && m.spreadsheetId) dedup.set(m.county, m);
  });

  const header = SETTINGS.COUNTY_VOL_SHEET.MAP_HEADERS;
  const rows = [];

  dedup.forEach(m => {
    if (existing.has(m.county)) return;
    rows.push(rowFromObj_(header, {
      County: m.county,
      Spreadsheet_ID: m.spreadsheetId,
      Spreadsheet_Name: m.spreadsheetName,
      Notes: 'Auto-backfilled from Election Root search'
    }));
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);
  }

  return rows.length;
}

function findCountyVolunteerSpreadsheetFile_(county, cfg) {
  const mappedId = getCountySpreadsheetIdFromMap_(county, cfg);
  if (mappedId) {
    try {
      return DriveApp.getFileById(mappedId);
    } catch (e) {
      throw new Error(`Mapped Spreadsheet_ID for "${county}" is invalid or inaccessible: ${mappedId}`);
    }
  }

  throw new Error(
    `No Spreadsheet_ID mapping found for county "${county}" in ` +
    `"${SETTINGS.COUNTY_VOL_SHEET.MAP_TAB_NAME}".`
  );
}

function formatAssignmentLine_(row, i) {
  const d = row[`A${i}_Date`] || '';
  const s = row[`A${i}_Start`] || '';
  const e = row[`A${i}_End`] || '';
  const ln = row[`A${i}_LocationName`] || '';
  const addr = row[`A${i}_Address`] || '';

  if (!d && !s && !e && !ln && !addr) return '';

  const time = [s || '', e ? (`-${e}`) : ''].join('');
  const pieces = [d, time, ln, addr].filter(Boolean);
  return pieces.join(' | ');
}

function buildCountyAssignedVolunteerRows_(volList) {
  const nowStamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );

  return volList
    .slice()
    .sort((a,b) => String(a.Volunteer_Name || '').localeCompare(String(b.Volunteer_Name || '')))
    .map(v => {
      const out = {
        VAN_ID: v.VAN_ID || '',
        Volunteer_Name: v.Volunteer_Name || '',
        Volunteer_Email: v.Volunteer_Email || '',
        Volunteer_Phone: v.Volunteer_Phone || '',
        'Volunteer Address': (v['Volunteer Address'] || v['Volunteer_Address'] || v.Volunteer_Address || '').toString(),
        Home_County: v.County || '',
        Assignment_County: v.Assignment_County || '',
        Assignment_Count: v.Assignment_Count || '',
        Assignments_JSON: v.Assignments_JSON || '',
        Assignments_JSON_Last_Updated: nowStamp
      };

      for (let i = 1; i <= SETTINGS.MAX_ASSIGNMENTS; i++) {
        out[`Assignment_${i}`] = formatAssignmentLine_(v, i);
      }
      return out;
    });
}

function upsertCountyAssignedVolunteersTab_(county, volList, cfg) {
  const file = findCountyVolunteerSpreadsheetFile_(county, cfg);
  const ss = SpreadsheetApp.openById(file.getId());

  const tabName = SETTINGS.COUNTY_VOL_SHEET.TAB_NAME;
  const header = SETTINGS.COUNTY_VOL_SHEET.HEADERS;

  let sh = ss.getSheetByName(tabName);
  if (!sh) {
    sh = ss.insertSheet(tabName);
  } else {
    const filter = sh.getFilter();
    if (filter) filter.remove();
    sh.clearContents();
    sh.clearFormats();
  }

  sh.getRange(1, 1, 1, header.length).setValues([header]);

  const rows = buildCountyAssignedVolunteerRows_(volList);
  if (rows.length) {
    const values = rows.map(r => rowFromObj_(header, r));
    sh.getRange(2, 1, values.length, header.length).setValues(values);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, header.length);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow > 1) {
    sh.getRange(1, 1, lastRow, lastCol).createFilter();
  }

  return {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    tabName
  };
}

function utilBackfillCountyFileMap() {
  const { cfg } = readConfig_();
  const count = backfillCountyFileMap_FromElectionRoot_(cfg);
  toast_(`County file map backfilled: ${count} row(s) added.`, 5);
}

function utilWriteAssignedVolunteersTabs_AllCounties() {
  const { cfg } = readConfig_();
  const assignmentsSs = getAssignmentsSs_(cfg);
  const master = getTab_(assignmentsSs, SETTINGS.TABS.MASTER);

  const vals = getData_(master);
  if (!vals.length || vals.length < 2) {
    toast_('No master data found.', 5);
    return;
  }

  const header = vals[0] || [];
  const rows = vals.slice(1).map(r => objFromRow_(header, r));
  const byCounty = groupBy_(rows, r => r.Assignment_County || 'Unknown');

  let success = 0;
  let failed = 0;

  Object.keys(byCounty).forEach(cty => {
    try {
      upsertCountyAssignedVolunteersTab_(cty, byCounty[cty], cfg);
      success++;
    } catch (e) {
      failed++;
      log_(cfg, 'Assigned Volunteers tab update failed', {
        county: cty,
        error: e.message
      });
    }
  });

  toast_(`Assigned Volunteers tabs written. Success: ${success}; Failed: ${failed}`, 8);
}

/** ========================================
 *  CORE CONFIGURATION & CONTENT FUNCTIONS
 *  ======================================== */

function ensureSheets_(cfg) {
  const assignmentsSs = getAssignmentsSs_(cfg);

  getOrCreateWithHeader_(assignmentsSs, SETTINGS.TABS.MASTER, SETTINGS.MASTER_HEADERS);
  getOrCreateWithHeader_(assignmentsSs, SETTINGS.TABS.COUNTY_ROLLUP, SETTINGS.COUNTY_HEADERS);
  getOrCreate_(assignmentsSs, SETTINGS.TABS.LOGS);

  /**************************************************************
   * ADDED — ensure county map tab exists
   **************************************************************/
  getOrCreateWithHeader_(
    assignmentsSs,
    SETTINGS.COUNTY_VOL_SHEET.MAP_TAB_NAME,
    SETTINGS.COUNTY_VOL_SHEET.MAP_HEADERS
  );
}


function readConfig_() {
  const MASTER_CONFIG_SPREADSHEET_ID = '1RcSFICZyYN7hXpo-aiEG_ucF-661G57x7cNV2JQsLzA';

  if (!MASTER_CONFIG_SPREADSHEET_ID) {
    throw new Error('Master Election Configuration Spreadsheet ID not set.');
  }

  const masterConfigSs = SpreadsheetApp.openById(MASTER_CONFIG_SPREADSHEET_ID);
  const masterConfigSh = masterConfigSs.getSheetByName('Elections');

  if (!masterConfigSh) {
    throw new Error('Missing "Elections" tab in Master Election Configuration spreadsheet');
  }

  const masterData = masterConfigSh.getRange(
    1, 1,
    Math.max(1, masterConfigSh.getLastRow()),
    Math.max(1, masterConfigSh.getLastColumn())
  ).getDisplayValues();

  const masterHeader = masterData[0] || [];

  const statusCol        = masterHeader.indexOf('Status');
  const emailCompCol     = masterHeader.indexOf('Email Composition Sheet ID');
  const assignmentsCol   = masterHeader.indexOf('Assignments Spreadsheet ID');
  const runYearCol       = masterHeader.indexOf('Run_Year');
  const electionKeyCol =
    masterHeader.indexOf('Election Key') >= 0
      ? masterHeader.indexOf('Election Key')
      : masterHeader.indexOf('Election_Key');
  const electionDateCol  = masterHeader.indexOf('Election_Date');
  const electionTypeCol  = masterHeader.indexOf('Election_Type');
  const timezoneCol      = masterHeader.indexOf('Calendar_Timezone');
  const rootFolderCol    = masterHeader.indexOf('Election Root Folder ID');

  // Data Mapping column (flexible header support)
  const dataMappingCol =
  masterHeader.indexOf('Data Mapping Document ID') >= 0
    ? masterHeader.indexOf('Data Mapping Document ID')
    : masterHeader.indexOf('Data Mapping Spreadsheet ID') >= 0
    ? masterHeader.indexOf('Data Mapping Spreadsheet ID')
    : masterHeader.indexOf('Data Mapping Sheet ID') >= 0
    ? masterHeader.indexOf('Data Mapping Sheet ID')
    : masterHeader.indexOf('Data_Mapping_Sheet_ID');


  if (statusCol < 0) {
    throw new Error('Master Election Configuration must have "Status" column');
  }
  if (emailCompCol < 0) {
    throw new Error('Master Election Configuration must have "Email Composition Sheet ID" column');
  }
  if (assignmentsCol < 0) {
    throw new Error('Master Election Configuration must have "Assignments Spreadsheet ID" column');
  }

  let activeRow = null;
  for (let i = 1; i < masterData.length; i++) {
    const status = String(masterData[i][statusCol]).trim().toLowerCase();
    if (status === 'active') {
      activeRow = masterData[i];
      break;
    }
  }

  if (!activeRow) {
    throw new Error('No active election found in Master Election Configuration (Status = "Active")');
  }

  const cfg = {};

  cfg.email_sheet_id       = extractSpreadsheetId_(activeRow[emailCompCol] || '');
  cfg.assignments_sheet_id = extractSpreadsheetId_(activeRow[assignmentsCol] || '');

  // Data Mapping spreadsheet ID
  cfg.Data_Mapping_Sheet_ID =
    dataMappingCol >= 0 ? extractSpreadsheetId_(activeRow[dataMappingCol] || '') : '';

  // Aliases so downstream helpers resolve it reliably
  cfg.DataMappingSheetId           = cfg.Data_Mapping_Sheet_ID;
  cfg.Data_Mapping_Spreadsheet_ID  = cfg.Data_Mapping_Sheet_ID;
  cfg.data_mapping_spreadsheet_id  = cfg.Data_Mapping_Sheet_ID;

  if (!cfg.email_sheet_id) {
    throw new Error('Email Composition Sheet ID not found for active election');
  }
  if (!cfg.assignments_sheet_id) {
    throw new Error('Assignments Spreadsheet ID not found for active election');
  }
  if (!cfg.Data_Mapping_Sheet_ID) {
    throw new Error('Data Mapping Spreadsheet ID not found for active election');
  }

  cfg.Run_Year                = runYearCol      >= 0 ? (activeRow[runYearCol]      || '') : '';
  cfg.Election_Key            = electionKeyCol  >= 0 ? (activeRow[electionKeyCol]  || '') : '';
  cfg.Election_Date           = electionDateCol >= 0 ? (activeRow[electionDateCol] || '') : '';
  cfg.Election_Type           = electionTypeCol >= 0 ? (activeRow[electionTypeCol] || '') : '';
  cfg.Calendar_Timezone       = timezoneCol     >= 0 ? (activeRow[timezoneCol]     || '') : '';
  cfg.Election_Root_Folder_ID = rootFolderCol   >= 0 ? (activeRow[rootFolderCol]   || '') : '';

  // --- Merge Email Composition Config tab ---
  const emailSs = SpreadsheetApp.openById(cfg.email_sheet_id);
  const configSh = emailSs.getSheetByName(SETTINGS.TABS.CONFIG);

  if (configSh) {
    const configData = configSh.getRange(
      1, 1,
      Math.max(1, configSh.getLastRow()),
      Math.max(2, configSh.getLastColumn())
    ).getDisplayValues();

    for (let i = 1; i < configData.length; i++) {
      const k = String(configData[i][0] || '').trim();
      const v = String(configData[i][1] || '').trim();
      if (k && !cfg[k]) {
        cfg[k] = v;
      }
    }
  }

  cfg.Calendar_Timezone =
    cfg.Calendar_Timezone || Session.getScriptTimeZone() || 'America/New_York';

  // --- Output folder resolution ---
  const derived = resolveOutputFolderIdsFromElectionRoot_(cfg);
    if (derived) {
      // Always prefer folders derived from the active election root
      cfg.Volunteers_Output_Folder_ID = derived.volunteersFolderId;
      cfg.Counties_Output_Folder_ID   = derived.countiesFolderId;
    } else {
      cfg.Volunteers_Output_Folder_ID =
        (cfg.Volunteers_Output_Folder_ID || '').trim() || SETTINGS.OUTPUT.VOLUNTEERS_FOLDER_ID;
      cfg.Counties_Output_Folder_ID =
        (cfg.Counties_Output_Folder_ID || '').trim() || SETTINGS.OUTPUT.COUNTIES_FOLDER_ID;
    }

  return { cfg };
}


function readContentBlocksAdvancedFromEmailSs_(emailSs) {
  if (!emailSs) throw new Error('readContentBlocksAdvancedFromEmailSs_: emailSs is required.');

  const sh = getTab_(emailSs, SETTINGS.TABS.CONTENT_BLOCKS);
  const lastRow = Math.max(1, sh.getLastRow());

  const text = {};
  const images = {};

  if (lastRow < 2) return { text, images };

  const overImgs = sh.getImages ? sh.getImages() : [];
  const imgByCell = new Map();
  if (overImgs && overImgs.forEach) {
    overImgs.forEach(img => {
      try {
        const a = img.getAnchorCell();
        imgByCell.set(`${a.getRow()},${a.getColumn()}`, img.getBlob());
      } catch (_) {}
    });
  }

  for (let r = 2; r <= lastRow; r++) {
    const key = (sh.getRange(r, 1).getDisplayValue() || '').trim();
    if (!key) continue;

    const cell = sh.getRange(r, 2);
    const disp = cell.getDisplayValue();
    const rich = cell.getRichTextValue();

    let html = '';
    if (rich) {
      const runs = rich.getRuns();
      for (const run of runs) {
        const rawTxt = run.getText() || '';
        const txt = escapeHtml_(rawTxt);
        const style = run.getTextStyle ? run.getTextStyle() : null;
        const link = run.getLinkUrl && run.getLinkUrl();

        let piece = txt;
        if (link) {
          piece = `<a href="${link}" target="_blank" rel="noopener noreferrer" style="color:#1155cc;text-decoration:underline;">${txt}</a>`;
        }

        if (style) {
          if (style.isBold && style.isBold()) piece = `<b>${piece}</b>`;
          if (style.isItalic && style.isItalic()) piece = `<i>${piece}</i>`;
          if (style.isUnderline && style.isUnderline() && !link) piece = `<u>${piece}</u>`;
        }

        html += piece;
      }
      html = html.replace(/\n/g, '<br>');
    } else {
      html = escapeHtml_(disp).replace(/\n/g, '<br>');
    }

    if (html) text[key] = html;

    let blob = null;

    const mapKey = `${r},2`;
    if (imgByCell.has(mapKey)) blob = imgByCell.get(mapKey);

    if (!blob) {
      const formula = cell.getFormula();
      if (formula && /^=IMAGE\(/i.test(formula)) {
        const m = formula.match(/=IMAGE\(\s*"([^"]+)"/i);
        if (m && m[1]) {
          try { blob = UrlFetchApp.fetch(m[1]).getBlob(); } catch (_) {}
        }
      }
    }

    if (!blob) {
      const raw = cell.getValue();
      const asText = (typeof raw === 'string' || typeof raw === 'number') ? String(raw) : '';
      if (asText && /^[A-Za-z0-9_-]{25,}$/.test(asText)) {
        try { blob = DriveApp.getFileById(asText).getBlob(); } catch (_) {}
      }
    }

    if (blob) images[key] = blob;
  }

  return { text, images };
}

function readEmailOrder_() {
  const { cfg } = readConfig_();
  const emailSs = SpreadsheetApp.openById(cfg.email_sheet_id);
  
  const TAB = SETTINGS.TABS.P4_EMAIL_ORDER;
  const sh = P4_getEmailOrderTab_(emailSs, cfg);
  if (!sh) throw new Error('Missing tab in Email Composition sheet: ' + TAB);

  const rows = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), Math.max(1, sh.getLastColumn()))
                 .getDisplayValues();
  if (rows.length < 2) return {};

  const header = rows[0].map(h => (h || '').toString().trim());
  const body   = rows.slice(1);

  let labelColIdx = 0;
  for (let c = 0; c < header.length; c++) {
    const h = (header[c] || '').toString().trim().toLowerCase();
    if (h === 'block name' || h === 'block' || h === 'label') { labelColIdx = c; break; }
  }

  const order = {};
  for (let cIdx = 0; cIdx < header.length; cIdx++) {
    if (cIdx === labelColIdx) continue;
    const colName = header[cIdx];
    if (!colName) continue;

    let subjectKeyOrText = '';
    const blocks = [];

    for (let r = 0; r < body.length; r++) {
      const rowLabel = (body[r][labelColIdx] || '').toString().trim();
      const val      = (body[r][cIdx]       || '').toString().trim();
      if (!val) continue;

      if (/\bsubject\b/i.test(rowLabel)) {
        subjectKeyOrText = val;
        continue;
      }
      const m = val.match(/^subject\s*:\s*(.*)$/i);
      if (m) {
        subjectKeyOrText = m[1].trim();
        continue;
      }

      blocks.push(val);
    }

    order[colName] = { subjectKey: subjectKeyOrText, blocks };
  }

  try { Logger.log('ORDER columns: ' + JSON.stringify(Object.keys(order))); } catch (_){}
  return order;
}

var __BRE_CACHE = null;

function getBRESheet_(cfg) {
  const idRaw = (cfg.BRE_Workbook_ID || SETTINGS.DEFAULT_BRE.WORKBOOK_ID || '').trim();
  const tab = cfg.BRE_Merge_Tab_Name || SETTINGS.DEFAULT_BRE.TAB_NAME || 'BRE Merge Sheet';
  const useActive = !idRaw || idRaw.toUpperCase() === 'ACTIVE';
  const ss = useActive ? SpreadsheetApp.getActive() : SpreadsheetApp.openById(idRaw);
  return getTab_(ss, tab);
}

function readBRE_ByCounty_(cfg) {
  const sh = getBRESheet_(cfg);
  const vals = getData_(sh);
  const header = vals[0] || [];
  const out = new Map();
  const idxCounty = header.indexOf('County');
  for (let r=1;r<vals.length;r++) {
    const key = ((idxCounty>=0 ? vals[r][idxCounty] : '') || '').toString().trim();
    if (key) {
      const rec = objFromRow_(header, vals[r]);
      out.set(key, rec);
    }
  }
  return out;
}

function getBREData_(cfg) {
  if (__BRE_CACHE && (Date.now() - __BRE_CACHE.time) < 5*60*1000) return __BRE_CACHE.data;

  const sh = getBRESheet_(cfg);
  const vals = sh.getRange(1,1, sh.getLastRow()||1, sh.getLastColumn()||1).getDisplayValues();
  const header = vals[0] || [];
  const rows   = vals.slice(1);

  const iCounty = header.findIndex(h => String(h).trim().toLowerCase() === 'county');
  const map = new Map();
  rows.forEach(r => {
    const key = normCountyKey_(r[iCounty]);
    if (key) map.set(key, r);
  });

  __BRE_CACHE = { time: Date.now(), data: { header, map } };
  return __BRE_CACHE.data;
}

function getBRE_ForCounty_(cfg, county) {
  if (!county) return null;
  const data = getBREData_(cfg);
  const r = data.map.get(normCountyKey_(county));
  return r ? { breHeader: data.header, breRow: r } : null;
}

function utilClearBRECache(){ __BRE_CACHE = null; toast_('BRE cache cleared.', 3); }

function readContentBlocksAdvanced_() {
  const { cfg } = readConfig_();
  
  if (!cfg.email_sheet_id) {
    throw new Error('Email Composition Spreadsheet ID not configured');
  }
  
  const emailSs = SpreadsheetApp.openById(cfg.email_sheet_id);
  return readContentBlocksAdvancedFromEmailSs_(emailSs);
}

/** ===============================
 *  MERGE TOKEN & ADDRESS HELPERS
 *  =============================== */

function replaceMergeTokens_(txt, row, cfg, blocks, extras) {
  var map = buildMergeMap_(row, cfg, blocks, extras);
  function repToken(match, p1) {
    var key = String(p1).trim();
    return Object.prototype.hasOwnProperty.call(map, key) ? String(map[key]) : match;
  }
  return String(txt)
    .replace(/\*\|\s*([^|]+?)\s*\|\*/g, repToken)
    .replace(/\|\s*([^|]+?)\s*\|/g, repToken)
    .replace(/\{\{\s*([^}]+?)\s*\}\}/g, repToken)
    .replace(/\{\s*([^}]+?)\s*\}/g, repToken)
    .replace(/\[\s*([^\]]+?)\s*\]/g, repToken);
}

function buildMergeMap_(row, cfg, blocks, extras) {
  row = row || {};
  cfg = cfg || {};
  blocks = blocks || { text: {} };
  extras = extras || {};

  // -----------------------------
  // Helpers
  // -----------------------------
  const tz =
    cfg.Calendar_Timezone ||
    Session.getScriptTimeZone() ||
    'America/New_York';

  function pref(cfgKey, cbKey) {
    let v = (cfg[cfgKey] != null && cfg[cfgKey] !== '') ? cfg[cfgKey] : null;
    if (v == null && blocks.text && cbKey) v = blocks.text[cbKey];
    return v || '';
  }

  function normKeyForReserve_(k) {
    return String(k || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, ' ');
  }

  function addIfNew_(map, k, v) {
    if (!k) return;
    if (!Object.prototype.hasOwnProperty.call(map, k)) map[k] = v;
  }

  // -----------------------------
  // Name bits
  // -----------------------------
  const name = String(row.Volunteer_Name || '').trim();
  const parts = name ? name.split(/\s+/) : [];
  const first = parts[0] || '';
  const last = parts.length > 1 ? parts[parts.length - 1] : '';

  // -----------------------------
  // On/During (single vs multi)
  // -----------------------------
  const onDuring = (Number(row.Assignment_Count || 0) > 1) ? 'during' : 'on';

  // -----------------------------
  // Pretty letter date
  // -----------------------------
  const letterDateRaw =
    (cfg.Letter_Date && String(cfg.Letter_Date).trim())
      ? cfg.Letter_Date
      : new Date();
  const letterDate = formatPrettyDate_(letterDateRaw, tz);

  // -----------------------------
  // Volunteer mailing address (DO NOT read plain "Address")
  // -----------------------------
  const addr1 = pick_(
    row,
    'Volunteer Address', 'Volunteer_Address', 'Mailing Address', 'Mailing_Address',
    'Address Line 1', 'Address1'
  );
  const addr2 = pick_(row, 'Address Line 2', 'Address2', 'Apt', 'Unit');
  const city = pick_(row, 'City', 'Volunteer_City', 'Mailing_City');
  const state = pick_(row, 'State', 'St', 'Province');
  const zip = pick_(row, 'Zip', 'ZIP', 'Postal', 'Postal_Code', 'Postal Code');

  let addrBlock = '';
  if (addr1 || addr2 || city || state || zip) {
    const line2 = [city, state].filter(Boolean).join(', ');
    addrBlock = [
      addr1,
      addr2,
      [line2, zip].filter(Boolean).join(' ')
    ].filter(Boolean).join('\n');
  } else if (row.VAN_ID) {
    // optional safety net: fall back to LBJ mailing address by VAN
    const lbj = lookupVolunteerAddressByVan_(row.VAN_ID);
    if (lbj) {
      const line2 = [lbj.city, lbj.state].filter(Boolean).join(', ');
      addrBlock = [
        lbj.addr1,
        lbj.addr2,
        [line2, lbj.zip].filter(Boolean).join(' ')
      ].filter(Boolean).join('\n');
    }
  }

  // -----------------------------
  // County token (row first; optional BRE fallback)
  // -----------------------------
  let countyTok = String(row.Assignment_County || row.County || '').trim();
  if (!countyTok && extras.breHeader && extras.breRow) {
    const iCounty = extras.breHeader.findIndex(
      h => String(h || '').trim().toLowerCase() === 'county'
    );
    if (iCounty >= 0) countyTok = String(extras.breRow[iCounty] || '').trim();
  }

  // -----------------------------
  // Assignment/Credential Dates (from Data Mapping doc)
  // -----------------------------
  const assignmentDatesBullets = getTwoColDatesBullets_(cfg, 'Assignment Dates') || '';
  const credentialDatesBullets = getTwoColDatesBullets_(cfg, 'Credential Dates') || '';

  // -----------------------------
  // Base map (core tokens)
  // -----------------------------
  const map = {
    // Names
    'Volunteer Name': name,
    'First Name': first,
    'FirstName': first,
    'Volunteer First Name': first,
    'First_Name': first,
    'Last Name': last,
    'LastName': last,
    'Volunteer Last Name': last,
    'Last_Name': last,

    // Address (mailing)
    'Address': addrBlock,
    'Mailing Address': addrBlock,
    'Volunteer Address': addrBlock,

    // Dates/titles
    'Letter Date': letterDate,
    'OnDuring': onDuring,
    'On/During': onDuring,
    'On or During': onDuring,

    'Voting Type': pref('Voting_Type', 'Voting Type'),

    'Election Date': pref('Election_Date', 'Election Date'),
    'Election_Date': pref('Election_Date', 'Election Date'),

    'Election Title': pref('Election_Type', 'Election Title'),
    'Election_Type': pref('Election_Type', 'Election Title'),

    'Election Name': pref('Election_Name', 'Election Name') || pref('Election_Type', 'Election Title'),
    'Election_Name': pref('Election_Name', 'Election Name') || pref('Election_Type', 'Election Title'),

    // Credential/Assignment Dates
    'Assignment Dates': assignmentDatesBullets,
    'Credential Dates': credentialDatesBullets,
    'Credential_Dates': credentialDatesBullets,

    // Misc
    'County': countyTok
  };

  // -----------------------------
  // Training + shift dates
  // -----------------------------
  map['Training Deadlines EVIP'] = getTrainingDeadline_(cfg, 'EVIP') || '';
  map['Training Deadlines EDay'] = getTrainingDeadline_(cfg, 'EDay') || '';

  const shiftDates = row.VAN_ID ? getShiftDatesBullets_(row.VAN_ID, cfg) : '';
  map['Shift Dates'] = shiftDates || '';

  // -----------------------------
  // Links from Data Mapping ("Links" section)
  // -----------------------------
  const linkSpecs = [
    { key: 'Poll Watcher Manual Link', aliases: ['Poll Watcher Manual'] },
    { key: 'Poll Watcher FAQ Link', aliases: ['Poll Watcher FAQ'] },
    { key: 'Poll Watcher EDay FAQ Link', aliases: ['Poll Watcher Election Day FAQ Link'] },
    { key: 'Sign Up Form Link', aliases: ['Signup Form Link'] }
  ];

  for (const spec of linkSpecs) {
    const v = getLinkValueFromDataMapping_(cfg, spec.key);
    if (!v) continue;
    map[spec.key] = v;
    (spec.aliases || []).forEach(a => { map[a] = v; });
  }

  // -----------------------------
  // Merge BRE row columns as tokens, without clobbering core tokens
  // -----------------------------
  if (extras.breHeader && extras.breRow) {
    const H = extras.breHeader;
    const R = extras.breRow;

    // tokens we refuse to overwrite (normalized)
    const reserved = new Set(Object.keys(map).map(normKeyForReserve_));

    for (let i = 0; i < H.length; i++) {
      const rawKey = String(H[i] || '').trim();
      if (!rawKey) continue;

      const val = (R[i] != null) ? R[i] : '';

      const normSpaces = rawKey.replace(/\s+/g, ' ').trim();
      const noHyphens = normSpaces.replace(/-/g, ' ');
      const noUnders = normSpaces.replace(/_/g, ' ');
      const squashAll = normSpaces.replace(/[-_]/g, ' ');

      // Always expose BRE-namespaced tokens
      addIfNew_(map, 'BRE ' + normSpaces, val);
      addIfNew_(map, 'BRE ' + noHyphens, val);
      addIfNew_(map, 'BRE ' + noUnders, val);
      addIfNew_(map, 'BRE ' + squashAll, val);

      // Only add plain keys if they won't override core tokens
      const candidates = [rawKey, normSpaces, noHyphens, noUnders, squashAll];
      for (const c of candidates) {
        if (!reserved.has(normKeyForReserve_(c))) addIfNew_(map, c, val);
      }
    }
  }

  return map;
}


function lookupVolunteerAddressByVan_(van) {
  try {
    const {cfg} = readConfig_();
    const assignmentsSs = SpreadsheetApp.openById(cfg.assignments_sheet_id);
    const sh = getTab_(assignmentsSs, SETTINGS.TABS.LBJ_IMPORT);
    const vals = sh.getRange(1,1, sh.getLastRow()||1, sh.getLastColumn()||1).getDisplayValues();
    const header = vals[0] || [];
    const rows   = vals.slice(1);

    const idx = mapHeaderIdxFlexible_(header, {
      VAN_ID: SETTINGS.LBJ_FIELDS.VAN_ID,
      A1: ['Address','Address 1','Street Address','Mailing Address','Volunteer Address'],
      A2: ['Address 2','Apt','Unit','Address Line 2'],
      CITY: ['City','Mailing City','Volunteer City'],
      STATE:['State','St','Province'],
      ZIP:  ['Zip','ZIP','Postal','Postal Code','Postal_Code']
    });

    for (const r of rows) {
      const v = getValIdx_(r, idx.VAN_ID);
      if (v && String(v).trim() === String(van).trim()) {
        return {
          addr1: getValIdx_(r, idx.A1),
          addr2: getValIdx_(r, idx.A2),
          city:  getValIdx_(r, idx.CITY),
          state: getValIdx_(r, idx.STATE),
          zip:   getValIdx_(r, idx.ZIP)
        };
      }
    }
  } catch (_) {}
  return null;
}

function lookupVolunteerAddressByVan_SingleLine_(van) {
  try {
    const {cfg} = readConfig_();
    const assignmentsSs = SpreadsheetApp.openById(cfg.assignments_sheet_id);
    const sh = getTab_(assignmentsSs, SETTINGS.TABS.LBJ_IMPORT);
    const vals = sh.getRange(1,1, sh.getLastRow()||1, sh.getLastColumn()||1).getDisplayValues();
    const header = vals[0] || [];
    const rows   = vals.slice(1);

    const idx = mapHeaderIdxFlexible_(header, {
      VAN_ID: SETTINGS.LBJ_FIELDS.VAN_ID,
      ADDRESS_EXACT: ['Address'],
      A1: ['Address','Address 1','Street Address','Mailing Address','Volunteer Address'],
      A2: ['Address 2','Apt','Unit','Address Line 2'],
      CITY: ['City','Mailing City','Volunteer City','Volunteer City'],
      STATE:['State','St','Province'],
      ZIP:  ['Zip','ZIP','Postal','Postal Code','Postal_Code']
    });

    for (const r of rows) {
      const v = getValIdx_(r, idx.VAN_ID);
      if (v && String(v).trim() === String(van).trim()) {
        const exact = getValIdx_(r, idx.ADDRESS_EXACT);
        if (exact) return exact.replace(/\r?\n/g, ' ').trim();

        const a1   = getValIdx_(r, idx.A1);
        const a2   = getValIdx_(r, idx.A2);
        const city = getValIdx_(r, idx.CITY);
        const st   = getValIdx_(r, idx.STATE);
        const zip  = getValIdx_(r, idx.ZIP);
        const line2 = [city, st].filter(Boolean).join(', ');
        return [a1, a2, [line2, zip].filter(Boolean).join(' ')].filter(Boolean).join(' ').replace(/\s{2,}/g,' ').trim();
      }
    }
  } catch (_) {}
  return '';
}

// Reads date sections from the "4 Assignments" tab
function getTwoColDatesBullets_(cfg, labelName) {
  labelName = labelName || 'Assignment Dates';

  const ssRefRaw = String(
    (cfg && (
      cfg.Data_Mapping_Sheet_ID ||
      cfg.DataMappingSheetId ||
      cfg.Data_Mapping_Doc_ID ||
      cfg['Data Mapping Document ID'] ||
      cfg['Data Mapping Spreadsheet ID'] ||
      cfg.data_mapping_spreadsheet_id
    )) || ''
  ).trim();

  const ssId = extractSpreadsheetId_(ssRefRaw);
  if (!ssId) return '';

  const tabName = (cfg.Assignments_Tab_Name || '4 Assignments').trim();

  let sh;
  try {
    sh = SpreadsheetApp.openById(ssId).getSheetByName(tabName);
  } catch (e) {
    return '';
  }
  if (!sh) return '';

  const grid = sh.getDataRange().getDisplayValues();
  const want = labelName.toLowerCase().replace(/\s+/g, ' ').trim();

  let anchorR = -1, anchorC = -1;
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      const v = String(grid[r][c] || '').trim().toLowerCase().replace(/\s+/g, ' ');
      if (v === want || v.startsWith(want + ':') || v.startsWith(want + ' ')) {
        anchorR = r; anchorC = c; break;
      }
    }
    if (anchorR >= 0) break;
  }
  if (anchorR < 0) return '';

  const out = [];
  for (let r = anchorR + 1; r < grid.length; r++) {
    const left  = String(grid[r][anchorC] || '').trim();        // type
    const right = String(grid[r][anchorC + 1] || '').trim();    // date/value

    // STOP: new section header (e.g., "Links" in col A, blank in col B)
    if (left && !right) break;

    // STOP: fully blank row
    if (!left && !right) break;

    // skip partial rows
    if (!left || !right) continue;

    const pretty = formatPrettyDateWithOrdinal_(
      right,
      cfg.Calendar_Timezone || Session.getScriptTimeZone()
    );

    out.push(`● ${escapeHtml_(left)}: ${escapeHtml_(pretty)}`);
  }

  return out.join('<br>');
}


function getOneColDateListBullets_(cfg, labelName) {
  labelName = labelName || 'EVIP Dates';

  const ssRefRaw = String(
    (cfg && (
      cfg.Data_Mapping_Sheet_ID ||
      cfg.DataMappingSheetId ||
      cfg.Data_Mapping_Doc_ID ||
      cfg['Data Mapping Document ID'] ||
      cfg['Data Mapping Spreadsheet ID'] ||
      cfg.data_mapping_spreadsheet_id
    )) || ''
  ).trim();

  const ssId = extractSpreadsheetId_(ssRefRaw);
  if (!ssId) return '';

  const tabName = (cfg.Assignments_Tab_Name || '4 Assignments').trim();

  let sh;
  try {
    sh = SpreadsheetApp.openById(ssId).getSheetByName(tabName);
  } catch (e) {
    return '';
  }
  if (!sh) return '';

  const grid = sh.getDataRange().getDisplayValues();
  const want = labelName.toLowerCase().replace(/\s+/g, ' ').trim();

  let anchorR = -1, anchorC = -1;
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      const v = String(grid[r][c] || '').trim().toLowerCase().replace(/\s+/g, ' ');
      if (v === want || v.startsWith(want + ':') || v.startsWith(want + ' ')) {
        anchorR = r; anchorC = c; break;
      }
    }
    if (anchorR >= 0) break;
  }
  if (anchorR < 0) return '';

  const out = [];
  const dateCol = anchorC + 1; // dates are in the next column (B relative to label column)

  for (let r = anchorR + 1; r < grid.length; r++) {
    const left  = String(grid[r][anchorC] || '').trim();     // used only to detect section headers
    const date  = String(grid[r][dateCol] || '').trim();

    // STOP: new section header
    if (left && !date) break;

    // STOP: blank line
    if (!left && !date) break;

    // in a one-col list, we only care about the date/value
    if (!date) continue;

    const pretty = formatPrettyDateWithOrdinal_(
      date,
      cfg.Calendar_Timezone || Session.getScriptTimeZone()
    );
    out.push(`● ${escapeHtml_(pretty)}`);
  }

  return out.join('<br>');
}


function formatPrettyDateWithOrdinal_(d, tz) {
  const base = formatPrettyDate_(d, tz); // expects "September 22, 2026"

  const m = base.match(/^([A-Za-z]+)\s+(\d{1,2}),\s+(\d{4})$/);
  if (!m) return base;

  const month = m[1];
  const day = Number(m[2]);
  const year = m[3];

  const suffix =
    (day % 100 >= 11 && day % 100 <= 13) ? 'th' :
    (day % 10 === 1) ? 'st' :
    (day % 10 === 2) ? 'nd' :
    (day % 10 === 3) ? 'rd' : 'th';

  return `${month} ${day}${suffix}, ${year}`;
}


function getShiftDatesBullets_(van, cfg) {
  try {
    // 🔒 HARD-CODED MEC SPREADSHEET ID
    const MEC_SHEET_ID = '1RcSFICZyYN7hXpo-aiEG_ucF-661G57x7cNV2JQsLzA';

    if (!van) return '';

    // ---- Open MEC → Elections ----
    const mecSs = SpreadsheetApp.openById(MEC_SHEET_ID);
    const electionsSh = mecSs.getSheetByName('Elections');
    if (!electionsSh) throw new Error('Missing MEC tab: Elections');

    const eVals = electionsSh.getDataRange().getValues();
    if (eVals.length < 2) return '';

    const eHdr = eVals[0].map(h => normalizeKey_(h));
    const idx = (name) => eHdr.indexOf(normalizeKey_(name));

    const statusCol   = idx('Status');
    const commitIdCol = idx('Commitment Response Sheet ID');

    if (statusCol < 0)   throw new Error('Missing "Status" column in MEC Elections.');
    if (commitIdCol < 0) throw new Error('Missing "Commitment Response Sheet ID" column in MEC Elections.');

    // Prefer the Active election row
    let commitmentSheetId = '';
    for (let r = 1; r < eVals.length; r++) {
      const status = String(eVals[r][statusCol] || '').trim().toLowerCase();
      if (status === 'active') {
        commitmentSheetId = String(eVals[r][commitIdCol] || '').trim();
        break;
      }
    }

    if (!commitmentSheetId) return '';

    // ---- Open Commitment Sheet → Availability Shifts (Parsed) ----
    const ss = SpreadsheetApp.openById(commitmentSheetId);
    const sh = ss.getSheetByName('Availability Shifts (Parsed)');
    if (!sh) throw new Error('Missing tab: Availability Shifts (Parsed)');

    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return '';

    const hdr = vals[0].map(h => normalizeKey_(h));
    const hIdx = (name) => hdr.indexOf(normalizeKey_(name));

    const vanCol   = hIdx('myc_van_id');
    const labelCol = hIdx('shift_label');
    const ansCol   = hIdx('google_form_answer');

    if (vanCol < 0 || labelCol < 0 || ansCol < 0) return '';

    const vanStr = String(van).trim();
    const seen = new Set();
    const out = [];

    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      if (String(row[vanCol]).trim() !== vanStr) continue;

      const label = String(row[labelCol] || '').trim();
      const ans   = String(row[ansCol] || '').trim();

      if (!label || !ans) continue;
      if (ans.toLowerCase() === 'i am not available') continue;

      const key = label + '::' + ans;
      if (seen.has(key)) continue;
      seen.add(key);

      out.push(
        `<div style="margin:4px 0;"><b>Commitment on ${escapeHtml_(label)}:</b> ${escapeHtml_(ans)}</div>`
      );
    }

    return out.join('');

  } catch (e) {
    Logger.log('getShiftDatesBullets_ error: ' + e);
    return '';
  }
}












/**
 * Attempts to parse labels like "Monday, April 27, 2026"
 * Returns { prettyNoYear: "Monday, April 27", sortKey: <ms since epoch> } or null
 */
function tryParsePrettyLabelDate_(label) {
  // If it's "Election Day" / "RUNOFF Election Day" etc, don't parse
  if (!label.includes(',') || /election day/i.test(label)) return null;

  // Remove leading weekday if present, then parse remainder as date
  // Example: "Monday, April 27, 2026" -> "April 27, 2026"
  const parts = label.split(',').map(s => s.trim());
  if (parts.length < 2) return null;

  const weekday = parts[0]; // "Monday"
  const dateStr = parts.slice(1).join(', ').trim(); // "April 27, 2026"

  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return null;

  // Build "Monday, April 27" (no year)
  const month = d.toLocaleString('en-US', { month: 'long' });
  const day = d.getDate(); // numeric
  return {
    prettyNoYear: `${weekday}, ${month} ${day}`,
    sortKey: d.getTime()
  };
}


/** ==============================
 *  BUILD MASTER HELPERS
 *  ============================== */

function renderLinksTableToHtml_(rows) {
  // rows = [["Sign Up Form","https://..."], ...]
  const clean = (v) => String(v ?? '').trim();

  const items = [];
  for (const r of (rows || [])) {
    const label = clean(r[0]);
    const url   = clean(r[1]);

    if (!label && !url) continue;

    if (url) {
      items.push(
        `<li><a href="${escapeAttr_(url)}" target="_blank" rel="noopener noreferrer" ` +
        `style="color:#1155cc;text-decoration:underline;">${escapeHtml_(label || url)}</a></li>`
      );
    } else {
      items.push(`<li>${escapeHtml_(label)}</li>`);
    }
  }

  if (!items.length) return '';
  return `<ul style="margin:8px 0 8px 20px;padding:0;">${items.join('')}</ul>`;
}


function utilTestManualLinkToken() {
  const { cfg } = readConfig_();
  const url = getLinkValueFromDataMapping_(cfg, 'Poll Watcher Manual Link');
  Logger.log('Manual link resolved to: ' + url);
  SpreadsheetApp.getActive().toast('Manual link: ' + (url || '(blank)'), 'Project 4', 8);
}


function readLBJ_AndGenerateYesNo_(cfg) {
  const assignmentsSs = SpreadsheetApp.openById(cfg.assignments_sheet_id);
  const sh = getTab_(assignmentsSs, SETTINGS.TABS.LBJ_IMPORT);

  const lastRow = Math.max(2, sh.getLastRow());
  const lastCol = Math.max(1, sh.getLastColumn());

  const vals  = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const dvals = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const rtx   = (lastRow > 1) ? sh.getRange(2, 1, lastRow - 1, lastCol).getRichTextValues() : [];

  if (vals.length < 2) return [];

  const header = vals[0];
  const rows   = vals.slice(1);
  const drows  = dvals.slice(1);

  const yesCol = header.indexOf('Yes_Link');
  const noCol  = header.indexOf('No_Link');

  const direct   = (cfg[SETTINGS.FORM_LINK.BASE_URL_CONFIG_KEY] || '').trim();
  const emailSs = getEmailSs_(cfg);
  const fromCell = getCellByA1_(emailSs, cfg[SETTINGS.FORM_LINK.BASE_URL_CELL_KEY] || '');

  const baseUrl  = direct || fromCell || '';

  const fieldIdx = mapHeaderIdxFlexible_(header, SETTINGS.LBJ_FIELDS);
  const out = [];

  rows.forEach((r, i) => {
    const dr = drows[i] || r;

    if (fieldIdx.DATE  >= 0) r[fieldIdx.DATE]  = dr[fieldIdx.DATE];
    if (fieldIdx.START >= 0) r[fieldIdx.START] = dr[fieldIdx.START];
    if (fieldIdx.END   >= 0) r[fieldIdx.END]   = dr[fieldIdx.END];

    if (yesCol >= 0 && rtx[i] && rtx[i][yesCol]) {
      const link = rtx[i][yesCol].getLinkUrl && rtx[i][yesCol].getLinkUrl();
      if (link && !String(r[yesCol] || '').startsWith('http')) r[yesCol] = link;
    }
    if (noCol >= 0 && rtx[i] && rtx[i][noCol]) {
      const link = rtx[i][noCol].getLinkUrl && rtx[i][noCol].getLinkUrl();
      if (link && !String(r[noCol] || '').startsWith('http')) r[noCol] = link;
    }

    const haveYes = (yesCol >= 0 && r[yesCol] && String(r[yesCol]).trim());
    const haveNo  = (noCol  >= 0 && r[noCol]  && String(r[noCol]).trim());

    if ((!haveYes || !haveNo) && baseUrl) {
      const first = getValIdx_(r, fieldIdx.FIRST);
      const last  = getValIdx_(r, fieldIdx.LAST);
      const email = getValIdx_(r, fieldIdx.EMAIL);
      const phone = getValIdx_(r, fieldIdx.PHONE);
      if (yesCol >= 0 && !haveYes) r[yesCol] = buildPrefilledLink_(baseUrl, first, last, email, phone, SETTINGS.FORM_LINK.YES_PARAM);
      if (noCol  >= 0 && !haveNo)  r[noCol]  = buildPrefilledLink_(baseUrl, first, last, email, phone, SETTINGS.FORM_LINK.NO_PARAM);
    }

    if (yesCol < 0 && noCol < 0 && baseUrl) {
      const first = getValIdx_(r, fieldIdx.FIRST);
      const last  = getValIdx_(r, fieldIdx.LAST);
      const email = getValIdx_(r, fieldIdx.EMAIL);
      const phone = getValIdx_(r, fieldIdx.PHONE);
      sh.getRange(i + 2, 23).setValue(buildPrefilledLink_(baseUrl, first, last, email, phone, SETTINGS.FORM_LINK.YES_PARAM));
      sh.getRange(i + 2, 24).setValue(buildPrefilledLink_(baseUrl, first, last, email, phone, SETTINGS.FORM_LINK.NO_PARAM));
    }

    out.push(objFromRow_(header, r));
  });

  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  return out;
}

function buildMasterRows_(lbjRows, cfg) {
  const byVan = groupBy_(lbjRows, r => valByAliases_(r, SETTINGS.LBJ_FIELDS.VAN_ID) || '');
  const result = [];

  Object.keys(byVan).forEach(van => {
    if (!van) return;
    const items = byVan[van];
    const any = items[0] || {};

    const first  = valByAliases_(any, SETTINGS.LBJ_FIELDS.FIRST);
    const last   = valByAliases_(any, SETTINGS.LBJ_FIELDS.LAST);
    const email  = valByAliases_(any, SETTINGS.LBJ_FIELDS.EMAIL);
    const phone  = valByAliases_(any, SETTINGS.LBJ_FIELDS.PHONE);
    const homeCounty = valByAliases_(any, SETTINGS.LBJ_FIELDS.COUNTY);
    const assignCountyAny = valByAliases_(any, SETTINGS.LBJ_FIELDS.ASSIGNMENT_COUNTY) || homeCounty;
    const volAddress = valByAliases_(any, SETTINGS.LBJ_FIELDS.VOL_ADDRESS);
    const yesLink= any['Yes_Link'] || '';
    const noLink = any['No_Link']  || '';
    const sheetTag = valByAliases_(any, SETTINGS.LBJ_FIELDS.SHEET_TAG) || inferSheetTag_();
    const volunteerName = [first,last].filter(Boolean).join(' ').trim();

    const assigns = items.map(it => {
      const rowAssignCounty = valByAliases_(it, SETTINGS.LBJ_FIELDS.ASSIGNMENT_COUNTY) || assignCountyAny;
      return {
        date: (valByAliases_(it, SETTINGS.LBJ_FIELDS.DATE) || '').trim(),
        start: (valByAliases_(it, SETTINGS.LBJ_FIELDS.START) || '').trim(),
        end:   (valByAliases_(it, SETTINGS.LBJ_FIELDS.END)   || '').trim(),
        locationName: (valByAliases_(it, SETTINGS.LBJ_FIELDS.LOCATION_NAME) || '').trim(),
        address:      (valByAliases_(it, SETTINGS.LBJ_FIELDS.LOCATION_ADDR) || '').trim(),
        county:       rowAssignCounty
      };
    });

    assigns.sort((a,b)=>
      (a.date||'').localeCompare(b.date||'') ||
      (a.start||'').localeCompare(b.start||'') ||
      (a.locationName||'').localeCompare(b.locationName||'') ||
      (a.address||'').localeCompare(b.address||'')
    );
    const dedup = dedupeAssignments_(assigns);
    const count = dedup.length;
    const docMode = (count <= 1) ? 'SINGLE' : 'TABLE';

    const flat = {};
    dedup.slice(0, SETTINGS.MAX_ASSIGNMENTS).forEach((a, idx) => {
      const p = `A${idx+1}_`;
      flat[p+'Date'] = a.date;
      flat[p+'Start'] = a.start;
      flat[p+'End'] = a.end;
      flat[p+'LocationName'] = a.locationName;
      flat[p+'Address'] = a.address;
    });

    const canon = canonicalizeAssignmentsForHash_(dedup, cfg.Calendar_Timezone);
    const hash  = sha256_(canon);
    const assignmentCountyForRow = (dedup[0] && dedup[0].county) || assignCountyAny || '';

    result.push({
      VAN_ID: String(van).trim(),
      Volunteer_Name: volunteerName,
      Volunteer_Email: email,
      Volunteer_Phone: phone,
      Volunteer_Address: volAddress,
      'Volunteer Address': volAddress,
      County: homeCounty,
      Assignment_County: assignmentCountyForRow,
      Run_Year: String(cfg.Run_Year||''),
      Yes_Link: yesLink,
      No_Link: noLink,
      Assignment_Count: count,
      Doc_Mode: docMode,
      Assignments_JSON: canon,
      Cred_Hash: hash,
      ...flat,
      Volunteer_PDF_File_ID: '',
      ICS_File_Name: '',
      Credential_Sent_On: '',
      Credential_Sent_By: '',
      Source_LBJ_Sheet_Tag: sheetTag,
      Needs_Confirmation_Send: true,
      Needs_Credential_Send: true,
      Errors: ''
    });
  });

  return result;
}

function upsertMaster_(rows, cfg) {
  const ss = getAssignmentsSs_(cfg);
  const sh = getOrCreateWithHeader_(ss, SETTINGS.TABS.MASTER, SETTINGS.MASTER_HEADERS);
  const header   = getHeader_(sh);
  const existing = getData_(sh).slice(1).map(r => objFromRow_(header, r));

  const vanKey_ = v => String(v || '').trim();
  const byVanExisting = new Map(existing.map(r => [vanKey_(r.VAN_ID), r]));
  const writeRows = [];

  rows.forEach(r => {
    const prior = byVanExisting.get(vanKey_(r.VAN_ID));
    if (prior) {
      r.Credential_Sent_On    = prior.Credential_Sent_On || '';
      r.Credential_Sent_By    = prior.Credential_Sent_By || '';
      r.Source_LBJ_Sheet_Tag  = r.Source_LBJ_Sheet_Tag || prior.Source_LBJ_Sheet_Tag || '';
      r.Volunteer_PDF_File_ID = prior.Volunteer_PDF_File_ID || '';

      const priorAssignHash = assignmentsHashFromJson_(prior.Assignments_JSON);
      const newAssignHash   = assignmentsHashFromJson_(r.Assignments_JSON);
      const changed = priorAssignHash !== newAssignHash;

      const prevM1 = asBool_(prior.Needs_Confirmation_Send);
      const prevM2 = asBool_(prior.Needs_Credential_Send);

      r.Needs_Confirmation_Send = prevM1 || changed;
      r.Needs_Credential_Send   = prevM2 || changed;
    } else {
      r.Needs_Confirmation_Send = true;
      r.Needs_Credential_Send   = true;
    }

    writeRows.push(rowFromObj_(header, r));
  });

  if (writeRows.length) {
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }
    sh.getRange(2, 1, writeRows.length, header.length).setValues(writeRows);
  }
}

function getTrainingDeadline_(cfg, which) {
  if (!cfg) return '';

  // --- STEP 1: resolve the Data Mapping spreadsheet ID safely ---
  const rawId =
    cfg.data_mapping_sheet_id ||
    cfg.Data_Mapping_Sheet_ID ||
    cfg.DataMappingSheetId ||
    cfg['Data Mapping Document ID'] ||
    cfg['Data Mapping Spreadsheet ID'] ||
    '';

  const dmId = normalizeSheetId_(rawId);
  if (!dmId) {
    throw new Error(
      'getTrainingDeadline_: Missing or invalid Data Mapping spreadsheet ID. ' +
      'Raw value was: ' + rawId
    );
  }

  // --- STEP 2: open the spreadsheet + tab ---
  const ss = SpreadsheetApp.openById(dmId);

  const tabName = (cfg.Assignments_Tab_Name || '4 Assignments').trim();
  const sh = ss.getSheetByName(tabName);
  if (!sh) {
    throw new Error('getTrainingDeadline_: Missing tab "' + tabName + '" in Data Mapping sheet');
  }

  // --- STEP 3: read the grid ---
  const grid = sh.getDataRange().getDisplayValues();

  // --- STEP 4: find the "Training Deadlines" header ---
  let headerRow = -1;
  for (let r = 0; r < grid.length; r++) {
    const v = String(grid[r][0] || '').trim().toLowerCase();
    if (v === 'training deadlines') {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) return '';

  // --- STEP 5: scan rows under it for EVIP / EDay ---
  const want = String(which || '').trim().toLowerCase();

  for (let r = headerRow + 1; r < grid.length; r++) {
    const k = String(grid[r][0] || '').trim();
    const v = String(grid[r][1] || '').trim();

    // stop at end of section
    if (!k && !v) break;
    if (k && !v) break;

    if (k.toLowerCase() === want) {
      return v || '';
    }
  }

  return '';
}



function refreshCountyRollup_AllCounties_(cfg) {
  const assignmentsSs = getAssignmentsSs_(cfg);
  const master = getTab_(assignmentsSs, SETTINGS.TABS.MASTER);

  const vals = getData_(master);
  if (!vals.length) return;
  const header = vals[0];
  const rows = vals.slice(1).map(r => objFromRow_(header, r));

  const byCountyAll = groupBy_(rows, r => r.Assignment_County || 'Unknown');

  const roll = getOrCreateWithHeader_(assignmentsSs, SETTINGS.TABS.COUNTY_ROLLUP, SETTINGS.COUNTY_HEADERS);
  const rollHeader = SETTINGS.COUNTY_HEADERS;

  if (roll.getLastRow()>1) roll.getRange(2,1,roll.getLastRow()-1, roll.getLastColumn()).clearContent();

  const out = [];
  Object.keys(byCountyAll).forEach(cty => {
    const list = byCountyAll[cty];
    out.push(rowFromObj_(rollHeader, {
      County: cty,
      Run_Year: cfg.Run_Year || '',
      Volunteer_Count: list.length,
      VAN_ID_List: list.map(v=>v.VAN_ID).join(','),
      County_PDF_File_ID: '',
      County_CSV_File_ID: '',
      County_Email_Sent_On: '',
      Errors: ''
    }));
  });

  if (out.length) roll.getRange(2,1,out.length, rollHeader.length).setValues(out);
}

function buildPrefilledLink_(baseUrl, first, last, email, phone, yesOrNo) {
  if (!baseUrl) return '#';
  const ph = SETTINGS.FORM_LINK.PLACEHOLDERS;
  const ids = SETTINGS.FORM_LINK.ENTRY_IDS;

  const hasPlaceholders =
    baseUrl.includes(ph.first) || baseUrl.includes(ph.last) ||
    baseUrl.includes(ph.phone) || baseUrl.includes(ph.email)  || baseUrl.includes(ph.avail);

  if (hasPlaceholders) {
    return baseUrl
      .replace(new RegExp(ph.first, 'g'), encodeURIComponent(first || ''))
      .replace(new RegExp(ph.last,  'g'), encodeURIComponent(last  || ''))
      .replace(new RegExp(ph.phone, 'g'), encodeURIComponent(phone || ''))
      .replace(new RegExp(ph.email, 'g'), encodeURIComponent(email || ''))
      .replace(new RegExp(ph.avail, 'g'), encodeURIComponent(yesOrNo || ''));
  }

  const parts = [];
  if (first)   parts.push(`${ids.first}=${encodeURIComponent(first)}`);
  if (last)    parts.push(`${ids.last}=${encodeURIComponent(last)}`);
  if (phone)   parts.push(`${ids.phone}=${encodeURIComponent(phone)}`);
  if (email)   parts.push(`${ids.email}=${encodeURIComponent(email)}`);
  if (yesOrNo) parts.push(`${ids.avail}=${encodeURIComponent(yesOrNo)}`);

  const sep = baseUrl.includes('?') ? (baseUrl.endsWith('?') || baseUrl.endsWith('&') ? '' : '&') : '?';
  return baseUrl + (parts.length ? (sep + parts.join('&')) : '');
}

/**
 * Normalize keys so "Poll Watcher Manual", "Poll Watcher Manual Link",
 * "*|Poll Watcher Manual Link|*", etc. can match reliably.
 */
function normalizeKey_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\*?\|\s*/g, '')   // strip "*|" and "|"
    .replace(/\s*\|\*?/g, '')
    .replace(/[^a-z0-9]+/g, ' ') // punctuation -> spaces
    .replace(/\s+/g, ' ')
    .trim();
}

/** Open the Data Mapping sheet (Spreadsheet + tab) used for "4 Assignments" */
function openDataMappingAssignmentsSheet_(cfg) {
  const ssRefRaw = String(
    (cfg && (
      cfg.Data_Mapping_Sheet_ID ||
      cfg.DataMappingSheetId ||
      cfg.Data_Mapping_Doc_ID ||
      cfg['Data Mapping Document ID'] ||
      cfg['Data Mapping Spreadsheet ID'] ||
      cfg.data_mapping_spreadsheet_id
    )) || ''
  ).trim();

  const ssId = extractSpreadsheetId_(ssRefRaw);
  if (!ssId) return null;

  const tabName = (cfg.Assignments_Tab_Name || '4 Assignments').trim();

  try {
    const ss = SpreadsheetApp.openById(ssId);
    const sh = ss.getSheetByName(tabName);
    return sh || null;
  } catch (e) {
    return null;
  }
}

/**
 * Read a two-column section (e.g., "Links") into a map:
 *  Left column = label, right column = value.
 * Stops when it hits a fully blank row OR a new section header (left has text, right is blank).
 */
function readTwoColSectionMap_(cfg, sectionName) {
  const sh = openDataMappingAssignmentsSheet_(cfg);
  if (!sh) return {};

  const rng = sh.getDataRange();
  const grid = rng.getDisplayValues();
  const rtx  = rng.getRichTextValues();
  const fml  = rng.getFormulas(); // to catch =HYPERLINK("url","text")

  const want = normalizeKey_(sectionName);

  // Find the section header cell anywhere in the grid
  let anchorR = -1, anchorC = -1;
  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      const v = normalizeKey_(grid[r][c]);
      if (v === want) { anchorR = r; anchorC = c; break; }
    }
    if (anchorR >= 0) break;
  }
  if (anchorR < 0) return {};

  const out = {};

  for (let r = anchorR + 1; r < grid.length; r++) {
    const left  = String(grid[r][anchorC] || '').trim();
    const right = String(grid[r][anchorC + 1] || '').trim();

    // end on blank row
    if (!left && !right) break;

    // end if new section header: text in left col, blank in right col
    if (left && !right) break;

    if (!left) continue;

    // Prefer actual URL from rich text hyperlink (if present)
        // Prefer actual URL from rich text hyperlink (if present)
    let value = right;

    try {
      const rt = rtx[r] && rtx[r][anchorC + 1];
      const url = rt && rt.getLinkUrl && rt.getLinkUrl();
      if (url) value = url;
    } catch (_) {}

    // If it's a HYPERLINK() formula, extract the first argument as URL
    if (!value || value === right) {
      const formula = (fml[r] && fml[r][anchorC + 1]) ? String(fml[r][anchorC + 1]) : '';
      const m = formula.match(/^\s*=\s*HYPERLINK\(\s*"([^"]+)"/i);
      if (m && m[1]) value = m[1];
    }

    value = normalizeUrl_(value);

    if (!value) continue;

    out[normalizeKey_(left)] = String(value).trim();

  }

  return out;
}


/**
 * Get a value from the "Links" section by merge field name or by sheet label,
 * using alias lists when needed.
 */
function getLinkValueFromDataMapping_(cfg, mergeFieldName) {
  // 1) Read the "Links" section from the Data Mapping sheet
  const links = readTwoColSectionMap_(cfg, 'Links');
  if (!links || !Object.keys(links).length) return '';

  // 2) Clean the merge token wrappers if they are still present
  const cleaned = String(mergeFieldName || '')
    .replace(/^\*\|\s*/g, '')
    .replace(/\s*\|\*$/g, '')
    .trim();

  // 3) Normalize
  const k = normalizeKey_(cleaned); // e.g. "poll watcher manual link"

  // 4) Aliases (all normalized-ish human variants)
  //    Keyed by normalized merge-field name.
  const ALIASES = {
    'poll watcher manual link': [
      'poll watcher manual',
      'poll watcher manual link',
      'pw manual',
      'manual'
    ],
    'poll watcher faq link': [
      'poll watcher faq',
      'poll watcher faq link',
      'pw faq',
      'faq'
    ],
    'poll watcher eday faq link': [
      'poll watcher eday faq',
      'poll watcher election day faq',
      'poll watcher e-day faq',
      'eday faq',
      'election day faq'
    ],
    'sign up form link': [
      'sign up form',
      'signup form',
      'sign-up form',
      'pw sign up',
      'poll watcher sign up form'
    ],
    'training form link': [
      'training form',
      'training signup',
      'training sign up form',
      'reschedule training form'
    ],
    'lbj office hours link': [
      'lbj office hours',
      'office hours',
      'lbj hours'
    ],
    'my voter page link': [
      'my voter page - ga secretary of state',
      'my voter page',
      'mvp',
      'ga mvp'
    ]
  };

  // 5) Build candidate lookup keys
  //    - exact normalized merge field
  //    - version without trailing " link"
  //    - any explicit aliases for that key
  const kNoLink = k.replace(/\s+link$/i, '').trim();

  const candidatesRaw = []
    .concat([k])
    .concat(kNoLink && kNoLink !== k ? [kNoLink] : [])
    .concat(ALIASES[k] || []);

  // 6) Look up against the map (map keys are already normalized by readTwoColSectionMap_)
  for (const cand of candidatesRaw) {
    const key = normalizeKey_(cand);
      const v = links[key];
    if (v) return normalizeUrl_(String(v).trim());

  }

  return '';
}



function TEST_trainingDeadlinesMerge() {
  const out = readConfig_();        // <-- known-good
  const cfg = out.cfg;

  const row = {
    Volunteer_Name: 'Test Person',
    Assignment_Count: 1,
    Assignment_County: 'Fulton'
  };

  const map = buildMergeMap_(row, cfg, { text: {} }, {});

  Logger.log('EVIP deadline => %s', map['Training Deadlines EVIP']);
  Logger.log('EDay deadline => %s', map['Training Deadlines EDay']);

  const sample =
    'EVIP: *|Training Deadlines EVIP|* | ' +
    'EDay: *|Training Deadlines EDay|*';

  const rendered = sample
    .replace('*|Training Deadlines EVIP|*', map['Training Deadlines EVIP'] || '')
    .replace('*|Training Deadlines EDay|*', map['Training Deadlines EDay'] || '');

  Logger.log('Rendered => %s', rendered);
}


function normalizeSheetId_(v) {
  const s = String(v || '').trim();
  if (!s) return '';

  // If they pasted a full URL, extract the /d/<ID>/ part
  const m = s.match(/\/d\/([a-zA-Z0-9-_]{25,})/);
  if (m && m[1]) return m[1];

  // If they pasted an ID directly, keep it
  const m2 = s.match(/^[a-zA-Z0-9-_]{25,}$/);
  if (m2) return s;

  // Fallback: sometimes IDs appear as a long token in text
  const m3 = s.match(/([a-zA-Z0-9-_]{25,})/);
  return (m3 && m3[1]) ? m3[1] : '';
}


function TEST_shiftDates() {
  const out = readConfig_();
  const cfg = out.cfg;

  const row = { VAN_ID: '12345' }; // use your real VAN
  const html = getShiftDatesBullets_(row.VAN_ID, cfg);

  Logger.log('Shift Dates => %s', html);

  const sample = 'Dates to Remember<br>*|Shift Dates|*';
  const rendered = sample.replace('*|Shift Dates|*', html || '');
  Logger.log('Rendered => %s', rendered);
}

function getCfgValue_(cfg, aliases) {
  if (!cfg) return '';

  // direct hits
  for (const a of aliases) {
    if (cfg[a] !== undefined && cfg[a] !== null && String(cfg[a]).trim() !== '') {
      return String(cfg[a]).trim();
    }
  }

  // normalized + contains hits
  const norm = (s) => normalizeKey_(String(s || ''));
  const want = aliases.map(a => norm(a));

  const keys = Object.keys(cfg);
  for (const k of keys) {
    const kN = norm(k);
    if (want.includes(kN)) {
      const v = cfg[k];
      if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
    }
  }

  // contains match (last resort)
  for (const k of keys) {
    const kN = norm(k);
    if (want.some(w => kN.includes(w) || w.includes(kN))) {
      const v = cfg[k];
      if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
    }
  }

  return '';
}


function utilDebugOutputFolders() {
  const { cfg } = readConfig_();

  Logger.log('Election_Root_Folder_ID: ' + (cfg.Election_Root_Folder_ID || ''));
  Logger.log('Volunteers_Output_Folder_ID: ' + (cfg.Volunteers_Output_Folder_ID || ''));
  Logger.log('Counties_Output_Folder_ID: ' + (cfg.Counties_Output_Folder_ID || ''));

  let volName = '';
  let ctyName = '';

  try {
    volName = DriveApp.getFolderById(cfg.Volunteers_Output_Folder_ID).getName();
  } catch (e) {
    volName = 'ERROR: ' + e.message;
  }

  try {
    ctyName = DriveApp.getFolderById(cfg.Counties_Output_Folder_ID).getName();
  } catch (e) {
    ctyName = 'ERROR: ' + e.message;
  }

  Logger.log('Volunteers output folder name: ' + volName);
  Logger.log('Counties output folder name: ' + ctyName);

  SpreadsheetApp.getUi().alert(
    'Election Root Folder ID:\n' + (cfg.Election_Root_Folder_ID || '') +
    '\n\nVolunteers Output Folder ID:\n' + (cfg.Volunteers_Output_Folder_ID || '') +
    '\nFolder Name: ' + volName +
    '\n\nCounties Output Folder ID:\n' + (cfg.Counties_Output_Folder_ID || '') +
    '\nFolder Name: ' + ctyName
  );
}

/** ===============================
 *  RENDERING (ORDER-BASED HTML)
 *  =============================== */

function getOrderTemplate_(order, name) {
  if (!order || !name) return { subjectKey: '', blocks: [] };
  const target = String(name).trim().toLowerCase();

  for (const k in order) {
    if (Object.prototype.hasOwnProperty.call(order, k)) {
      if (String(k).trim().toLowerCase() === target) return order[k];
    }
  }
  for (const k in order) {
    if (Object.prototype.hasOwnProperty.call(order, k)) {
      if (String(k).toLowerCase().includes(target)) return order[k];
    }
  }
  return { subjectKey: '', blocks: [] };
}

function renderFromOrder_(templateColName, row, blocks, order, cfg, extras) {
  const tpl = getOrderTemplate_(order, templateColName) || { subjectKey:'', blocks:[] };

  let _extras = extras || null;
  if ((!_extras || !_extras.breHeader) && row && (row.Assignment_County || row.County)) {
    const by = row.Assignment_County || row.County;
    const found = getBRE_ForCounty_(cfg, by);
    if (found) _extras = Object.assign({}, extras || {}, found);
  }

  let subject = '';
  if (tpl.subjectKey) {
    const fromCB = blocks && blocks.text && blocks.text[tpl.subjectKey];
    subject = fromCB ? stripHtml_(fromCB) : stripHtml_(tpl.subjectKey);
  }

  const useDataUrls = !!(_extras && _extras.imageMode === 'dataurl');
  const bag = { __useCid: !useDataUrls };

  const parts = [];
  if (tpl.blocks && tpl.blocks.length) {
    const seq = tpl.blocks.slice();
    for (let i = 0; i < seq.length - 1; i++) {
      if (/^signature\s*1$/i.test(seq[i]) && /^signature\s*2$/i.test(seq[i + 1])) {
        seq.splice(i, 2, 'Signatures Row');
        break;
      }
    }
    for (const key of seq) {
      const html = blockToHtml_(key, row, blocks, cfg, _extras || extras, bag);
      if (html) parts.push(wrapWithSpacer_(html));
    }
  }

  if (!parts.length) {
    const spaced = [];
    if (row) {
      const first = ((row.Volunteer_Name||'').trim().split(/\s+/)[0]) || 'there';
      spaced.push(wrapWithSpacer_(`<p>Hi ${escapeHtml_(first)},</p>`));
      const multi = Number(row.Assignment_Count || 0) > 1;
      spaced.push(wrapWithSpacer_(multi ? renderAssignmentTable_Multiple_(row) : renderAssignmentSingle_(row)));
      const confirm = blockToHtml_('Assignment Confirmation', row, blocks, cfg, _extras, bag);
      if (confirm) spaced.push(wrapWithSpacer_(confirm));
    } else if (_extras && _extras.breRow) {
      spaced.push(wrapWithSpacer_(`<p>Please find the county packet attached.</p>`));
    }
    const htmlBody = `<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.45">${spaced.join('\n')}</div>`;
    const subjFallback = subject || cfg.Volunteer_Subject || cfg.County_Subject || 'Update';
    return { subject: subjFallback, html: htmlBody, inlineImages: useDataUrls ? {} : cleanInlineImages_(bag) };
  }

  const appendix = (_extras && _extras.appendix) ? wrapWithSpacer_(_extras.appendix) : '';
  const htmlBody = `
    <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.45">
      ${parts.join('\n')}
      ${appendix}
    </div>
  `;
  return { subject, html: htmlBody, inlineImages: useDataUrls ? {} : cleanInlineImages_(bag) };
}

function blockToHtml_(keyRaw, row, blocks, cfg, extras, bag) {
  if (!keyRaw) return '';
  const key = String(keyRaw).trim();

  if (/^logo\s*(header|image)?$/i.test(key)) {
    if (extras && extras.forceHeaderLogo) return '';
    const blob = findLogoBlob_(cfg, blocks);
    if (!blob) return '';
    const src = addInlineImage_(bag || {}, 'logo_header', blob);
    const maxW = Number(cfg.Logo_Max_Width_Px || 120);
    return `<div style="text-align:center;margin:8px 0 14px;">
              <img src="${src}" alt="Logo" style="max-width:${maxW}px;width:100%;height:auto;display:inline-block;">
            </div>`;
  }

  if (/^single assignment$/i.test(key)) {
    if (!row) return '';
    return renderAssignmentSingle_(row);
  }
  if (/^multiple assignment$/i.test(key)) {
    if (!row) return '';
    return renderAssignmentTable_Multiple_(row);
  }

  if (/^assignment confirmation$/i.test(key)) {
    if (!row) return '';
    const yesTxt = (blocks && blocks.text && blocks.text['Google Form Link Yes']) || 'Yes, I can make it';
    const noTxt  = (blocks && blocks.text && blocks.text['Google Form Link No'])  || 'No, I cannot';
    let yes = row.Yes_Link || '';
    let no  = row.No_Link  || '';

    const emailSs = getEmailSs_(cfg);
    const baseUrl = (cfg[SETTINGS.FORM_LINK.BASE_URL_CONFIG_KEY] ||
                 getCellByA1_(emailSs, cfg[SETTINGS.FORM_LINK.BASE_URL_CELL_KEY] || '') || '').trim();

    if ((!yes || yes === '#') && baseUrl) {
      const parts = (row.Volunteer_Name || '').trim().split(/\s+/);
      yes = buildPrefilledLink_(baseUrl, parts[0]||'', parts.slice(1).join(' ')||'', row.Volunteer_Phone || '', SETTINGS.FORM_LINK.YES_PARAM);
    }
    if ((!no || no === '#') && baseUrl) {
      const parts = (row.Volunteer_Name || '').trim().split(/\s+/);
      no = buildPrefilledLink_(baseUrl, parts[0]||'', parts.slice(1).join(' ')||'', row.Volunteer_Phone || '', SETTINGS.FORM_LINK.NO_PARAM);
    }
    const left  = yes ? `<a href="${yes}" target="_blank" rel="noopener noreferrer" style="color:#1155cc;text-decoration:underline;">${escapeHtml_(stripHtml_(yesTxt))}</a>` : '';
    const right = no  ? `<a href="${no}"  target="_blank" rel="noopener noreferrer" style="color:#1155cc;text-decoration:underline;">${escapeHtml_(stripHtml_(noTxt))}</a>`  : '';
    if (left && right) return `<div>${left} &nbsp; | &nbsp; ${right}</div>`;
    if (left) return `<div>${left}</div>`;
    if (right) return `<div>${right}</div>`;
    return '';
  }

  if (/^google form link click here$/i.test(key)) {
    const clickTxt = (blocks && blocks.text && blocks.text['Google Form Link Click Here']) || 'click here';
    const href = (row && row.Yes_Link) ? row.Yes_Link : '';
    if (href) return `<a href="${href}" target="_blank" rel="noopener noreferrer" style="color:#1155cc;text-decoration:underline;">${escapeHtml_(stripHtml_(clickTxt))}</a>`;
    return '';
  }

  if (/^signature(\s*\d+)?\s*image$/i.test(key)) {
    const blob = getImageFromKey_(key, cfg, blocks);
    if (!blob) return '';
    const cidName = key.toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'');
    const src = addInlineImage_(bag || {}, cidName, blob);
    const h = Number(cfg.Signature_Height_Px || 72);
    return `<div style="display:inline-block;vertical-align:top;margin:10px 18px 0 0;text-align:left;">
              <img src="${src}" alt="${escapeHtml_(key)}" style="height:${h}px;display:block;">
            </div>`;
  }

  if (/^signatures?(?:\s*(row|side\s*by\s*side))?$/i.test(key)) {
    function sig(n) {
      const blob  = getImageFromKey_(`Signature ${n} Image`, cfg, blocks);
      const src   = blob ? addInlineImage_(bag || {}, `signature_${n}_img`, blob) : '';
      const h     = Number(cfg.Signature_Height_Px || 72);
      const name  = cfg[`Signature_${n}_Name`]  || '';
      const title = cfg[`Signature_${n}_Title`] || '';
      return `
        <div style="display:block;text-align:left;">
          ${src ? `<img src="${src}" alt="Signature ${n}" style="height:${h}px;display:block;margin:0 0 4px 0;">` : ''}
          ${name  ? `<div style="font-weight:700;margin:0;text-align:left;">${escapeHtml_(name)}</div>` : ''}
          ${title ? `<div style="margin:0;text-align:left;">${escapeHtml_(title)}</div>` : ''}
        </div>`;
    }

    const tableHtml = `
      <table role="presentation" style="width:100%;border-collapse:collapse;margin-top:8px;">
        <tr>
          <td style="vertical-align:top;padding:0;border:0;">${sig(1)}</td>
          <td style="vertical-align:top;padding:0;border:0;">${sig(2)}</td>
        </tr>
      </table>`;

    const spacer = (extras && extras.pageBreakAfterSignatures)
      ? `<p data-sig-spacer="1" style="margin:0;font-size:1pt;line-height:1;">&nbsp;</p>`
      : '';

    const marker = (extras && extras.pageBreakAfterSignatures)
      ? `<div style="font-size:0;line-height:0;color:#fff">[[[PAGE_BREAK_AFTER_SIGNATURES]]]</div>`
      : '';

    return tableHtml + spacer + marker;
  }

  if (/^signature\s*[12]$/i.test(key)) {
    const num   = /\d/.test(key) ? String(key).match(/\d/)[0] : '1';
    const blob  = getImageFromKey_(`Signature ${num} Image`, cfg, blocks);
    const name  = cfg[`Signature_${num}_Name`]  || '';
    const title = cfg[`Signature_${num}_Title`] || '';

    const src = blob ? addInlineImage_(bag || {}, `signature_${num}_img`, blob) : '';
    const h   = Number(cfg.Signature_Height_Px || 72);

    const img  = src ? `<img src="${src}" alt="Signature ${num}" style="height:${h}px;max-width:100%;display:block;">` : '';
    const meta = (name || title)
      ? `<div style="margin-top:4px;"><div style="font-weight:700;text-align:left;">${escapeHtml_(name)}</div>${title ? `<div style="text-align:left;">${escapeHtml_(title)}</div>` : ''}</div>`
      : '';

    return `<div class="sig sig-${num}" style="display:inline-block;width:48%;min-width:220px;vertical-align:top;box-sizing:border-box;padding-right:8px;">${img}${meta}</div>`;
  }

  if (/^county pick\s*up$/i.test(key) && extras && extras.breHeader && extras.breRow) {
    const get = (k)=> {
      const i = extras.breHeader.indexOf(k);
      return i>=0 ? (extras.breRow[i]||'') : '';
    };
    const loc  = get('Pick Up Location Name');
    const addr = get('Pick Up Location Address');
    const time = get('Pick Up Time');
    if (loc || addr || time) {
      return `<div><b>Pick-up:</b> ${escapeHtml_(loc)} — ${escapeHtml_(addr)} — ${escapeHtml_(time)}</div>`;
    }
    return '';
  }

  let raw = (blocks && blocks.text && (blocks.text[key] || blocks.text[keyRaw])) || '';
  if (!raw) raw = tryFindTextLoose_(blocks ? blocks.text : null, key);
    if (raw) {
    const merged = replaceMergeTokens_(String(raw), row, cfg, blocks, extras);
    const linkified = linkifyUrlsIfNoAnchors_(merged);
    return `<div>${linkified}</div>`;

  }

  return '';
}

function applyContentBlockDirectives_(text) {
  if (!text) return text;

  // Wrap-block link directive:
  // [[LINK url=...]] ... [[/LINK]]
  text = text.replace(
    /\[\[LINK\s+url=(.+?)\]\]([\s\S]*?)\[\[\/LINK\]\]/g,
    (m, url, inner) => {
      const href = String(url || '').trim().replace(/^"(.*)"$/, '$1');
      const label = String(inner || '').trim();
      if (!href || !label) return label || '';
      return `<a href="${escapeHtml_(href)}">${escapeHtml_(label)}</a>`;
    }
  );

  return text;
}


function renderAssignmentSingle_(row) {
  const d    = row['A1_Date']         || '';
  const s    = row['A1_Start']        || '';
  const e    = row['A1_End']          || '';
  const ln   = row['A1_LocationName'] || '';
  const addr = row['A1_Address']      || '';

  // ── Normalize date to consistent display format regardless of what's stored ──
  let displayDate = d;
  if (d) {
    try {
      const parsed = parseDateFlexible_(d, 'America/New_York');
      displayDate = formatAssignmentDate_(parsed.y, parsed.m, parsed.d);
    } catch(_) {
      displayDate = String(d);
    }
  }

  const timeDisplay = [s, e].filter(Boolean).join(' - ');

  return `
    <div><b>Date:</b> ${escapeHtml_(displayDate)}</div>
    <div><b>Time:</b> ${escapeHtml_(timeDisplay)}</div>
    <div><b>Location:</b> ${escapeHtml_(ln)}</div>
    <div><b>Address:</b> ${escapeHtml_(addr)}</div>
  `;
}

function renderAssignmentTable_Multiple_(row) {
  const tableStyle = 'border-collapse:collapse;width:100%;border:1px solid #d0d7de;';
  const thStyle = 'text-align:left;border:1px solid #d0d7de;padding:8px;background:#f6f8fa;font-weight:700;';
  const tdStyle = 'border:1px solid #d0d7de;padding:8px;vertical-align:top;';

  const rows = [];
  for (let i = 1; i <= SETTINGS.MAX_ASSIGNMENTS; i++) {
    const d    = row[`A${i}_Date`];
    const s    = row[`A${i}_Start`];
    const e    = row[`A${i}_End`];
    const ln   = row[`A${i}_LocationName`];
    const addr = row[`A${i}_Address`];

    if (!d && !s && !ln && !addr) break;

    // ── Normalize date to consistent display format regardless of what's stored ──
    let displayDate = '';
    if (d) {
      try {
        const parsed = parseDateFlexible_(d, 'America/New_York');
        displayDate = formatAssignmentDate_(parsed.y, parsed.m, parsed.d);
      } catch(_) {
        displayDate = String(d); // safe fallback
      }
    }

    const shift = [s || '', e ? ('- ' + e) : ''].filter(Boolean).join(' ');
    rows.push(
      `<tr>
        <td style="${tdStyle}">${escapeHtml_(displayDate)}</td>
        <td style="${tdStyle}">${escapeHtml_(shift)}</td>
        <td style="${tdStyle}">${escapeHtml_(ln || '')}</td>
        <td style="${tdStyle}">${escapeHtml_(addr || '')}</td>
      </tr>`
    );
  }

  return `
    <div><b>Assignments (${row.Assignment_Count || 0})</b></div>
    <table role="presentation" style="${tableStyle}">
      <thead>
        <tr>
          <th style="${thStyle}">Date</th>
          <th style="${thStyle}">Time</th>
          <th style="${thStyle}">Location Name</th>
          <th style="${thStyle}">Location Address</th>
        </tr>
      </thead>
      <tbody>${rows.join('')}</tbody>
    </table>
  `;
}

function renderVolunteerLetterHtml_FromOrder_(row, blocks, order, cfg, opts) {
  const isSingle = String(row.Doc_Mode) === 'SINGLE';
  const colKey = isSingle ? SETTINGS.ORDER_KEYS.LETTER_SINGLE : SETTINGS.ORDER_KEYS.LETTER_MULTI;

  const built = renderFromOrder_(colKey, row, blocks, order, cfg, {
    imageMode: 'dataurl',
    forceHeaderLogo: true
  });

  const allowBodyLogo = (opts && Object.prototype.hasOwnProperty.call(opts, 'injectBodyLogo'))
    ? !!opts.injectBodyLogo
    : true;

  let fallbackTop = '';
  if (allowBodyLogo && asBool_(cfg.Logo_PDF_Fallback_In_Body)) {
    const logoBlob = findLogoBlob_(cfg, blocks);
    if (logoBlob) {
      const src = addInlineImage_({ __useCid:false }, 'logo_pdf_fallback', logoBlob);
      const maxW = Number(
        (cfg.Logo_PDF_Fallback_Width_Px && String(cfg.Logo_PDF_Fallback_Width_Px).trim()) ||
        (cfg.Logo_Header_Width_Px      && String(cfg.Logo_Header_Width_Px).trim()) ||
        (cfg.Logo_Max_Width_Px         && String(cfg.Logo_Max_Width_Px).trim()) ||
        100
      );
      fallbackTop = `
        <div style="text-align:center;margin:6px 0 10px;">
          <img src="${src}" alt="Logo" width="${maxW}"
               style="width:${maxW}px;height:auto;max-width:100%;display:inline-block;">
        </div>`;
    }
  }

  const css = `
    <style>
      @page { size: letter; margin: 0.6in; }
      body { font-family: Arial, Helvetica, sans-serif; font-size: 11pt; color: #222; }
      table.assign { width: 100%; border-collapse: collapse; margin: 8px 0; }
      table.assign th, table.assign td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; vertical-align: top; }
      table.assign th { background:#f0f0f0; }
    </style>
  `;

  return `<!doctype html><html><head><meta charset="utf-8" />${css}</head>
  <body>${fallbackTop}${built.html}</body></html>`;
}

function renderCountyPacketHtml_(county, volList, blocks, cfg, order) {
  const parts = volList.map((row, idx) => {
    const html = renderVolunteerLetterHtml_FromOrder_(row, blocks, order, cfg, {
      injectBodyLogo: true
    });

    const m = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    const inner = m ? m[1] : html;

    const after = (idx < volList.length - 1)
      ? '<p>###PAGE_BREAK_BETWEEN_LETTERS###</p>'
      : '';

    return `${inner}\n${after}`;
  });

  const css = `
    <style>
      @page { size: letter; margin: 0.6in; }
      body {
        font-family: Arial, Helvetica, sans-serif;
        font-size: 11pt;
        color: #222;
        line-height: 1.45;
      }
      table.assign {
        width: 100%;
        border-collapse: collapse;
        margin: 8px 0;
      }
      table.assign th,
      table.assign td {
        border: 1px solid #ccc;
        padding: 6px 8px;
        text-align: left;
        vertical-align: top;
      }
      table.assign th {
        background:#f0f0f0;
      }
    </style>
  `;

  return `<!doctype html>
<html>
<head><meta charset="utf-8">${css}</head>
<body>
${parts.join('\n')}
</body>
</html>`;
}

/** ===========================
 *  HTML TO PDF CONVERSION
 *  =========================== */

function htmlToPdf_(html, fileName, folder, cfg, blocks) {
  const folderId = folder.getId();
  const gdocId = createGoogleDocFromHtml_(html, fileName, folderId);

  Utilities.sleep(400);  // was 800 — Drive is usually ready faster

  postProcessDocForPdf_(gdocId, cfg, blocks);

  Utilities.sleep(150);  // was 300

  const pdfBlob = exportGDocToPdf_(gdocId, fileName);
  try { trashFile_(gdocId); } catch (e) {}
  return pdfBlob;
}

/** ====================
 *  ICS CALENDAR FILES
 *  ==================== */

function buildVolunteerIcs_(row, tz) {
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Org//Poll Watcher//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH'
  ];
  const now = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");

  for (let i = 1; i <= SETTINGS.MAX_ASSIGNMENTS; i++) {
    const dRaw = row[`A${i}_Date`];
    const sRaw = row[`A${i}_Start`];
    const eRaw = row[`A${i}_End`];
    const ln   = row[`A${i}_LocationName`];
    const addr = row[`A${i}_Address`];

    if (!dRaw && !sRaw && !ln) break;

    const dtStart = toLocalIcsDateTime_(dRaw, sRaw, tz);
    const dtEnd   = toLocalIcsDateTime_(dRaw, eRaw, tz);

    const dateKey = dtStart ? dtStart.slice(0, 8) : '00000000';
    const timeKey = dtStart ? dtStart.slice(9, 13) : '0000';
    const uid = `${row.VAN_ID || 'van'}-${dateKey}-${timeKey}@yourorg.org`;

    lines.push(
      'BEGIN:VEVENT',
      `UID:${uid}`,
      `DTSTAMP:${now}`,
      dtStart ? `DTSTART;TZID=${tz}:${dtStart}` : '',
      dtEnd   ? `DTEND;TZID=${tz}:${dtEnd}`     : '',
      'SUMMARY:Poll Watcher Shift',
      `LOCATION:${icsEscape_(`${ln || ''}, ${addr || ''}`)}`,
      'END:VEVENT'
    );
  }

  lines.push('END:VCALENDAR');
  return Utilities.newBlob(
    lines.filter(Boolean).join('\r\n'),
    'text/calendar',
    `Volunteer_${row.VAN_ID || 'van'}.ics`
  );
}

/** ====================
 *  COUNTY CSV BUILDER
 *  ==================== */

function buildCountyCsv_(county, volList, cfg) {
  let maxAssign = 0;
  volList.forEach(v => {
    let count = 0;
    for (let i = 1; i <= SETTINGS.MAX_ASSIGNMENTS; i++) {
      const hasAny = (v[`A${i}_LocationName`] || v[`A${i}_Address`]);
      if (hasAny) count++; else break;
    }
    if (count > maxAssign) maxAssign = count;
  });

  const header = ['Volunteer_Name', 'Volunteer Address', 'Assignment_County'];
  for (let i = 1; i <= maxAssign; i++) header.push(`Location ${i} Name`, `Location ${i} Address`);
  const lines = [header.map(csvEscape_).join(',')];

  volList.forEach(v => {
    const volunteerName = v.Volunteer_Name || '';

    let volAddr = lookupVolunteerAddressByVan_SingleLine_(v.VAN_ID);
    if (!volAddr) {
      volAddr = (v['Volunteer Address'] || v['Volunteer_Address'] || '').toString().replace(/\r?\n/g, ' ').trim();
    }

    const assignCounty = v.Assignment_County || '';

    const row = [volunteerName, volAddr, assignCounty];

    for (let i = 1; i <= maxAssign; i++) {
      const ln   = v[`A${i}_LocationName`] || '';
      const addr = v[`A${i}_Address`]      || '';
      row.push(ln, addr);
    }

    lines.push(row.map(csvEscape_).join(','));
  });

  return Utilities.newBlob(
    lines.join('\r\n'),
    'text/csv',
    `County_${normalizeCountyCore_(county)}_${cfg.Run_Year || ''}.csv`
  );
}

/** ====================
 *  EMAIL SENDER
 *  ==================== */

function sendEmail_({to, cc, subject, htmlBody, attachments=[], inlineImages={}, replyTo, cfg}) {
  const testMode = String(cfg.Test_Mode||'').toUpperCase() === 'TRUE';
  const testRecipients = (cfg.Test_Recipients || '').split(',').map(s=>s.trim()).filter(Boolean);
  let finalTo = to, finalCc = cc || '';
  let finalSubject = subject;

  if (testMode) {
    finalSubject = `[TEST] ${subject}`;
    if (testRecipients.length) { finalTo = testRecipients.join(','); finalCc = ''; }
  }

  const imgs = cleanInlineImages_(inlineImages);
  const opts = {
    name: cfg.From_Name || 'The Georgia Democrats Voter Protection Team',
    htmlBody,
    replyTo: replyTo || '',
    attachments
  };
  if (finalCc) opts.cc = finalCc;
  if (Object.keys(imgs).length) opts.inlineImages = imgs;
  
  log_(cfg, 'sendEmail', {
    testMode,
    finalTo,
    finalCc,
    subject: finalSubject,
    attachmentsCount: (attachments||[]).length,
    inlineImagesCount: Object.keys(imgs||{}).length
  });

  MailApp.sendEmail(finalTo, finalSubject, '(HTML only)', opts);
}

/** =============================
 *  COUNTY ROLLUP LOOKUP HELPERS
 *  ============================= */

function findCountyRollupByCounty_(rollSheet, countyName, cfg) {
  const vals = rollSheet.getRange(1,1, rollSheet.getLastRow()||1, rollSheet.getLastColumn()||1).getDisplayValues();
  const header = vals[0] || [];
  const rows   = vals.slice(1);

  const h = {};
  header.forEach((k,i)=> h[k]=i);

  const yrWanted = (cfg && cfg.Run_Year) ? String(cfg.Run_Year).trim() : '';
  let best = null;

  for (const r of rows) {
    if (!r || !r.length) continue;
    const county = (r[h['County']]||'').trim();
    if (!county || county.toLowerCase() !== String(countyName||'').trim().toLowerCase()) continue;

    if (yrWanted && r[h['Run_Year']] && String(r[h['Run_Year']]).trim() === yrWanted) {
      best = r; break;
    }
    if (!best) best = r;
  }
  return { header, row: best };
}

function findGeneralMMRollupByCounty_(rollSheet, countyName, cfg) {
  const vals = rollSheet.getRange(1,1, rollSheet.getLastRow()||1, rollSheet.getLastColumn()||1).getDisplayValues();
  const header = vals[0] || [];
  const rows   = vals.slice(1);

  const h = {};
  header.forEach((k,i)=> h[k]=i);

  const yrWanted = (cfg && cfg.Run_Year) ? String(cfg.Run_Year).trim() : '';
  let best = null;

  for (const r of rows) {
    if (!r || !r.length) continue;
    const county = (r[h['County']]||'').trim();
    if (!county || county.toLowerCase() !== String(countyName||'').trim().toLowerCase()) continue;

    if (yrWanted && r[h['Run_Year']] && String(r[h['Run_Year']]).trim() === yrWanted) {
      best = r; break;
    }
    if (!best) best = r;
  }
  return { header, row: best };
}

/** ==============================
 *  COMMAND FUNCTIONS
 *  ============================== */

function cmdBuildUpdateMaster() {
  const { cfg } = readConfig_();
  const blocks = readContentBlocksAdvanced_();

  ensureSheets_(cfg);

  const lbjRows    = readLBJ_AndGenerateYesNo_(cfg);
  const masterRows = buildMasterRows_(lbjRows, cfg);

  upsertMaster_(masterRows, cfg);
  refreshCountyRollup_AllCounties_(cfg);

  log_(cfg, 'Build/Update Master', {countLBJ: lbjRows.length, countMaster: masterRows.length});
  toast_('Master updated.', 5);
}


function cmdPreviewDiffs() {
  const {cfg} = readConfig_();
  const assignmentsSs = getAssignmentsSs_(cfg);
  const master = getTab_(assignmentsSs, SETTINGS.TABS.MASTER);
  const vals = getData_(master);
  const header = vals[0] || [];
  const rows = vals.slice(1).map(r => objFromRow_(header, r));
  const flagged = rows.filter(r => String(r.Needs_Credential_Send).toUpperCase() === 'TRUE');

  const byCountyFlagged = groupBy_(flagged, r => r.Assignment_County || 'Unknown');

  const breMap = readBRE_ByCounty_(cfg);
  const breIssues = [];
  for (const county of breMap.keys()) {
    const rec = breMap.get(county);
    if (!rec.To) breIssues.push(`${county}: missing To`);
  }

  const msg =
    `Preview Diffs\n\n`+
    `Flagged volunteers (M2 new/changed): ${flagged.length}\n` +
    `Flagged counties impacted: ${Object.keys(byCountyFlagged).length}\n`+
    (breIssues.length ? `\nBRE Issues:\n- ${breIssues.join('\n- ')}` : `\nBRE Issues: none\n`);
  SpreadsheetApp.getUi().alert(msg);
}

function cmdGenerateVolunteerPDFs() {
  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const assignmentsSs = getAssignmentsSs_(cfg);
  const master = getTab_(assignmentsSs, SETTINGS.TABS.MASTER);

  const valsT = getData_(master);
  const valsD = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
  const header = valsT[0] || [];
  const dataT  = valsT.slice(1);
  const dataD  = valsD.slice(1);

  const outFolder = DriveApp.getFolderById(cfg.Volunteers_Output_Folder_ID);

  let gen = 0;
  for (let i=0; i<dataT.length; i++) {
    const rowT = dataT[i], rowD = dataD[i];
    const objT = objFromRow_(header, rowT);
    const objD = objFromRow_(header, rowD);

    if (String(objT.Needs_Credential_Send).toUpperCase() !== 'TRUE') continue;

    const html = renderVolunteerLetterHtml_FromOrder_(objD, blocks, order, cfg);

    const last = fileSafe_(extractLastName_(objD.Volunteer_Name));
    const pdfBlob = htmlToPdf_(html, `Credential_${last}.pdf`, outFolder, cfg, blocks);

    const fileId = ensureSavedInFolder_(pdfBlob, outFolder);
    const pdfCol = header.indexOf('Volunteer_PDF_File_ID');
    if (pdfCol >= 0) {
      rowT[pdfCol] = fileId;
      master.getRange(i+2, 1, 1, rowT.length).setValues([rowT]);
      gen++;
    }
  }
  log_(cfg, 'Generate Volunteer PDFs', {generated: gen});
  toast_(`Volunteer PDFs generated: ${gen}`, 5);
}

function cmdGenerateCountyPackets() {
  const CHUNK_SIZE = 8; // counties per run — tune based on avg volunteer count
  const props = PropertiesService.getScriptProperties();
  const resumeFrom = parseInt(props.getProperty('P4_COUNTY_PKT_RESUME') || '0', 10);

  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const ss = getAssignmentsSs_(cfg);
  const master = getTab_(ss, SETTINGS.TABS.MASTER);

  const valsT = getData_(master);
  const valsD = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
  const header = valsT[0] || [];
  const dataT  = valsT.slice(1).map(r => objFromRow_(header, r));
  const dataD  = valsD.slice(1).map(r => objFromRow_(header, r));

  const byCountyDisp = groupBy_(dataD, r => r.Assignment_County || 'Unknown');
  const countyNames  = Object.keys(byCountyDisp).sort();
  const outFolder = DriveApp.getFolderById(cfg.Counties_Output_Folder_ID || SETTINGS.OUTPUT.COUNTIES_FOLDER_ID);
  const roll = getOrCreateWithHeader_(ss, SETTINGS.TABS.COUNTY_ROLLUP, SETTINGS.COUNTY_HEADERS);
  const rollHeader = SETTINGS.COUNTY_HEADERS;

  if (resumeFrom === 0 && roll.getLastRow() > 1) {
    roll.getRange(2, 1, roll.getLastRow()-1, roll.getLastColumn()).clearContent();
  }

  const slice = countyNames.slice(resumeFrom, resumeFrom + CHUNK_SIZE);
  let created = 0;

  for (const cty of slice) {
    const listDisp = byCountyDisp[cty].sort((a,b)=> (a.Volunteer_Name||'').localeCompare(b.Volunteer_Name||''));
    const bigHtml  = renderCountyPacketHtml_(cty, listDisp, blocks, cfg, order);
    const pdfBlob  = htmlToPdf_(bigHtml, `County_${cty}_${cfg.Run_Year||''}.pdf`, outFolder, cfg, blocks);
    const pdfId    = ensureSavedInFolder_(pdfBlob, outFolder);

    const listTyped = dataT.filter(r => (r.Assignment_County||'Unknown') === cty);
    const csvBlob   = buildCountyCsv_(cty, listTyped, cfg);
    const csvFile   = outFolder.createFile(csvBlob);

    roll.appendRow(rowFromObj_(rollHeader, {
      County: cty,
      Run_Year: cfg.Run_Year || '',
      Volunteer_Count: listDisp.length,
      VAN_ID_List: listDisp.map(v=>v.VAN_ID).join(','),
      County_PDF_File_ID: pdfId,
      County_CSV_File_ID: csvFile.getId(),
      County_Email_Sent_On: '',
      Errors: ''
    }));
    created++;
  }

  const nextStart = resumeFrom + slice.length;
  const remaining = countyNames.length - nextStart;

  if (remaining > 0) {
    props.setProperty('P4_COUNTY_PKT_RESUME', String(nextStart));
    toast_(`Batch done: ${created} counties (${nextStart}/${countyNames.length}). Run again to continue.`, 8);
  } else {
    props.deleteProperty('P4_COUNTY_PKT_RESUME');
    log_(cfg, 'Generate County Packets', { counties: nextStart });
    toast_(`All county packets complete: ${nextStart} total.`, 5);
  }
}

function cmdSendVolunteerM1() {
  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const ss = getAssignmentsSs_(cfg);
  const master = getTab_(ss, SETTINGS.TABS.MASTER);

  const valsT = getData_(master);
  const valsD = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
  const header = valsT[0] || [];
  const dataT  = valsT.slice(1);
  const dataD  = valsD.slice(1);

  const idxFlagM1 = header.indexOf('Needs_Confirmation_Send');
  if (idxFlagM1 < 0) { toast_('Missing Needs_Confirmation_Send column.', 5); return; }

  const maxBatch = Number(cfg.Max_Batch_Size || 400);
  let sent = 0;
  const dirtyRows = []; // ← collect writes here instead of inside loop

  for (let i = 0; i < dataT.length; i++) {
    if (sent >= maxBatch) break;
    const rowT = dataT[i], rowD = dataD[i];
    const objT = objFromRow_(header, rowT);
    const objD = objFromRow_(header, rowD);

    if (String(objT.Needs_Confirmation_Send).toUpperCase() !== 'TRUE') continue;
    if (!objD.Volunteer_Email) continue;

    const isSingle = String(objD.Doc_Mode) === 'SINGLE';
    const key = isSingle ? SETTINGS.ORDER_KEYS.VOL_M1_SINGLE : SETTINGS.ORDER_KEYS.VOL_M1_MULTI;

    const built = renderFromOrder_(key, objD, blocks, order, cfg);
    sendEmail_({
      to: objD.Volunteer_Email,
      subject: built.subject || (cfg.Volunteer_Subject || SETTINGS.SUBJECTS.VOL_M1),
      htmlBody: built.html,
      replyTo: cfg.Reply_To || cfg.From_Email,
      cfg,
      inlineImages: built.inlineImages
    });

    rowT[idxFlagM1] = false;
    dirtyRows.push({ sheetRow: i + 2, data: rowT }); // ← stage the write
    sent++;
  }

  // ── Single bulk write pass ──────────────────────────────────────────────────
  for (const w of dirtyRows) {
    master.getRange(w.sheetRow, 1, 1, w.data.length).setValues([w.data]);
  }

  log_(cfg, 'Send Volunteer M1', { sent });
  toast_(`M1 sent: ${sent}`, 5);
}

function cmdSendVolunteerM2() {
  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const ss = getAssignmentsSs_(cfg);
  const master = getTab_(ss, SETTINGS.TABS.MASTER);

  const valsT = getData_(master);
  const valsD = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
  const header = valsT[0] || [];
  const dataT  = valsT.slice(1);
  const dataD  = valsD.slice(1);

  const tz = cfg.Calendar_Timezone || 'America/New_York';
  const outFolder = DriveApp.getFolderById(cfg.Volunteers_Output_Folder_ID || SETTINGS.OUTPUT.VOLUNTEERS_FOLDER_ID);

  const maxBatch = Number(cfg.Max_Batch_Size || 400);
  let sent = 0;
  const dirtyRows = []; // ← stage all sheet writes here; flush once after the loop

  for (let i = 0; i < dataT.length; i++) {
    if (sent >= maxBatch) break;

    const rowT = dataT[i], rowD = dataD[i];
    const objT = objFromRow_(header, rowT);
    const objD = objFromRow_(header, rowD);

    if (String(objT.Needs_Credential_Send).toUpperCase() !== 'TRUE') continue;
    if (!objD.Volunteer_Email) continue;

    // ensure PDF
    let pdfBlob = null;
    if (objT.Volunteer_PDF_File_ID) {
      try { pdfBlob = DriveApp.getFileById(objT.Volunteer_PDF_File_ID).getBlob(); } catch(e){}
    }
    if (!pdfBlob) {
      const htmlLetter = renderVolunteerLetterHtml_FromOrder_(objD, blocks, order, cfg);
      const last = fileSafe_(extractLastName_(objD.Volunteer_Name));
      pdfBlob = htmlToPdf_(htmlLetter, `Credential_${last}.pdf`, outFolder, cfg, blocks);
    }

    const icsBlob = buildVolunteerIcs_(objT, tz);

    const key = String(objD.Doc_Mode) === 'SINGLE'
      ? SETTINGS.ORDER_KEYS.VOL_M2_SINGLE
      : SETTINGS.ORDER_KEYS.VOL_M2_MULTI;

    const built = renderFromOrder_(key, objD, blocks, order, cfg);
    sendEmail_({
      to: objD.Volunteer_Email,
      subject: built.subject || (cfg.Volunteer_Subject || SETTINGS.SUBJECTS.VOL_M2),
      htmlBody: built.html,
      replyTo: cfg.Reply_To || cfg.From_Email,
      attachments: [pdfBlob, icsBlob],
      cfg,
      inlineImages: built.inlineImages
    });

    // Stage the state update — do NOT write to the sheet inside this loop
    const updatedRow = [...rowT];
    updatedRow[header.indexOf('Credential_Sent_On')]    = new Date();
    updatedRow[header.indexOf('Credential_Sent_By')]    = Session.getActiveUser().getEmail();
    updatedRow[header.indexOf('Needs_Credential_Send')] = false;

    dirtyRows.push({ sheetRow: i + 2, data: updatedRow });
    sent++;
  }

  // ── Single bulk write pass after all emails are sent ─────────────────────────
  // Writing row-by-row inside the loop was the primary cause of timeouts —
  // each setValues() is a separate Sheets API round-trip. Batching here cuts
  // that overhead from N calls down to N calls sequenced after the email work
  // is already done, so a timeout mid-flush doesn't block any sends.
  for (const w of dirtyRows) {
    master.getRange(w.sheetRow, 1, 1, w.data.length).setValues([w.data]);
  }

  log_(cfg, 'Send Volunteer M2', { sent });
  toast_(`M2 sent: ${sent}`, 5);
}

function cmdSendCountyM3() {
  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const assignmentsSs = getAssignmentsSs_(cfg);
  const bre  = getTab_(assignmentsSs, SETTINGS.TABS.BRE_MERGE);
  const roll = getTab_(assignmentsSs, SETTINGS.TABS.COUNTY_ROLLUP);

  const breData   = bre.getRange(1,1, bre.getLastRow()||1, bre.getLastColumn()||1).getDisplayValues();
  const breHeader = breData[0] || [];
  const breRows   = breData.slice(1);

  const idx = {};
  breHeader.forEach((k,i)=> idx[(k||'').toString().trim().toLowerCase()] = i);
  const iCounty = idx['county'];
  const iTo     = idx['to'];
  const iCc     = idx['cc'];

  if (iCounty == null || iTo == null) throw new Error('BRE Merge Sheet must include columns: County, To (and optional CC).');

  let sent = 0;

  for (const r of breRows) {
    const county = (r[iCounty] || '').trim();
    const to     = (r[iTo] || '').trim();
    const cc     = (iCc != null ? (r[iCc] || '').trim() : '');

    if (!county || !to) continue;

    const {header: rollHdr, row: rollRow} = findCountyRollupByCounty_(roll, county, cfg);
    if (!rollRow) { log_(cfg, 'County packet missing for ' + county, {}); continue; }

    const col = (name)=> rollHdr.indexOf(name);

    let { pdfBlob, csvBlob } = findLatestCountyArtifactsFromFolder_(county, cfg);

    if (!pdfBlob || !csvBlob) {
      const master = getTab_(assignmentsSs, SETTINGS.TABS.MASTER);
      const valsT  = getData_(master);
      const valsD  = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
      const header = valsT[0] || [];
      const dataT  = valsT.slice(1).map(x=>objFromRow_(header, x));
      const dataD  = valsD.slice(1).map(x=>objFromRow_(header, x));

      const listTyped = dataT.filter(x => normalizeCountyCore_(x.Assignment_County||'') === normalizeCountyCore_(county));
      const listDisp  = dataD.filter(x => normalizeCountyCore_(x.Assignment_County||'') === normalizeCountyCore_(county));

      const bigHtml   = renderCountyPacketHtml_(county, listDisp, blocks, cfg, order);
      const outFolder = DriveApp.getFolderById(cfg.Counties_Output_Folder_ID || SETTINGS.OUTPUT.COUNTIES_FOLDER_ID);
      const pdfName   = `County_${normalizeCountyCore_(county)}_${cfg.Run_Year||''}.pdf`;
      const newPdf    = htmlToPdf_(bigHtml, pdfName, outFolder, cfg, blocks);
      if (newPdf && newPdf.getBytes && newPdf.getBytes().length) pdfBlob = newPdf;

      const newCsv    = buildCountyCsv_(county, listTyped, cfg);
      if (newCsv && newCsv.getBytes && newCsv.getBytes().length) csvBlob = newCsv;
    }

    const attachments = [];
    if (pdfBlob) attachments.push(pdfBlob);
    if (csvBlob) attachments.push(csvBlob);

    const built = renderFromOrder_(SETTINGS.ORDER_KEYS.COUNTY_EMAIL, null, blocks, order, cfg, {
      breHeader: breHeader,
      breRow:    r
    });

    const inlineImagesFinal = ensureInlineLogoForCid_(built.html, built.inlineImages, cfg, blocks);

    sendEmail_({
      to,
      cc,
      subject: built.subject || (cfg.County_Subject || SETTINGS.SUBJECTS.COUNTY_M3),
      htmlBody: built.html,
      replyTo: cfg.Reply_To || cfg.From_Email,
      attachments,
      cfg,
      inlineImages: inlineImagesFinal
    });

    const sentCol = col('County_Email_Sent_On');
    if (sentCol >= 0) {
      const rollVals = roll.getRange(2,1, roll.getLastRow()-1, roll.getLastColumn()).getDisplayValues();
      for (let i=0; i<rollVals.length; i++) {
        const rowCounty = (rollVals[i][col('County')]||'').trim();
        const rowYear   = (col('Run_Year')>=0 ? (rollVals[i][col('Run_Year')]||'').trim() : '');
        if (rowCounty.toLowerCase() === county.toLowerCase()
            && (!cfg.Run_Year || !rowYear || rowYear === String(cfg.Run_Year))) {
          roll.getRange(i+2, sentCol+1).setValue(new Date());
          break;
        }
      }
    }
    sent++;
  }
  toast_(`County emails sent: ${sent}`, 5);
}

function cmdSendGenMM() {
  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const assignmentsSs = getAssignmentsSs_(cfg);
  const bre  = getTab_(assignmentsSs, SETTINGS.TABS.BRE_MERGE);
  const roll = getTab_(assignmentsSs, SETTINGS.TABS.GENERAL_MM_ROLLUP);

  const breData   = bre.getRange(1,1, bre.getLastRow()||1, bre.getLastColumn()||1).getDisplayValues();
  const breHeader = breData[0] || [];
  const breRows   = breData.slice(1);

  const idx = {};
  breHeader.forEach((k,i)=> idx[(k||'').toString().trim().toLowerCase()] = i);
  const iCounty = idx['county'];
  const iTo     = idx['to'];
  const iCc     = idx['cc'];

  if (iCounty == null || iTo == null) throw new Error('BRE Merge Sheet must include columns: County, To (and optional CC).');

  let sent = 0;

  for (const r of breRows) {
    const county = (r[iCounty] || '').trim();
    const to     = (r[iTo] || '').trim();
    const cc     = (iCc != null ? (r[iCc] || '').trim() : '');

    if (!county || !to) continue;

    const {header: rollHdr, row: rollRow} = findGeneralMMRollupByCounty_(roll, county, cfg);
    if (!rollRow) { log_(cfg, 'County packet missing for ' + county, {}); continue; }

    const built = renderFromOrder_(SETTINGS.ORDER_KEYS.GENERAL_MM, null, blocks, order, cfg, {
      breHeader: breHeader,
      breRow:    r
    });

    const inlineImagesFinal = ensureInlineLogoForCid_(built.html, built.inlineImages, cfg, blocks);

    sendEmail_({
      to,
      cc,
      subject: built.subject || (cfg.General_Subject || SETTINGS.SUBJECTS.GEN_MM),
      htmlBody: built.html,
      replyTo: cfg.Reply_To || cfg.From_Email,
      cfg,
      inlineImages: inlineImagesFinal
    });

    const col = (name)=> rollHdr.indexOf(name);
    const sentCol = col('General_MM_Sent_On');
    if (sentCol >= 0) {
      const rollVals = roll.getRange(2,1, roll.getLastRow()-1, roll.getLastColumn()).getDisplayValues();
      for (let i=0; i<rollVals.length; i++) {
        const rowCounty = (rollVals[i][col('County')]||'').trim();
        const rowYear   = (col('Run_Year')>=0 ? (rollVals[i][col('Run_Year')]||'').trim() : '');
        if (rowCounty.toLowerCase() === county.toLowerCase()
            && (!cfg.Run_Year || !rowYear || rowYear === String(cfg.Run_Year))) {
          roll.getRange(i+2, sentCol+1).setValue(new Date());
          break;
        }
      }
    }
    sent++;
  }
  toast_(`MM sent: ${sent}`, 5);
}

/** ==================
 *  UTILITY COMMANDS
 *  ================== */

function utilRecomputeHashes() {
  const { cfg } = readConfig_();
  const ss = getAssignmentsSs_(cfg);
  const sh = getTab_(ss, SETTINGS.TABS.MASTER);
  const vals = getData_(sh);
  const header = vals[0] || [];
  const rows = vals.slice(1);
  const idxHash = header.indexOf('Cred_Hash');
  const idxFlag = header.indexOf('Needs_Credential_Send');
  if (idxHash < 0 || idxFlag < 0) return;

  const tz = Session.getScriptTimeZone();
  rows.forEach(r => {
    const assigns = [];
    for (let i=1; i<=SETTINGS.MAX_ASSIGNMENTS; i++) {
      const d = r[header.indexOf(`A${i}_Date`)];
      const s = r[header.indexOf(`A${i}_Start`)];
      const e = r[header.indexOf(`A${i}_End`)];
      const ln = r[header.indexOf(`A${i}_LocationName`)];
      const addr = r[header.indexOf(`A${i}_Address`)];
      if (!d && !s && !ln && !addr) break;
      assigns.push({ date:d, start:s, end:e, locationName:ln, address:addr });
    }
    const canon = canonicalizeAssignmentsForHash_(assigns, tz);
    r[header.indexOf('Assignments_JSON')] = canon;
    r[idxHash] = sha256_(canon);
    r[idxFlag] = true;
  });

  if (rows.length) sh.getRange(2,1,rows.length, header.length).setValues(rows);
  toast_('Hashes recomputed from row fields; flags set TRUE.', 5);
}

function utilClearFlags() {
  const { cfg } = readConfig_();
  const ss = getAssignmentsSs_(cfg);
  const sh = getTab_(ss, SETTINGS.TABS.MASTER);
  const vals = getData_(sh);
  const header = vals[0] || [];
  const rows = vals.slice(1);

  const iM1 = header.indexOf('Needs_Confirmation_Send');
  const iM2 = header.indexOf('Needs_Credential_Send');
  if (rows.length && (iM1 >= 0 || iM2 >= 0)) {
    rows.forEach(r => {
      if (iM1 >= 0) r[iM1] = false;
      if (iM2 >= 0) r[iM2] = false;
    });
    sh.getRange(2,1,rows.length, header.length).setValues(rows);
  }
  toast_('All send flags (M1 & M2) cleared.', 5);
}

function utilExportLogs() {
  const { cfg } = readConfig_();
  const ss = getAssignmentsSs_(cfg);
  const sh = getOrCreate_(ss, SETTINGS.TABS.LOGS);
  const csv = getData_(sh).map(r=>r.map(csvEscape_).join(',')).join('\r\n');
  const blob = Utilities.newBlob(
    csv,
    'text/csv',
    `Project4_Logs_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.csv`
  );
  DriveApp.createFile(blob);
  toast_('Logs exported from active election workbook to Drive root.', 5);
}

function utilRepairHeaders() {
  const { cfg } = readConfig_();
  const ss = getAssignmentsSs_(cfg);

  getOrCreateWithHeader_(ss, SETTINGS.TABS.MASTER, SETTINGS.MASTER_HEADERS);
  getOrCreateWithHeader_(ss, SETTINGS.TABS.COUNTY_ROLLUP, SETTINGS.COUNTY_HEADERS);
  getOrCreateWithHeader_(ss, SETTINGS.TABS.GENERAL_MM_ROLLUP, SETTINGS.GENERAL_MM_HEADERS);
  getOrCreateWithHeader_(ss, SETTINGS.COUNTY_VOL_SHEET.MAP_TAB_NAME, SETTINGS.COUNTY_VOL_SHEET.MAP_HEADERS);

  toast_('Headers repaired for active election workbook.', 5);
}

function utilSendCountyM3_TestOne(countyName) {
  const { cfg } = readConfig_();
  const blocks  = readContentBlocksAdvanced_();
  const order   = readEmailOrder_();

  const assignmentsSs = getAssignmentsSs_(cfg);
  const bre  = getBRESheet_(cfg);
  const breVals   = bre.getRange(1,1, bre.getLastRow()||1, bre.getLastColumn()||1).getDisplayValues();
  const breHeader = breVals[0] || [];
  const rows      = breVals.slice(1);
  const iCounty = breHeader.findIndex(h => String(h).trim().toLowerCase()==='county');
  const iTo     = breHeader.findIndex(h => String(h).trim().toLowerCase()==='to');
  const iCc     = breHeader.findIndex(h => String(h).trim().toLowerCase()==='cc');

  if (iCounty < 0 || iTo < 0) throw new Error('BRE needs columns: County, To (and optional CC).');

  const wantCore = normalizeCountyCore_(countyName);
  const row = rows.find(r => normalizeCountyCore_(r[iCounty]||'') === wantCore);
  if (!row) throw new Error('County not found in BRE: ' + countyName);

  let { pdfBlob, csvBlob } = findLatestCountyArtifactsFromFolder_(row[iCounty], cfg);

  if (!pdfBlob || !csvBlob) {
    const master  = getTab_(assignmentsSs, SETTINGS.TABS.MASTER);
    const valsT   = getData_(master);
    const valsD   = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
    const header  = valsT[0] || [];
    const asObjT  = valsT.slice(1).map(x=>objFromRow_(header,x));
    const asObjD  = valsD.slice(1).map(x=>objFromRow_(header,x));
    const core    = normalizeCountyCore_(row[iCounty]);

    const listTyped = asObjT.filter(x => normalizeCountyCore_(x.Assignment_County||'') === core);
    const listDisp  = asObjD.filter(x => normalizeCountyCore_(x.Assignment_County||'') === core);

    const bigHtml   = renderCountyPacketHtml_(row[iCounty], listDisp, blocks, cfg, order);
    const outFolder = DriveApp.getFolderById(cfg.Counties_Output_Folder_ID || SETTINGS.OUTPUT.COUNTIES_FOLDER_ID);
    const pdfName   = `County_${core}_${cfg.Run_Year||''}.pdf`;
    const newPdf    = htmlToPdf_(bigHtml, pdfName, outFolder, cfg, blocks);
    if (newPdf && newPdf.getBytes && newPdf.getBytes().length) pdfBlob = newPdf;

    const newCsv = buildCountyCsv_(row[iCounty], listTyped, cfg);
    if (newCsv && newCsv.getBytes && newCsv.getBytes().length) csvBlob = newCsv;
  }

  const attachments = [];
  if (pdfBlob) attachments.push(pdfBlob);
  if (csvBlob) attachments.push(csvBlob);

  const built = renderFromOrder_(SETTINGS.ORDER_KEYS.COUNTY_EMAIL, null, blocks, order, cfg, {
    breHeader, breRow: row
  });

  const inlineImagesFinal = ensureInlineLogoForCid_(built.html, built.inlineImages, cfg, blocks);

  sendEmail_({
    to: row[iTo],
    cc: iCc >= 0 ? row[iCc] : '',
    subject: built.subject || (cfg.County_Subject || SETTINGS.SUBJECTS.COUNTY_M3),
    htmlBody: built.html,
    attachments,
    replyTo: cfg.Reply_To || cfg.From_Email,
    cfg,
    inlineImages: inlineImagesFinal
  });

  log_(cfg, 'utilSendCountyM3_TestOne', {
    county: row[iCounty],
    attachmentsCount: attachments.length,
    hasLogoInline: !!(inlineImagesFinal && inlineImagesFinal['logo_header'])
  });

  toast_('County M3 test queued for: ' + row[iCounty], 5);
}

function run_Test_M3() { utilSendCountyM3_TestOne('DeKalb'); }

function utilSendGenMM_TestOne(countyName) {
  const { cfg } = readConfig_();
  const blocks  = readContentBlocksAdvanced_();
  const order   = readEmailOrder_();

  const bre  = getBRESheet_(cfg);
  const breVals   = bre.getRange(1,1, bre.getLastRow()||1, bre.getLastColumn()||1).getDisplayValues();
  const breHeader = breVals[0] || [];
  const rows      = breVals.slice(1);
  const iCounty = breHeader.findIndex(h => String(h).trim().toLowerCase()==='county');
  const iTo     = breHeader.findIndex(h => String(h).trim().toLowerCase()==='to');
  const iCc     = breHeader.findIndex(h => String(h).trim().toLowerCase()==='cc');

  if (iCounty < 0 || iTo < 0) throw new Error('BRE needs columns: County, To (and optional CC).');

  const wantCore = normalizeCountyCore_(countyName);
  const row = rows.find(r => normalizeCountyCore_(r[iCounty]||'') === wantCore);
  if (!row) throw new Error('County not found in BRE: ' + countyName);

  const built = renderFromOrder_(SETTINGS.ORDER_KEYS.GENERAL_MM, null, blocks, order, cfg, {
    breHeader, breRow: row
  });

  const inlineImagesFinal = ensureInlineLogoForCid_(built.html, built.inlineImages, cfg, blocks);

  sendEmail_({
    to: row[iTo],
    cc: iCc >= 0 ? row[iCc] : '',
    subject: built.subject || (cfg.General_Subject || SETTINGS.SUBJECTS.GEN_MM),
    htmlBody: built.html,
    replyTo: cfg.Reply_To || cfg.From_Email,
    cfg,
    inlineImages: inlineImagesFinal
  });

  log_(cfg, 'utilSendGenMM_TestOne', {
    county: row[iCounty],
    hasLogoInline: !!(inlineImagesFinal && inlineImagesFinal['logo_header'])
  });

  toast_('General MM test queued for: ' + row[iCounty], 5);
}

function run_Test_MM() { utilSendGenMM_TestOne('DeKalb'); }

/** ======== Menu ======== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Project 4')
      .addItem('Build/Update Master', 'cmdBuildUpdateMaster')
      .addItem('Preview Diffs', 'cmdPreviewDiffs')
      .addSeparator()
      .addItem('Generate Volunteer PDFs', 'cmdGenerateVolunteerPDFs')
      .addItem('Generate County Packets', 'cmdGenerateCountyPackets')
      .addSeparator()
      .addItem('Send Volunteer Emails — M1 (LBJ)', 'cmdSendVolunteerM1')
      .addItem('Send Volunteer Emails — M2 (LBJ)', 'cmdSendVolunteerM2')
      .addItem('Send County Emails — M3 (BRE)', 'cmdSendCountyM3')
      .addItem('Send Mail Merge — MM (Gen)', 'cmdSendGenMM')
      .addSeparator()
      .addSubMenu(
        SpreadsheetApp.getUi().createMenu('Utilities')
          .addItem('Recompute Hashes', 'utilRecomputeHashes')
          .addItem('Clear Send Flags', 'utilClearFlags')
          .addItem('Export Logs', 'utilExportLogs')
          /**************************************************************
           * ADDED — utilities for county file map + assigned volunteer tabs
           **************************************************************/
          .addItem('Backfill County File Map', 'utilBackfillCountyFileMap')
          .addItem('Write Assigned Volunteers Tabs', 'utilWriteAssignedVolunteersTabs_AllCounties')
          .addItem('Reset County Packet Progress', 'utilResetCountyPacketResume')
      )
    .addToUi();
}

function formatAssignmentDate_(y, m, d) {
  const days   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const dow    = new Date(y, m - 1, d).getDay();
  return `${days[dow]} ${months[m - 1]} ${String(d).padStart(2,'0')} ${y}`;
}

function utilResetCountyPacketResume() {
  PropertiesService.getScriptProperties().deleteProperty('P4_COUNTY_PKT_RESUME');
  toast_('County packet resume pointer cleared.', 3);
}

function utilDebugDates() {
  const { cfg } = readConfig_();

  Logger.log('cfg.email_sheet_id=' + cfg.email_sheet_id);
  Logger.log('cfg.assignments_sheet_id=' + cfg.assignments_sheet_id);

  const keys = [
    'Data_Mapping_Sheet_ID',
    'DataMappingSheetId',
    'Data_Mapping_Doc_ID',
    'Data Mapping Document ID',
    'Data Mapping Spreadsheet ID',
    'data_mapping_spreadsheet_id',
    'Assignments_Tab_Name'
  ];
  keys.forEach(k => Logger.log(k + '=' + (cfg[k] || '')));

  Logger.log('Assignment Dates bullets=' + getTwoColDatesBullets_(cfg, 'Assignment Dates'));
  Logger.log('Credential Dates bullets=' + getTwoColDatesBullets_(cfg, 'Credential Dates'));
}
