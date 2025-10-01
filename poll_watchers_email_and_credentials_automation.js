/**************************************************************
 * Project 4 — Automated Credential Generation & Distribution
 **************************************************************/

const SETTINGS = {
  TABS: {
    CONFIG: 'Config',
    CONTENT_BLOCKS: 'Content Blocks',
    ORDER: 'Project 4 Email Order',        // keep for safety
    P4_EMAIL_ORDER: 'Project 4 Email Order',
    LBJ_IMPORT: 'LBJ_Import (latest)',
    MASTER: 'Master_Assignments',
    COUNTY_ROLLUP: 'County_Rollup',
    BRE_MERGE: 'BRE Merge Sheet',
    LOGS: 'Logs',
    README: 'README',
  },

  IMAGE_DEFAULTS: {
    LOGO_MAX_WIDTH: 360,    // px
    SIGNATURE_HEIGHT: 96    // px
  },

  // BRE workbook (set to ACTIVE/blank to use this workbook)
  DEFAULT_BRE: { WORKBOOK_ID: '', TAB_NAME: 'BRE Merge Sheet' },

  OUTPUT: {
    VOLUNTEERS_FOLDER_ID: 'PUT_VOLUNTEER_PDFS_FOLDER_ID',
    COUNTIES_FOLDER_ID: 'PUT_COUNTY_OUTPUT_FOLDER_ID',
  },

  MAX_ASSIGNMENTS: 15,

  MASTER_HEADERS: [
    'VAN_ID','Volunteer_Name','Volunteer_Email','Volunteer_Phone','Volunteer Address','County','Run_Year',
    'Yes_Link','No_Link',
    'Assignment_Count','Doc_Mode','Assignments_JSON','Cred_Hash',
    ...Array.from({length: 15}).flatMap((_,i)=>[
      `A${i+1}_Date`,`A${i+1}_Start`,`A${i+1}_End`,`A${i+1}_LocationName`,`A${i+1}_Address`
    ]),
    'Volunteer_PDF_File_ID','ICS_File_Name',
    'Credential_Sent_On','Credential_Sent_By','Source_LBJ_Sheet_Tag','Needs_Credential_Send','Errors'
  ],

  COUNTY_HEADERS: [
    'County','Run_Year','Volunteer_Count','VAN_ID_List',
    'County_PDF_File_ID','County_CSV_File_ID','County_Email_Sent_On','Errors'
  ],

  // LBJ header aliases
  LBJ_FIELDS: {
    VAN_ID:          ['VAN_ID','VanID','VAN ID'],
    FIRST:           ['First Name','FirstName','VolFirstName','Volunteer First Name'],
    LAST:            ['Last Name','LastName','VolLastName','Volunteer Last Name'],
    EMAIL:           ['Email','Email Address'],
    PHONE:           ['Phone Number','Phone','Cell Phone','Cell Phone Number','Mobile','VolPhone'],
    COUNTY:          ['County'],
    DATE:            ['Date','Assignment Date'],
    START:           ['Start Time','StartTime'],
    END:             ['End Time','EndTime'],
    LOCATION_NAME:   ['Polling Location','Polling Location Name','Location Name','LocationName'],

    // LOCATION address = polling-site-only (DO NOT include plain "Address")
    LOCATION_ADDR:   ['Polling Location Address','Location Address','LocationAddress'],

    SHEET_TAG:       ['LBJ_Sheet_Tag','Sheet_Tag'],

    // volunteer MAILING address = LBJ "Address" (plus common variants)
    VOL_ADDRESS: [
      'Address','Volunteer Address','Mailing Address','Mailing_Address',
      'Address Line 1','Address1','Street Address','Home Address','Residential Address'
    ]
  },

  // Prefill form placeholders + entry IDs
  FORM_LINK: {
    BASE_URL_CONFIG_KEY: 'Form_Base_URL',        // direct URL (recommended)
    BASE_URL_CELL_KEY:   'Form_Base_URL_Cell',   // e.g., "Form Link!A1" (optional)
    YES_PARAM: 'Yes',
    NO_PARAM: 'No',
    PLACEHOLDERS: { first:'VolFirstName', last:'VolLastName', phone:'VolPhone', avail:'AvailGood' },
    ENTRY_IDS:   { first:'entry.576587497', last:'entry.1363204004', phone:'entry.436579453', avail:'entry.1913347883' }
  },

  SUBJECTS: {
    VOL_M1: 'Assignment Confirmation — Action Requested',
    VOL_M2: 'Your Official Poll Watcher Credential',
    COUNTY_M3: 'County Packet: Poll Watcher Credentials & Roster'
  },

  ORDER_KEYS: {
    VOL_M1_SINGLE: 'Single Assignment Confirmation Email',
    VOL_M1_MULTI:  'Multiple Assignment Confirmation Email',
    VOL_M2_SINGLE: 'Single Assignment Credential Email',
    VOL_M2_MULTI:  'Multiple Assignment Credential Email',
    COUNTY_EMAIL:  'Email to Counties',
    LETTER_SINGLE: 'Single Assignment Credential Letter',
    LETTER_MULTI:  'Multiple Assignment Credential Letter'
  }
};

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
      .addSeparator()
      .addSubMenu(
        SpreadsheetApp.getUi().createMenu('Utilities')
          .addItem('Recompute Hashes', 'utilRecomputeHashes')
          .addItem('Clear Send Flags', 'utilClearFlags')
          .addItem('Export Logs', 'utilExportLogs')
      )
    .addToUi();
}

/** =================
 *  1) CORE COMMANDS
 *  ================= */
function cmdBuildUpdateMaster() {
  const {cfg} = readConfig_();
  ensureSheets_();

  const lbjRows    = readLBJ_AndGenerateYesNo_(cfg);
  const masterRows = buildMasterRows_(lbjRows, cfg);
  upsertMaster_(masterRows, cfg);

  refreshCountyRollup_AllCounties_(cfg);
  log_('Build/Update Master', {countLBJ: lbjRows.length, countMaster: masterRows.length});
  toast_('Master updated.', 5);
}

function cmdPreviewDiffs() {
  const ss = SpreadsheetApp.getActive();
  const master = getTab_(ss, SETTINGS.TABS.MASTER);
  const vals = getData_(master);
  const header = vals[0] || [];
  const rows = vals.slice(1).map(r => objFromRow_(header, r));
  const flagged = rows.filter(r => String(r.Needs_Credential_Send).toUpperCase() === 'TRUE');
  const byCountyFlagged = groupBy_(flagged, r => r.County || 'Unknown');

  const {cfg} = readConfig_();
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

  const ss = SpreadsheetApp.getActive();
  const master = getTab_(ss, SETTINGS.TABS.MASTER);

  const valsT = getData_(master);
  const valsD = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
  const header = valsT[0] || [];
  const dataT  = valsT.slice(1);
  const dataD  = valsD.slice(1);

  const outFolder = DriveApp.getFolderById(cfg.Volunteers_Output_Folder_ID || SETTINGS.OUTPUT.VOLUNTEERS_FOLDER_ID);

  let gen = 0;
  for (let i=0; i<dataT.length; i++) {
    const rowT = dataT[i], rowD = dataD[i];
    const objT = objFromRow_(header, rowT);
    const objD = objFromRow_(header, rowD);

    if (String(objT.Needs_Credential_Send).toUpperCase() !== 'TRUE') continue;

    // Build the letter HTML once…
    const html = renderVolunteerLetterHtml_FromOrder_(objD, blocks, order, cfg);

    // …then convert to PDF (declare pdfBlob here)
    const last = fileSafe_(extractLastName_(objD.Volunteer_Name));
    const pdfBlob = htmlToPdf_(html, `Credential_${last}_${objD.VAN_ID}.pdf`, outFolder, cfg, blocks);

    // Save and store the file id back to the sheet
    const fileId = ensureSavedInFolder_(pdfBlob, outFolder);
    const pdfCol = header.indexOf('Volunteer_PDF_File_ID');
    if (pdfCol >= 0) {
      rowT[pdfCol] = fileId;
      master.getRange(i+2, 1, 1, rowT.length).setValues([rowT]);
      gen++;
    }
  }
  log_('Generate Volunteer PDFs', {generated: gen});
  toast_(`Volunteer PDFs generated: ${gen}`, 5);
}


function cmdGenerateCountyPackets() {
  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const ss = SpreadsheetApp.getActive();
  const master = getTab_(ss, SETTINGS.TABS.MASTER);

  const valsT = getData_(master);
  const valsD = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
  const header = valsT[0] || [];
  const dataT  = valsT.slice(1).map(r => objFromRow_(header, r));
  const dataD  = valsD.slice(1).map(r => objFromRow_(header, r));

  const byCountyDisp = groupBy_(dataD, r => r.County || 'Unknown');
  const outFolder = DriveApp.getFolderById(cfg.Counties_Output_Folder_ID || SETTINGS.OUTPUT.COUNTIES_FOLDER_ID);
  const roll = getOrCreateWithHeader_(ss, SETTINGS.TABS.COUNTY_ROLLUP, SETTINGS.COUNTY_HEADERS);
  const rollHeader = SETTINGS.COUNTY_HEADERS;

  if (roll.getLastRow()>1) roll.getRange(2,1,roll.getLastRow()-1, roll.getLastColumn()).clearContent();

  let created = 0;
  Object.keys(byCountyDisp).forEach(cty => {
    const listDisp = byCountyDisp[cty].sort((a,b)=> (a.Volunteer_Name||'').localeCompare(b.Volunteer_Name||''));
    const bigHtml = renderCountyPacketHtml_(cty, listDisp, blocks, cfg, order);
    const pdfBlob = htmlToPdf_(bigHtml, `County_${cty}_${cfg.Run_Year||''}.pdf`, outFolder, cfg, blocks);
    const pdfId = ensureSavedInFolder_(pdfBlob, outFolder);

    const listTyped = dataT.filter(r => (r.County||'Unknown') === cty);
    const csvBlob = buildCountyCsv_(cty, listTyped, cfg);
    const csvFile = outFolder.createFile(csvBlob);
    const csvId = csvFile.getId();

    roll.appendRow(rowFromObj_(rollHeader, {
      County: cty,
      Run_Year: cfg.Run_Year || '',
      Volunteer_Count: listDisp.length,
      VAN_ID_List: listDisp.map(v=>v.VAN_ID).join(','),
      County_PDF_File_ID: pdfId,
      County_CSV_File_ID: csvId,
      County_Email_Sent_On: '',
      Errors: ''
    }));
    created++;
  });

  log_('Generate County Packets', {counties: created});
  toast_(`County packets generated: ${created}`, 5);
}

// --- M1: Assignment confirmation to volunteers ---
function cmdSendVolunteerM1() {
  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const ss = SpreadsheetApp.getActive();
  const master = getTab_(ss, SETTINGS.TABS.MASTER);

  const valsT = getData_(master);
  const valsD = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
  const header = valsT[0] || [];
  const dataD  = valsD.slice(1);

  let sent = 0;
  for (let i=0; i<dataD.length; i++) {
    const objD = objFromRow_(header, dataD[i]);
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
    sent++;
  }
  toast_(`M1 sent: ${sent}`, 5);
}

// --- M2: Credential with PDF + ICS to volunteers ---
function cmdSendVolunteerM2() {
  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const ss = SpreadsheetApp.getActive();
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

  for (let i=0; i<dataT.length; i++) {
    const rowT = dataT[i], rowD = dataD[i];
    const objT = objFromRow_(header, rowT);
    const objD = objFromRow_(header, rowD);

    if (String(objT.Needs_Credential_Send).toUpperCase() !== 'TRUE') continue;
    if (!objD.Volunteer_Email) continue;

    // ensure PDF
    let pdfBlob = null;
    if (objT.Volunteer_PDF_File_ID) { try { pdfBlob = DriveApp.getFileById(objT.Volunteer_PDF_File_ID).getBlob(); } catch(e){} }
    if (!pdfBlob) {
      const htmlLetter = renderVolunteerLetterHtml_FromOrder_(objD, blocks, order, cfg);
      const last = fileSafe_(extractLastName_(objD.Volunteer_Name));
      pdfBlob = htmlToPdf_(htmlLetter, `Credential_${last}_${objD.VAN_ID}.pdf`, outFolder, cfg, blocks);
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

    // state update
    const now = new Date();
    rowT[header.indexOf('Credential_Sent_On')]   = now;
    rowT[header.indexOf('Credential_Sent_By')]   = Session.getActiveUser().getEmail();
    rowT[header.indexOf('Needs_Credential_Send')] = false;
    master.getRange(i+2, 1, 1, rowT.length).setValues([rowT]);

    sent++;
    if (sent >= maxBatch) break;
  }
  toast_(`M2 sent: ${sent}`, 5);
}

// Look up the most relevant row in County_Rollup for a county
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
      best = r; break; // exact match on year
    }
    if (!best) best = r;
  }
  return { header, row: best };
}

// ---------------- County sender (Mail Merge 3) ----------------
function cmdSendCountyM3() {
  const {cfg} = readConfig_();
  const blocks = readContentBlocksAdvanced_();
  const order  = readEmailOrder_();

  const ss   = SpreadsheetApp.getActive();
  const bre  = getTab_(ss, SETTINGS.TABS.BRE_MERGE);
  const roll = getTab_(ss, SETTINGS.TABS.COUNTY_ROLLUP);

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
    if (!rollRow) { log_('County packet missing for ' + county, {}); continue; }

    const col = (name)=> rollHdr.indexOf(name);
    const pdfId = (col('County_PDF_File_ID')>=0) ? (rollRow[col('County_PDF_File_ID')]||'') : '';
    const csvId = (col('County_CSV_File_ID')>=0) ? (rollRow[col('County_CSV_File_ID')]||'') : '';

    // --- ATTACHMENTS: pull straight from Counties_Output_Folder_ID ---
    let { pdfBlob, csvBlob } = findLatestCountyArtifactsFromFolder_(county, cfg);

    // Fallback: if the combined county packet isn't present, regenerate it now
    if (!pdfBlob || !csvBlob) {
      const master = getTab_(ss, SETTINGS.TABS.MASTER);
      const valsT  = getData_(master);
      const valsD  = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
      const header = valsT[0] || [];
      const dataT  = valsT.slice(1).map(x=>objFromRow_(header, x));
      const dataD  = valsD.slice(1).map(x=>objFromRow_(header, x));

      const listTyped = dataT.filter(x => normalizeCountyCore_(x.County||'') === normalizeCountyCore_(county));
      const listDisp  = dataD.filter(x => normalizeCountyCore_(x.County||'') === normalizeCountyCore_(county));

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

    // --- BODY from ORDER / Content Blocks (loads BRE row via extras) ---
    const built = renderFromOrder_(SETTINGS.ORDER_KEYS.COUNTY_EMAIL, null, blocks, order, cfg, {
      breHeader: breHeader,
      breRow:    r
    });

    // Make sure CID logo actually has a blob if referenced
    const inlineImagesFinal = ensureInlineLogoForCid_(built.html, built.inlineImages, cfg, blocks);

    // --- SEND ---
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

/** ========================
 *  2) BUILD MASTER HELPERS
 *  ======================== */
function readLBJ_AndGenerateYesNo_(cfg) {
  const ss = SpreadsheetApp.getActive();
  const sh = getTab_(ss, SETTINGS.TABS.LBJ_IMPORT);

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
  const fromCell = getCellByA1_(ss, cfg[SETTINGS.FORM_LINK.BASE_URL_CELL_KEY] || '');
  const baseUrl  = direct || fromCell || '';

  const fieldIdx = mapHeaderIdxFlexible_(header, SETTINGS.LBJ_FIELDS);
  const out = [];

  rows.forEach((r, i) => {
    const dr = drows[i] || r;

    // Keep EXACT sheet display for date/time fields
    if (fieldIdx.DATE  >= 0) r[fieldIdx.DATE]  = dr[fieldIdx.DATE];
    if (fieldIdx.START >= 0) r[fieldIdx.START] = dr[fieldIdx.START];
    if (fieldIdx.END   >= 0) r[fieldIdx.END]   = dr[fieldIdx.END];

    // Pull hyperlink URLs if W/X are hyperlink cells
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
      const phone = getValIdx_(r, fieldIdx.PHONE);
      if (yesCol >= 0 && !haveYes) r[yesCol] = buildPrefilledLink_(baseUrl, first, last, phone, SETTINGS.FORM_LINK.YES_PARAM);
      if (noCol  >= 0 && !haveNo)  r[noCol]  = buildPrefilledLink_(baseUrl, first, last, phone, SETTINGS.FORM_LINK.NO_PARAM);
    }

    if (yesCol < 0 && noCol < 0 && baseUrl) {
      const first = getValIdx_(r, fieldIdx.FIRST);
      const last  = getValIdx_(r, fieldIdx.LAST);
      const phone = getValIdx_(r, fieldIdx.PHONE);
      sh.getRange(i + 2, 23).setValue(buildPrefilledLink_(baseUrl, first, last, phone, SETTINGS.FORM_LINK.YES_PARAM)); // W
      sh.getRange(i + 2, 24).setValue(buildPrefilledLink_(baseUrl, first, last, phone, SETTINGS.FORM_LINK.NO_PARAM));  // X
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
    const county = valByAliases_(any, SETTINGS.LBJ_FIELDS.COUNTY);
    const volAddress = valByAliases_(any, SETTINGS.LBJ_FIELDS.VOL_ADDRESS);
    const yesLink= any['Yes_Link'] || '';
    const noLink = any['No_Link']  || '';
    const sheetTag = valByAliases_(any, SETTINGS.LBJ_FIELDS.SHEET_TAG) || inferSheetTag_();

    const volunteerName = [first,last].filter(Boolean).join(' ').trim();

    // Build raw assignments (keep display look for emails/PDF),
    // but compute hash from a *normalized* version only.
    const assigns = items.map(it => ({
      date: (valByAliases_(it, SETTINGS.LBJ_FIELDS.DATE) || '').trim(),
      start: (valByAliases_(it, SETTINGS.LBJ_FIELDS.START) || '').trim(),
      end:   (valByAliases_(it, SETTINGS.LBJ_FIELDS.END)   || '').trim(),
      locationName: (valByAliases_(it, SETTINGS.LBJ_FIELDS.LOCATION_NAME) || '').trim(),
      address:      (valByAliases_(it, SETTINGS.LBJ_FIELDS.LOCATION_ADDR) || '').trim(),
      county
    }));

    // Order-insensitive & dedup (by all fields)
    assigns.sort((a,b)=>
      (a.date||'').localeCompare(b.date||'') ||
      (a.start||'').localeCompare(b.start||'') ||
      (a.locationName||'').localeCompare(b.locationName||'') ||
      (a.address||'').localeCompare(b.address||'')
    );
    const dedup = dedupeAssignments_(assigns);

    const count = dedup.length;
    const docMode = (count <= 1) ? 'SINGLE' : 'TABLE';

    // Flatten into A1_*, A2_* … for templates
    const flat = {};
    dedup.slice(0, SETTINGS.MAX_ASSIGNMENTS).forEach((a, idx) => {
      const p = `A${idx+1}_`;
      flat[p+'Date'] = a.date;
      flat[p+'Start'] = a.start;
      flat[p+'End'] = a.end;
      flat[p+'LocationName'] = a.locationName;
      flat[p+'Address'] = a.address;
    });

    // ***** IMPORTANT: hash only the normalized assignment set *****
    const canon = canonicalizeAssignmentsForHash_(dedup, cfg.Calendar_Timezone);
    const hash  = sha256_(canon);

    result.push({
      VAN_ID: String(van).trim(),
      Volunteer_Name: volunteerName,
      Volunteer_Email: email,
      Volunteer_Phone: phone,
      Volunteer_Address: volAddress,
      'Volunteer Address': volAddress,
      County: county,
      Run_Year: String(cfg.Run_Year||''),
      Yes_Link: yesLink,
      No_Link: noLink,
      Assignment_Count: count,
      Doc_Mode: docMode,
      Assignments_JSON: canon,             // store canon used for the hash
      Cred_Hash: hash,
      ...flat,
      Volunteer_PDF_File_ID: '',
      ICS_File_Name: '',
      Credential_Sent_On: '',
      Credential_Sent_By: '',
      Source_LBJ_Sheet_Tag: sheetTag,
      Needs_Credential_Send: true,        // will be corrected in upsert
      Errors: ''
    });
  });

  return result;
}

function upsertMaster_(rows, cfg) {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreateWithHeader_(ss, SETTINGS.TABS.MASTER, SETTINGS.MASTER_HEADERS);
  const header = getHeader_(sh);
  const existing = getData_(sh).slice(1).map(r => objFromRow_(header, r));

  // normalize VAN keys so "12345" === 12345
  const vanKey_ = v => String(v || '').trim();

  const byVanExisting = new Map(existing.map(r => [vanKey_(r.VAN_ID), r]));

  const writeRows = [];

  rows.forEach(r => {
    const prior = byVanExisting.get(vanKey_(r.VAN_ID));
    if (prior) {
      // Preserve state fields from prior row
      r.Credential_Sent_On     = prior.Credential_Sent_On || '';
      r.Credential_Sent_By     = prior.Credential_Sent_By || '';
      r.Source_LBJ_Sheet_Tag   = r.Source_LBJ_Sheet_Tag || prior.Source_LBJ_Sheet_Tag || '';
      r.Volunteer_PDF_File_ID  = prior.Volunteer_PDF_File_ID || '';

      const priorAssignHash = assignmentsHashFromJson_(prior.Assignments_JSON);
      const newAssignHash   = assignmentsHashFromJson_(r.Assignments_JSON);
      const changed = priorAssignHash !== newAssignHash;

      const prevFlag = asBool_(prior.Needs_Credential_Send);
      r.Needs_Credential_Send = prevFlag || changed;
    } else {
      r.Needs_Credential_Send = true; // brand new VAN_ID
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


function refreshCountyRollup_AllCounties_(cfg) {
  const ss = SpreadsheetApp.getActive();
  const master = getTab_(ss, SETTINGS.TABS.MASTER);
  const vals = getData_(master);
  if (!vals.length) return;
  const header = vals[0];
  const rows = vals.slice(1).map(r => objFromRow_(header, r));
  const byCountyAll = groupBy_(rows, r => r.County || 'Unknown');

  const roll = getOrCreateWithHeader_(ss, SETTINGS.TABS.COUNTY_ROLLUP, SETTINGS.COUNTY_HEADERS);
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

// Robust prefill: replace placeholders OR append entry.* params
function buildPrefilledLink_(baseUrl, first, last, phone, yesOrNo) {
  if (!baseUrl) return '#';
  const ph = SETTINGS.FORM_LINK.PLACEHOLDERS;
  const ids = SETTINGS.FORM_LINK.ENTRY_IDS;

  const hasPlaceholders =
    baseUrl.includes(ph.first) || baseUrl.includes(ph.last) ||
    baseUrl.includes(ph.phone) || baseUrl.includes(ph.avail);

  if (hasPlaceholders) {
    return baseUrl
      .replace(new RegExp(ph.first, 'g'), encodeURIComponent(first || ''))
      .replace(new RegExp(ph.last,  'g'), encodeURIComponent(last  || ''))
      .replace(new RegExp(ph.phone, 'g'), encodeURIComponent(phone || ''))
      .replace(new RegExp(ph.avail, 'g'), encodeURIComponent(yesOrNo || ''));
  }

  const parts = [];
  if (first)   parts.push(`${ids.first}=${encodeURIComponent(first)}`);
  if (last)    parts.push(`${ids.last}=${encodeURIComponent(last)}`);
  if (phone)   parts.push(`${ids.phone}=${encodeURIComponent(phone)}`);
  if (yesOrNo) parts.push(`${ids.avail}=${encodeURIComponent(yesOrNo)}`);

  const sep = baseUrl.includes('?') ? (baseUrl.endsWith('?') || baseUrl.endsWith('&') ? '' : '&') : '?';
  return baseUrl + (parts.length ? (sep + parts.join('&')) : '');
}

function ensureSavedInFolder_(pdfBlob, folder) {
  const name = pdfBlob.getName();
  const files = folder.getFilesByName(name);
  while (files.hasNext()) { const f = files.next(); try { f.setTrashed(true); } catch(e){} }
  const f = folder.createFile(pdfBlob);
  return f.getId();
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

// ---- DROP-IN: merge token helpers ----
// Supports *|Token|*, |Token|, {{Token}}, {Token}, [Token]
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
  row = row || {}; cfg = cfg || {}; blocks = blocks || { text: {} };

  // ---- name bits ----
  const name = (row.Volunteer_Name || '').toString().trim();
  const parts = name ? name.split(/\s+/) : [];
  const first = parts[0] || '';
  const last  = parts.length > 1 ? parts[parts.length - 1] : '';

  // ---- on/during (single vs multi) ----
  const onDuring = (Number(row.Assignment_Count || 0) > 1) ? 'during' : 'on';

  // ---- pretty letter date ----
  const tz = cfg.Calendar_Timezone || Session.getScriptTimeZone() || 'America/New_York';
  const letterDateRaw = cfg.Letter_Date && String(cfg.Letter_Date).trim() ? cfg.Letter_Date : new Date();
  const letterDate = formatPrettyDate_(letterDateRaw, tz);

  // ---- volunteer mailing address (DO NOT read plain "Address") ----
  const addr1 = pick_(row,
    'Volunteer Address','Volunteer_Address','Mailing Address','Mailing_Address',
    'Address Line 1','Address1'
  );
  const addr2 = pick_(row, 'Address Line 2','Address2','Apt','Unit');
  const city  = pick_(row, 'City','Volunteer_City','Mailing_City');
  const state = pick_(row, 'State','St','Province');
  const zip   = pick_(row, 'Zip','ZIP','Postal','Postal_Code','Postal Code');

  let addrBlock = '';
  if (addr1 || addr2 || city || state || zip) {
    const line2 = [city, state].filter(Boolean).join(', ');
    addrBlock = [addr1, addr2, [line2, zip].filter(Boolean).join(' ')].filter(Boolean).join('\n');
  } else if (row.VAN_ID) {
    // optional safety net: fall back to LBJ mailing address by VAN
    const lbj = lookupVolunteerAddressByVan_(row.VAN_ID);
    if (lbj) {
      const line2 = [lbj.city, lbj.state].filter(Boolean).join(', ');
      addrBlock = [lbj.addr1, lbj.addr2, [line2, lbj.zip].filter(Boolean).join(' ')].filter(Boolean).join('\n');
    }
  }

  function pref(cfgKey, cbKey) {
    let v = (cfg[cfgKey] != null && cfg[cfgKey] !== '') ? cfg[cfgKey] : null;
    if (v == null && blocks.text && cbKey) v = blocks.text[cbKey];
    return v || '';
  }
  
  let countyTok = (row.County || '').toString().trim();
  if (!countyTok && extras && extras.breHeader && extras.breRow) {
    const iCounty = extras.breHeader.findIndex(h => String(h).trim().toLowerCase() === 'county');
    if (iCounty >= 0) countyTok = (extras.breRow[iCounty] || '').toString().trim();
  }

  // ---- base map (core tokens) ----
  const map = {
    // Names
    'Volunteer Name': name,
    'First Name': first, 'FirstName': first, 'Volunteer First Name': first,
    'Last Name':  last,  'LastName':  last,  'Volunteer Last Name':  last,

    // Address (mailing)
    'Address': addrBlock,
    'Mailing Address': addrBlock,
    'Volunteer Address': addrBlock,

    // Dates/titles
    'Letter Date':   letterDate,
    'OnDuring':      onDuring, 'On/During': onDuring, 'On or During': onDuring,
    'Voting Type':   pref('Voting_Type',   'Voting Type'),
    'Election Date': pref('Election_Date', 'Election Date'),
    'Election Title':pref('Election_Type', 'Election Title'),

    // Misc
    'County': countyTok
  };

  // ---- merge BRE row columns as tokens, without clobbering core tokens ----
  if (extras && extras.breHeader && extras.breRow) {
    const H = extras.breHeader, R = extras.breRow;

    // tokens we refuse to overwrite
    const RESERVED = new Set(Object.keys(map).map(k => k.toLowerCase().replace(/\s+/g,' ')));

    for (let i = 0; i < H.length; i++) {
      const rawKey = (H[i] || '').toString().trim();
      if (!rawKey) continue;
      const val = (R[i] != null) ? R[i] : '';

      const norm = rawKey.replace(/\s+/g, ' ').trim();
      const normLC = norm.toLowerCase();

      // Always expose a BRE-namespaced token
      map['BRE ' + norm] = val;

      // Only add the plain key if it wouldn't override a core token
      if (!RESERVED.has(normLC) && !Object.prototype.hasOwnProperty.call(map, rawKey)) {
        map[rawKey] = val;
        map[norm]   = val;
      }
    }
  }

  return map;
}


function lookupVolunteerAddressByVan_(van) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = getTab_(ss, SETTINGS.TABS.LBJ_IMPORT);
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


// ---- helpers ----
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

/** ======================================
 *  3) RENDERING VIA ORDER (HTML → PDF/ICS)
 *  ====================================== */

// ----- Order reader -----
function readEmailOrder_() {
  const TAB = (SETTINGS && SETTINGS.TABS && SETTINGS.TABS.ORDER) || 'Project 4 Email Order';
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(TAB);
  if (!sh) throw new Error('Missing tab: ' + TAB);

  const rows = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), Math.max(1, sh.getLastColumn()))
                 .getDisplayValues();
  if (rows.length < 2) return {};

  const header = rows[0].map(h => (h || '').toString().trim());
  const body   = rows.slice(1);

  // find a likely "label" column, else default to 0
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
      if (!val) continue; // skip empty cells

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

// ----- Generic renderer from ORDER -----
function renderFromOrder_(templateColName, row, blocks, order, cfg, extras) {
  const tpl = getOrderTemplate_(order, templateColName) || { subjectKey:'', blocks:[] };

  // Auto-load the BRE row by county if not explicitly provided
  let _extras = extras || null;
  if ((!_extras || !_extras.breHeader) && row && row.County) {
    const found = getBRE_ForCounty_(cfg, row.County);
    if (found) _extras = Object.assign({}, extras || {}, found);
  }

  // Subject: allow either a block key or literal text
  let subject = '';
  if (tpl.subjectKey) {
    const fromCB = blocks && blocks.text && blocks.text[tpl.subjectKey];
    subject = fromCB ? stripHtml_(fromCB) : stripHtml_(tpl.subjectKey);
  }

  // Images: data URLs for letters/PDFs, CID for emails
  const useDataUrls = !!(_extras && _extras.imageMode === 'dataurl');
  const bag = { __useCid: !useDataUrls };

  // Render blocks with spacing after EACH component
  const parts = [];
  if (tpl.blocks && tpl.blocks.length) {
    const seq = tpl.blocks.slice();
    // Merge "Signature 1" + "Signature 2" into a single row if adjacent
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

  // Safe fallback (also spaced)
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


// ----- Block to HTML (dynamic tokens + logo/signature handling) -----
function blockToHtml_(keyRaw, row, blocks, cfg, extras, bag) {
  if (!keyRaw) return '';
  const key = String(keyRaw).trim();

  // ---- logo (emails: inline CID; letters: injected in Doc header so skip if forced) ----
  if (/^logo\s*(header|image)?$/i.test(key)) {
    if (extras && extras.forceHeaderLogo) return ''; // for PDFs, we put it in the Doc header
    const blob = findLogoBlob_(cfg, blocks);        // <— tolerant lookup
    if (!blob) return '';
    const src = addInlineImage_(bag, 'logo_header', blob); // emails: cid:logo_header
    const maxW = Number(cfg.Logo_Max_Width_Px || 320);
    return `<div style="text-align:center;margin:8px 0 14px;">
              <img src="${src}" alt="Logo" style="max-width:${maxW}px;width:100%;height:auto;display:inline-block;">
            </div>`;
  }


  // ---- dynamic blocks ----
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
    const yesTxt = blocks.text['Google Form Link Yes'] || 'Yes, I can make it';
    const noTxt  = blocks.text['Google Form Link No']  || 'No, I cannot';
    let yes = row.Yes_Link || '';
    let no  = row.No_Link  || '';
    const baseUrl = (cfg[SETTINGS.FORM_LINK.BASE_URL_CONFIG_KEY] ||
                     getCellByA1_(SpreadsheetApp.getActive(), cfg[SETTINGS.FORM_LINK.BASE_URL_CELL_KEY] || '') || '').trim();
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
    const clickTxt = blocks.text['Google Form Link Click Here'] || 'click here';
    const href = (row && row.Yes_Link) ? row.Yes_Link : '';
    if (href) return `<a href="${href}" target="_blank" rel="noopener noreferrer" style="color:#1155cc;text-decoration:underline;">${escapeHtml_(stripHtml_(clickTxt))}</a>`;
    return '';
  }

  // ---- signatures: raw image only ("Signature Image", etc.) ----
  if (/^signature(\s*\d+)?\s*image$/i.test(key)) {
    const blob = getImageFromKey_(key, cfg, blocks);
    if (!blob) return '';
    const cidName = key.toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'');
    const src = addInlineImage_(bag, cidName, blob);
    const h = Number(cfg.Signature_Height_Px || 72);
    return `<div style="display:inline-block;vertical-align:top;margin:10px 18px 0 0;text-align:left;">
              <img src="${src}" alt="${escapeHtml_(key)}" style="height:${h}px;display:block;">
            </div>`;
  }

  // ---- Signatures side-by-side (borderless; names left-justified) ----
  if (/^signatures?(?:\s*(row|side\s*by\s*side))?$/i.test(key)) {
    function sig(n) {
      const blob  = getImageFromKey_(`Signature ${n} Image`, cfg, blocks);
      const src   = blob ? addInlineImage_(bag, `signature_${n}_img`, blob) : '';
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
    return `
      <table role="presentation" style="width:100%;border-collapse:collapse;margin-top:8px;">
        <tr>
          <td style="vertical-align:top;padding:0;border:0;">${sig(1)}</td>
          <td style="vertical-align:top;padding:0;border:0;">${sig(2)}</td>
        </tr>
      </table>`;
  }

  // ---- "Signature 1"/"Signature 2" half-width fallback ----
  if (/^signature\s*[12]$/i.test(key)) {
    const num   = /\d/.test(key) ? String(key).match(/\d/)[0] : '1';
    const blob  = getImageFromKey_(`Signature ${num} Image`, cfg, blocks);
    const name  = cfg[`Signature_${num}_Name`]  || '';
    const title = cfg[`Signature_${num}_Title`] || '';

    const src = blob ? addInlineImage_(bag, `signature_${num}_img`, blob) : '';
    const h   = Number(cfg.Signature_Height_Px || 72);

    const img  = src ? `<img src="${src}" alt="Signature ${num}" style="height:${h}px;max-width:100%;display:block;">` : '';
    const meta = (name || title)
      ? `<div style="margin-top:4px;"><div style="font-weight:700;text-align:left;">${escapeHtml_(name)}</div>${title ? `<div style="text-align:left;">${escapeHtml_(title)}</div>` : ''}</div>`
      : '';

    return `<div class="sig sig-${num}" style="display:inline-block;width:48%;min-width:220px;vertical-align:top;box-sizing:border-box;padding-right:8px;">${img}${meta}</div>`;
  }

  // ---- county pick-up (BRE) ----
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

  // ---- plain text Content Block (with tokens) ----
  let raw = (blocks.text && (blocks.text[key] || blocks.text[keyRaw])) || '';
  if (!raw) raw = tryFindTextLoose_(blocks.text, key);
  if (raw) return `<div>${replaceMergeTokens_(String(raw), row, cfg, blocks, extras)}</div>`;

  return '';
}

// ----- Letters (PDF) from ORDER -----
function renderVolunteerLetterHtml_FromOrder_(row, blocks, order, cfg) {
  const isSingle = String(row.Doc_Mode) === 'SINGLE';
  const colKey = isSingle ? SETTINGS.ORDER_KEYS.LETTER_SINGLE : SETTINGS.ORDER_KEYS.LETTER_MULTI;

  // For PDFs: use data URLs and tell blocks to skip inline "Logo" if present
  const built = renderFromOrder_(colKey, row, blocks, order, cfg, {
    imageMode: 'dataurl',
    forceHeaderLogo: true
  });

  // Optional PDF fallback: if header somehow doesn’t render, also place a small logo at top of body.
  // Enable by setting Config key: Logo_PDF_Fallback_In_Body = TRUE
  let fallbackTop = '';
  if (asBool_(cfg.Logo_PDF_Fallback_In_Body)) {
    const logoBlob = findLogoBlob_(cfg, blocks);
    if (logoBlob) {
      const src = addInlineImage_({ __useCid:false }, 'logo_pdf_fallback', logoBlob);
      const maxW = Number(cfg.Logo_Max_Width_Px || 320);
      fallbackTop = `<div style="text-align:center;margin:6px 0 10px;"><img src="${src}" alt="Logo" style="max-width:${maxW}px;height:auto;"></div>`;
    }
  }

  const css = `
    <style>
      @page { size: letter; margin: 0.6in; }
      body { font-family: Arial, Helvetica, sans-serif; font-size: 11pt; color: #222; }
      table.assign { width: 100%; border-collapse: collapse; margin: 8px 0; }
      table.assign th, table.assign td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; vertical-align: top; }
      table.assign th { background:#f0f0f0; }
      .page-break { page-break-before: always; }
    </style>
  `;

  return `<!doctype html><html><head><meta charset="utf-8" />${css}</head><body>${fallbackTop}${built.html}</body></html>`;
}


// ----- County packet PDF -----
function renderCountyPacketHtml_(county, volList, blocks, cfg, order) {
  const parts = volList.map((r, idx) => {
    const html = renderVolunteerLetterHtml_FromOrder_(r, blocks, order, cfg);
    return (idx===0) ? html : html.replace('<body>','<body><div class="page-break"></div>');
  });
  return parts.join('\n');
}

// ----- Assignment renderers -----
function renderAssignmentSingle_(row) {
  const d = row['A1_Date']||'', s=row['A1_Start']||'', e=row['A1_End']||'';
  const ln=row['A1_LocationName']||'', addr=row['A1_Address']||'';
  return `
    <div><b>Date/Time:</b> ${escapeHtml_(d)} ${escapeHtml_(s)}${e?('–'+escapeHtml_(e)) : ''}</div>
    <div><b>Location:</b> ${escapeHtml_(ln)}</div>
    <div><b>Address:</b> ${escapeHtml_(addr)}</div>
  `;
}

function renderAssignmentTable_Multiple_(row) {
  const tableStyle = 'border-collapse:collapse;width:100%;border:1px solid #d0d7de;';
  const thStyle = 'text-align:left;border:1px solid #d0d7de;padding:8px;background:#f6f8fa;font-weight:700;';
  const tdStyle = 'border:1px solid #d0d7de;padding:8px;vertical-align:top;';

  const rows = [];
  let day = 1; // kept for any logic you might add later; not rendered
  for (let i = 1; i <= SETTINGS.MAX_ASSIGNMENTS; i++) {
    const d    = row[`A${i}_Date`];
    const s    = row[`A${i}_Start`];
    const e    = row[`A${i}_End`];
    const ln   = row[`A${i}_LocationName`];
    const addr = row[`A${i}_Address`];

    if (!d && !s && !ln && !addr) break;

    const shift = [s || '', e ? ('–' + e) : ''].join('');
    rows.push(
      `<tr>
        <td style="${tdStyle}">${escapeHtml_(d || '')}</td>
        <td style="${tdStyle}">${escapeHtml_(shift)}</td>
        <td style="${tdStyle}">${escapeHtml_(ln || '')}</td>
        <td style="${tdStyle}">${escapeHtml_(addr || '')}</td>
      </tr>`
    );
    day++; // not displayed
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

function renderCountyEmailHtml_(breHeader, breRow, blocks, order, cfg) {
  // Subject: prefer ORDER subject, else Config fallback
  let subject = '';
  try {
    const tpl = getOrderTemplate_(order, SETTINGS.ORDER_KEYS.COUNTY_EMAIL) || {};
    if (tpl.subjectKey) {
      const fromCB = blocks && blocks.text && blocks.text[tpl.subjectKey];
      subject = fromCB ? stripHtml_(fromCB) : stripHtml_(tpl.subjectKey);
    }
  } catch (_) {}
  if (!subject) subject = cfg.County_Subject || SETTINGS.SUBJECTS.COUNTY_M3;

  // Pull proper county name from BRE row
  const iCounty = breHeader ? breHeader.findIndex(h => String(h).trim().toLowerCase() === 'county') : -1;
  const countyProper = (iCounty >= 0 && breRow) ? (breRow[iCounty] || '').toString().trim() : '';

  // Minimal row so merge tokens like *|County|* resolve to the BRE value
  const row = { County: countyProper };

  // Inline logo (CID) at the top
  const bag = { __useCid: true };
  const logoHtml = blockToHtml_('Logo', row, blocks, cfg, /*extras*/ null, bag) || '';

  // Exact body you requested (tokens will be merged from BRE via extras)
  const bodyTpl = [
    `${logoHtml}`,
    `<p>Dear *|County|* Elections Office,</p>`,
    `<p>Pursuant to O.C.G.A. § 21-2-408, I am writing on behalf of the Democratic Party of Georgia to designate *|Voting Type|* poll watchers for the *|Election Date|* *|Election Title|*.</p>`,
    `<p>Please see the attached poll watching list and designation letters.</p>`,
    `<p><b>Regarding poll watcher badges:</b> *|BRE Badge Blurb|*</p>`,
    `<p>Thank you and please let us know if you have any questions.</p>`,
    `<p>Best,</p>`,
    `<p>Cecilia Ugarte Baldwin<br>Voter Protection Director</p>`
  ].join('\n');

  const html = replaceMergeTokens_(bodyTpl, row, cfg, blocks, { breHeader, breRow });

  const htmlBody = `
    <div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#222;line-height:1.45">
      ${html}
    </div>`;

  return { subject, html: htmlBody, inlineImages: cleanInlineImages_(bag) };
}


/** =========================
 *  HTML → PDF helpers
 *  ========================= */
function htmlToPdf_(html, fileName, folder, cfg, blocks) {
  const folderId = folder.getId();
  const gdocId = createGoogleDocFromHtml_(html, fileName, folderId);

  // Wait a moment for HTML→Doc conversion
  Utilities.sleep(800);

  // Insert header logo, then wait for the revision to commit
  try { insertHeaderLogo_(gdocId, cfg, blocks); } catch (e) { try { Logger.log('insertHeaderLogo_ error: ' + e); } catch(_){} }
  Utilities.sleep(800);

  const pdfBlob = exportGDocToPdf_(gdocId, fileName);
  try { trashFile_(gdocId); } catch (e) {}
  return pdfBlob;
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

  // v3 export via UrlFetch
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

function trashFile_(fileId) {
  try {
    if (Drive.Files && typeof Drive.Files.trash === 'function') {
      Drive.Files.trash(fileId); // v2
      return;
    }
  } catch (e) {}
  try {
    if (Drive.Files && typeof Drive.Files.update === 'function') {
      Drive.Files.update({ trashed: true }, fileId); // v3
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

  // v3
  const resource = { name, mimeType: 'application/vnd.google-apps.document', parents: [folderId] };
  const gdoc = Drive.Files.create(resource, blobHtml, {supportsAllDrives: true});
  return gdoc.id;
}

// Inserts the org logo into the Google Doc HEADER (centered, fixed width)
function insertHeaderLogo_(gdocId, cfg, blocks) {
  const logoBlob = findLogoBlob_(cfg, blocks);
  if (!logoBlob) return;

  // Retry in case the doc is still finishing conversion
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const doc = DocumentApp.openById(gdocId);
      let header = doc.getHeader();
      if (!header) header = doc.addHeader();
      header.clear();

      const img = header.appendImage(logoBlob);
      const widthPx = Number(cfg.Logo_Header_Width_Px || cfg.Logo_Max_Width_Px || 180);
      img.setWidth(widthPx);
      img.getParent().asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);

      doc.saveAndClose();
      return; // success
    } catch (e) {
      // brief backoff then retry
      Utilities.sleep(600);
      if (attempt === 2) throw e;
    }
  }
}



/** ICS builder */
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

/** ============================
 *  4) EMAIL SENDER
 *  ============================ */
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
    name: cfg.From_Name || 'Credentialing',
    htmlBody,
    replyTo: replyTo || '',
    attachments
  };
  if (finalCc) opts.cc = finalCc;
  if (Object.keys(imgs).length) opts.inlineImages = imgs;
  
  log_('sendEmail', {
    testMode,
    finalTo,
    finalCc,
    subject: finalSubject,
    attachmentsCount: (attachments||[]).length,
    inlineImagesCount: Object.keys(imgs||{}).length
  });

  MailApp.sendEmail(finalTo, finalSubject, '(HTML only)', opts);
}

/** ======================
 *  5) COUNTY CSV BUILDER
 *  ====================== */
function buildCountyCsv_(county, volList, cfg) {
  const header = ['Volunteer_Name','VAN_ID','Email','Phone','Date','Start','End','Location_Name','Location_Address','County'];
  const lines = [header.join(',')];
  volList.forEach(v => {
    for (let i=1;i<=SETTINGS.MAX_ASSIGNMENTS;i++){
      const d=v[`A${i}_Date`], s=v[`A${i}_Start`], e=v[`A${i}_End`], ln=v[`A${i}_LocationName`], addr=v[`A${i}_Address`];
      if (!d && !s && !ln) break;
      const row = [v.Volunteer_Name||'', v.VAN_ID||'', v.Volunteer_Email||'', v.Volunteer_Phone||'',
                   d||'', s||'', e||'', ln||'', addr||'', county||''].map(csvEscape_);
      lines.push(row.join(','));
    }
  });
  return Utilities.newBlob(lines.join('\r\n'), 'text/csv', `County_${county}_${cfg.Run_Year||''}.csv`);
}

/** ==============
 *  6) CONTENT BLOCKS (TEXT + IMAGES)
 *  ============== */
function readContentBlocksAdvanced_() {
  const ss = SpreadsheetApp.getActive();
  const sh = getTab_(ss, SETTINGS.TABS.CONTENT_BLOCKS);
  const lastRow = Math.max(1, sh.getLastRow());
  const text = {};
  const images = {};

  if (lastRow < 2) return { text, images };

  // map over-grid images by anchor cell
  const overImgs = sh.getImages ? sh.getImages() : [];
  const imgByCell = new Map(); // "r,c" -> blob
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

    // TEXT: keep inline links/bold/italic/underline + line breaks
    let html = '';
    if (rich) {
      const runs = rich.getRuns();
      for (const run of runs) {
        const txt = escapeHtml_(run.getText() || '');
        const style = run.getTextStyle ? run.getTextStyle() : null;
        const link = run.getLinkUrl && run.getLinkUrl();
        let piece = txt;
        if (link) piece = `<a href="${link}" target="_blank" rel="noopener noreferrer" style="color:#1155cc;text-decoration:underline;">${txt}</a>`;
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

    // IMAGES: over-grid → =IMAGE("url") → Drive file id in text
    let blob = null;
    const mapKey = `${r},2`;
    if (imgByCell.has(mapKey)) blob = imgByCell.get(mapKey);

    if (!blob) {
      const formula = cell.getFormula();
      if (formula && /^=IMAGE\(/i.test(formula)) {
        const m = formula.match(/=IMAGE\(\s*"([^"]+)"/i);
        if (m && m[1]) { try { blob = UrlFetchApp.fetch(m[1]).getBlob(); } catch (_) {} }
      }
    }
    if (!blob) {
      const raw = cell.getValue();
      const asText = (typeof raw === 'string' || typeof raw === 'number') ? String(raw) : '';
      if (asText && /^[A-Za-z0-9_-]{25,}$/.test(asText)) { try { blob = DriveApp.getFileById(asText).getBlob(); } catch (_) {} }
    }
    if (blob) images[key] = blob;
  }

  return { text, images };
}

/** ==============
 *  7) UTILITIES
 *  ============== */
function ensureSheets_() {
  const ss = SpreadsheetApp.getActive();
  getOrCreateWithHeader_(ss, SETTINGS.TABS.MASTER, SETTINGS.MASTER_HEADERS);
  getOrCreateWithHeader_(ss, SETTINGS.TABS.COUNTY_ROLLUP, SETTINGS.COUNTY_HEADERS);
  getOrCreate_(ss, SETTINGS.TABS.LOGS);
  getTab_(ss, SETTINGS.TABS.P4_EMAIL_ORDER); // presence check
}

function readConfig_() {
  const ss = SpreadsheetApp.getActive();
  const sh = getTab_(ss, SETTINGS.TABS.CONFIG);

  const data = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), Math.max(2, sh.getLastColumn()))
                 .getDisplayValues();

  const cfg = {};
  for (let i = 1; i < data.length; i++) {
    const k = (data[i][0] || '').toString().trim();
    const v = (data[i][1] || '').toString().trim();
    if (k) cfg[k] = v;
  }

  cfg.Volunteers_Output_Folder_ID = cfg.Volunteers_Output_Folder_ID || SETTINGS.OUTPUT.VOLUNTEERS_FOLDER_ID;
  cfg.Counties_Output_Folder_ID   = cfg.Counties_Output_Folder_ID   || SETTINGS.OUTPUT.COUNTIES_FOLDER_ID;
  cfg.Calendar_Timezone          = cfg.Calendar_Timezone            || Session.getScriptTimeZone() || 'America/New_York';
  return { cfg };
}

// BRE helpers
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

// --- BRE cache + lookup by county ---
var __BRE_CACHE = null;

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
function sha256_(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str);
  return raw.map(b=>{ const s=(b<0? b+256:b).toString(16); return s.length===1?'0'+s:s; }).join('');
}
function icsEscape_(s){ return (s||'').replace(/[,;]/g, '\\$&'); }
function csvEscape_(s){ const v=(s||'').toString(); return /[",\n]/.test(v) ? `"${v.replace(/"/g,'""')}"` : v; }
function escapeHtml_(s){ return (s||'').toString().replace(/[&<>"']/g, c=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function stripHtml_(s){ return String(s||'').replace(/<[^>]*>/g,''); }

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
function log_(action, payload){ const ss=SpreadsheetApp.getActive(); const sh=getOrCreate_(ss, SETTINGS.TABS.LOGS); sh.appendRow([new Date(), action, JSON.stringify(payload||{})]); }
function getCellByA1_(ss, a1){ if (!a1) return ''; const [tab, cell] = a1.split('!'); const sh=ss.getSheetByName(tab); if (!sh) return ''; return (sh.getRange(cell).getValue()||'').toString(); }

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

/** Normalize + DEDUPE + sort -> stable canonical JSON for hashing. */
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

function asBool_(v) {
  if (v === true)  return true;
  if (v === false) return false;
  const s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === 't' || s === 'yes' || s === 'y' || s === '1';
}

function parseJsonSafe_(s) {
  try { return JSON.parse(String(s || '{}')); } catch (e) { return {}; }
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

/** Robust hash reader for prior/new Assignments_JSON fields. */
function assignmentsHashFromJson_(jsonStr) {
  const parsed = parseJsonSafe_(jsonStr);

  // If it's already our normalized array [{d,s,e,n,a}, ...]
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

  // Legacy shapes: {assignments:[...]}, {items:[...]} etc.
  const arr = (parsed && (parsed.assignments || parsed.items)) || [];
  return sha256_(canonicalizeAssignments_(arr));
}



// Works for both emails (CID) and letters/PDFs (data URL).
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

// Robust image lookup: Content Blocks, then Config (File ID or URL)
function resolveBlobFromRef_(ref) {
  if (!ref) return null;
  const s = String(ref).trim();
  try {
    if (/^[A-Za-z0-9_-]{25,}$/.test(s)) {            // Drive file id
      return DriveApp.getFileById(s).getBlob();
    }
    if (/^https?:\/\//i.test(s)) {                   // URL
      const resp = UrlFetchApp.fetch(s, {muteHttpExceptions:true});
      if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) return resp.getBlob();
    }
  } catch (_) {}
  return null;
}

function getImageFromKey_(keyRaw, cfg, blocks) {
  const key = String(keyRaw || '').trim();
  if (!key) return null;

  // Prefer exact or loose match from Content Blocks first
  const fromBlocksExact = (blocks && blocks.images && (blocks.images[key] || blocks.images[keyRaw])) || null;
  if (fromBlocksExact) return fromBlocksExact;

  // If caller is asking for any flavor of logo, use the tolerant finder
  const norm = key.toLowerCase();
  if (norm.includes('logo')) {
    const b = findLogoBlob_(cfg, blocks);
    if (b) return b;
  }

  // Signature aliases remain as-is
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

  // Last chance: try loose match from Content Blocks by similarity to the requested key
  const loose = tryFindImageLoose_(blocks, key);
  return loose || null;
}


function extractLastName_(fullName) {
  const parts = String(fullName || '').trim().split(/\s+/);
  return parts.length ? parts[parts.length - 1] : 'Unknown';
}
function fileSafe_(s) {
  return String(s || '').replace(/[\\/:*?"<>|]+/g, '').trim();
}

// Extremely tolerant logo lookup: Content Blocks (any key containing "logo"),
// then Config (scan for any key with "logo" plus "image|url|file|id"), then
// some common aliases. Returns a Blob or null.
function findLogoBlob_(cfg, blocks) {
  // 1) Content Blocks images: any key containing "logo"
  if (blocks && blocks.images) {
    for (const k of Object.keys(blocks.images)) {
      if (String(k).toLowerCase().includes('logo')) return blocks.images[k];
    }
  }

  // 2) Config: scan all keys that look like logo refs
  if (cfg) {
    for (const k of Object.keys(cfg)) {
      const kl = k.toLowerCase();
      if (kl.includes('logo') && (kl.includes('image') || kl.includes('url') || kl.includes('file') || kl.includes('id'))) {
        const blob = resolveBlobFromRef_(cfg[k]);
        if (blob) return blob;
      }
    }
  }

  // 3) Common aliases as a final fallback
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

function wrapWithSpacer_(html) {
  if (!html) return '';
  // ~1 line of visual space in email/PDF; tweak 12px if you want more/less.
  return `<div style="margin:0 0 12px 0;">${html}</div>`;
}

function normCountyKey_(s){
  let t = String(s || '').toLowerCase().trim();
  t = t.replace(/^ga[\s\-_]*/, '');   // drop "ga-"
  t = t.replace(/\bcounty\b/g, '');   // drop trailing "county"
  t = t.replace(/[^a-z0-9]+/g, ' ');  // unify separators
  return t.trim();
}

// --- Name normalization & matching -----------------------------
function normalizeCountyCore_(s) {
  return String(s||'')
    .trim()
    .replace(/^ga[-_\s]*/i, '')        // strip "ga-"
    .replace(/\s+/g,' ')               // collapse spaces
    .replace(/[^\w\s-]/g,'')           // remove punctuation
    .trim();
}
function countySlugVariants_(county) {
  const core = normalizeCountyCore_(county).toLowerCase();
  const underscore = core.replace(/\s+/g,'_');
  const hyphen     = core.replace(/\s+/g,'-');
  const prefUnderscore = `ga-${core}`.replace(/\s+/g,'_');
  const prefHyphen     = `ga-${core}`.replace(/\s+/g,'-');
  return new Set([
    underscore, hyphen,
    `ga_${underscore}`, `ga-${hyphen}`,
    `ga${underscore.startsWith('_')?'':'_'}${underscore}`,
    `ga${hyphen.startsWith('-')?'':'-'}${hyphen}`
  ]);
}

// --- Drive fetchers --------------------------------------------
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

    // only consider files with "county_" prefix to avoid single-volunteer PDFs
    if (!/^county[_\s-]/.test(name)) continue;

    // must contain one of our slug variants and have _<something> after it
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

// --- Ensure inline logo when HTML references cid:logo_header ---
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


/** ==== Utility Commands ==== */
function utilRecomputeHashes() {
  const ss = SpreadsheetApp.getActive();
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
  const ss = SpreadsheetApp.getActive();
  const sh = getTab_(ss, SETTINGS.TABS.MASTER);
  const vals = getData_(sh);
  const header = vals[0] || [];
  const rows = vals.slice(1);
  const idxFlag = header.indexOf('Needs_Credential_Send');
  if (idxFlag<0) return;
  rows.forEach(r => r[idxFlag] = false);
  if (rows.length) sh.getRange(2,1,rows.length, header.length).setValues(rows);
  toast_('All send flags cleared.', 5);
}

function utilExportLogs() {
  const ss = SpreadsheetApp.getActive();
  const sh = getOrCreate_(ss, SETTINGS.TABS.LOGS);
  const csv = getData_(sh).map(r=>r.map(csvEscape_).join(',')).join('\r\n');
  const blob = Utilities.newBlob(csv, 'text/csv', `Project4_Logs_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.csv`);
  DriveApp.createFile(blob);
  toast_('Logs exported to Drive root.', 5);
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

// County M3 Test (folder-driven attachments + guaranteed inline logo)
function utilSendCountyM3_TestOne(countyName) {
  const { cfg } = readConfig_();
  const blocks  = readContentBlocksAdvanced_();
  const order   = readEmailOrder_();

  // BRE row (accept "ga-dekalb" | "DeKalb")
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

  // --- Attachments from Counties_Output_Folder_ID (regen if missing)
  let { pdfBlob, csvBlob } = findLatestCountyArtifactsFromFolder_(row[iCounty], cfg);

  if (!pdfBlob || !csvBlob) {
    const ss      = SpreadsheetApp.getActive();
    const master  = getTab_(ss, SETTINGS.TABS.MASTER);
    const valsT   = getData_(master);
    const valsD   = master.getRange(1,1, master.getLastRow()||1, master.getLastColumn()||1).getDisplayValues();
    const header  = valsT[0] || [];
    const asObjT  = valsT.slice(1).map(x=>objFromRow_(header,x));
    const asObjD  = valsD.slice(1).map(x=>objFromRow_(header,x));
    const core    = normalizeCountyCore_(row[iCounty]);

    const listTyped = asObjT.filter(x => normalizeCountyCore_(x.County||'') === core);
    const listDisp  = asObjD.filter(x => normalizeCountyCore_(x.County||'') === core);

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

  // --- Build body from ORDER/Content Blocks with BRE tokens
  const built = renderFromOrder_(SETTINGS.ORDER_KEYS.COUNTY_EMAIL, null, blocks, order, cfg, {
    breHeader, breRow: row
  });

  // Ensure cid:logo_header has a blob
  const inlineImagesFinal = ensureInlineLogoForCid_(built.html, built.inlineImages, cfg, blocks);

  // --- Send
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

  log_('utilSendCountyM3_TestOne', {
    county: row[iCounty],
    attachmentsCount: attachments.length,
    hasLogoInline: !!(inlineImagesFinal && inlineImagesFinal['logo_header'])
  });

  toast_('County M3 test queued for: ' + row[iCounty], 5);
}

function run_Test_M3() { utilSendCountyM3_TestOne('DeKalb'); }
