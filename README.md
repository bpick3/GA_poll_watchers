# Project 4 — Operator Checklist

_A step-by-step runbook to build, review, and send volunteer credentials and county packets safely and consistently._

![Status](https://img.shields.io/badge/status-production-blue) ![Platform](https://img.shields.io/badge/platform-Google%20Apps%20Script-000) ![License](https://img.shields.io/badge/license-MIT-green)

---

## Quick Purpose

- Automate credential letters (PDF), calendar invites (ICS), and county packets (PDF + CSV).  
- Prevent accidental re-sends unless assignments **materially** change.  
- Keep content/layout centralized in **Content Blocks** and **Project 4 Email Order**.

---

## Table of Contents

- [Architecture](#architecture)
- [Security & Privacy](#security--privacy)
- [One-Time Setup (per cycle)](#one-time-setup-per-cycle)
- [Standard Run (each batch)](#standard-run-each-batch)
- [What Triggers a Re-send](#what-triggers-a-re-send-material-changes)
- [County Attachments Behavior](#county-attachments-behavior)
- [Troubleshooting](#troubleshooting)
- [Reference — Menu Map](#reference--menu-map)
- [FAQs](#faqs)
- [Contributing](#contributing)
- [License](#license)

---

## Architecture

> **Platform:** Google Sheets + Apps Script + Gmail + Google Drive

**Tabs used**
- `Config` — runtime configuration (names, emails, folder IDs, timezone, batch size, etc.).
- `Content Blocks` — reusable text blocks and images (logos/signatures).
- `Project 4 Email Order` — columns for M1/M2/M3 and Letters (Single/Multi), including subjects.
- `LBJ_Import (latest)` — latest assignment export (input).
- `Master_Assignments` — normalized, deduped volunteers + assignments; re-send flags.
- `County_Rollup` — one row per county with packet artifact references.
- `BRE Merge Sheet` — county recipients and merge data.
- `Logs` — audit rows for actions (sends, generations).

**Key features**
- Change-aware hashing of assignments to avoid unnecessary re-sends.
- One-click generation of PDF letters and County packets (PDF + CSV).
- Email sending flows for volunteers (M2) and counties (M3).
- Test Mode with `Test_Recipients` to validate safely.
- Inline logo handling (CID in emails; header injection for letters).

---

## Security & Privacy

- **Never commit** real spreadsheet data, Drive File IDs, or emails to a public repo.
- Keep all secrets/IDs/emails in the **`Config`** tab (not in source).
- Use **placeholder** IDs and example values in docs/screenshots.
- Logs may contain recipient emails—do not export or commit those.

> See `CONFIGURATION.md` (optional) for a redacted sample configuration pattern.

---

## One-Time Setup (per cycle)

### ✅ Config tab

- [ ] `From_Name` / `From_Email` / `Reply_To` filled.  
- [ ] `Run_Year` set.  
- [ ] `Volunteers_Output_Folder_ID` and `Counties_Output_Folder_ID` point to correct Drive folders.  
- [ ] `Calendar_Timezone` is correct (e.g., `America/New_York`).  
- [ ] `Test_Mode = TRUE` while testing; set to `FALSE` for live sends.  
- [ ] `Max_Batch_Size` set to your comfort level.  
- [ ] **Logo** and **Signature 1/2**: Names/Titles + Images (File ID or URL) configured.  
- [ ] (If using M1 confirmations) `Form_Base_URL` or `Form_Base_URL_Cell` present.

### ✅ Content & order

- [ ] **Content Blocks**: All required blocks present (greetings, assignment confirmation, badge blurbs, etc.).  
- [ ] **Project 4 Email Order**: Columns exist for **M1/M2/M3** + **Letters (Single/Multi)**; subject lines present.

### ✅ BRE

- [ ] `BRE_Workbook_ID` / `BRE_Merge_Tab_Name` (or leave blank/`ACTIVE` to use current workbook).  
- [ ] In BRE sheet: each county has **To** (and optional **CC**).  
- [ ] **Accuracy Check** is OK for counties you plan to email.

---

## Standard Run (each batch)

### 1) Refresh LBJ
- Paste the latest export into **`LBJ_Import (latest)`**.
- Confirm key columns exist (aliases supported):
  - `VAN_ID`, `First Name`, `Last Name`, `Email`, `Phone`, `County`
  - `Date`, `Start`, `End`, `Polling Location Name`, `Polling Location Address` (polling-site address)
  - `Address` (volunteer mailing address used on the letter)
- (Optional) If `Form_Base_URL` is set, the sheet will (re)build `Yes_Link` / `No_Link`.

### 2) Build/Update Master
Menu: **Project 4 → Build/Update Master**

- Generates/updates `Master_Assignments` with **deduped** assignments.  
- Computes the **assignment change hash** and sets `Needs_Credential_Send` only when assignments changed.

**Sanity checks in Master**
- [ ] `Volunteer_Address` populated (from LBJ Address).  
- [ ] `A1_… / A2_…` fields show correct dates/times/locations/addresses.  
- [ ] `Needs_Credential_Send` is **TRUE** only for new/changed assignments.

### 3) Preview Diffs
Menu: **Project 4 → Preview Diffs**

- Review who would be re-sent and any BRE issues (e.g., missing **To**, Accuracy Check ≠ OK).  
- Fix BRE issues before county sends.

### 4) Test path (highly recommended)
Set `Test_Mode = TRUE` and provide `Test_Recipients` (comma-separated).

- [ ] **Generate Volunteer PDFs** (optional preview)  
  Menu: **Project 4 → Generate Volunteer PDFs**  
  Check Drive folder: filenames follow `Credential_<LastName>_<VANID>.pdf`
- [ ] **Send Volunteer Emails — M2** (still in Test Mode)  
  Menu: **Project 4 → Send Volunteer Emails — M2**  
  Verify in test inbox: inline logo, spacing, signatures side-by-side, **PDF + ICS** attached.
- [ ] **Generate County Packets**  
  Menu: **Project 4 → Generate County Packets**  
  County PDF + CSV appear in Counties folder; `County_Rollup` updated.
- [ ] **Send County Emails — M3** (still in Test Mode)  
  Menu: **Project 4 → Send County Emails — M3**  
  Verify inline logo present; badge/pick-up blurbs correct; attachments attached.

### 5) Live send
- Set `Test_Mode = FALSE` in **Config**.  
- (Optional) Adjust `Max_Batch_Size`.  
- **Send Volunteer Emails — M2** → Only rows with `Needs_Credential_Send = TRUE` go out.  
- **Send County Emails — M3** → Sends to county contacts from BRE and records timestamps.

### 6) Post-run checks
**Master_Assignments**
- [ ] `Credential_Sent_On` / `Credential_Sent_By` populated.  
- [ ] `Needs_Credential_Send` flipped to **FALSE** for sent volunteers.

**County_Rollup**
- [ ] `County_Email_Sent_On` populated per county.

(Optionally) **Export Logs**  
Menu: **Project 4 → Utilities → Export Logs**.

---

## What Triggers a Re-send (Material Changes)

We re-send **only** if any assignment changes:
- **Date**, **Start**, **End**, **Location Name**, **Location Address**

**Non-material** (will **not** trigger re-send): name, email, phone, volunteer mailing address, etc.

**Force a resend (manual override):**
- In `Master_Assignments`, set `Needs_Credential_Send = TRUE` for that row; **or**
- Edit an assignment field; **or**
- (Nuclear) **Utilities → Recompute Hashes** (flags **all** rows to TRUE — use sparingly).

---

## County Attachments Behavior

- M3 sender looks for the latest `County_<County>_<Run_Year>.pdf` and matching `.csv` in the **Counties** folder.  
- If either is missing, it **regenerates** from `Master_Assignments` on the fly.  
- Inline logo is ensured whenever the email HTML references `cid:logo_header`.

---

## Troubleshooting

- **Flags unexpectedly TRUE after rebuild**  
  Check if LBJ assignment values changed display format (the system normalizes, but raw display changes can matter). If constant, flags will remain FALSE.

- **“Dear \|County\|” empty in test**  
  Use the M3 sender or `utilSendCountyM3_TestOne(<County>)` which passes the BRE row; ensure BRE has the county name.

- **Logo/signatures missing**  
  Confirm **Content Blocks** images or **Config** image refs (File IDs/URLs). For letters, header logo is injected during Doc conversion.

- **BRE edits not showing**  
  Run **Utilities → Clear BRE Cache** to refresh cached BRE data.

- **ICS duplicates**  
  Static `.ics` files can’t prevent de-dupe in recipients’ calendars. Avoid re-sending unless assignments truly change.

---

## Reference — Menu Map

- **Build/Update Master** — Rebuild `Master_Assignments` and change flags.  
- **Preview Diffs** — Show who would be re-sent + BRE issues.  
- **Generate Volunteer PDFs** — Create per-volunteer letter PDFs.  
- **Generate County Packets** — Build per-county PDF + CSV.  
- **Send Volunteer Emails — M2** — Email volunteers with PDF + ICS; respects flags and Test Mode.  
- **Send County Emails — M3** — Email counties with attachments; logs send time.  

**Utilities**
- **Recompute Hashes** — sets all flags TRUE  
- **Clear Send Flags** — sets all flags FALSE  
- **Export Logs**  
- **Clear BRE Cache**

---

## FAQs

**Q: Do I need Advanced Drive Service?**  
A: Some PDF export flows use `Drive.Files` (Advanced Drive). Enable in Apps Script > Services if required.

**Q: Will Test Mode stop real sends?**  
A: Yes. When `Test_Mode = TRUE`, messages go to `Test_Recipients` and subjects are prefixed with `[TEST]`.

**Q: How are re-sends prevented?**  
A: A normalized assignment JSON is hashed. Only material changes flip `Needs_Credential_Send` to TRUE.

---

## Contributing

- Open an issue describing the improvement or bug.  
- Submit a small, scoped PR with tests where feasible (unit tests for pure utils).  
- Do **not** include any PII, real folder IDs, or screenshots of live data.

---

## License

MIT © stac labs
