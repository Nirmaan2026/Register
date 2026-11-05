/**
 * NIRMAAN 4.0 — registration receiver
 * Saves each registration as a row in this Google Sheet and the payment
 * screenshot in a Google Drive folder.
 *
 * Setup (once):
 *   1. Paste this whole file into Extensions → Apps Script (replace everything).
 *   2. Choose the function "setup" at the top and press Run. Allow access.
 *   3. Deploy → New deployment → type "Web app"
 *        Execute as: Me      Who has access: Anyone
 *      Press Deploy and copy the Web app URL (ends in /exec).
 */

const SHEET_NAME  = 'Registrations';
const FOLDER_NAME = 'NIRMAAN 4.0 Payment Screenshots';

// Must match the website. [early-bird fee, on-spot fee] in rupees; 0 = free event.
const EARLY_FEE_UNTIL = new Date('2026-11-10T23:59:59+05:30');
const FEES = {
  'Arch Tech (Arc Bridge)': [49, 69],
  'Blindfolded Builder (Blind O Civil)': [49, 69],
  'Civiotrix (Robotic Arm with Hydraulic System)': [49, 69],
  'Earthquake Challenge': [49, 69],
  'Engineering Maze': [49, 69],
  'Logo Undercover': [49, 69],
  'Mystery Matrix': [49, 69],
  'Pick and Speak': [49, 69],
  'Popsicle Bridge': [49, 69],
  'Site-to-Solution': [49, 69],
  'TechVOX (Tech Tark)': [49, 69],
  'Waste to Value Construction': [49, 69],
  'Be a Roadie': [59, 79],
  'BGMI': [100, 130],
  'Blind Date': [79, 99],
  'Chess': [49, 69],
  'COD': [100, 120],
  'Dance Mania': [69, 89],
  'Face Painting': [49, 69],
  'Fancy Flow': [69, 89],
  'Flavour Feast': [0, 0],
  'Free Fire': [100, 130],
  'Meme King': [49, 69],
  'Music Melodies': [49, 69],
  'Short Film Making': [99, 119],
  'Treasure Hunt': [149, 169],
  'Trade War': [149, 169],
  'UNO': [49, 69]
};

const HEADERS = ['Timestamp', 'Reg ID', 'Name', 'Mobile', 'Category', 'College', 'Year',
  'Branch / Dept', 'Roll No.', 'Event', 'Participation', 'Team members', 'Fee (₹)', 'Fee type', 'Payment screenshot',
  'Status', 'Checked by', 'Remarks'];
const COL = name => HEADERS.indexOf(name) + 1;

/** Run once from the editor: creates the sheet layout and the Drive folder. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#3b1d2c').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  const data = sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns());
  data.clearDataValidations();
  data.setNumberFormat('General');
  const status = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Verified', 'Rejected'], true).build();
  sh.getRange(2, COL('Status'), sh.getMaxRows() - 1, 1).setDataValidation(status);
  sh.getRange(2, COL('Mobile'), sh.getMaxRows() - 1, 1).setNumberFormat('@');
  sh.getRange(2, COL('Roll No.'), sh.getMaxRows() - 1, 1).setNumberFormat('@');
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  const folder = getFolder_();
  Logger.log('Ready. Screenshots folder: ' + folder.getUrl());
}

function doGet() {
  return json_({ ok: true, message: 'NIRMAAN 4.0 registration service is running.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);

    // ---- check the data
    const name = clean_(d.name, 80);
    const mobile = String(d.mobile || '').replace(/\D/g, '');
    const category = clean_(d.category, 40);
    const event = String(d.event || '');
    if (!name) throw new Error('Name is missing.');
    if (!/^[6-9]\d{9}$/.test(mobile)) throw new Error('Mobile number is not valid.');
    if (['GCE Keonjhar student', 'Other college student', 'Faculty'].indexOf(category) < 0) throw new Error('Category is not valid.');
    if (!(event in FEES)) throw new Error('Unknown event.');
    if (!FEES[event]) throw new Error('Registration for this event is not open yet.');
    const isStudent = category !== 'Faculty';
    if (isStudent && (!d.year || !clean_(d.roll, 40))) throw new Error('Year or roll number is missing.');
    const group = d.participation === 'Group';
    const team = group ? (Array.isArray(d.team) ? d.team : []).map(n => clean_(n, 80)).filter(String).slice(0, 9) : [];
    if (group && !team.length) throw new Error('Team member names are missing.');

    const early = new Date() <= EARLY_FEE_UNTIL;
    const fee = early ? FEES[event][0] : FEES[event][1];
    let bytes = null;
    if (fee > 0) {
      if (!d.shot || !d.shot.data || !/^image\//.test(d.shot.type || '')) throw new Error('Payment screenshot is missing.');
      bytes = Utilities.base64Decode(d.shot.data);
      if (bytes.length > 10 * 1024 * 1024) throw new Error('Screenshot is too large.');
    }

    lock.waitLock(30000);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sh) throw new Error('Sheet not set up. Run setup() first.');

    const last = sh.getLastRow();
    const regId = 'N4-' + String(last).padStart(4, '0');   // header is row 1 → first entry N4-0001
    const roll = clean_(d.roll, 40);

    // ---- save the screenshot
    let shotCell = 'Free event';
    if (bytes) {
      const ext = (d.shot.type.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
      const fileName = [regId, isStudent ? roll : mobile, event].join('_').replace(/[^\w.-]+/g, '-') + '.' + ext;
      const file = getFolder_().createFile(Utilities.newBlob(bytes, d.shot.type, fileName));
      shotCell = '=HYPERLINK("' + file.getUrl() + '","View")';
    }

    // ---- remarks for the checking team
    const remarks = [];
    if (last > 1) {
      const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
      rows.forEach(r => {
        if (String(r[COL('Mobile') - 1]) === mobile && r[COL('Event') - 1] === event) {
          remarks.push('Same mobile already registered for this event: ' + r[COL('Reg ID') - 1]);
        }
      });
    }
    if (Number(d.fee) !== fee) remarks.push('Page showed ₹' + d.fee);

    sh.appendRow([
      new Date(), regId, name, mobile, category,
      clean_(d.college, 120), isStudent ? clean_(d.year, 20) : '', clean_(d.branch, 80),
      isStudent ? roll : '', event, group ? 'Group of ' + (team.length + 1) : 'Solo', team.join('; '), fee, fee === 0 ? 'Free' : (early ? 'Early-bird' : 'On-spot'),
      shotCell,
      fee === 0 ? 'Verified' : 'Pending', '', remarks.join('; ')
    ]);

    return json_({ ok: true, id: regId, fee: fee });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const found = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = found.hasNext() ? found.next() : DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

// Text as plain text: trimmed, length-limited, never treated as a formula.
function clean_(v, max) {
  let s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
