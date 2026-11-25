// French tutoring sign-up backend. Lives inside the Google Sheet (Extensions → Apps Script).
// Sessions tab: one row per 1-hour session (start date + time, tutor, online tick box).
// Sign-ups tab: filled in by the website. Delete a row to remove a student.

const CAP = 3;        // max students per session
const CUTOFF_H = 24;  // no sign-ups or cancellations within this many hours of a session
const HOUR = 3600e3;
const SESSIONS = ['Session start', 'Tutor', 'Online'];
const SIGNUPS = ['Session', 'Name', 'Student', 'Parent phone', 'Signed up', 'Cancel code'];

// Run once from the editor: creates both tabs.
function setup() {
  const sessions = tab_('Sessions', SESSIONS);
  sessions.getRange('A2:A').setNumberFormat('ddd mmm d, yyyy h:mm am/pm');
  sessions.getRange('C2:C').insertCheckboxes();
  tab_('Sign-ups', SIGNUPS).getRange('A2:A').setNumberFormat('ddd mmm d, yyyy h:mm am/pm');
}

function doGet() {
  return json_(listSessions_().map(({ start, ...s }) => s));
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // one sign-up at a time, so two people can't take the last spot
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'book') return json_(book_(body));
    if (body.action === 'cancel') return json_(cancel_(body));
    return json_({ error: 'Unknown action.' });
  } catch (err) {
    console.error(err);
    return json_({ error: 'Something went wrong. Please try again.' });
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function listSessions_() {
  const taken = {};
  for (const r of rows_('Sign-ups', SIGNUPS)) {
    if (r[0] instanceof Date) taken[key_(r[0])] = (taken[key_(r[0])] || 0) + 1;
  }
  const now = new Date();
  return rows_('Sessions', SESSIONS)
    .filter(r => r[0] instanceof Date && r[0] > now)
    .sort((a, b) => a[0] - b[0])
    .map(([start, tutor, online]) => ({
      id: key_(start),
      day: fmt_(start, 'EEEE, MMMM d'),
      time: `${fmt_(start, 'h:mm a')}–${fmt_(new Date(+start + HOUR), 'h:mm a')}`,
      taken: taken[key_(start)] || 0,
      cap: CAP,
      closed: pastCutoff_(start),
      tutor: String(tutor).trim(),
      online: online === true,
      start,
    }));
}

function book_({ slot, name, student, phone }) {
  [name, student, phone] = [name, student, phone].map(v => String(v ?? '').trim());
  if (!name || !student || !phone) return { error: 'Please fill in every field.' };
  if (name.length > 80 || student.length > 80) return { error: 'Names must be under 80 characters.' };
  if (!/^[0-9+()\-.\s]{7,20}$/.test(phone)) return { error: 'Please enter a valid phone number.' };
  const s = listSessions_().find(s => s.id === slot);
  if (!s) return { error: 'That session is no longer available.' };
  if (s.closed) return { error: `Sign-ups close ${CUTOFF_H} hours before the session.` };
  if (s.taken >= CAP) return { error: 'Sorry, that session just filled up.' };
  const token = Utilities.getUuid();
  // The leading ' stores text exactly as typed, so a name starting with "=" can't run as a formula.
  tab_('Sign-ups', SIGNUPS).appendRow([s.start, "'" + name, "'" + student, "'" + phone, new Date(), token]);
  return { token, label: `${s.day}, ${s.time}` };
}

function cancel_({ token }) {
  token = String(token || '');
  const sheet = tab_('Sign-ups', SIGNUPS);
  const rows = sheet.getDataRange().getValues();
  const i = token ? rows.findIndex((r, i) => i > 0 && r[5] === token) : -1;
  if (i < 1) return { error: 'Sign-up not found. It may already be cancelled.' };
  if (pastCutoff_(rows[i][0])) return { error: `Cancellations close ${CUTOFF_H} hours before the session.` };
  sheet.deleteRow(i + 1);
  return { ok: true };
}

function tab_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

const rows_ = (name, headers) => tab_(name, headers).getDataRange().getValues().slice(1);
const fmt_ = (d, pattern) => Utilities.formatDate(d, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), pattern);
const key_ = d => fmt_(d, 'yyyy-MM-dd HH:mm');
const pastCutoff_ = d => d - Date.now() < CUTOFF_H * HOUR;
const json_ = obj => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
