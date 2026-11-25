// Runs Code.gs against a fake in-memory Google Sheet and checks the sign-up rules. Run: node test.js
const vm = require('node:vm');
const fs = require('node:fs');
const assert = require('node:assert');

// Just enough of Apps Script to run Code.gs.
function fakeAppsScript() {
  const sheets = {};
  const range = { setFontWeight: () => range, setNumberFormat: () => range, insertCheckboxes: () => range };
  const makeSheet = () => {
    const rows = [];
    return {
      rows,
      appendRow: r => rows.push(r.map(v => (typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v))),
      getDataRange: () => ({ getValues: () => rows.map(r => [...r]) }),
      deleteRow: n => rows.splice(n - 1, 1),
      setFrozenRows() {},
      getRange: () => range,
    };
  };
  const ss = {
    getSheetByName: n => sheets[n],
    insertSheet: n => (sheets[n] = makeSheet()),
    getSpreadsheetTimeZone: () => 'America/New_York',
  };
  const part = (d, opts) => new Intl.DateTimeFormat('en-US', { timeZone: 'America/New_York', ...opts }).format(d);
  const formats = {
    'yyyy-MM-dd HH:mm': d => d.toISOString(),
    'EEEE, MMMM d': d => part(d, { weekday: 'long', month: 'long', day: 'numeric' }),
    'h:mm a': d => part(d, { hour: 'numeric', minute: '2-digit' }),
  };
  const appended = [];
  return {
    sheets, appended,
    SpreadsheetApp: { getActive: () => ss, flush() {} },
    Utilities: { formatDate: (d, tz, p) => formats[p](d), getUuid: () => crypto.randomUUID() },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: text => ({ text, setMimeType() { return this; } }) },
    console: { error() {} }, // quiet: the bad-JSON check logs on purpose
  };
}

const g = fakeAppsScript();
vm.createContext(g);
vm.runInContext(fs.readFileSync(__dirname + '/Code.gs', 'utf8'), g);
const get = () => JSON.parse(g.doGet().text);
const post = body => JSON.parse(g.doPost({ postData: { contents: JSON.stringify(body) } }).text);

const SheetDate = vm.runInContext('Date', g); // dates must come from the script's own realm for instanceof
const H = 3600e3, now = Date.now();
g.setup();
const sessions = g.sheets['Sessions'].rows;
sessions.push([new SheetDate(now + 72 * H), 'Marie', true]);  // open
const inside24h = new SheetDate(now + 2 * H);
sessions.push([inside24h, '', false]);                   // inside 24h
sessions.push([new SheetDate(now + 96 * H), '', false]);      // open, no tutor
sessions.push([new SheetDate(now - 5 * H), '', false]);       // already happened
sessions.push(['', '', false]);                          // blank row

const slots = get();
assert.equal(slots.length, 3, 'past and blank rows are hidden');
const [soon, open, other] = slots;
assert.equal(soon.closed, true);
assert.equal(open.closed, false);
assert.equal(open.tutor, 'Marie');
assert.equal(open.online, true);
assert(!('start' in open), 'internal fields stay private');

const who = { name: 'Ann Lee', student: 'Sam Lee', phone: '(555) 123-4567' };
const tokens = [1, 2, 3].map(() => post({ action: 'book', slot: open.id, ...who }).token);
assert(tokens.every(Boolean), 'first three sign-ups work');
assert.match(post({ action: 'book', slot: open.id, ...who }).error, /filled up/, '4th is refused');
assert.match(post({ action: 'book', slot: open.id, ...who, phone: 'abc' }).error, /phone/);
assert.match(post({ action: 'book', slot: open.id, ...who, student: ' ' }).error, /every field/);
assert.match(post({ action: 'book', slot: soon.id, ...who }).error, /close 24 hours/);
assert.match(post({ action: 'book', slot: 'nope', ...who }).error, /no longer available/);
assert.equal(get()[1].taken, 3);

assert.equal(post({ action: 'cancel', token: tokens[0] }).ok, true);
assert.match(post({ action: 'cancel', token: tokens[0] }).error, /not found/, 'double cancel');
assert.match(post({ action: 'cancel', token: '' }).error, /not found/, 'blank code');
assert(post({ action: 'book', slot: open.id, ...who }).token, 'freed spot can be taken');

g.sheets['Sign-ups'].rows.push([inside24h, 'X', 'Y', '555 555 5555', new Date(), 'late']);
assert.match(post({ action: 'cancel', token: 'late' }).error, /close 24 hours/, 'no cancelling inside 24h');

// Typed text must reach the Sheet with a leading ' so "=..." can't become a formula.
const realAppend = g.sheets['Sign-ups'].appendRow;
g.sheets['Sign-ups'].appendRow = r => { g.appended.push(r); realAppend(r); };
post({ action: 'book', slot: other.id, ...who, name: '=IMPORTXML("x")' });
assert.equal(g.appended[0][1], `'=IMPORTXML("x")`);

assert.match(JSON.parse(g.doPost({ postData: { contents: 'not json' } }).text).error, /went wrong/);
console.log('all checks passed');
