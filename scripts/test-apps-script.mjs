// Runs apps-script/Code.gs against an in-memory Google Sheets mock.
// No real spreadsheet, Telegram or email is touched.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

class Sheet {
  constructor(name) { this.name = name; this.cells = []; this.notes = []; this.maxRows = 1000; }
  getName() { return this.name; }
  getLastRow() { return this.cells.length; }
  getLastColumn() { return this.cells.reduce((m, r) => Math.max(m, r.length), 0); }
  getMaxRows() { return this.maxRows; }
  appendRow(row) { this.cells.push(Array.from(row)); this.notes.push([]); }
  setFrozenRows() {}
  cell(r, c) { return (this.cells[r - 1] || [])[c - 1] ?? ''; }
  set(r, c, v) {
    while (this.cells.length < r) { this.cells.push([]); this.notes.push([]); }
    const row = this.cells[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
  deleteColumn(c) {
    this.cells.forEach(r => r.splice(c - 1, 1));
    this.notes.forEach(r => r.splice(c - 1, 1));
  }
  moveColumns(range, dest) {
    const from = range.c;
    const move = rows => rows.forEach(r => {
      const width = Math.max(r.length, from, dest);
      while (r.length < width) r.push(undefined);
      const [v] = r.splice(from - 1, 1);
      r.splice(dest > from ? dest - 2 : dest - 1, 0, v);
    });
    move(this.cells); move(this.notes);
    this.cells.forEach(r => { for (let i = 0; i < r.length; i++) if (r[i] === undefined) r[i] = ''; });
  }
}

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  map(fn) { return Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => fn(this.r + i, this.c + j))); }
  getValues() { return this.map((r, c) => this.sh.cell(r, c)); }
  setValues(v) { v.forEach((row, i) => row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x))); return this; }
  setValue(v) { this.sh.set(this.r, this.c, v); return this; }
  getNotes() { return this.map((r, c) => (this.sh.notes[r - 1] || [])[c - 1] || ''); }
  setNotes(v) { v.forEach((row, i) => row.forEach((x, j) => this.setNoteAt(this.r + i, this.c + j, x))); return this; }
  setNote(v) { this.setNoteAt(this.r, this.c, v); return this; }
  setNoteAt(r, c, v) { while (this.sh.notes.length < r) this.sh.notes.push([]); this.sh.notes[r - 1][c - 1] = v; }
  setFontWeight() { return this; }
}

function load({ token = 'TEST-TOKEN' } = {}) {
  const sheets = [];
  const ss = {
    getSheetByName: n => sheets.find(s => s.name === n) || null,
    insertSheet: n => { const s = new Sheet(n); sheets.push(s); return s; },
    getSheets: () => sheets.slice()
  };
  const sent = { telegram: [], mail: [] };
  const fmt = (d, tz, p) => {
    const z = n => String(n).padStart(2, '0');
    return p === 'yyMMdd' ? String(d.getFullYear()).slice(2) + z(d.getMonth() + 1) + z(d.getDate())
      : d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate());
  };
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, getUi: () => ({}) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => k === 'TG_TOKEN' ? token : null }) },
    Utilities: { formatDate: fmt },
    Session: { getScriptTimeZone: () => 'Europe/Kyiv' },
    ContentService: { createTextOutput: t => ({ text: t, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    MailApp: { sendEmail: m => sent.mail.push(m), getRemainingDailyQuota: () => 100 },
    UrlFetchApp: { fetch: (url, o) => { sent.telegram.push({ url, ...o.payload }); return { getContentText: () => '{"ok":true}' }; } },
    Logger: { log() {} },
    JSON, Math, Date, String, Number, Object, parseInt, Error
  };
  vm.createContext(ctx);
  vm.runInContext(readFileSync(new URL('../apps-script/Code.gs', import.meta.url), 'utf8'), ctx);
  return { ctx, ss, sheets, sent };
}

const post = (ctx, d) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(d) } }).text);

const LEAD = {
  fullName: 'Тест Тестенко', phone: '+380674704617', email: 'Oktobus109@gmail.com',
  fromCity: 'Львів', toCity: 'Bologna', routeLabel: 'Львів → Bologna', route: 'lviv-bologna',
  date: '2026-09-08', dateLabel: '08.09.2026', departureDate: '2026-09-09', departureDateLabel: '09.09.2026',
  seats: '14', paxCount: 1, comment: 'Тестова заявка', channel: 'Сайт', lang: 'ua',
  extraKg: 0, extraKgEur: 0, totalEur: 130, passengersFlat: '1) Тест Тестенко · +380674704617 · 14'
};

const EXPECTED_HEADERS = ['Дата поїздки', 'Ім’я та прізвище', 'Телефон', 'Email', 'Звідки', 'Куди', 'Дата заявки',
  'Місце', 'Кількість пасажирів', 'Коментар', 'Канал', 'Мова', 'Додатковий багаж, кг', 'Доплата за багаж, €', 'Сума, €', 'Усі пасажири'];

let groups = 0;
const pass = m => { groups++; console.log('PASS ' + m); };

// 1. Telegram text: exact compact format, no emoji, real destination, total as sent.
{
  const { ctx, sent } = load();
  assert.equal(ctx.telegramText_(LEAD), '08.09.2026\nЛьвів → Bologna\nТест Тестенко\n+380674704617\nМісце: 14\nOktobus109@gmail.com\nСума: 130 €');
  assert.equal(ctx.telegramText_({ ...LEAD, toCity: 'Fano', totalEur: 170 }).split('\n')[1], 'Львів → Fano');
  assert.equal(ctx.telegramText_({ ...LEAD, toCity: 'Fano', totalEur: 170 }).split('\n')[6], 'Сума: 170 €');
  assert.equal(ctx.telegramText_({ ...LEAD, seats: '14, 15', paxCount: 2 }).split('\n')[4], 'Місце: 14, 15');
  assert.equal(ctx.telegramText_({ ...LEAD, seats: '' }).split('\n')[4], 'Місце: —');
  assert.equal(ctx.telegramText_({ ...LEAD, email: '' }).split('\n').length, 6);
  assert.equal(ctx.telegramText_({ ...LEAD, dateLabel: '', date: '2026-09-08' }).split('\n')[0], '08.09.2026');
  assert.equal(ctx.telegramText_({ ...LEAD, totalEur: 131.5 }).split('\n').pop(), 'Сума: 131.5 €');
  const emoji = /[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{FE0F}]/u;
  ctx.testTelegram();
  const msg = sent.telegram[0];
  assert(!emoji.test(msg.text), msg.text);
  assert(!emoji.test(msg.reply_markup), msg.reply_markup);
  assert.equal(msg.parse_mode, undefined);
  for (const label of ['Нове бронювання', 'Маршрут:', 'Дата:', 'Телефон:', 'Email:', 'APP-', 'TEST-0001']) assert(!msg.text.includes(label), label);
  assert(msg.url.endsWith('/botTEST-TOKEN/sendMessage'));
  const noToken = load({ token: null });
  noToken.ctx.testTelegram();
  assert.equal(noToken.sent.telegram.length, 0);
  pass('Telegram: exact 7-line format, no emoji or labels, real city, multiple seats, total from booking');
}

// 2. New sheet: header and every column of a real doPost row.
{
  const { ctx, sheets, sent } = load();
  const res = post(ctx, LEAD);
  assert.equal(res.success, true);
  assert.match(res.id, /^APP-\d{6}-\d{4}$/);
  const sh = sheets.find(s => s.name === 'ЗАЯВКИ');
  assert.deepEqual(sh.cells[0], EXPECTED_HEADERS);
  assert(!sh.cells[0].includes('ID'));
  const row = Object.fromEntries(sh.cells[0].map((h, i) => [h, sh.cells[1][i]]));
  assert.equal(sh.cells[1].length, EXPECTED_HEADERS.length);
  assert.equal(row['Дата поїздки'], '08.09.2026');
  assert(row['Дата заявки'] instanceof Date);
  assert.equal(row['Ім’я та прізвище'], 'Тест Тестенко');
  assert.equal(row['Телефон'], "'+380674704617");
  assert.equal(row['Email'], 'Oktobus109@gmail.com');
  assert.equal(row['Звідки'], 'Львів');
  assert.equal(row['Куди'], 'Bologna');
  assert.equal(row['Місце'], '14');
  assert.equal(row['Кількість пасажирів'], 1);
  assert.equal(row['Коментар'], 'Тестова заявка');
  assert.equal(row['Канал'], 'Сайт');
  assert.equal(row['Мова'], 'ua');
  assert.equal(row['Додатковий багаж, кг'], 0);
  assert.equal(row['Доплата за багаж, €'], 0);
  assert.equal(row['Сума, €'], 130);
  assert.equal(row['Усі пасажири'], '1) Тест Тестенко · +380674704617 · 14');
  assert(!sh.cells[1].some(v => String(v).startsWith('APP-')));
  assert.equal(sh.notes[1][6], res.id, 'internal ID kept as a note on «Дата заявки»');
  const flight = sheets.find(s => s.name === 'Рейс 08.09.2026');
  assert.deepEqual(flight.cells[0], EXPECTED_HEADERS);
  assert.deepEqual(flight.cells[1], sh.cells[1]);
  // internal ID still used everywhere else
  assert(sent.mail.some(m => m.subject.includes(res.id)));
  assert.equal(sent.telegram.length, 1);
  // duplicate detection by phone + trip date still returns the same internal ID
  const dup = post(ctx, { ...LEAD, phone: '0674704617' });
  assert.deepEqual([dup.success, dup.error, dup.id], [false, 'duplicate', res.id]);
  assert.equal(sh.cells.length, 2);
  assert.equal(post(ctx, { ...LEAD, date: '2026-09-15', dateLabel: '15.09.2026' }).success, true);
  pass('new lead table: no ID column, dates swapped, all 16 columns match headers, ID kept for duplicates and emails');
}

// 3. Existing sheet in the OLD layout, before migration: new rows still land under the right headers.
const OLD_HEADERS = ['ID', 'Дата заявки', 'Ім’я та прізвище', 'Телефон', 'Email', 'Звідки', 'Куди', 'Дата поїздки', 'Місце',
  'Кількість пасажирів', 'Коментар', 'Канал', 'Мова', 'Додатковий багаж, кг', 'Доплата за багаж, €', 'Сума, €', 'Усі пасажири'];
const oldRow = (id, name, phone, trip) => [id, new Date('2026-10-01T10:00:00Z'), name, "'" + phone, 'old@mail.com', 'Київ', 'Rimini',
  trip, '7', 1, 'стара', 'Сайт', 'ua', 2, 3, 143, '1) old'];
{
  const { ctx, ss } = load();
  const sh = ss.insertSheet('ЗАЯВКИ');
  sh.appendRow(OLD_HEADERS);
  sh.appendRow(oldRow('APP-261001-1111', 'Стара Заявка', '+380501112233', '13.10.2026'));
  const res = post(ctx, LEAD);
  assert.equal(res.success, true);
  const row = Object.fromEntries(sh.cells[0].map((h, i) => [h, sh.cells[2][i]]));
  assert.equal(row.ID, '');
  assert.equal(row['Дата поїздки'], '08.09.2026');
  assert(row['Дата заявки'] instanceof Date);
  assert.equal(row['Телефон'], "'+380674704617");
  assert.equal(row['Куди'], 'Bologna');
  assert.equal(row['Сума, €'], 130);
  const dup = post(ctx, { ...LEAD, phone: '+380501112233', dateLabel: '13.10.2026', date: '2026-10-13' });
  assert.deepEqual([dup.error, dup.id], ['duplicate', 'APP-261001-1111']);
  pass('unmigrated old sheet: rows written by header name, ID cell left empty, old duplicates still found');
}

// 4. Migration of an old sheet: data moves with its header, IDs become notes, nothing shifts.
{
  const { ctx, ss } = load();
  const sh = ss.insertSheet('ЗАЯВКИ');
  const flight = ss.insertSheet('Рейс 13.10.2026');
  const seats = ss.insertSheet('ПОСАДОЧНІ МІСЦЯ');
  seats.appendRow(['Маршрут', 'Дата', 'Місце', 'Статус']);
  for (const s of [sh, flight]) {
    s.appendRow([...OLD_HEADERS, 'Нотатки менеджера']);
    s.appendRow([...oldRow('APP-261001-1111', 'Перша Особа', '+380501112233', '13.10.2026'), 'дзвонив']);
    s.appendRow([...oldRow('APP-261002-2222', 'Друга Особа', '+380671234567', '13.10.2026'), '']);
  }
  const before = sh.cells.map(r => r.slice());
  ctx.migrateLeadSheetsDryRun();
  assert.deepEqual(sh.cells, before, 'dry run must not change anything');
  ctx.migrateLeadSheets();
  for (const s of [sh, flight]) {
    assert.deepEqual(s.cells[0], [...EXPECTED_HEADERS, 'Нотатки менеджера']);
    const rec = i => Object.fromEntries(s.cells[0].map((h, c) => [h, s.cells[i][c]]));
    assert.equal(rec(1)['Дата поїздки'], '13.10.2026');
    assert(rec(1)['Дата заявки'] instanceof Date);
    assert.equal(rec(1)['Ім’я та прізвище'], 'Перша Особа');
    assert.equal(rec(1)['Телефон'], "'+380501112233");
    assert.equal(rec(1)['Email'], 'old@mail.com');
    assert.equal(rec(1)['Куди'], 'Rimini');
    assert.equal(rec(1)['Сума, €'], 143);
    assert.equal(rec(1)['Нотатки менеджера'], 'дзвонив');
    assert.equal(rec(2)['Ім’я та прізвище'], 'Друга Особа');
    assert.equal(s.notes[1][6], 'APP-261001-1111');
    assert.equal(s.notes[2][6], 'APP-261002-2222');
  }
  assert.deepEqual(seats.cells[0], ['Маршрут', 'Дата', 'Місце', 'Статус']);
  // after migration: new leads and duplicates work, running migration again changes nothing
  const res = post(ctx, LEAD);
  assert.equal(res.success, true);
  assert.equal(sh.cells[3][0], '08.09.2026');
  assert.equal(sh.cells[3][2], "'+380674704617");
  assert.equal(sh.cells[3][16], '');
  const dup = post(ctx, { ...LEAD, phone: '+380671234567', dateLabel: '13.10.2026', date: '2026-10-13' });
  assert.deepEqual([dup.error, dup.id], ['duplicate', 'APP-261002-2222']);
  const snapshot = JSON.stringify(sh.cells);
  ctx.migrateLeadSheets();
  assert.equal(JSON.stringify(sh.cells), snapshot);
  pass('migration: ID column removed into notes, dates swapped with their data, custom columns kept, rerun is a no-op');
}

console.log(`${groups} Apps Script groups passed (in-memory mock, nothing sent).`);
