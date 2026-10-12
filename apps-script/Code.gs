/**
 * Avtobus v Italiu — місток між сайтом і Google Таблицею.
 *
 *   GET  ?action=seats&route=...&date=...  → статуси місць
 *   POST { ... }                           → рядок у «ЗАЯВКИ» + листи + Telegram
 */

var SHEET_APPS = 'ЗАЯВКИ';
var SHEET_SEATS = 'ПОСАДОЧНІ МІСЦЯ';
var SEAT_COUNT = 50;

var STATUS_FREE = 'ВІЛЬНЕ';
var STATUS_TAKEN = 'ЗАЙНЯТЕ';
var STATUS_BLOCKED = 'ЗАБЛОКОВАНЕ';

/** Номери менеджерів для листа клієнту. */
var WA_UA = '380674704617';
var WA_IT = '393245958718';

/** Куди надсилати листи про нові заявки. */
var NOTIFY_EMAILS = 'Oktobus109@gmail.com, alexalixmanov@gmail.com';

/**
 * Telegram: токен бота і чат, куди надсилати заявки.
 * Токен зберігається у властивостях скрипта (Налаштування проєкту → Властивості скрипта),
 * ключ TG_TOKEN, а не в коді.
 */
var TG_TOKEN = PropertiesService.getScriptProperties().getProperty('TG_TOKEN') || '';
var TG_CHAT = '-5432756087';

/**
 * Єдина схема таблиці заявок: заголовок → значення.
 * І заголовки, і рядки будуються тільки з цього списку, тому вони не можуть розійтися.
 * Колонки ID немає: номер заявки зберігається приміткою в клітинці «Дата заявки».
 */
var COL_TRIP_DATE = 'Дата поїздки';
var COL_CREATED = 'Дата заявки';
var COL_PHONE = 'Телефон';
var LEGACY_ID = 'ID';

var APP_COLUMNS = [
  [COL_TRIP_DATE,            function (d) { return clean_(d.dateLabel || d.date, 40); }],
  ['Ім’я та прізвище',       function (d) { return clean_(d.fullName, 120); }],
  [COL_PHONE,                function (d) { return clean_(d.phone, 40); }],
  ['Email',                  function (d) { return clean_(d.email, 120); }],
  ['Звідки',                 function (d) { return clean_(d.fromCity, 80); }],
  ['Куди',                   function (d) { return clean_(d.toCity, 80); }],
  [COL_CREATED,              function (d, now) { return now; }],
  ['Місце',                  function (d) { return clean_(d.seats, 60); }],
  ['Кількість пасажирів',    function (d) { return parseInt(d.paxCount, 10) || 1; }],
  ['Коментар',               function (d) { return clean_(d.comment, 500); }],
  ['Канал',                  function (d) { return clean_(d.channel, 40); }],
  ['Мова',                   function (d) { return clean_(d.lang, 10); }],
  ['Додатковий багаж, кг',   function (d) { return parseInt(d.extraKg, 10) || 0; }],
  ['Доплата за багаж, €',    function (d) { return Number(d.extraKgEur) || 0; }],
  ['Сума, €',                function (d) { return Number(d.totalEur) || 0; }],
  ['Усі пасажири',           function (d) { return clean_(d.passengersFlat, 900); }]
];

var APP_HEADERS = APP_COLUMNS.map(function (c) { return c[0]; });

var SEAT_HEADERS = ['Маршрут', 'Дата', 'Місце', 'Статус'];

/* ---------- допоміжні ---------- */

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

/** Номери колонок за заголовками першого рядка: { 'Телефон': 3, ... } (з 1). */
function headerIndex_(sh) {
  var width = sh.getLastColumn();
  var out = {};
  if (!width) return out;
  var row = sh.getRange(1, 1, 1, width).getValues()[0];
  for (var i = 0; i < row.length; i++) {
    var h = String(row[i]).trim();
    if (h && !out[h]) out[h] = i + 1;
  }
  return out;
}

/** Значення заявки за заголовками схеми. */
function leadRecord_(d, now) {
  var rec = {};
  APP_COLUMNS.forEach(function (c) { rec[c[0]] = c[1](d, now); });
  return rec;
}

/**
 * Дописує заявку в аркуш за ЙОГО заголовками, а не за позицією.
 * Якщо аркуш ще у старому форматі (з ID), дані все одно потраплять у свої колонки,
 * а клітинка ID залишиться порожньою. Відсутні заголовки додаються в кінець.
 */
function appendLead_(sh, rec, id) {
  var idx = headerIndex_(sh);
  APP_HEADERS.forEach(function (h) {
    if (!idx[h]) {
      var col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(h).setFontWeight('bold');
      idx[h] = col;
    }
  });
  var width = sh.getLastColumn();
  var row = [];
  for (var c = 0; c < width; c++) row.push('');
  APP_HEADERS.forEach(function (h) { row[idx[h] - 1] = rec[h]; });
  sh.appendRow(row);
  if (id) sh.getRange(sh.getLastRow(), idx[COL_CREATED]).setNote(id);
}

function norm_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v).trim().toLowerCase().replace(/\s+/g, ' ');
  var m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
  if (m) {
    return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  return s;
}

function normStatus_(v) {
  var s = String(v || '').trim().toUpperCase();
  if (s === STATUS_TAKEN) return 'taken';
  if (s === STATUS_BLOCKED) return 'blocked';
  return 'free';
}

function clean_(v, max) {
  var s = String(v === null || v === undefined ? '' : v);
  s = s.replace(/[\u0000-\u001F\u007F]/g, ' ').trim();
  if (s.charAt(0) === '=' || s.charAt(0) === '+' || s.charAt(0) === '@') s = "'" + s;
  return s.slice(0, max || 300);
}

/** Аркуш для конкретної дати виїзду: «Рейс 02.11.2026». */
function dateSheet_(dateLabel) {
  var label = clean_(dateLabel, 40);
  if (!label) return null;
  var iso = norm_(label);
  var p = String(iso).split('-');
  var human = p.length === 3 ? p[2] + '.' + p[1] + '.' + p[0] : label;
  var name = ('Рейс ' + human).slice(0, 90);
  return sheet_(name, APP_HEADERS);
}

/* ---------- GET: статуси місць ---------- */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (p.action && p.action !== 'seats') {
      return json_({ success: false, error: 'Невідома операція' });
    }

    var wantRoutes = [norm_(p.route), norm_(p.routeLabel)].filter(String);
    var wantDates = [norm_(p.date), norm_(p.dateLabel)].filter(String);
    if (!wantRoutes.length || !wantDates.length) {
      return json_({ success: false, error: 'Не вказано маршрут або дату' });
    }

    var sh = sheet_(SHEET_SEATS, SEAT_HEADERS);
    var seats = [];
    var last = sh.getLastRow();

    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, 4).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (wantRoutes.indexOf(norm_(r[0])) === -1) continue;
        if (wantDates.indexOf(norm_(r[1])) === -1) continue;
        var n = parseInt(r[2], 10);
        if (!n || n < 1 || n > SEAT_COUNT) continue;
        seats.push({ seat: n, status: normStatus_(r[3]) });
      }
    }

    return json_({
      success: true,
      route: p.route || '',
      date: p.date || '',
      seatCount: SEAT_COUNT,
      seats: seats
    });
  } catch (err) {
    return json_({ success: false, error: 'Не вдалося отримати посадочні місця' });
  }
}

/* ---------- POST: нова заявка ---------- */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (!clean_(d.fullName) || !clean_(d.phone)) {
      return json_({ success: false, error: 'Не вдалося зберегти заявку' });
    }

    // одна людина — одна заявка на дату (з будь-якого браузера чи пристрою)
    var dup = duplicate_(d);
    if (dup) {
      return json_({ success: false, error: 'duplicate', id: dup });
    }

    // місце резервується під блокуванням — дві заявки на одне місце не пройдуть
    var conflict = seatConflicts_(d);
    if (conflict.length) {
      return json_({ success: false, error: 'seat_taken', seats: conflict });
    }

    var sh = sheet_(SHEET_APPS, APP_HEADERS);
    var id = 'APP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd') +
      '-' + String(Math.floor(Math.random() * 9000) + 1000);

    var rec = leadRecord_(d, new Date());

    appendLead_(sh, rec, id);

    try {
      var ds = dateSheet_(d.dateLabel || d.date);
      if (ds) appendLead_(ds, rec, id);
    } catch (e3) {
      // аркуш на дату не критичний — заявка вже в «ЗАЯВКИ»
    }

    markSeats_(d);
    notify_(id, d);
    confirm_(id, d);
    telegram_(id, d);

    return json_({ success: true, id: id });
  } catch (err) {
    return json_({ success: false, error: 'Не вдалося зберегти заявку' });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Лист менеджеру про нову заявку. */
function notify_(id, d) {
  if (!NOTIFY_EMAILS) return;
  try {
    var route = clean_(d.routeLabel || (clean_(d.fromCity) + ' → ' + clean_(d.toCity)), 140);
    var date = clean_(d.dateLabel || d.date, 40);
    var name = clean_(d.fullName, 120);
    var phone = clean_(d.phone, 40);
    var waPhone = phone.replace(/[^0-9]/g, '');

    var rows = [
      ['Номер заявки', id],
      ['Маршрут', route],
      ['Дата поїздки', date],
      ['Ім’я та прізвище', name],
      ['Телефон', phone],
      ['Email', clean_(d.email, 120) || '—'],
      ['Місце', clean_(d.seats, 60) || '—'],
      ['Пасажирів', String(parseInt(d.paxCount, 10) || 1)],
      ['Додатковий багаж', (parseInt(d.extraKg, 10) || 0) + ' кг · ' + (Number(d.extraKgEur) || 0) + ' €'],
      ['Сума', (Number(d.totalEur) || 0) + ' €'],
      ['Канал', clean_(d.channel, 40) || 'Сайт'],
      ['Коментар', clean_(d.comment, 500) || '—'],
      ['Усі пасажири', clean_(d.passengersFlat, 900) || '—']
    ];

    var plain = rows.map(function (r) { return r[0] + ': ' + r[1]; }).join('\n');

    var trs = rows.map(function (r) {
      return '<tr>' +
        '<td style="padding:9px 14px;border-bottom:1px solid #E6EBF2;color:#4B6584;font:600 13px/1.4 Arial,sans-serif;white-space:nowrap">' + r[0] + '</td>' +
        '<td style="padding:9px 14px;border-bottom:1px solid #E6EBF2;color:#08182F;font:700 15px/1.4 Arial,sans-serif">' + r[1] + '</td>' +
        '</tr>';
    }).join('');

    var html =
      '<div style="background:#F5F8FC;padding:26px 14px;font-family:Arial,sans-serif">' +
        '<div style="max-width:600px;margin:0 auto;background:#FFFFFF;border-radius:14px;overflow:hidden;border:1px solid #DEE7F1">' +
          '<div style="background:#061225;padding:20px 24px">' +
            '<div style="color:#F2B01E;font:700 11px/1.4 Arial,sans-serif;letter-spacing:2px;text-transform:uppercase">Avtobus v Italiu</div>' +
            '<div style="color:#FFFFFF;font:700 21px/1.3 Arial,sans-serif;margin-top:7px">Нова заявка</div>' +
            '<div style="color:#B9CBE1;font:600 14px/1.4 Arial,sans-serif;margin-top:5px">' + route + ' · ' + date + '</div>' +
          '</div>' +
          '<table style="width:100%;border-collapse:collapse">' + trs + '</table>' +
          '<div style="padding:20px 24px">' +
            (waPhone
              ? '<a href="https://wa.me/' + waPhone + '" style="display:inline-block;padding:13px 22px;background:#25D366;color:#062B14;border-radius:9px;font:700 15px Arial,sans-serif;text-decoration:none">Написати клієнту в WhatsApp</a>'
              : '') +
            '<a href="tel:' + phone + '" style="display:inline-block;margin-left:8px;padding:13px 22px;background:#F5F8FC;color:#08182F;border:1px solid #CFDBEA;border-radius:9px;font:700 15px Arial,sans-serif;text-decoration:none">Зателефонувати</a>' +
          '</div>' +
        '</div>' +
      '</div>';

    MailApp.sendEmail({
      to: NOTIFY_EMAILS,
      subject: 'Нова заявка: ' + route + ' · ' + date + ' · ' + name,
      body: plain,
      htmlBody: html,
      name: 'Avtobus v Italiu'
    });
  } catch (err) {
    // лист не критичний — заявка вже в таблиці
  }
}

/** Лист-підтвердження клієнту. */
function confirm_(id, d) {
  var to = clean_(d.email, 120);
  if (!to || to.indexOf('@') < 1) return;

  var L = {
    ua: {
      subj: 'Ваша бронь', hi: 'Дякуємо за бронювання!',
      lead: 'Ми отримали вашу заявку. Менеджер зв’яжеться з вами для підтвердження місця.',
      no: 'Номер заявки', route: 'Маршрут', date: 'Дата виїзду', pax: 'Пасажирів',
      seat: 'Місце', name: 'Пасажир', phone: 'Телефон', total: 'До сплати',
      pay: 'Розрахунок готівкою при посадці.',
      bag: '30 кг багажу та 5 кг ручної поклажі входять у квиток. Кожен додатковий кілограм — 1,5 €.',
      ask: 'Маєте питання? Напишіть нам:',
      ua: 'Україна', it: 'Італія',
      foot: 'Посадка та висадка проводяться у призначених місцях за маршрутом. Додаткову посадку чи висадку потрібно узгодити з менеджером.'
    },
    ru: {
      subj: 'Ваша бронь', hi: 'Спасибо за бронирование!',
      lead: 'Мы получили вашу заявку. Менеджер свяжется с вами для подтверждения места.',
      no: 'Номер заявки', route: 'Маршрут', date: 'Дата выезда', pax: 'Пассажиров',
      seat: 'Место', name: 'Пассажир', phone: 'Телефон', total: 'К оплате',
      pay: 'Расчёт наличными при посадке.',
      bag: '30 кг багажа и 5 кг ручной клади входят в билет. Каждый дополнительный килограмм — 1,5 €.',
      ask: 'Есть вопросы? Напишите нам:',
      ua: 'Украина', it: 'Италия',
      foot: 'Посадка и высадка проводятся в назначенных местах по маршруту. Дополнительные посадку или высадку нужно согласовать с менеджером.'
    },
    it: {
      subj: 'La tua prenotazione', hi: 'Grazie per la prenotazione!',
      lead: 'Abbiamo ricevuto la tua richiesta. L’operatore ti contatterà per confermare il posto.',
      no: 'Numero richiesta', route: 'Tratta', date: 'Data di partenza', pax: 'Passeggeri',
      seat: 'Posto', name: 'Passeggero', phone: 'Telefono', total: 'Totale',
      pay: 'Pagamento in contanti alla partenza.',
      bag: '30 kg di bagaglio e 5 kg a bordo sono inclusi. Ogni chilo extra costa 1,5 €.',
      ask: 'Domande? Scrivici:',
      ua: 'Ucraina', it: 'Italia',
      foot: 'La salita e la discesa avvengono nei punti prestabiliti lungo il percorso. Fermate aggiuntive vanno concordate con l’operatore.'
    }
  };
  var t = L[clean_(d.lang, 4)] || L.ua;

  try {
    var route = clean_(d.routeLabel || (clean_(d.fromCity) + ' → ' + clean_(d.toCity)), 140);
    var date = clean_(d.dateLabel || d.date, 40);
    var rows = [
      [t.no, id],
      [t.route, route],
      [t.date, date],
      [t.name, clean_(d.fullName, 120)],
      [t.phone, clean_(d.phone, 40)],
      [t.seat, clean_(d.seats, 60) || '—'],
      [t.pax, String(parseInt(d.paxCount, 10) || 1)],
      [t.total, (Number(d.totalEur) || 0) + ' €']
    ];

    var plain = t.hi + '\n\n' + t.lead + '\n\n' +
      rows.map(function (r) { return r[0] + ': ' + r[1]; }).join('\n') +
      '\n\n' + t.pay + '\n' + t.bag +
      '\n\n' + t.ask + '\nWhatsApp ' + t.ua + ': https://wa.me/' + WA_UA +
      '\nWhatsApp ' + t.it + ': https://wa.me/' + WA_IT +
      '\nViber: +' + WA_UA +
      '\n\n' + t.foot;

    var trs = rows.map(function (r) {
      return '<tr>' +
        '<td style="padding:10px 14px;border-bottom:1px solid #E6EBF2;color:#4B6584;font:600 13px/1.4 Arial,sans-serif;white-space:nowrap">' + r[0] + '</td>' +
        '<td style="padding:10px 14px;border-bottom:1px solid #E6EBF2;color:#08182F;font:700 15px/1.4 Arial,sans-serif;text-align:right">' + r[1] + '</td>' +
        '</tr>';
    }).join('');

    var btn = function (href, bg, fg, label) {
      return '<a href="' + href + '" style="display:inline-block;margin:0 6px 8px 0;padding:13px 20px;background:' + bg +
        ';color:' + fg + ';border-radius:9px;font:700 14px Arial,sans-serif;text-decoration:none">' + label + '</a>';
    };

    var html =
      '<div style="background:#F5F8FC;padding:26px 14px;font-family:Arial,sans-serif">' +
        '<div style="max-width:600px;margin:0 auto;background:#FFFFFF;border-radius:16px;overflow:hidden;border:1px solid #DEE7F1">' +
          '<div style="background:#061225;padding:24px">' +
            '<div style="color:#F2B01E;font:700 11px/1.4 Arial,sans-serif;letter-spacing:2px;text-transform:uppercase">Avtobus v Italiu</div>' +
            '<div style="color:#FFFFFF;font:700 22px/1.3 Arial,sans-serif;margin-top:9px">' + t.hi + '</div>' +
            '<div style="color:#B9CBE1;font:600 14px/1.55 Arial,sans-serif;margin-top:8px">' + t.lead + '</div>' +
          '</div>' +
          '<table style="width:100%;border-collapse:collapse">' + trs + '</table>' +
          '<div style="padding:18px 24px;background:#FBFCFE;border-bottom:1px solid #E6EBF2">' +
            '<div style="color:#08182F;font:700 14px/1.5 Arial,sans-serif">' + t.pay + '</div>' +
            '<div style="color:#4B6584;font:600 13px/1.6 Arial,sans-serif;margin-top:6px">' + t.bag + '</div>' +
          '</div>' +
          '<div style="padding:20px 24px">' +
            '<div style="color:#4B6584;font:600 13px/1.4 Arial,sans-serif;margin-bottom:12px">' + t.ask + '</div>' +
            btn('https://wa.me/' + WA_UA, '#25D366', '#062B14', 'WhatsApp · ' + t.ua) +
            '<div style="height:4px"></div>' +
            btn('https://wa.me/' + WA_IT, '#25D366', '#062B14', 'WhatsApp · ' + t.it) +
          '</div>' +
          '<div style="padding:16px 24px 22px;color:#77899F;font:600 12px/1.6 Arial,sans-serif;border-top:1px solid #E6EBF2">' + t.foot + '</div>' +
        '</div>' +
      '</div>';

    MailApp.sendEmail({
      to: to,
      subject: t.subj + ' № ' + id + ' — ' + date,
      replyTo: 'Oktobus109@gmail.com',
      body: plain,
      htmlBody: html,
      name: 'Avtobus v Italiu'
    });
  } catch (err) {
    // лист клієнту не критичний
  }
}

/** Текст для Telegram без формул-захисту таблиці (без апострофа перед «+»). */
function plain_(v, max) {
  var s = String(v === null || v === undefined ? '' : v);
  return s.replace(/[\u0000-\u001F\u007F]/g, ' ').trim().slice(0, max || 300);
}

/** 2026-09-08, 08.09.2026, 8/9/2026 або Date → 08.09.2026. */
function dmy_(v) {
  var iso = norm_(v);
  var m = String(iso).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? m[3] + '.' + m[2] + '.' + m[1] : plain_(v, 40);
}

/**
 * Коротке повідомлення без emoji:
 *   08.09.2026
 *   Львів → Bologna
 *   Тест Тестенко
 *   +380674704617
 *   Місце: 14
 *   Oktobus109@gmail.com   (рядок пропускається, якщо email не вказано)
 *   Сума: 130 €
 * Сума береться з уже розрахованого totalEur, тут нічого не перераховується.
 */
function telegramText_(d) {
  var from = plain_(d.fromCity, 80), to = plain_(d.toCity, 80);
  var route = from && to ? from + ' → ' + to : plain_(d.routeLabel, 140);
  var lines = [
    dmy_(d.dateLabel || d.date),
    route,
    plain_(d.fullName, 120),
    plain_(d.phone, 40),
    'Місце: ' + (plain_(d.seats, 60) || '—')
  ];
  var email = plain_(d.email, 120);
  if (email) lines.push(email);
  lines.push('Сума: ' + (Number(d.totalEur) || 0) + ' €');
  return lines.join('\n');
}

/** Повідомлення в Telegram. */
function telegram_(id, d) {
  if (!TG_TOKEN || !TG_CHAT) return;
  try {
    var waPhone = plain_(d.phone, 40).replace(/[^0-9]/g, '');

    var payload = {
      chat_id: TG_CHAT,
      text: telegramText_(d),
      disable_web_page_preview: true
    };

    if (waPhone) {
      payload.reply_markup = JSON.stringify({
        inline_keyboard: [[
          { text: 'WhatsApp', url: 'https://wa.me/' + waPhone }
        ]]
      });
    }

    var send = function (chat) {
      payload.chat_id = chat;
      var r = UrlFetchApp.fetch('https://api.telegram.org/bot' + TG_TOKEN + '/sendMessage', {
        method: 'post',
        payload: payload,
        muteHttpExceptions: true
      });
      return JSON.parse(r.getContentText() || '{}');
    };
    var res = send(TG_CHAT);
    // група стала супергрупою — Telegram дає новий ID, пробуємо його
    if (!res.ok && res.parameters && res.parameters.migrate_to_chat_id) {
      Logger.log('Новий ID чату: ' + res.parameters.migrate_to_chat_id + ' — впишіть його в TG_CHAT');
      res = send(String(res.parameters.migrate_to_chat_id));
    }
    Logger.log(JSON.stringify(res));
  } catch (err) {
    Logger.log('Telegram error: ' + err);
  }
}

/** Останні 9 цифр телефону: +380 67…, 067… і 38067… — один номер. */
function phoneKey_(v) {
  var s = String(v === null || v === undefined ? '' : v).replace(/[^0-9]/g, '');
  return s.length > 9 ? s.slice(-9) : s;
}

/** ID наявної заявки з тим самим телефоном на ту саму дату, або ''. */
function duplicate_(d) {
  var key = phoneKey_(d.phone);
  if (key.length < 7) return '';
  var wantDates = [norm_(d.date), norm_(d.dateLabel)].filter(String);
  if (!wantDates.length) return '';

  var sh = sheet_(SHEET_APPS, APP_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return '';
  var idx = headerIndex_(sh);
  if (!idx[COL_PHONE] || !idx[COL_TRIP_DATE]) return '';
  var n = last - 1;
  var phones = sh.getRange(2, idx[COL_PHONE], n, 1).getValues();
  var dates = sh.getRange(2, idx[COL_TRIP_DATE], n, 1).getValues();
  var notes = idx[COL_CREATED] ? sh.getRange(2, idx[COL_CREATED], n, 1).getNotes() : [];
  var ids = idx[LEGACY_ID] ? sh.getRange(2, idx[LEGACY_ID], n, 1).getValues() : [];

  for (var i = n - 1; i >= 0; i--) {
    if (phoneKey_(phones[i][0]) !== key) continue;
    if (wantDates.indexOf(norm_(dates[i][0])) === -1) continue;
    var id = String((notes[i] && notes[i][0]) || (ids[i] && ids[i][0]) || '').trim();
    return id || '—';
  }
  return '';
}

/** Номери місць із заявки, які вже ЗАЙНЯТЕ або ЗАБЛОКОВАНЕ. */
function seatConflicts_(d) {
  var want = String(d.seats || '').split(',')
    .map(function (x) { return parseInt(x, 10); })
    .filter(function (n) { return n >= 1 && n <= SEAT_COUNT; });
  if (!want.length) return [];

  var sh = sheet_(SHEET_SEATS, SEAT_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 4).getValues();

  var wantRoutes = [norm_(d.route), norm_(d.routeLabel)].filter(String);
  var wantDates = [norm_(d.date), norm_(d.dateLabel)].filter(String);
  var out = [];

  for (var i = 0; i < rows.length; i++) {
    if (wantRoutes.indexOf(norm_(rows[i][0])) === -1) continue;
    if (wantDates.indexOf(norm_(rows[i][1])) === -1) continue;
    var n = parseInt(rows[i][2], 10);
    if (want.indexOf(n) === -1) continue;
    if (normStatus_(rows[i][3]) !== 'free' && out.indexOf(n) === -1) out.push(n);
  }
  return out;
}

/** Позначає обрані місця як ЗАЙНЯТЕ. */
function markSeats_(d) {
  var seats = String(d.seats || '').split(',')
    .map(function (x) { return parseInt(x, 10); })
    .filter(function (n) { return n >= 1 && n <= SEAT_COUNT; });
  if (!seats.length) return;

  var route = clean_(d.routeLabel || d.route, 120);
  var dateLabel = clean_(d.dateLabel || d.date, 40);
  var sh = sheet_(SHEET_SEATS, SEAT_HEADERS);
  var last = sh.getLastRow();
  var rows = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];

  var wantRoutes = [norm_(d.route), norm_(d.routeLabel)].filter(String);
  var wantDates = [norm_(d.date), norm_(d.dateLabel)].filter(String);

  seats.forEach(function (n) {
    var found = -1;
    for (var i = 0; i < rows.length; i++) {
      if (wantRoutes.indexOf(norm_(rows[i][0])) === -1) continue;
      if (wantDates.indexOf(norm_(rows[i][1])) === -1) continue;
      if (parseInt(rows[i][2], 10) !== n) continue;
      found = i + 2;
      break;
    }
    if (found > 0) {
      sh.getRange(found, 4).setValue(STATUS_TAKEN);
    } else {
      sh.appendRow([route, dateLabel, n, STATUS_TAKEN]);
    }
  });
}

/* ---------- меню в таблиці ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Рейси')
    .addItem('Створити рейс (50 вільних місць)', 'createFlight')
    .addItem('Створити аркуші', 'setupSheets')
    .addToUi();
}

function createFlight() {
  var ui = SpreadsheetApp.getUi();

  var r1 = ui.prompt('Новий рейс', 'Маршрут (наприклад: Львів → Bologna)', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var route = clean_(r1.getResponseText(), 120);
  if (!route) return ui.alert('Маршрут не вказано.');

  var r2 = ui.prompt('Новий рейс', 'Дата рейсу (наприклад: 08.09.2026)', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var date = clean_(r2.getResponseText(), 40);
  if (!date) return ui.alert('Дату не вказано.');

  var sh = sheet_(SHEET_SEATS, SEAT_HEADERS);
  var rows = [];
  for (var n = 1; n <= SEAT_COUNT; n++) rows.push([route, date, n, STATUS_FREE]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);

  ui.alert('Готово: ' + route + ' · ' + date + ' — ' + SEAT_COUNT + ' вільних місць.');
}

function setupSheets() {
  sheet_(SHEET_APPS, APP_HEADERS);
  sheet_(SHEET_SEATS, SEAT_HEADERS);
}

/* ---------- перехід на нову схему заявок (запускати вручну, один раз) ---------- */

/** Аркуші із заявками: «ЗАЯВКИ» і всі «Рейс …». */
function leadSheets_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(function (sh) {
    var name = sh.getName();
    return name === SHEET_APPS || name.indexOf('Рейс ') === 0;
  });
}

/**
 * Для кожного аркуша заявок:
 *  1) номер із колонки ID переносить у примітку клітинки «Дата заявки» і видаляє колонку ID;
 *  2) переставляє колонки в порядку APP_HEADERS разом із даними
 *     («Дата поїздки» стає першою, «Дата заявки» — на місці колишньої «Дата поїздки»).
 * Колонки, яких немає в схемі (власні колонки менеджера), залишаються праворуч.
 * dryRun = true — лише пише в журнал, що буде зроблено.
 */
function migrateLeadSheet_(sh, dryRun) {
  var log = [];
  if (!sh.getLastColumn() || !sh.getLastRow()) return ['порожній аркуш — пропущено'];
  var idx = headerIndex_(sh);
  var last = sh.getLastRow();

  if (idx[LEGACY_ID]) {
    if (!idx[COL_CREATED]) throw new Error(sh.getName() + ': немає колонки «' + COL_CREATED + '»');
    log.push('перенести ID у примітки «' + COL_CREATED + '» і видалити колонку ID (' + idx[LEGACY_ID] + ')');
    if (!dryRun) {
      if (last > 1) {
        var ids = sh.getRange(2, idx[LEGACY_ID], last - 1, 1).getValues();
        var target = sh.getRange(2, idx[COL_CREATED], last - 1, 1);
        var notes = target.getNotes();
        var merged = ids.map(function (r, i) { return [String(r[0] || '').trim() || notes[i][0]]; });
        target.setNotes(merged);
      }
      sh.deleteColumn(idx[LEGACY_ID]);
    }
  }

  var headers = APP_HEADERS.slice();
  if (dryRun) {
    var now = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(function (h) { return String(h).trim(); })
      .filter(function (h) { return h && h !== LEGACY_ID; });
    log.push('порядок зараз: ' + now.join(' | '));
    log.push('порядок після: ' + headers.filter(function (h) { return now.indexOf(h) !== -1; }).join(' | '));
    return log;
  }

  var pos = 1;
  headers.forEach(function (h) {
    var cur = headerIndex_(sh)[h];
    if (!cur) return;
    if (cur !== pos) sh.moveColumns(sh.getRange(1, cur, sh.getMaxRows(), 1), pos);
    pos++;
  });
  log.push('колонки впорядковано за схемою');
  return log;
}

/** Показує в журналі, що зміниться. Нічого не змінює. */
function migrateLeadSheetsDryRun() {
  leadSheets_().forEach(function (sh) {
    Logger.log(sh.getName() + ':\n  ' + migrateLeadSheet_(sh, true).join('\n  '));
  });
}

/** Переводить аркуші заявок на нову схему. Спочатку — на копії таблиці. */
function migrateLeadSheets() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    leadSheets_().forEach(function (sh) {
      Logger.log(sh.getName() + ': ' + migrateLeadSheet_(sh, false).join('; '));
    });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- перевірка ---------- */

var TEST_DATA = {
  fullName: 'Тест Тестенко',
  phone: '+380674704617',
  email: 'Oktobus109@gmail.com',
  fromCity: 'Львів',
  toCity: 'Bologna',
  routeLabel: 'Львів → Bologna',
  dateLabel: '08.09.2026',
  date: '2026-09-08',
  seats: '14',
  paxCount: 1,
  comment: 'Тестова заявка',
  channel: 'Тест',
  lang: 'ua',
  extraKg: 0,
  extraKgEur: 0,
  totalEur: 130
};

/** Тестове повідомлення в Telegram. */
function testTelegram() {
  telegram_('TEST-0001', TEST_DATA);
}

/** Тестовий лист менеджеру. */
function testEmail() {
  notify_('TEST-0001', TEST_DATA);
}

/** Тестовий лист клієнту — прийде на пошту з TEST_DATA.email. */
function testClientEmail() {
  confirm_('TEST-0001', TEST_DATA);
  Logger.log('Залишок листів на сьогодні: ' + MailApp.getRemainingDailyQuota());
}