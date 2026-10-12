import assert from 'node:assert/strict';

// Fares as they were before Fano/Marotta/Senigallia/Civitanova were added.
// Columns: base (Trieste…Rimini), pesaro, ancona, pescara.
const FARES = {
  kryvyi: [150, 160, 170, 170], kropyv: [150, 160, 170, 170], cherkasy: [150, 160, 160, 160],
  kyiv: [140, 150, 150, 150], zhytomyr: [140, 150, 150, 150], rivne: [130, 140, 140, 140],
  lviv: [120, 120, 120, 120], stryi: [120, 120, 120, 120], uzhhorod: [120, 120, 120, 120]
};
const TIER = {
  trieste: 0, venezia: 0, padova: 0, rovigo: 0, ferrara: 0, bologna: 0, imola: 0, forli: 0, rimini: 0,
  pesaro: 1, ancona: 2, pescara: 3
};
const ANCONA_ALIASES = ['fano', 'marotta', 'senigallia', 'civitanova'];

const UA_POINTS = {
  kryvyi: '95 квартал', kropyv: 'автостанція', cherkasy: 'автостанція 2',
  kyiv: 'вул. Симона Петлюри, 32 / Центральний залізничний вокзал',
  zhytomyr: 'автостанція', rivne: 'автостанція', lviv: 'Центральний автовокзал'
};
const IT_POINTS = {
  trieste: 'Piazza della Libertà, Via Flavio Gioia 2', venezia: 'Viale Stazione 3',
  padova: 'Viale della Pace, autostazione, corsia 12', rovigo: 'Piazza Riconoscenza 5',
  ferrara: 'Piazza Stazione', bologna: 'Piazza XX Settembre 6', imola: 'Via Romagnoli',
  forli: 'Via Alessandro Volta, autostazione', rimini: 'Piazzale Cesare Battisti 22',
  pesaro: 'Piazza Marcello Stefanini, Palasport',
  fano: 'Casello A14 Fano, parcheggio',
  marotta: 'Точну точку повідомить менеджер',
  senigallia: 'Viale Giordano Bruno, distributore Enilive',
  ancona: 'Piazza Rosselli, stazione F.S.',
  civitanova: 'Via Luigi Einaudi, centro commerciale Megauno',
  pescara: 'Piazzale Terminal Bus, fronte stazione FS'
};
const IT_ORDER = ['trieste', 'venezia', 'padova', 'rovigo', 'ferrara', 'bologna', 'imola', 'forli', 'rimini',
  'pesaro', 'fano', 'marotta', 'senigallia', 'ancona', 'civitanova', 'pescara'];

export async function testRoutes(browser, base, pass) {
  const { Component } = await import('../src/booking.js');
  const model = new Component({ defaultLang: 'ua', initialNow: '2026-10-12T09:00:00Z' });
  const setState = patch => { model.state = { ...model.state, ...patch }; };

  // Pricing: existing fares unchanged; aliases use the Ancona column of the same UA city.
  for (const [ua, fares] of Object.entries(FARES)) {
    for (const [it, tier] of Object.entries(TIER)) assert.equal(model.price(ua, it), fares[tier], ua + '→' + it);
    for (const it of ANCONA_ALIASES) assert.equal(model.price(ua, it), model.price(ua, 'ancona'), ua + '→' + it);
  }
  assert.equal(model.price('lviv', 'bologna'), 120);
  setState({ dir: 'ua_it', from: 'kryvyi', to: 'fano', pax: '2', extraKg: 0 });
  assert.equal(model.total(), 340);

  for (const lang of ['ua', 'ru', 'it']) {
    setState({ lang });
    for (const [k, text] of Object.entries(IT_POINTS)) {
      setState({ dir: 'ua_it', from: 'lviv', to: k });
      const expected = k === 'marotta' ? model.t().ui.addressPending : text;
      assert.equal(model.routeNames().tp, expected, lang + ' ' + k);
    }
  }
  setState({ lang: 'ua' });
  for (const [k, text] of Object.entries(UA_POINTS)) {
    setState({ dir: 'ua_it', from: k, to: 'bologna' });
    assert.equal(model.routeNames().fp, text, k);
  }
  for (const k of ['stryi', 'uzhhorod']) {
    setState({ dir: 'ua_it', from: k, to: 'bologna' });
    assert.equal(model.routeNames().fp, 'Точну точку повідомить менеджер', k);
  }
  setState({ lang: 'ru', dir: 'ua_it', from: 'kyiv', to: 'bologna' });
  assert.equal(model.routeNames().fp, 'ул. Симона Петлюры, 32 / Центральный железнодорожный вокзал');
  setState({ lang: 'ua' });

  // Destination stays the real city in the payload and message; fare is Ancona's.
  for (const k of ANCONA_ALIASES) {
    setState({ dir: 'ua_it', from: 'kryvyi', to: k, pax: '1', extraKg: 0, dateOut: '2026-10-20', seats: [],
      lead: { first: 'Тест', last: 'Тестенко', phone: '0674704617' }, extra: [], email: '', comment: '' });
    const p = model.payload('site');
    const name = { fano: 'Fano', marotta: 'Marotta', senigallia: 'Senigallia', civitanova: 'Civitanova' }[k];
    assert.equal(p.toCity, name);
    assert(p.routeLabel.includes(name), p.routeLabel);
    assert(!p.routeLabel.includes('Ancona'), p.routeLabel);
    assert(p.message.includes('Висадка: ' + name + ' — '), p.message);
    assert.equal(p.totalEur, 170);
    assert.equal(p.dropoffAddress, model.routeNames().tp);
  }

  // Unconfirmed departures: not offered as Italian boarding cities, no crash without a schedule.
  setState({ dir: 'ua_it', from: 'lviv', to: 'fano' });
  assert.equal(model.timeLabel(model.legInfo(true, 'lviv', 'fano').arr), model.t().durationNA);
  setState({ dir: 'it_ua', from: 'fano', to: 'lviv' });
  assert.equal(model.boardingOffset('it_ua', 'fano'), 0);
  pass('fares unchanged, Fano/Marotta/Senigallia/Civitanova priced as Ancona with own names, boarding points in UA/RU/IT');

  // Browser: route selects, address block with map link, real destination in the POST.
  const context = await browser.newContext({ viewport: { width: 390, height: 844 }, isMobile: true, hasTouch: true });
  const page = await context.newPage();
  let payload;
  await page.route('https://script.google.com/**', route => {
    if (route.request().method() === 'POST') {
      payload = route.request().postDataJSON();
      return route.fulfill({ json: { success: true, id: 'ROUTE-TEST' } });
    }
    return route.fulfill({ json: { success: true, seats: [] } });
  });
  await page.goto(base + '/', { waitUntil: 'networkidle' });
  const select = page.locator('#v3search select');
  const values = async n => (await select.nth(n).locator('option').evaluateAll(o => o.map(x => x.value))).filter(Boolean);
  assert.deepEqual(await values(1), IT_ORDER);
  assert.equal(new Set(await values(1)).size, IT_ORDER.length);
  await select.nth(0).selectOption('kryvyi');
  await select.nth(1).selectOption('fano');
  await select.nth(2).selectOption({ index: 1 });
  await page.locator('.trip-details > summary').first().click();
  const details = page.locator('.trip-details').first();
  const text = await details.innerText();
  assert(text.includes('Кривий Ріг — 95 квартал'), text);
  assert(text.includes('Fano — Casello A14 Fano, parcheggio'), text);
  assert.equal(await details.locator('a[href*="maps.app.goo.gl"]').count(), 0);
  await page.locator('#passenger-1-first').fill('Тест');
  await page.locator('#passenger-1-last').fill('Тестенко');
  await page.locator('#passenger-1-phone').fill('0674704617');
  await page.locator('input[type=checkbox]').check();
  await page.getByRole('button', { name: 'Надіслати заявку', exact: true }).first().click();
  await page.getByRole('dialog').waitFor();
  assert.equal(payload.toCity, 'Fano');
  assert.equal(payload.fromCity, 'Кривий Ріг');
  assert.equal(payload.totalEur, 170);
  assert.equal(payload.boardingAddress, '95 квартал');
  assert.equal(payload.dropoffAddress, 'Casello A14 Fano, parcheggio');

  await page.keyboard.press('Escape');
  await page.locator('[data-v=direction]').nth(1).click();
  const itFrom = await values(0);
  for (const k of ANCONA_ALIASES) assert(!itFrom.includes(k), 'unscheduled departure offered: ' + k);
  assert(itFrom.includes('ancona') && itFrom.includes('pescara'));
  await context.close();
  pass('route select order without duplicates, written address for Fano, real destination and Ancona fare in the mock POST');
}
