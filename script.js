/* Fire Timesheet — replicates the FIRE.xlsx workbook.
   Blue cells in the workbook are the <input> fields here; everything else is calculated. */
(function () {
  'use strict';

  // ---- Configuration (constants that were hard-coded in the workbook formulas) ----
  const REGULAR_BASE = 112;      // REGULAR = 112 - LEAVE
  const FLSA_THRESHOLD = 182;    // OT 1/2 = FLSA regular hours - 182
  const STORAGE_KEY = 'fire-timesheet-v1';
  const DATES_STORAGE_KEY = 'fire-timesheet-dates-v1';

  const DEFAULT_PERIODS = [
    { start: '2026-08-15', end: '2026-08-28' },
    { start: '2026-08-29', end: '2026-09-11' },
    { start: '2026-09-12', end: '2026-09-25', usesFlsa: true }
  ];
  // Workbook: FLSA 08-19 .. 09-12, with the 24 listed days starting the day after the first date.
  const DEFAULT_FLSA = { start: '2026-08-19', end: '2026-09-12' };

  // Input columns (workbook columns C..I)
  const COLS = [
    { key: 'reg',   label: 'Regular',          code: '700100', group: 'Hours worked' },
    { key: 'ot15',  label: 'O/T @ 1.5',        code: '700300', group: 'Hours worked' },
    { key: 'othalf',label: 'O/T',              code: '700300', group: '½ Rate Hrs.' },
    { key: 'hol',   label: 'Hol',              code: '700500', group: '½ Rate Hrs.' },
    { key: 'fire',  label: 'Fire Engineer',    code: '700500' },
    { key: 'moveup',label: '* Hours Move-up',  code: '700900' },
    { key: 'leave', label: '* Hours Leave',    code: '700100' }
  ];
  const IDX = Object.fromEntries(COLS.map((c, i) => [c.key, i]));

  // ---- Date helpers (UTC to avoid timezone drift) ----
  const DAY = 86400000;
  const parse = iso => { const [y, m, d] = iso.split('-').map(Number); return Date.UTC(y, m - 1, d); };
  const toIso = ms => new Date(ms).toISOString().slice(0, 10);
  const WD = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const fmtShort = iso => { const d = new Date(parse(iso)); return `${String(d.getUTCMonth() + 1).padStart(2, '0')}/${String(d.getUTCDate()).padStart(2, '0')}/${d.getUTCFullYear()}`; };
  const fmtDay = iso => `${WD[new Date(parse(iso)).getUTCDay()]} ${fmtShort(iso).slice(0, 5)}`;
  const range = (a, b) => { const out = []; for (let t = parse(a); t <= parse(b); t += DAY) out.push(toIso(t)); return out; };

  function loadDates() {
    try {
      const dates = JSON.parse(localStorage.getItem(DATES_STORAGE_KEY));
      if (Array.isArray(dates.periods) && dates.periods.length === DEFAULT_PERIODS.length &&
          dates.periods.every(p => isValidDate(p.start) && isValidDate(p.end) && p.start <= p.end) &&
          dates.flsa && isValidDate(dates.flsa.start) && isValidDate(dates.flsa.end) &&
          dates.flsa.start < dates.flsa.end) {
        return dates;
      }
    } catch (e) { /* storage unavailable or invalid — use defaults */ }
    return { periods: DEFAULT_PERIODS.map(({ start, end }) => ({ start, end })), flsa: { ...DEFAULT_FLSA } };
  }
  function isValidDate(iso) {
    if (typeof iso !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return false;
    const ms = parse(iso);
    return Number.isFinite(ms) && toIso(ms) === iso;
  }

  const dates = loadDates();
  let PERIODS = dates.periods.map((p, i) => ({ ...p, usesFlsa: i === 2, days: [] }));
  let FLSA = { ...dates.flsa };
  let FLSA_DAYS = [];

  function updateDays() {
    PERIODS.forEach(p => { p.days = range(p.start, p.end); });
    FLSA_DAYS = range(toIso(parse(FLSA.start) + DAY), FLSA.end);
  }
  updateDays();

  // ---- State: { 'YYYY-MM-DD': [7 numbers|null] } ----
  let state = load();

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return JSON.parse(raw);
    } catch (e) { /* storage unavailable — fall through */ }
    return clone(window.SAMPLE_DATA || {});
  }
  function save() { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch (e) { /* ignore */ } }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  const num = v => (v === null || v === undefined || v === '' || isNaN(+v)) ? 0 : +v;
  const get = (iso, i) => num(state[iso] && state[iso][i]);
  const round = n => Math.round(n * 100) / 100;
  const fmt = n => String(round(n));

  // ---- Calculations (mirror workbook formulas) ----
  function rowTotal(iso) { return COLS.reduce((s, _, i) => s + get(iso, i), 0); }          // =SUM(C:I)
  function colTotal(p, i) { return p.days.reduce((s, d) => s + get(d, i), 0); }              // =SUM(col)
  function flsaTotal() { return FLSA_DAYS.reduce((s, d) => s + get(d, IDX.reg), 0); }       // =SUM(VLOOKUP(date,B:D,2))
  function flsaOT() { return Math.max(0, flsaTotal() - FLSA_THRESHOLD); }                    // =R3-182 (floored at 0)

  function periodTotals(p) {
    const t = COLS.map((_, i) => colTotal(p, i));
    if (p.usesFlsa) t[IDX.othalf] = flsaOT();                                                // E58 = N49
    const grand = p.days.reduce((s, d) => s + rowTotal(d), 0);                               // J58 = SUM(J rows)
    const leave = t[IDX.leave];
    return {
      cols: t,
      grand,
      summary: [
        ['REGULAR', REGULAR_BASE - leave],
        ['OT', t[IDX.ot15]],
        ['LEAVE', leave],
        ['SPDTYP', t[IDX.fire]],
        ['MOVEUP', t[IDX.moveup]],
        ...(p.usesFlsa ? [['OT 1/2', flsaOT()]] : [])
      ]
    };
  }

  // ---- Rendering ----
  const $ = (sel, root = document) => root.querySelector(sel);
  const el = (tag, attrs = {}, html) => {
    const e = document.createElement(tag);
    for (const k in attrs) e.setAttribute(k, attrs[k]);
    if (html !== undefined) e.innerHTML = html;
    return e;
  };

  function buildPeriod(p, n) {
    const sec = el('section', { class: 'card', 'aria-labelledby': `pp-${n}` });
    sec.dataset.period = n;

    sec.appendChild(el('div', { class: 'card-head' },
      `<h2 id="pp-${n}">Pay period ${n + 1}</h2>
       <p class="meta">${fmtShort(p.start)} – ${fmtShort(p.end)} · FLSA ${fmtShort(FLSA.start)} – ${fmtShort(FLSA.end)}</p>`));

    // group header spans (merged cells in workbook: "Hours worked" over 2 cols, "1/2 Rate Hrs." over 2 cols)
    const groups = [['Hours worked', 2], ['½ Rate Hrs.', 2], ['', 3]];
    const thead = `
      <tr class="grp"><th></th>${groups.map(([t, s]) => `<th colspan="${s}">${t}</th>`).join('')}<th></th></tr>
      <tr><th scope="col" class="day">Day</th>${COLS.map(c =>
        `<th scope="col">${c.label}<span class="code">${c.code}</span></th>`).join('')}<th scope="col">Total Hours</th></tr>`;

    const rows = p.days.map(iso => `
      <tr data-date="${iso}">
        <th scope="row" class="day">${fmtDay(iso)}<span class="yr">${iso.slice(0, 4)}</span></th>
        ${COLS.map((c, i) => `<td class="in"><input type="number" inputmode="decimal" min="0" step="any"
            data-date="${iso}" data-col="${i}" aria-label="${fmtDay(iso)} ${c.label} ${c.code}"></td>`).join('')}
        <td class="calc row-total">0</td>
      </tr>`).join('');

    const tfoot = `<tr class="tot"><th scope="row" class="day">Total</th>${COLS.map((c, i) =>
        `<td class="col-total" data-col="${i}"${p.usesFlsa && i === IDX.othalf ? ' title="Calculated from the FLSA work period"' : ''}>0</td>`).join('')}
        <td class="grand">0</td></tr>`;

    const wrap = el('div', { class: 'table-wrap' },
      `<table><thead>${thead}</thead><tbody>${rows}</tbody><tfoot>${tfoot}</tfoot></table>`);
    sec.appendChild(wrap);

    sec.appendChild(el('dl', { class: 'summary' }));
    if (p.usesFlsa) {
      sec.appendChild(el('p', { class: 'note' }, 'O/T ½ total is calculated from the FLSA work period below, not from the column inputs.'));
    }
    return sec;
  }

  function buildAll() {
    const host = $('#periods');
    host.replaceChildren();
    PERIODS.forEach((p, n) => host.appendChild(buildPeriod(p, n)));

    // FLSA chips
    $('#flsa-meta').textContent = `${fmtShort(FLSA_DAYS[0])} – ${fmtShort(FLSA.end)} · ${FLSA_DAYS.length} days · Regular hours only`;
    $('#flsa-grid').innerHTML = FLSA_DAYS.map((iso, i) =>
      `<div class="chip" data-flsa="${iso}"><span class="n">${i + 1}</span><span class="d">${fmtDay(iso)}</span><span class="h">0</span></div>`).join('');
    $('#flsa-threshold').textContent = FLSA_THRESHOLD;

    fillInputs();
  }

  function fillInputs() {
    document.querySelectorAll('input[data-date]').forEach(inp => {
      const v = state[inp.dataset.date] && state[inp.dataset.date][+inp.dataset.col];
      inp.value = (v === null || v === undefined) ? '' : v;
    });
  }

  function update() {
    PERIODS.forEach((p, n) => {
      const sec = $(`section[data-period="${n}"]`);
      const t = periodTotals(p);
      p.days.forEach(iso => {
        $(`tr[data-date="${iso}"] .row-total`, sec).textContent = fmt(rowTotal(iso));
      });
      sec.querySelectorAll('.col-total').forEach(td => { td.textContent = fmt(t.cols[+td.dataset.col]); });
      $('.grand', sec).textContent = fmt(t.grand);
      $('.summary', sec).innerHTML = t.summary.map(([k, v]) =>
        `<div><dt>${k}</dt><dd>${fmt(v)}</dd></div>`).join('');
    });
    FLSA_DAYS.forEach(iso => {
      const chip = $(`.chip[data-flsa="${iso}"]`);
      const h = get(iso, IDX.reg);
      $('.h', chip).textContent = fmt(h);
      chip.classList.toggle('zero', h === 0);
    });
    $('#flsa-total').textContent = fmt(flsaTotal());
    $('#flsa-ot').textContent = fmt(flsaOT());
  }

  // ---- Events ----
  document.addEventListener('input', e => {
    const inp = e.target;
    if (!inp.matches || !inp.matches('input[data-date]')) return;
    const iso = inp.dataset.date, i = +inp.dataset.col;
    if (!state[iso]) state[iso] = new Array(COLS.length).fill(null);
    state[iso][i] = inp.value === '' ? null : parseFloat(inp.value);
    save(); update();
  });

  $('#btn-sample').addEventListener('click', () => { state = clone(window.SAMPLE_DATA || {}); save(); fillInputs(); update(); });
  $('#btn-clear').addEventListener('click', () => {
    if (!confirm('Clear all entered hours?')) return;
    state = {}; save(); fillInputs(); update();
  });
  $('#btn-print').addEventListener('click', () => window.print());

  const datesForm = $('#dates-form');
  const dateInput = id => $(`#${id}`);
  function fillDateInputs() {
    PERIODS.forEach((p, i) => {
      dateInput(`period-${i}-start`).value = p.start;
      dateInput(`period-${i}-end`).value = p.end;
    });
    dateInput('flsa-start').value = FLSA.start;
    dateInput('flsa-end').value = FLSA.end;
  }

  datesForm.addEventListener('submit', e => {
    e.preventDefault();
    const nextPeriods = PERIODS.map((p, i) => ({
      start: dateInput(`period-${i}-start`).value,
      end: dateInput(`period-${i}-end`).value,
      usesFlsa: p.usesFlsa
    }));
    const nextFlsa = { start: dateInput('flsa-start').value, end: dateInput('flsa-end').value };
    const error = $('#dates-error');

    if (nextPeriods.some(p => !isValidDate(p.start) || !isValidDate(p.end) || p.start > p.end)) {
      error.textContent = 'Each pay period needs a valid start date on or before its end date.';
      return;
    }
    if (nextPeriods.some((p, i) => i > 0 && p.start <= nextPeriods[i - 1].end)) {
      error.textContent = 'Pay periods must be in order and must not overlap.';
      return;
    }
    if (!isValidDate(nextFlsa.start) || !isValidDate(nextFlsa.end) || nextFlsa.start >= nextFlsa.end) {
      error.textContent = 'The FLSA start date must be before its end date.';
      return;
    }

    PERIODS = nextPeriods.map(p => ({ ...p, days: [] }));
    FLSA = nextFlsa;
    updateDays();
    try {
      localStorage.setItem(DATES_STORAGE_KEY, JSON.stringify({
        periods: PERIODS.map(({ start, end }) => ({ start, end })),
        flsa: FLSA
      }));
      error.textContent = '';
    } catch (e) {
      error.textContent = 'Dates were applied, but could not be saved in this browser.';
    }
    buildAll();
    update();
  });

  fillDateInputs();
  buildAll();
  update();

  // Expose for testing
  window.__fire = { periodTotals, flsaTotal, flsaOT, get PERIODS() { return PERIODS; } };
})();
