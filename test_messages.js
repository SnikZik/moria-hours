// Simulates pickMessage_ from dist/Code.gs against an in-memory log tab, no Google needed.
const fs = require('fs');
const vm = require('vm');

function fakeSheet() {
  const rows = [['תאריך ושעה', 'פעולה', 'משפט']];
  const sheet = {
    rows,
    getLastRow: () => rows.length,
    getRange: (r, c, nr, nc) => ({
      getValue: () => (rows[r - 1] || [])[c - 1],
      setValue() { rows[r - 1][c - 1] = 'משפט'; return this; },
      setFontWeight() { return this; },
      getValues: () => rows.slice(r - 1, r - 1 + nr).map((row) => row.slice(c - 1, c - 1 + nc)),
    }),
  };
  return sheet;
}

const ctx = {
  console, Math, Date,
  Utilities: { formatDate: (d) => d.toISOString().slice(0, 10) },
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(process.argv[2], 'utf8'), ctx);

const ACTIONS = ['start', 'break_start', 'break_end', 'end'];

function runMonth(month, days, log) {
  const ss = { getSheetByName: () => log };
  const shown = [];
  let blanks = 0;
  for (let day = 1; day <= days; day++) {
    const now = new Date(`${month}-${String(day).padStart(2, '0')}T12:00:00Z`);
    for (const action of ACTIONS) {
      const m = ctx.pickMessage_(ss, action, now);
      log.rows.push([now, action, m]);
      if (m) shown.push(m); else blanks++;
    }
  }
  return { shown, blanks };
}

const log = fakeSheet();
const a = runMonth('2026-12', 31, log);
const repeatsA = a.shown.length - new Set(a.shown).size;
console.log(`every day of a 31 day month: shown=${a.shown.length} repeats=${repeatsA} blanks=${a.blanks}`);

const b = runMonth('2027-01', 2, log);
const carried = b.shown.filter((m) => a.shown.includes(m)).length;
console.log(`next month: shown=${b.shown.length} blanks=${b.blanks} lines available again=${carried}`);

const extra = fakeSheet();
const ss = { getSheetByName: () => extra };
const now = new Date('2026-12-01T12:00:00Z');
const startTotal = ctx.MESSAGES.start.length;
const outs = [];
for (let i = 0; i < startTotal + 3; i++) {
  const m = ctx.pickMessage_(ss, 'start', now);
  extra.rows.push([now, 'start', m]);
  outs.push(m);
}
const nonBlank = outs.filter(Boolean);
console.log(`exhausting one list: ${startTotal} lines, ${outs.length} presses, repeats=${nonBlank.length - new Set(nonBlank).size} blanks=${outs.length - nonBlank.length}`);

if (repeatsA || a.blanks || nonBlank.length !== new Set(nonBlank).size || !b.shown.length) process.exit(1);
console.log('PASS');
