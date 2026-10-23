// Run by tests/timezone.test.mjs in a child process with TZ set: prints what the rules make of yesterday's and today's
// visits at several local times on Fri 10/23/2026. Synthetic patients only (made-up MRNs 9006xx).
'use strict';
const R = require('../../rules.js');

const visit = (date) => ({ date, time: '9:00 AM', event: 'Med Management', loc: 'San Rafael', rendering: 'Reyes, Sam MD', status: 'Expected' });
// Kaiser med management with no auth: a flag for each upcoming visit, so a past visit that still counts shows up.
const pt = (mrn, last, first, dates) => ({ 'Per Nbr': mrn.slice(-3), MRN: mrn, 'Last Name': last, 'First Name': first, DOB: '1/1/1980', 'In Scope': 'Yes',
  'Primary Payer': 'Kaiser Synthetic', x: { location: 'San Rafael', appointments: dates.map(visit) } });
const patients = [pt('900601', 'Yesterly', 'Pastina', ['10/22/2026']), pt('900602', 'Todayson', 'Nowella', ['10/23/2026'])];

const cases = {
  morning: new Date(2026, 9, 23, 10, 0),
  justAfterMidnight: new Date(2026, 9, 23, 0, 30),
  lateEvening: new Date(2026, 9, 23, 23, 30),
  dayValue: new Date(Date.UTC(2026, 9, 23)), // a parseDate / dayOf day, as app.js passes the clinic run date
};
const out = { offset: new Date(2026, 9, 23, 12).getTimezoneOffset(), cases: {} };
for (const [name, now] of Object.entries(cases)) {
  const flags = R.allFlags(patients, [], { today: now });
  const rows = R.clinicRows('San Rafael', patients, [], flags, now);
  const pack = R.clinicPack('San Rafael', { patients, results: [], review: [], tasks: [], cfg: {}, runAt: now, version: 'test' });
  const se = R.clinicUnpack(Object.fromEntries(pack.sheets.map((s) => [s.name, s.rows])));
  out.cases[name] = {
    clinic: rows.map((r) => r.key),
    flagged: [...flags].filter(([, f]) => f.length).map(([k]) => k),
    runDate: rows.length ? rows[0]['Run Date'] : '',
    email: R.clinicEmail('San Rafael', rows, { runDate: now }).split('\n')[1],
    packed: pack.counts.patients,
    packRunDate: pack.meta['Run Date'],
    derived: R.clinicDerive(se, [], now).rows.map((r) => r.key),
  };
}
process.stdout.write(JSON.stringify(out));
