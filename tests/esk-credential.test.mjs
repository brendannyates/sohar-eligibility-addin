import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const R = require('../rules.js');
const PR = require('../payer-rules.js');

// R13 esketamine payer credential rules. Made-up payers only: the real list lives on each workbook's Payer Credential
// Rules sheet (imported from OneDrive) and is never committed.
const TODAY = new Date(2026, 9, 15); // Thu 10/15/2026 local
const SHEET = [
  PR.HEADERS,
  ['Acme Health', '', 'NP', 'Y', 'Yes', 'Yes', 'Yes', 'Yes', 'acme health', ''],
  ['Acme Health', '', 'MD', 'Y', 'N/A', 'Yes', 'Yes', 'Yes', 'acme health', ''],
  ['Zenith CA', 'CA', 'NP', 'N', 'No', 'No', 'No', 'No', 'zenith', ''],
  ['Zenith CA', 'CA', 'MD', 'N', 'N/A', 'Yes', 'Yes', 'Yes', 'zenith', ''],
  ['Zenith WA', '', 'NP', 'N', 'Yes', 'Yes', 'Yes', 'Yes', 'zenith', ''],
  ['Zenith WA', '', 'MD', 'N', 'N/A', 'Yes', 'Yes', 'Yes', 'zenith', ''],
  ['Orbit (Fed)', '', 'NP', 'Y', 'No', 'No', 'No', 'Yes', 'orbit; orbit federal', 'NP may do observation only'],
  ['Orbit (Fed)', '', 'MD', 'Y', 'N/A', 'Yes', 'Yes', 'Yes', 'orbit; orbit federal', ''],
  ['Orbit Medicare', '', 'NP', '', 'Yes', 'Yes', 'Yes', 'Yes', 'orbit medicare', ''],
  ['Orbit Medicare', '', 'MD', '', 'N/A', 'Yes', 'Yes', 'Yes', 'orbit medicare', ''],
];
const RULES = PR.read(SHEET).rules;
const CFG = { today: TODAY, sfhpProviders: [], reverifyDays: 7, payerRules: RULES, locationState: { 'Bayview': 'WA', 'Harbor Point': 'CA' } };

const NP = 'Lark, Jordan PMHNP', MD = 'Reyes, Sam MD';
const appt = (o) => Object.assign({ date: '10/20/2026', time: '09:00', event: 'Esketamine Follow Up', loc: 'Harbor Point', rendering: NP, status: '' }, o);
function pat(payer, appts, o = {}) {
  return Object.assign({ 'Per Nbr': '3001', MRN: 'M3001', 'Last Name': 'Testerson', 'First Name': 'Alexa', State: 'CA', 'In Scope': 'Yes', 'Primary Payer': payer, 'Secondary Payer': '' }, o, {
    x: { finClass: 'Contracted Commercial', finClassLstEnc: 'Contracted Commercial', appointments: appts, verifyRows: [] },
  });
}
const ESK = ['ESK_CRED', 'ESK_STATE', 'BESSE_NO', 'BESSE_PENDING'];
const esk = (p, cfg = CFG) => R.patientFlags(p, [], cfg).filter((f) => ESK.includes(f.code));
const codes = (fl) => fl.map((f) => f.code);

test('pills: ESK credential high/red, Besse not approved medium/amber, state check and Besse pending low', () => {
  const by = Object.fromEntries(R.PILLS.map((p) => [p.code, p]));
  assert.deepEqual([by.ESK_CRED.severity, by.ESK_CRED.color], ['high', 'red']);
  assert.deepEqual([by.BESSE_NO.severity, by.BESSE_NO.color], ['medium', 'amber']);
  assert.equal(by.ESK_STATE.severity, 'info');
  assert.equal(by.BESSE_PENDING.severity, 'info');
});

test('isESKEvent: esketamine, Spravato and ESK; not IV ketamine or TMS', () => {
  for (const e of ['Esketamine Follow Up', 'Spravato First Dose', 'ESK Obs', 'Esketamine Tx Start']) assert.ok(R.isESKEvent(e), e);
  for (const e of ['IV Ketamine Follow Up', 'TMS Treatment Unilateral', 'Desk time', 'Med Management Telepsych']) assert.ok(!R.isESKEvent(e), e);
});

test('no sheet (no rules) = no flags; payer not on the list = no flags; non-ESK visits = no flags', () => {
  const p = pat('Zenith Blue', [appt({ event: 'Esketamine First Treatment' })]);
  assert.deepEqual(esk(p, { ...CFG, payerRules: undefined }), []);
  assert.deepEqual(esk(p, { ...CFG, payerRules: [] }), []);
  assert.deepEqual(esk(p, { ...CFG, payerRules: R.payerRulesFrom([]) }), []);
  assert.deepEqual(esk(pat('Unlisted Plan PPO', [appt({ event: 'Esketamine First Treatment' })])), []);
  assert.deepEqual(esk(pat('Zenith Blue', [appt({ event: 'TMS Treatment Unilateral' }), appt({ event: 'IV Ketamine Follow Up' })])), []);
  assert.deepEqual(esk(pat('Zenith Blue', [appt({ date: '10/01/2026', event: 'Esketamine First Treatment' })])), [], 'past visits are not checked');
});

test('NP blocked: high flag with the reason, event, date and rendering; MD allowed', () => {
  const f = esk(pat('Zenith Blue PPO', [appt({ event: 'Esketamine First Treatment' })]));
  const cred = f.find((x) => x.code === 'ESK_CRED');
  assert.equal(cred.severity, 'high');
  assert.equal(cred.detail, 'Zenith CA doesn’t allow an NP to bill the esketamine first treatment: Esketamine First Treatment on 10/20/2026 with Lark, Jordan PMHNP');
  assert.deepEqual(cred.appt, { date: '10/20/2026', event: 'Esketamine First Treatment', rendering: NP });
  assert.ok(!codes(esk(pat('Zenith Blue PPO', [appt({ event: 'Esketamine First Treatment', rendering: MD })]))).includes('ESK_CRED'));
  assert.ok(!codes(esk(pat('Acme Health HMO', [appt({ event: 'Esketamine First Treatment' })]))).includes('ESK_CRED'));
  // An NP whose surname is Do is still an NP.
  assert.ok(codes(esk(pat('Zenith Blue PPO', [appt({ rendering: 'Do, Minh NP' })]))).includes('ESK_CRED'));
  assert.ok(!codes(esk(pat('Zenith Blue PPO', [appt({ rendering: 'Reyes, Dana DO' })]))).includes('ESK_CRED'));
});

test('observation-only NP: follow-ups pass, first treatment and start are flagged', () => {
  const p = pat('Orbit Federal Services', [
    appt({ date: '10/16/2026', event: 'Esketamine Tx Start' }),
    appt({ date: '10/19/2026', event: 'Esketamine First Treatment' }),
    appt({ date: '10/22/2026', event: 'Esketamine Follow Up' }),
  ]);
  const f = esk(p).filter((x) => x.code === 'ESK_CRED');
  assert.deepEqual(f.map((x) => x.appt.date), ['10/16/2026', '10/19/2026']);
  assert.match(f[0].detail, /^Orbit \(Fed\) doesn’t allow an NP to bill the esketamine start \(NP may do observation only\)/);
  assert.match(f[1].detail, /esketamine first treatment/);
  assert.deepEqual(esk(pat('Orbit Federal', [appt({ event: 'Spravato Observation' })])), []);
});

test('state rules: the patient’s state decides, else the visit location; still unknown = low review item', () => {
  const visit = [appt({ event: 'Esketamine Follow Up' })];
  assert.ok(codes(esk(pat('Zenith Blue', visit, { State: 'CA' }))).includes('ESK_CRED'));
  assert.ok(!codes(esk(pat('Zenith Blue', visit, { State: 'WA' }))).includes('ESK_CRED'));
  assert.ok(!codes(esk(pat('Zenith Blue', [appt({ loc: 'Bayview' })], { State: '' }))).includes('ESK_CRED'), 'Config location state WA');
  const amb = esk(pat('Zenith Blue', [appt({ loc: 'Somewhere' })], { State: '' }));
  assert.deepEqual(codes(amb), ['BESSE_NO', 'ESK_STATE'], 'both state rules say not Besse-approved; NP answer disagrees');
  const st = amb.find((x) => x.code === 'ESK_STATE');
  assert.equal(st.severity, 'info');
  assert.match(st.detail, /^Zenith Blue matches Zenith CA and Zenith WA on Payer Credential Rules: check which state the plan is from/);
  // Same answer in every state: no review item.
  assert.deepEqual(codes(esk(pat('Zenith Blue', [appt({ loc: 'Somewhere', rendering: MD })], { State: '' }))), ['BESSE_NO']);
});

test('Besse: not approved = medium, pending (blank) = low, approved = nothing; one per patient', () => {
  const two = [appt({ date: '10/16/2026', rendering: MD }), appt({ date: '10/20/2026', rendering: MD })];
  const no = esk(pat('Zenith Blue', two));
  assert.deepEqual(codes(no), ['BESSE_NO']);
  assert.equal(no[0].severity, 'medium');
  assert.equal(no[0].detail, "Payer isn't Besse-approved: Zenith CA (Esketamine Follow Up 10/16/2026)");
  const pend = esk(pat('Orbit Medicare Advantage', two));
  assert.deepEqual(codes(pend), ['BESSE_PENDING']);
  assert.equal(pend[0].detail, 'Besse approval pending for Orbit Medicare (Esketamine Follow Up 10/16/2026)');
  assert.deepEqual(esk(pat('Acme Health', two)), []);
});

test('Review sheet, clinic rows and collapsed lines carry the flags with their fix text', () => {
  const p = pat('Orbit Federal', [appt({ date: '10/16/2026', event: 'Esketamine Tx Start' }), appt({ date: '10/17/2026', event: 'Esketamine Start' })]);
  const fm = R.allFlags([p], [], CFG);
  const rv = R.reviewRows(fm, [p]).filter((r) => r.Issue === 'ESK credential');
  assert.equal(rv.length, 2, 'one Review row per visit');
  assert.equal(rv[0].Severity, 'High');
  assert.equal(rv[0].Fix, 'Reassign to a provider the payer allows, or confirm with billing.');
  const c = R.collapseFlags(fm.get('M3001')).filter((f) => f.code === 'ESK_CRED');
  assert.equal(c.length, 1);
  assert.equal(c[0].detail, 'Orbit (Fed) doesn’t allow an NP to bill the esketamine start (NP may do observation only): 2 visits with Lark, Jordan PMHNP (10/16, 10/17)');
  // Different reasons for one provider stay separate lines.
  const mixed = R.allFlags([pat('Orbit Federal', [appt({ date: '10/16/2026', event: 'Esketamine Start' }), appt({ date: '10/17/2026', event: 'Esketamine First Treatment' })])], [], CFG);
  assert.equal(R.collapseFlags(mixed.get('M3001')).filter((f) => f.code === 'ESK_CRED').length, 2);
  const row = R.clinicRows('All', [p], [], fm, TODAY)[0];
  assert.match(row.Pills, /ESK credential/);
  assert.match(row['Review Items'], /2 visits with Lark, Jordan PMHNP/);
  const pend = R.allFlags([pat('Orbit Medicare', [appt({ rendering: MD })])], [], CFG);
  const cr = R.clinicReviewRows('All', [pat('Orbit Medicare', [appt({ rendering: MD })])], [], pend, TODAY).find((r) => r.Issue === 'BESSE_PENDING');
  assert.equal(cr.Severity, 'Low');
  assert.match(R.clinicRows('All', [pat('Orbit Medicare', [appt({ rendering: MD })])], [], pend, TODAY)[0].Notes, /Besse approval pending for Orbit Medicare/);
});

test('sheet rows: written rows read back as the same rules, also after a column is moved in Excel', () => {
  const rows = R.payerRuleRows(RULES);
  assert.deepEqual(Object.keys(rows[0]), PR.HEADERS);
  const back = R.payerRulesFrom(rows);
  const sig = (rs) => rs.map((r) => [r.payer, r.state, r.cred, r.besse, r.first, r.start, r.obs, r.match.join(), r.notes]);
  assert.deepEqual(sig(back), sig(RULES));
  const moved = rows.map(({ Notes, ...rest }) => ({ Notes, ...rest }));
  assert.deepEqual(sig(R.payerRulesFrom(moved)), sig(RULES));
  // A staff edit on the sheet changes the answer.
  const edited = rows.map((r) => (r.Payer === 'Zenith CA' && r.Credential === 'NP' ? { ...r, 'Esketamine Observation': 'Yes' } : r));
  assert.ok(!codes(esk(pat('Zenith Blue', [appt()]), { ...CFG, payerRules: R.payerRulesFrom(edited) })).includes('ESK_CRED'));
});
