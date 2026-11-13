import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const F = require('../fees.js');

// v0.31 plan-year estimate, per-visit code and the Flow rateTable (docs/ARCHITECTURE.md section 4). Synthetic data only:
// payers, payer IDs, locations and rates are invented (rates end in .11).

test('visitCode: the section 4 table', () => {
  const code = (ev, payer) => F.visitCode(ev, '', payer).code;
  for (const ev of ['New Patient', 'Intake', 'Consult Telepsych', 'TMS Consult', 'Esketamine Consult', 'IV Ketamine Consult']) assert.equal(code(ev), '90792', ev);
  for (const ev of ['Med Management', 'MMI', 'Follow Up', 'TMS F/U', 'TFU', 'TTF', 'Telepsych']) assert.equal(code(ev), '99214', ev);
  for (const ev of ['Transition', 'ADHD Testing Review']) assert.equal(code(ev), '99215', ev);
  const ivk = F.visitCode('IV Ketamine Treatment Start', '', 'Acme Health PPO');
  assert.deepEqual([ivk.code, ivk.medication], ['99214', true], 'IVK: 99214 (not 99215 at the start), plus the medication');
  assert.equal(code('IVK', 'Kaiser Permanente'), '', 'Kaiser IVK: no estimate');
  assert.equal(code('IVK', 'VA Community Care'), '');
  for (const ev of ['Esketamine', 'Spravato Treatment', 'TMS', 'TMS Treatment', 'MT', 'MTR', 'TMS Mapping']) {
    const v = F.visitCode(ev);
    assert.equal(v.code, '', ev);
    assert.ok(v.none, ev);
  }
  assert.match(F.visitCode('Esketamine').note, /tracker BIDF/);
  assert.equal(F.visitCode('Med Management', 'ESK').code, '', 'the line given wins');
});

test('Estimate Codes seeds: IVK start 99214; transition and ADHD testing review 99215; old untouched seeds upgrade', () => {
  const seed = (k) => F.ESTIMATE_CODE_SEEDS.find((r) => r['Event keyword'] === k).Codes;
  assert.equal(seed('IV ketamine treatment start'), '99214 x1');
  assert.equal(seed('transition'), '99215 x1');
  assert.equal(seed('ADHD testing review'), '99215 x1');
  const old = [{ 'Event keyword': 'transition', Codes: '99214 x1', Source: 'default, review', Note: 'Transition visit; confirm it is 30 minutes' },
    { 'Event keyword': 'ADHD testing review', Codes: '99213 x1', Source: 'default, review', Note: 'mine' }];
  const up = F.upgradeInterventionSeeds(old);
  assert.equal(up.changed, 1);
  assert.equal(up.rows[0].Codes, '99215 x1');
  assert.equal(up.rows[1].Codes, '99213 x1', 'an edited row is kept');
});

test('planYearCodes: Config Plan-year estimate codes, default 90792 / 99214 / 99215', () => {
  assert.deepEqual(F.planYearCodes({}), ['90792', '99214', '99215']);
  assert.deepEqual(F.planYearCodes({ 'Plan-year estimate codes': '99214, 99215; bad' }), ['99214', '99215']);
  assert.equal(F.PLAN_YEAR_CONFIG_ROWS[0].setting, 'Plan-year estimate codes');
});

const pol = (o) => ({ 'Sohar Status': 'complete.eligible', 'Sohar Payer': 'Acme Health', 'Coverage Start': '01/01/2026', 'OOP Max': '5000', 'OOP Remaining': '4000', ...o });
const est = (at) => ({ 'Pulled At': at, 'Est 99214': '45.11' });

test('planYearEstimateDue: once per plan year; calendar year when coverage starts January 1', () => {
  const today = new Date(Date.UTC(2026, 10, 12));
  assert.equal(F.planYearEstimateDue(pol(), [], today).due, true);
  assert.equal(F.planYearEstimateDue(pol(), [], today).basis, 'Calendar');
  assert.equal(F.planYearEstimateDue(pol(), [est('2026-02-03T18:00:00Z')], today).due, false, 'estimated this calendar year');
  assert.equal(F.planYearEstimateDue(pol(), [est('2025-12-20T18:00:00Z')], today).due, true, 'last year');
  assert.equal(F.planYearEstimateDue(pol(), [{ 'Pulled At': '2026-03-01', Deductible: '500' }], today).due, true, 'a check without an estimate does not count');
});

test('planYearEstimateDue: no estimate for VA / TriWest / TRICARE, OOP met, not eligible', () => {
  const today = new Date(Date.UTC(2026, 10, 12));
  for (const payer of ['TriWest Healthcare Alliance', 'TRICARE West', 'VA Community Care', 'CHAMPVA']) {
    const r = F.planYearEstimateDue(pol({ 'Sohar Payer': payer }), [], today);
    assert.equal(r.due, false, payer);
    assert.match(r.reason, /\$0/);
  }
  assert.equal(F.planYearEstimateDue(pol({ 'Sohar Payer': 'Kaiser Permanente' }), [], today).due, true, 'Kaiser gets the estimate (assumption in the plan)');
  assert.match(F.planYearEstimateDue(pol({ 'OOP Remaining': '0' }), [], today).reason, /OOP max met/);
  assert.equal(F.planYearEstimateDue(pol({ 'OOP Status': 'Met', 'OOP Remaining': '' }), [], today).due, false);
  assert.equal(F.planYearEstimateDue(pol({ 'Sohar Status': 'complete.ineligible' }), [], today).due, false);
});

test('planYearEstimateDue: plan start not January 1 -> Unknown, estimated at both January 1 and the plan date', () => {
  const p = pol({ 'Coverage Start': '07/01/2025' });
  const nov = new Date(Date.UTC(2026, 10, 12)), feb = new Date(Date.UTC(2027, 1, 10));
  assert.equal(F.resetBasisOf(p, []), 'Unknown');
  const r = F.planYearEstimateDue(p, [est('2026-07-15T18:00:00Z')], nov);
  assert.deepEqual([r.due, r.basis, r.since], [false, 'Unknown', '07/01/2026']);
  const jan = F.planYearEstimateDue(p, [est('2026-07-15T18:00:00Z')], feb);
  assert.deepEqual([jan.due, jan.since], [true, '01/01/2027'], 'January 1 also gets one while the reset is unknown');
  assert.equal(F.planYearEstimateDue(p, [est('2026-07-15T18:00:00Z'), est('2027-01-05T18:00:00Z')], feb).due, false);
});

test('resetBasisOf: the deductible shows which reset is real; Plan and Calendar change the schedule', () => {
  const p = pol({ 'Coverage Start': '07/01/2025' });
  const chk = (at, rem) => ({ 'Pulled At': at, Deductible: '1000', 'Deductible Remaining': rem });
  // Reset in January (remaining back to the total), not at 7/1.
  assert.equal(F.resetBasisOf(p, [chk('2025-12-10', '200'), chk('2026-01-08', '1000')]), 'Calendar');
  assert.equal(F.resetBasisOf(p, [chk('2026-06-20', '300'), chk('2026-07-08', '300')]), 'Calendar', 'no reset at the plan date');
  // Reset at the plan date.
  assert.equal(F.resetBasisOf(p, [chk('2026-06-20', '300'), chk('2026-07-08', '1000')]), 'Plan');
  assert.equal(F.resetBasisOf(p, [chk('2025-12-10', '200'), chk('2026-01-08', '150')]), 'Plan', 'no reset in January');
  assert.equal(F.resetBasisOf(p, [chk('2026-02-01', '500'), chk('2026-03-01', '400')]), 'Unknown', 'no boundary crossed');
  assert.equal(F.resetBasisOf({ ...p, 'Reset Basis': 'plan' }, []), 'Plan', 'a saved basis wins');
  // Plan basis: once from the plan date only, so January does not ask again.
  const plan = { ...p, 'Reset Basis': 'Plan' };
  assert.equal(F.planYearEstimateDue(plan, [est('2026-07-15T18:00:00Z')], new Date(Date.UTC(2027, 1, 10))).due, false);
  const cal = { ...p, 'Reset Basis': 'Calendar' };
  assert.equal(F.planYearEstimateDue(cal, [est('2026-01-15T18:00:00Z')], new Date(Date.UTC(2026, 10, 12))).due, false, 'Calendar: the July date is ignored');
});

// A tiny synthetic fee index: tab "North" (Harbor Point), payer columns Acme and Zenith, MD and NP.
const ROWS = [];
for (const [payer, md] of [['Acme', 100], ['Zenith', 120]]) {
  for (const [code, f] of [['90792', 2], ['99214', 1], ['99215', 1.4]]) {
    ROWS.push({ Tab: 'North', Code: code, Payer: payer, Level: 'MD', Rate: md * f + 0.11 });
    if (!(payer === 'Zenith' && code === '99215')) ROWS.push({ Tab: 'North', Code: code, Payer: payer, Level: 'NP', Rate: md * f * 0.85 + 0.11 });
  }
}
const CTX = () => ({
  index: F.feeIndex(ROWS), teleKeywords: ['Tele'],
  locationRows: [{ 'NextGen Location': 'Harbor Point', 'Fee Tab': 'North' }],
  homeRows: [{ Provider: 'Doe MD, Jamie', 'Home Location': 'Harbor Point' }],
  payerMap: [{ Payer: 'Acme Health', 'Sohar payerId': 'AC001', 'Fee Tab': 'North', 'Fee Payer': 'Acme', Confidence: 'High', Source: 'auto' },
    { Payer: 'Zenith Plan', 'Sohar payerId': 'ZN002', 'Fee Tab': 'North', 'Fee Payer': 'Zenith', Confidence: 'High', Source: 'auto' },
    { Payer: 'Lowly Health', 'Sohar payerId': 'LW003', 'Fee Tab': 'North', 'Fee Payer': 'Acme', Confidence: 'Low', Source: 'auto' }],
  codes: F.PLAN_YEAR_CODES,
  payers: [{ payerId: 'AC001', payerName: 'Acme Health' }, { payerId: 'ZN002', payerName: 'Zenith Plan' }, { payerId: 'LW003', payerName: 'Lowly Health' }, { payerId: 'AC001', payerName: 'Acme Health' }],
});

test('rateTableFor: chart payer + discovery payers at the visit\'s fee tab and MD/NP level; incomplete or Low left out', () => {
  const md = F.rateTableFor({ location: 'Harbor Point', provider: 'Doe MD, Jamie', event: 'Med Management' }, CTX());
  assert.deepEqual(Object.keys(md.rateTable), ['AC001', 'ZN002'], 'Low confidence is left out; one entry per payer');
  assert.deepEqual(md.rateTable.AC001, [{ code: '90792', contractedRate: 200.11 }, { code: '99214', contractedRate: 100.11 }, { code: '99215', contractedRate: 140.11 }]);
  assert.equal(md.level, 'MD');
  const np = F.rateTableFor({ location: 'Harbor Point', provider: 'Roe PMHNP, Sam', event: 'Med Management' }, CTX());
  assert.deepEqual(Object.keys(np.rateTable), ['AC001'], 'Zenith has no NP 99215: left out');
  const tele = F.rateTableFor({ location: 'Telehealth', provider: 'Doe MD, Jamie', event: 'Telepsych' }, CTX());
  assert.equal(tele.feeTab, 'North', 'telehealth: the provider\'s home clinic');
  assert.match(F.rateTableFor({ location: 'Nowhere', provider: 'Doe MD, Jamie' }, CTX()).note, /No Fee Locations row/);
  assert.match(F.rateTableFor({ location: 'Harbor Point', provider: 'Jamie Doe' }, CTX()).note, /MD or NP/);
  assert.equal(JSON.stringify(md.rateTable).includes('"contractedRate":200.11'), true);
});

test('discoveryPayers: the payers discovery finds most, per state, per patient', () => {
  const o = (mrn, id, name, kind = 'Discovery') => ({ Key: `${mrn}|${id}`, MRN: mrn, Kind: kind, 'Returned payerId': id, 'Returned Payer': name });
  const st = { M1: 'CA', M2: 'CA', M3: 'CA', M4: 'TX' };
  const m = F.discoveryPayers([o('M1', 'AC001', 'Acme Health'), o('M2', 'AC001', 'Acme Health'), o('M2', 'AC001', 'Acme Health'), o('M3', 'ZN002', 'Zenith Plan'),
    o('M4', 'ZN002', 'Zenith Plan'), o('M1', 'QQ9', 'Verify only', 'Verification')], (k) => st[k]);
  assert.deepEqual(m.get('CA').map((x) => [x.payerId, x.n]), [['AC001', 2], ['ZN002', 1]]);
  assert.deepEqual(m.get('TX').map((x) => x.payerId), ['ZN002']);
});
