import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const F = require('../fees.js');
const REC = require('../record.js');

// v0.32 (Brendan, 11/13): the rotation's ledger on Sohar Results, where each quote came from, and every booked visit
// priced (Sohar's quote first, then the PR calculator as the deductible is used). Synthetic data only.

test('estimateSourceOf: fee-math quotes are PR Calculator Estimate, the rest Sohar Cost Estimate', () => {
  assert.equal(REC.estimateSourceOf({ 'Estimate Note': 'Codes quoted 90792/99214/99215; add-in fee math; No NextGen payer on file' }), 'PR Calculator Estimate');
  assert.equal(REC.estimateSourceOf({ 'Estimate Note': 'Codes quoted 99214; Sohar Cost Estimate 2026-11-01' }), 'Sohar Cost Estimate');
});

test('ledgerResults: one result per rotation check, soonest upcoming visit, Import matches left out', () => {
  const V = 'https://app.soharhealth.com/verifications/22222222-2222-2222-2222-222222222222';
  const L = (o) => ({ MRN: '880001', LASFIR: 'SYNCCC', 'Sohar Status': 'complete.eligible', 'Verification ID': 'v-2', 'Sohar Link': V, Location: 'Austin', Event: 'Telepsych',
    Provider: 'Wood, Ash MD', 'Sohar Payer': 'Aetna', Deductible: '500', 'Deductible Remaining': '300', 'Last Checked': '2026-11-13 07:30',
    'Cost Estimate Link': 'https://app.soharhealth.com/cost-estimates/x', ...o });
  const ledger = [L({ 'Appt Date': '2026-11-01', Time: '9:00 AM' }), L({ 'Appt Date': '2026-11-20', Time: '10:00 AM' }), L({ 'Appt Date': '2026-11-16', Time: '9:00 AM' }),
    L({ MRN: '880002', 'Verification ID': 'v-known', 'Sohar Link': '', 'Appt Date': '2026-11-16' }),
    L({ MRN: '880003', 'Verification ID': '', 'Discovery ID': 'd-9', 'Sohar Link': '', 'Discovery Link': 'https://app.soharhealth.com/discoveries/d-9', 'Sohar Status': 'complete.found', 'Appt Date': '2026-11-18' }),
    L({ MRN: '880004', 'Sohar Status': '', 'Verification ID': 'v-none' })];
  const out = REC.ledgerResults(ledger, [{ MRN: '880002', 'Sohar Ref': 'V-KNOWN' }], { asOf: '11/13/2026' });
  assert.equal(out.length, 2, 'the check the Import matched and the row with no result are left out');
  const a = out.find((r) => r.MRN === '880001');
  assert.equal(a['Next Appt'], '11/16/2026'); assert.equal(a['Upcoming Appts'], '2'); assert.equal(a['Next Location'], 'Austin');
  assert.equal(a.Source, 'Rotation'); assert.equal(a['Sohar Ref'], 'v-2'); assert.equal(a['Sent As'], 'Verification'); assert.equal(a['Deductible Remaining'], '300');
  assert.match(a['Cost Estimate Link'], /cost-estimates/);
  const d = out.find((r) => r.MRN === '880003');
  assert.equal(d['Sent As'], 'Discovery'); assert.match(d['Sohar Link'], /discoveries/); assert.equal(d['Discovery Link'], '');
});

test('visitPlan: Sohar quote on the first visit, then the PR calculator as the deductible is used', () => {
  const fin = { 'Sohar Status': 'complete.eligible', Deductible: '500', 'Deductible Remaining': '300', 'OOP Max': '1000', 'OOP Remaining': '400', Coinsurance: '20%', 'Sohar Payer': 'Aetna', 'Sohar Date': '2026-11-13' };
  const vs = ['11/16/2026', '11/23/2026', '11/30/2026', '12/07/2026', '01/04/2027'].map((date) => ({ MRN: '880001', lasfir: 'SYNCCC', date, time: '9:00 AM', event: 'Telepsych', location: 'Austin' }));
  vs.push({ MRN: '880001', date: '11/18/2026', time: '8:00 AM', event: 'TMS Treatment Unilateral' });
  const opts = { codeOf: (v, l, p) => F.visitCode(v.event, l, p), benefitsOf: () => fin, quotesOf: () => ({ 99214: { amount: 150, source: REC.SOHAR_ESTIMATE, date: '2026-11-13' } }), rateOf: () => 150, math: F.estimateMath };
  const rows = REC.visitPlan(vs, opts);
  const by = (d) => rows.find((r) => r['Appt Date'] === d);
  assert.equal(by('11/16/2026')['Est Patient'], '$150.00'); assert.equal(by('11/16/2026')['Estimate Source'], 'Sohar Cost Estimate');
  assert.equal(by('11/18/2026')['Est Patient'], '', 'a TMS treatment session is never estimated'); assert.match(by('11/18/2026')['Estimate Note'], /TMS treatment/);
  assert.equal(by('11/23/2026')['Est Patient'], '$150.00'); assert.equal(by('11/23/2026')['Estimate Source'], 'PR Calculator Estimate');
  assert.equal(by('11/23/2026')['Ded Left After'], '$0.00'); assert.equal(by('11/23/2026').Milestone, 'Deductible met at this visit');
  assert.equal(by('11/30/2026')['Est Patient'], '$30.00', '20% of $150 once the deductible is met');
  assert.equal(by('12/07/2026')['Est Patient'], '$30.00'); assert.equal(by('12/07/2026')['OOP Left After'], '$40.00');
  const jan = by('01/04/2027');
  assert.match(jan['Estimate Note'], /New plan year/); assert.equal(jan['Deductible Remaining'], '$500.00'); assert.equal(jan['Est Patient'], '$150.00');
  assert.equal(jan['Estimate Source'], 'PR Calculator Estimate', "last year's Sohar quote is re-priced");
  // No rate: the quote stands, capped at the OOP max left; a quote equal to the copay is copay only.
  const cp = REC.visitPlan(vs.slice(0, 2), { ...opts, rateOf: () => null, benefitsOf: () => ({ ...fin, Copay: '25' }), quotesOf: () => ({ 99214: { amount: 25, source: REC.SOHAR_ESTIMATE } }) });
  assert.deepEqual(cp.map((r) => r['Est Patient']), ['$25.00', '$25.00']); assert.equal(cp[0]['Ded Left After'], '$300.00', 'a copay never counts toward the deductible');
  const inel = REC.visitPlan(vs.slice(0, 1), { ...opts, benefitsOf: () => ({ 'Sohar Status': 'complete.ineligible' }) });
  assert.match(inel[0]['Estimate Note'], /Not eligible/); assert.equal(inel[0]['Est Patient'], '');
  const ivk = REC.visitPlan([{ MRN: '880009', date: '11/16/2026', event: 'IV Ketamine Follow Up' }], { ...opts, medOf: () => 400 });
  assert.equal(ivk[0]['Est Patient'], '$550.00'); assert.match(ivk[0]['Estimate Note'], /\$400\.00 IV ketamine medication/);
  for (const c of REC.VISIT_PLAN_COLS) assert.ok(c in rows[0], c);
  assert.ok(!REC.VISIT_PLAN_COLS.some((c) => /name|dob|member/i.test(c)), 'no names, DOBs or member IDs');
});

test('quotesByMrn: newest per code from Flow Results and the ledger, sandbox never', () => {
  const q = REC.quotesByMrn([{ MRN: '0880001', 'Est 99214': '35.00', 'Est 90792': '60', 'Last Updated': '2026-11-10', 'Estimate Note': 'add-in fee math' },
    { MRN: '880001', 'Est 99214': '999', File: 'Flow Results (SANDBOX).csv', 'Last Updated': '2026-11-14' }],
  [{ MRN: '880001', 'Est Codes': '99214', 'Sohar Est Patient': '$40.00', 'Last Checked': '2026-11-13', 'Estimate Note': 'Sohar Cost Estimate' }]);
  const m = q.get('880001');
  assert.equal(m['99214'].amount, 40); assert.equal(m['99214'].source, 'Sohar Cost Estimate');
  assert.equal(m['90792'].amount, 60); assert.equal(m['90792'].source, 'PR Calculator Estimate');
});

test('projectionFromPlan: the patient view bars follow the priced visits; next plan year is listed, not applied', () => {
  const fin = { Deductible: '500', 'Deductible Remaining': '300', 'OOP Max': '1000', 'OOP Remaining': '400' };
  const plan = [
    { 'Appt Date': '11/16/2026', Event: 'Telepsych', 'Est Code': '99214', 'Est Patient': '$150.00', 'Estimate Source': 'Sohar Cost Estimate', 'Ded Left After': '$150.00', 'OOP Left After': '$250.00', Milestone: '' },
    { 'Appt Date': '11/23/2026', Event: 'Telepsych', 'Est Code': '99214', 'Est Patient': '$150.00', 'Estimate Source': 'PR Calculator Estimate', 'Ded Left After': '$0.00', 'OOP Left After': '$100.00', Milestone: 'Deductible met at this visit' },
    { 'Appt Date': '11/25/2026', Event: 'TMS Treatment', 'Est Code': '', 'Est Patient': '', 'Estimate Note': 'TMS treatment session: verification only' },
    { 'Appt Date': '01/04/2027', Event: 'Telepsych', 'Est Code': '99214', 'Est Patient': '$150.00', 'Estimate Source': 'PR Calculator Estimate', 'Ded Left After': '$350.00', 'OOP Left After': '$850.00', 'Estimate Note': 'New plan year: deductible and OOP max start over' }];
  const p = REC.projectionFromPlan(fin, plan);
  assert.equal(p.plan, true);
  assert.equal(p.ded.applied, 300); assert.equal(p.ded.at['Appt Date'], '11/23/2026');
  assert.equal(p.oop.applied, 300, 'the January visit is not applied to this year');
  assert.equal(p.rows.length, 4); assert.equal(p.rows[3].nextYear, true); assert.equal(p.rows[0].source, 'Sohar Cost Estimate');
  assert.equal(p.missing.length, 1);
  assert.equal(REC.projectionFromPlan({}, plan), null, 'no benefit figures, no projection');
});
