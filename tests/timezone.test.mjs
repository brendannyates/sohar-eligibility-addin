import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

// rules.dayOf turns a local date into a UTC-midnight day. A day value converted a second time west of UTC became the day
// before, so yesterday's visit still counted as the next visit (clinic set, flags, the Clinic Eligibility file). Each zone
// runs in its own process (TZ is read at start-up); the probe is tests/helpers/tz-probe.cjs. Synthetic patients only.
const PROBE = fileURLToPath(new URL('./helpers/tz-probe.cjs', import.meta.url));
// Minutes behind UTC on 10/23/2026 (getTimezoneOffset), to prove the zone really applied.
const ZONES = { 'America/Los_Angeles': 420, 'America/New_York': 240, 'UTC': 0, 'Asia/Tokyo': -540, 'Pacific/Auckland': -780 };

for (const [tz, offset] of Object.entries(ZONES)) {
  test(`timezone ${tz}: a visit yesterday is not in the clinic set, today's is`, () => {
    const out = JSON.parse(execFileSync(process.execPath, [PROBE], { env: { ...process.env, TZ: tz }, encoding: 'utf8' }));
    assert.equal(out.offset, offset, `the child ran in ${tz}`);
    for (const [name, c] of Object.entries(out.cases)) {
      assert.deepEqual(c.clinic, ['900602'], `${name}: clinic rows`);
      assert.deepEqual(c.flagged, ['900602'], `${name}: only today's visit is flagged`);
      assert.equal(c.runDate, '10/23/2026', `${name}: Run Date`);
      assert.equal(c.email, 'Run 10/23/2026. 1 of 1 patients need attention.', `${name}: clinic email`);
      assert.equal(c.packed, 1, `${name}: Clinic Eligibility file patients`);
      assert.equal(c.packRunDate, '10/23/2026', `${name}: file Run Date`);
      assert.deepEqual(c.derived, ['900602'], `${name}: CTT derivation`);
    }
  });
}
