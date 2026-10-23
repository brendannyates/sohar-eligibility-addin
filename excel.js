/* excel.js — everything that touches the workbook. Table and sheet names must match the template. */
(function (root) {
  'use strict';

  const SHEETS = { import: 'Import', payerMap: 'Payer Map', sohar: 'Sohar Payers', config: 'Config',
    patients: 'Patients', verification: 'Verification', discovery: 'Discovery', review: 'Review', dashboard: 'Dashboard', results: 'Sohar Results', tasks: 'NextGen Tasks' };
  const TABLES = { payerMap: 'tblPayerMap', sohar: 'tblSoharPayers', config: 'tblConfig', locState: 'tblLocationState',
    patients: 'tblPatients', verification: 'tblVerification', discovery: 'tblDiscovery', review: 'tblReview' };
  // Created by the add-in on first use; not checked by templateCheck.
  const RESULTS_TABLE = 'tblSoharResults';
  const TASKS_TABLE = 'tblNextGenTasks';
  const CHUNK = 1000;

  // delayForCellEdit: wait out a cell someone is still typing in instead of failing (ExcelApi 1.11).
  async function run(fn) {
    return Office.context.requirements.isSetSupported('ExcelApi', '1.11') ? Excel.run({ delayForCellEdit: true }, fn) : Excel.run(fn);
  }

  async function readImportRows() {
    return run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(SHEETS.import);
      const used = ws.getUsedRangeOrNullObject(true);
      used.load('rowCount,columnCount,rowIndex,columnIndex,isNullObject');
      await ctx.sync();
      if (used.isNullObject) return [];
      const rows = [];
      const cols = used.columnCount;
      for (let start = 0; start < used.rowCount; start += 5000) {
        const n = Math.min(5000, used.rowCount - start);
        const rg = ws.getRangeByIndexes(used.rowIndex + start, used.columnIndex, n, cols);
        rg.load('values');
        await ctx.sync();
        for (const r of rg.values) rows.push(r);
      }
      return rows;
    });
  }

  // Rows per read. Excel on the web caps a response at 5MB; Sohar Uploads rows carry ~2KB of JSON each.
  const READ_CHUNK = 500;

  // ---------- fail-safe saves ----------
  // Sheets + tables the add-in rebuilds with writeAddinTable. hidden: kept off the tab bar.
  const OWNED = {
    tblSoharResults: { sheet: 'Sohar Results' }, tblNextGenTasks: { sheet: 'NextGen Tasks' }, tblExport: { sheet: 'Export' },
    tblIrCrossRef: { sheet: 'IR Cross-Ref' }, tblBenefits: { sheet: 'Benefits' }, tblDiscoveryVerify: { sheet: 'Discovery Verify' },
    tblVerifyFailed: { sheet: 'Discovery from Verify' }, tblPayerEvidence: { sheet: 'Payer Evidence' },
    tblSoharUploads: { sheet: 'Sohar Uploads', hidden: true }, tblIrUploads: { sheet: 'IR Uploads', hidden: true },
    // Payer learning (v0.10): the outcomes ledger is hidden (it has Per Nbr); the log and the rollup have no patient identifiers.
    tblPayerOutcomes: { sheet: 'Payer Outcomes', hidden: true }, tblPayerMapLog: { sheet: 'Payer Map Log' }, tblPayerLearning: { sheet: 'Payer Learning' },
    // Sohar ref -> Per Nbr memory (append-only, kept by "Clear saved Sohar results") and "Re-verify anyway" requests.
    tblSoharRefs: { sheet: 'Sohar Refs', hidden: true }, tblReverifyRequests: { sheet: 'Re-verify Requests', hidden: true },
    // Fee estimates (v0.11, a test, Off by default). Created only when Config "Fee estimates" is On; every one is safe to
    // delete. The schedule and its info hold rates, so they are hidden; the mapping tables are visible and editable.
    tblFeeSchedule: { sheet: 'Fee Schedule', hidden: true }, tblFeeMeta: { sheet: 'Fee Schedule Info', hidden: true },
    tblFeeLocations: { sheet: 'Fee Locations' }, tblProviderHome: { sheet: 'Provider Home Clinic' },
    tblFeePayerMap: { sheet: 'Fee Payer Map' }, tblFeePayerMapLog: { sheet: 'Fee Payer Map Log' },
    tblEstimateCodes: { sheet: 'Estimate Codes' }, tblEstimateDefaults: { sheet: 'Estimate Defaults' },
    // v0.12. Guarantor resubmission requests and payer test rows carry MRNs: hidden. The payer mapping discovery rows are
    // visible (the file's rows, like Discovery Verify). Metrics, their history and the guarantor themes have no identifiers.
    tblGuarantorRequests: { sheet: 'Guarantor Requests', hidden: true }, tblPayerTests: { sheet: 'Payer Tests', hidden: true },
    tblPayerMapping: { sheet: 'Payer Mapping' }, tblPayerMetrics: { sheet: 'Payer Metrics' }, tblPayerMetricHistory: { sheet: 'Payer Metrics History' },
    tblGuarantorThemes: { sheet: 'Guarantor Themes' },
    // Inside a clinic file: the clinic's own "Re-verify anyway" ticks, brought back to the admin workbook.
    tblClinicRequests: { sheet: 'Clinic Requests' },
    // v0.14: notes and payer notes from returned Clinic Eligibility files. Visible: a person reads them and acts.
    tblClinicNotes: { sheet: 'Clinic Notes' }, tblClinicPayerNotes: { sheet: 'Clinic Payer Notes' },
    // Which Clinic Work rows were already brought back (a file sent twice changes nothing). Keys are MRNs: hidden.
    tblClinicWorkApplied: { sheet: 'Clinic Work Applied', hidden: true },
    // Esketamine payer credential rules (R13): imported from the shared CSV in OneDrive, visible and editable. No patient data.
    tblPayerCredRules: { sheet: 'Payer Credential Rules' },
  };
  // Fee estimate tables by short name (app.js uses readFee / writeFee with these).
  const FEE_TABLES = { schedule: 'tblFeeSchedule', meta: 'tblFeeMeta', locations: 'tblFeeLocations', homes: 'tblProviderHome',
    payerMap: 'tblFeePayerMap', payerLog: 'tblFeePayerMapLog', codes: 'tblEstimateCodes', defaults: 'tblEstimateDefaults' };
  // A save builds the new rows on a hidden staging sheet (STAGE_SHEET) and only then makes them a table named STG.
  // So a STG table always holds complete new data. The old sheet is deleted only in the swap that renames the
  // staging sheet and then the table into place. "Real table missing + STG present" can only mean an interrupted
  // swap, and rescue() finishes it. Sheet names max 31 chars: '~saving tblDiscoveryVerify' is 26.
  const STAGE_SHEET = (t) => `~saving ${t}`, STG = (t) => `${t}_stg`, STAGE_IMPORT = '~saving Import';

  /** Finish an interrupted swap: rename the staging sheet and its STG table to the real names. True when done. */
  async function rescue(ctx, tableName) {
    const { sheet, hidden } = OWNED[tableName], wb = ctx.workbook;
    try {
      const stg = wb.tables.getItem(STG(tableName)), sws = stg.worksheet;
      sws.load('id,name');
      const clash = wb.worksheets.getItemOrNullObject(sheet); clash.load('isNullObject,id');
      await ctx.sync();
      // A different sheet still holding the real name has already lost its table: it is the old sheet, half-deleted.
      if (!clash.isNullObject && clash.id !== sws.id) clash.delete();
      sws.visibility = hidden ? Excel.SheetVisibility.hidden : Excel.SheetVisibility.visible;
      if (sws.name !== sheet) sws.name = sheet;
      stg.name = tableName; // last: the real name only ever appears on complete data
      await ctx.sync();
      return true;
    } catch (e) { console.warn(`could not finish the interrupted save of ${sheet}`, e); return false; }
  }

  /**
   * The table to read for tableName: the real one, else (interrupted save) the STG copy after trying to finish
   * the swap, else null. Never treats a recoverable save as missing.
   */
  async function settle(ctx, tableName) {
    const wb = ctx.workbook;
    const probe = () => {
      const real = wb.tables.getItemOrNullObject(tableName); real.load('isNullObject');
      const stg = OWNED[tableName] ? wb.tables.getItemOrNullObject(STG(tableName)) : null;
      if (stg) stg.load('isNullObject');
      return { real, stg };
    };
    let p = probe();
    await ctx.sync();
    if (!p.real.isNullObject) return p.real;
    if (!p.stg || p.stg.isNullObject) return null;
    if (await rescue(ctx, tableName)) recovered.add(OWNED[tableName].sheet);
    p = probe();
    await ctx.sync();
    return !p.real.isNullObject ? p.real : !p.stg.isNullObject ? p.stg : null;
  }
  const recovered = new Set(); // sheets whose interrupted save was finished, for the pane to report

  /** Table rows as objects. Throws if the table is missing, unless missingOk (then []). Other errors always throw. */
  async function readTable(tableName, missingOk) {
    return run(async (ctx) => {
      const t = await settle(ctx, tableName);
      if (!t) {
        if (missingOk) return [];
        throw new Error(`Table ${tableName} is missing. Use the Sohar Eligibility template workbook.`);
      }
      const h = t.getHeaderRowRange(); h.load('values');
      const b = t.getDataBodyRange(); b.load('rowIndex,columnIndex,rowCount,columnCount');
      await ctx.sync();
      const headers = h.values[0].map((x) => String(x).trim());
      const values = [];
      for (let i = 0; i < b.rowCount; i += READ_CHUNK) {
        const rg = t.worksheet.getRangeByIndexes(b.rowIndex + i, b.columnIndex, Math.min(READ_CHUNK, b.rowCount - i), b.columnCount);
        rg.load('values');
        await ctx.sync();
        for (const r of rg.values) values.push(r);
      }
      return values
        .filter((r) => r.some((v) => v !== '' && v != null))
        .map((r) => Object.fromEntries(headers.map((k, i) => [k, r[i]])));
    });
  }
  // For tables the add-in creates on first use: missing means "nothing saved yet"; any other error stops the run
  // so a failed read can never be mistaken for an empty table and overwritten.
  const readOwned = (tableName) => readTable(tableName, true);

  async function readConfig() {
    const rows = await readTable(TABLES.config);
    const cfg = {};
    for (const r of rows) cfg[String(r.Setting || '').trim()] = r.Value;
    const loc = {};
    for (const r of await readTable(TABLES.locState)) if (r.Location) loc[String(r.Location).trim()] = String(r.State || '').trim();
    const scope = String(cfg['Scope'] || 'upcoming').toLowerCase().startsWith('all') ? 'all' : 'upcoming';
    let asOf = null;
    const raw = cfg['As-of date'];
    if (raw !== '' && raw != null) asOf = root.SoharCore.parseDate(raw);
    return {
      asOf, scope, locationState: loc,
      defaultState: String(cfg['Default state'] || '').trim(),
      teleKeywords: String(cfg['Telehealth keywords'] || 'Tele').split(/[;,]/).map((s) => s.trim()).filter(Boolean),
      posTele: String(cfg['POS code: telehealth'] || '10').trim(),
      posOffice: String(cfg['POS code: office'] || '11').trim(),
      specialty: String(cfg['Specialty code'] || 'PSY').trim(),
      useGuarantorAsSubscriber: /^y|true/i.test(String(cfg['Use guarantor as subscriber'] || 'No')),
      // "Approaching" = remaining at or under this share of the deductible or OOP max. Optional Config row; default 20%.
      approachingPct: (() => { const n = parseFloat(String(cfg['Benefit approaching %'] ?? '').replace('%', '')); return isNaN(n) ? 0.2 : (n > 1 ? n / 100 : n); })(),
      // Optional rows, never required: SFHP-approved rendering providers (names or last names, ; separated),
      // days before a Sohar verification counts as stale, and the Clinic tab's last location.
      sfhpProviders: String(cfg['SFHP providers'] ?? '').split(/[;\n]/).map((s) => s.trim()).filter(Boolean),
      reverifyDays: (() => { const n = parseInt(String(cfg['Re-verify days'] ?? ''), 10); return isNaN(n) || n < 0 ? 7 : n; })(),
      // Re-verify skip window: Skip recent verifications (On), Re-verify days (7), Re-verify in new month (Off). See core.reverifyConfig.
      reverify: root.SoharCore.reverifyConfig(cfg),
      // v0.12: payer mapping discovery (On) and payer tests (Off), see core.payerTestConfig; and when Sohar ids became MRNs.
      payerTests: root.SoharCore.payerTestConfig(cfg),
      mrnSince: cfg['MRN ids since'] == null ? '' : cfg['MRN ids since'],
      clinicLocation: String(cfg['Clinic location'] ?? '').trim(),
      // Payer learning: optional rows (Learning, Learn: ...), defaults when missing. See core.learningConfig.
      learning: root.SoharCore.learningConfig(cfg),
      // Fee estimates (v0.11): Off unless Config "Fee estimates" is On. Off when fees.js is not loaded.
      fees: root.SoharFees ? root.SoharFees.feeConfig(cfg) : { on: false },
    };
  }

  /**
   * Add Config rows that are missing (by Setting), with their plain-English description in the third column when the
   * table has one. Existing rows are never changed. rows: [{setting, value, what}]. Returns the settings added.
   */
  async function ensureConfigRows(rows) {
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItem(TABLES.config);
      const h = t.getHeaderRowRange(); h.load('values');
      const b = t.getDataBodyRange(); b.load('values');
      await ctx.sync();
      const headers = h.values[0].map((x) => String(x).trim());
      const si = Math.max(0, headers.indexOf('Setting')), vi = headers.indexOf('Value') >= 0 ? headers.indexOf('Value') : 1;
      const di = headers.findIndex((x, j) => j !== si && j !== vi);
      const have = new Set(b.values.map((r) => String(r[si] ?? '').trim().toLowerCase()).filter(Boolean));
      const add = (rows || []).filter((r) => !have.has(String(r.setting).trim().toLowerCase()));
      if (!add.length) return [];
      t.rows.add(null, add.map((r) => headers.map((x, j) => (j === si ? r.setting : j === vi ? r.value : j === di ? (r.what || '') : ''))));
      await ctx.sync();
      return add.map((r) => r.setting);
    });
  }

  async function readPayerMap() {
    const rows = await readTable(TABLES.payerMap);
    return rows.map((r) => ({
      nextgenPayerName: r['NextGen Payer Name'], payerId: r['Sohar payerId'], payerName: r['Sohar payerName'],
      confidence: r['Confidence'], notes: r['Notes'],
    }));
  }

  async function readSoharPayers() {
    const rows = await readTable(TABLES.sohar);
    return rows.map((r) => ({ payerId: String(r.payerId || ''), payerName: String(r.payerName || ''), aliasPayerNames: String(r.aliasPayerNames || '') }));
  }

  async function upsertPayerMap(entry, opts) { return upsertPayerMapMany([entry], opts); }

  /**
   * Add or update Payer Map rows by NextGen name in one pass (auto-mapping can add dozens at once).
   * Notes: the new note goes first and earlier notes are kept after " | " (core.mergeNotes, ~500 characters).
   * Last reviewed: stamped with today only for a person's write (opts.human, or entry.human); automatic writes keep
   * the row's value (new automatic rows leave it blank). It is always written as text MM/DD/YYYY; an older date
   * serial on a row that is rewritten becomes that text too. Other columns on an updated row are kept.
   */
  async function upsertPayerMapMany(entries, opts) {
    if (!entries || !entries.length) return 0;
    const C = root.SoharCore;
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItem(TABLES.payerMap);
      const h = t.getHeaderRowRange(); h.load('values');
      const b = t.getDataBodyRange(); b.load('values,rowCount');
      await ctx.sync();
      const headers = h.values[0].map((x) => String(x).trim());
      const col = (name) => headers.indexOf(name);
      const get = (vals, name) => { const j = col(name); return j >= 0 ? vals[j] : ''; };
      const set = (vals, name, v) => { const j = col(name); if (j >= 0) vals[j] = v; };
      const today = C.fmtDate(new Date(Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate())));
      const asText = (v) => (typeof v === 'number' ? (v > 59 ? C.fmtDate(C.excelSerialToDate(v)) : '') : v instanceof Date ? C.fmtDate(v) : String(v ?? '').trim());
      const keyOf = (v) => String(v ?? '').trim().toLowerCase();
      const existing = new Map();
      const body = b.values.map((r) => r.slice()); // local copy: rows written in this batch are updated here
      body.forEach((r, i) => { const k = keyOf(r[col('NextGen Payer Name')]); if (k && !existing.has(k)) existing.set(k, i); });
      let blank = b.rowCount === 1 && body[0].every((v) => v === '' || v == null) ? 0 : -1;
      // Sohar payerIds like 01111 must stay text, or Excel stores 1111 and verification gets the wrong payerId.
      // Last reviewed is text too, so the column never mixes date serials and text.
      const textCols = ['Sohar payerId', 'Last reviewed'].map(col).filter((j) => j >= 0);
      const writeRow = (i, rowVals) => {
        const rg = b.getRow(i);
        for (const j of textCols) rg.getCell(0, j).numberFormat = [['@']];
        rg.values = [rowVals];
      };
      const adds = [], addAt = new Map();
      for (const entry of entries) {
        const i = existing.get(keyOf(entry.nextgenPayerName));
        const a = addAt.get(keyOf(entry.nextgenPayerName)); // same name twice in one batch: the later entry wins
        const prev = i != null ? body[i] : a != null ? adds[a] : null;
        const rowVals = prev ? prev.slice() : headers.map(() => '');
        const human = !!((opts && opts.human) || entry.human);
        set(rowVals, 'NextGen Payer Name', entry.nextgenPayerName);
        set(rowVals, 'Sohar payerId', entry.payerId);
        set(rowVals, 'Sohar payerName', entry.payerName);
        set(rowVals, 'Confidence', entry.confidence);
        set(rowVals, 'Action', entry.payerId && C.VERIFIABLE.has(keyOf(entry.confidence)) ? 'verification' : 'discovery');
        set(rowVals, 'Notes', C.mergeNotes(entry.notes, prev ? get(prev, 'Notes') : ''));
        set(rowVals, 'Last reviewed', human ? today : prev ? asText(get(prev, 'Last reviewed')) : '');
        if (i != null) { writeRow(i, rowVals); body[i] = rowVals; }
        else if (a != null) adds[a] = rowVals;
        else if (blank === 0) { writeRow(0, rowVals); blank = -1; existing.set(keyOf(entry.nextgenPayerName), 0); body[0] = rowVals; }
        else { addAt.set(keyOf(entry.nextgenPayerName), adds.length); adds.push(rowVals); }
      }
      if (adds.length) {
        t.rows.add(null, adds);
        await ctx.sync();
        if (textCols.length) {
          // Re-write the new rows' text columns as text (rows.add may have turned "01111" into 1111).
          const nb = t.getDataBodyRange(); nb.load('rowCount');
          await ctx.sync();
          for (const j of textCols) {
            const cells = nb.getCell(nb.rowCount - adds.length, j).getResizedRange(adds.length - 1, 0);
            cells.numberFormat = adds.map(() => ['@']);
            cells.values = adds.map((r) => [r[j]]);
          }
        }
      }
      await ctx.sync();
      return entries.length;
    });
  }

  /**
   * Replace all body rows of a template table with `rows` (array of objects keyed by column header).
   * Not fail-safe (rewritten in place); these tables are rebuilt from the Import sheet on every refresh.
   */
  async function replaceTable(tableName, cols, rows) {
    return replaceTableInPlace(tableName, cols, rows).catch((e) => {
      if (!/is missing|columns; this version/.test(e.message)) e.message = `Writing ${tableName} failed: ${e.message}. That table may be incomplete; click refresh (⟳) to rebuild it.`;
      throw e;
    });
  }
  async function replaceTableInPlace(tableName, cols, rows) {
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItemOrNullObject(tableName);
      t.load('isNullObject');
      await ctx.sync();
      if (t.isNullObject) throw new Error(`Table ${tableName} is missing. Use the Sohar Eligibility template workbook.`);
      const ws = t.worksheet;
      t.clearFilters(); // a filtered table can't have its rows deleted and shifted reliably
      const hdr = t.getHeaderRowRange(); hdr.load('rowIndex,columnIndex,columnCount');
      const body = t.getDataBodyRange(); body.load('rowCount');
      await ctx.sync();
      if (hdr.columnCount !== cols.length) throw new Error(`${tableName} has ${hdr.columnCount} columns; this version of the add-in expects ${cols.length}. Use the matching template.`);
      // Same columns, older names (v0.12: Review "Per Nbr" is "MRN"): rename the header cells in place.
      hdr.load('values');
      await ctx.sync();
      if (hdr.values[0].some((h, j) => String(h).trim() !== cols[j])) { hdr.values = [cols]; await ctx.sync(); }
      ctx.application.suspendApiCalculationUntilNextSync();
      // Empty the body down to one blank row. Excel refuses to delete a table's last row (InsertDeleteConflict),
      // so the first row is cleared and only rows 2..n are deleted.
      body.clear(Excel.ClearApplyTo.contents);
      if (body.rowCount > 1) body.getRow(1).getBoundingRect(body.getLastRow()).delete(Excel.DeleteShiftDirection.up);
      await ctx.sync();
      // Everything is written as text so member IDs and ZIPs keep leading zeros.
      const values = rows.map((r) => cols.map((c) => { const v = r[c]; return v == null ? '' : v; }));
      if (!values.length) { t.getDataBodyRange().numberFormat = [cols.map(() => '@')]; await ctx.sync(); return; }
      const canResize = Office.context.requirements.isSetSupported('ExcelApi', '1.13');
      if (canResize) {
        const target = ws.getRangeByIndexes(hdr.rowIndex + 1, hdr.columnIndex, values.length, cols.length);
        target.numberFormat = values.map((r) => r.map(() => '@'));
        await ctx.sync();
        for (let i = 0; i < values.length; i += CHUNK) {
          ctx.application.suspendApiCalculationUntilNextSync();
          const chunk = values.slice(i, i + CHUNK);
          ws.getRangeByIndexes(hdr.rowIndex + 1 + i, hdr.columnIndex, chunk.length, cols.length).values = chunk;
          await ctx.sync();
        }
        t.resize(ws.getRangeByIndexes(hdr.rowIndex, hdr.columnIndex, values.length + 1, cols.length));
        await ctx.sync();
      } else {
        // Older Excel: format the blank row first so appended rows inherit text format, then append and drop the blank.
        t.getDataBodyRange().numberFormat = [cols.map(() => '@')]; // one blank row, one '@' per column
        await ctx.sync();
        for (let i = 0; i < values.length; i += CHUNK) {
          ctx.application.suspendApiCalculationUntilNextSync();
          t.rows.add(null, values.slice(i, i + CHUNK));
          await ctx.sync();
        }
        const after = t.getDataBodyRange(); after.load('rowCount');
        await ctx.sync();
        if (after.rowCount > values.length) after.getRow(0).delete(Excel.DeleteShiftDirection.up);
        await ctx.sync();
      }
    });
  }

  // ---------- file upload ----------
  const canInsertWorkbook = () => Office.context.requirements.isSetSupported('ExcelApi', '1.13');

  /** Replace the Import sheet with CSV rows, written as text so IDs and ZIPs keep leading zeros. */
  async function writeImportRows(rows) {
    const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
    const grid = rows.map((r) => { const o = r.slice(); while (o.length < width) o.push(''); return o; });
    const failed = (s) => (e) => { e.message = `Could not load the report (${s}): ${e.message}. The Import sheet was not changed.`; throw e; };
    return run(async (ctx) => {
      const wb = ctx.workbook, ws = wb.worksheets.getItem(SHEETS.import);
      if (grid.length && width && Office.context.requirements.isSetSupported('ExcelApi', '1.9')) {
        // Fail-safe: rows go to a hidden staging sheet first, then onto Import in one batch (clear + server-side
        // copy), so a cut-off upload leaves Import untouched, empty or complete, never a partial report to build from.
        const old = wb.worksheets.getItemOrNullObject(STAGE_IMPORT); old.load('isNullObject');
        await ctx.sync().catch(failed('checking the workbook'));
        if (!old.isNullObject) old.delete();
        const stage = wb.worksheets.add(STAGE_IMPORT);
        stage.visibility = Excel.SheetVisibility.hidden;
        await ctx.sync().catch(failed('preparing'));
        for (const [i, chunk] of writeChunks(grid)) {
          ctx.application.suspendApiCalculationUntilNextSync();
          const rg = stage.getRangeByIndexes(i, 0, chunk.length, width);
          rg.numberFormat = chunk.map((r) => r.map(() => '@'));
          rg.values = chunk;
          await ctx.sync().catch(failed(`rows ${i + 1}-${i + chunk.length}`));
        }
        ws.getRange().clear();
        ws.getRangeByIndexes(0, 0, grid.length, width).copyFrom(stage.getRangeByIndexes(0, 0, grid.length, width), Excel.RangeCopyType.all);
        await ctx.sync().catch((e) => { e.message = `Could not load the report: ${e.message}. Upload it again.`; throw e; });
        stage.delete();
        await ctx.sync().catch((e) => console.warn('report staging sheet left', e));
        return;
      }
      ws.getRange().clear();
      await ctx.sync();
      if (!grid.length || !width) return;
      for (let i = 0; i < grid.length; i += CHUNK) {
        ctx.application.suspendApiCalculationUntilNextSync();
        const chunk = grid.slice(i, i + CHUNK);
        const rg = ws.getRangeByIndexes(i, 0, chunk.length, width);
        rg.numberFormat = chunk.map((r) => r.map(() => '@'));
        rg.values = chunk;
        await ctx.sync();
      }
    });
  }

  /** Insert an .xlsx (base64) as temporary sheets, hand the used range of the first (or the one matching prefer) to fn, then remove them. */
  async function withTempSheet(base64, fn, prefer) {
    if (!canInsertWorkbook()) throw new Error('This version of Excel cannot open .xlsx files from the pane. Save the file as CSV and upload that, or paste it into the sheet.');
    return run(async (ctx) => {
      const inserted = ctx.workbook.insertWorksheetsFromBase64(base64, { positionType: Excel.WorksheetPositionType.end });
      await ctx.sync();
      const ids = inserted.value || []; // worksheet IDs, not names
      if (!ids.length) throw new Error('That workbook has no sheets.');
      try {
        const sheets = ids.map((id) => { const s = ctx.workbook.worksheets.getItem(id); s.load('name'); return s; });
        await ctx.sync();
        const pick = (prefer && sheets.find((s) => prefer.test(s.name))) || sheets[0];
        const used = pick.getUsedRangeOrNullObject(true);
        used.load('isNullObject,rowCount,columnCount');
        await ctx.sync();
        return await fn(ctx, used);
      } finally {
        for (const id of ids) ctx.workbook.worksheets.getItem(id).delete();
        await ctx.sync();
      }
    });
  }

  /** Copy an uploaded appointment report (.xlsx) onto the Import sheet, formats included, as if pasted. */
  async function importXlsx(base64) {
    return withTempSheet(base64, async (ctx, used) => {
      const ws = ctx.workbook.worksheets.getItem(SHEETS.import);
      ws.getRange().clear();
      if (used.isNullObject) { await ctx.sync(); return; }
      ws.getRangeByIndexes(0, 0, used.rowCount, used.columnCount).copyFrom(used, Excel.RangeCopyType.all);
      await ctx.sync();
    });
  }

  /** Values of the first sheet of an uploaded .xlsx (or the sheet whose name matches prefer, e.g. the IR Portal's Filtered tab). */
  async function readXlsxRows(base64, prefer) {
    return withTempSheet(base64, async (ctx, used) => {
      if (used.isNullObject) return [];
      used.load('rowIndex,columnIndex');
      await ctx.sync();
      const rows = [];
      for (let i = 0; i < used.rowCount; i += 5000) { // chunked: Excel on the web caps a response at 5MB
        const rg = used.worksheet.getRangeByIndexes(used.rowIndex + i, used.columnIndex, Math.min(5000, used.rowCount - i), used.columnCount);
        rg.load('values');
        await ctx.sync();
        for (const r of rg.values) rows.push(r);
      }
      return rows;
    }, prefer);
  }

  const writeResults = (cols, rows) => writeAddinTable(SHEETS.results, RESULTS_TABLE, cols, rows, true);
  const writeTasks = (cols, rows) => writeAddinTable(SHEETS.tasks, TASKS_TABLE, cols, rows, false);
  // Excel on the web cannot download from the pane, so files go to an Export sheet (File > Export > Download as CSV).
  const writeExport = (cols, rows) => writeAddinTable('Export', 'tblExport', cols, rows, true);
  const writeIr = (cols, rows) => writeAddinTable('IR Cross-Ref', 'tblIrCrossRef', cols, rows, true);
  // Benefits: one row per verified patient, OOP closest to met first. Percent columns stay numbers so they show as % with data bars.
  const writeBenefits = (cols, rows, pctCols) => writeAddinTable('Benefits', 'tblBenefits', cols, rows, false, { pctCols });
  async function readIr() { return readOwned('tblIrCrossRef'); }
  // Discovery Verify: import-ready verification rows for patients discovery found coverage for. Admins can type member IDs here.
  const writeDiscoveryVerify = (cols, rows) => writeAddinTable('Discovery Verify', 'tblDiscoveryVerify', cols, rows, false);
  async function readDiscoveryVerify() { return readOwned('tblDiscoveryVerify'); }
  // Discovery from Verify: import-ready discovery rows for patients whose verification failed. Admins can type ZIPs here.
  const writeVerifyFailed = (cols, rows) => writeAddinTable('Discovery from Verify', 'tblVerifyFailed', cols, rows, false);
  async function readVerifyFailed() { return readOwned('tblVerifyFailed'); }
  // Every Sohar result uploaded so far, one JSON record per row, on a hidden sheet.
  async function readUploads() { return readOwned('tblSoharUploads'); }
  async function writeUploads(cols, rows) {
    await writeAddinTable('Sohar Uploads', 'tblSoharUploads', cols, rows, false); // hidden: OWNED
  }
  // Insurance-related IR Portal items from the last IR upload, so the cross-check re-runs on open. Hidden sheet.
  async function readIrUploads() {
    return (await readOwned('tblIrUploads')).map((r) => { try { return JSON.parse(r.Record); } catch (e) { return null; } }).filter(Boolean);
  }
  async function writeIrUploads(records) {
    await writeAddinTable('IR Uploads', 'tblIrUploads', ['Record'], records.map((r) => ({ Record: JSON.stringify(r) })), false); // hidden: OWNED
  }
  /** Rows in a table the add-in owns, or 0 if it does not exist yet. */
  async function tableRowCount(tableName) {
    return run(async (ctx) => {
      const t = await settle(ctx, tableName);
      if (!t) return 0;
      const b = t.getDataBodyRange(); b.load('values');
      await ctx.sync();
      return b.values.filter((r) => r.some((v) => v !== '' && v != null)).length;
    });
  }
  async function readTasks() { return readOwned(TASKS_TABLE); }

  // Payer Evidence: one row per patient-insurance observation, upserted by Key each refresh (rules.payerEvidence).
  async function readPayerEvidence() { return readOwned('tblPayerEvidence'); }
  const writePayerEvidence = (cols, rows) => writeAddinTable('Payer Evidence', 'tblPayerEvidence', cols, rows, false);

  // Payer learning (v0.10). The ledger and the log are append-only: callers pass every earlier row plus the new ones
  // (core.appendOutcomes / log rows appended), and each save replaces the sheet fail-safe like the others.
  async function readPayerOutcomes() { return readOwned('tblPayerOutcomes'); }
  const writePayerOutcomes = (cols, rows) => writeAddinTable('Payer Outcomes', 'tblPayerOutcomes', cols, rows, false); // hidden: OWNED
  async function readPayerMapLog() { return readOwned('tblPayerMapLog'); }
  const writePayerMapLog = (cols, rows) => writeAddinTable('Payer Map Log', 'tblPayerMapLog', cols, rows, false);
  /** Read the Payer Map Log and save it with rows added at the end. Returns the number added. */
  async function appendPayerMapLog(cols, rows) {
    if (!rows || !rows.length) return 0;
    const prev = await readPayerMapLog();
    await writePayerMapLog(cols, prev.concat(rows));
    return rows.length;
  }
  const writePayerLearning = (cols, rows) => writeAddinTable('Payer Learning', 'tblPayerLearning', cols, rows, false);

  // Sohar ref memory (core.SOHAR_REF_COLS): append-only through core.appendRefs, so callers pass every earlier row.
  async function readSoharRefs() { return readOwned('tblSoharRefs'); }
  const writeSoharRefs = (cols, rows) => writeAddinTable('Sohar Refs', 'tblSoharRefs', cols, rows, false); // hidden: OWNED
  // "Re-verify anyway" requests (core.REVERIFY_REQUEST_COLS), pruned by core.pruneReverifyRequests on each build.
  async function readReverifyRequests() { return readOwned('tblReverifyRequests'); }
  const writeReverifyRequests = (cols, rows) => writeAddinTable('Re-verify Requests', 'tblReverifyRequests', cols, rows, false); // hidden: OWNED

  // v0.12: guarantor resubmission requests, payer tests (hidden), the payer mapping discovery rows, metrics and themes.
  const readGuarantorRequests = () => readOwned('tblGuarantorRequests');
  const writeGuarantorRequests = (cols, rows) => writeAddinTable('Guarantor Requests', 'tblGuarantorRequests', cols, rows, false); // hidden: OWNED
  const readPayerTests = () => readOwned('tblPayerTests');
  const writePayerTests = (cols, rows) => writeAddinTable('Payer Tests', 'tblPayerTests', cols, rows, false); // hidden: OWNED
  const writePayerMapping = (cols, rows) => writeAddinTable('Payer Mapping', 'tblPayerMapping', cols, rows, false);
  const readPayerMapping = () => readOwned('tblPayerMapping');
  const writePayerMetrics = (cols, rows) => writeAddinTable('Payer Metrics', 'tblPayerMetrics', cols, rows, false);
  const readPayerMetricHistory = () => readOwned('tblPayerMetricHistory');
  const writePayerMetricHistory = (cols, rows) => writeAddinTable('Payer Metrics History', 'tblPayerMetricHistory', cols, rows, false);
  const writeGuarantorThemes = (cols, rows) => writeAddinTable('Guarantor Themes', 'tblGuarantorThemes', cols, rows, false);
  const readClinicRequests = () => readOwned('tblClinicRequests');
  // Payer Credential Rules: missing sheet = [] (the rule is silent until the CSV is imported).
  const readPayerCredRules = () => readOwned('tblPayerCredRules');
  const writePayerCredRules = (cols, rows) => writeAddinTable('Payer Credential Rules', 'tblPayerCredRules', cols, rows, false);
  const writeClinicRequests = (cols, rows) => writeAddinTable('Clinic Requests', 'tblClinicRequests', cols, rows, false);
  // Clinic Notes / Clinic Payer Notes (rules.CLINIC_NOTE_COLS / CLINIC_PAYER_NOTE_COLS): missing sheet = none yet.
  const readClinicNotes = () => readOwned('tblClinicNotes');
  const writeClinicNotes = (cols, rows) => writeAddinTable('Clinic Notes', 'tblClinicNotes', cols, rows, false);
  const readClinicPayerNotes = () => readOwned('tblClinicPayerNotes');
  const writeClinicPayerNotes = (cols, rows) => writeAddinTable('Clinic Payer Notes', 'tblClinicPayerNotes', cols, rows, false);
  const readClinicWorkApplied = () => readOwned('tblClinicWorkApplied'); // rules.CLINIC_APPLIED_COLS
  const writeClinicWorkApplied = (cols, rows) => writeAddinTable('Clinic Work Applied', 'tblClinicWorkApplied', cols, rows, false); // hidden: OWNED

  // Fee estimates (v0.11): read / write by short name (FEE_TABLES). Missing table = nothing saved yet (readOwned).
  function feeTable(kind) {
    const t = FEE_TABLES[kind];
    if (!t) throw new Error(`Unknown fee table ${kind}`);
    return t;
  }
  async function readFee(kind) { return readOwned(feeTable(kind)); }
  async function writeFee(kind, cols, rows) { const t = feeTable(kind); return writeAddinTable(OWNED[t].sheet, t, cols, rows, false); }

  /** True when the workbook has this table (used to tell the Sohar template from a clinic file). */
  async function tableExists(tableName) {
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItemOrNullObject(tableName);
      t.load('isNullObject');
      await ctx.sync();
      return !t.isNullObject;
    });
  }

  /** A clinic file saved by the Admin tab: {meta, rows, review}, or null when this is not one. */
  async function readClinicFile() {
    if (!(await tableExists('tblClinicMeta'))) return null;
    const meta = (await readTable('tblClinicMeta'))[0] || {};
    const rows = await readTable('tblClinic').catch(() => []);
    const review = await readTable('tblClinicReview').catch(() => []);
    const requests = await readOwned('tblClinicRequests').catch(() => []);
    return { meta, rows, review, requests };
  }

  /** Mark one NextGen task done (or not) by Task Key. */
  async function setTaskDone(key, done) {
    return run(async (ctx) => {
      const t = await settle(ctx, TASKS_TABLE);
      if (!t) return;
      const h = t.getHeaderRowRange(); h.load('values');
      const b = t.getDataBodyRange(); b.load('values');
      await ctx.sync();
      const hd = h.values[0].map(String), kc = hd.indexOf('Task Key'), dc = hd.indexOf('Done');
      const i = b.values.findIndex((r) => String(r[kc]) === key);
      if (i >= 0 && dc >= 0) { b.getCell(i, dc).values = [[done ? 'Yes' : '']]; await ctx.sync(); }
    });
  }

  /** Row blocks for writing: at most CHUNK rows and about 2MB each (Excel on the web rejects a request over ~5MB). */
  function writeChunks(all) {
    const out = [];
    for (let i = 0; i < all.length;) {
      let j = i, bytes = 0;
      while (j < all.length && j - i < CHUNK && (j === i || bytes < 2e6)) bytes += JSON.stringify(all[j++]).length;
      out.push([i, all.slice(i, j)]);
      i = j;
    }
    return out;
  }

  /**
   * Create (or rebuild) a sheet + table the add-in owns (OWNED). Fail-safe: the new rows are written to a hidden
   * staging sheet and become a table only when every row is in; the old sheet is replaced in one small swap at the
   * end. If anything fails before the swap, the old sheet is untouched. If the swap is cut off, the complete new
   * table is on the staging sheet and the next read, save or pane open finishes it (rescue). Writes cells first and
   * makes the table in one step: no table resize (that path left Sohar Results empty in Excel).
   */
  async function writeAddinTable(sheetName, tableName, cols, rows, activate, opts) {
    const own = OWNED[tableName] || { sheet: sheetName };
    const pct = new Set((opts && opts.pctCols) || []);
    const values = rows.map((r) => cols.map((c) => { const v = r[c]; return v == null ? '' : v; }));
    const fmt = (r, hdr) => r.map((v, j) => (!hdr && pct.has(cols[j]) ? '0%' : '@'));
    const unchanged = (s) => (e) => { e.message = `Could not save the ${sheetName} sheet (${s}): ${e.message}. The ${sheetName} sheet was not changed; its previous data is still there.`; throw e; };
    await run(async (ctx) => {
      const wb = ctx.workbook;
      const probe = () => {
        const p = { real: wb.tables.getItemOrNullObject(tableName), stg: wb.tables.getItemOrNullObject(STG(tableName)),
          stage: wb.worksheets.getItemOrNullObject(STAGE_SHEET(tableName)), ws: wb.worksheets.getItemOrNullObject(sheetName) };
        p.real.load('isNullObject'); p.stg.load('isNullObject'); p.stage.load('isNullObject,id'); p.ws.load('isNullObject,id,position,visibility');
        return p;
      };
      let p = probe();
      const active = wb.worksheets.getActiveWorksheet(); active.load('id');
      const prot = Office.context.requirements.isSetSupported('ExcelApi', '1.7') ? wb.protection : null;
      if (prot) prot.load('protected');
      await ctx.sync().catch(unchanged('checking the workbook'));
      // Saves add, rename and delete sheets, which a structure-protected workbook refuses.
      if (prot && prot.protected) unchanged('checking the workbook')(new Error('the workbook structure is protected; unprotect it (Review > Protect Workbook) and try again'));
      // An earlier save was cut off mid-swap: its complete rows are the only copy. Finish it before anything else.
      if (p.real.isNullObject && !p.stg.isNullObject && OWNED[tableName]) {
        if (!(await rescue(ctx, tableName))) throw new Error(`Could not save the ${sheetName} sheet: an earlier save of it was cut off and could not be finished. Nothing was overwritten: its rows are safe on the hidden sheet "${STAGE_SHEET(tableName)}". Close and reopen the pane.`);
        recovered.add(sheetName);
        p = probe();
        await ctx.sync().catch(unchanged('checking the workbook'));
      }
      const wasActive = !p.ws.isNullObject && active.id === p.ws.id; // deleting it would bounce the user to a neighbour
      let oldTableSheetId = null;
      if (!p.real.isNullObject) { const w = p.real.worksheet; w.load('id'); await ctx.sync().catch(unchanged('checking the workbook')); oldTableSheetId = w.id; }

      // Leftovers from an unfinished save are partial or superseded (the real table is intact): sweep them, by name
      // in this same batch (see sweepStage), so a sheet another window has just swapped into place is never hit.
      if (!p.stage.isNullObject && p.stage.id !== oldTableSheetId) {
        if (!p.real.isNullObject) sweepStage(ctx, tableName);
        else wb.worksheets.getItem(STAGE_SHEET(tableName)).delete(); // first save ever: nothing real to protect
      }
      // A stray STG table off the staging sheet (the real table is intact here) would block the name: drop it.
      else if (!p.stg.isNullObject) wb.tables.getItem(STG(tableName)).delete();
      const stage = wb.worksheets.add(STAGE_SHEET(tableName));
      stage.visibility = Excel.SheetVisibility.hidden;
      stage.load('id');
      await ctx.sync().catch(unchanged('preparing'));
      const stageId = stage.id;

      const all = [cols, ...values];
      for (const [i, chunk] of writeChunks(all)) {
        ctx.application.suspendApiCalculationUntilNextSync();
        const rg = stage.getRangeByIndexes(i, 0, chunk.length, cols.length);
        rg.numberFormat = chunk.map((r, k) => fmt(r, i + k === 0)); // text, so IDs and ZIPs keep leading zeros; % columns stay numbers
        rg.values = chunk;
        await ctx.sync().catch(unchanged(`rows ${i + 1}-${i + chunk.length}`));
      }
      // The completeness marker: only now does a table named STG exist.
      stage.tables.add(stage.getRangeByIndexes(0, 0, Math.max(2, all.length), cols.length), true).name = STG(tableName);
      await ctx.sync().catch(unchanged('making the table'));

      // The swap, in an order where any cut-off point leaves either the old table intact or a complete STG table.
      const vis = own.hidden ? Excel.SheetVisibility.hidden : p.ws.isNullObject ? Excel.SheetVisibility.visible : p.ws.visibility;
      stage.visibility = vis; // before the delete, so a visible sheet always remains
      if (!p.real.isNullObject && oldTableSheetId !== (p.ws.isNullObject ? null : p.ws.id)) p.real.delete(); // old table on a renamed sheet
      if (!p.ws.isNullObject) p.ws.delete(); // takes the old table with it
      stage.name = sheetName;
      wb.tables.getItem(STG(tableName)).name = tableName;
      if (!p.ws.isNullObject) stage.position = p.ws.position;
      try { await ctx.sync(); } catch (e) {
        // Find out how far the swap got. Decide by where the real table lives, never by names alone: another
        // Excel window can be saving the same table with the same staging names.
        const later = () => { throw finishLater(sheetName, e); };
        const q = probe();
        await ctx.sync().catch(later);
        if (q.real.isNullObject && !q.stg.isNullObject && !(await rescue(ctx, tableName))) later(); // cut off mid-swap
        const now = wb.tables.getItemOrNullObject(tableName); now.load('isNullObject');
        await ctx.sync().catch(later);
        if (now.isNullObject) throw Object.assign(e, { message: `Saving the ${sheetName} sheet failed and the sheet is gone (${e.message}). Run the step again to rebuild it.` });
        const home = now.worksheet; home.load('id');
        await ctx.sync().catch(later);
        if (home.id !== stageId) {
          // The real table is not the one built here: the old rows are still in place. Drop this save's own staging
          // sheet (by id: it is ours whatever it is named now) and say the save did not happen.
          const mine = wb.worksheets.getItemOrNullObject(stageId); mine.load('isNullObject');
          await ctx.sync().catch(() => {});
          if (mine.isNullObject === false) { mine.delete(); await ctx.sync().catch(() => {}); }
          unchanged('swapping in the new sheet')(e);
        }
        // Otherwise the swap (or the rescue) put this save's table in place: the save stands.
      }

      try { // cosmetic only: never fails a save
        const ws = wb.worksheets.getItem(sheetName), tbl = wb.tables.getItem(tableName);
        ws.freezePanes.freezeRows(1);
        // Hidden sheets (uploads, the outcomes ledger) are never looked at: autofit on a big ledger is wasted time.
        if (!own.hidden) tbl.getRange().format.autofitColumns();
        if (values.length && Office.context.requirements.isSetSupported('ExcelApi', '1.6')) {
          for (const c of pct) {
            const j = cols.indexOf(c); if (j < 0) continue;
            const bar = ws.getRangeByIndexes(1, j, values.length, 1).conditionalFormats.add('DataBar');
            bar.dataBar.lowerBoundRule = { type: 'Number', formula: '0' };
            bar.dataBar.upperBoundRule = { type: 'Number', formula: '1' };
          }
        }
        if ((activate || wasActive) && vis !== Excel.SheetVisibility.hidden) ws.activate();
        await ctx.sync();
      } catch (e) { console.warn(`formatting of ${sheetName} skipped`, e); }
    });
    return values.length;
  }
  const finishLater = (sheetName, e) => {
    e.message = `Saving the ${sheetName} sheet did not finish: ${e.message}. Nothing was lost: the add-in finishes or undoes the save the next time it opens. Close and reopen the pane.`;
    return e;
  };

  /**
   * On pane open: finish any save that was cut off mid-swap and sweep unfinished staging sheets.
   * Returns the sheet names that were put back (plus any finished by reads earlier in this session).
   */
  async function recoverAll() {
    await run(async (ctx) => {
      const wb = ctx.workbook, names = Object.keys(OWNED);
      const ps = names.map((t) => {
        const p = { real: wb.tables.getItemOrNullObject(t), stg: wb.tables.getItemOrNullObject(STG(t)), stage: wb.worksheets.getItemOrNullObject(STAGE_SHEET(t)) };
        p.real.load('isNullObject'); p.stg.load('isNullObject'); p.stage.load('isNullObject');
        return p;
      });
      const imp = wb.worksheets.getItemOrNullObject(STAGE_IMPORT); imp.load('isNullObject');
      await ctx.sync();
      if (!imp.isNullObject) { imp.delete(); await ctx.sync().catch((e) => console.warn('report staging sheet left', e)); }
      for (let i = 0; i < names.length; i++) {
        const p = ps[i];
        if (p.real.isNullObject && !p.stg.isNullObject) { if (await rescue(ctx, names[i])) recovered.add(OWNED[names[i]].sheet); }
        else if (!p.stage.isNullObject) { await sweepStage(ctx, names[i]); await ctx.sync().catch((e) => console.warn('staging sheet left', e)); }
      }
    });
    return takeRecovered();
  }
  /**
   * Queue (no sync) the removal of an unfinished save's staging sheet, guarded so it can never remove the only copy.
   * Fresh proxies resolved by name when the batch runs: a sheet proxy loaded in an earlier sync follows the sheet's
   * id, so if another window's save had since renamed its staging sheet into place, it would delete the live sheet.
   * The first op fails the batch (nothing deleted) unless the real table exists at that moment.
   */
  function sweepStage(ctx, tableName) {
    ctx.workbook.tables.getItem(tableName).getRange().load('address');
    ctx.workbook.worksheets.getItem(STAGE_SHEET(tableName)).delete();
  }
  /** Sheets whose interrupted save was finished since the last call (by a read, a save or recoverAll). */
  function takeRecovered() { const out = [...recovered]; recovered.clear(); return out; }

  async function setConfigValue(setting, value) {
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItem(TABLES.config);
      const b = t.getDataBodyRange(); b.load('values');
      await ctx.sync();
      const i = b.values.findIndex((r) => String(r[0]).trim() === setting);
      if (i >= 0) b.getCell(i, 1).values = [[value]];
      else t.rows.add(null, [[setting, value, '']]);
      await ctx.sync();
    });
  }

  async function activate(sheetName) {
    return run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      ws.load('isNullObject');
      await ctx.sync();
      if (!ws.isNullObject) { ws.activate(); await ctx.sync(); }
    });
  }

  async function templateCheck() {
    return run(async (ctx) => {
      const names = Object.values(TABLES);
      const items = names.map((n) => { const t = ctx.workbook.tables.getItemOrNullObject(n); t.load('isNullObject'); return t; });
      await ctx.sync();
      return names.filter((n, i) => items[i].isNullObject);
    });
  }

  root.SoharExcel = { SHEETS, TABLES, readImportRows, readTable, readConfig, readPayerMap, readSoharPayers, upsertPayerMap, upsertPayerMapMany, replaceTable,
    readPayerEvidence, writePayerEvidence, tableExists, readClinicFile, ensureConfigRows,
    readPayerOutcomes, writePayerOutcomes, readPayerMapLog, writePayerMapLog, appendPayerMapLog, writePayerLearning,
    readSoharRefs, writeSoharRefs, readReverifyRequests, writeReverifyRequests, FEE_TABLES, readFee, writeFee,
    readGuarantorRequests, writeGuarantorRequests, readPayerTests, writePayerTests, writePayerMapping, readPayerMapping, writePayerMetrics,
    readPayerMetricHistory, writePayerMetricHistory, writeGuarantorThemes, readClinicRequests, writeClinicRequests, readPayerCredRules, writePayerCredRules,
    readClinicNotes, writeClinicNotes, readClinicPayerNotes, writeClinicPayerNotes, readClinicWorkApplied, writeClinicWorkApplied,
    writeImportRows, importXlsx, readXlsxRows, writeResults, writeTasks, writeExport, writeIr, readIr, readTasks,
    writeDiscoveryVerify, readDiscoveryVerify, writeVerifyFailed, readVerifyFailed,writeBenefits, readUploads, writeUploads, readIrUploads, writeIrUploads, tableRowCount, setTaskDone, recoverAll, takeRecovered, canInsertWorkbook, setConfigValue, activate, templateCheck };
})(typeof self !== 'undefined' ? self : this);
