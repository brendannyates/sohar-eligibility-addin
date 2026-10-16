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

  /** Table rows as objects. Throws if the table is missing, unless missingOk (then []). Other errors always throw. */
  async function readTable(tableName, missingOk) {
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItemOrNullObject(tableName);
      t.load('isNullObject');
      await ctx.sync();
      if (t.isNullObject) {
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
      clinicLocation: String(cfg['Clinic location'] ?? '').trim(),
    };
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

  async function upsertPayerMap(entry) { return upsertPayerMapMany([entry]); }

  /** Add or update Payer Map rows by NextGen name in one pass (auto-mapping can add dozens at once). */
  async function upsertPayerMapMany(entries) {
    if (!entries || !entries.length) return 0;
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItem(TABLES.payerMap);
      const h = t.getHeaderRowRange(); h.load('values');
      const b = t.getDataBodyRange(); b.load('values,rowCount');
      await ctx.sync();
      const headers = h.values[0].map((x) => String(x).trim());
      const col = (name) => headers.indexOf(name);
      const set = (vals, name, v) => { const j = col(name); if (j >= 0) vals[j] = v; };
      const today = root.SoharCore.fmtDate(new Date(Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate())));
      const verifiable = root.SoharCore.VERIFIABLE;
      const keyOf = (v) => String(v ?? '').trim().toLowerCase();
      const existing = new Map();
      b.values.forEach((r, i) => { const k = keyOf(r[col('NextGen Payer Name')]); if (k && !existing.has(k)) existing.set(k, i); });
      let blank = b.rowCount === 1 && b.values[0].every((v) => v === '' || v == null) ? 0 : -1;
      // Sohar payerIds like 01111 must stay text, or Excel stores 1111 and verification gets the wrong payerId.
      const idCol = col('Sohar payerId');
      const writeRow = (i, rowVals) => {
        const rg = b.getRow(i);
        if (idCol >= 0) rg.getCell(0, idCol).numberFormat = [['@']];
        rg.values = [rowVals];
      };
      const adds = [];
      for (const entry of entries) {
        const rowVals = headers.map(() => '');
        set(rowVals, 'NextGen Payer Name', entry.nextgenPayerName);
        set(rowVals, 'Sohar payerId', entry.payerId);
        set(rowVals, 'Sohar payerName', entry.payerName);
        set(rowVals, 'Confidence', entry.confidence);
        set(rowVals, 'Action', entry.payerId && verifiable.has(keyOf(entry.confidence)) ? 'verification' : 'discovery');
        set(rowVals, 'Notes', entry.notes || '');
        set(rowVals, 'Last reviewed', today);
        const i = existing.get(keyOf(entry.nextgenPayerName));
        if (i != null) writeRow(i, rowVals);
        else if (blank === 0) { writeRow(0, rowVals); blank = -1; existing.set(keyOf(entry.nextgenPayerName), 0); }
        else adds.push(rowVals);
      }
      if (adds.length) {
        t.rows.add(null, adds);
        await ctx.sync();
        if (idCol >= 0) {
          // Re-write the new rows' payerIds as text (rows.add may have turned "01111" into 1111).
          const nb = t.getDataBodyRange(); nb.load('rowCount');
          await ctx.sync();
          const ids = nb.getCell(nb.rowCount - adds.length, idCol).getResizedRange(adds.length - 1, 0);
          ids.numberFormat = adds.map(() => ['@']);
          ids.values = adds.map((r) => [r[idCol]]);
        }
      }
      await ctx.sync();
      return entries.length;
    });
  }

  /** Replace all body rows of a table with `rows` (array of objects keyed by column header). */
  async function replaceTable(tableName, cols, rows) {
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
    return run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(SHEETS.import);
      ws.getRange().clear();
      await ctx.sync();
      const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
      if (!rows.length || !width) return;
      const grid = rows.map((r) => { const o = r.slice(); while (o.length < width) o.push(''); return o; });
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
    await writeAddinTable('Sohar Uploads', 'tblSoharUploads', cols, rows, false);
    await run(async (ctx) => { ctx.workbook.worksheets.getItem('Sohar Uploads').visibility = Excel.SheetVisibility.hidden; await ctx.sync(); }).catch(() => {});
  }
  // Insurance-related IR Portal items from the last IR upload, so the cross-check re-runs on open. Hidden sheet.
  async function readIrUploads() {
    return (await readOwned('tblIrUploads')).map((r) => { try { return JSON.parse(r.Record); } catch (e) { return null; } }).filter(Boolean);
  }
  async function writeIrUploads(records) {
    await writeAddinTable('IR Uploads', 'tblIrUploads', ['Record'], records.map((r) => ({ Record: JSON.stringify(r) })), false);
    await run(async (ctx) => { ctx.workbook.worksheets.getItem('IR Uploads').visibility = Excel.SheetVisibility.hidden; await ctx.sync(); }).catch(() => {});
  }
  /** Rows in a table the add-in owns, or 0 if it does not exist yet. */
  async function tableRowCount(tableName) {
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItemOrNullObject(tableName);
      t.load('isNullObject');
      await ctx.sync();
      if (t.isNullObject) return 0;
      const b = t.getDataBodyRange(); b.load('values');
      await ctx.sync();
      return b.values.filter((r) => r.some((v) => v !== '' && v != null)).length;
    });
  }
  async function readTasks() { return readOwned(TASKS_TABLE); }

  // Payer Evidence: one row per patient-insurance observation, upserted by Key each refresh (rules.payerEvidence).
  async function readPayerEvidence() { return readOwned('tblPayerEvidence'); }
  const writePayerEvidence = (cols, rows) => writeAddinTable('Payer Evidence', 'tblPayerEvidence', cols, rows, false);

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
    return { meta, rows, review };
  }

  /** Mark one NextGen task done (or not) by Task Key. */
  async function setTaskDone(key, done) {
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItemOrNullObject(TASKS_TABLE);
      t.load('isNullObject');
      await ctx.sync();
      if (t.isNullObject) return;
      const h = t.getHeaderRowRange(); h.load('values');
      const b = t.getDataBodyRange(); b.load('values');
      await ctx.sync();
      const hd = h.values[0].map(String), kc = hd.indexOf('Task Key'), dc = hd.indexOf('Done');
      const i = b.values.findIndex((r) => String(r[kc]) === key);
      if (i >= 0 && dc >= 0) { b.getCell(i, dc).values = [[done ? 'Yes' : '']]; await ctx.sync(); }
    });
  }

  /**
   * Create (or rebuild) a sheet + table the add-in owns. Not part of the template.
   * Writes every row as plain cells first, then turns the block into a table in one step,
   * so there is no table resize (that path left Sohar Results empty in Excel).
   */
  async function writeAddinTable(sheetName, tableName, cols, rows, activate, opts) {
    const pct = new Set((opts && opts.pctCols) || []);
    const values = rows.map((r) => cols.map((c) => { const v = r[c]; return v == null ? '' : v; }));
    const fmt = (r, hdr) => r.map((v, j) => (!hdr && pct.has(cols[j]) ? '0%' : '@'));
    const stage = (s) => (e) => { e.message = `Writing the ${sheetName} sheet (${s}): ${e.message}`; throw e; };
    await run(async (ctx) => {
      const old = ctx.workbook.tables.getItemOrNullObject(tableName);
      old.load('isNullObject');
      let ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      ws.load('isNullObject');
      await ctx.sync();
      if (!old.isNullObject) old.delete();
      const existed = !ws.isNullObject;
      if (existed) ws.tables.load('items/name');
      else ws = ctx.workbook.worksheets.add(sheetName);
      await ctx.sync().catch(stage('preparing the sheet'));
      // clear() leaves tables in place; a renamed or extra table on this sheet would make tables.add overlap.
      if (existed) for (const t of ws.tables.items) t.delete();
      ws.getRange().clear();
      await ctx.sync().catch(stage('clearing'));

      const all = [cols, ...values];
      for (let i = 0; i < all.length; i += CHUNK) {
        ctx.application.suspendApiCalculationUntilNextSync();
        const chunk = all.slice(i, i + CHUNK);
        const rg = ws.getRangeByIndexes(i, 0, chunk.length, cols.length);
        rg.numberFormat = chunk.map((r, k) => fmt(r, i + k === 0)); // text, so IDs and ZIPs keep leading zeros; % columns stay numbers
        rg.values = chunk;
        await ctx.sync().catch(stage(`rows ${i + 1}-${i + chunk.length}`));
      }
      const tbl = ws.tables.add(ws.getRangeByIndexes(0, 0, Math.max(2, all.length), cols.length), true);
      tbl.name = tableName;
      await ctx.sync().catch(stage('making the table'));
      ws.freezePanes.freezeRows(1);
      tbl.getRange().format.autofitColumns();
      if (values.length && Office.context.requirements.isSetSupported('ExcelApi', '1.6')) {
        try { // cosmetic only: never let data bars break a save
          for (const c of pct) {
            const j = cols.indexOf(c); if (j < 0) continue;
            const bar = ws.getRangeByIndexes(1, j, values.length, 1).conditionalFormats.add('DataBar');
            bar.dataBar.lowerBoundRule = { type: 'Number', formula: '0' };
            bar.dataBar.upperBoundRule = { type: 'Number', formula: '1' };
          }
        } catch (e) { console.warn('data bars skipped', e); }
      }
      if (activate) ws.activate();
      await ctx.sync().catch(() => {}); // cosmetic only
    });
    return values.length;
  }

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
    readPayerEvidence, writePayerEvidence, tableExists, readClinicFile,
    writeImportRows, importXlsx, readXlsxRows, writeResults, writeTasks, writeExport, writeIr, readIr, readTasks,
    writeDiscoveryVerify, readDiscoveryVerify, writeVerifyFailed, readVerifyFailed,writeBenefits, readUploads, writeUploads, readIrUploads, writeIrUploads, tableRowCount, setTaskDone, canInsertWorkbook, setConfigValue, activate, templateCheck };
})(typeof self !== 'undefined' ? self : this);
