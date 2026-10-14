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

  async function run(fn) { return Excel.run(fn); }

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

  async function readTable(tableName) {
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItemOrNullObject(tableName);
      t.load('isNullObject');
      await ctx.sync();
      if (t.isNullObject) throw new Error(`Table ${tableName} is missing. Use the Sohar Eligibility template workbook.`);
      const h = t.getHeaderRowRange(); h.load('values');
      const b = t.getDataBodyRange(); b.load('values');
      await ctx.sync();
      const headers = h.values[0].map((x) => String(x).trim());
      return b.values
        .filter((r) => r.some((v) => v !== '' && v != null))
        .map((r) => Object.fromEntries(headers.map((k, i) => [k, r[i]])));
    });
  }

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

  async function upsertPayerMap(entry) {
    return run(async (ctx) => {
      const t = ctx.workbook.tables.getItem(TABLES.payerMap);
      const h = t.getHeaderRowRange(); h.load('values');
      const b = t.getDataBodyRange(); b.load('values,rowCount');
      await ctx.sync();
      const headers = h.values[0].map((x) => String(x).trim());
      const col = (name) => headers.indexOf(name);
      const today = root.SoharCore.fmtDate(new Date(Date.UTC(new Date().getFullYear(), new Date().getMonth(), new Date().getDate())));
      const rowVals = headers.map(() => '');
      rowVals[col('NextGen Payer Name')] = entry.nextgenPayerName;
      rowVals[col('Sohar payerId')] = entry.payerId;
      rowVals[col('Sohar payerName')] = entry.payerName;
      rowVals[col('Confidence')] = entry.confidence;
      rowVals[col('Action')] = entry.payerId ? 'verification' : 'discovery';
      rowVals[col('Notes')] = entry.notes || '';
      if (col('Last reviewed') >= 0) rowVals[col('Last reviewed')] = today;
      const i = b.values.findIndex((r) => String(r[col('NextGen Payer Name')]).trim().toLowerCase() === String(entry.nextgenPayerName).trim().toLowerCase());
      if (i >= 0) {
        b.getRow(i).values = [rowVals];
      } else {
        const blank = b.values.findIndex((r) => r.every((v) => v === '' || v == null));
        if (blank >= 0 && b.rowCount === 1) b.getRow(blank).values = [rowVals];
        else t.rows.add(null, [rowVals]);
      }
      await ctx.sync();
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
      const hdr = t.getHeaderRowRange(); hdr.load('rowIndex,columnIndex,columnCount');
      const body = t.getDataBodyRange(); body.load('rowCount');
      await ctx.sync();
      if (hdr.columnCount !== cols.length) throw new Error(`${tableName} has ${hdr.columnCount} columns; this version of the add-in expects ${cols.length}. Use the matching template.`);
      ctx.application.suspendApiCalculationUntilNextSync();
      if (body.rowCount > 0) { body.delete(Excel.DeleteShiftDirection.up); await ctx.sync(); }
      // Excel leaves one empty body row. Everything is written as text so member IDs and ZIPs keep leading zeros.
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

  /** Insert the first sheet of an .xlsx (base64) as a temporary sheet, hand its used range to fn, then remove it. */
  async function withTempSheet(base64, fn) {
    if (!canInsertWorkbook()) throw new Error('This version of Excel cannot open .xlsx files from the pane. Save the file as CSV and upload that, or paste it into the sheet.');
    return run(async (ctx) => {
      const inserted = ctx.workbook.insertWorksheetsFromBase64(base64, { positionType: Excel.WorksheetPositionType.end });
      await ctx.sync();
      const names = inserted.value || [];
      if (!names.length) throw new Error('That workbook has no sheets.');
      try {
        const used = ctx.workbook.worksheets.getItem(names[0]).getUsedRangeOrNullObject(true);
        used.load('isNullObject,rowCount,columnCount');
        await ctx.sync();
        return await fn(ctx, used);
      } finally {
        for (const n of names) ctx.workbook.worksheets.getItem(n).delete();
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

  /** Values of the first sheet of an uploaded .xlsx (Sohar results saved from Excel). */
  async function readXlsxRows(base64) {
    return withTempSheet(base64, async (ctx, used) => {
      if (used.isNullObject) return [];
      used.load('values');
      await ctx.sync();
      return used.values;
    });
  }

  const writeResults = (cols, rows) => writeAddinTable(SHEETS.results, RESULTS_TABLE, cols, rows, true);
  const writeTasks = (cols, rows) => writeAddinTable(SHEETS.tasks, TASKS_TABLE, cols, rows, false);
  async function readTasks() { try { return await readTable(TASKS_TABLE); } catch (e) { return []; } }

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

  /** Create (or rebuild) a sheet + table the add-in owns, then fill it. Not part of the template. */
  async function writeAddinTable(sheetName, tableName, cols, rows, activate) {
    await run(async (ctx) => {
      const t = ctx.workbook.tables.getItemOrNullObject(tableName);
      t.load('isNullObject');
      await ctx.sync();
      if (!t.isNullObject) {
        const h = t.getHeaderRowRange(); h.load('values');
        await ctx.sync();
        if (h.values[0].join('|') === cols.join('|')) return;
        t.delete();
        await ctx.sync();
      }
      let ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
      ws.load('isNullObject');
      await ctx.sync();
      if (ws.isNullObject) ws = ctx.workbook.worksheets.add(sheetName);
      else ws.getRange().clear();
      const hdr = ws.getRangeByIndexes(0, 0, 1, cols.length);
      hdr.values = [cols];
      const nt = ws.tables.add(hdr, true);
      nt.name = tableName;
      ws.freezePanes.freezeRows(1);
      await ctx.sync();
    });
    await replaceTable(tableName, cols, rows);
    await run(async (ctx) => {
      const t = ctx.workbook.tables.getItem(tableName);
      t.getRange().format.autofitColumns();
      if (activate) t.worksheet.activate();
      await ctx.sync();
    });
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

  root.SoharExcel = { SHEETS, TABLES, readImportRows, readTable, readConfig, readPayerMap, readSoharPayers, upsertPayerMap, replaceTable,
    writeImportRows, importXlsx, readXlsxRows, writeResults, writeTasks, readTasks, setTaskDone, canInsertWorkbook, setConfigValue, activate, templateCheck };
})(typeof self !== 'undefined' ? self : this);
