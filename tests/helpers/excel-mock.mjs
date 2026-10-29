/* excel-mock.mjs — in-memory Excel JS API subset with Office.js-style deferred batching and failure injection.
 * Loads the real excel.js against it. Synthetic data only.
 *
 *   const m = createMock();            // workbook with one visible sheet, "Import"
 *   await m.SoharExcel.writeTasks(cols, rows);
 *   m.inject({ failAtOp: 3 });          // 3rd mutating op from now throws; earlier ops stay applied (no rollback)
 *   m.inject({ failAtSync: 2 });        // 2nd sync from now rejects before running anything
 *   m.inject({ killAfterSync: 2 });     // 2nd sync applies, then "Excel closes": that batch never resumes, its Excel.run
 *                                       // rejects, and every later Excel.run rejects until inject() is called again
 *   m.inject({ payloadLimit: 5e6 });    // reject a sync whose request (setter payload) or response (loaded values) is too big
 *   m.inject();                         // clear injection, revive, reset opCount/syncCount
 */
import vm from 'node:vm';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const EXCEL_JS = new URL('../../excel.js', import.meta.url);
const MAX_ROWS = 1048576, MAX_COLS = 16384;
const NEVER = () => new Promise(() => {});

export function xlErr(code, message) { const e = new Error(message); e.name = 'RichApi.Error'; e.code = code; return e; }
const notLoaded = (p) => xlErr('PropertyNotLoaded', `The property '${p}' is not available. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`);

// ---------- addresses ----------
const k = (r, c) => r * MAX_COLS + c;
const rc = (key) => [Math.floor(key / MAX_COLS), key % MAX_COLS];
export function colName(c) { let s = ''; for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s; return s; }
function parseCell(a) {
  const m = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(a);
  if (!m) return null;
  let c = 0; for (const ch of m[1].toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64;
  return { r: +m[2] - 1, c: c - 1 };
}
function parseAddress(addr) {
  let sheet = null, ref = String(addr);
  const bang = ref.lastIndexOf('!');
  if (bang >= 0) { sheet = ref.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'"); ref = ref.slice(bang + 1); }
  const [a, b] = ref.split(':'); const p = parseCell(a), q = b ? parseCell(b) : p;
  if (!p || !q) throw xlErr('InvalidArgument', `The argument is invalid or missing or has an incorrect format. (address ${addr})`);
  return { sheet, r: Math.min(p.r, q.r), c: Math.min(p.c, q.c), rows: Math.abs(q.r - p.r) + 1, cols: Math.abs(q.c - p.c) + 1 };
}
const quoteSheet = (n) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
const a1 = (g) => `${colName(g.c)}${g.r + 1}` + (g.rows > 1 || g.cols > 1 ? `:${colName(g.c + g.cols - 1)}${g.r + g.rows}` : '');
const intersects = (a, b) => a.r < b.r + b.rows && b.r < a.r + a.rows && a.c < b.c + b.cols && b.c < a.c + a.cols;
const inside = (g, r, c) => r >= g.r && r < g.r + g.rows && c >= g.c && c < g.c + g.cols;
function looksLikeCellRef(n) {
  if (/^[RC]$/i.test(n) || /^R\d*C\d*$/i.test(n)) return true;
  const p = parseCell(n);
  return !!p && p.c < MAX_COLS && p.r < MAX_ROWS;
}

// ---------- workbook model ----------
class Workbook {
  constructor(M) { this.M = M; this.sheets = []; this.tables = []; this.active = null; this.seq = 0; this.structureProtected = false; }
  // Review > Protect Workbook (structure): Excel refuses to add, delete, rename, hide/unhide or move sheets.
  guardStructure() { if (this.structureProtected) throw xlErr('AccessDenied', 'You cannot perform the requested operation. (workbook structure is protected)'); }
  sheet(key) {
    const s = this.sheets.find((x) => x.id === key) || this.sheets.find((x) => x.name.toLowerCase() === String(key).toLowerCase());
    return s || null;
  }
  table(name) { return this.tables.find((t) => t.name.toLowerCase() === String(name).toLowerCase()) || null; }
  tablesOn(s) { return this.tables.filter((t) => t.sheet === s); }
  visible() { return this.sheets.filter((s) => s.visibility === 'Visible'); }

  checkSheetName(name, self) {
    name = String(name);
    if (!name.length || name.length > 31 || /[[\]:*?/\\]/.test(name) || /^'|'$/.test(name) || /^history$/i.test(name))
      throw xlErr('InvalidArgument', `The argument is invalid or missing or has an incorrect format. (sheet name "${name}")`);
    const dup = this.sheet(name);
    if (dup && dup !== self) throw xlErr('ItemAlreadyExists', `A resource with the same name or identifier already exists. (sheet "${name}")`);
    return name;
  }
  addSheet(name) {
    if (name == null) { let n = this.sheets.length + 1; while (this.sheet(`Sheet${n}`)) n++; name = `Sheet${n}`; }
    this.guardStructure();
    name = this.checkSheetName(name, null);
    const s = { id: `{00000000-0001-0000-0000-${String(++this.seq).padStart(12, '0')}}`, name, visibility: 'Visible',
      vals: new Map(), fmts: new Map(), frozenRows: 0, cfs: [], autofits: 0, deleted: false };
    this.sheets.push(s);
    if (!this.active) this.active = s;
    return s;
  }
  deleteSheet(s) {
    this.guardStructure();
    if (s.visibility === 'Visible' && this.visible().length === 1) throw xlErr('InvalidOperation', 'This operation is not permitted for the current object. (cannot delete the last visible sheet)');
    for (const t of this.tablesOn(s)) this.dropTable(t);
    this.sheets.splice(this.sheets.indexOf(s), 1);
    s.deleted = true;
    if (this.active === s) this.active = this.visible()[0] || null;
  }
  setVisibility(s, v) {
    const norm = { visible: 'Visible', hidden: 'Hidden', veryhidden: 'VeryHidden' }[String(v).toLowerCase()];
    if (!norm) throw xlErr('InvalidArgument', `Bad visibility ${v}`);
    if (norm !== s.visibility) this.guardStructure();
    if (norm !== 'Visible' && s.visibility === 'Visible' && this.visible().length === 1) throw xlErr('InvalidOperation', 'This operation is not permitted for the current object. (cannot hide the last visible sheet)');
    s.visibility = norm;
    if (norm !== 'Visible' && this.active === s) this.active = this.visible()[0];
  }
  setPosition(s, pos) {
    if (!Number.isInteger(pos) || pos < 0 || pos >= this.sheets.length) throw xlErr('InvalidArgument', `Bad position ${pos}`);
    if (this.sheets.indexOf(s) !== pos) this.guardStructure();
    this.sheets.splice(this.sheets.indexOf(s), 1);
    this.sheets.splice(pos, 0, s);
  }
  activate(s) {
    if (s.visibility !== 'Visible') throw xlErr('InvalidOperation', 'This operation is not permitted for the current object. (cannot activate a hidden sheet)');
    this.active = s;
  }
  copySheet(s, positionType, relativeTo) {
    let base = `${s.name}`, n = 2, name;
    do { const suf = ` (${n++})`; name = base.slice(0, 31 - suf.length) + suf; } while (this.sheet(name));
    const c = this.addSheet(name);
    c.vals = new Map(s.vals); c.fmts = new Map(s.fmts); c.frozenRows = s.frozenRows; c.cfs = s.cfs.map((x) => ({ ...x }));
    this.sheets.pop();
    const pt = String(positionType || 'None').toLowerCase();
    let at = this.sheets.indexOf(s) + 1; // None: right after the source (Excel's default)
    if (pt === 'beginning') at = 0;
    else if (pt === 'end') at = this.sheets.length;
    else if (pt === 'before' || pt === 'after') { at = this.sheets.indexOf(relativeTo); if (at < 0) throw xlErr('InvalidArgument', 'relativeTo'); if (pt === 'after') at++; }
    this.sheets.splice(at, 0, c);
    for (const t of this.tablesOn(s)) {
      let i = 2; while (this.table(`${t.name}${i}`)) i++;
      this.tables.push({ ...t, id: `{T${++this.seq}}`, name: `${t.name}${i}`, sheet: c, deleted: false });
    }
    return c;
  }

  // cells
  cellFmt(s, r, c) { return s.fmts.get(k(r, c)) || 'General'; }
  setValue(s, r, c, x) {
    if (x === null || x === undefined) return; // Office.js: null leaves the cell unchanged
    const key = k(r, c);
    if (typeof x === 'object') throw xlErr('InvalidArgument', 'The argument is invalid or missing or has an incorrect format. (object in values)');
    if (typeof x === 'number' && !Number.isFinite(x)) throw xlErr('InvalidArgument', 'Non-finite number in values');
    if (typeof x === 'string' && x.length > 32767) throw xlErr('InvalidArgument', 'Text values in cells are limited to 32,767 characters.');
    if (x === '') { s.vals.delete(key); return; }
    if (typeof x === 'string' && this.cellFmt(s, r, c) !== '@') {
      // General cells parse what you type: "01111" -> 1111, "true" -> TRUE. Text (@) cells keep the string.
      const t = x.trim();
      if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) x = Number(t);
      else if (/^(true|false)$/i.test(t)) x = /^true$/i.test(t);
    }
    s.vals.set(key, x);
  }
  setFmt(s, r, c, f) {
    if (f === null || f === undefined) return;
    if (f === 'General' || f === '') s.fmts.delete(k(r, c)); else s.fmts.set(k(r, c), String(f));
  }
  grid(g, v, what) {
    if (!Array.isArray(v)) {
      if (g.rows * g.cols > 1e6) throw xlErr('InvalidArgument', `Mock: refusing to broadcast ${what} over ${g.rows}x${g.cols}`);
      return Array.from({ length: g.rows }, () => new Array(g.cols).fill(v));
    }
    if (v.length !== g.rows || v.some((row) => !Array.isArray(row) || row.length !== g.cols))
      throw xlErr('InvalidArgument', `The number of rows or columns in the input array doesn't match the size or dimensions of the range. (${what}: range ${g.rows}x${g.cols}, got ${v.length}x${Array.isArray(v[0]) ? v[0].length : '?'})`);
    return v;
  }
  setValues(g, v) { const a = this.grid(g, v, 'values'); for (let i = 0; i < g.rows; i++) for (let j = 0; j < g.cols; j++) this.setValue(g.sheet, g.r + i, g.c + j, a[i][j]); this.fixHeaders(g.sheet); }
  setFormats(g, v) { const a = this.grid(g, v, 'numberFormat'); for (let i = 0; i < g.rows; i++) for (let j = 0; j < g.cols; j++) this.setFmt(g.sheet, g.r + i, g.c + j, a[i][j]); }
  readGrid(g, map, dflt) {
    if (g.rows * g.cols > 5e6) throw xlErr('ResponsePayloadSizeLimitExceeded', 'The response payload size has exceeded the limit.');
    return Array.from({ length: g.rows }, (_, i) => Array.from({ length: g.cols }, (_, j) => { const v = map.get(k(g.r + i, g.c + j)); return v === undefined ? dflt : v; }));
  }
  eachKey(map, g, fn) {
    if (g.rows * g.cols > map.size) { for (const key of [...map.keys()]) { const [r, c] = rc(key); if (inside(g, r, c)) fn(key); } }
    else for (let i = 0; i < g.rows; i++) for (let j = 0; j < g.cols; j++) { const key = k(g.r + i, g.c + j); if (map.has(key)) fn(key); }
  }
  clear(g, applyTo) {
    const a = String(applyTo || 'All').toLowerCase();
    if (a === 'all' || a === 'contents') this.eachKey(g.sheet.vals, g, (key) => g.sheet.vals.delete(key));
    if (a === 'all' || a === 'formats') { this.eachKey(g.sheet.fmts, g, (key) => g.sheet.fmts.delete(key)); g.sheet.cfs = g.sheet.cfs.filter((cf) => !intersects(cf.g, g)); }
    this.fixHeaders(g.sheet);
  }
  usedRange(s, valuesOnly) {
    let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
    for (const map of valuesOnly ? [s.vals] : [s.vals, s.fmts]) for (const key of map.keys()) {
      const [r, c] = rc(key); r0 = Math.min(r0, r); c0 = Math.min(c0, c); r1 = Math.max(r1, r); c1 = Math.max(c1, c);
    }
    return r1 < 0 ? null : { sheet: s, r: r0, c: c0, rows: r1 - r0 + 1, cols: c1 - c0 + 1 };
  }
  // Move cells in a band: keys matching `pick` are dropped if `drop`, else moved by `to`.
  moveCells(s, pick, to) {
    for (const map of [s.vals, s.fmts]) {
      const moved = [];
      for (const [key, v] of map) { const [r, c] = rc(key); const p = pick(r, c); if (!p) continue; map.delete(key); if (p === 'move') moved.push([to(r, c), v]); }
      for (const [[r, c], v] of moved) map.set(k(r, c), v);
    }
  }
  deleteCells(g, shift) {
    const s = g.sheet, conflict = () => xlErr('InsertDeleteConflict', 'The insert or delete operation attempted resulted in a conflict.');
    if (String(shift).toLowerCase() === 'left') {
      for (const t of this.tablesOn(s)) { const tg = this.tableGeom(t); if (tg.r < g.r + g.rows && g.r < tg.r + tg.rows && tg.c + tg.cols > g.c) throw conflict(); }
      this.moveCells(s, (r, c) => (r < g.r || r >= g.r + g.rows || c < g.c ? null : c < g.c + g.cols ? 'drop' : 'move'), (r, c) => [r, c - g.cols]);
      return;
    }
    const plan = [];
    for (const t of this.tablesOn(s)) {
      const tg = this.tableGeom(t), within = g.c <= tg.c && g.c + g.cols >= tg.c + tg.cols, overlapCols = tg.c < g.c + g.cols && g.c < tg.c + tg.cols;
      if (!intersects(tg, g)) { if (tg.r >= g.r + g.rows && overlapCols) { if (!within) throw conflict(); plan.push(() => { t.r -= g.rows; }); } continue; }
      if (!within) throw conflict();
      if (g.r <= tg.r) { if (g.r + g.rows >= tg.r + tg.rows) plan.push(() => this.dropTable(t)); else throw conflict(); continue; }
      const n = Math.min(g.r + g.rows, tg.r + tg.rows) - g.r;
      if (n >= tg.rows - 1) throw conflict(); // Excel will not delete a table's last body row
      plan.push(() => { t.rows -= n; });
    }
    this.moveCells(s, (r, c) => (c < g.c || c >= g.c + g.cols || r < g.r ? null : r < g.r + g.rows ? 'drop' : 'move'), (r, c) => [r - g.rows, c]);
    for (const f of plan) f();
  }
  copyFrom(dst, src, type) {
    const t = String(type || 'All').toLowerCase();
    const vals = t !== 'formats' ? this.readGrid(src, src.sheet.vals, undefined) : null;
    const fmts = t === 'all' || t === 'formats' ? this.readGrid(src, src.sheet.fmts, undefined) : null;
    const s = dst.sheet;
    for (let i = 0; i < src.rows; i++) for (let j = 0; j < src.cols; j++) {
      const key = k(dst.r + i, dst.c + j);
      if (vals) { if (vals[i][j] === undefined) s.vals.delete(key); else s.vals.set(key, vals[i][j]); }
      if (fmts) { if (fmts[i][j] === undefined) s.fmts.delete(key); else s.fmts.set(key, fmts[i][j]); }
    }
    this.fixHeaders(s);
  }

  // tables
  tableGeom(t, part) {
    if (part === 'header') return { sheet: t.sheet, r: t.r, c: t.c, rows: 1, cols: t.cols };
    if (part === 'body') return { sheet: t.sheet, r: t.r + 1, c: t.c, rows: t.rows - 1, cols: t.cols };
    return { sheet: t.sheet, r: t.r, c: t.c, rows: t.rows, cols: t.cols };
  }
  checkTableName(name, self) {
    name = String(name);
    if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name) || name.length > 255 || looksLikeCellRef(name))
      throw xlErr('InvalidArgument', `The argument is invalid or missing or has an incorrect format. (table name "${name}")`);
    const dup = this.table(name);
    if (dup && dup !== self) throw xlErr('InvalidArgument', `A table with the name "${name}" already exists.`);
    return name;
  }
  addTable(g, hasHeaders) {
    if (!hasHeaders) throw xlErr('InvalidArgument', 'Mock: tables.add(range, false) is not modelled');
    const geo = { ...g, rows: Math.max(2, g.rows) };
    if (geo.r + geo.rows > MAX_ROWS) throw xlErr('InvalidArgument', 'Table off the sheet');
    for (const t of this.tablesOn(g.sheet)) if (intersects(this.tableGeom(t), geo))
      throw xlErr('InvalidOperation', "A table can't overlap a range that contains a PivotTable report, query results, protected cells or another table.");
    let n = 1; while (this.table(`Table${n}`)) n++;
    const t = { id: `{T${++this.seq}}`, name: `Table${n}`, sheet: g.sheet, r: geo.r, c: geo.c, rows: geo.rows, cols: geo.cols, deleted: false };
    this.tables.push(t);
    this.fixHeaders(g.sheet);
    return t;
  }
  dropTable(t) { this.tables.splice(this.tables.indexOf(t), 1); t.deleted = true; }
  deleteTable(t) { const g = this.tableGeom(t); this.dropTable(t); this.clear(g, 'All'); } // Table.delete removes its cells too (convertToRange keeps them)
  resizeTable(t, g) {
    if (g.sheet !== t.sheet || g.r !== t.r || g.rows < 2 || !intersects(g, this.tableGeom(t)))
      throw xlErr('InvalidArgument', 'The header row must stay in the same row and the new range must overlap the table.');
    for (const o of this.tablesOn(t.sheet)) if (o !== t && intersects(this.tableGeom(o), g)) throw xlErr('InvalidOperation', "A table can't overlap another table.");
    Object.assign(t, { c: g.c, rows: g.rows, cols: g.cols });
    this.fixHeaders(t.sheet);
  }
  addRows(t, index, values) {
    if (!Array.isArray(values) || !values.length || values.some((r) => !Array.isArray(r) || r.length !== t.cols))
      throw xlErr('InvalidArgument', `The number of rows or columns in the input array doesn't match the size or dimensions of the range. (rows.add on ${t.cols} columns)`);
    const body = t.rows - 1, at = index == null || index === -1 ? body : index;
    if (!Number.isInteger(at) || at < 0 || at > body) throw xlErr('InvalidArgument', `Bad row index ${index}`);
    const R = t.r + 1 + at, n = values.length, s = t.sheet;
    if (t.r + t.rows + n > MAX_ROWS) throw xlErr('InvalidArgument', 'Rows would fall off the sheet');
    for (const o of this.tablesOn(s)) {
      if (o === t || o.r < R || !(o.c < t.c + t.cols && t.c < o.c + o.cols)) continue;
      if (o.c < t.c || o.c + o.cols > t.c + t.cols) throw xlErr('InsertDeleteConflict', 'The insert or delete operation attempted resulted in a conflict.');
      o.r += n;
    }
    this.moveCells(s, (r, c) => (c >= t.c && c < t.c + t.cols && r >= R ? 'move' : null), (r, c) => [r + n, c]);
    const src = at > 0 ? R - 1 : R + n; // new rows take the format of the row above (or below, at the top)
    for (let i = 0; i < n; i++) for (let j = 0; j < t.cols; j++) this.setFmt(s, R + i, t.c + j, s.fmts.get(k(src, t.c + j)));
    t.rows += n;
    for (let i = 0; i < n; i++) for (let j = 0; j < t.cols; j++) this.setValue(s, R + i, t.c + j, values[i][j]);
  }
  // Excel keeps header cells non-empty text: blanks become ColumnN, numbers become text, duplicates get a number.
  fixHeaders(s) {
    for (const t of this.tablesOn(s)) {
      const seen = new Set();
      for (let j = 0; j < t.cols; j++) {
        const key = k(t.r, t.c + j);
        let h = s.vals.get(key); h = h === undefined || h === '' ? `Column${j + 1}` : String(h);
        let name = h, n = 2; while (seen.has(name.toLowerCase())) name = `${h}${n++}`;
        seen.add(name.toLowerCase());
        s.vals.set(key, name);
      }
    }
  }
}

// ---------- proxies ----------
function props(p) {
  if (p == null) return null;
  if (typeof p === 'object' && !Array.isArray(p)) p = p.select || '';
  return (Array.isArray(p) ? p : String(p).split(',')).map((x) => x.trim()).filter(Boolean);
}

class Proxy0 {
  constructor(ctx, resolve, perSync) { this._ctx = ctx; this._resolve = resolve; this._perSync = perSync; this._b = undefined; this._bn = -1; this._l = {}; }
  get _M() { return this._ctx._M; }
  get _wb() { return this._ctx._M.wb; }
  // Ranges re-evaluate their path every sync; entities (sheets, tables) keep their identity once found.
  _bind() {
    if (this._perSync ? this._bn !== this._M.syncs : this._b == null) { this._b = this._resolve(); this._bn = this._M.syncs; }
    return this._b;
  }
  _m() {
    const m = this._bind();
    if (m === null) throw xlErr('ItemNotFound', 'The requested resource doesn\'t exist. (method called on a null object)');
    if (m.deleted || (m.sheet && m.sheet.deleted)) throw xlErr('ItemNotFound', 'The requested resource doesn\'t exist. (object was deleted)');
    return m;
  }
  _q(mut, run, bytes = 0) { this._ctx._push(mut, run, bytes); }
  _need(set, ver) { if (!this._M.supported(set, ver)) throw xlErr('ApiNotFound', `The API you are trying to use could not be found. (${set} ${ver})`); }
  _v(p) { if (!(p in this._l)) throw notLoaded(p); return this._l[p]; }
  get isNullObject() { return this._v('isNullObject'); }
  load(p) {
    const list = props(p) || this._defaults;
    this._q(false, () => {
      const m = this._bind();
      const snap = { isNullObject: m === null };
      if (m !== null) { this._m(); for (const name of list) if (name !== 'isNullObject') snap[name] = this._read(m, name); }
      Object.assign(this._l, snap);
    });
    return this;
  }
  _read(m, p) { throw xlErr('InvalidArgument', `Mock: ${this.constructor.name}.${p} is not modelled`); }
  _cache(key, make) { return (this['_c_' + key] ||= make()); }
  toJSON() { return { ...this._l }; }
}

class CollectionP extends Proxy0 {
  // owner() -> model owning the list; list(owner) -> item models; wrap(model) -> item proxy
  constructor(ctx, owner, list, wrap) { super(ctx, owner, false); this._list = list; this._wrap = wrap; this._items = null; }
  load(p) {
    const list = (props(p) || []).map((x) => x.replace(/^items\//, '')).filter((x) => x !== 'items');
    this._q(false, () => {
      const owner = this._m();
      this._items = this._list(owner).map((m) => {
        const it = this._wrap(m);
        it._b = m;
        it._l.isNullObject = false;
        for (const name of list.length ? list : it._defaults) it._l[name] = it._read(m, name);
        return it;
      });
      this._l.items = true;
    });
    return this;
  }
  get items() { this._v('items'); return this._items; }
}

class WorkbookP extends Proxy0 {
  constructor(ctx) { super(ctx, () => ctx._M.wb, false); }
  get worksheets() { return this._cache('ws', () => new WorksheetsP(this._ctx)); }
  get tables() { return this._cache('tables', () => new TablesP(this._ctx, () => this._wb, null)); }
  get protection() { return this._cache('prot', () => new ProtectionP(this._ctx)); }
  insertWorksheetsFromBase64(b64, opts) {
    const result = { _set: false, _v: undefined, get value() { if (!this._set) throw notLoaded('value'); return this._v; } };
    this._q(true, () => {
      this._need('ExcelApi', '1.13');
      const fx = this._M.xlsx.get(b64);
      if (!fx) throw xlErr('InvalidArgument', 'Mock: unknown base64 workbook (register it with mock.registerXlsx)');
      const ids = fx.map((sh) => {
        let name = sh.name, n = 2; while (this._wb.sheet(name)) name = `${sh.name} (${n++})`;
        const s = this._wb.addSheet(name);
        (sh.rows || []).forEach((row, i) => row.forEach((v, j) => this._wb.setValue(s, i, j, v)));
        return s.id;
      });
      result._v = ids; result._set = true;
    });
    return result;
  }
}

class ProtectionP extends Proxy0 {
  constructor(ctx) { super(ctx, () => ctx._M.wb, false); this._defaults = ['protected']; }
  _read(m, p) { this._need('ExcelApi', '1.7'); if (p === 'protected') return m.structureProtected; return super._read(m, p); }
  get protected() { return this._v('protected'); }
}

class WorksheetsP extends CollectionP {
  constructor(ctx) { super(ctx, () => ctx._M.wb, (wb) => wb.sheets.slice(), (m) => new WorksheetP(ctx, () => m)); }
  getItem(key) { return new WorksheetP(this._ctx, () => { const s = this._wb.sheet(key); if (!s) throw xlErr('ItemNotFound', `The requested resource doesn't exist. (sheet "${key}")`); return s; }); }
  getItemOrNullObject(key) { return new WorksheetP(this._ctx, () => { this._need('ExcelApi', '1.4'); return this._wb.sheet(key); }); }
  getActiveWorksheet() { return new WorksheetP(this._ctx, () => this._wb.active); }
  add(name) {
    let made = null;
    this._q(true, () => { made = this._wb.addSheet(name); });
    return new WorksheetP(this._ctx, () => made);
  }
}

class WorksheetP extends Proxy0 {
  constructor(ctx, resolve) { super(ctx, resolve, false); this._defaults = ['name', 'id', 'position', 'visibility']; }
  _read(m, p) {
    if (p === 'name') return m.name;
    if (p === 'id') return m.id;
    if (p === 'position') return this._wb.sheets.indexOf(m);
    if (p === 'visibility') return m.visibility;
    return super._read(m, p);
  }
  get name() { return this._v('name'); }
  set name(v) { this._q(true, () => { const m = this._m(); if (m.name !== String(v)) this._wb.guardStructure(); m.name = this._wb.checkSheetName(v, m); }); }
  get id() { return this._v('id'); }
  get position() { return this._v('position'); }
  set position(v) { this._q(true, () => this._wb.setPosition(this._m(), v)); }
  get visibility() { return this._v('visibility'); }
  set visibility(v) { this._q(true, () => this._wb.setVisibility(this._m(), v)); }
  delete() { this._q(true, () => this._wb.deleteSheet(this._m())); }
  activate() { this._q(true, () => this._wb.activate(this._m())); }
  copy(positionType, relativeTo) {
    let made = null;
    this._q(true, () => { this._need('ExcelApi', '1.10'); made = this._wb.copySheet(this._m(), positionType, relativeTo ? relativeTo._m() : null); });
    return new WorksheetP(this._ctx, () => made);
  }
  getRange(addr) {
    return new RangeP(this._ctx, () => {
      const s = this._m();
      if (!addr) return { sheet: s, r: 0, c: 0, rows: MAX_ROWS, cols: MAX_COLS };
      const a = parseAddress(addr);
      return { sheet: s, r: a.r, c: a.c, rows: a.rows, cols: a.cols };
    });
  }
  getRangeByIndexes(r, c, rows, cols) {
    return new RangeP(this._ctx, () => {
      const ok = [r, c, rows, cols].every(Number.isInteger) && r >= 0 && c >= 0 && rows >= 1 && cols >= 1 && r + rows <= MAX_ROWS && c + cols <= MAX_COLS;
      if (!ok) throw xlErr('InvalidArgument', `The argument is invalid or missing or has an incorrect format. (getRangeByIndexes ${r},${c},${rows},${cols})`);
      return { sheet: this._m(), r, c, rows, cols };
    });
  }
  getUsedRangeOrNullObject(valuesOnly) { return new RangeP(this._ctx, () => { this._need('ExcelApi', '1.4'); return this._wb.usedRange(this._m(), !!valuesOnly); }); }
  getUsedRange(valuesOnly) { return new RangeP(this._ctx, () => this._wb.usedRange(this._m(), !!valuesOnly) || { sheet: this._m(), r: 0, c: 0, rows: 1, cols: 1 }); }
  get tables() { return this._cache('tables', () => new TablesP(this._ctx, () => this._m(), this)); }
  get freezePanes() {
    return this._cache('fp', () => ({
      freezeRows: (n = 1) => this._q(true, () => { this._need('ExcelApi', '1.7'); this._m().frozenRows = n; }),
      unfreeze: () => this._q(true, () => { this._need('ExcelApi', '1.7'); this._m().frozenRows = 0; }),
    }));
  }
}

class TablesP extends CollectionP {
  // sheetP: null for workbook.tables
  constructor(ctx, owner, sheetP) {
    super(ctx, owner, (o) => (sheetP ? ctx._M.wb.tablesOn(o) : ctx._M.wb.tables.slice()), (m) => new TableP(ctx, () => m));
    this._sheetP = sheetP;
  }
  getItem(name) {
    return new TableP(this._ctx, () => {
      const t = this._wb.table(name);
      if (!t || (this._sheetP && t.sheet !== this._m())) throw xlErr('ItemNotFound', `The requested resource doesn't exist. (table "${name}")`);
      return t;
    });
  }
  getItemOrNullObject(name) {
    return new TableP(this._ctx, () => {
      this._need('ExcelApi', '1.4');
      const t = this._wb.table(name);
      return t && (!this._sheetP || t.sheet === this._m()) ? t : null;
    });
  }
  add(range, hasHeaders) {
    let made = null;
    this._q(true, () => {
      let g;
      if (typeof range === 'string') {
        const a = parseAddress(range);
        const s = a.sheet ? this._wb.sheet(a.sheet) : this._sheetP ? this._m() : this._wb.active;
        if (!s) throw xlErr('ItemNotFound', `No sheet ${a.sheet}`);
        g = { sheet: s, r: a.r, c: a.c, rows: a.rows, cols: a.cols };
      } else g = range._m();
      if (this._sheetP && g.sheet !== this._m()) throw xlErr('InvalidArgument', 'Range is on another sheet');
      made = this._wb.addTable(g, hasHeaders);
    });
    return new TableP(this._ctx, () => made);
  }
}

class TableP extends Proxy0 {
  constructor(ctx, resolve) { super(ctx, resolve, false); this._defaults = ['name', 'id']; }
  _read(m, p) { if (p === 'name') return m.name; if (p === 'id') return m.id; return super._read(m, p); }
  get name() { return this._v('name'); }
  set name(v) { this._q(true, () => { const m = this._m(); m.name = this._wb.checkTableName(v, m); }); }
  get id() { return this._v('id'); }
  get worksheet() { return this._cache('ws', () => new WorksheetP(this._ctx, () => this._m().sheet)); }
  getRange() { return new RangeP(this._ctx, () => this._wb.tableGeom(this._m())); }
  getHeaderRowRange() { return new RangeP(this._ctx, () => this._wb.tableGeom(this._m(), 'header')); }
  getDataBodyRange() { return new RangeP(this._ctx, () => this._wb.tableGeom(this._m(), 'body')); }
  delete() { this._q(true, () => this._wb.deleteTable(this._m())); }
  convertToRange() { let g = null; this._q(true, () => { const m = this._m(); g = this._wb.tableGeom(m); this._wb.dropTable(m); }); return new RangeP(this._ctx, () => g); }
  clearFilters() { this._q(true, () => { this._m(); }); }
  resize(range) {
    this._q(true, () => {
      this._need('ExcelApi', '1.13');
      const m = this._m();
      let g;
      if (typeof range === 'string') { const a = parseAddress(range); g = { sheet: a.sheet ? this._wb.sheet(a.sheet) : m.sheet, r: a.r, c: a.c, rows: a.rows, cols: a.cols }; }
      else g = range._m();
      this._wb.resizeTable(m, g);
    });
  }
  get rows() {
    return this._cache('rows', () => ({
      add: (index, values) => { this._q(true, () => this._wb.addRows(this._m(), index, values), JSON.stringify(values || []).length); return {}; },
      getCount: () => {
        const res = { _set: false, get value() { if (!this._set) throw notLoaded('value'); return this._v; } };
        this._q(false, () => { res._v = this._m().rows - 1; res._set = true; });
        return res;
      },
    }));
  }
}

const CF_TYPES = new Set(['Custom', 'DataBar', 'ColorScale', 'IconSet', 'TopBottom', 'PresetCriteria', 'ContainsText', 'CellValue']);

class RangeP extends Proxy0 {
  constructor(ctx, resolve) { super(ctx, resolve, true); this._defaults = ['address', 'rowCount', 'columnCount', 'rowIndex', 'columnIndex', 'values', 'numberFormat']; }
  _read(g, p) {
    switch (p) {
      case 'values': { const v = this._wb.readGrid(g, g.sheet.vals, ''); this._M.responseBytes += JSON.stringify(v).length; this._M.valueLoads.push(g.rows); return v; }
      case 'numberFormat': return this._wb.readGrid(g, g.sheet.fmts, 'General');
      case 'rowCount': return g.rows;
      case 'columnCount': return g.cols;
      case 'rowIndex': return g.r;
      case 'columnIndex': return g.c;
      case 'cellCount': return g.rows * g.cols;
      case 'address': return `${quoteSheet(g.sheet.name)}!${a1(g)}`;
      case 'addressLocal': return `${quoteSheet(g.sheet.name)}!${a1(g)}`;
      default: return super._read(g, p);
    }
  }
  get values() { return this._v('values'); }
  set values(v) { this._q(true, () => this._wb.setValues(this._m(), v), JSON.stringify(v ?? null).length); }
  get numberFormat() { return this._v('numberFormat'); }
  set numberFormat(v) { this._q(true, () => this._wb.setFormats(this._m(), v), JSON.stringify(v ?? null).length); }
  get rowCount() { return this._v('rowCount'); }
  get columnCount() { return this._v('columnCount'); }
  get rowIndex() { return this._v('rowIndex'); }
  get columnIndex() { return this._v('columnIndex'); }
  get cellCount() { return this._v('cellCount'); }
  get address() { return this._v('address'); }
  get worksheet() { return this._cache('ws', () => new WorksheetP(this._ctx, () => this._m().sheet)); }
  clear(applyTo) { this._q(true, () => this._wb.clear(this._m(), applyTo)); }
  delete(shift) { this._q(true, () => this._wb.deleteCells(this._m(), shift)); }
  copyFrom(source, copyType, skipBlanks, transpose) {
    if (skipBlanks || transpose) throw new Error('Mock: copyFrom skipBlanks/transpose not modelled');
    this._q(true, () => {
      this._need('ExcelApi', '1.9');
      const dst = this._m();
      let src;
      if (typeof source === 'string') { const a = parseAddress(source); const s = a.sheet ? this._wb.sheet(a.sheet) : dst.sheet; if (!s) throw xlErr('ItemNotFound', `No sheet ${a.sheet}`); src = { sheet: s, r: a.r, c: a.c, rows: a.rows, cols: a.cols }; }
      else src = source._m();
      if (dst.r + src.rows > MAX_ROWS || dst.c + src.cols > MAX_COLS) throw xlErr('InvalidArgument', 'copyFrom would fall off the sheet');
      this._wb.copyFrom(dst, src, copyType);
    });
  }
  _sub(f) { return new RangeP(this._ctx, () => { const g = this._m(); const s = f(g); if (s.rows < 1 || s.cols < 1 || s.r < 0 || s.c < 0 || s.r + s.rows > MAX_ROWS || s.c + s.cols > MAX_COLS) throw xlErr('InvalidArgument', 'The argument is invalid or missing or has an incorrect format. (range out of bounds)'); return { sheet: g.sheet, ...s }; }); }
  _within(i, n) { if (!Number.isInteger(i) || i < 0 || i >= n) throw xlErr('InvalidArgument', `Index ${i} out of range`); }
  getRow(i) { return this._sub((g) => { this._within(i, g.rows); return { r: g.r + i, c: g.c, rows: 1, cols: g.cols }; }); }
  getColumn(j) { return this._sub((g) => { this._within(j, g.cols); return { r: g.r, c: g.c + j, rows: g.rows, cols: 1 }; }); }
  getCell(i, j) { return this._sub((g) => ({ r: g.r + i, c: g.c + j, rows: 1, cols: 1 })); }
  getLastRow() { return this._sub((g) => ({ r: g.r + g.rows - 1, c: g.c, rows: 1, cols: g.cols })); }
  getLastColumn() { return this._sub((g) => ({ r: g.r, c: g.c + g.cols - 1, rows: g.rows, cols: 1 })); }
  getResizedRange(dr, dc) { return this._sub((g) => ({ r: g.r, c: g.c, rows: g.rows + dr, cols: g.cols + dc })); }
  getOffsetRange(dr, dc) { return this._sub((g) => ({ r: g.r + dr, c: g.c + dc, rows: g.rows, cols: g.cols })); }
  getBoundingRect(other) {
    return this._sub((g) => {
      const o = typeof other === 'string' ? { ...parseAddress(other), sheet: g.sheet } : other._m();
      if (o.sheet !== g.sheet) throw xlErr('InvalidArgument', 'Ranges are on different sheets');
      const r = Math.min(g.r, o.r), c = Math.min(g.c, o.c);
      return { r, c, rows: Math.max(g.r + g.rows, o.r + o.rows) - r, cols: Math.max(g.c + g.cols, o.c + o.cols) - c };
    });
  }
  get format() {
    return this._cache('fmt', () => ({
      autofitColumns: () => this._q(true, () => { this._need('ExcelApi', '1.2'); this._m().sheet.autofits++; }),
      autofitRows: () => this._q(true, () => { this._need('ExcelApi', '1.2'); this._m().sheet.autofits++; }),
    }));
  }
  get conditionalFormats() {
    return this._cache('cfs', () => ({
      add: (type) => {
        let cf = null;
        this._q(true, () => {
          this._need('ExcelApi', '1.6');
          if (!CF_TYPES.has(type)) throw xlErr('InvalidArgument', `Bad conditional format type ${type}`);
          const g = this._m(); cf = { type, g, rules: {} }; g.sheet.cfs.push(cf);
        });
        const setRule = (name) => ({ set: (v) => this._q(true, () => {
          if (!v || typeof v !== 'object' || !v.type) throw xlErr('InvalidArgument', `Bad ${name}`);
          if (cf.type !== 'DataBar') throw xlErr('InvalidOperation', 'dataBar on a non-DataBar format');
          cf.rules[name] = v;
        }) });
        const dataBar = Object.defineProperties({}, { lowerBoundRule: setRule('lowerBoundRule'), upperBoundRule: setRule('upperBoundRule') });
        return { dataBar };
      },
      clearAll: () => this._q(true, () => { this._need('ExcelApi', '1.6'); const g = this._m(); g.sheet.cfs = g.sheet.cfs.filter((cf) => !intersects(cf.g, g)); }),
    }));
  }
}

class Ctx {
  constructor(M) {
    this._M = M; this._queue = []; this._dead = false; this._kill = null;
    this.workbook = new WorkbookP(this);
    this.application = { suspendApiCalculationUntilNextSync: () => this._push(false, () => { if (!M.supported('ExcelApi', '1.6')) throw xlErr('ApiNotFound', 'suspendApiCalculationUntilNextSync needs ExcelApi 1.6'); }) };
    this.trackedObjects = { add() {}, remove() {} };
  }
  _push(mut, run, bytes = 0) { this._queue.push({ mut, run, bytes }); }
  sync() { return this._M._sync(this); }
}

// ---------- the mock ----------
const cmpVer = (a, b) => { const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number); for (let i = 0; i < Math.max(x.length, y.length); i++) { const d = (x[i] || 0) - (y[i] || 0); if (d) return d; } return 0; };

export function createMock(opts = {}) {
  const M = {
    wb: null, syncs: 0, ops: 0, injected: {}, killed: false, responseBytes: 0, valueLoads: [], xlsx: new Map(),
    sets: { ExcelApi: '1.17', DialogApi: '1.2', ...(opts.sets || {}) },
    supported(name, ver) { const have = this.sets[name]; return have != null && (ver == null || cmpVer(have, ver) >= 0); },
    _sync(ctx) {
      const q = ctx._queue; ctx._queue = [];
      if (ctx._dead) return NEVER();
      const closed = () => { ctx._dead = true; ctx._kill(xlErr('ExcelClosed', `Mock: Excel closed after sync ${this.syncs}`)); return NEVER(); };
      if (this.killed) return closed();
      const n = ++this.syncs, inj = this.injected;
      return Promise.resolve().then(() => {
        if (inj.failAtSync === n) throw xlErr('InjectedFailure', `Mock: sync ${n} failed (failAtSync)`);
        if (inj.payloadLimit && q.reduce((s, o) => s + o.bytes, 0) > inj.payloadLimit) throw xlErr('RequestPayloadSizeLimitExceeded', 'The request payload size has exceeded the limit.');
        this.responseBytes = 0;
        for (const op of q) {
          if (op.mut && ++this.ops === inj.failAtOp) throw xlErr('InjectedFailure', `Mock: operation ${this.ops} failed (failAtOp)`);
          op.run();
        }
        if (inj.payloadLimit && this.responseBytes > inj.payloadLimit) throw xlErr('ResponsePayloadSizeLimitExceeded', 'The response payload size has exceeded the limit.');
        if (inj.killAfterSync === n) { this.killed = true; return closed(); }
      });
    },
    run(a, b) {
      const fn = typeof a === 'function' ? a : b;
      if (typeof fn !== 'function') return Promise.reject(new TypeError('Excel.run needs a batch function'));
      if (M.killed) return Promise.reject(xlErr('ExcelClosed', 'Mock: Excel is closed'));
      const ctx = new Ctx(M);
      return new Promise((resolve, reject) => {
        ctx._kill = reject;
        Promise.resolve().then(() => fn(ctx))
          .then(async (v) => { if (ctx._queue.length) await ctx.sync(); return v; }) // Excel.run flushes what is still queued
          .then((v) => { if (!ctx._dead) resolve(v); }, (e) => { if (!ctx._dead) reject(e); });
      });
    },
  };
  M.wb = new Workbook(M);
  for (const name of opts.sheets || ['Import']) M.wb.addSheet(name);

  const Excel = {
    run: (a, b) => M.run(a, b),
    ClearApplyTo: { all: 'All', formats: 'Formats', contents: 'Contents', hyperlinks: 'Hyperlinks', removeHyperlinks: 'RemoveHyperlinks' },
    DeleteShiftDirection: { up: 'Up', left: 'Left' },
    InsertShiftDirection: { down: 'Down', right: 'Right' },
    SheetVisibility: { visible: 'Visible', hidden: 'Hidden', veryHidden: 'VeryHidden' },
    RangeCopyType: { all: 'All', formulas: 'Formulas', values: 'Values', formats: 'Formats' },
    WorksheetPositionType: { none: 'None', before: 'Before', after: 'After', beginning: 'Beginning', end: 'End' },
    ConditionalFormatType: Object.fromEntries([...CF_TYPES].map((t) => [t[0].toLowerCase() + t.slice(1), t])),
    ErrorCodes: { itemNotFound: 'ItemNotFound', invalidArgument: 'InvalidArgument', invalidOperation: 'InvalidOperation', insertDeleteConflict: 'InsertDeleteConflict' },
  };
  const Office = {
    context: { requirements: { isSetSupported: (name, ver) => M.supported(name, ver) } },
    onReady: (cb) => Promise.resolve(cb ? cb({ host: 'Excel' }) : { host: 'Excel' }),
    HostType: { Excel: 'Excel' },
  };

  // The real excel.js, run in this realm (so arrays/objects it returns compare cleanly) with its own globals.
  // opts.knowledge: also load the built-in payer knowledge (payer-knowledge.js), as index.html does. Off by default so
  // Payer Map tests see only their own synthetic rows.
  const self = { SoharCore: require('../../core.js'), ...(opts.knowledge ? { PayerKnowledge: require('../../payer-knowledge.js') } : {}) };
  // opts.source or EXCEL_MOCK_SOURCE=path: run a candidate excel.js without touching the real file.
  const src = opts.source ?? fs.readFileSync(process.env.EXCEL_MOCK_SOURCE || EXCEL_JS, 'utf8');
  vm.runInThisContext(`(function (self, globalThis, Excel, Office, console) {\n${src}\n})`, { filename: fileURLToPath(EXCEL_JS) })
    .call(self, self, self, Excel, Office, opts.console || console);

  const wb = M.wb;
  const sheetOrThrow = (name) => { const s = wb.sheet(name); if (!s) throw new Error(`no sheet ${name}`); return s; };
  const tableOrThrow = (name) => { const t = wb.table(name); if (!t) throw new Error(`no table ${name}`); return t; };
  const range = (sheet, address) => { const s = sheetOrThrow(sheet); const a = parseAddress(address); return { sheet: s, r: a.r, c: a.c, rows: a.rows, cols: a.cols }; };

  return {
    SoharExcel: self.SoharExcel, Excel, Office, wb, M,
    inject(o) { M.injected = { ...(o || {}) }; M.killed = false; M.ops = 0; M.syncs = 0; M.valueLoads = []; },
    setSupport(name, ver) { if (ver == null) delete M.sets[name]; else M.sets[name] = ver; },
    opCount: () => M.ops,
    syncCount: () => M.syncs,
    valueLoads: () => M.valueLoads.slice(),
    killed: () => M.killed,
    sheetNames: () => wb.sheets.map((s) => s.name),
    visibleSheets: () => wb.visible().map((s) => s.name),
    sheetVisibility: (name) => sheetOrThrow(name).visibility,
    activeSheet: () => (wb.active ? wb.active.name : null),
    frozenRows: (name) => sheetOrThrow(name).frozenRows,
    autofits: (name) => sheetOrThrow(name).autofits,
    conditionalFormats: (name) => sheetOrThrow(name).cfs.map((cf) => ({ type: cf.type, address: a1(cf.g), rules: cf.rules })),
    tableNames: () => wb.tables.map((t) => t.name),
    tableInfo(name) { const t = tableOrThrow(name); return { sheet: t.sheet.name, address: a1(wb.tableGeom(t)), rows: t.rows - 1, cols: t.cols }; },
    /** Body rows as objects by header. Fully blank rows are skipped unless all is true. */
    tableRows(name, all) {
      const t = tableOrThrow(name);
      const [hdr] = wb.readGrid(wb.tableGeom(t, 'header'), t.sheet.vals, '');
      return wb.readGrid(wb.tableGeom(t, 'body'), t.sheet.vals, '')
        .filter((r) => all || r.some((v) => v !== ''))
        .map((r) => Object.fromEntries(hdr.map((h, j) => [h, r[j]])));
    },
    values: (sheet, address) => { const g = range(sheet, address); return wb.readGrid(g, g.sheet.vals, ''); },
    numberFormats: (sheet, address) => { const g = range(sheet, address); return wb.readGrid(g, g.sheet.fmts, 'General'); },
    usedAddress: (sheet) => { const g = wb.usedRange(sheetOrThrow(sheet), true); return g ? a1(g) : null; },
    addSheet(name, rows, opts2 = {}) {
      const s = wb.sheet(name) || wb.addSheet(name);
      (rows || []).forEach((row, i) => row.forEach((v, j) => { if (opts2.text) wb.setFmt(s, i, j, '@'); wb.setValue(s, i, j, v); }));
      if (opts2.hidden) wb.setVisibility(s, 'Hidden');
      return name;
    },
    /** Put a table straight into the model (no ops counted). rows: objects by column; [] leaves one blank body row. */
    seedTable(sheetName, tableName, cols, rows, opts2 = {}) {
      const s = wb.sheet(sheetName) || wb.addSheet(sheetName);
      const r0 = opts2.row || 0, c0 = opts2.col || 0;
      const grid = [cols, ...rows.map((r) => cols.map((c) => (r[c] == null ? '' : r[c])))];
      grid.forEach((row, i) => row.forEach((v, j) => { if (opts2.text !== false) wb.setFmt(s, r0 + i, c0 + j, '@'); wb.setValue(s, r0 + i, c0 + j, v); }));
      const t = wb.addTable({ sheet: s, r: r0, c: c0, rows: grid.length, cols: cols.length }, true);
      t.name = wb.checkTableName(tableName, t);
      if (opts2.hidden) wb.setVisibility(s, 'Hidden');
      return tableName;
    },
    registerXlsx(b64, sheets) { M.xlsx.set(b64, sheets); },
    protectStructure(on) { wb.structureProtected = !!on; },
  };
}
