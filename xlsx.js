/* xlsx.js — tiny .xlsx writer. Pure: no Office.js, no DOM, no dependencies, no network.
   buildWorkbook(spec) -> Uint8Array (a valid Office Open XML workbook, ZIP STORE, no compression).
   Each sheet gets a real Excel table, a frozen header row, text-format ID columns and optional hyperlinks.
   Optional webExtension tags the file so the add-in task pane opens with the document. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SoharXlsx = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const MAX_CELL = 32767;
  // Columns whose values are identifiers: written as text with the '@' format so leading zeros survive.
  const ID_COL_RE = /(^|[^a-z])(id|nbr|mrn|zip|zip3|prefix|ref|payerid|npi|tin)([^a-z]|$)|member id|payer ?id|per nbr|pol nbr/i;
  const STYLE = { normal: 0, header: 1, text: 2, pct: 3, link: 4 };

  // ---------- small helpers ----------

  // Strip characters XML 1.0 forbids (C0 controls except tab/LF/CR, U+FFFE/FFFF, lone surrogates).
  function stripControl(s) {
    return String(s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
      .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '$1');
  }
  function escapeXml(v) {
    if (v === null || v === undefined) return '';
    return stripControl(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }

  // 1 -> A, 26 -> Z, 27 -> AA, 702 -> ZZ, 703 -> AAA (1-based).
  function colLetter(n) {
    n = Math.floor(n);
    if (!(n >= 1 && n <= 16384)) throw new Error('Column number out of range: ' + n);
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function cellRef(col, row) { return colLetter(col) + row; } // both 1-based
  // Header row 1 + dataRows (an empty table still needs one data row in Excel).
  function tableRef(nCols, dataRows) {
    const last = 1 + Math.max(1, dataRows | 0);
    return 'A1:' + colLetter(Math.max(1, nCols)) + last;
  }

  // Table column names must be non-empty and unique (case-insensitive). Duplicates get " 2", " 3"...
  function uniqueNames(names) {
    const seen = new Set();
    return (names || []).map((raw, i) => {
      let base = stripControl(raw === null || raw === undefined ? '' : raw).replace(/[\r\n\t]+/g, ' ').trim();
      if (!base) base = 'Column' + (i + 1);
      let name = base, k = 2;
      while (seen.has(name.toLowerCase())) name = base + ' ' + (k++);
      seen.add(name.toLowerCase());
      return name;
    });
  }

  // Sheet names: max 31 chars, none of []:*?/\ , no leading/trailing apostrophe, unique (case-insensitive).
  function sanitizeSheetName(name, used) {
    let s = stripControl(name || '').replace(/[\[\]:*?\/\\]/g, ' ').replace(/^'+|'+$/g, '').trim() || 'Sheet';
    s = s.slice(0, 31).trim();
    if (used) {
      let out = s, k = 2;
      while (used.has(out.toLowerCase())) { const suf = ' (' + (k++) + ')'; out = s.slice(0, 31 - suf.length).trim() + suf; }
      used.add(out.toLowerCase());
      return out;
    }
    return s;
  }
  // Table names: letters, digits, underscore, period; start with a letter or underscore; not a cell reference; unique.
  function sanitizeTableName(name, used) {
    let s = stripControl(name || '').replace(/[^A-Za-z0-9_.]/g, '_');
    if (!/^[A-Za-z_]/.test(s)) s = 'tbl_' + s;
    if (/^[A-Za-z]{1,3}\d+$/.test(s) || /^[RrCc]$/.test(s) || /^[Rr]\d*[Cc]\d*$/.test(s)) s = 'tbl_' + s;
    s = s.slice(0, 255);
    if (used) {
      let out = s, k = 2;
      while (used.has(out.toLowerCase())) out = s + '_' + (k++);
      used.add(out.toLowerCase());
      return out;
    }
    return s;
  }

  // ---------- CRC32 + ZIP (STORE) ----------

  const CRC_TABLE = (function () {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function utf8(s) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(s);
    const out = [];
    const str = unescape(encodeURIComponent(s));
    for (let i = 0; i < str.length; i++) out.push(str.charCodeAt(i));
    return new Uint8Array(out);
  }
  function toBytes(d) { return d instanceof Uint8Array ? d : utf8(String(d)); }

  function crc32(data) {
    const b = toBytes(data);
    let c = 0xFFFFFFFF;
    for (let i = 0; i < b.length; i++) c = CRC_TABLE[(c ^ b[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function dosDateTime(d) {
    const y = Math.max(1980, d.getFullYear());
    return {
      time: ((d.getHours() << 11) | (d.getMinutes() << 5) | (Math.floor(d.getSeconds() / 2))) & 0xFFFF,
      date: (((y - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF,
    };
  }

  // files: [{ name, data: string | Uint8Array }] -> Uint8Array zip, all entries STORED.
  function zip(files, when) {
    const dt = dosDateTime(when instanceof Date ? when : new Date());
    const parts = [], central = [];
    let offset = 0;
    for (const f of files) {
      const name = utf8(f.name), data = toBytes(f.data), crc = crc32(data);
      const lh = new Uint8Array(30 + name.length), lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, 0x0800, true); lv.setUint16(8, 0, true);
      lv.setUint16(10, dt.time, true); lv.setUint16(12, dt.date, true); lv.setUint32(14, crc, true);
      lv.setUint32(18, data.length, true); lv.setUint32(22, data.length, true); lv.setUint16(26, name.length, true); lv.setUint16(28, 0, true);
      lh.set(name, 30);
      const ch = new Uint8Array(46 + name.length), cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true); cv.setUint16(12, dt.time, true); cv.setUint16(14, dt.date, true); cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true); cv.setUint32(24, data.length, true); cv.setUint16(28, name.length, true);
      cv.setUint16(30, 0, true); cv.setUint16(32, 0, true); cv.setUint16(34, 0, true); cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true); cv.setUint32(42, offset, true);
      ch.set(name, 46);
      parts.push(lh, data); central.push(ch);
      offset += lh.length + data.length;
    }
    const cdSize = central.reduce((a, c) => a + c.length, 0);
    const eocd = new Uint8Array(22), ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true); ev.setUint16(4, 0, true); ev.setUint16(6, 0, true);
    ev.setUint16(8, files.length, true); ev.setUint16(10, files.length, true);
    ev.setUint32(12, cdSize, true); ev.setUint32(16, offset, true); ev.setUint16(20, 0, true);
    const all = parts.concat(central, [eocd]);
    const out = new Uint8Array(offset + cdSize + 22);
    let p = 0;
    for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  // Read back a STORE zip (used by tests and for sanity checks). Returns [{ name, data: Uint8Array, crc }].
  function unzipStored(u8) {
    const v = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let e = u8.length - 22;
    while (e >= 0 && v.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error('No end of central directory');
    const count = v.getUint16(e + 10, true), cdOff = v.getUint32(e + 16, true);
    const dec = typeof TextDecoder !== 'undefined' ? new TextDecoder() : null;
    const out = [];
    let p = cdOff;
    for (let i = 0; i < count; i++) {
      if (v.getUint32(p, true) !== 0x02014b50) throw new Error('Bad central header at ' + p);
      const crc = v.getUint32(p + 16, true), size = v.getUint32(p + 20, true), nl = v.getUint16(p + 28, true);
      const xl = v.getUint16(p + 30, true), cl = v.getUint16(p + 32, true), lo = v.getUint32(p + 42, true);
      const name = dec ? dec.decode(u8.subarray(p + 46, p + 46 + nl)) : String.fromCharCode.apply(null, u8.subarray(p + 46, p + 46 + nl));
      if (v.getUint32(lo, true) !== 0x04034b50) throw new Error('Bad local header for ' + name);
      const lnl = v.getUint16(lo + 26, true), lxl = v.getUint16(lo + 28, true);
      const start = lo + 30 + lnl + lxl;
      out.push({ name, crc, data: u8.subarray(start, start + size) });
      p += 46 + nl + xl + cl;
    }
    return out;
  }

  function toBase64(u8) {
    if (typeof Buffer !== 'undefined' && typeof Buffer.from === 'function') return Buffer.from(u8).toString('base64');
    let s = '';
    for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
    return btoa(s);
  }

  // ---------- workbook parts ----------

  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function localDate(d) { return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  function isSafeUrl(u) { return /^(https?:\/\/|mailto:)[^\s"<>]+$/i.test(String(u || '').trim()); }

  function stylesXml() {
    return XML_DECL + '<styleSheet xmlns="' + NS_MAIN + '">' +
      '<fonts count="3">' +
      '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
      '<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
      '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>' +
      '</fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="5">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
      '<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
      '</styleSheet>';
  }

  function inlineStr(ref, text, style) {
    const t = stripControl(text).slice(0, MAX_CELL);
    return '<c r="' + ref + '"' + (style ? ' s="' + style + '"' : '') + ' t="inlineStr"><is><t xml:space="preserve">' +
      escapeXml(t) + '</t></is></c>';
  }

  // Prepare one sheet: resolved names, cell XML, hyperlinks.
  function buildSheet(sheet, idx, usedTables) {
    const names = uniqueNames(sheet.columns || []);
    const src = sheet.columns || [];
    const nCols = Math.max(1, names.length);
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    const linkCols = new Set(sheet.linkColumns || []);
    const pctCols = new Set(sheet.pctColumns || []);
    const textCols = new Set(sheet.textColumns || []);
    const autoText = sheet.autoTextColumns !== false;
    const linkText = sheet.linkText || {};
    const isText = src.map(c => textCols.has(c) || (autoText && ID_COL_RE.test(String(c || ''))));
    const links = []; // { ref, url }
    const widths = names.map(n => Math.min(60, Math.max(8, String(n).length + 2)));

    let xml = '<row r="1" spans="1:' + nCols + '">';
    names.forEach((n, c) => { xml += inlineStr(cellRef(c + 1, 1), n, STYLE.header); });
    xml += '</row>';

    rows.forEach((r, ri) => {
      const rn = ri + 2;
      let cells = '';
      src.forEach((col, ci) => {
        const ref = cellRef(ci + 1, rn);
        let v = r ? r[col] : undefined;
        if (v === undefined || v === null || v === '') return;
        if (v instanceof Date) v = isNaN(v) ? '' : localDate(v);
        if (linkCols.has(col)) {
          const url = typeof v === 'object' ? v.url : v;
          const text = typeof v === 'object' ? (v.text || v.url) : (linkText[col] || v);
          if (isSafeUrl(url)) {
            links.push({ ref, url: String(url).trim() });
            cells += inlineStr(ref, text, STYLE.link);
            widths[ci] = Math.max(widths[ci], Math.min(60, String(text).length + 2));
            return;
          }
          v = typeof v === 'object' ? (v.text || '') : v;
          if (v === '') return;
        }
        if (typeof v === 'number' && isFinite(v) && !isText[ci]) {
          cells += '<c r="' + ref + '"' + (pctCols.has(col) ? ' s="' + STYLE.pct + '"' : '') + '><v>' + v + '</v></c>';
          widths[ci] = Math.max(widths[ci], Math.min(60, String(v).length + 2));
        } else if (typeof v === 'boolean' && !isText[ci]) {
          cells += '<c r="' + ref + '" t="b"><v>' + (v ? 1 : 0) + '</v></c>';
        } else {
          const s = typeof v === 'object' ? JSON.stringify(v) : String(v);
          cells += inlineStr(ref, s, isText[ci] ? STYLE.text : 0);
          widths[ci] = Math.max(widths[ci], Math.min(60, s.length + 2));
        }
      });
      xml += '<row r="' + rn + '" spans="1:' + nCols + '">' + cells + '</row>';
    });

    // colWidths: array by position or { column: width }.
    const cw = sheet.colWidths;
    if (Array.isArray(cw)) cw.forEach((w, i) => { if (i < nCols && w > 0) widths[i] = w; });
    else if (cw && typeof cw === 'object') src.forEach((c, i) => { if (cw[c] > 0) widths[i] = cw[c]; });

    const ref = tableRef(nCols, rows.length);
    return {
      names, src, nCols, ref, links, widths, rowsXml: xml,
      table: sanitizeTableName(sheet.table || ('Table' + (idx + 1)), usedTables),
      hidden: !!sheet.hidden,
      textColIdx: isText.map((t, i) => t ? i : -1).filter(i => i >= 0),
    };
  }

  function sheetXml(s, selected) {
    const urlIds = new Map();
    s.links.forEach(l => { if (!urlIds.has(l.url)) urlIds.set(l.url, 'rId' + (urlIds.size + 2)); }); // rId1 = table
    let cols = '<cols>';
    s.widths.forEach((w, i) => {
      cols += '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + (Math.round(w * 100) / 100) + '" customWidth="1"' +
        (s.textColIdx.indexOf(i) >= 0 ? ' style="' + STYLE.text + '"' : '') + '/>';
    });
    cols += '</cols>';
    let xml = XML_DECL + '<worksheet xmlns="' + NS_MAIN + '" xmlns:r="' + NS_R + '">' +
      '<dimension ref="' + s.ref + '"/>' +
      '<sheetViews><sheetView ' + (selected ? 'tabSelected="1" ' : '') + 'workbookViewId="0">' +
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' + cols +
      '<sheetData>' + s.rowsXml + '</sheetData>';
    if (s.links.length) {
      xml += '<hyperlinks>' + s.links.map(l => '<hyperlink ref="' + l.ref + '" r:id="' + urlIds.get(l.url) + '"/>').join('') + '</hyperlinks>';
    }
    xml += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
      '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>';
    let rels = XML_DECL + '<Relationships xmlns="' + NS_PKG_REL + '">';
    rels += '<Relationship Id="rId1" Type="' + REL + 'table" Target="../tables/table' + s.tableIndex + '.xml"/>';
    urlIds.forEach((id, url) => {
      rels += '<Relationship Id="' + id + '" Type="' + REL + 'hyperlink" Target="' + escapeXml(url) + '" TargetMode="External"/>';
    });
    rels += '</Relationships>';
    return { xml, rels };
  }

  function tableXml(s) {
    return XML_DECL + '<table xmlns="' + NS_MAIN + '" id="' + s.tableIndex + '" name="' + escapeXml(s.table) +
      '" displayName="' + escapeXml(s.table) + '" ref="' + s.ref + '" totalsRowShown="0">' +
      '<autoFilter ref="' + s.ref + '"/>' +
      '<tableColumns count="' + s.names.length + '">' +
      s.names.map((n, i) => '<tableColumn id="' + (i + 1) + '" name="' + escapeXml(n) + '"/>').join('') +
      '</tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>' +
      '</table>';
  }

  // Deterministic GUID-looking id for the web extension part.
  function guidFrom(seed) {
    let h = '';
    for (let i = 0; h.length < 32; i++) h += crc32(seed + ':' + i).toString(16).padStart(8, '0');
    return '{' + [h.slice(0, 8), h.slice(8, 12), '4' + h.slice(13, 16), '8' + h.slice(17, 20), h.slice(20, 32)].join('-').toUpperCase() + '}';
  }

  function webExtensionParts(we) {
    const ref = '<we:reference id="' + escapeXml(we.id) + '" version="' + escapeXml(we.version || '1.0.0.0') +
      '" store="' + escapeXml(we.store || 'EXCatalog') + '" storeType="' + escapeXml(we.storeType || 'EXCatalog') + '"/>';
    const props = [['Office.AutoShowTaskpaneWithDocument', we.autoShow === false ? 'false' : 'true']]
      .concat(Object.entries(we.properties || {}))
      .map(([k, v]) => '<we:property name="' + escapeXml(k) + '" value="' + escapeXml(v) + '"/>').join('');
    const webext = XML_DECL + '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="' +
      guidFrom(String(we.id)) + '">' + ref + '<we:alternateReferences/>' +
      '<we:properties>' + props + '</we:properties><we:bindings/>' +
      '<we:snapshot xmlns:r="' + NS_R + '"/></we:webextension>';
    const taskpanes = XML_DECL + '<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">' +
      '<wetp:taskpane dockstate="right" visibility="0" width="' + (we.width || 350) + '" row="' + (we.row || 4) + '">' +
      '<wetp:webextensionref xmlns:r="' + NS_R + '" r:id="rId1"/></wetp:taskpane></wetp:taskpanes>';
    const tpRels = XML_DECL + '<Relationships xmlns="' + NS_PKG_REL + '">' +
      '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/>' +
      '</Relationships>';
    return [
      { name: 'xl/webextensions/taskpanes.xml', data: taskpanes },
      { name: 'xl/webextensions/_rels/taskpanes.xml.rels', data: tpRels },
      { name: 'xl/webextensions/webextension1.xml', data: webext },
    ];
  }

  // spec: { sheets: [{ name, columns, rows, table, hidden, colWidths, linkColumns, linkText, pctColumns, textColumns, autoTextColumns }],
  //         webExtension?: { id, version, store, storeType, autoShow, width, properties }, creator?, now? }
  // Returns { files: [{ name, data }] } — the package parts before zipping.
  function buildParts(spec) {
    spec = spec || {};
    const now = spec.now instanceof Date ? spec.now : new Date();
    const input = (spec.sheets || []).filter(Boolean);
    if (!input.length) throw new Error('buildWorkbook needs at least one sheet');
    const usedSheets = new Set(), usedTables = new Set();
    const sheets = input.map((sh, i) => {
      const s = buildSheet(sh, i, usedTables);
      s.name = sanitizeSheetName(sh.name || ('Sheet' + (i + 1)), usedSheets);
      s.tableIndex = i + 1;
      return s;
    });
    const active = sheets.findIndex(s => !s.hidden);
    if (active < 0) throw new Error('At least one sheet must be visible');
    const we = spec.webExtension && spec.webExtension.id ? spec.webExtension : null;

    const files = [];
    let ct = XML_DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
    sheets.forEach((s, i) => {
      ct += '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      ct += '<Override PartName="/xl/tables/table' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>';
    });
    if (we) {
      ct += '<Override PartName="/xl/webextensions/taskpanes.xml" ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>';
      ct += '<Override PartName="/xl/webextensions/webextension1.xml" ContentType="application/vnd.ms-office.webextension+xml"/>';
    }
    ct += '</Types>';
    files.push({ name: '[Content_Types].xml', data: ct });

    files.push({ name: '_rels/.rels', data: XML_DECL + '<Relationships xmlns="' + NS_PKG_REL + '">' +
      '<Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="' + REL + 'extended-properties" Target="docProps/app.xml"/>' +
      (we ? '<Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" Target="xl/webextensions/taskpanes.xml"/>' : '') +
      '</Relationships>' });

    const iso = now.toISOString().replace(/\.\d{3}Z$/, 'Z');
    const creator = escapeXml(spec.creator || 'Sohar Eligibility add-in');
    files.push({ name: 'docProps/core.xml', data: XML_DECL +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<dc:creator>' + creator + '</dc:creator><cp:lastModifiedBy>' + creator + '</cp:lastModifiedBy>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + iso + '</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">' + iso + '</dcterms:modified></cp:coreProperties>' });
    files.push({ name: 'docProps/app.xml', data: XML_DECL +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application></Properties>' });

    let wb = XML_DECL + '<workbook xmlns="' + NS_MAIN + '" xmlns:r="' + NS_R + '"><workbookPr/>' +
      '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300" activeTab="' + active + '"/></bookViews><sheets>';
    sheets.forEach((s, i) => {
      wb += '<sheet name="' + escapeXml(s.name) + '" sheetId="' + (i + 1) + '"' + (s.hidden ? ' state="hidden"' : '') + ' r:id="rId' + (i + 1) + '"/>';
    });
    wb += '</sheets></workbook>';
    files.push({ name: 'xl/workbook.xml', data: wb });

    let wbRels = XML_DECL + '<Relationships xmlns="' + NS_PKG_REL + '">';
    sheets.forEach((s, i) => { wbRels += '<Relationship Id="rId' + (i + 1) + '" Type="' + REL + 'worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>'; });
    wbRels += '<Relationship Id="rId' + (sheets.length + 1) + '" Type="' + REL + 'styles" Target="styles.xml"/></Relationships>';
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: wbRels });
    files.push({ name: 'xl/styles.xml', data: stylesXml() });

    sheets.forEach((s, i) => {
      const x = sheetXml(s, i === active);
      files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: x.xml });
      files.push({ name: 'xl/worksheets/_rels/sheet' + (i + 1) + '.xml.rels', data: x.rels });
      files.push({ name: 'xl/tables/table' + (i + 1) + '.xml', data: tableXml(s) });
    });
    if (we) webExtensionParts(we).forEach(f => files.push(f));
    return { files, sheets: sheets.map(s => ({ name: s.name, table: s.table, ref: s.ref, columns: s.names })) };
  }

  function buildWorkbook(spec) {
    const parts = buildParts(spec);
    return zip(parts.files, spec && spec.now instanceof Date ? spec.now : new Date());
  }

  return {
    MIME, ID_COL_RE,
    buildWorkbook, buildParts, toBase64,
    zip, unzipStored, crc32, utf8,
    escapeXml, stripControl, colLetter, cellRef, tableRef, uniqueNames, sanitizeSheetName, sanitizeTableName,
  };
});
