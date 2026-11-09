/* =========================================================
   XLSX READER — pembaca file Excel (.xlsx) tanpa library luar
   Berisi: inflate (DEFLATE) murni JS, pembaca ZIP, dan parser
   XML workbook/spreadsheet. Dipakai offline oleh panel admin.
   ========================================================= */
(function () {
  'use strict';

  var LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  var LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  var DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  var DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];

  /* ---------------- Inflate (RFC 1951) ---------------- */
  function buildTree(lengths) {
    var counts = new Uint16Array(16);
    var offs = new Uint16Array(16);
    var symbols = new Uint16Array(lengths.length);
    var i;
    for (i = 0; i < lengths.length; i++) counts[lengths[i]]++;
    counts[0] = 0;
    for (i = 1; i < 16; i++) offs[i] = offs[i - 1] + counts[i - 1];
    for (i = 0; i < lengths.length; i++) {
      if (lengths[i]) symbols[offs[lengths[i]]++] = i;
    }
    return { counts: counts, symbols: symbols };
  }

  function inflateRaw(data) {
    var pos = 0;
    var bitBuf = 0;
    var bitCnt = 0;
    var out = new Uint8Array(Math.max(1024, data.length * 4));
    var outLen = 0;

    function ensure(n) {
      if (outLen + n <= out.length) return;
      var next = out.length * 2;
      while (next < outLen + n) next *= 2;
      var neo = new Uint8Array(next);
      neo.set(out.subarray(0, outLen));
      out = neo;
    }

    function bits(n) {
      while (bitCnt < n) {
        bitBuf |= data[pos++] << bitCnt;
        bitCnt += 8;
      }
      var v = bitBuf & ((1 << n) - 1);
      bitBuf >>>= n;
      bitCnt -= n;
      return v;
    }

    function decodeSym(tree) {
      var code = 0, first = 0, index = 0;
      for (var len = 1; len <= 15; len++) {
        code |= bits(1);
        var count = tree.counts[len];
        if (code - first < count) return tree.symbols[index + (code - first)];
        index += count;
        first = (first + count) << 1;
        code <<= 1;
      }
      throw new Error('Data Excel rusak (kode Huffman tidak valid).');
    }

    var last = 0;
    while (!last) {
      last = bits(1);
      var type = bits(2);

      if (type === 0) {
        bitBuf = 0;
        bitCnt = 0;
        var len = data[pos] | (data[pos + 1] << 8);
        pos += 4;
        ensure(len);
        out.set(data.subarray(pos, pos + len), outLen);
        outLen += len;
        pos += len;
        continue;
      }
      if (type === 3) throw new Error('Data Excel rusak (tipe blok tidak dikenal).');

      var lt, dt;
      if (type === 1) {
        var fixed = new Uint8Array(288);
        for (var i = 0; i < 144; i++) fixed[i] = 8;
        for (; i < 256; i++) fixed[i] = 9;
        for (; i < 280; i++) fixed[i] = 7;
        for (; i < 288; i++) fixed[i] = 8;
        lt = buildTree(fixed);
        var dfixed = new Uint8Array(30);
        for (var j = 0; j < 30; j++) dfixed[j] = 5;
        dt = buildTree(dfixed);
      } else {
        var hlit = bits(5) + 257;
        var hdist = bits(5) + 1;
        var hclen = bits(4) + 4;
        var order = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
        var clLens = new Uint8Array(19);
        for (var k = 0; k < hclen; k++) clLens[order[k]] = bits(3);
        var clTree = buildTree(clLens);
        var lens = new Uint8Array(hlit + hdist);
        var idx = 0;
        while (idx < lens.length) {
          var sym = decodeSym(clTree);
          if (sym < 16) {
            lens[idx++] = sym;
          } else if (sym === 16) {
            var prev = idx ? lens[idx - 1] : 0;
            var rep = 3 + bits(2);
            while (rep-- && idx < lens.length) lens[idx++] = prev;
          } else if (sym === 17) {
            var rep17 = 3 + bits(3);
            while (rep17-- && idx < lens.length) lens[idx++] = 0;
          } else {
            var rep18 = 11 + bits(7);
            while (rep18-- && idx < lens.length) lens[idx++] = 0;
          }
        }
        lt = buildTree(lens.subarray(0, hlit));
        dt = buildTree(lens.subarray(hlit));
      }

      for (;;) {
        var s = decodeSym(lt);
        if (s < 256) {
          ensure(1);
          out[outLen++] = s;
        } else if (s === 256) {
          break;
        } else {
          s -= 257;
          if (s >= 29) throw new Error('Data Excel rusak (panjang blok tidak valid).');
          var length = LEN_BASE[s] + bits(LEN_EXTRA[s]);
          var ds = decodeSym(dt);
          if (ds >= 30) throw new Error('Data Excel rusak (jarak blok tidak valid).');
          var dist = DIST_BASE[ds] + bits(DIST_EXTRA[ds]);
          if (dist > outLen) throw new Error('Data Excel rusak (referensi jarak tidak valid).');
          ensure(length);
          for (var m = 0; m < length; m++) {
            out[outLen] = out[outLen - dist];
            outLen++;
          }
        }
      }
    }

    return out.subarray(0, outLen);
  }

  /* ---------------- ZIP ---------------- */
  function u16(d, o) { return d[o] | (d[o + 1] << 8); }
  function u32(d, o) { return (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24)) >>> 0; }

  function unzip(buffer) {
    var data = new Uint8Array(buffer);
    var len = data.length;
    var eocd = -1;
    var min = Math.max(0, len - 65558);
    for (var p = len - 22; p >= min; p--) {
      if (u32(data, p) === 0x06054b50) { eocd = p; break; }
    }
    if (eocd < 0) throw new Error('File bukan arsip Excel (.xlsx) yang valid.');

    var count = u16(data, eocd + 10);
    var cdOff = u32(data, eocd + 16);
    if (!count || cdOff >= len) throw new Error('Struktur file Excel tidak lengkap.');

    var entries = {};
    var offset = cdOff;
    for (var i = 0; i < count; i++) {
      if (u32(data, offset) !== 0x02014b50) throw new Error('Daftar isi file Excel rusak.');
      var method = u16(data, offset + 10);
      var csize = u32(data, offset + 20);
      var nameLen = u16(data, offset + 28);
      var extraLen = u16(data, offset + 30);
      var commentLen = u16(data, offset + 32);
      var localOff = u32(data, offset + 42);
      var name = decodeText(data.subarray(offset + 46, offset + 46 + nameLen));
      entries[name] = { method: method, csize: csize, localOff: localOff };
      offset += 46 + nameLen + extraLen + commentLen;
    }

    function read(name) {
      var e = entries[name];
      if (!e) return null;
      var lo = e.localOff;
      if (u32(data, lo) !== 0x04034b50) throw new Error('Entri file Excel rusak: ' + name);
      var nLen = u16(data, lo + 26);
      var xLen = u16(data, lo + 28);
      var start = lo + 30 + nLen + xLen;
      var raw = data.subarray(start, start + e.csize);
      if (e.method === 0) return raw;
      if (e.method === 8) return inflateRaw(raw);
      throw new Error('Metode kompresi Excel tidak didukung: ' + e.method);
    }

    return { names: Object.keys(entries), read: read };
  }

  var textDecoder = typeof TextDecoder !== 'undefined' ? new TextDecoder('utf-8') : null;
  function decodeText(bytes) {
    if (textDecoder) return textDecoder.decode(bytes);
    var s = '';
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return decodeURIComponent(escape(s));
  }

  /* ---------------- XML helpers ---------------- */
  function parseXml(bytes, label) {
    var text = decodeText(bytes);
    var doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('XML ' + label + ' tidak valid.');
    return doc;
  }
  function elements(root, localName) {
    var out = [];
    var all = root.getElementsByTagName('*');
    for (var i = 0; i < all.length; i++) {
      if ((all[i].localName || all[i].nodeName) === localName) out.push(all[i]);
    }
    return out;
  }
  function childText(el, localName) {
    if (!el) return '';
    var kids = el.getElementsByTagName('*');
    for (var i = 0; i < kids.length; i++) {
      if ((kids[i].localName || kids[i].nodeName) === localName) return kids[i].textContent || '';
    }
    return '';
  }
  function relId(el) {
    return el.getAttribute('r:id') || el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || '';
  }
  function resolve(base, target) {
    if (!target) return '';
    if (target.charAt(0) === '/') return target.slice(1);
    var parts = base.split('/').slice(0, -1).concat(target.split('/'));
    var out = [];
    parts.forEach(function (part) {
      if (part === '..') out.pop();
      else if (part !== '.' && part !== '') out.push(part);
    });
    return out.join('/');
  }

  function sharedStrings(zip) {
    var bytes = zip.read('xl/sharedStrings.xml');
    if (!bytes) return [];
    var doc = parseXml(bytes, 'sharedStrings');
    var sis = elements(doc, 'si');
    return sis.map(function (si) {
      var ts = elements(si, 't');
      return ts.map(function (t) { return t.textContent || ''; }).join('');
    });
  }

  var BUILTIN_DATE = {};
  [14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]
    .forEach(function (id) { BUILTIN_DATE[id] = true; });

  function dateStyles(zip) {
    var bytes = zip.read('xl/styles.xml');
    if (!bytes) return [];
    var doc = parseXml(bytes, 'styles');
    var custom = {};
    elements(doc, 'numFmt').forEach(function (nf) {
      custom[nf.getAttribute('numFmtId')] = nf.getAttribute('formatCode') || '';
    });
    var xfs = elements(doc, 'cellXfs');
    if (!xfs.length) return [];
    var list = elements(xfs[0], 'xf');
    return list.map(function (xf) {
      var id = +(xf.getAttribute('numFmtId') || 0);
      if (BUILTIN_DATE[id]) return true;
      var code = custom[id];
      if (!code) return false;
      var clean = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
      return /[ymd]/i.test(clean) && /[\/\-.]|\by\b|\bm\b|\bd\b/i.test(clean);
    });
  }

  function columnIndex(ref) {
    var letters = String(ref || '').replace(/[^A-Za-z]/g, '').toUpperCase();
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }
  function rowNumber(ref) {
    var m = String(ref || '').match(/\d+/);
    return m ? parseInt(m[0], 10) : 0;
  }

  function isoFromSerial(serial) {
    if (!isFinite(serial) || serial < 1 || serial > 2958465) return '';
    var ms = Math.round((serial - 25569) * 86400000);
    var d = new Date(ms);
    if (isNaN(d.getTime())) return '';
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  }

  function parseSheet(bytes, strings, styles) {
    var doc = parseXml(bytes, 'worksheet');
    var sheetData = elements(doc, 'sheetData');
    if (!sheetData.length) return [];
    var rowEls = elements(sheetData[0], 'row');
    var rowMap = {};
    var maxRow = 0;
    rowEls.forEach(function (rowEl) {
      var r = rowNumber(rowEl.getAttribute('r')) || (maxRow + 1);
      var cells = [];
      var cellEls = elements(rowEl, 'c');
      cellEls.forEach(function (c) {
        var idx = columnIndex(c.getAttribute('r'));
        if (!isFinite(idx) || idx < 0) idx = cells.length;
        var type = c.getAttribute('t');
        var style = c.getAttribute('s');
        var value = null;
        var vEl = null;
        var cKids = c.getElementsByTagName('*');
        for (var i = 0; i < cKids.length; i++) {
          if ((cKids[i].localName || cKids[i].nodeName) === 'v') { vEl = cKids[i]; break; }
        }
        if (type === 's') {
          var si = vEl ? +vEl.textContent : -1;
          value = (si >= 0 && strings[si] !== undefined) ? strings[si] : '';
        } else if (type === 'inlineStr') {
          var texts = elements(c, 't');
          value = texts.map(function (t) { return t.textContent || ''; }).join('');
        } else if (type === 'str') {
          value = vEl ? (vEl.textContent || '') : '';
        } else if (type === 'b') {
          value = vEl && vEl.textContent === '1' ? 'TRUE' : 'FALSE';
        } else if (type === 'e') {
          value = '';
        } else if (vEl) {
          var raw = (vEl.textContent || '').trim();
          if (raw === '') {
            value = '';
          } else {
            var num = parseFloat(raw);
            if (isFinite(num) && /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(raw)) {
              if (style !== null && styles[+style]) value = isoFromSerial(num);
              else value = num;
            } else {
              value = raw;
            }
          }
        }
        if (value !== null && value !== '') cells[idx] = value;
      });
      if (cells.length) {
        rowMap[r] = cells;
        if (r > maxRow) maxRow = r;
      }
    });

    var rows = [];
    for (var rn = 1; rn <= maxRow; rn++) {
      var row = rowMap[rn] || [];
      var dense = [];
      for (var c = 0; c < row.length; c++) dense[c] = row[c] === undefined ? '' : row[c];
      rows.push(dense);
    }
    while (rows.length && rows[rows.length - 1].every(function (v) { return v === '' || v === undefined; })) rows.pop();
    return rows;
  }

  /* ---------------- Public API ---------------- */
  function read(buffer) {
    var zip = unzip(buffer);

    var strings = sharedStrings(zip);
    var styles = dateStyles(zip);

    var relsDoc = zip.read('xl/_rels/workbook.xml.rels');
    var relMap = {};
    if (relsDoc) {
      var rdoc = parseXml(relsDoc, 'workbook rels');
      elements(rdoc, 'Relationship').forEach(function (rel) {
        var type = rel.getAttribute('Type') || '';
        if (/worksheet/i.test(type)) relMap[rel.getAttribute('Id')] = rel.getAttribute('Target');
      });
    }

    var workbookDoc = zip.read('xl/workbook.xml');
    var sheetDefs = [];
    if (workbookDoc) {
      var wdoc = parseXml(workbookDoc, 'workbook');
      elements(wdoc, 'sheet').forEach(function (s) {
        sheetDefs.push({ name: s.getAttribute('name') || 'Sheet', rid: relId(s) });
      });
    }
    if (!sheetDefs.length) {
      zip.names.filter(function (n) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(n); })
        .sort()
        .forEach(function (n, i) { sheetDefs.push({ name: 'Sheet' + (i + 1), path: n }); });
    }

    var sheets = [];
    sheetDefs.forEach(function (s) {
      var path = s.path || resolve('xl/workbook.xml', relMap[s.rid] || '');
      if (!path || !zip.read(path)) return;
      var rows = parseSheet(zip.read(path), strings, styles);
      if (rows.length) sheets.push({ name: s.name, rows: rows });
    });

    if (!sheets.length) throw new Error('Tidak ada sheet berisi data pada file Excel ini.');
    return { sheets: sheets };
  }

  /* ---------------- CSV ---------------- */
  function parseDelimited(text, forced) {
    text = String(text || '').replace(/^\ufeff/, '');
    var firstLine = text.split(/\r?\n/)[0] || '';
    var delimiter = forced;
    if (!delimiter) {
      var cands = [';', ',', '\t'];
      var best = ',';
      var bestCount = -1;
      cands.forEach(function (c) {
        var n = firstLine.split(c).length;
        if (n > bestCount) { bestCount = n; best = c; }
      });
      delimiter = best;
    }
    var rows = [];
    var row = [];
    var field = '';
    var quoted = false;
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (quoted) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; }
          else quoted = false;
        } else {
          field += ch;
        }
      } else if (ch === '"') {
        quoted = true;
      } else if (ch === delimiter) {
        row.push(field); field = '';
      } else if (ch === '\n') {
        row.push(field); field = '';
        if (row.length > 1 || row[0] !== '') rows.push(row);
        row = [];
      } else if (ch === '\r') {
        /* abaikan */
      } else {
        field += ch;
      }
    }
    row.push(field);
    if (row.length > 1 || row[0] !== '') rows.push(row);
    return rows;
  }

  window.XLSXReader = { read: read, parseDelimited: parseDelimited, inflateRaw: inflateRaw };
})();
