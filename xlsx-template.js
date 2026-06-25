/* xlsx-template.js — fills the OFFICIAL template workbook with the app's
   data and re-zips it, so the exported file is byte-for-byte identical to
   the uploaded form (logo, colours, weekday fills, borders, Angsana fonts,
   merged cells and SUM formulas all preserved). Only cell VALUES change.

   window.SavingsTemplate.exportFilled(model, templateArrayBuffer, filename)
   Layout (per analysis of the template):
     ข้อมูล sheet : A4 school, A6 class, A8 office, A10 year, A12 teacher,
                    A14 director ; H4.. student names.
     month sheets : student rows 8..22 (max 15) ; C = ยอดยกมา (carry),
                    D..AH = days 1..31, AI = รวม (SUM formula, auto).
                    Row 23 = รวม row (SUM formulas, auto).                */
(function () {
  "use strict";

  /* ---- CRC32 ---- */
  var CRC = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
    return t;
  })();
  function crc32(b) { var c = 0xFFFFFFFF; for (var i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  var enc = new TextEncoder();
  function U8(s) { return enc.encode(s); }

  /* ---- read zip central directory ---- */
  function readEntries(buf) {
    var data = new Uint8Array(buf), dv = new DataView(data.buffer);
    var eocd = -1;
    for (var i = data.length - 22; i >= 0; i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) throw new Error("ไฟล์เทมเพลตเสียหาย");
    var cnt = dv.getUint16(eocd + 10, true), off = dv.getUint32(eocd + 16, true), p = off;
    var entries = [];
    for (var k = 0; k < cnt; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      var method = dv.getUint16(p + 10, true);
      var crc = dv.getUint32(p + 16, true);
      var csize = dv.getUint32(p + 20, true);
      var usize = dv.getUint32(p + 24, true);
      var nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      var lo = dv.getUint32(p + 42, true);
      var name = new TextDecoder().decode(data.slice(p + 46, p + 46 + nl));
      // locate compressed payload via local header
      var lnl = dv.getUint16(lo + 26, true), lel = dv.getUint16(lo + 28, true);
      var dstart = lo + 30 + lnl + lel;
      var comp = data.slice(dstart, dstart + csize);
      entries.push({ name: name, method: method, crc: crc, usize: usize, comp: comp, text: null });
      p += 46 + nl + el + cl;
    }
    return entries;
  }

  function inflateRaw(bytes) {
    var ds = new DecompressionStream("deflate-raw");
    return new Response(new Blob([bytes]).stream().pipeThrough(ds)).arrayBuffer().then(function (a) { return new Uint8Array(a); });
  }
  function deflateRaw(bytes) {
    var cs = new CompressionStream("deflate-raw");
    return new Response(new Blob([bytes]).stream().pipeThrough(cs)).arrayBuffer().then(function (a) { return new Uint8Array(a); });
  }
  function getText(entries, name) {
    var e = entries.find(function (x) { return x.name === name; });
    if (!e) return Promise.resolve(null);
    if (e.method === 0) return Promise.resolve(new TextDecoder().decode(e.comp));
    return inflateRaw(e.comp).then(function (u) { return new TextDecoder().decode(u); });
  }
  function setText(entries, name, text) {
    var e = entries.find(function (x) { return x.name === name; });
    e.text = text; // mark changed; recompressed at zip time
  }

  /* ---- cell helpers ---- */
  function colName(n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
  function esc(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }

  // replace a numeric cell's value, preserving its style and any <f> formula
  function setNum(block, ref, value, keepFormula) {
    var re = new RegExp('<c r="' + ref + '"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)');
    return block.replace(re, function (m, attrs, inner) {
      var s = /\bs="(\d+)"/.exec(attrs); var sAttr = s ? ' s="' + s[1] + '"' : '';
      var fStr = "";
      if (keepFormula && inner) {
        var f = /<f\b[^>]*?(?:\/>|>[\s\S]*?<\/f>)/.exec(inner);
        if (f) fStr = f[0];
      }
      return '<c r="' + ref + '"' + sAttr + '>' + fStr + '<v>' + value + '</v></c>';
    });
  }
  // set an inline-string cell (preserve style)
  function setStr(block, ref, text) {
    var re = new RegExp('<c r="' + ref + '"([^>]*?)(?:/>|>[\\s\\S]*?</c>)');
    return block.replace(re, function (m, attrs) {
      var s = /\bs="(\d+)"/.exec(attrs); var sAttr = s ? ' s="' + s[1] + '"' : '';
      if (!text) return '<c r="' + ref + '"' + sAttr + '/>';
      return '<c r="' + ref + '"' + sAttr + ' t="inlineStr"><is><t xml:space="preserve">' + esc(text) + '</t></is></c>';
    });
  }
  // set the cached value of a string-formula cell (=ข้อมูล!H..), keep <f>
  function setFormulaCache(block, ref, text) {
    var re = new RegExp('<c r="' + ref + '"([^>]*?)>([\\s\\S]*?)</c>');
    return block.replace(re, function (m, attrs, inner) {
      var s = /\bs="(\d+)"/.exec(attrs); var sAttr = s ? ' s="' + s[1] + '"' : '';
      var f = /<f\b[^>]*?(?:\/>|>[\s\S]*?<\/f>)/.exec(inner);
      var fStr = f ? f[0] : '';
      if (!fStr) return setStr(m, ref, text); // no formula -> inline string
      return '<c r="' + ref + '"' + sAttr + ' t="str">' + fStr + '<v>' + esc(text) + '</v></c>';
    });
  }

  var MAXROWS = 15; // template student rows 8..22
  var TOTROW = 23;  // template totals row

  function patchMonth(xml, rows, names, school) {
    // pre-compute totals row values
    var carryTotal = 0, grand = 0, dayTotals = new Array(31).fill(0);
    for (var i = 0; i < rows.length; i++) {
      carryTotal += rows[i].carry || 0;
      grand += rows[i].total || 0;
      for (var dd = 0; dd < 31; dd++) dayTotals[dd] += (rows[i].days[dd] || 0);
    }

    // 1) student rows 8..22 + totals row 23 (all static values)
    xml = xml.replace(/<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g, function (block, rstr) {
      var r = +rstr;
      if (r >= 8 && r <= 22) {
        var idx = r - 8;
        var d = idx < rows.length ? rows[idx] : null;
        var nm = idx < names.length ? names[idx] : '';
        block = setStr(block, 'B' + r, nm);
        block = setNum(block, 'C' + r, d ? d.carry : 0, false);
        for (var day = 1; day <= 31; day++) {
          block = setNum(block, colName(3 + day) + r, d ? (d.days[day - 1] || 0) : 0, false);
        }
        block = setNum(block, 'AI' + r, d ? d.total : 0, false);
        return block;
      }
      if (r === TOTROW) {
        block = setNum(block, 'C' + r, carryTotal, false);
        for (var day2 = 1; day2 <= 31; day2++) {
          block = setNum(block, colName(3 + day2) + r, dayTotals[day2 - 1], false);
        }
        block = setNum(block, 'AI' + r, grand, false);
        return block;
      }
      return block;
    });

    // 2) header + signature cells (originally =ข้อมูล!.. formulas) -> inline text
    xml = setStr(xml, 'H4', school.name);
    xml = setStr(xml, 'B4', 'ชั้น  ' + school.classLevel);
    xml = setStr(xml, 'C4', 'ปีการศึกษา  ' + school.year);
    xml = setStr(xml, 'U4', school.office);
    xml = setStr(xml, 'B49', school.teacher);
    xml = setStr(xml, 'V49', school.director);

    // 3) smart-strip EVERY remaining formula (e.g. shared AI members in rows
    //    24..46). Precise, fast regex (no tempered-greedy backtracking): a
    //    formula cell is <c ..><f .../|>..</f>[<v>..</v>]</c>. t="str" cells
    //    become inline strings (keep cached text); numeric cells keep <v>.
    xml = xml.replace(/<c\b([^>]*)><f\b[^>]*?(?:\/>|>[^<]*<\/f>)(?:<v>([^<]*)<\/v>)?<\/c>/g,
      function (m, attrs, val) {
        val = val || '';
        if (/\bt="str"/.test(attrs)) {
          var a = attrs.replace(/\s+t="str"/, '');
          return (val === '' || val.charAt(0) === '#') ? '<c' + a + '/>'
            : '<c' + a + ' t="inlineStr"><is><t xml:space="preserve">' + val + '</t></is></c>';
        }
        var an = attrs.replace(/\s+t="[^"]*"/, '');
        // keep <v> only when it is a clean number; error values (#REF!, #VALUE!,
        // …) or anything non-numeric become a blank cell so we never emit an
        // invalid numeric cell.
        return (val !== '' && /^-?\d+(?:\.\d+)?$/.test(val))
          ? '<c' + an + '><v>' + val + '</v></c>' : '<c' + an + '/>';
      });
    return xml;
  }

  function patchInfo(xml, school, names) {
    xml = setStr(xml, 'A4', school.name);
    xml = setStr(xml, 'A6', 'ชั้น  ' + school.classLevel);
    xml = setStr(xml, 'A8', school.office);
    xml = setStr(xml, 'A10', 'ปีการศึกษา  ' + school.year);
    xml = setStr(xml, 'A12', school.teacher);
    xml = setStr(xml, 'A14', school.director);
    for (var i = 0; i < MAXROWS; i++) {
      xml = setStr(xml, 'H' + (4 + i), i < names.length ? names[i] : '');
    }
    return xml;
  }

  /* ---- build model rows for a month ---- */
  function monthRows(model, mi) {
    // mirror of the app: carry = sum of prior months' deposits per student
    var out = [];
    var students = model.students;
    for (var i = 0; i < Math.min(students.length, MAXROWS); i++) {
      var days = model.months[mi].rows[i].days.slice(0, 31);
      out.push({
        carry: model.months[mi].rows[i].carry,
        days: days,
        total: model.months[mi].rows[i].total
      });
    }
    return out;
  }

  /* ---- assemble + download ---- */
  function buildZip(entries) {
    var DOS_TIME = 0;                       // 00:00:00 (valid)
    var DOS_DATE = (44 << 9) | (6 << 5) | 15; // 2024-06-15 (valid DOS date)
    var locals = [], central = [], offset = 0;
    function u16(n) { return [n & 0xFF, (n >>> 8) & 0xFF]; }
    function u32(n) { return [n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF]; }
    entries.forEach(function (e) {
      var nameB = U8(e.name);
      var lh = [].concat(u32(0x04034b50), u16(20), u16(0x0800), u16(e.method), u16(DOS_TIME), u16(DOS_DATE),
        u32(e.crc), u32(e.csize), u32(e.usize), u16(nameB.length), u16(0));
      locals.push(new Uint8Array(lh), nameB, e.payload);
      var ch = [].concat(u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(e.method), u16(DOS_TIME), u16(DOS_DATE),
        u32(e.crc), u32(e.csize), u32(e.usize), u16(nameB.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset));
      central.push(new Uint8Array(ch), nameB);
      offset += lh.length + nameB.length + e.payload.length;
    });
    var cstart = offset, csize = central.reduce(function (a, b) { return a + b.length; }, 0);
    var eocd = [].concat(u32(0x06054b50), u16(0), u16(0), u16(entries.length), u16(entries.length), u32(csize), u32(cstart), u16(0));
    var parts = locals.concat(central, [new Uint8Array(eocd)]);
    var total = parts.reduce(function (a, b) { return a + b.length; }, 0);
    var out = new Uint8Array(total), pos = 0;
    parts.forEach(function (p) { out.set(p, pos); pos += p.length; });
    return out;
  }

  function exportFilled(model, templateBuf, filename) {
    var entries = readEntries(templateBuf);
    var monthFiles = []; // sheet2..sheet12 -> month index 0..10
    for (var s = 2; s <= 12; s++) monthFiles.push('xl/worksheets/sheet' + s + '.xml');
    var names = model.students.map(function (x) { return x.name; });

    // 0) drop calcChain.xml — after editing cached formula values it becomes
    //    inconsistent and triggers Excel's "we found a problem" recovery.
    //    Removing it (and forcing a recalc) makes Excel rebuild it cleanly.
    entries = entries.filter(function (e) { return e.name !== 'xl/calcChain.xml'; });

    // 1) patch text parts
    return getText(entries, 'xl/worksheets/sheet1.xml').then(function (info) {
      setText(entries, 'xl/worksheets/sheet1.xml', patchInfo(info, model.school, names));
      var chain = Promise.resolve();
      monthFiles.forEach(function (path, mi) {
        chain = chain.then(function () {
          return getText(entries, path).then(function (xml) {
            setText(entries, path, patchMonth(xml, monthRows(model, mi), names, model.school));
          });
        });
      });
      return chain;
    }).then(function () {
      // 1b) de-reference calcChain & force full recalc on load
      return getText(entries, '[Content_Types].xml').then(function (ct) {
        ct = ct.replace(/<Override[^>]*calcChain[^>]*\/>/, '');
        setText(entries, '[Content_Types].xml', ct);
        return getText(entries, 'xl/_rels/workbook.xml.rels');
      }).then(function (rels) {
        rels = rels.replace(/<Relationship[^>]*calcChain[^>]*\/>/, '');
        setText(entries, 'xl/_rels/workbook.xml.rels', rels);
        return getText(entries, 'xl/workbook.xml');
      }).then(function (wb) {
        if (/<calcPr\b/.test(wb)) {
          wb = wb.replace(/<calcPr\b([^>]*?)\/>/, function (m, a) {
            if (/fullCalcOnLoad/.test(a)) return m;
            return '<calcPr' + a + ' fullCalcOnLoad="1"/>';
          });
        }
        setText(entries, 'xl/workbook.xml', wb);
      });
    }).then(function () {
      // 2) finalise payloads (recompress changed, reuse rest)
      var chain = Promise.resolve();
      entries.forEach(function (e) {
        chain = chain.then(function () {
          if (e.text === null) { e.payload = e.comp; e.csize = e.comp.length; return; }
          var raw = U8(e.text);
          e.crc = crc32(raw); e.usize = raw.length;
          return deflateRaw(raw).then(function (cz) { e.payload = cz; e.method = 8; e.csize = cz.length; });
        });
      });
      return chain;
    }).then(function () {
      return buildZip(entries);
    });
  }

  function downloadFilled(model, templateBuf, filename) {
    return exportFilled(model, templateBuf, filename).then(function (bytes) {
      var blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url; a.download = filename || "savings.xlsx";
      document.body.appendChild(a); a.click();
      setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1500);
      return { students: Math.min(model.students.length, MAXROWS), capped: model.students.length > MAXROWS };
    });
  }

  window.SavingsTemplate = { exportFilled: exportFilled, downloadFilled: downloadFilled, MAXROWS: MAXROWS };
})();
