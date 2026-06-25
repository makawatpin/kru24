/* xlsx-export.js — minimal styled .xlsx writer for the savings book.
   No dependencies. Builds a multi-sheet workbook with borders, merged
   cells, a Thai font and pastel header fills, then triggers a download.
   Exposes window.SavingsXlsx.download(model, filename).            */
(function () {
  "use strict";

  /* ---------- CRC32 + STORE-method ZIP ---------- */
  var CRC_TABLE = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function strBytes(s) { return new TextEncoder().encode(s); }

  function zip(files) {
    // files: [{name, data(Uint8Array)}]
    var locals = [], central = [], offset = 0;
    function u16(n) { return [n & 0xFF, (n >>> 8) & 0xFF]; }
    function u32(n) { return [n & 0xFF, (n >>> 8) & 0xFF, (n >>> 16) & 0xFF, (n >>> 24) & 0xFF]; }
    files.forEach(function (f) {
      var nameB = strBytes(f.name);
      var crc = crc32(f.data);
      var size = f.data.length;
      var lh = [].concat(
        u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(0), u16(0),
        u32(crc), u32(size), u32(size), u16(nameB.length), u16(0)
      );
      locals.push(new Uint8Array(lh), nameB, f.data);
      var ch = [].concat(
        u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(0),
        u32(crc), u32(size), u32(size), u16(nameB.length), u16(0), u16(0), u16(0), u16(0),
        u32(0), u32(offset)
      );
      central.push(new Uint8Array(ch), nameB);
      offset += lh.length + nameB.length + size;
    });
    var centralStart = offset;
    var centralSize = central.reduce(function (a, b) { return a + b.length; }, 0);
    var eocd = [].concat(
      u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length),
      u32(centralSize), u32(centralStart), u16(0)
    );
    var parts = locals.concat(central, [new Uint8Array(eocd)]);
    var total = parts.reduce(function (a, b) { return a + b.length; }, 0);
    var out = new Uint8Array(total), pos = 0;
    parts.forEach(function (p) { out.set(p, pos); pos += p.length; });
    return out;
  }

  /* ---------- helpers ---------- */
  function colName(n) { // 1-based -> A, B, ... AA
    var s = "";
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }
  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function cell(ref, style, value, isNum) {
    if (value === "" || value === null || value === undefined) {
      return '<c r="' + ref + '" s="' + style + '"/>';
    }
    if (isNum) return '<c r="' + ref + '" s="' + style + '"><v>' + value + '</v></c>';
    return '<c r="' + ref + '" s="' + style + '" t="inlineStr"><is><t xml:space="preserve">' + esc(value) + '</t></is></c>';
  }

  /* ---------- styles.xml ----------
     Style indexes (cellXfs):
     0 base | 1 title | 2 subtitle | 3 header-fill | 4 no-cell | 5 name-cell
     6 num-cell | 7 carry-cell | 8 total-cell | 9 total-row | 10 holiday-cell
     11 sign-text | 12 sign-name                                              */
  var STYLES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="6">' +
      '<font><sz val="14"/><name val="TH Sarabun New"/></font>' +                                 // 0
      '<font><b/><sz val="18"/><name val="TH Sarabun New"/></font>' +                              // 1 title
      '<font><sz val="15"/><name val="TH Sarabun New"/></font>' +                                  // 2 subtitle
      '<font><b/><sz val="14"/><name val="TH Sarabun New"/></font>' +                              // 3 header
      '<font><sz val="13"/><name val="TH Sarabun New"/></font>' +                                  // 4 small num
      '<font><b/><sz val="14"/><color rgb="FF2E7D6B"/><name val="TH Sarabun New"/></font>' +       // 5 total
    '</fonts>' +
    '<fills count="5">' +
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFD8F0E8"/><bgColor indexed="64"/></patternFill></fill>' + // 2 header mint
      '<fill><patternFill patternType="solid"><fgColor rgb="FFEFEAE0"/><bgColor indexed="64"/></patternFill></fill>' + // 3 holiday gray
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFBF6E8"/><bgColor indexed="64"/></patternFill></fill>' + // 4 total row cream
    '</fills>' +
    '<borders count="2">' +
      '<border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"><color rgb="FF9FB8B0"/></left><right style="thin"><color rgb="FF9FB8B0"/></right><top style="thin"><color rgb="FF9FB8B0"/></top><bottom style="thin"><color rgb="FF9FB8B0"/></bottom></border>' +
    '</borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="13">' +
      '<xf xfId="0" fontId="0" fillId="0" borderId="0"/>' +                                                                                   // 0 base
      '<xf xfId="0" fontId="1" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +          // 1 title
      '<xf xfId="0" fontId="2" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +          // 2 subtitle
      '<xf xfId="0" fontId="3" fillId="2" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' + // 3 header
      '<xf xfId="0" fontId="0" fillId="0" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +          // 4 no
      '<xf xfId="0" fontId="0" fillId="0" borderId="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +            // 5 name
      '<xf xfId="0" fontId="4" fillId="0" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +          // 6 num
      '<xf xfId="0" fontId="0" fillId="0" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +          // 7 carry
      '<xf xfId="0" fontId="5" fillId="0" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +          // 8 total
      '<xf xfId="0" fontId="5" fillId="4" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +          // 9 total row
      '<xf xfId="0" fontId="4" fillId="3" borderId="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +          // 10 holiday
      '<xf xfId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +          // 11 sign text
      '<xf xfId="0" fontId="3" fillId="4" borderId="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +            // 12 name total
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '</styleSheet>';

  /* ---------- one month worksheet ---------- */
  function monthSheet(model, m) {
    // m: {short, full, yearLabel, dayCount, rows:[{no,name,carry,days[31],total}],
    //     colTotals[31], carryTotal, dayTotal, grand, holidays:Set}
    var DAYS = 31;
    var rows = [];
    var lastCol = colName(3 + DAYS + 1); // C(3)+31 days + total = col 35 = AI
    var rT = "1:" + (8 + model.students.length); // not used; dimension below

    // column widths
    var cols = '<cols>' +
      '<col min="1" max="1" width="4.5" customWidth="1"/>' +
      '<col min="2" max="2" width="24" customWidth="1"/>' +
      '<col min="3" max="3" width="8" customWidth="1"/>' +
      '<col min="4" max="' + (3 + DAYS) + '" width="3.6" customWidth="1"/>' +
      '<col min="' + (4 + DAYS) + '" max="' + (4 + DAYS) + '" width="8.5" customWidth="1"/>' +
    '</cols>';

    function r(n, cells, ht) {
      return '<row r="' + n + '"' + (ht ? ' ht="' + ht + '" customHeight="1"' : '') + '>' + cells + '</row>';
    }

    // R1 title
    rows.push(r(1, cell("A1", 1, "บันทึกการออมเงินของนักเรียน", false), 26));
    // R2 school / class / year
    rows.push(r(2, cell("A2", 2,
      "โรงเรียน" + model.school.name + "    ชั้น " + model.school.classLevel + "    ปีการศึกษา " + model.school.year, false), 20));
    // R3 office
    rows.push(r(3, cell("A3", 2, model.school.office, false), 20));
    // R4 month
    rows.push(r(4, cell("A4", 2, "ประจำเดือน " + m.full + " " + m.yearLabel, false), 20));

    // header R5
    var h5 = cell("A5", 3, "ที่", false) + cell("B5", 3, "ชื่อ - สกุล", false) + cell("C5", 3, "ยอดยกมา", false) +
      cell("D5", 3, "วันที่", false);
    for (var c = 5; c <= 3 + DAYS; c++) h5 += cell(colName(c) + "5", 3, "", false);
    h5 += cell(colName(4 + DAYS) + "5", 3, "รวม", false);
    rows.push(r(5, h5, 22));
    // header R6 day numbers
    var h6 = cell("A6", 3, "", false) + cell("B6", 3, "", false) + cell("C6", 3, "", false);
    for (var d = 1; d <= DAYS; d++) h6 += cell(colName(3 + d) + "6", 3, d, true);
    h6 += cell(colName(4 + DAYS) + "6", 3, "", false);
    rows.push(r(6, h6, 16));

    // student rows from R7
    var rowNum = 7;
    m.rows.forEach(function (st) {
      var line = cell("A" + rowNum, 4, st.no, true) +
        cell("B" + rowNum, 5, st.name, false) +
        cell("C" + rowNum, 7, st.carry ? st.carry : "", true);
      for (var dd = 1; dd <= DAYS; dd++) {
        var styleIdx = m.holidays.has(dd) ? 10 : 6;
        var v = st.days[dd - 1];
        line += cell(colName(3 + dd) + rowNum, styleIdx, (v ? v : ""), true);
      }
      line += cell(colName(4 + DAYS) + rowNum, 8, st.total, true);
      rows.push(r(rowNum, line, 19));
      rowNum++;
    });

    // totals row
    var tline = cell("A" + rowNum, 9, "", false) + cell("B" + rowNum, 12, "รวม", false) +
      cell("C" + rowNum, 9, m.carryTotal, true);
    for (var dt = 1; dt <= DAYS; dt++) {
      tline += cell(colName(3 + dt) + rowNum, 9, (m.colTotals[dt - 1] ? m.colTotals[dt - 1] : ""), true);
    }
    tline += cell(colName(4 + DAYS) + rowNum, 9, m.grand, true);
    rows.push(r(rowNum, tline, 20));
    var totalsRow = rowNum;
    rowNum += 2;

    // signature block
    var sgnA = rowNum, sgnB = rowNum + 1, sgnC = rowNum + 2;
    rows.push(r(sgnA,
      cell("B" + sgnA, 11, "ลงชื่อ.....................................................ผู้รับเงิน", false) +
      cell("U" + sgnA, 11, "ลงชื่อ.....................................................ผู้รับรอง", false), 22));
    rows.push(r(sgnB,
      cell("B" + sgnB, 11, "( " + model.school.teacher + " )", false) +
      cell("U" + sgnB, 11, "( " + model.school.director + " )", false), 20));
    rows.push(r(sgnC,
      cell("B" + sgnC, 11, "ครูประจำชั้น", false) +
      cell("U" + sgnC, 11, "ผู้อำนวยการ" + model.school.name, false), 20));

    var merges = [
      "A1:" + lastCol + "1", "A2:" + lastCol + "2", "A3:" + lastCol + "3", "A4:" + lastCol + "4",
      "A5:A6", "B5:B6", "C5:C6", "D5:" + colName(3 + DAYS) + "5", colName(4 + DAYS) + "5:" + colName(4 + DAYS) + "6",
      "B" + sgnA + ":" + "I" + sgnA, "U" + sgnA + ":AB" + sgnA,
      "B" + sgnB + ":I" + sgnB, "U" + sgnB + ":AB" + sgnB,
      "B" + sgnC + ":I" + sgnC, "U" + sgnC + ":AB" + sgnC
    ];
    var mergeXml = '<mergeCells count="' + merges.length + '">' +
      merges.map(function (x) { return '<mergeCell ref="' + x + '"/>'; }).join("") + '</mergeCells>';

    var dim = "A1:" + lastCol + sgnC;
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="' + dim + '"/>' +
      '<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="18"/>' +
      cols +
      '<sheetData>' + rows.join("") + '</sheetData>' +
      mergeXml +
      '<pageMargins left="0.2" right="0.2" top="0.3" bottom="0.3" header="0.2" footer="0.2"/>' +
      '<pageSetup paperSize="9" orientation="landscape" scale="70"/>' +
      '</worksheet>';
  }

  /* ---------- info (ข้อมูล) sheet ---------- */
  function infoSheet(model) {
    var rows = [], n = 1;
    function r(num, cells, ht) { return '<row r="' + num + '"' + (ht ? ' ht="' + ht + '" customHeight="1"' : '') + '>' + cells + '</row>'; }
    rows.push(r(n, cell("A" + n, 1, "ข้อมูลพื้นฐาน", false), 26)); n += 2;
    var info = [
      ["ชื่อโรงเรียน", model.school.name],
      ["ระดับชั้น", model.school.classLevel],
      ["ปีการศึกษา", model.school.year],
      ["สังกัด", model.school.office],
      ["ครูประจำชั้น", model.school.teacher],
      ["ผู้อำนวยการสถานศึกษา", model.school.director]
    ];
    info.forEach(function (p) {
      rows.push(r(n, cell("A" + n, 3, p[0], false) + cell("B" + n, 5, p[1], false), 20)); n++;
    });
    n++;
    rows.push(r(n, cell("A" + n, 1, "รายชื่อนักเรียน", false), 24)); n++;
    rows.push(r(n, cell("A" + n, 3, "ที่", false) + cell("B" + n, 3, "ชื่อ - สกุล", false) + cell("C" + n, 3, "รวมทั้งปี (บาท)", false), 20)); n++;
    model.students.forEach(function (s, i) {
      rows.push(r(n, cell("A" + n, 4, i + 1, true) + cell("B" + n, 5, s.name, false) + cell("C" + n, 8, model.yearTotals[i], true), 19)); n++;
    });
    rows.push(r(n, cell("A" + n, 9, "", false) + cell("B" + n, 12, "รวมทั้งหมด", false) + cell("C" + n, 9, model.grandYear, true), 20));
    var cols = '<cols><col min="1" max="1" width="6" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/><col min="3" max="3" width="16" customWidth="1"/></cols>';
    var merges = ["A1:C1", "B" + (n) + ":B" + n];
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<dimension ref="A1:C' + n + '"/>' +
      '<sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="18"/>' + cols +
      '<sheetData>' + rows.join("") + '</sheetData>' +
      '<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>' +
      '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
      '</worksheet>';
  }

  /* ---------- assemble workbook ---------- */
  function build(model) {
    var sheets = [{ name: "ข้อมูล", xml: infoSheet(model) }];
    model.months.forEach(function (m) { sheets.push({ name: m.short, xml: monthSheet(model, m) }); });

    var sheetTags = sheets.map(function (s, i) {
      return '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
    }).join("");
    var workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets>' + sheetTags + '</sheets></workbook>';

    var wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map(function (s, i) {
        return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
      }).join("") +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>';

    var contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map(function (s, i) {
        return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      }).join("") +
      '</Types>';

    var rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>';

    var files = [
      { name: "[Content_Types].xml", data: strBytes(contentTypes) },
      { name: "_rels/.rels", data: strBytes(rootRels) },
      { name: "xl/workbook.xml", data: strBytes(workbook) },
      { name: "xl/_rels/workbook.xml.rels", data: strBytes(wbRels) },
      { name: "xl/styles.xml", data: strBytes(STYLES) }
    ];
    sheets.forEach(function (s, i) {
      files.push({ name: "xl/worksheets/sheet" + (i + 1) + ".xml", data: strBytes(s.xml) });
    });
    return zip(files);
  }

  function download(model, filename) {
    var bytes = build(model);
    var blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = filename || "savings.xlsx";
    document.body.appendChild(a); a.click();
    setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1500);
  }

  /* ================= IMPORT (read .xlsx) ================= */
  function inflateRaw(bytes) {
    if (typeof DecompressionStream === "undefined") {
      return Promise.reject(new Error("เบราว์เซอร์นี้ไม่รองรับการอ่านไฟล์ (DecompressionStream)"));
    }
    var ds = new DecompressionStream("deflate-raw");
    var stream = new Blob([bytes]).stream().pipeThrough(ds);
    return new Response(stream).arrayBuffer().then(function (ab) { return new Uint8Array(ab); });
  }

  function unzip(buf) {
    // returns Promise<{ name: string -> string(utf8) }>
    var data = new Uint8Array(buf);
    var dv = new DataView(data.buffer);
    var eocd = -1;
    for (var i = data.length - 22; i >= 0; i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) return Promise.reject(new Error("ไฟล์ไม่ใช่ .xlsx ที่ถูกต้อง"));
    var cnt = dv.getUint16(eocd + 10, true), off = dv.getUint32(eocd + 16, true);
    var entries = [], p = off;
    for (var k = 0; k < cnt; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      var method = dv.getUint16(p + 10, true);
      var compSize = dv.getUint32(p + 20, true);
      var nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      var lo = dv.getUint32(p + 42, true);
      var name = new TextDecoder().decode(data.slice(p + 46, p + 46 + nl));
      entries.push({ name: name, method: method, compSize: compSize, lo: lo });
      p += 46 + nl + el + cl;
    }
    var out = {};
    var chain = Promise.resolve();
    entries.forEach(function (e) {
      chain = chain.then(function () {
        var lnl = dv.getUint16(e.lo + 26, true), lel = dv.getUint16(e.lo + 28, true);
        var dstart = e.lo + 30 + lnl + lel;
        var raw = data.slice(dstart, dstart + e.compSize);
        if (e.method === 0) { out[e.name] = new TextDecoder().decode(raw); return; }
        return inflateRaw(raw).then(function (u) { out[e.name] = new TextDecoder().decode(u); });
      });
    });
    return chain.then(function () { return out; });
  }

  function parseSharedStrings(xml) {
    var arr = [];
    if (!xml) return arr;
    var re = /<si>([\s\S]*?)<\/si>/g, m;
    while ((m = re.exec(xml))) {
      var tre = /<t[^>]*>([\s\S]*?)<\/t>/g, tm, s = "";
      while ((tm = tre.exec(m[1]))) s += tm[1];
      arr.push(unesc(s));
    }
    return arr;
  }
  function unesc(s) {
    return String(s).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&apos;/g, "'").replace(/&amp;/g, "&");
  }

  function parseSheet(xml, shared) {
    // returns map "A1" -> string value
    var cells = {};
    var cre = /<c\s+r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, cm;
    while ((cm = cre.exec(xml))) {
      var ref = cm[1], attrs = cm[2] || "", inner = cm[3] || "";
      var tM = /\bt="([^"]+)"/.exec(attrs);
      var t = tM ? tM[1] : null;
      var val = "";
      if (t === "inlineStr") {
        var im = /<t[^>]*>([\s\S]*?)<\/t>/.exec(inner);
        if (im) val = unesc(im[1]);
      } else {
        var vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
        if (vm) {
          if (t === "s") val = shared[parseInt(vm[1], 10)] || "";
          else if (t === "str") val = unesc(vm[1]);
          else val = vm[1];
        }
      }
      if (val !== "") cells[ref] = val;
    }
    return cells;
  }

  function colRow(ref) {
    var m = /^([A-Z]+)(\d+)$/.exec(ref);
    return { col: m[1], row: parseInt(m[2], 10) };
  }

  /* Resolve the worksheet file whose tab name is `wanted` (default ข้อมูล),
     falling back to the first sheet. Returns parsed cell map.            */
  function readInfoSheet(files) {
    var shared = parseSharedStrings(files["xl/sharedStrings.xml"]);
    var wb = files["xl/workbook.xml"] || "";
    var rels = files["xl/_rels/workbook.xml.rels"] || "";
    // map rId -> target
    var relMap = {};
    var rre = /<Relationship\s+[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*\/>/g, rm;
    while ((rm = rre.exec(rels))) relMap[rm[1]] = rm[2].replace(/^\/?xl\//, "").replace(/^\//, "");
    // sheets in order
    var sheets = [];
    var sre = /<sheet\s+([^>]*?)\/>/g, sm;
    while ((sm = sre.exec(wb))) {
      var a = sm[1];
      var nameM = /name="([^"]*)"/.exec(a);
      var ridM = /r:id="([^"]+)"/.exec(a);
      sheets.push({ name: nameM ? unesc(nameM[1]) : "", rid: ridM ? ridM[1] : null });
    }
    var pick = sheets.find(function (s) { return s.name.indexOf("ข้อมูล") >= 0; }) || sheets[0];
    var target = pick && pick.rid && relMap[pick.rid] ? relMap[pick.rid] : "worksheets/sheet1.xml";
    var path = "xl/" + target.replace(/^xl\//, "");
    var xml = files[path] || files["xl/worksheets/sheet1.xml"];
    return parseSheet(xml || "", shared);
  }

  /* Extract students (ลำดับที่ col G, ชื่อ-สกุล col H, from row 4) from the
     ข้อมูล sheet of an uploaded workbook matching the official form.       */
  function readStudents(buf) {
    return unzip(buf).then(function (files) {
      var cells = readInfoSheet(files);
      // find which columns hold "ที่" and "ชื่อ - สกุล" by scanning header row
      var numCol = "G", nameCol = "H", headerRow = 3;
      for (var ref in cells) {
        var v = cells[ref];
        if (typeof v !== "string") continue;
        if (v.replace(/\s/g, "") === "ที่") { var cr = colRow(ref); numCol = cr.col; headerRow = cr.row; }
        else if (v.indexOf("ชื่อ") >= 0 && v.indexOf("สกุล") >= 0) { nameCol = colRow(ref).col; }
      }
      var students = [];
      var blanks = 0;
      for (var r = headerRow + 1; r <= headerRow + 200; r++) {
        var name = cells[nameCol + r];
        var no = cells[numCol + r];
        if (name && String(name).trim()) {
          students.push({ no: no ? String(no).trim() : String(students.length + 1), name: String(name).trim() });
          blanks = 0;
        } else {
          blanks++;
          if (blanks >= 15 && students.length) break;
        }
      }
      // also pull school info (best-effort) from label/value pairs in col A
      var school = {};
      var labels = { "ชื่อโรงเรียน": "name", "ระดับชั้น": "classLevel", "ปีการศึกษา": "year", "สังกัด": "office", "ครูประจำชั้น": "teacher", "ผู้อำนวยการ": "director" };
      var colA = [];
      for (var ref2 in cells) { var crr = colRow(ref2); if (crr.col === "A") colA.push({ row: crr.row, v: cells[ref2] }); }
      colA.sort(function (a, b) { return a.row - b.row; });
      for (var i = 0; i < colA.length; i++) {
        var txt = String(colA[i].v).trim();
        for (var lab in labels) {
          if (txt.indexOf(lab) === 0 || txt === lab) {
            // value is in the next non-empty A cell below
            for (var j = i + 1; j < colA.length; j++) {
              var nv = String(colA[j].v).trim();
              var isLabel = Object.keys(labels).some(function (L) { return nv.indexOf(L) === 0; });
              if (nv && !isLabel) { school[labels[lab]] = nv; break; }
            }
          }
        }
      }
      return { students: students, school: school };
    });
  }

  window.SavingsXlsx = { build: build, download: download, readStudents: readStudents };
})();
