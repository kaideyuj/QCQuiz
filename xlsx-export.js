(function initQCQuizExcel(global) {
  "use strict";

  const encoder = new TextEncoder();

  function xmlEscape(value) {
    return String(value)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&apos;");
  }

  function columnName(index) {
    let value = index + 1;
    let name = "";
    while (value > 0) {
      value -= 1;
      name = String.fromCharCode(65 + (value % 26)) + name;
      value = Math.floor(value / 26);
    }
    return name;
  }

  function displayWidth(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return 20;
    }
    return [...String(value ?? "")].reduce(
      (width, char) => width + (char.codePointAt(0) > 255 ? 2 : 1),
      0
    );
  }

  function cellXml(value, rowIndex, columnIndex, styleIndex = 0) {
    const reference = `${columnName(columnIndex)}${rowIndex + 1}`;
    const isDate = value instanceof Date && !Number.isNaN(value.getTime());
    const finalStyleIndex = isDate ? 2 : styleIndex;
    const style = finalStyleIndex ? ` s="${finalStyleIndex}"` : "";

    if (isDate) {
      const localTimestamp = Date.UTC(
        value.getFullYear(),
        value.getMonth(),
        value.getDate(),
        value.getHours(),
        value.getMinutes(),
        value.getSeconds()
      );
      const excelSerial = (localTimestamp / 86400000) + 25569;
      return `<c r="${reference}"${style}><v>${excelSerial}</v></c>`;
    }

    if (typeof value === "number" && Number.isFinite(value)) {
      return `<c r="${reference}"${style}><v>${value}</v></c>`;
    }
    if (typeof value === "boolean") {
      return `<c r="${reference}" t="b"${style}><v>${value ? 1 : 0}</v></c>`;
    }
    if (value === null || value === undefined || value === "") {
      return `<c r="${reference}"${style}/>`;
    }

    return `<c r="${reference}" t="inlineStr"${style}><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
  }

  function worksheetXml(sheet) {
    const headers = Array.isArray(sheet.headers) ? sheet.headers : [];
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    const allRows = [headers, ...rows];
    const columnCount = Math.max(1, headers.length, ...rows.map((row) => row.length));
    const lastCell = `${columnName(columnCount - 1)}${Math.max(1, allRows.length)}`;
    const widths = Array.from({ length: columnCount }, (_, columnIndex) => {
      const maxWidth = Math.max(
        displayWidth(headers[columnIndex]),
        ...rows.map((row) => displayWidth(row[columnIndex]))
      );
      return Math.max(10, Math.min(42, maxWidth + 2));
    });
    const columns = widths.map(
      (width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`
    ).join("");
    const data = allRows.map((row, rowIndex) => {
      const cells = Array.from({ length: columnCount }, (_, columnIndex) =>
        cellXml(row[columnIndex], rowIndex, columnIndex, rowIndex === 0 ? 1 : 0)
      ).join("");
      return `<row r="${rowIndex + 1}"${rowIndex === 0 ? ' ht="22" customHeight="1"' : ""}>${cells}</row>`;
    }).join("");
    const filter = headers.length ? `<autoFilter ref="A1:${columnName(columnCount - 1)}${Math.max(1, allRows.length)}"/>` : "";

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:${lastCell}"/>
  <sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>${columns}</cols>
  <sheetData>${data}</sheetData>
  ${filter}
</worksheet>`;
  }

  function workbookXml(sheets) {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="22000" windowHeight="12000"/></bookViews>
  <sheets>${sheets.map((sheet, index) => `<sheet name="${xmlEscape(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("")}</sheets>
</workbook>`;
  }

  function workbookRelationshipsXml(sheets) {
    const worksheetRelationships = sheets.map(
      (_, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
    ).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${worksheetRelationships}
  <Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }

  function contentTypesXml(sheets) {
    const worksheetOverrides = sheets.map(
      (_, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    ).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${worksheetOverrides}
</Types>`;
  }

  const rootRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Aptos"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos"/><family val="2"/></font>
  </fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1D4ED8"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="2"><border/><border><bottom style="thin"><color rgb="FFD9E2F3"/></bottom></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="top" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment vertical="top"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  let crcTable = null;
  function getCrcTable() {
    if (crcTable) {
      return crcTable;
    }
    crcTable = new Uint32Array(256);
    for (let index = 0; index < 256; index += 1) {
      let value = index;
      for (let bit = 0; bit < 8; bit += 1) {
        value = (value & 1) ? (0xEDB88320 ^ (value >>> 1)) : (value >>> 1);
      }
      crcTable[index] = value >>> 0;
    }
    return crcTable;
  }

  function crc32(bytes) {
    const table = getCrcTable();
    let crc = 0xFFFFFFFF;
    bytes.forEach((byte) => {
      crc = table[(crc ^ byte) & 0xFF] ^ (crc >>> 8);
    });
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  function writeUint16(buffer, offset, value) {
    new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength).setUint16(offset, value, true);
  }

  function writeUint32(buffer, offset, value) {
    new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength).setUint32(offset, value >>> 0, true);
  }

  function concatBytes(chunks) {
    const result = new Uint8Array(chunks.reduce((total, chunk) => total + chunk.length, 0));
    let offset = 0;
    chunks.forEach((chunk) => {
      result.set(chunk, offset);
      offset += chunk.length;
    });
    return result;
  }

  function zipFiles(files) {
    const localChunks = [];
    const centralChunks = [];
    let localOffset = 0;

    files.forEach((file) => {
      const name = encoder.encode(file.name);
      const data = typeof file.data === "string" ? encoder.encode(file.data) : file.data;
      const checksum = crc32(data);
      const localHeader = new Uint8Array(30 + name.length);
      writeUint32(localHeader, 0, 0x04034B50);
      writeUint16(localHeader, 4, 20);
      writeUint16(localHeader, 6, 0x0800);
      writeUint16(localHeader, 8, 0);
      writeUint32(localHeader, 14, checksum);
      writeUint32(localHeader, 18, data.length);
      writeUint32(localHeader, 22, data.length);
      writeUint16(localHeader, 26, name.length);
      localHeader.set(name, 30);
      localChunks.push(localHeader, data);

      const centralHeader = new Uint8Array(46 + name.length);
      writeUint32(centralHeader, 0, 0x02014B50);
      writeUint16(centralHeader, 4, 20);
      writeUint16(centralHeader, 6, 20);
      writeUint16(centralHeader, 8, 0x0800);
      writeUint16(centralHeader, 10, 0);
      writeUint32(centralHeader, 16, checksum);
      writeUint32(centralHeader, 20, data.length);
      writeUint32(centralHeader, 24, data.length);
      writeUint16(centralHeader, 28, name.length);
      writeUint32(centralHeader, 42, localOffset);
      centralHeader.set(name, 46);
      centralChunks.push(centralHeader);

      localOffset += localHeader.length + data.length;
    });

    const central = concatBytes(centralChunks);
    const end = new Uint8Array(22);
    writeUint32(end, 0, 0x06054B50);
    writeUint16(end, 8, files.length);
    writeUint16(end, 10, files.length);
    writeUint32(end, 12, central.length);
    writeUint32(end, 16, localOffset);
    return concatBytes([...localChunks, central, end]);
  }

  function uniqueSheetNames(sheets) {
    const used = new Set();
    return sheets.map((sheet, index) => {
      const base = String(sheet.name || `工作表${index + 1}`)
        .replace(/[\\/*?:[\]]/g, "-")
        .slice(0, 31) || `工作表${index + 1}`;
      let name = base;
      let suffix = 2;
      while (used.has(name)) {
        const tail = `-${suffix}`;
        name = `${base.slice(0, 31 - tail.length)}${tail}`;
        suffix += 1;
      }
      used.add(name);
      return { ...sheet, name };
    });
  }

  function buildWorkbook(inputSheets) {
    const sheets = uniqueSheetNames(Array.isArray(inputSheets) ? inputSheets : []);
    if (!sheets.length) {
      throw new Error("至少需要一個工作表。");
    }

    const files = [
      { name: "[Content_Types].xml", data: contentTypesXml(sheets) },
      { name: "_rels/.rels", data: rootRelationships },
      { name: "xl/workbook.xml", data: workbookXml(sheets) },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRelationshipsXml(sheets) },
      { name: "xl/styles.xml", data: stylesXml },
      ...sheets.map((sheet, index) => ({
        name: `xl/worksheets/sheet${index + 1}.xml`,
        data: worksheetXml(sheet)
      }))
    ];
    return zipFiles(files);
  }

  function downloadWorkbook(filename, sheets) {
    const bytes = buildWorkbook(sheets);
    const blob = new Blob(
      [bytes],
      { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
    );
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  global.QCQuizExcel = { buildWorkbook, downloadWorkbook };
})(globalThis);
