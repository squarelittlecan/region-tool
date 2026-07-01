window.XLSX = (() => {
  const textDecoder = new TextDecoder("utf-8");
  const textEncoder = new TextEncoder();

  const utils = {
    encode_col(index) {
      let name = "";
      let n = index;
      do {
        name = String.fromCharCode(65 + (n % 26)) + name;
        n = Math.floor(n / 26) - 1;
      } while (n >= 0);
      return name;
    },
    decode_col(name) {
      let n = 0;
      for (const char of String(name).toUpperCase()) {
        n = n * 26 + char.charCodeAt(0) - 64;
      }
      return n - 1;
    },
    decode_range(ref) {
      const [start, end = start] = String(ref).split(":");
      return { s: decodeCell(start), e: decodeCell(end) };
    },
    sheet_to_json(sheet, options = {}) {
      const rows = sheet && sheet.__rows ? sheet.__rows : [];
      if (options.header === 1) {
        return rows.map(row => row.map(value => value == null ? (options.defval ?? "") : value));
      }
      return rows;
    },
    aoa_to_sheet(rows) {
      const sheet = { __rows: rows, "!ref": rowsToRef(rows) };
      rows.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
          const address = `${utils.encode_col(colIndex)}${rowIndex + 1}`;
          if (value && typeof value === "object" && value.f) {
            sheet[address] = value;
          } else if (value !== "") {
            sheet[address] = { t: typeof value === "number" ? "n" : "s", v: value };
          }
        });
      });
      return sheet;
    },
    book_new() {
      return { SheetNames: [], Sheets: {} };
    },
    book_append_sheet(workbook, sheet, name) {
      let safe = cleanSheetName(name);
      let suffix = 1;
      while (workbook.Sheets[safe]) {
        const base = safe.slice(0, 28);
        safe = `${base}_${suffix++}`.slice(0, 31);
      }
      workbook.SheetNames.push(safe);
      workbook.Sheets[safe] = sheet;
    }
  };

  async function read(input, options = {}) {
    const bytes = input instanceof ArrayBuffer ? new Uint8Array(input) : new Uint8Array(input.buffer || input);
    const csvProbe = textDecoder.decode(bytes.slice(0, Math.min(bytes.length, 256)));
    if (!csvProbe.includes("PK\u0003\u0004")) return readCsv(textDecoder.decode(bytes));
    const zip = await readZip(bytes);
    const workbookXml = await zip.text("xl/workbook.xml");
    const workbookRels = await zip.text("xl/_rels/workbook.xml.rels");
    const rels = parseWorkbookRels(workbookRels);
    const sharedStrings = zip.has("xl/sharedStrings.xml") ? parseSharedStrings(await zip.text("xl/sharedStrings.xml")) : [];
    const workbookDoc = xml(workbookXml);
    const workbook = { SheetNames: [], Sheets: {} };
    const sheetNodes = [...workbookDoc.getElementsByTagName("sheet")];
    for (const node of sheetNodes) {
      const name = node.getAttribute("name");
      const rid = node.getAttribute("r:id") || node.getAttribute("id");
      const target = normalizePath("xl/" + rels[rid]);
      if (!zip.has(target)) continue;
      const sheet = parseSheet(await zip.text(target), sharedStrings);
      workbook.SheetNames.push(name);
      workbook.Sheets[name] = sheet;
    }
    return workbook;
  }

  async function writeFile(workbook, fileName) {
    const bytes = buildWorkbook(workbook);
    const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function decodeCell(address) {
    const match = String(address).match(/^([A-Z]+)(\d+)$/i);
    return { c: utils.decode_col(match ? match[1] : "A"), r: match ? Number(match[2]) - 1 : 0 };
  }

  function rowsToRef(rows) {
    let maxCol = 0;
    rows.forEach(row => { maxCol = Math.max(maxCol, row.length); });
    return `A1:${utils.encode_col(Math.max(0, maxCol - 1))}${Math.max(1, rows.length)}`;
  }

  function cleanSheetName(name) {
    return String(name || "Sheet").replace(/[\\/?*[\]:]/g, " ").slice(0, 31).trim() || "Sheet";
  }

  function xml(source) {
    return new DOMParser().parseFromString(source, "application/xml");
  }

  function parseWorkbookRels(source) {
    const doc = xml(source);
    const rels = {};
    for (const rel of [...doc.getElementsByTagName("Relationship")]) {
      rels[rel.getAttribute("Id")] = rel.getAttribute("Target");
    }
    return rels;
  }

  function parseSharedStrings(source) {
    const doc = xml(source);
    return [...doc.getElementsByTagName("si")].map(si =>
      [...si.getElementsByTagName("t")].map(t => t.textContent || "").join("")
    );
  }

  function parseSheet(source, sharedStrings) {
    const doc = xml(source);
    const rows = [];
    const cells = {};
    for (const rowNode of [...doc.getElementsByTagName("row")]) {
      const rowIndex = Number(rowNode.getAttribute("r") || rows.length + 1) - 1;
      if (!rows[rowIndex]) rows[rowIndex] = [];
      for (const cellNode of [...rowNode.getElementsByTagName("c")]) {
        const address = cellNode.getAttribute("r") || `${utils.encode_col(rows[rowIndex].length)}${rowIndex + 1}`;
        const colIndex = decodeCell(address).c;
        const type = cellNode.getAttribute("t");
        const formulaNode = cellNode.getElementsByTagName("f")[0];
        const valueNode = cellNode.getElementsByTagName("v")[0];
        const inlineNode = cellNode.getElementsByTagName("is")[0];
        let value = "";
        if (type === "s") {
          value = sharedStrings[Number(valueNode?.textContent || 0)] || "";
        } else if (type === "inlineStr") {
          value = [...inlineNode?.getElementsByTagName("t") || []].map(t => t.textContent || "").join("");
        } else if (valueNode) {
          const raw = valueNode.textContent || "";
          value = raw !== "" && !Number.isNaN(Number(raw)) ? Number(raw) : raw;
        } else if (formulaNode) {
          value = { f: formulaNode.textContent || "" };
        }
        rows[rowIndex][colIndex] = value;
        cells[address] = formulaNode ? { t: "s", f: formulaNode.textContent || "", v: value } : { t: type || "s", v: value };
      }
    }
    const normalized = rows.map(row => {
      const copy = [];
      for (let i = 0; i < row.length; i += 1) copy[i] = row[i] == null ? "" : row[i];
      return copy;
    });
    return Object.assign(cells, { __rows: normalized, "!ref": rowsToRef(normalized) });
  }

  function readCsv(text) {
    const rows = [];
    let row = [];
    let value = "";
    let quoted = false;
    for (let i = 0; i < text.length; i += 1) {
      const char = text[i];
      if (quoted && char === '"' && text[i + 1] === '"') {
        value += '"';
        i += 1;
      } else if (char === '"') {
        quoted = !quoted;
      } else if (!quoted && char === ",") {
        row.push(value);
        value = "";
      } else if (!quoted && (char === "\n" || char === "\r")) {
        if (char === "\r" && text[i + 1] === "\n") i += 1;
        row.push(value);
        rows.push(row);
        row = [];
        value = "";
      } else {
        value += char;
      }
    }
    if (value || row.length) {
      row.push(value);
      rows.push(row);
    }
    const sheet = utils.aoa_to_sheet(rows);
    return { SheetNames: ["Sheet1"], Sheets: { Sheet1: sheet } };
  }

  async function readZip(bytes) {
    const eocd = findEocd(bytes);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const entriesCount = view.getUint16(eocd + 10, true);
    const centralOffset = view.getUint32(eocd + 16, true);
    const files = new Map();
    let offset = centralOffset;
    for (let i = 0; i < entriesCount; i += 1) {
      if (view.getUint32(offset, true) !== 0x02014b50) throw new Error("XLSX 中央目录读取失败");
      const method = view.getUint16(offset + 10, true);
      const compressedSize = view.getUint32(offset + 20, true);
      const uncompressedSize = view.getUint32(offset + 24, true);
      const nameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      const localOffset = view.getUint32(offset + 42, true);
      const name = textDecoder.decode(bytes.slice(offset + 46, offset + 46 + nameLength));
      files.set(normalizePath(name), { method, compressedSize, uncompressedSize, localOffset });
      offset += 46 + nameLength + extraLength + commentLength;
    }
    return {
      has(name) {
        return files.has(normalizePath(name));
      },
      async text(name) {
        const data = await extractZipEntry(bytes, files.get(normalizePath(name)));
        return textDecoder.decode(data);
      }
    };
  }

  function findEocd(bytes) {
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 66000); i -= 1) {
      if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) return i;
    }
    throw new Error("不是有效的 xlsx 文件");
  }

  async function extractZipEntry(bytes, entry) {
    if (!entry) throw new Error("XLSX 缺少必要文件");
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const local = entry.localOffset;
    if (view.getUint32(local, true) !== 0x04034b50) throw new Error("XLSX 本地文件头读取失败");
    const nameLength = view.getUint16(local + 26, true);
    const extraLength = view.getUint16(local + 28, true);
    const dataStart = local + 30 + nameLength + extraLength;
    const data = bytes.slice(dataStart, dataStart + entry.compressedSize);
    if (entry.method === 0) return data;
    if (entry.method !== 8) throw new Error("暂不支持此 XLSX 压缩格式");
    if (!("DecompressionStream" in window)) throw new Error("浏览器不支持解压 XLSX，请使用新版 Chrome 或 Edge");
    const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  function normalizePath(path) {
    const parts = [];
    String(path).replace(/\\/g, "/").split("/").forEach(part => {
      if (!part || part === ".") return;
      if (part === "..") parts.pop();
      else parts.push(part);
    });
    return parts.join("/");
  }

  function buildWorkbook(workbook) {
    const files = new Map();
    const sheetNames = workbook.SheetNames.map(cleanSheetName);
    files.set("[Content_Types].xml", contentTypes(sheetNames.length));
    files.set("_rels/.rels", rootRels());
    files.set("xl/workbook.xml", workbookXml(sheetNames));
    files.set("xl/_rels/workbook.xml.rels", workbookRels(sheetNames.length));
    files.set("xl/styles.xml", stylesXml());
    sheetNames.forEach((name, index) => {
      files.set(`xl/worksheets/sheet${index + 1}.xml`, sheetXml(workbook.Sheets[workbook.SheetNames[index]]));
    });
    return writeZip(files);
  }

  function contentTypes(count) {
    let overrides = "";
    for (let i = 1; i <= count; i += 1) {
      overrides += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${overrides}</Types>`;
  }

  function rootRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
  }

  function workbookXml(names) {
    const sheets = names.map((name, index) => `<sheet name="${escapeXml(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets}</sheets><calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>`;
  }

  function workbookRels(count) {
    let rels = "";
    for (let i = 1; i <= count; i += 1) {
      rels += `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`;
    }
    rels += `<Relationship Id="rId${count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
  }

  function stylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`;
  }

  function sheetXml(sheet) {
    const rows = sheet.__rows || [];
    const lines = rows.map((row, rowIndex) => {
      const cells = row.map((value, colIndex) => {
        const address = `${utils.encode_col(colIndex)}${rowIndex + 1}`;
        const cell = sheet[address];
        const actual = cell && cell.f ? cell : value;
        return cellXml(address, actual);
      }).join("");
      return `<row r="${rowIndex + 1}">${cells}</row>`;
    }).join("");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="${rowsToRef(rows)}"/><sheetData>${lines}</sheetData></worksheet>`;
  }

  function cellXml(address, value) {
    if (value && typeof value === "object" && value.f) {
      return `<c r="${address}"><f>${escapeXml(value.f)}</f></c>`;
    }
    if (typeof value === "number") return `<c r="${address}"><v>${value}</v></c>`;
    return `<c r="${address}" t="inlineStr"><is><t>${escapeXml(value == null ? "" : String(value))}</t></is></c>`;
  }

  function escapeXml(value) {
    return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function writeZip(files) {
    const chunks = [];
    const central = [];
    let offset = 0;
    for (const [name, text] of files) {
      const nameBytes = textEncoder.encode(name);
      const data = textEncoder.encode(text);
      const crc = crc32(data);
      const local = concat([
        u32(0x04034b50), u16(20), u16(0), u16(0), u16(0), u16(0),
        u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), nameBytes, data
      ]);
      chunks.push(local);
      central.push(concat([
        u32(0x02014b50), u16(20), u16(20), u16(0), u16(0), u16(0), u16(0),
        u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), u16(0),
        u16(0), u16(0), u32(0), u32(offset), nameBytes
      ]));
      offset += local.length;
    }
    const centralBytes = concat(central);
    const eocd = concat([
      u32(0x06054b50), u16(0), u16(0), u16(files.size), u16(files.size),
      u32(centralBytes.length), u32(offset), u16(0)
    ]);
    return concat([...chunks, centralBytes, eocd]);
  }

  function concat(parts) {
    const total = parts.reduce((sum, part) => sum + part.length, 0);
    const out = new Uint8Array(total);
    let offset = 0;
    parts.forEach(part => {
      out.set(part, offset);
      offset += part.length;
    });
    return out;
  }

  function u16(value) {
    const out = new Uint8Array(2);
    new DataView(out.buffer).setUint16(0, value, true);
    return out;
  }

  function u32(value) {
    const out = new Uint8Array(4);
    new DataView(out.buffer).setUint32(0, value >>> 0, true);
    return out;
  }

  const crcTable = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n += 1) {
      let c = n;
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (const byte of bytes) crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    return (crc ^ 0xffffffff) >>> 0;
  }

  return { read, writeFile, utils };
})();
