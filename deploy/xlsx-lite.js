// 최소 XLSX 파서: zip 해제(DecompressionStream) + 시트/스타일 XML 파싱.
// "정책표 입력" 엑셀 첨부 기능 전용 - 고정 양식(1~3행 헤더, A열 기종명, 이후 값 열) 가정.

async function unzipXlsx(arrayBuffer) {
  const buf = new Uint8Array(arrayBuffer);
  const dv = new DataView(buf.buffer);
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("올바른 xlsx 파일이 아닙니다.");
  const cdOffset = dv.getUint32(eocd + 16, true);
  const cdEntries = dv.getUint16(eocd + 10, true);
  let pos = cdOffset;
  const entries = [];
  for (let i = 0; i < cdEntries; i++) {
    const compMethod = dv.getUint16(pos + 10, true);
    const compSize = dv.getUint32(pos + 20, true);
    const nameLen = dv.getUint16(pos + 28, true);
    const extraLen = dv.getUint16(pos + 30, true);
    const commentLen = dv.getUint16(pos + 32, true);
    const localOffset = dv.getUint32(pos + 42, true);
    const name = new TextDecoder().decode(buf.subarray(pos + 46, pos + 46 + nameLen));
    entries.push({ name, compMethod, compSize, localOffset });
    pos += 46 + nameLen + extraLen + commentLen;
  }
  async function extract(entry) {
    const lpos = entry.localOffset;
    const lNameLen = dv.getUint16(lpos + 26, true);
    const lExtraLen = dv.getUint16(lpos + 28, true);
    const dataStart = lpos + 30 + lNameLen + lExtraLen;
    const compData = buf.subarray(dataStart, dataStart + entry.compSize);
    if (entry.compMethod === 0) return compData.slice();
    const ds = new DecompressionStream("deflate-raw");
    const stream = new Blob([compData]).stream().pipeThrough(ds);
    const ab = await new Response(stream).arrayBuffer();
    return new Uint8Array(ab);
  }
  const files = {};
  for (const e of entries) files[e.name] = await extract(e);
  return files;
}

const bytesToText = bytes => new TextDecoder("utf-8").decode(bytes);

function colLetterToIndex(letter) {
  let n = 0;
  for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
  return n;
}

// 빌트인 날짜/시간 서식 번호 (ECMA-376).
const BUILTIN_DATE_FMT_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

function parseStyles(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const fonts = [];
  const fontsNode = doc.getElementsByTagName("fonts")[0];
  if (fontsNode) {
    Array.from(fontsNode.children).forEach(f => {
      const colorEl = f.getElementsByTagName("color")[0];
      let black = true;
      if (colorEl) {
        const rgb = colorEl.getAttribute("rgb");
        const theme = colorEl.getAttribute("theme");
        const indexed = colorEl.getAttribute("indexed");
        if (rgb) black = rgb.slice(-6).toUpperCase() === "000000";
        else if (theme != null) black = theme === "1";
        else if (indexed != null) black = indexed === "0" || indexed === "8";
      }
      fonts.push({ black });
    });
  }
  const customFmts = {};
  const numFmtsNode = doc.getElementsByTagName("numFmts")[0];
  if (numFmtsNode) {
    Array.from(numFmtsNode.children).forEach(nf => {
      customFmts[nf.getAttribute("numFmtId")] = nf.getAttribute("formatCode") || "";
    });
  }
  const xfs = [];
  const cellXfsNode = doc.getElementsByTagName("cellXfs")[0];
  if (cellXfsNode) {
    Array.from(cellXfsNode.children).forEach(xf => {
      const numFmtId = parseInt(xf.getAttribute("numFmtId") || "0", 10);
      let isDate = BUILTIN_DATE_FMT_IDS.has(numFmtId);
      if (!isDate && customFmts[numFmtId] != null) {
        const code = customFmts[numFmtId];
        // 날짜/시간 서식 코드로 판단: y/m/d/h/s 토큰 포함, 색상 코드/일반 서식 등 제외.
        isDate = /[ymdhs]/i.test(code) && !/^general$/i.test(code) && !code.includes("%") && !code.includes("@");
      }
      xfs.push({ fontId: parseInt(xf.getAttribute("fontId") || "0", 10), isDate });
    });
  }
  return { fonts, xfs };
}

function isStyleBlack(styles, styleIdx) {
  const xf = styles.xfs[styleIdx || 0];
  if (!xf) return true;
  const font = styles.fonts[xf.fontId];
  return font ? font.black : true;
}

function isStyleDate(styles, styleIdx) {
  const xf = styles.xfs[styleIdx || 0];
  return xf ? !!xf.isDate : false;
}

// Excel 날짜 일련번호(1900 날짜 체계, Unix epoch 기준) → "YYYY-MM-DD" 문자열.
function excelSerialToDateStr(serial) {
  const num = parseFloat(serial);
  if (isNaN(num)) return String(serial);
  const ms = Math.round((num - 25569) * 86400 * 1000);
  const d = new Date(ms);
  const y = d.getUTCFullYear(), m = d.getUTCMonth() + 1, day = d.getUTCDate();
  return y + "-" + String(m).padStart(2, "0") + "-" + String(day).padStart(2, "0");
}

function parseSharedStrings(xml) {
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return Array.from(doc.getElementsByTagName("si")).map(si => Array.from(si.getElementsByTagName("t")).map(t => t.textContent).join(""));
}

function parseSheetRows(xml, sharedStrings, styles) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  return Array.from(doc.getElementsByTagName("row")).map(rowEl => {
    const cells = {};
    Array.from(rowEl.getElementsByTagName("c")).forEach(c => {
      const ref = c.getAttribute("r") || "";
      const m = ref.match(/[A-Z]+/);
      if (!m) return;
      const col = colLetterToIndex(m[0]);
      const styleIdx = parseInt(c.getAttribute("s") || "0", 10);
      const type = c.getAttribute("t");
      let value = "";
      const isEl = c.getElementsByTagName("is")[0];
      const vEl = c.getElementsByTagName("v")[0];
      if (isEl) value = Array.from(isEl.getElementsByTagName("t")).map(t => t.textContent).join("");
      else if (vEl) {
        value = vEl.textContent;
        if (type === "s") value = sharedStrings[parseInt(value, 10)] || "";
        else if (!type && isStyleDate(styles, styleIdx) && value !== "") value = excelSerialToDateStr(value);
      }
      cells[col] = { value, black: isStyleBlack(styles, styleIdx) };
    });
    return cells;
  });
}

// 반환: { columns:[{col,plan,join,disc}], deviceRows:[{device, values:[{plan,join,disc,value,black}]}] }
// 고정 양식: 1행=요금제, 2행=가입유형, 3행=할인유형, A열=라벨/기종명, 4행부터 데이터.
async function readXlsxCostGrid(arrayBuffer) {
  const files = await unzipXlsx(arrayBuffer);
  const sheetName = Object.keys(files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort()[0];
  if (!sheetName) throw new Error("워크시트를 찾을 수 없습니다.");
  const styles = parseStyles(files["xl/styles.xml"] ? bytesToText(files["xl/styles.xml"]) : "");
  const sharedStrings = parseSharedStrings(files["xl/sharedStrings.xml"] ? bytesToText(files["xl/sharedStrings.xml"]) : "");
  const rows = parseSheetRows(bytesToText(files[sheetName]), sharedStrings, styles);
  if (rows.length < 4) throw new Error("시트 형식을 인식할 수 없습니다 (헤더 3행 + 기종 행이 필요합니다).");
  const [planRow, joinRow, discRow] = rows;
  const maxCol = Math.max(0, ...Object.keys(planRow).map(Number));
  const columns = [];
  for (let col = 2; col <= maxCol; col++) {
    if (!planRow[col] || !planRow[col].value) continue;
    columns.push({
      col,
      plan: (planRow[col].value || "").trim(),
      join: (joinRow[col] ? joinRow[col].value : "").trim(),
      disc: (discRow[col] ? discRow[col].value : "").trim()
    });
  }
  if (!columns.length) throw new Error("요금제 헤더(1행)를 읽지 못했습니다.");
  const deviceRows = [];
  for (let i = 3; i < rows.length; i++) {
    const r = rows[i];
    const name = r[1] ? (r[1].value || "").trim() : "";
    if (!name) continue;
    const values = columns.map(cm => {
      const cell = r[cm.col];
      const num = cell ? parseFloat(cell.value) : NaN;
      return { plan: cm.plan, join: cm.join, disc: cm.disc, value: isNaN(num) ? null : num, black: cell ? cell.black : false };
    });
    deviceRows.push({ device: name, values });
  }
  return { columns, deviceRows };
}
window.readXlsxCostGrid = readXlsxCostGrid;

// 범용 시트 리더: 행 배열({colIndex:{value,black}}) 그대로 반환. 재고 엑셀처럼
// 고정 헤더 없이 임의 열(A,B,F,G,J,M,N,O,T 등)만 뽑아 쓰는 용도.
async function readXlsxSheetRows(arrayBuffer) {
  const files = await unzipXlsx(arrayBuffer);
  const sheetName = Object.keys(files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort()[0];
  if (!sheetName) throw new Error("워크시트를 찾을 수 없습니다.");
  const styles = parseStyles(files["xl/styles.xml"] ? bytesToText(files["xl/styles.xml"]) : "");
  const sharedStrings = parseSharedStrings(files["xl/sharedStrings.xml"] ? bytesToText(files["xl/sharedStrings.xml"]) : "");
  return parseSheetRows(bytesToText(files[sheetName]), sharedStrings, styles);
}
window.readXlsxSheetRows = readXlsxSheetRows;
