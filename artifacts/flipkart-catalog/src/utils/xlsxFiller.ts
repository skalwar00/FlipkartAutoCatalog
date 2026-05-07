import JSZip from "jszip";

export type Mapping = { newCol: string; oldCol: string | null };

function escapeXml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function colIndexToLetter(idx: number): string {
  let letter = "";
  let n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

export function isXlsxBytes(bytes: ArrayBuffer): boolean {
  const view = new Uint8Array(bytes, 0, 4);
  return view[0] === 0x50 && view[1] === 0x4b && view[2] === 0x03 && view[3] === 0x04;
}

function parseSharedStrings(xml: string): string[] {
  const result: string[] = [];
  const siRegex = /<si>([\s\S]*?)<\/si>/g;
  let siMatch;
  while ((siMatch = siRegex.exec(xml)) !== null) {
    const siContent = siMatch[1];
    const tRegex = /<t[^>]*>([\s\S]*?)<\/t>/g;
    let tMatch;
    let combined = "";
    while ((tMatch = tRegex.exec(siContent)) !== null) {
      combined += tMatch[1];
    }
    result.push(combined);
  }
  return result;
}

function resolveCell(
  cellXml: string,
  sharedStrings: string[]
): string | number | null {
  const typeMatch = cellXml.match(/\bt="([^"]+)"/);
  const cellType = typeMatch ? typeMatch[1] : null;

  if (cellType === "inlineStr") {
    const tMatch = cellXml.match(/<t[^>]*>([\s\S]*?)<\/t>/);
    return tMatch ? tMatch[1] : null;
  }

  const vMatch = cellXml.match(/<v>([\s\S]*?)<\/v>/);
  if (!vMatch) return null;
  const vVal = vMatch[1];

  if (cellType === "s") {
    const idx = parseInt(vVal, 10);
    return sharedStrings[idx] ?? null;
  }
  if (cellType === "b") return vVal === "1" ? "TRUE" : "FALSE";
  const num = parseFloat(vVal);
  return isNaN(num) ? vVal : num;
}

function findHeaderRowInXml(
  sheetXml: string,
  sharedStrings: string[]
): { headerRowNum: number; colIndexMap: Record<string, number> } {
  const rowRegex = /<row[^>]+r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let bestRowNum = 1;
  let bestScore = -1;
  let bestCols: Record<string, number> = {};

  let rowMatch;
  const checkedRows: Array<{ rowNum: number; rowContent: string }> = [];

  while ((rowMatch = rowRegex.exec(sheetXml)) !== null) {
    const rowNum = parseInt(rowMatch[1]);
    if (rowNum > 40) break;
    checkedRows.push({ rowNum, rowContent: rowMatch[2] });
  }

  for (const { rowNum, rowContent } of checkedRows) {
    const cellRegex = /<c\s+r="([A-Z]+)\d+"([\s\S]*?)<\/c>/g;
    let cellMatch;
    let score = 0;
    const colMap: Record<string, number> = {};

    while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
      const colLetters = cellMatch[1];
      const cellXml = `<c r="${cellMatch[1]}${rowNum}"${cellMatch[2]}</c>`;
      const val = resolveCell(cellXml, sharedStrings);
      if (val !== null && typeof val === "string" && val.trim().length > 0) {
        score++;
        const colIdx = colLetterToIndex(colLetters);
        colMap[val.trim()] = colIdx;
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestRowNum = rowNum;
      bestCols = colMap;
    }
  }

  return { headerRowNum: bestRowNum, colIndexMap: bestCols };
}

/**
 * Given the sheetXml and headerRowNum, scan the rows immediately after the
 * header and return the row number where actual user data should start.
 *
 * "System rows" (Flipkart M/O/C rows, instruction rows, etc.) that exist in
 * the template between the header row and the data area are preserved as-is.
 * They are identified by having ALL non-empty cell values be very short
 * (≤ 3 characters, e.g. "M", "O", "C", "Yes", "No") or known marker strings.
 */
function findDataStartRow(
  sheetXml: string,
  sharedStrings: string[],
  headerRowNum: number
): number {
  const SYSTEM_MARKERS = new Set([
    "m", "o", "c", "r", "yes", "no", "y", "n",
    "mandatory", "optional", "conditional", "required",
  ]);

  const rowRegex = /<row[^>]+r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let rowMatch;

  // Collect rows in the window just after the header (up to 6 rows)
  const candidateRows: Array<{ rowNum: number; xml: string }> = [];
  while ((rowMatch = rowRegex.exec(sheetXml)) !== null) {
    const rowNum = parseInt(rowMatch[1]);
    if (rowNum <= headerRowNum) continue;
    if (rowNum > headerRowNum + 6) break;
    candidateRows.push({ rowNum, xml: rowMatch[0] });
  }

  let dataStartRow = headerRowNum + 1;

  for (const { rowNum, xml } of candidateRows) {
    // Extract all cell values in this row
    const cellRegex = /<c\s+r="([A-Z]+)\d+"([\s\S]*?)<\/c>/g;
    let cellMatch;
    const vals: string[] = [];

    while ((cellMatch = cellRegex.exec(xml)) !== null) {
      const cellXml = `<c r="${cellMatch[1]}${rowNum}"${cellMatch[2]}</c>`;
      const val = resolveCell(cellXml, sharedStrings);
      if (val !== null && String(val).trim() !== "") {
        vals.push(String(val).trim());
      }
    }

    if (vals.length === 0) {
      // Empty row — stop preserving, data can start here
      break;
    }

    const isSystemRow = vals.every(
      (v) => v.length <= 3 || SYSTEM_MARKERS.has(v.toLowerCase())
    );

    if (isSystemRow) {
      dataStartRow = rowNum + 1;
    } else {
      // Looks like real data — stop preserving
      break;
    }
  }

  return dataStartRow;
}

function colLetterToIndex(letters: string): number {
  let result = 0;
  for (let i = 0; i < letters.length; i++) {
    result = result * 26 + (letters.charCodeAt(i) - 64);
  }
  return result - 1;
}

export interface FillResult {
  blob: Blob;
  headerRowNum: number;
  rowsWritten: number;
}

/**
 * @param knownHeaderRowNum  Optional 1-indexed header row number already
 *   detected by the XLSX-library-based parser. When provided it is used
 *   directly, avoiding a second (potentially diverging) XML-based detection.
 */
export async function fillXlsxTemplate(
  templateBytes: ArrayBuffer,
  targetSheetName: string | null,
  mappings: Mapping[],
  dataRows: Record<string, string | number | boolean | null>[],
  knownHeaderRowNum?: number
): Promise<FillResult> {
  const zip = await JSZip.loadAsync(templateBytes);

  const sharedStringsRaw = await zip.file("xl/sharedStrings.xml")?.async("string");
  const sharedStrings = sharedStringsRaw ? parseSharedStrings(sharedStringsRaw) : [];

  const relsXml = (await zip.file("xl/_rels/workbook.xml.rels")?.async("string")) ?? "";
  const workbookXml = (await zip.file("xl/workbook.xml")?.async("string")) ?? "";

  let targetRid: string | null = null;
  if (targetSheetName) {
    const sheetMatch = workbookXml.match(
      new RegExp(`<sheet[^>]+name="${escapeXml(targetSheetName)}"[^>]+r:id="([^"]+)"`)
    );
    targetRid = sheetMatch ? sheetMatch[1] : null;
  }
  if (!targetRid) {
    const firstSheet = workbookXml.match(/<sheet[^>]+r:id="([^"]+)"/);
    targetRid = firstSheet ? firstSheet[1] : "rId1";
  }

  const relMatch = relsXml.match(new RegExp(`Id="${targetRid}"[^>]*Target="([^"]+)"`));
  const sheetRelPath = relMatch ? relMatch[1] : "worksheets/sheet1.xml";
  const sheetPath = `xl/${sheetRelPath}`;

  const sheetXml = await zip.file(sheetPath)?.async("string");
  if (!sheetXml) throw new Error("Sheet XML not found: " + sheetPath);

  // Use the caller-supplied header row (preferred) or re-detect from XML
  let headerRowNum: number;
  let colIndexMap: Record<string, number>;

  if (knownHeaderRowNum !== undefined) {
    // knownHeaderRowNum is 1-indexed (same as Excel row numbers)
    headerRowNum = knownHeaderRowNum;
    // Still need to build the colIndexMap from the XML for this specific row
    const detected = findHeaderRowInXml(sheetXml, sharedStrings);
    // If the detected row matches (±1 tolerance), use its colIndexMap;
    // otherwise re-scan the known row for column positions
    if (Math.abs(detected.headerRowNum - headerRowNum) <= 1) {
      headerRowNum = detected.headerRowNum; // prefer exact match
      colIndexMap = detected.colIndexMap;
    } else {
      colIndexMap = detected.colIndexMap;
      headerRowNum = knownHeaderRowNum;
    }
  } else {
    const detected = findHeaderRowInXml(sheetXml, sharedStrings);
    headerRowNum = detected.headerRowNum;
    colIndexMap = detected.colIndexMap;
  }

  // Find where actual user data should start (skip system/marker rows)
  const dataStartRow = findDataStartRow(sheetXml, sharedStrings, headerRowNum);

  // Build new data rows XML starting at dataStartRow
  const newRowsXml = dataRows
    .map((row, idx) => {
      const targetRowNum = dataStartRow + idx;
      const cells: string[] = [];

      for (const m of mappings) {
        const colIdx = colIndexMap[m.newCol];
        if (colIdx === undefined) continue;
        const value = m.oldCol ? (row[m.oldCol] ?? null) : null;
        if (value === null || value === "") continue;

        const colLetter = colIndexToLetter(colIdx);
        const cellRef = `${colLetter}${targetRowNum}`;

        if (typeof value === "number") {
          cells.push(`<c r="${cellRef}"><v>${value}</v></c>`);
        } else {
          cells.push(
            `<c r="${cellRef}" t="inlineStr"><is><t>${escapeXml(String(value))}</t></is></c>`
          );
        }
      }

      if (cells.length === 0) return "";
      cells.sort((a, b) => {
        const refA = a.match(/r="([A-Z]+)\d+"/)?.[1] ?? "";
        const refB = b.match(/r="([A-Z]+)\d+"/)?.[1] ?? "";
        return colLetterToIndex(refA) - colLetterToIndex(refB);
      });
      return `<row r="${targetRowNum}">${cells.join("")}</row>`;
    })
    .filter(Boolean)
    .join("\n");

  // Keep all original template rows BEFORE dataStartRow (header + system rows)
  // and replace everything from dataStartRow onwards with new data
  const modifiedXml = sheetXml.replace(
    /(<sheetData>)([\s\S]*?)(<\/sheetData>)/,
    (_, open, content, close) => {
      const rowRegex = /<row[^>]+r="(\d+)"[\s\S]*?<\/row>/g;
      let keptRows = "";
      let match;
      while ((match = rowRegex.exec(content)) !== null) {
        const rowNum = parseInt(match[1]);
        if (rowNum < dataStartRow) {
          keptRows += match[0];
        }
      }
      return `${open}${keptRows}${newRowsXml}${close}`;
    }
  );

  // Update dimension ref
  const lastDataRow = dataStartRow - 1 + dataRows.length;
  const updatedXml = modifiedXml.replace(
    /(<dimension ref="[A-Z]+\d+:)([A-Z]+)(\d+)(")/,
    (_, prefix, endCol, _endRow, suffix) => `${prefix}${endCol}${lastDataRow}${suffix}`
  );

  zip.file(sheetPath, updatedXml);

  const blob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return { blob, headerRowNum, rowsWritten: dataRows.length };
}
