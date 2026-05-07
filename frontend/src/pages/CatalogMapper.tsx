import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { fillXlsxTemplate, isXlsxBytes, type Mapping } from "@/utils/xlsxFiller";

type SheetData = {
  headers: string[];
  rows: Record<string, string | number | boolean | null>[];
  sheetNames: string[];
  selectedSheet: string;
  workbook: XLSX.WorkBook;
  headerRowIndex: number;
  rawBytes: ArrayBuffer;
  isXlsx: boolean;
  colIndexMap: Record<string, number>;
};

type ColumnMapping = {
  newCol: string;
  oldCol: string | null;
};

type EditableRow = Record<string, string>;

type ColTransform = { prefix: string; suffix: string };

type ImgConfig = {
  imgKeyCol: string;
  oldKeyCol: string;
  imgUrlCol: string;
  newTargetCol: string;
};

function findHeaderRow(ws: XLSX.WorkSheet): {
  headerRowIndex: number;
  headers: string[];
  colIndexMap: Record<string, number>;
} {
  if (!ws["!ref"]) return { headerRowIndex: 0, headers: [], colIndexMap: {} };
  const range = XLSX.utils.decode_range(ws["!ref"]);
  let bestRow = range.s.r;
  let bestScore = -1;
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 30); r++) {
    let score = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "string" && cell.v.trim().length > 0) score++;
    }
    if (score > bestScore) { bestScore = score; bestRow = r; }
  }
  const headers: string[] = [];
  const colIndexMap: Record<string, number> = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: bestRow, c })];
    if (cell && cell.v !== null && cell.v !== undefined) {
      const key = String(cell.v).trim();
      if (key) { headers.push(key); colIndexMap[key] = c; }
    }
  }
  return { headerRowIndex: bestRow, headers, colIndexMap };
}

function parseWorkbook(wb: XLSX.WorkBook, sheetName: string, rawBytes: ArrayBuffer, isXlsx: boolean): SheetData {
  const ws = wb.Sheets[sheetName];
  const fallback: SheetData = {
    headers: [], rows: [], sheetNames: wb.SheetNames,
    selectedSheet: sheetName, workbook: wb,
    headerRowIndex: 0, rawBytes, isXlsx, colIndexMap: {}
  };
  if (!ws) return fallback;
  const { headerRowIndex, headers, colIndexMap } = findHeaderRow(ws);
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
  const rows: Record<string, string | number | boolean | null>[] = [];
  for (let r = headerRowIndex + 1; r <= range.e.r; r++) {
    const row: Record<string, string | number | boolean | null> = {};
    let hasData = false;
    for (const [header, c] of Object.entries(colIndexMap)) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      const val = cell ? (cell.v ?? null) : null;
      row[header] = val as string | number | boolean | null;
      if (val !== null && val !== "") hasData = true;
    }
    if (hasData) rows.push(row);
  }
  return { headers, rows, sheetNames: wb.SheetNames, selectedSheet: sheetName, workbook: wb, headerRowIndex, rawBytes, isXlsx, colIndexMap };
}

function normalizeHeader(h: string): string {
  return h.toLowerCase().replace(/[\s_\-().*/\\]+/g, "").trim();
}

function autoMatch(newHeaders: string[], oldHeaders: string[]): Record<string, string | null> {
  const result: Record<string, string | null> = {};
  for (const nh of newHeaders) {
    const normNh = normalizeHeader(nh);
    let best: string | null = null;
    for (const oh of oldHeaders) {
      if (normalizeHeader(oh) === normNh) { best = oh; break; }
    }
    if (!best) {
      for (const oh of oldHeaders) {
        const normOh = normalizeHeader(oh);
        if (normOh.includes(normNh) || normNh.includes(normOh)) { best = oh; break; }
      }
    }
    result[nh] = best;
  }
  return result;
}

function buildEditableRows(
  mappings: ColumnMapping[],
  oldRows: Record<string, string | number | boolean | null>[],
  imgLookup?: Map<string, string>,
  imgOldKeyCol?: string,
  imgNewTargetCol?: string
): EditableRow[] {
  return oldRows.map((oldRow) => {
    const row: EditableRow = {};
    for (const m of mappings) {
      const val = m.oldCol ? (oldRow[m.oldCol] ?? "") : "";
      row[m.newCol] = val === null ? "" : String(val);
    }
    if (imgLookup && imgOldKeyCol && imgNewTargetCol) {
      const key = String(oldRow[imgOldKeyCol] ?? "").trim();
      if (key) {
        const url = imgLookup.get(key);
        if (url) row[imgNewTargetCol] = url;
      }
    }
    return row;
  });
}

function findBestHeader(target: string, headers: string[]): string {
  const normTarget = target.toLowerCase().replace(/\s+/g, " ").trim();
  const exact = headers.find((h) => h === target);
  if (exact) return exact;
  const caseInsensitive = headers.find((h) => h.toLowerCase() === normTarget);
  if (caseInsensitive) return caseInsensitive;
  const partial = headers.find((h) => h.toLowerCase().includes(normTarget) || normTarget.includes(h.toLowerCase()));
  if (partial) return partial;
  return "";
}

function applyTransform(value: string, t: ColTransform | undefined): string {
  if (!t) return value;
  if (value === "") return value;
  return `${t.prefix}${value}${t.suffix}`;
}

export default function CatalogMapper() {
  const [oldData, setOldData] = useState<SheetData | null>(null);
  const [newData, setNewData] = useState<SheetData | null>(null);
  const [imgData, setImgData] = useState<SheetData | null>(null);
  const [mappings, setMappings] = useState<ColumnMapping[]>([]);
  const [step, setStep] = useState<1 | 2 | 3 | 4>(1);
  const [oldFileName, setOldFileName] = useState("");
  const [newFileName, setNewFileName] = useState("");
  const [imgFileName, setImgFileName] = useState("");
  const [dragOver, setDragOver] = useState<"old" | "new" | "img" | null>(null);
  const [generating, setGenerating] = useState(false);
  const [editableRows, setEditableRows] = useState<EditableRow[]>([]);
  const [editingCell, setEditingCell] = useState<{ rowIdx: number; col: string } | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [columnTransforms, setColumnTransforms] = useState<Record<string, ColTransform>>({});
  const [transformPanelOpen, setTransformPanelOpen] = useState(false);
  const [findReplaceOpen, setFindReplaceOpen] = useState(false);
  const [colSelectorOpen, setColSelectorOpen] = useState(false);
  const [findVal, setFindVal] = useState("");
  const [replaceVal, setReplaceVal] = useState("");
  const [findReplaceCols, setFindReplaceCols] = useState<string[]>([]);
  const [findReplaceCount, setFindReplaceCount] = useState<number | null>(null);
  const [imgConfig, setImgConfig] = useState<ImgConfig>({ imgKeyCol: "", oldKeyCol: "", imgUrlCol: "", newTargetCol: "" });
  const [imgConfigOpen, setImgConfigOpen] = useState(false);
  const [imgMatchCount, setImgMatchCount] = useState<number | null>(null);
  const oldInputRef = useRef<HTMLInputElement>(null);
  const newInputRef = useRef<HTMLInputElement>(null);
  const imgInputRef = useRef<HTMLInputElement>(null);

  const detectImgSheet = useCallback((wb: XLSX.WorkBook): string => {
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      if (!ws) continue;
      const { headers } = findHeaderRow(ws);
      if (headers.some((h) => h.toLowerCase().includes("url"))) return name;
    }
    return wb.SheetNames[0];
  }, []);

  const loadFile = useCallback((file: File, type: "old" | "new" | "img") => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const rawBytes = e.target?.result as ArrayBuffer;
      const isXlsx = isXlsxBytes(rawBytes);
      const data = new Uint8Array(rawBytes);
      const wb = XLSX.read(data, { type: "array" });
      const sheetName = type === "img" ? detectImgSheet(wb) : wb.SheetNames[0];
      const sheetData = parseWorkbook(wb, sheetName, rawBytes, isXlsx);
      if (type === "old") { setOldData(sheetData); setOldFileName(file.name); }
      else if (type === "new") { setNewData(sheetData); setNewFileName(file.name); }
      else { setImgData(sheetData); setImgFileName(file.name); setImgConfig({ imgKeyCol: "", oldKeyCol: "", imgUrlCol: "", newTargetCol: "" }); setImgMatchCount(null); }
    };
    reader.readAsArrayBuffer(file);
  }, [detectImgSheet]);

  const handleFile = useCallback((file: File | undefined, type: "old" | "new" | "img") => {
    if (!file) return;
    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
      alert("Sirf .xlsx, .xls ya .csv files allowed hain");
      return;
    }
    loadFile(file, type);
  }, [loadFile]);

  useEffect(() => {
    if (!imgData) return;
    setImgConfig((prev) => ({
      imgKeyCol: prev.imgKeyCol || findBestHeader("Seller SKU ID", imgData.headers),
      imgUrlCol: prev.imgUrlCol || findBestHeader("Main Image URL", imgData.headers),
      oldKeyCol: prev.oldKeyCol || findBestHeader("Brand Color", oldData?.headers ?? []),
      newTargetCol: prev.newTargetCol || findBestHeader("Main Image URL", newData?.headers ?? []),
    }));
  }, [imgData, oldData, newData]);

  useEffect(() => {
    if (editableRows.length === 0) return;
    const cols = mappings.filter((m) => m.oldCol !== null).map((m) => m.newCol);
    const brandCols = cols.filter(isBrandCol);
    setFindReplaceCols(brandCols);
    setFindReplaceCount(null);
    // Auto-fill findVal: prefer column with "brand" in name, then any brand col
    const brandNameCol =
      brandCols.find((c) => c.toLowerCase().includes("brand")) ?? brandCols[0];
    if (brandNameCol) {
      const firstVal = editableRows
        .slice(3)
        .map((r) => String(r[brandNameCol] ?? "").trim())
        .find((v) => v.length > 0);
      if (firstVal) setFindVal(firstVal);
    }
  }, [editableRows.length]);

  const handleSheetChange = useCallback((type: "old" | "new", sheetName: string) => {
    if (type === "old" && oldData) {
      const sd = parseWorkbook(oldData.workbook, sheetName, oldData.rawBytes, oldData.isXlsx);
      setOldData({ ...sd, sheetNames: oldData.sheetNames });
    } else if (type === "new" && newData) {
      const sd = parseWorkbook(newData.workbook, sheetName, newData.rawBytes, newData.isXlsx);
      setNewData({ ...sd, sheetNames: newData.sheetNames });
    }
  }, [oldData, newData]);

  const handleProceedToMapping = () => {
    if (!oldData || !newData) return;
    const autoMapped = autoMatch(newData.headers, oldData.headers);
    setMappings(newData.headers.map((h) => ({ newCol: h, oldCol: autoMapped[h] ?? null })));
    setStep(2);
  };

  const handleMappingChange = (newCol: string, oldCol: string) => {
    setMappings((prev) =>
      prev.map((m) => (m.newCol === newCol ? { ...m, oldCol: oldCol === "__none__" ? null : oldCol } : m))
    );
  };

  const handleProceedToReview = () => {
    if (!oldData || !newData) return;

    let imgLookup: Map<string, string> | undefined;
    let matchCount = 0;

    if (imgData && imgConfig.imgKeyCol && imgConfig.imgUrlCol && imgConfig.oldKeyCol && imgConfig.newTargetCol) {
      imgLookup = new Map();
      for (const row of imgData.rows) {
        const key = String(row[imgConfig.imgKeyCol] ?? "").trim();
        const url = String(row[imgConfig.imgUrlCol] ?? "").trim();
        if (key && url) imgLookup.set(key, url);
      }
      for (const row of oldData.rows) {
        const key = String(row[imgConfig.oldKeyCol] ?? "").trim();
        if (key && imgLookup.has(key)) matchCount++;
      }
      setImgMatchCount(matchCount);
    } else {
      setImgMatchCount(null);
    }

    const rows = buildEditableRows(
      mappings,
      oldData.rows,
      imgLookup,
      imgConfig.oldKeyCol || undefined,
      imgConfig.newTargetCol || undefined
    );
    setEditableRows(rows);
    setSearchQuery("");
    setEditingCell(null);
    setColumnTransforms({});
    setTransformPanelOpen(true);
    setFindReplaceOpen(true);
    setColSelectorOpen(false);
    setStep(3);
  };

  const handleCellChange = (rowIdx: number, col: string, value: string) => {
    setEditableRows((prev) => {
      const updated = [...prev];
      updated[rowIdx] = { ...updated[rowIdx], [col]: value };
      return updated;
    });
  };

  const handleTransformChange = (col: string, field: "prefix" | "suffix", value: string) => {
    setColumnTransforms((prev) => ({
      ...prev,
      [col]: { prefix: prev[col]?.prefix ?? "", suffix: prev[col]?.suffix ?? "", [field]: value },
    }));
  };

  const clearTransform = (col: string) => {
    setColumnTransforms((prev) => {
      const next = { ...prev };
      delete next[col];
      return next;
    });
  };

  const activeTransformCount = Object.values(columnTransforms).filter(
    (t) => t.prefix.trim() !== "" || t.suffix.trim() !== ""
  ).length;

  const BRAND_KEYWORDS = ["brand", "manufacturer", "packer", "description"];
  const isBrandCol = (col: string) => BRAND_KEYWORDS.some((kw) => col.toLowerCase().includes(kw));

  const handleFindReplace = () => {
    if (!findVal.trim() || findReplaceCols.length === 0) return;
    let count = 0;
    setEditableRows((rows) =>
      rows.map((row) => {
        const updated = { ...row };
        for (const col of findReplaceCols) {
          if (col in updated) {
            const newVal = updated[col].split(findVal).join(replaceVal);
            if (newVal !== updated[col]) { updated[col] = newVal; count++; }
          }
        }
        return updated;
      })
    );
    setFindReplaceCount(count);
  };

  const imgConfigComplete =
    !!imgData &&
    !!imgConfig.imgKeyCol &&
    !!imgConfig.imgUrlCol &&
    !!imgConfig.oldKeyCol &&
    !!imgConfig.newTargetCol;

  const generateOutput = async () => {
    if (!newData || !oldData) return;
    setGenerating(true);
    try {
      const mapsForFiller: Mapping[] = mappings.map((m) => ({ newCol: m.newCol, oldCol: m.newCol }));

      const LOCKED_ROWS = 3;
      const dataRows: Record<string, string | number | boolean | null>[] = editableRows.map((row, idx) => {
        const converted: Record<string, string | number | boolean | null> = {};
        for (const [k, v] of Object.entries(row)) {
          const transformed = idx < LOCKED_ROWS ? v : applyTransform(v, columnTransforms[k]);
          if (transformed === "" || transformed === null) { converted[k] = null; continue; }
          const strVal = String(transformed);
          const num = Number(strVal);
          converted[k] = (idx >= LOCKED_ROWS && !isNaN(num) && strVal.trim() !== "" && columnTransforms[k]?.prefix === "" && columnTransforms[k]?.suffix === "") ? num : strVal;
        }
        return converted;
      });

      if (newData.isXlsx) {
        const knownHeaderRowNum = newData.headerRowIndex + 1;
        const { blob } = await fillXlsxTemplate(
          newData.rawBytes,
          newData.selectedSheet,
          mapsForFiller,
          dataRows,
          knownHeaderRowNum
        );
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = newFileName || "flipkart_catalog_filled.xlsx";
        a.click();
        URL.revokeObjectURL(url);
      } else {
        const oldSheetIdx = oldData.workbook.SheetNames.indexOf(oldData.selectedSheet);
        const templateSheetIdx = newData.workbook.SheetNames.indexOf(newData.selectedSheet);

        const formData = new FormData();
        formData.append("template", new Blob([newData.rawBytes], { type: "application/vnd.ms-excel" }), newFileName);
        formData.append("oldCatalog", new Blob([oldData.rawBytes], { type: "application/vnd.ms-excel" }), oldFileName);
        formData.append("mappings", JSON.stringify(mapsForFiller));
        formData.append("templateSheetIdx", String(Math.max(0, templateSheetIdx)));
        formData.append("templateHeaderRow", String(newData.headerRowIndex));
        formData.append("oldSheetIdx", String(Math.max(0, oldSheetIdx)));
        formData.append("oldHeaderRow", String(oldData.headerRowIndex));
        formData.append("dataRows", JSON.stringify(dataRows));

        const apiBaseUrl = (import.meta.env.VITE_API_BASE_URL ?? "").replace(/\/$/, "");
        const fillXlsUrl = apiBaseUrl
          ? `${apiBaseUrl}/api/fill-xls`
          : `${import.meta.env.BASE_URL}api/fill-xls`;

        const response = await fetch(fillXlsUrl, {
          method: "POST",
          body: formData,
        });

        if (!response.ok) {
          const err = await response.json().catch(() => ({ error: "Server error" }));
          throw new Error(err.error || "Server error");
        }

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = newFileName || "flipkart_catalog_filled.xls";
        a.click();
        URL.revokeObjectURL(url);
      }
      setStep(4);
    } finally {
      setGenerating(false);
    }
  };

  const reset = () => {
    setOldData(null); setNewData(null); setImgData(null); setMappings([]); setStep(1);
    setOldFileName(""); setNewFileName(""); setImgFileName("");
    setEditableRows([]); setEditingCell(null); setSearchQuery("");
    setColumnTransforms({}); setTransformPanelOpen(false);
    setImgConfig({ imgKeyCol: "", oldKeyCol: "", imgUrlCol: "", newTargetCol: "" });
    setImgConfigOpen(false); setImgMatchCount(null);
    if (oldInputRef.current) oldInputRef.current.value = "";
    if (newInputRef.current) newInputRef.current.value = "";
    if (imgInputRef.current) imgInputRef.current.value = "";
  };

  const mappedCount = mappings.filter((m) => m.oldCol !== null).length;
  const totalCount = mappings.length;

  const visibleColumns = mappings.filter((m) => m.oldCol !== null).map((m) => m.newCol);

  const liveMatchCount = findVal.trim() && findReplaceCols.length > 0
    ? editableRows.slice(3).reduce((acc, row) => {
        for (const col of findReplaceCols) {
          if (String(row[col] ?? "").includes(findVal)) acc++;
        }
        return acc;
      }, 0)
    : null;

  const TRANSFORM_ALLOWED = ["Seller SKU ID", "Style Code"];
  const transformColumns = visibleColumns.filter((col) => TRANSFORM_ALLOWED.includes(col));

  const filteredRowIndices = editableRows
    .map((row, idx) => ({ row, idx }))
    .filter(({ row }) => {
      if (!searchQuery.trim()) return true;
      const q = searchQuery.toLowerCase();
      return Object.values(row).some((v) => String(v ?? "").toLowerCase().includes(q));
    });

  const STEPS = [
    { n: 1, label: "Files Upload" },
    { n: 2, label: "Column Mapping" },
    { n: 3, label: "Review & Edit" },
    { n: 4, label: "Download" },
  ];

  return (
    <div className="min-h-screen bg-gradient-to-br from-[#1a1a2e] via-[#16213e] to-[#0f3460] text-white">
      <div className="border-b border-white/10 bg-white/5 backdrop-blur-sm">
        <div className="max-w-5xl mx-auto px-4 py-4 flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-[#f9a825] flex items-center justify-center font-bold text-[#1a1a2e] text-sm">FK</div>
          <div>
            <h1 className="text-base font-semibold leading-tight">Flipkart Catalog Mapper</h1>
            <p className="text-xs text-white/50">Purani catalog se naye template mein data transfer karein</p>
          </div>
          {step > 1 && (
            <button onClick={reset} className="ml-auto text-xs text-white/40 hover:text-white/80 border border-white/10 hover:border-white/30 px-3 py-1.5 rounded-md transition-colors">
              Dobara shuru karein
            </button>
          )}
        </div>
      </div>

      <div className="max-w-5xl mx-auto px-4 py-5">
        {/* Steps */}
        <div className="flex items-center gap-2 mb-8">
          {STEPS.map((s, i) => (
            <div key={s.n} className="flex items-center gap-2">
              <div className={`flex items-center gap-2 text-sm ${step >= s.n ? "text-white" : "text-white/30"}`}>
                <div className={`w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold transition-colors ${step > s.n ? "bg-green-500 text-white" : step === s.n ? "bg-[#f9a825] text-[#1a1a2e]" : "bg-white/10"}`}>
                  {step > s.n ? "✓" : s.n}
                </div>
                <span className="hidden sm:inline">{s.label}</span>
              </div>
              {i < STEPS.length - 1 && <div className={`h-px w-6 sm:w-10 ${step > s.n ? "bg-green-500/50" : "bg-white/10"}`} />}
            </div>
          ))}
        </div>

        {/* STEP 1 */}
        {step === 1 && (
          <div className="space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {(["old", "new"] as const).map((type) => {
                const data = type === "old" ? oldData : newData;
                const fileName = type === "old" ? oldFileName : newFileName;
                const ref = type === "old" ? oldInputRef : newInputRef;
                const label = type === "old" ? "Purani Catalog File" : "Flipkart ka Naya Template";
                const hint = type === "old" ? "Apni pehli catalog yahan drag karein ya click karein" : "Flipkart ka naya template yahan drag karein ya click karein";
                const icon = type === "old" ? "📁" : "📋";
                return (
                  <div
                    key={type}
                    className={`rounded-xl border-2 border-dashed transition-all p-6 cursor-pointer ${dragOver === type ? "border-[#f9a825] bg-[#f9a825]/10" : data ? "border-green-500/50 bg-green-500/5" : "border-white/20 hover:border-white/40 bg-white/5"}`}
                    onDragOver={(e) => { e.preventDefault(); setDragOver(type); }}
                    onDragLeave={() => setDragOver(null)}
                    onDrop={(e) => { e.preventDefault(); setDragOver(null); handleFile(e.dataTransfer.files[0], type); }}
                    onClick={() => ref.current?.click()}
                  >
                    <input ref={ref} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleFile(e.target.files?.[0], type)} />
                    <div className="text-center">
                      <div className={`text-3xl mb-3 ${data ? "text-green-400" : "text-white/40"}`}>{data ? "✓" : icon}</div>
                      <p className="font-medium text-sm mb-1">{data ? (type === "old" ? "Purani catalog load ho gayi" : "Naya template load ho gaya") : label}</p>
                      {data ? (
                        <div className="space-y-1">
                          <p className="text-xs text-green-400 font-medium truncate">{fileName}</p>
                          <p className="text-xs text-white/40">{data.rows.length} products • {data.headers.length} columns</p>
                          {type === "new" && !data.isXlsx && (
                            <p className="text-xs text-yellow-400 mt-1">⚠ .xls format — styling partially preserve hogi.</p>
                          )}
                          {type === "new" && data.isXlsx && (
                            <p className="text-xs text-green-400 mt-1">✓ .xlsx format — styling 100% preserve hogi</p>
                          )}
                        </div>
                      ) : (
                        <p className="text-xs text-white/40">{hint}</p>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Optional Image URL File */}
            <div
              className={`rounded-xl border-2 border-dashed transition-all p-5 cursor-pointer ${dragOver === "img" ? "border-blue-400 bg-blue-400/10" : imgData ? "border-blue-400/50 bg-blue-400/5" : "border-white/10 hover:border-white/25 bg-white/[0.02]"}`}
              onDragOver={(e) => { e.preventDefault(); setDragOver("img"); }}
              onDragLeave={() => setDragOver(null)}
              onDrop={(e) => { e.preventDefault(); setDragOver(null); handleFile(e.dataTransfer.files[0], "img"); }}
              onClick={() => imgInputRef.current?.click()}
            >
              <input ref={imgInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleFile(e.target.files?.[0], "img")} />
              <div className="flex items-center gap-4">
                <div className={`text-2xl flex-shrink-0 ${imgData ? "text-blue-400" : "text-white/25"}`}>{imgData ? "✓" : "🖼️"}</div>
                <div className="flex-1 min-w-0">
                  <div className="flex items-center gap-2">
                    <p className="font-medium text-sm">{imgData ? "Image URL File load ho gayi" : "Image URL File"}</p>
                    <span className="text-[10px] text-white/30 bg-white/10 px-1.5 py-0.5 rounded-full">Optional</span>
                  </div>
                  {imgData ? (
                    <div className="flex items-center gap-3 mt-0.5">
                      <p className="text-xs text-blue-400 font-medium truncate">{imgFileName}</p>
                      <p className="text-xs text-white/30">{imgData.rows.length} rows • {imgData.headers.length} columns</p>
                      <button
                        onClick={(e) => { e.stopPropagation(); setImgData(null); setImgFileName(""); setImgConfig({ imgKeyCol: "", oldKeyCol: "", imgUrlCol: "", newTargetCol: "" }); if (imgInputRef.current) imgInputRef.current.value = ""; }}
                        className="text-xs text-white/20 hover:text-red-400 transition-colors ml-auto flex-shrink-0"
                      >
                        ✕ Hatao
                      </button>
                    </div>
                  ) : (
                    <p className="text-xs text-white/25 mt-0.5">Agar alag file mein image URLs hain to yahan upload karein — Step 2 mein link karein</p>
                  )}
                </div>
              </div>
            </div>

            {(oldData?.sheetNames?.length ?? 0) > 1 && (
              <div className="bg-white/5 rounded-xl p-4 border border-white/10">
                <p className="text-xs text-white/50 mb-2">Purani file ka sheet chunein:</p>
                <div className="flex flex-wrap gap-2">
                  {oldData!.sheetNames.map((s) => (
                    <button key={s} onClick={() => handleSheetChange("old", s)} className={`text-xs px-3 py-1.5 rounded-md border transition-colors ${oldData!.selectedSheet === s ? "bg-[#f9a825] text-[#1a1a2e] border-[#f9a825] font-medium" : "border-white/20 text-white/60 hover:border-white/40"}`}>{s}</button>
                  ))}
                </div>
              </div>
            )}
            {(newData?.sheetNames?.length ?? 0) > 1 && (
              <div className="bg-white/5 rounded-xl p-4 border border-white/10">
                <p className="text-xs text-white/50 mb-2">Naye template ka data sheet chunein:</p>
                <div className="flex flex-wrap gap-2">
                  {newData!.sheetNames.map((s) => (
                    <button key={s} onClick={() => handleSheetChange("new", s)} className={`text-xs px-3 py-1.5 rounded-md border transition-colors ${newData!.selectedSheet === s ? "bg-[#f9a825] text-[#1a1a2e] border-[#f9a825] font-medium" : "border-white/20 text-white/60 hover:border-white/40"}`}>{s}</button>
                  ))}
                </div>
              </div>
            )}

            <button
              disabled={!oldData || !newData}
              onClick={handleProceedToMapping}
              className="w-full py-3 rounded-xl font-semibold text-sm transition-all disabled:opacity-30 disabled:cursor-not-allowed bg-[#f9a825] text-[#1a1a2e] hover:bg-[#fbb200] active:scale-[0.99]"
            >
              {!oldData || !newData ? "Dono files upload karein" : "Column Mapping Karein →"}
            </button>
          </div>
        )}

        {/* STEP 2 */}
        {step === 2 && newData && oldData && (
          <div className="space-y-4">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div>
                <h2 className="text-base font-semibold">Columns Map Karein</h2>
                <p className="text-xs text-white/40 mt-0.5">{mappedCount} / {totalCount} columns automatically match hue</p>
              </div>
              <div className="flex items-center gap-3 text-xs text-white/40">
                <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-green-400 inline-block" />Auto-matched</span>
                <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-[#f9a825] inline-block" />Manual</span>
                <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full bg-white/20 inline-block" />Empty</span>
              </div>
            </div>

            <div className="bg-white/5 rounded-xl border border-white/10 divide-y divide-white/5 max-h-[60vh] overflow-y-auto">
              <div className="grid grid-cols-2 gap-4 px-4 py-2.5 bg-white/5 sticky top-0 z-10">
                <span className="text-xs font-medium text-white/50 uppercase tracking-wider">Naye Template ka Column</span>
                <span className="text-xs font-medium text-white/50 uppercase tracking-wider">Purani Catalog ka Column</span>
              </div>
              {mappings.map((m) => {
                const autoMatched = autoMatch([m.newCol], oldData.headers)[m.newCol] === m.oldCol && m.oldCol !== null;
                return (
                  <div key={m.newCol} className="grid grid-cols-2 gap-4 px-4 py-2.5 items-center hover:bg-white/5 transition-colors">
                    <div className="flex items-center gap-2 min-w-0">
                      <div className={`w-2 h-2 rounded-full flex-shrink-0 ${m.oldCol !== null ? (autoMatched ? "bg-green-400" : "bg-[#f9a825]") : "bg-white/20"}`} />
                      <span className="text-sm truncate" title={m.newCol}>{m.newCol}</span>
                    </div>
                    <select
                      value={m.oldCol ?? "__none__"}
                      onChange={(e) => handleMappingChange(m.newCol, e.target.value)}
                      className="bg-white/10 border border-white/20 rounded-lg px-2 py-1.5 text-sm text-white focus:outline-none focus:border-[#f9a825] transition-colors w-full"
                    >
                      <option value="__none__" className="bg-[#1a1a2e] text-white/40">-- koi column nahi --</option>
                      {oldData.headers.map((h) => (
                        <option key={h} value={h} className="bg-[#1a1a2e]">{h}</option>
                      ))}
                    </select>
                  </div>
                );
              })}
            </div>

            {/* Image URL Mapping Panel */}
            {imgData && (
              <div className="rounded-xl border border-blue-400/30 overflow-hidden">
                <button
                  onClick={() => setImgConfigOpen((p) => !p)}
                  className="w-full flex items-center justify-between px-4 py-3 bg-blue-400/10 hover:bg-blue-400/15 transition-colors text-left"
                >
                  <div className="flex items-center gap-2">
                    <span className="text-lg">🖼️</span>
                    <span className="text-sm font-medium text-blue-300">Image URL File se Link Karein</span>
                    {imgConfigComplete && (
                      <span className="text-xs bg-blue-400 text-[#1a1a2e] font-bold px-1.5 py-0.5 rounded-full">✓ Ready</span>
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs text-white/30">{imgFileName}</span>
                    <svg className={`w-4 h-4 text-white/40 transition-transform ${imgConfigOpen ? "rotate-180" : ""}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                    </svg>
                  </div>
                </button>

                {imgConfigOpen && (
                  <div className="bg-blue-400/5 p-4 space-y-3">
                    <p className="text-xs text-white/40">Image file aur purani catalog ke beech common key choose karo, phir URL column aur target column batao.</p>
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="text-xs text-white/50 block mb-1">Image file ka KEY column <span className="text-blue-300">(jaise SKU, Product ID)</span></label>
                        <select
                          value={imgConfig.imgKeyCol}
                          onChange={(e) => setImgConfig((p) => ({ ...p, imgKeyCol: e.target.value }))}
                          className="w-full bg-white/10 border border-white/20 rounded-lg px-2 py-1.5 text-sm text-white focus:outline-none focus:border-blue-400 transition-colors"
                        >
                          <option value="" className="bg-[#1a1a2e] text-white/40">-- chunein --</option>
                          {imgData.headers.map((h) => (
                            <option key={h} value={h} className="bg-[#1a1a2e]">{h}</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="text-xs text-white/50 block mb-1">Purani catalog ka MATCHING column <span className="text-blue-300">(same key)</span></label>
                        <select
                          value={imgConfig.oldKeyCol}
                          onChange={(e) => setImgConfig((p) => ({ ...p, oldKeyCol: e.target.value }))}
                          className="w-full bg-white/10 border border-white/20 rounded-lg px-2 py-1.5 text-sm text-white focus:outline-none focus:border-blue-400 transition-colors"
                        >
                          <option value="" className="bg-[#1a1a2e] text-white/40">-- chunein --</option>
                          {oldData.headers.map((h) => (
                            <option key={h} value={h} className="bg-[#1a1a2e]">{h}</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="text-xs text-white/50 block mb-1">Image file ka URL column <span className="text-blue-300">(jahan URL hai)</span></label>
                        <select
                          value={imgConfig.imgUrlCol}
                          onChange={(e) => setImgConfig((p) => ({ ...p, imgUrlCol: e.target.value }))}
                          className="w-full bg-white/10 border border-white/20 rounded-lg px-2 py-1.5 text-sm text-white focus:outline-none focus:border-blue-400 transition-colors"
                        >
                          <option value="" className="bg-[#1a1a2e] text-white/40">-- chunein --</option>
                          {imgData.headers.map((h) => (
                            <option key={h} value={h} className="bg-[#1a1a2e]">{h}</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="text-xs text-white/50 block mb-1">Naye template mein KAHAN daalna hai <span className="text-blue-300">(target column)</span></label>
                        <select
                          value={imgConfig.newTargetCol}
                          onChange={(e) => setImgConfig((p) => ({ ...p, newTargetCol: e.target.value }))}
                          className="w-full bg-white/10 border border-white/20 rounded-lg px-2 py-1.5 text-sm text-white focus:outline-none focus:border-blue-400 transition-colors"
                        >
                          <option value="" className="bg-[#1a1a2e] text-white/40">-- chunein --</option>
                          {newData.headers.map((h) => (
                            <option key={h} value={h} className="bg-[#1a1a2e]">{h}</option>
                          ))}
                        </select>
                      </div>
                    </div>

                    {imgConfigComplete && (
                      <div className="flex items-center gap-2 bg-blue-400/10 border border-blue-400/30 rounded-lg px-3 py-2">
                        <span className="text-blue-300 text-sm">✓</span>
                        <span className="text-xs text-blue-200">
                          <span className="font-medium">{imgData.rows.length}</span> rows mein se matching products ke image URLs automatically fill ho jayenge
                        </span>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            <div className="flex gap-3">
              <button onClick={() => setStep(1)} className="px-4 py-2.5 rounded-xl text-sm border border-white/20 text-white/60 hover:border-white/40 hover:text-white transition-colors">
                ← Wapas
              </button>
              <button
                onClick={handleProceedToReview}
                className="flex-1 py-2.5 rounded-xl font-semibold text-sm bg-[#f9a825] text-[#1a1a2e] hover:bg-[#fbb200] active:scale-[0.99] transition-all"
              >
                Data Review & Edit Karein ({oldData.rows.length} products) →
              </button>
            </div>
          </div>
        )}

        {/* STEP 3 — Review & Edit */}
        {step === 3 && newData && oldData && (
          <div className="space-y-4">
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div>
                <h2 className="text-base font-semibold">Data Review & Edit Karein</h2>
                <p className="text-xs text-white/40 mt-0.5">Kisi bhi cell par click karke value badal sakte hain, ya pura column prefix/suffix lagayein</p>
              </div>
              <div className="text-xs text-white/40 bg-white/5 border border-white/10 rounded-lg px-3 py-1.5">
                {editableRows.length} products • {visibleColumns.length} columns
              </div>
            </div>

            {imgMatchCount !== null && (
              <div className="flex items-center gap-2 bg-blue-400/10 border border-blue-400/30 rounded-lg px-4 py-2.5">
                <span className="text-blue-300 text-lg">🖼️</span>
                <div>
                  <span className="text-sm text-blue-200 font-medium">{imgMatchCount} products</span>
                  <span className="text-xs text-blue-200/70"> ke image URLs fill ho gaye "{imgConfig.newTargetCol}" column mein</span>
                </div>
              </div>
            )}

            {/* Column Transforms Panel */}
            <div className="rounded-xl border border-white/10 overflow-hidden">
              <button
                onClick={() => setTransformPanelOpen((p) => !p)}
                className="w-full flex items-center justify-between px-4 py-3 bg-white/5 hover:bg-white/10 transition-colors text-left"
              >
                <div className="flex items-center gap-2">
                  <svg className="w-4 h-4 text-[#f9a825]" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 7h.01M7 3h5c.512 0 1.024.195 1.414.586l7 7a2 2 0 010 2.828l-7 7a2 2 0 01-2.828 0l-7-7A1.994 1.994 0 013 12V7a4 4 0 014-4z" />
                  </svg>
                  <span className="text-sm font-medium">Column Prefix / Suffix</span>
                  {activeTransformCount > 0 && (
                    <span className="text-xs bg-[#f9a825] text-[#1a1a2e] font-bold px-1.5 py-0.5 rounded-full">{activeTransformCount}</span>
                  )}
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-white/30">Kisi bhi column ke saare values mein prefix/suffix lagayein</span>
                  <svg className={`w-4 h-4 text-white/40 transition-transform ${transformPanelOpen ? "rotate-180" : ""}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                  </svg>
                </div>
              </button>

              {transformPanelOpen && (
                <div className="bg-white/[0.03] divide-y divide-white/5">
                  <div className="grid grid-cols-[1fr_160px_160px_80px] gap-3 px-4 py-2 border-b border-white/10">
                    <span className="text-xs text-white/30 uppercase tracking-wider">Column</span>
                    <span className="text-xs text-white/30 uppercase tracking-wider">Prefix (aage lagega)</span>
                    <span className="text-xs text-white/30 uppercase tracking-wider">Suffix (peeche lagega)</span>
                    <span className="text-xs text-white/30 uppercase tracking-wider">Preview</span>
                  </div>
                  {transformColumns.map((col) => {
                    const t = columnTransforms[col];
                    const prefix = t?.prefix ?? "";
                    const suffix = t?.suffix ?? "";
                    const sampleVal = editableRows[0]?.[col] ?? "12345";
                    const preview = sampleVal ? applyTransform(sampleVal, { prefix, suffix }) : "—";
                    const hasTransform = prefix !== "" || suffix !== "";
                    return (
                      <div key={col} className="grid grid-cols-[1fr_160px_160px_80px] gap-3 px-4 py-2.5 items-center hover:bg-white/5">
                        <div className="flex items-center gap-2 min-w-0">
                          {hasTransform && <span className="w-1.5 h-1.5 rounded-full bg-[#f9a825] flex-shrink-0" />}
                          <span className="text-sm text-white/80 truncate" title={col}>{col}</span>
                        </div>
                        <input
                          type="text"
                          placeholder="jaise: FK-"
                          value={prefix}
                          onChange={(e) => handleTransformChange(col, "prefix", e.target.value)}
                          className="bg-white/10 border border-white/15 rounded-lg px-2.5 py-1.5 text-sm text-white placeholder-white/20 focus:outline-none focus:border-[#f9a825] transition-colors w-full font-mono"
                        />
                        <input
                          type="text"
                          placeholder="jaise: -IND"
                          value={suffix}
                          onChange={(e) => handleTransformChange(col, "suffix", e.target.value)}
                          className="bg-white/10 border border-white/15 rounded-lg px-2.5 py-1.5 text-sm text-white placeholder-white/20 focus:outline-none focus:border-[#f9a825] transition-colors w-full font-mono"
                        />
                        <div className="flex items-center gap-1.5 min-w-0">
                          <span className={`text-xs font-mono truncate ${hasTransform ? "text-[#f9a825]" : "text-white/30"}`} title={preview}>
                            {preview}
                          </span>
                          {hasTransform && (
                            <button onClick={() => clearTransform(col)} className="text-white/20 hover:text-red-400 transition-colors flex-shrink-0 text-xs" title="Clear">✕</button>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Find & Replace */}
            <div className="rounded-xl border border-white/10 overflow-hidden">
              <button
                onClick={() => setFindReplaceOpen((p) => !p)}
                className="w-full flex items-center justify-between px-4 py-3 bg-white/5 hover:bg-white/10 transition-colors text-left"
              >
                <div className="flex items-center gap-2">
                  <svg className="w-4 h-4 text-purple-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                  </svg>
                  <span className="text-sm font-medium">Find & Replace (Brand Name)</span>
                  {findReplaceCount !== null && (
                    <span className="text-xs bg-purple-500 text-white font-bold px-1.5 py-0.5 rounded-full">{findReplaceCount} replaced</span>
                  )}
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-white/30">Brand / Manufacturer name puri sheet mein badlein</span>
                  <svg className={`w-4 h-4 text-white/40 transition-transform ${findReplaceOpen ? "rotate-180" : ""}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                  </svg>
                </div>
              </button>

              {findReplaceOpen && (
                <div className="bg-purple-400/5 p-4 space-y-4">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                      <label className="text-xs text-white/50 block mb-1">Dhundhna hai <span className="text-purple-300">(purana text)</span></label>
                      <input
                        type="text"
                        placeholder="jaise: Nike"
                        value={findVal}
                        onChange={(e) => { setFindVal(e.target.value); setFindReplaceCount(null); }}
                        className="w-full bg-white/10 border border-white/20 rounded-lg px-3 py-2 text-sm text-white placeholder-white/25 focus:outline-none focus:border-purple-400 transition-colors font-mono"
                      />
                    </div>
                    <div>
                      <label className="text-xs text-white/50 block mb-1">Badalna hai <span className="text-purple-300">(naya text)</span></label>
                      <input
                        type="text"
                        placeholder="jaise: Adidas"
                        value={replaceVal}
                        onChange={(e) => { setReplaceVal(e.target.value); setFindReplaceCount(null); }}
                        className="w-full bg-white/10 border border-white/20 rounded-lg px-3 py-2 text-sm text-white placeholder-white/25 focus:outline-none focus:border-purple-400 transition-colors font-mono"
                      />
                    </div>
                  </div>

                  <div>
                    <button
                      onClick={() => setColSelectorOpen((p) => !p)}
                      className="flex items-center gap-2 text-xs text-white/50 hover:text-white/70 transition-colors w-full text-left"
                    >
                      <svg className={`w-3 h-3 transition-transform flex-shrink-0 ${colSelectorOpen ? "rotate-90" : ""}`} fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                      </svg>
                      <span>Kin columns mein badlein</span>
                      <span className="text-purple-300 font-medium">({findReplaceCols.length} selected)</span>
                    </button>
                    {colSelectorOpen && (
                      <div className="flex flex-wrap gap-2 mt-2">
                        {visibleColumns.map((col) => {
                          const checked = findReplaceCols.includes(col);
                          return (
                            <button
                              key={col}
                              onClick={() => setFindReplaceCols((prev) => checked ? prev.filter((c) => c !== col) : [...prev, col])}
                              className={`px-2.5 py-1 rounded-lg text-xs font-medium border transition-colors ${checked ? "bg-purple-500/20 border-purple-400/50 text-purple-200" : "bg-white/5 border-white/15 text-white/40 hover:border-white/30 hover:text-white/60"}`}
                            >
                              {col}
                            </button>
                          );
                        })}
                      </div>
                    )}
                  </div>

                  <div className="flex items-center gap-3 flex-wrap">
                    <button
                      onClick={handleFindReplace}
                      disabled={!findVal.trim() || findReplaceCols.length === 0}
                      className="px-4 py-2 rounded-lg text-sm font-medium bg-purple-500 hover:bg-purple-600 disabled:opacity-30 disabled:cursor-not-allowed text-white transition-colors"
                    >
                      Replace All
                    </button>
                    {liveMatchCount !== null && findReplaceCount === null && (
                      <span className={`text-xs font-medium ${liveMatchCount > 0 ? "text-purple-300" : "text-white/30"}`}>
                        {liveMatchCount > 0 ? `${liveMatchCount} cells match honge` : "Koi match nahi mila"}
                      </span>
                    )}
                    {findReplaceCount !== null && (
                      <span className="text-xs text-purple-300">
                        {findReplaceCount > 0 ? `✓ ${findReplaceCount} cells mein replace hua` : "Koi match nahi mila"}
                      </span>
                    )}
                  </div>
                </div>
              )}
            </div>

            {/* Search */}
            <div className="relative">
              <svg className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-white/30" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-4.35-4.35M17 11A6 6 0 1 1 5 11a6 6 0 0 1 12 0z" />
              </svg>
              <input
                type="text"
                placeholder="Koi bhi value search karein..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full bg-white/5 border border-white/10 rounded-lg pl-8 pr-3 py-2 text-sm text-white placeholder-white/30 focus:outline-none focus:border-[#f9a825] transition-colors"
              />
              {searchQuery && (
                <button onClick={() => setSearchQuery("")} className="absolute right-3 top-1/2 -translate-y-1/2 text-white/30 hover:text-white/60 text-xs">✕</button>
              )}
            </div>

            {visibleColumns.length === 0 ? (
              <div className="text-center py-10 text-white/30 text-sm">Koi column mapped nahi hai. Wapas jaake columns map karein.</div>
            ) : (
              <div className="rounded-xl border border-white/10 overflow-auto max-h-[50vh]">
                <table className="w-full text-sm border-collapse min-w-max">
                  <thead className="sticky top-0 z-10">
                    <tr className="bg-[#0f3460]">
                      <th className="px-3 py-2.5 text-left text-xs font-medium text-white/40 uppercase tracking-wider border-b border-white/10 w-10">#</th>
                      {visibleColumns.map((col) => {
                        const t = columnTransforms[col];
                        const hasT = t && (t.prefix !== "" || t.suffix !== "");
                        const isImgTarget = imgMatchCount !== null && col === imgConfig.newTargetCol;
                        return (
                          <th key={col} className={`px-3 py-2.5 text-left border-b border-white/10 min-w-[150px] max-w-[240px] ${isImgTarget ? "bg-blue-400/10" : ""}`}>
                            <span className={`text-xs font-medium uppercase tracking-wider truncate block ${isImgTarget ? "text-blue-300" : "text-white/60"}`} title={col}>
                              {isImgTarget && "🖼️ "}{col}
                            </span>
                            {hasT && (
                              <span className="text-[10px] text-[#f9a825]/70 font-mono mt-0.5 block truncate">
                                {t.prefix && <span>"{t.prefix}"</span>}{t.prefix && t.suffix && " + "}{t.suffix && <span>"{t.suffix}"</span>}
                              </span>
                            )}
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-white/5">
                    {filteredRowIndices.length === 0 ? (
                      <tr>
                        <td colSpan={visibleColumns.length + 1} className="text-center py-8 text-white/30">Koi result nahi mila</td>
                      </tr>
                    ) : (
                      filteredRowIndices.map(({ row, idx }) => {
                        const isLocked = idx < 3;
                        return (
                        <tr key={idx} className={`transition-colors group ${isLocked ? "bg-white/[0.02]" : "hover:bg-white/5"}`}>
                          <td className="px-3 py-1.5 text-xs text-white/20 select-none">
                            <div className="flex items-center gap-1">
                              {isLocked ? (
                                <span title="Yeh row lock hai — edit nahi ho sakti" className="text-white/20">🔒</span>
                              ) : (
                                <span>{idx + 1}</span>
                              )}
                            </div>
                          </td>
                          {visibleColumns.map((col) => {
                            const isEditing = !isLocked && editingCell?.rowIdx === idx && editingCell?.col === col;
                            const rawVal = row[col] ?? "";
                            const t = columnTransforms[col];
                            const hasT = !isLocked && t && (t.prefix !== "" || t.suffix !== "");
                            const isImgTarget = imgMatchCount !== null && col === imgConfig.newTargetCol;
                            return (
                              <td key={col} className={`px-1.5 py-1 min-w-[150px] max-w-[240px] ${isImgTarget && rawVal ? "bg-blue-400/5" : ""}`}>
                                {isEditing ? (
                                  <input
                                    autoFocus
                                    value={rawVal}
                                    onChange={(e) => handleCellChange(idx, col, e.target.value)}
                                    onBlur={() => setEditingCell(null)}
                                    onKeyDown={(e) => {
                                      if (e.key === "Enter" || e.key === "Escape") setEditingCell(null);
                                    }}
                                    className="w-full bg-[#f9a825]/10 border border-[#f9a825] rounded px-2 py-1 text-sm text-white focus:outline-none"
                                  />
                                ) : isLocked ? (
                                  <div className="px-2 py-1 truncate select-none" title="Lock hai — edit nahi ho sakta">
                                    {rawVal === "" ? (
                                      <span className="text-white/15 italic">—</span>
                                    ) : (
                                      <span className="text-white/35">{rawVal}</span>
                                    )}
                                  </div>
                                ) : (
                                  <div
                                    onClick={() => setEditingCell({ rowIdx: idx, col })}
                                    title="Click karein edit karne ke liye"
                                    className={`px-2 py-1 rounded cursor-pointer truncate transition-colors group-hover:bg-white/5 hover:bg-white/10 ${rawVal === "" ? "text-white/20 italic" : ""}`}
                                  >
                                    {rawVal === "" ? (
                                      <span>—</span>
                                    ) : hasT ? (
                                      <span>
                                        {t.prefix && <span className="text-[#f9a825]/80 font-mono">{t.prefix}</span>}
                                        <span className="text-white/80">{rawVal}</span>
                                        {t.suffix && <span className="text-[#f9a825]/80 font-mono">{t.suffix}</span>}
                                      </span>
                                    ) : (
                                      <span className={isImgTarget && rawVal ? "text-blue-300" : "text-white/80"}>{rawVal}</span>
                                    )}
                                  </div>
                                )}
                              </td>
                            );
                          })}
                        </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            )}

            {searchQuery && (
              <p className="text-xs text-white/30 text-center">{filteredRowIndices.length} / {editableRows.length} rows dikh rahe hain</p>
            )}

            <div className="flex gap-3">
              <button onClick={() => setStep(2)} className="px-4 py-2.5 rounded-xl text-sm border border-white/20 text-white/60 hover:border-white/40 hover:text-white transition-colors">
                ← Wapas
              </button>
              <button
                onClick={generateOutput}
                disabled={generating}
                className="flex-1 py-2.5 rounded-xl font-semibold text-sm bg-[#f9a825] text-[#1a1a2e] hover:bg-[#fbb200] active:scale-[0.99] transition-all disabled:opacity-60 disabled:cursor-wait"
              >
                {generating ? "File ban rahi hai..." : `Filled Template Download Karein (${editableRows.length} products)`}
              </button>
            </div>
          </div>
        )}

        {/* STEP 4 — Success */}
        {step === 4 && (
          <div className="text-center py-16 space-y-6">
            <div className="text-6xl">🎉</div>
            <div>
              <h2 className="text-xl font-bold mb-2">File Download Ho Gayi!</h2>
              <p className="text-white/50 text-sm">
                {newData?.isXlsx
                  ? "Flipkart template ki original styling bilkul waise hi rahegi — colors, borders, merged cells sab preserve."
                  : "File ready hai. Best results ke liye aage se .xlsx template use karein."}
                <br />Ab isse seedha Flipkart Seller Hub mein upload kar sakte hain.
              </p>
            </div>
            <div className="bg-white/5 rounded-xl border border-white/10 p-4 max-w-sm mx-auto text-left space-y-2">
              <div className="flex justify-between text-sm">
                <span className="text-white/40">Products transfer kiye</span>
                <span className="font-medium text-green-400">{editableRows.length}</span>
              </div>
              <div className="flex justify-between text-sm">
                <span className="text-white/40">Columns mapped</span>
                <span className="font-medium">{mappedCount} / {totalCount}</span>
              </div>
              {imgMatchCount !== null && (
                <div className="flex justify-between text-sm">
                  <span className="text-white/40">Image URLs fill hue</span>
                  <span className="font-medium text-blue-300">{imgMatchCount} products</span>
                </div>
              )}
              {activeTransformCount > 0 && (
                <div className="flex justify-between text-sm">
                  <span className="text-white/40">Prefix/Suffix lagaye</span>
                  <span className="font-medium text-[#f9a825]">{activeTransformCount} columns</span>
                </div>
              )}
              <div className="flex justify-between text-sm">
                <span className="text-white/40">Format</span>
                <span className={`font-medium ${newData?.isXlsx ? "text-green-400" : "text-yellow-400"}`}>{newData?.isXlsx ? "XLSX (full style)" : "XLS (partial)"}</span>
              </div>
            </div>
            <div className="flex justify-center gap-3">
              <button onClick={reset} className="px-6 py-2.5 rounded-xl text-sm font-medium bg-[#f9a825] text-[#1a1a2e] hover:bg-[#fbb200] transition-colors">
                Dobara karein
              </button>
              <button onClick={() => setStep(3)} className="px-6 py-2.5 rounded-xl text-sm border border-white/20 text-white/60 hover:text-white hover:border-white/40 transition-colors">
                Data phir se edit karein
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
