import { useState, useRef, useCallback, useEffect } from "react";
import JSZip from "jszip";

type LoadedImage = {
  name: string;
  orig: HTMLImageElement;
  objectUrl: string;
};

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  return [
    parseInt(h.substring(0, 2), 16),
    parseInt(h.substring(2, 4), 16),
    parseInt(h.substring(4, 6), 16),
  ];
}

function rgbToHex(r: number, g: number, b: number): string {
  return `#${[r, g, b].map(v => v.toString(16).padStart(2, "0")).join("")}`;
}

/**
 * AUTO BACKGROUND DETECTION — FLOOD FILL FROM EDGES
 */
function medianOf(arr: number[]): number {
  arr.sort((a, b) => a - b);
  return arr[Math.floor(arr.length / 2)];
}

function processImage(
  img: HTMLImageElement,
  newR: number,
  newG: number,
  newB: number,
  tolerance: number,
  fringeIter: number
): { canvas: HTMLCanvasElement; detectedBg: [number, number, number]; confidence: number } {
  const canvas = document.createElement("canvas");
  canvas.width  = img.naturalWidth;
  canvas.height = img.naturalHeight;
  const ctx = canvas.getContext("2d")!;
  ctx.drawImage(img, 0, 0);

  const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const d = imageData.data;
  const W = canvas.width;
  const H = canvas.height;
  const total = W * H;

  // ── Step 1: sample ALL 4 border edges, use median per channel ─────────────
  const STEP = Math.max(1, Math.floor(Math.min(W, H) / 40)); // ~40 samples per edge
  const rs: number[] = [], gs: number[] = [], bs: number[] = [];
  const addSample = (x: number, y: number) => {
    const i = (y * W + x) * 4;
    rs.push(d[i]); gs.push(d[i + 1]); bs.push(d[i + 2]);
  };
  for (let x = 0; x < W; x += STEP) { addSample(x, 0); addSample(x, H - 1); }
  for (let y = 0; y < H; y += STEP) { addSample(0, y); addSample(W - 1, y); }
  const bgR = medianOf(rs);
  const bgG = medianOf(gs);
  const bgB = medianOf(bs);

  function dist(r: number, g: number, b: number): number {
    const dr = r - bgR, dg = g - bgG, db = b - bgB;
    return Math.sqrt(dr * dr + dg * dg + db * db);
  }
  function similar(r: number, g: number, b: number): boolean {
    return dist(r, g, b) <= tolerance;
  }

  // ── Step 1b: confidence — check center region vs bg color ─────────────────
  // Low confidence = center pixels are also "similar" to bg (white-on-white)
  const cx = Math.floor(W / 2), cy = Math.floor(H / 2);
  const cw = Math.floor(W / 4), ch = Math.floor(H / 4);
  let centerTotal = 0, centerSimilar = 0;
  for (let y = cy - ch; y <= cy + ch; y += 4) {
    for (let x = cx - cw; x <= cx + cw; x += 4) {
      if (x < 0 || x >= W || y < 0 || y >= H) continue;
      const i = (y * W + x) * 4;
      centerTotal++;
      if (similar(d[i], d[i + 1], d[i + 2])) centerSimilar++;
    }
  }
  // confidence: 0 = all center pixels match bg (bad), 1 = none match (good)
  const confidence = centerTotal > 0 ? 1 - centerSimilar / centerTotal : 1;

  // ── Step 2: BFS flood-fill from all border pixels ─────────────────────────
  const visited = new Uint8Array(total);
  const isBg    = new Uint8Array(total);
  const queue: number[] = [];

  const seed = (idx: number) => {
    if (visited[idx]) return;
    visited[idx] = 1;
    const px = idx * 4;
    if (d[px + 3] === 0 || similar(d[px], d[px + 1], d[px + 2])) {
      isBg[idx] = 1;
      queue.push(idx);
    }
  };

  for (let x = 0; x < W; x++) { seed(x); seed((H - 1) * W + x); }
  for (let y = 1; y < H - 1; y++) { seed(y * W); seed(y * W + W - 1); }

  let head = 0;
  while (head < queue.length) {
    const idx = queue[head++];
    const x = idx % W;
    const y = Math.floor(idx / W);
    if (y > 0)     seed(idx - W);
    if (y < H - 1) seed(idx + W);
    if (x > 0)     seed(idx - 1);
    if (x < W - 1) seed(idx + 1);
  }

  // ── Step 3: fringe-removal pass ───────────────────────────────────────────
  for (let iter = 0; iter < fringeIter; iter++) {
    const next = new Uint8Array(isBg);
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        const idx = y * W + x;
        if (isBg[idx]) continue;
        const touchesBg =
          (y > 0     && isBg[idx - W]) ||
          (y < H - 1 && isBg[idx + W]) ||
          (x > 0     && isBg[idx - 1]) ||
          (x < W - 1 && isBg[idx + 1]);
        if (touchesBg) {
          const px = idx * 4;
          const bright = d[px] * 0.299 + d[px + 1] * 0.587 + d[px + 2] * 0.114;
          if (bright > 160) next[idx] = 1;
        }
      }
    }
    isBg.set(next);
  }

  // ── Step 4: apply replacement colour ─────────────────────────────────────
  for (let i = 0; i < total; i++) {
    if (isBg[i]) {
      d[i * 4]     = newR;
      d[i * 4 + 1] = newG;
      d[i * 4 + 2] = newB;
      d[i * 4 + 3] = 255;
    }
  }

  ctx.putImageData(imageData, 0, 0);
  return { canvas, detectedBg: [bgR, bgG, bgB], confidence };
}

const PRESETS = [
  { label: "White",      hex: "#ffffff" },
  { label: "Off-White",  hex: "#f5f5f5" },
  { label: "Light Grey", hex: "#e0e0e0" },
  { label: "Grey",       hex: "#9e9e9e" },
  { label: "Black",      hex: "#000000" },
  { label: "FK Blue",    hex: "#2874f0" },
];

export default function BgColorTweaker() {
  const [images,      setImages]      = useState<LoadedImage[]>([]);
  const [selectedIdx, setSelectedIdx] = useState(0);
  const [bgColor,     setBgColor]     = useState("#ffffff");
  const [tol,         setTol]         = useState(35);
  const [fringe,      setFringe]      = useState(3);
  const [saving,      setSaving]      = useState(false);
  const [status,      setStatus]      = useState("Koi image upload nahi hui.");
  const [detectedBg,  setDetectedBg]  = useState<[number, number, number] | null>(null);
  const [confidence,  setConfidence]  = useState<number | null>(null);

  const previewCanvasRef = useRef<HTMLCanvasElement>(null);
  const fileInputRef     = useRef<HTMLInputElement>(null);
  const colorInputRef    = useRef<HTMLInputElement>(null);

  const [nr, ng, nb] = hexToRgb(bgColor);

  const redraw = useCallback(() => {
    const canvas = previewCanvasRef.current;
    if (!canvas || images.length === 0) return;
    const img = images[selectedIdx]?.orig;
    if (!img) return;
    const { canvas: processed, detectedBg: dbg, confidence: conf } = processImage(img, nr, ng, nb, tol, fringe);
    setDetectedBg(dbg);
    setConfidence(conf);
    const maxW = (canvas.parentElement?.clientWidth ?? 600) - 32;
    const maxH = 440;
    const scale = Math.min(maxW / processed.width, maxH / processed.height, 1);
    canvas.width  = Math.round(processed.width  * scale);
    canvas.height = Math.round(processed.height * scale);
    canvas.getContext("2d")!.drawImage(processed, 0, 0, canvas.width, canvas.height);
  }, [images, selectedIdx, nr, ng, nb, tol, fringe]);

  useEffect(() => { redraw(); }, [redraw]);

  const handleUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files ?? []);
    if (!files.length) return;
    images.forEach(im => URL.revokeObjectURL(im.objectUrl));
    const loaded: LoadedImage[] = [];
    let done = 0;
    files.forEach(file => {
      const url = URL.createObjectURL(file);
      const imgEl = new Image();
      imgEl.onload = () => {
        loaded.push({ name: file.name, orig: imgEl, objectUrl: url });
        if (++done === files.length) {
          loaded.sort((a, b) => a.name.localeCompare(b.name));
          setImages(loaded);
          setSelectedIdx(0);
          setDetectedBg(null);
          setConfidence(null);
          setStatus(`${loaded.length} image(s) load ho gayi.`);
        }
      };
      imgEl.src = url;
    });
    e.target.value = "";
  };

  const handleSaveZip = async () => {
    if (!images.length) return;
    setSaving(true);
    setStatus("ZIP ban raha hai…");
    try {
      const zip = new JSZip();
      for (const im of images) {
        const { canvas: processed } = processImage(im.orig, nr, ng, nb, tol, fringe);
        const ext  = im.name.split(".").pop()?.toLowerCase() ?? "png";
        const mime = ext === "png" ? "image/png" : "image/jpeg";
        const blob: Blob = await new Promise(res =>
          processed.toBlob(b => res(b!), mime, 0.95)
        );
        zip.folder(im.name.replace(/\.[^.]+$/, ""))!.file(im.name, blob);
      }
      const zipBlob = await zip.generateAsync({ type: "blob" });
      const url = URL.createObjectURL(zipBlob);
      Object.assign(document.createElement("a"), { href: url, download: "processed_images.zip" }).click();
      URL.revokeObjectURL(url);
      setStatus(`✅ ${images.length} image(s) save ho gayi.`);
    } catch (err) {
      setStatus("❌ Export fail ho gaya.");
      console.error(err);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#1a1a2e] text-white px-4 py-8">
      <div className="max-w-5xl mx-auto space-y-6">

        <div>
          <h1 className="text-2xl font-bold text-[#f9a825]">Auto Background Remover</h1>
          <p className="text-white/50 text-sm mt-1">
            Background <em>auto-detect</em> hoga — model/product ko bilkul touch nahi karega.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-3">
          <button
            onClick={() => fileInputRef.current?.click()}
            className="px-5 py-2.5 rounded-xl bg-[#f9a825] text-[#1a1a2e] font-semibold text-sm hover:bg-[#fbb200] transition-colors"
          >
            📁 Images Upload Karo
          </button>
          <input ref={fileInputRef} type="file" accept=".png,.jpg,.jpeg" multiple className="hidden" onChange={handleUpload} />

          <button
            onClick={handleSaveZip}
            disabled={saving || images.length === 0}
            className="px-5 py-2.5 rounded-xl bg-[#2e7d32] text-white font-semibold text-sm hover:bg-[#388e3c] transition-colors disabled:opacity-40 disabled:cursor-not-allowed"
          >
            {saving ? "Saving…" : `💾 Save All to ZIP${images.length > 0 ? ` (${images.length})` : ""}`}
          </button>

          <span className="text-white/40 text-sm ml-auto">{status}</span>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-[300px_1fr] gap-6">

          {/* Left panel — all settings */}
          <div className="space-y-4">

            {/* Image list */}
            {images.length > 0 && (
              <div className="bg-[#12122a] rounded-2xl border border-white/10 p-3 space-y-1 max-h-48 overflow-y-auto">
                <p className="text-xs text-white/40 uppercase tracking-wider mb-2">
                  Images <span className="text-white/25 normal-case">({images.length})</span>
                </p>
                {images.map((im, i) => (
                  <button key={im.name} onClick={() => setSelectedIdx(i)}
                    className={`w-full text-left px-3 py-2 rounded-lg text-sm transition-colors truncate
                      ${i === selectedIdx ? "bg-[#f9a825]/20 text-[#f9a825] font-medium" : "text-white/60 hover:text-white hover:bg-white/5"}`}>
                    {im.name}
                  </button>
                ))}
              </div>
            )}

            {/* All settings in one card */}
            <div className="bg-[#12122a] rounded-2xl border border-white/10 p-4 space-y-5">
              <p className="text-xs text-white/40 uppercase tracking-wider">Settings</p>

              {/* Color picker */}
              <div className="space-y-3">
                <p className="text-xs text-white/60 font-semibold">Naya Background Color</p>

                {/* Presets */}
                <div className="flex flex-wrap gap-2">
                  {PRESETS.map((p) => (
                    <button
                      key={p.hex}
                      title={p.label}
                      onClick={() => setBgColor(p.hex)}
                      className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs border transition-colors ${
                        bgColor.toLowerCase() === p.hex
                          ? "border-[#f9a825] text-[#f9a825] bg-[#f9a825]/10"
                          : "border-white/15 text-white/50 hover:border-white/30 hover:text-white/80"
                      }`}
                    >
                      <span
                        className="w-3 h-3 rounded-sm border border-white/20 flex-shrink-0"
                        style={{ background: p.hex }}
                      />
                      {p.label}
                    </button>
                  ))}
                </div>

                {/* Color picker input */}
                <div
                  className="flex items-center gap-3 bg-white/5 border border-white/15 rounded-xl px-3 py-2.5 cursor-pointer hover:border-white/30 transition-colors"
                  onClick={() => colorInputRef.current?.click()}
                >
                  <span
                    className="w-8 h-8 rounded-lg border border-white/20 flex-shrink-0 shadow-md"
                    style={{ background: bgColor }}
                  />
                  <div className="flex-1 min-w-0">
                    <p className="text-xs text-white/40 mb-0.5">Custom color chunein</p>
                    <p className="text-sm font-mono text-white/80">{bgColor.toUpperCase()}</p>
                  </div>
                  <svg className="w-4 h-4 text-white/30" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 21a4 4 0 01-4-4V5a2 2 0 012-2h4a2 2 0 012 2v12a4 4 0 01-4 4zm0 0h12a2 2 0 002-2v-4a2 2 0 00-2-2h-2.343M11 7.343l1.657-1.657a2 2 0 012.828 0l2.829 2.829a2 2 0 010 2.828l-8.486 8.485M7 17h.01" />
                  </svg>
                  <input
                    ref={colorInputRef}
                    type="color"
                    value={bgColor}
                    onChange={(e) => setBgColor(e.target.value)}
                    className="sr-only"
                  />
                </div>

                {/* Auto-detected bg inline */}
                {detectedBg && (
                  <div className="flex items-center gap-2 text-xs text-white/40">
                    <span>Auto-detected BG:</span>
                    <span
                      className="w-4 h-4 rounded border border-white/20"
                      style={{ background: rgbToHex(...detectedBg) }}
                    />
                    <span className="font-mono">{rgbToHex(...detectedBg).toUpperCase()}</span>
                  </div>
                )}
              </div>

              <div className="border-t border-white/10" />

              {/* Sensitivity */}
              <div className="space-y-1.5">
                <div className="flex justify-between text-xs">
                  <span className="text-white/70 font-semibold">Sensitivity</span>
                  <span className="text-white/50 font-mono">{tol}</span>
                </div>
                <input type="range" min={5} max={100} value={tol}
                  onChange={e => setTol(Number(e.target.value))}
                  className="w-full h-2 rounded-full appearance-none cursor-pointer"
                  style={{ accentColor: "#f9a825" }} />
                <div className="flex justify-between text-[10px] text-white/25">
                  <span>Tight (safe)</span><span>Loose (aggressive)</span>
                </div>
              </div>

              {/* Shadow / Halo */}
              <div className="space-y-1.5">
                <div className="flex justify-between text-xs">
                  <span className="text-white/70 font-semibold">Shadow / Halo Removal</span>
                  <span className="text-white/50 font-mono">{fringe}</span>
                </div>
                <input type="range" min={0} max={12} value={fringe}
                  onChange={e => setFringe(Number(e.target.value))}
                  className="w-full h-2 rounded-full appearance-none cursor-pointer"
                  style={{ accentColor: "#9b59b6" }} />
                <div className="flex justify-between text-[10px] text-white/25">
                  <span>Off</span><span>Strong</span>
                </div>
              </div>

              <p className="text-[11px] text-white/25 leading-relaxed">
                BFS flood-fill corners se shuru hota hai — main subject kabhi cut nahi hoga.
              </p>
            </div>
          </div>

          {/* Preview */}
          <div className="space-y-3">
            {/* Confidence warning */}
            {confidence !== null && confidence < 0.4 && (
              <div className="flex items-start gap-3 bg-red-500/10 border border-red-500/30 rounded-xl px-4 py-3">
                <span className="text-red-400 text-lg flex-shrink-0">⚠️</span>
                <div>
                  <p className="text-sm font-medium text-red-300">Background detect nahi ho pa raha</p>
                  <p className="text-xs text-red-300/70 mt-0.5">
                    {confidence < 0.15
                      ? "Garment aur background ka color almost same hai (jaise white product on white bg). Yeh tool is image ke liye kaam nahi karega — color-based algorithm same colors mein fark nahi kar sakta."
                      : "Product ka color background se bahut similar hai. Sensitivity slider thoda badhao ya Fringe ko 0 karo."}
                  </p>
                </div>
              </div>
            )}
            {confidence !== null && confidence >= 0.4 && confidence < 0.65 && (
              <div className="flex items-start gap-3 bg-yellow-500/10 border border-yellow-500/20 rounded-xl px-4 py-3">
                <span className="text-yellow-400 text-base flex-shrink-0">⚡</span>
                <p className="text-xs text-yellow-300/80">
                  Kuch edges thodi tricky hain. Result check karo — agar garment cut ho raha hai to Sensitivity slider thoda kam karo.
                </p>
              </div>
            )}

            <div className="bg-[#12122a] rounded-2xl border border-white/10 flex items-center justify-center min-h-[320px] p-4 overflow-hidden">
              {images.length === 0 ? (
                <div onClick={() => fileInputRef.current?.click()}
                  className="flex flex-col items-center gap-3 text-white/30 cursor-pointer hover:text-white/50 transition-colors">
                  <span className="text-5xl">🖼️</span>
                  <p className="text-sm">Yahan preview dikhega</p>
                  <p className="text-xs">Click karke images upload karo</p>
                </div>
              ) : (
                <canvas ref={previewCanvasRef} className="max-w-full rounded-xl shadow-lg" />
              )}
            </div>
          </div>
        </div>

      </div>
    </div>
  );
}
