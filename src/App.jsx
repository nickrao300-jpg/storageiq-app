import { useState, useMemo, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ── Scoring engine ────────────────────────────────────────────────────────────
function scoreCard(fields) {
  let fin = 0, mkt = 0, comp = 0, ops = 0, deal = 0;
  const occ = parseFloat(fields.physOcc) || 0;
  fin += occ >= 90 ? 10 : occ >= 85 ? 8 : occ >= 80 ? 6 : occ >= 75 ? 4 : 0;
  const eocc = parseFloat(fields.econOcc) || 0;
  fin += eocc >= 88 ? 8 : eocc >= 82 ? 6 : eocc >= 75 ? 4 : 0;
  const noi = parseFloat(fields.noiMargin) || 0;
  fin += noi >= 55 ? 8 : noi >= 50 ? 6 : noi >= 45 ? 4 : 0;
  const cap = parseFloat(fields.capRate) || 0;
  fin += cap >= 7 ? 9 : cap >= 6 ? 7 : cap >= 5 ? 4 : 0;
  const mhi = parseFloat(fields.mhi) || 0;
  mkt += mhi >= 75000 ? 8 : mhi >= 65000 ? 6 : mhi >= 55000 ? 4 : mhi >= 50000 ? 2 : 0;
  mkt += fields.popGrowth === "growing" ? 7 : fields.popGrowth === "flat" ? 4 : 0;
  const sqft = parseFloat(fields.sqftPerCapita) || 0;
  mkt += sqft > 0 ? (sqft < 7 ? 10 : sqft <= 9 ? 6 : 0) : 0;
  const comps = parseInt(fields.competitors) || 0;
  comp += comps <= 2 ? 10 : comps <= 4 ? 7 : comps <= 6 ? 4 : 0;
  comp += fields.streetRates === "at-above" ? 10 : fields.streetRates === "slightly-below" ? 6 : fields.streetRates === "well-below" ? 2 : 0;
  const unitTypes = parseInt(fields.unitTypes) || 0;
  ops += unitTypes >= 4 ? 4 : unitTypes === 3 ? 2 : 0;
  const cc = parseFloat(fields.climateControl) || 0;
  ops += cc >= 30 ? 4 : cc >= 20 ? 2 : 0;
  const del = parseFloat(fields.delinquency) || 0;
  ops += del > 0 ? (del < 3 ? 4 : del <= 6 ? 2 : 0) : 0;
  deal += fields.sellerFinancing ? 4 : 0;
  deal += fields.valueAdd?.trim() ? 4 : 0;
  const total = fin + mkt + comp + ops + deal;
  const rec = total >= 80 ? "GO" : total >= 65 ? "CONDITIONAL GO" : total >= 50 ? "SOFT PASS" : "PASS";
  return { fin, mkt, comp, ops, deal, total, rec };
}

// ── Helpers ───────────────────────────────────────────────────────────────────
const recColor = (r) => ({ "GO": "bg-emerald-500/20 text-emerald-400 border-emerald-500/30", "CONDITIONAL GO": "bg-amber-500/20 text-amber-400 border-amber-500/30", "SOFT PASS": "bg-orange-500/20 text-orange-400 border-orange-500/30", "PASS": "bg-red-500/20 text-red-400 border-red-500/30" }[r] || "bg-zinc-700/40 text-zinc-400 border-zinc-600");
const scoreColor = (pct) => pct >= 80 ? "#10b981" : pct >= 65 ? "#f59e0b" : pct >= 50 ? "#f97316" : "#ef4444";
const statusColor = (s) => ({ pending: "bg-zinc-700/50 text-zinc-400", analyzing: "bg-amber-500/20 text-amber-400", complete: "bg-emerald-500/20 text-emerald-400", killed: "bg-red-500/20 text-red-400" }[s] || "bg-zinc-700/50 text-zinc-400");
const fmt$ = (v) => v ? `$${Number(v).toLocaleString()}` : "—";
const fmtPct = (v) => v ? `${v}%` : "—";

// ── Shared UI ─────────────────────────────────────────────────────────────────
const inputCls = "w-full bg-zinc-950 border border-zinc-700 rounded-lg px-3 py-2 text-sm text-zinc-100 focus:outline-none focus:border-amber-500 placeholder:text-zinc-600 transition-colors";
const labelCls = "block text-xs text-zinc-500 mb-1 uppercase tracking-wider";
const selectCls = "w-full bg-zinc-950 border border-zinc-700 rounded-lg px-3 py-2 text-sm text-zinc-100 focus:outline-none focus:border-amber-500 transition-colors";
const sectionHd = "text-xs font-semibold text-amber-500 uppercase tracking-widest mb-3 mt-1";

function ScoreBar({ value, max, label }) {
  const pct = Math.round((value / max) * 100);
  return (
    <div className="mb-3">
      <div className="flex justify-between text-xs mb-1">
        <span className="text-zinc-400">{label}</span>
        <span style={{ fontFamily: "monospace", color: scoreColor(pct) }}>{value}/{max}</span>
      </div>
      <div className="h-1.5 rounded-full bg-zinc-800">
        <div className="h-full rounded-full transition-all duration-700" style={{ width: `${pct}%`, backgroundColor: scoreColor(pct) }} />
      </div>
    </div>
  );
}

function Toggle({ value, onChange }) {
  return (
    <button onClick={() => onChange(!value)} className={`relative w-11 h-6 rounded-full transition-colors shrink-0 ${value ? "bg-amber-500" : "bg-zinc-700"}`}>
      <span className="absolute top-0.5 w-5 h-5 bg-white rounded-full shadow transition-all" style={{ left: value ? "1.375rem" : "0.125rem" }} />
    </button>
  );
}

// ── File type detection ───────────────────────────────────────────────────────
const FILE_TYPES = {
  pdf: { exts: ["pdf"], mime: ["application/pdf"], label: "PDF", icon: "📄" },
  image: { exts: ["jpg","jpeg","png","gif","webp"], mime: ["image/jpeg","image/png","image/gif","image/webp"], label: "Image", icon: "🖼️" },
  excel: { exts: ["xlsx","xls","csv"], mime: ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","application/vnd.ms-excel","text/csv"], label: "Spreadsheet", icon: "📊" },
  word: { exts: ["docx","doc"], mime: ["application/vnd.openxmlformats-officedocument.wordprocessingml.document","application/msword"], label: "Word Doc", icon: "📝" },
};

function detectFileType(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  for (const [type, cfg] of Object.entries(FILE_TYPES)) {
    if (cfg.exts.includes(ext) || cfg.mime.includes(file.type)) return type;
  }
  return "unknown";
}

// ── Global fill handler — called by Claude via sendPrompt response ─────────────
// The artifact listens for a special JSON payload sent back through the chat
window.__storageIQFill = null; // set by ScorecardForm to receive filled data

// ── File → Claude content block ───────────────────────────────────────────────
async function fileToContentBlock(file) {
  const type = detectFileType(file);

  if (type === "pdf") {
    const b64 = await toBase64(file);
    return { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } };
  }

  if (type === "image") {
    const b64 = await toBase64(file);
    const mimeType = file.type || "image/jpeg";
    return { type: "image", source: { type: "base64", media_type: mimeType, data: b64 } };
  }

  if (type === "excel") {
    const text = await excelToText(file);
    return { type: "text", text: `[SPREADSHEET: ${file.name}]\n${text}` };
  }

  if (type === "word") {
    const text = await wordToText(file);
    return { type: "text", text: `[WORD DOC: ${file.name}]\n${text}` };
  }

  // fallback — try to read as text
  const text = await file.text().catch(() => "[Could not read file]");
  return { type: "text", text: `[FILE: ${file.name}]\n${text}` };
}

function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("Read failed"));
    r.readAsDataURL(file);
  });
}

async function excelToText(file) {
  try {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: "array" });
    const lines = [];
    for (const sheetName of wb.SheetNames) {
      const ws = wb.Sheets[sheetName];
      const csv = XLSX.utils.sheet_to_csv(ws, { blankrows: false });
      if (csv.trim()) lines.push(`=== Sheet: ${sheetName} ===\n${csv}`);
    }
    return lines.join("\n\n");
  } catch (e) {
    return "[Could not parse spreadsheet]";
  }
}

async function wordToText(file) {
  // Extract raw text from docx (zip of XML files)
  try {
    const { default: mammoth } = await import("mammoth");
    const ab = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: ab });
    return result.value || "[No text found in document]";
  } catch (e) {
    // Fallback: try reading as text
    try {
      return await file.text();
    } catch {
      return "[Could not parse Word document]";
    }
  }
}

// ── AI extraction prompt ──────────────────────────────────────────────────────
const EXTRACTION_PROMPT = `You are an expert self-storage underwriter. Your job is to extract every piece of data relevant to deal analysis from this document. Be aggressive — infer values when you can (e.g. calculate NOI margin from NOI/revenue, cap rate from NOI/price, economic occupancy from collected vs potential rent).

Return ONLY a valid JSON object with these exact keys (null only if truly not findable or inferable):
{
  "docType": "T12|RentRoll|OM|Email|Spreadsheet|Scan|Other",
  "propertyName": null,
  "askingPrice": null,
  "grossRevenue": null,
  "noiT12": null,
  "noiMargin": null,
  "physOcc": null,
  "econOcc": null,
  "capRate": null,
  "unitCount": null,
  "unitTypes": null,
  "climateControl": null,
  "delinquency": null,
  "mhi": null,
  "sqftPerCapita": null,
  "competitors": null,
  "sellerFinancing": null,
  "strengths": [],
  "redFlags": [],
  "nextSteps": [],
  "summary": "1-2 sentence summary of what this document contains and key findings",
  "executiveSummary": null,
  "rentableSqft": null,
  "lastSaleYear": null,
  "floodZone": null,
  "reitPresent": null,
  "crimeIndex": null
}

EXTRACTION RULES — read carefully:
- noiMargin: percentage as plain number (e.g. 52.3 for 52.3%). Calculate as (NOI / Gross Revenue) * 100 if both present.
- physOcc: physical occupancy as plain percentage number (e.g. 91.5). Look for "occupancy", "occ %", occupied units / total units * 100.
- econOcc: economic occupancy as plain percentage (collected rent / potential gross rent * 100). Also called "economic occupancy" or "collections rate".
- capRate: cap rate as plain percentage (e.g. 6.5). Calculate as (NOI / Asking Price) * 100 if both are present.
- askingPrice: total purchase price as plain number, no $ or commas. Look for "asking", "list price", "purchase price", "sale price".
- grossRevenue: total annual revenue as plain number. Look for "gross revenue", "gross income", "EGI", "total income", "potential gross income". If monthly, multiply by 12.
- noiT12: trailing 12-month NOI as plain number. Look for "NOI", "net operating income", "T12 NOI". If monthly, multiply by 12.
- unitCount: total number of storage units/spaces.
- unitTypes: COUNT of distinct unit size categories present (e.g. 5x5, 5x10, 10x10, 10x15, 10x20 = 5 types).
- climateControl: percentage of revenue or units that are climate-controlled. Look for "CC", "climate", "interior".
- delinquency: delinquency rate as plain percentage. Look for "delinquent", "past due", "collections issue".
- mhi: median household income as plain number (e.g. 62000). Look in market/demographic sections.
- sqftPerCapita: square feet of storage per capita in the market. Usually in OM market sections.
- competitors: number of competing self-storage facilities within 3 miles. Look in competitive analysis sections.
- sellerFinancing: true if seller financing is mentioned/offered, false if explicitly declined, null if not mentioned.
- strengths: array of 3-5 short strings highlighting positive deal attributes found in the document.
- redFlags: array of 3-5 short strings highlighting risks, concerns, or missing data.
- nextSteps: array of 3-5 recommended due diligence actions based on what this document shows.
- executiveSummary: 2-3 sentence investment thesis paragraph summarizing the deal — financials, value-add angle, key risks. Write as if presenting to an investor. Null if insufficient data.
- rentableSqft: total rentable square footage of the facility as a plain number. Look for "rentable sqft", "net rentable area", "total square feet".
- lastSaleYear: 4-digit year the property last sold. Look for "sale date", "acquired", "purchased", "deed date". Null if not found.
- floodZone: true if property is in a FEMA flood zone, false if explicitly clear, null if not mentioned.
- reitPresent: true if a REIT (Extra Space, Public Storage, CubeSmart, Life Storage, etc.) is mentioned as nearby competitor, false if explicitly no REITs, null if not mentioned.
- crimeIndex: numeric crime index value if mentioned (national average = 100). Null if not found.
- All dollar values: plain numbers only, no $, no commas (e.g. 2500000 not "$2.5M").
- Return ONLY the JSON object. No markdown, no explanation, no fences.`;

// ── Extract from single file via Claude API ───────────────────────────────────
async function extractFromFile(file, onLog) {
  const log = (msg) => { console.log(`[StorageIQ] ${msg}`); onLog?.(msg); };

  log(`Reading file: ${file.name} (${(file.size / 1024).toFixed(1)} KB, ${file.type || "unknown type"})`);

  let contentBlock;
  try {
    contentBlock = await fileToContentBlock(file);
    log(`File converted to content block: type=${contentBlock.type}`);
  } catch (e) {
    throw new Error(`Failed to read file: ${e.message}`);
  }

  const endpoint = import.meta.env.VITE_API_URL || window.__STORAGEIQ_API__ || "";
  if (!endpoint) throw new Error("No API endpoint set — add VITE_API_URL to your .env file");

  log(`Sending to proxy: ${endpoint}/api/extract`);
  let resp;
  try {
    resp = await fetch(`${endpoint}/api/extract`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        max_tokens: 1500,
        messages: [{
          role: "user",
          content: [contentBlock, { type: "text", text: EXTRACTION_PROMPT }]
        }]
      })
    });
  } catch (e) {
    throw new Error(`Network error — could not reach ${endpoint}: ${e.message}`);
  }

  log(`API response status: ${resp.status} ${resp.statusText}`);

  let data;
  try {
    data = await resp.json();
  } catch (e) {
    throw new Error(`API returned non-JSON response (status ${resp.status})`);
  }

  if (!resp.ok || data.error) {
    const errMsg = data.error?.message || data.error || `HTTP ${resp.status}`;
    const errType = data.error?.type || "";
    throw new Error(`API error [${errType}]: ${errMsg}`);
  }

  const text = data.content?.map(c => c.text || "").join("") || "";
  log(`Raw response length: ${text.length} chars`);

  if (!text.trim()) {
    throw new Error("API returned empty response — no content extracted");
  }

  const clean = text.replace(/```json|```/g, "").trim();

  let parsed;
  try {
    parsed = JSON.parse(clean);
    const foundFields = Object.entries(parsed).filter(([k, v]) => v !== null && v !== undefined && v !== "" && !(Array.isArray(v) && v.length === 0)).map(([k]) => k);
    log(`Extracted ${foundFields.length} fields: ${foundFields.join(", ")}`);
    return parsed;
  } catch (e) {
    log(`JSON parse failed. Raw text: ${clean.substring(0, 200)}`);
    throw new Error(`Could not parse AI response as JSON: ${e.message}`);
  }
}

// ── Merge extracted results ───────────────────────────────────────────────────
function mergeExtractions(results) {
  const merged = { extracted: {}, docsSummary: [] };
  const SCALAR_KEYS = ["askingPrice","grossRevenue","noiT12","noiMargin","physOcc","econOcc","capRate","unitCount","unitTypes","climateControl","delinquency","mhi","sqftPerCapita","competitors","sellerFinancing","propertyName","executiveSummary","rentableSqft","lastSaleYear","floodZone","reitPresent","crimeIndex"];

  for (const { parsed, file } of results) {
    merged.docsSummary.push({ name: file.name, type: parsed.docType, summary: parsed.summary });
    for (const k of SCALAR_KEYS) {
      if (parsed[k] !== null && parsed[k] !== undefined) merged.extracted[k] = String(parsed[k]);
    }
    for (const k of ["strengths","redFlags","nextSteps"]) {
      if (!merged.extracted[k]) merged.extracted[k] = [];
      if (parsed[k]?.length) merged.extracted[k] = [...new Set([...merged.extracted[k], ...parsed[k]])];
    }
  }

  // Convert arrays to newline strings for textareas
  for (const k of ["strengths","redFlags","nextSteps"]) {
    if (Array.isArray(merged.extracted[k])) merged.extracted[k] = merged.extracted[k].join("\n");
  }

  return merged;
}

// ── Doc Upload Zone ───────────────────────────────────────────────────────────
const ACCEPTED = ".pdf,.doc,.docx,.xls,.xlsx,.csv,.jpg,.jpeg,.png,.gif,.webp";
const TYPE_ICONS = { pdf: "📄", image: "🖼️", excel: "📊", word: "📝", unknown: "📎" };

// ── Doc Upload Zone (standalone — real file upload + Vercel API) ──────────────
function DocUploadZone({ onExtracted, existingDocs = [], dealName = "" }) {
  const [docs, setDocs] = useState(existingDocs);
  const [status, setStatus] = useState("idle");
  const [progress, setProgress] = useState({ current: 0, total: 0, name: "" });
  const [errors, setErrors] = useState([]);
  const [fileResults, setFileResults] = useState([]);
  const [dragOver, setDragOver] = useState(false);
  const [logs, setLogs] = useState([]);
  const [showLogs, setShowLogs] = useState(false);
  const inputRef = useRef();
  const addLog = (msg) => setLogs(prev => [...prev, `${new Date().toLocaleTimeString()} — ${msg}`]);

  const processFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList);
    if (!files.length) return;
    setStatus("uploading");
    setErrors([]);
    setLogs([]);
    setFileResults(files.map(f => ({ name: f.name, status: "pending", error: null, fieldCount: 0 })));
    const results = [];
    const errs = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      setProgress({ current: i + 1, total: files.length, name: file.name });
      setFileResults(prev => prev.map((r, idx) => idx === i ? { ...r, status: "reading" } : r));
      try {
        const parsed = await extractFromFile(file, addLog);
        const fieldCount = Object.values(parsed).filter(v => v !== null && v !== undefined && v !== "" && !(Array.isArray(v) && v.length === 0)).length;
        results.push({ parsed, file });
        setFileResults(prev => prev.map((r, idx) => idx === i ? { ...r, status: "done", fieldCount } : r));
      } catch (e) {
        errs.push({ name: file.name, msg: e.message });
        addLog(`ERROR on ${file.name}: ${e.message}`);
        setFileResults(prev => prev.map((r, idx) => idx === i ? { ...r, status: "error", error: e.message } : r));
      }
    }
    if (results.length) {
      const merged = mergeExtractions(results);
      setDocs(prev => [...prev, ...merged.docsSummary]);
      onExtracted(merged.extracted, merged.docsSummary);
      setStatus("done");
    } else {
      setStatus("error");
    }
    setErrors(errs);
  }, [onExtracted]);

  const onDrop = (e) => { e.preventDefault(); setDragOver(false); processFiles(e.dataTransfer.files); };
  const uploading = status === "uploading";

  const FileStatus = ({ s }) => ({ pending: <span className="text-zinc-600">○</span>, reading: <span className="text-amber-400 animate-pulse">◉</span>, done: <span className="text-emerald-400">✓</span>, error: <span className="text-red-400">✕</span> }[s] || null);

  return (
    <div className="space-y-3">
      <div
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={onDrop}
        onClick={() => !uploading && inputRef.current?.click()}
        className={`border-2 border-dashed rounded-xl p-6 text-center transition-all
          ${uploading ? "pointer-events-none opacity-70 border-zinc-700 bg-zinc-950" :
            dragOver ? "border-amber-500 bg-amber-500/5 cursor-copy" :
            "border-zinc-700 hover:border-zinc-500 bg-zinc-950 cursor-pointer"}`}
      >
        <input ref={inputRef} type="file" accept={ACCEPTED} multiple className="hidden" onChange={e => processFiles(e.target.files)} />
        {uploading ? (
          <div>
            <div className="w-6 h-6 border-2 border-amber-500 border-t-transparent rounded-full animate-spin mx-auto mb-3" />
            <p className="text-sm text-amber-400 font-medium">{progress.name}</p>
            <p className="text-xs text-zinc-500 mt-1">File {progress.current} of {progress.total} — AI extracting…</p>
          </div>
        ) : (
          <div>
            <div className="flex justify-center gap-1 text-2xl mb-2">{Object.values(TYPE_ICONS).map((icon, i) => <span key={i}>{icon}</span>)}</div>
            <p className="text-sm text-zinc-300 font-medium">Drop files here or click to upload</p>
            <p className="text-xs text-zinc-500 mt-1">PDF · Word · Excel/CSV · Images · Screenshots</p>
            <p className="text-xs text-zinc-600 mt-1">AI reads all formats and auto-fills your scorecard</p>
          </div>
        )}
      </div>
      {uploading && (
        <div className="h-1 bg-zinc-800 rounded-full overflow-hidden">
          <div className="h-full bg-amber-500 rounded-full transition-all duration-300" style={{ width: `${(progress.current / progress.total) * 100}%` }} />
        </div>
      )}
      {fileResults.length > 0 && (
        <div className="space-y-1.5">
          {fileResults.map((r, i) => (
            <div key={i} className={`flex items-start gap-2 rounded-lg px-3 py-2 border text-xs ${r.status === "error" ? "bg-red-500/5 border-red-500/20" : r.status === "done" ? "bg-emerald-500/5 border-emerald-500/20" : "bg-zinc-900 border-zinc-800"}`}>
              <span className="mt-0.5 shrink-0"><FileStatus s={r.status} /></span>
              <div className="min-w-0 flex-1">
                <p className="font-medium text-zinc-300 truncate">{r.name}</p>
                {r.status === "done" && <p className="text-zinc-500">{r.fieldCount} fields extracted</p>}
                {r.status === "reading" && <p className="text-amber-400">Reading with AI…</p>}
                {r.status === "error" && <p className="text-red-400 break-all">{r.error}</p>}
              </div>
            </div>
          ))}
        </div>
      )}
      {status === "done" && errors.length === 0 && (
        <div className="bg-emerald-500/10 border border-emerald-500/20 rounded-lg px-3 py-2 text-xs text-emerald-400">
          ✓ All {fileResults.length} file{fileResults.length !== 1 ? "s" : ""} processed — amber fields auto-filled below
        </div>
      )}
      {status === "error" && (
        <div className="bg-red-500/10 border border-red-500/20 rounded-lg px-3 py-2 text-xs text-red-400">
          ✕ All files failed — check your API connection in Settings
        </div>
      )}
      {logs.length > 0 && (
        <div>
          <button onClick={() => setShowLogs(v => !v)} className="text-xs text-zinc-600 hover:text-zinc-400 transition-colors">
            {showLogs ? "▾ Hide" : "▸ Show"} debug log ({logs.length} entries)
          </button>
          {showLogs && (
            <div className="mt-2 bg-zinc-950 border border-zinc-800 rounded-lg p-3 space-y-1 max-h-48 overflow-y-auto">
              {logs.map((l, i) => <p key={i} className={`text-xs font-mono ${l.includes("ERROR") ? "text-red-400" : "text-zinc-500"}`}>{l}</p>)}
            </div>
          )}
        </div>
      )}
      {docs.length > 0 && (
        <div className="space-y-1.5">
          {docs.map((d, i) => (
            <div key={i} className="flex items-start gap-2 bg-zinc-900 border border-zinc-800 rounded-lg px-3 py-2">
              <span className="text-base shrink-0 mt-0.5">{TYPE_ICONS[detectFileType({ name: d.name || "", type: "" })] || "📎"}</span>
              <div className="min-w-0">
                <p className="text-xs font-medium text-zinc-300 truncate">{d.name}</p>
                <p className="text-xs text-zinc-500">{d.type}{d.summary ? ` — ${d.summary}` : ""}</p>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ── New Deal Form ────────────────────────────────────────────────────────────
function NewDealForm({ onSave, onCancel }) {
  const [f, setF] = useState({ name: "", address: "", city: "", state: "TX", zip: "" });
  const set = (k) => (e) => setF(prev => ({ ...prev, [k]: e.target.value }));

  return (
    <div className="fixed inset-0 bg-black/70 backdrop-blur-sm flex items-center justify-center z-50 p-4">
      <div className="bg-zinc-900 border border-zinc-700 rounded-2xl w-full max-w-md p-6">
        <h2 className="text-lg font-semibold text-zinc-100 mb-1">New Deal</h2>
        <p className="text-xs text-zinc-500 mb-5">Enter the property details — you can upload docs and auto-fill the scorecard after creating the deal.</p>
        <div className="space-y-4">
          <div><label className={labelCls}>Facility Name *</label><input className={inputCls} value={f.name} onChange={set("name")} placeholder="West Mountain Storage" autoFocus /></div>
          <div><label className={labelCls}>Street Address</label><input className={inputCls} value={f.address} onChange={set("address")} placeholder="8598 US Hwy 271 S" /></div>
          <div className="grid grid-cols-3 gap-2">
            <div className="col-span-1"><label className={labelCls}>City</label><input className={inputCls} value={f.city} onChange={set("city")} placeholder="Gilmer" /></div>
            <div><label className={labelCls}>State</label><input className={inputCls} value={f.state} onChange={set("state")} placeholder="TX" /></div>
            <div><label className={labelCls}>ZIP</label><input className={inputCls} value={f.zip} onChange={set("zip")} placeholder="75645" /></div>
          </div>
        </div>
        <div className="flex gap-3 mt-6">
          <button onClick={onCancel} className="flex-1 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 rounded-lg py-2.5 text-sm transition-colors">Cancel</button>
          <button onClick={() => f.name && onSave(f)} disabled={!f.name}
            className="flex-1 bg-amber-500 hover:bg-amber-400 text-zinc-950 font-semibold rounded-lg py-2.5 text-sm transition-colors disabled:opacity-40">
            Create Deal →
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Kill Switch Panel ─────────────────────────────────────────────────────────

function KillSwitchPanel({ deal, onSave }) {
  const [ks, setKs] = useState(deal.killSwitch || { reitPresent: false, reitDetail: "", mhi: "", mhiDetail: "", crimeIndex: "", crimeDetail: "", floodZone: false, floodDetail: "" });
  const set = (k) => (e) => setKs(prev => ({ ...prev, [k]: e.target.value }));
  const mhiPass = (parseFloat(ks.mhi) || 0) >= 50000;
  const crimePass = (parseFloat(ks.crimeIndex) || 999) <= 100;
  const overallPass = !ks.reitPresent && mhiPass && crimePass && !ks.floodZone;

  return (
    <div className="space-y-5">
      <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
        <div className="flex items-center justify-between mb-3">
          <div><p className="text-sm font-medium text-zinc-200">REIT within 5 miles</p><p className="text-xs text-zinc-500 mt-0.5">Large REIT presence pressures rates</p></div>
          <div className="flex items-center gap-2">
            <span className={`text-xs ${ks.reitPresent ? "text-red-400" : "text-zinc-500"}`}>{ks.reitPresent ? "PRESENT" : "NONE"}</span>
            <Toggle value={ks.reitPresent} onChange={(v) => setKs(p => ({ ...p, reitPresent: v }))} />
          </div>
        </div>
        <div><label className={labelCls}>Source / Notes</label><input className={inputCls + " mt-1"} value={ks.reitDetail} onChange={set("reitDetail")} placeholder="e.g. Extra Space 2.1 mi NW" /></div>
      </div>

      <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
        <div className="flex items-center justify-between mb-3">
          <div><p className="text-sm font-medium text-zinc-200">Median Household Income ≥ $50k</p><p className="text-xs text-zinc-500 mt-0.5">Minimum income threshold</p></div>
          <span className={`text-xs font-mono px-2 py-0.5 rounded-full border ${mhiPass ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : "bg-red-500/20 text-red-400 border-red-500/30"}`}>{mhiPass ? "PASS" : "FAIL"}</span>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div><label className={labelCls}>MHI Value ($)</label><input className={inputCls + " mt-1"} value={ks.mhi} onChange={set("mhi")} placeholder="62000" type="number" /></div>
          <div><label className={labelCls}>Source</label><input className={inputCls + " mt-1"} value={ks.mhiDetail} onChange={set("mhiDetail")} placeholder="Census ACS 2023" /></div>
        </div>
      </div>

      <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
        <div className="flex items-center justify-between mb-3">
          <div><p className="text-sm font-medium text-zinc-200">Crime Index ≤ 100</p><p className="text-xs text-zinc-500 mt-0.5">National average = 100</p></div>
          <span className={`text-xs font-mono px-2 py-0.5 rounded-full border ${crimePass ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : "bg-red-500/20 text-red-400 border-red-500/30"}`}>{crimePass ? "PASS" : "FAIL"}</span>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div><label className={labelCls}>Crime Index</label><input className={inputCls + " mt-1"} value={ks.crimeIndex} onChange={set("crimeIndex")} placeholder="82" type="number" /></div>
          <div><label className={labelCls}>Source</label><input className={inputCls + " mt-1"} value={ks.crimeDetail} onChange={set("crimeDetail")} placeholder="NeighborhoodScout" /></div>
        </div>
      </div>

      {/* Flood Zone */}
      <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
        <div className="flex items-center justify-between mb-3">
          <div>
            <p className="text-sm font-medium text-zinc-200">Flood Zone</p>
            <p className="text-xs text-zinc-500 mt-0.5">Property must not be in a FEMA flood zone</p>
          </div>
          <div className="flex items-center gap-2">
            <span className={`text-xs ${ks.floodZone ? "text-red-400" : "text-zinc-500"}`}>{ks.floodZone ? "IN FLOOD ZONE" : "CLEAR"}</span>
            <button onClick={() => setKs(p => ({ ...p, floodZone: !p.floodZone }))}
              className={`relative w-11 h-6 rounded-full transition-colors shrink-0 ${ks.floodZone ? "bg-red-500" : "bg-zinc-700"}`}>
              <span className="absolute top-0.5 w-5 h-5 bg-white rounded-full shadow transition-all" style={{ left: ks.floodZone ? "1.375rem" : "0.125rem" }} />
            </button>
          </div>
        </div>
        <div>
          <label className={labelCls}>Source / Notes</label>
          <input className={inputCls + " mt-1"} value={ks.floodDetail} onChange={e => setKs(p => ({ ...p, floodDetail: e.target.value }))} placeholder="e.g. FEMA Map Service Center — Zone X" />
        </div>
      </div>

      <div className={`rounded-xl p-4 text-center border ${overallPass ? "bg-emerald-500/10 border-emerald-500/30" : "bg-red-500/10 border-red-500/30"}`}>
        <p className={`text-lg font-bold ${overallPass ? "text-emerald-400" : "text-red-400"}`}>{overallPass ? "✓ DEAL PASSES KILL SWITCH" : "✕ DEAL KILLED"}</p>
        <p className="text-xs text-zinc-500 mt-1">{overallPass ? "Proceed to Scorecard" : "Do not underwrite further"}</p>
      </div>

      <button onClick={() => onSave({ ...ks, mhiPass, crimePass, overallPass })} className="w-full bg-amber-500 hover:bg-amber-400 text-zinc-950 font-semibold rounded-lg py-2.5 text-sm transition-colors">Save Kill Switch</button>
    </div>
  );
}

// ── Scorecard Form ────────────────────────────────────────────────────────────
const EMPTY_SC = { physOcc: "", econOcc: "", noiMargin: "", capRate: "", askingPrice: "", noiT12: "", grossRev: "", mhi: "", popGrowth: "", sqftPerCapita: "", competitors: "", streetRates: "", unitTypes: "", climateControl: "", delinquency: "", sellerFinancing: false, valueAdd: "", strengths: "", redFlags: "", nextSteps: "", executiveSummary: "" };

function EditPanel({ deal, onSave }) {
  const sc = deal.scorecard;
  const [f, setF] = useState(sc ? { ...EMPTY_SC, ...sc } : EMPTY_SC);
  const set = (k) => (e) => setF(prev => ({ ...prev, [k]: e.target?.value ?? e }));
  const score = useMemo(() => scoreCard(f), [f]);

  return (
    <div className="space-y-6">
      {/* Live score */}
      <div className="bg-zinc-950 border border-zinc-700 rounded-xl p-4 sticky top-0 z-10">
        <div className="flex items-center justify-between mb-3">
          <span className="text-xs text-zinc-500 uppercase tracking-wider">Live Score</span>
          <span className={`text-xs font-semibold px-2 py-0.5 rounded-full border ${recColor(score.rec)}`}>{score.rec}</span>
        </div>
        <div className="flex items-baseline gap-1 mb-3">
          <span className="text-3xl font-bold text-zinc-100" style={{ fontFamily: "monospace" }}>{score.total}</span>
          <span className="text-zinc-500 text-sm">/100</span>
        </div>
        <ScoreBar value={score.fin} max={35} label="Financial" />
        <ScoreBar value={score.mkt} max={25} label="Market" />
        <ScoreBar value={score.comp} max={20} label="Competitive" />
        <ScoreBar value={score.ops} max={12} label="Operational" />
        <ScoreBar value={score.deal} max={8} label="Deal Structure" />
      </div>

      {/* Financial */}
      <div>
        <p className={sectionHd}>Financial — 35 pts</p>
        <div className="grid grid-cols-2 gap-3">
          {[["physOcc","Physical Occ %","90"],["econOcc","Economic Occ %","85"],["noiMargin","NOI Margin %","52"],["capRate","Cap Rate %","6.5"],["askingPrice","Asking Price","2500000"],["noiT12","NOI T12","180000"],["grossRev","Gross Revenue","340000"]].map(([k,l,ph]) => (
            <div key={k}><label className={labelCls}>{l}</label><input className={inputCls} value={f[k]} onChange={set(k)} placeholder={ph} type="number" /></div>
          ))}
        </div>
      </div>

      {/* Market */}
      <div>
        <p className={sectionHd}>Market — 25 pts</p>
        <div className="grid grid-cols-2 gap-3">
          <div><label className={labelCls}>Median HHI ($)</label><input className={inputCls} value={f.mhi} onChange={set("mhi")} placeholder="62000" type="number" /></div>
          <div><label className={labelCls}>Sq Ft / Capita</label><input className={inputCls} value={f.sqftPerCapita} onChange={set("sqftPerCapita")} placeholder="7.2" type="number" /></div>
          <div className="col-span-2">
            <label className={labelCls}>Population Growth</label>
            <select className={selectCls} value={f.popGrowth} onChange={set("popGrowth")}>
              <option value="">Select…</option>
              <option value="growing">Growing (+7 pts)</option>
              <option value="flat">Flat (+4 pts)</option>
              <option value="declining">Declining (0 pts)</option>
            </select>
          </div>
        </div>
      </div>

      {/* Competitive */}
      <div>
        <p className={sectionHd}>Competitive — 20 pts</p>
        <div className="grid grid-cols-2 gap-3">
          <div><label className={labelCls}>Competitors within 3 mi</label><input className={inputCls} value={f.competitors} onChange={set("competitors")} placeholder="3" type="number" /></div>
          <div>
            <label className={labelCls}>Street Rates vs Market</label>
            <select className={selectCls} value={f.streetRates} onChange={set("streetRates")}>
              <option value="">Select…</option>
              <option value="at-above">At / Above (+10 pts)</option>
              <option value="slightly-below">5–10% Below (+6 pts)</option>
              <option value="well-below">10%+ Below (+2 pts)</option>
            </select>
          </div>
        </div>
      </div>

      {/* Operational */}
      <div>
        <p className={sectionHd}>Operational — 12 pts</p>
        <div className="grid grid-cols-3 gap-3">
          {[["unitTypes","Unit Types","4"],["climateControl","Climate Control %","35"],["delinquency","Delinquency %","2.5"]].map(([k,l,ph]) => (
            <div key={k}><label className={labelCls}>{l}</label><input className={inputCls} value={f[k]} onChange={set(k)} placeholder={ph} type="number" /></div>
          ))}
        </div>
      </div>

      {/* Deal Structure */}
      <div>
        <p className={sectionHd}>Deal Structure — 8 pts</p>
        <div className="flex items-center gap-3 mb-3">
          <Toggle value={!!f.sellerFinancing} onChange={(v) => setF(p => ({ ...p, sellerFinancing: v }))} />
          <span className="text-sm text-zinc-300">Seller Financing Available (+4 pts)</span>
        </div>
        <div><label className={labelCls}>Value-Add Opportunity (+4 pts if filled)</label><input className={inputCls} value={f.valueAdd} onChange={set("valueAdd")} placeholder="Rezone 2 acres for RV storage, add 40 units" /></div>
      </div>

      {/* Narratives */}
      <div>
        <p className={sectionHd}>Narratives</p>
        <div className="space-y-3">
          {[["executiveSummary","Executive Summary","Investment thesis…"],["strengths","Key Strengths","High occupancy, seller motivated"],["redFlags","Red Flags","Delinquency trending up"],["nextSteps","Next Steps","Order phase I, request rent rolls"]].map(([k,l,ph]) => (
            <div key={k}>
              <label className={labelCls}>{l}</label>
              <textarea className={`${inputCls} resize-y`} rows={3} value={f[k] || ""} onChange={set(k)} placeholder={ph} style={{minHeight:"60px"}} onInput={e => { e.target.style.height = "auto"; e.target.style.height = e.target.scrollHeight + "px"; }} />
            </div>
          ))}
        </div>
      </div>

      <button onClick={() => onSave({ ...f, score, docs: sc?.docs || [] })} className="w-full bg-amber-500 hover:bg-amber-400 text-zinc-950 font-bold rounded-lg py-3 text-sm transition-colors">
        Save — {score.total}/100 {score.rec}
      </button>
    </div>
  );
}

// ── Offer Multiple Calculator ─────────────────────────────────────────────────
function OfferMultipleCalc({ scorecard }) {
  const [monthlyRev, setMonthlyRev] = useState(() => {
    if (scorecard?.grossRev) return String(Math.round(parseFloat(scorecard.grossRev) / 12));
    return "";
  });
  const [askingPrice, setAskingPrice] = useState(() => scorecard?.askingPrice || "");

  const monthly = parseFloat(monthlyRev) || 0;
  const asking  = parseFloat(askingPrice) || 0;

  const multiples = [
    { label: "80x", value: 80, color: "#10b981", note: "Conservative" },
    { label: "90x", value: 90, color: "#f59e0b", note: "Target" },
    { label: "100x", value: 100, color: "#f97316", note: "Aggressive" },
  ];

  const impliedMultiple = monthly > 0 && asking > 0 ? asking / monthly : null;

  const askColor = () => {
    if (!impliedMultiple) return "text-zinc-400";
    if (impliedMultiple <= 90) return "text-emerald-400";
    if (impliedMultiple <= 110) return "text-amber-400";
    if (impliedMultiple <= 130) return "text-orange-400";
    return "text-red-400";
  };

  const askLabel = () => {
    if (!impliedMultiple) return "";
    if (impliedMultiple <= 80) return "Strong buy range";
    if (impliedMultiple <= 90) return "Within target range";
    if (impliedMultiple <= 100) return "At top of range";
    if (impliedMultiple <= 115) return "Slightly above range";
    if (impliedMultiple <= 130) return "Above range — negotiate down";
    return "Well above range — unlikely to pencil";
  };

  return (
    <div className="bg-zinc-950 border border-zinc-700 rounded-xl p-5">
      <p className="text-xs font-semibold text-amber-500 uppercase tracking-widest mb-4">Offer Multiple Calculator</p>

      {/* Inputs */}
      <div className="grid grid-cols-2 gap-3 mb-5">
        <div>
          <label className={labelCls}>Monthly Revenue ($)</label>
          <input
            className={inputCls}
            value={monthlyRev}
            onChange={e => setMonthlyRev(e.target.value)}
            placeholder="14,942"
            type="number"
          />
        </div>
        <div>
          <label className={labelCls}>Asking Price ($)</label>
          <input
            className={inputCls}
            value={askingPrice}
            onChange={e => setAskingPrice(e.target.value)}
            placeholder="2,585,000"
            type="number"
          />
        </div>
      </div>

      {/* Multiples table */}
      {monthly > 0 && (
        <div className="space-y-2 mb-4">
          {multiples.map(({ label, value, color, note }) => {
            const offer = monthly * value;
            const isAskAbove = asking > 0 && asking > offer;
            const pct = asking > 0 ? Math.min((offer / asking) * 100, 100) : 100;
            return (
              <div key={label} className="bg-zinc-900 border border-zinc-800 rounded-lg p-3">
                <div className="flex items-center justify-between mb-2">
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-zinc-400 w-8">{label}</span>
                    <span className="text-xs text-zinc-600">{note}</span>
                  </div>
                  <div className="flex items-center gap-3">
                    {asking > 0 && (
                      <span className={`text-xs ${isAskAbove ? "text-red-400" : "text-emerald-400"}`}>
                        {isAskAbove ? `Ask $${Math.round((asking - offer) / 1000)}k above` : `Ask $${Math.round((offer - asking) / 1000)}k below`}
                      </span>
                    )}
                    <span className="text-sm font-bold text-zinc-100" style={{ fontFamily: "monospace", color }}>
                      ${offer.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                    </span>
                  </div>
                </div>
                <div className="h-1.5 rounded-full bg-zinc-800">
                  <div className="h-full rounded-full transition-all duration-500" style={{ width: `${pct}%`, backgroundColor: color }} />
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Asking price analysis */}
      {impliedMultiple && (
        <div className={`rounded-lg p-3 border ${impliedMultiple <= 100 ? "bg-emerald-500/10 border-emerald-500/20" : impliedMultiple <= 115 ? "bg-amber-500/10 border-amber-500/20" : "bg-red-500/10 border-red-500/20"}`}>
          <div className="flex items-center justify-between">
            <div>
              <p className="text-xs text-zinc-500">Asking Price Implied Multiple</p>
              <p className={`text-xs mt-0.5 ${askColor()}`}>{askLabel()}</p>
            </div>
            <span className={`text-xl font-bold ${askColor()}`} style={{ fontFamily: "monospace" }}>
              {impliedMultiple.toFixed(1)}x
            </span>
          </div>
        </div>
      )}

      {!monthly && (
        <p className="text-xs text-zinc-600 text-center py-2">Enter monthly revenue to see offer ranges</p>
      )}
    </div>
  );
}

// ── Report View (Cactus-style) ────────────────────────────────────────────────
function ReportView({ deal, onUpdateDeal }) {
  const sc = deal.scorecard;
  const splitLines = (txt) => txt ? txt.split("\n").filter(Boolean) : [];

  const handleExtracted = useCallback((extracted, docsSummary) => {
    // Update scorecard fields
    const base = sc ? { ...EMPTY_SC, ...sc } : { ...EMPTY_SC };
    const FIELD_MAP = { grossRevenue: "grossRev" };
    for (const [k, v] of Object.entries(extracted)) {
      const key = FIELD_MAP[k] || k;
      if (key in base) base[key] = k === "sellerFinancing" ? v === "true" : v;
    }
    const newScore = scoreCard(base);
    const newDocs = [...(sc?.docs || []), ...docsSummary];

    // Update kill switch fields
    const ks = { ...(deal.killSwitch || { reitPresent: false, reitDetail: "", mhi: "", mhiDetail: "", crimeIndex: "", crimeDetail: "", floodZone: false, floodDetail: "" }) };
    if (extracted.mhi) ks.mhi = extracted.mhi;
    if (extracted.crimeIndex) ks.crimeIndex = extracted.crimeIndex;
    if (extracted.floodZone !== null && extracted.floodZone !== undefined) ks.floodZone = extracted.floodZone === "true" || extracted.floodZone === true;
    if (extracted.reitPresent !== null && extracted.reitPresent !== undefined) ks.reitPresent = extracted.reitPresent === "true" || extracted.reitPresent === true;
    const mhiPass = (parseFloat(ks.mhi) || 0) >= 50000;
    const crimePass = (parseFloat(ks.crimeIndex) || 999) <= 100;
    const ksWithStatus = { ...ks, mhiPass, crimePass, overallPass: !ks.reitPresent && mhiPass && crimePass && !ks.floodZone };

    // Update buy box fields
    const bb = { ...(deal.buyBox || {}) };
    if (extracted.mhi) bb.mhi = extracted.mhi;
    if (extracted.sqftPerCapita) bb.sqftPerCapita = extracted.sqftPerCapita;
    if (extracted.rentableSqft) bb.sqft = extracted.rentableSqft;
    if (extracted.lastSaleYear) bb.lastSaleYear = extracted.lastSaleYear;
    if (extracted.floodZone !== null && extracted.floodZone !== undefined) bb.floodZone = extracted.floodZone === "true" || extracted.floodZone === true;
    if (extracted.strengths) bb.marketNotes = Array.isArray(extracted.strengths) ? extracted.strengths.join("\n") : extracted.strengths;
    if (extracted.summary) bb.notes = extracted.summary;

    const updatedName = extracted.propertyName || deal.name;
    onUpdateDeal({ ...deal, name: updatedName, scorecard: { ...base, score: newScore, docs: newDocs }, killSwitch: ksWithStatus, buyBox: bb, status: "complete" });
  }, [deal, sc, onUpdateDeal]);

  if (!sc) {
    return (
      <div className="space-y-6">
        <div className="text-center py-10">
          <p className="text-5xl mb-4">📂</p>
          <p className="font-semibold text-zinc-300 text-lg mb-1">Upload documents to generate your report</p>
          <p className="text-zinc-500 text-sm">AI reads PDFs, Excel, Word, and images — auto-fills all metrics</p>
        </div>
        <DocUploadZone onExtracted={handleExtracted} dealName={deal.name} />
      </div>
    );
  }

  const score = sc.score;
  const strengths = splitLines(sc.strengths);
  const redFlags = splitLines(sc.redFlags);
  const nextSteps = splitLines(sc.nextSteps);

  const recBg = { "GO": "bg-emerald-500", "CONDITIONAL GO": "bg-amber-500", "SOFT PASS": "bg-orange-500", "PASS": "bg-red-500" }[score.rec] || "bg-zinc-600";

  return (
    <div className="space-y-5">
      {/* Header */}
      <div className="bg-gradient-to-br from-zinc-900 to-zinc-950 border border-zinc-700 rounded-2xl p-6">
        <div className="flex items-start justify-between mb-5">
          <div>
            <h2 className="text-2xl font-bold text-zinc-100">{deal.name}</h2>
            <p className="text-sm text-zinc-500 mt-1">{[deal.address, deal.city, deal.state].filter(Boolean).join(", ")}</p>
          </div>
          <div className="text-right shrink-0 ml-4">
            <div className="text-4xl font-bold text-zinc-100" style={{fontFamily:"monospace"}}>{score.total}<span className="text-lg text-zinc-500">/100</span></div>
            <span className={`inline-block mt-1 text-xs font-bold px-3 py-1 rounded-full text-white ${recBg}`}>{score.rec}</span>
          </div>
        </div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {[["Deal Score",`${score.total}/100`],["Recommendation",score.rec],["Physical Occ",sc.physOcc?`${sc.physOcc}%`:"—"],["Cap Rate",sc.capRate?`${sc.capRate}%`:"—"]].map(([l,v]) => (
            <div key={l} className="bg-zinc-950/60 border border-zinc-800 rounded-xl p-3">
              <p className="text-xs text-zinc-500 uppercase tracking-wider mb-1">{l}</p>
              <p className="text-sm font-bold text-zinc-200" style={{fontFamily:"monospace"}}>{v}</p>
            </div>
          ))}
        </div>
      </div>

      {/* Executive Summary */}
      {sc.executiveSummary && (
        <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-5">
          <p className="text-xs text-zinc-500 uppercase tracking-wider mb-3">Executive Summary</p>
          <p className="text-sm text-zinc-300 leading-relaxed">{sc.executiveSummary}</p>
        </div>
      )}

      {/* Score Breakdown */}
      <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-5">
        <p className="text-xs text-zinc-500 uppercase tracking-wider mb-4">Score Breakdown</p>
        <ScoreBar value={score.fin} max={35} label="Financial" />
        <ScoreBar value={score.mkt} max={25} label="Market" />
        <ScoreBar value={score.comp} max={20} label="Competitive" />
        <ScoreBar value={score.ops} max={12} label="Operational" />
        <ScoreBar value={score.deal} max={8} label="Deal Structure" />
      </div>

      {/* Key Metrics */}
      <div className="grid grid-cols-3 gap-3">
        {[["Asking Price",fmt$(sc.askingPrice)],["NOI T12",fmt$(sc.noiT12)],["Gross Revenue",fmt$(sc.grossRev)],["NOI Margin",fmtPct(sc.noiMargin)],["Econ. Occ",fmtPct(sc.econOcc)],["Units",sc.unitCount||"—"]].map(([l,v]) => (
          <div key={l} className="bg-zinc-900 border border-zinc-800 rounded-xl p-3 text-center">
            <p className="text-xs text-zinc-500 mb-1">{l}</p>
            <p className="text-sm font-semibold text-zinc-200" style={{fontFamily:"monospace"}}>{v}</p>
          </div>
        ))}
      </div>

      {/* Risks & Opportunities */}
      {(redFlags.length > 0 || strengths.length > 0) && (
        <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-5">
          <p className="text-xs text-zinc-500 uppercase tracking-wider mb-4">Risks & Opportunities</p>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
            {redFlags.length > 0 && (
              <div>
                <p className="text-xs font-semibold text-red-400 uppercase tracking-wider mb-3">⚠ Top Risks</p>
                <div className="space-y-2">
                  {redFlags.map((s,i) => (
                    <div key={i} className="bg-red-500/5 border-l-2 border-red-500/40 rounded-r-xl px-4 py-3">
                      <p className="text-sm text-zinc-300">{s}</p>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {strengths.length > 0 && (
              <div>
                <p className="text-xs font-semibold text-emerald-400 uppercase tracking-wider mb-3">✦ Opportunities</p>
                <div className="space-y-2">
                  {strengths.map((s,i) => (
                    <div key={i} className="bg-emerald-500/5 border-l-2 border-emerald-500/40 rounded-r-xl px-4 py-3">
                      <p className="text-sm text-zinc-300">{s}</p>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* Next Steps */}
      {nextSteps.length > 0 && (
        <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-5">
          <p className="text-xs text-zinc-500 uppercase tracking-wider mb-4">Strategic Next Steps</p>
          <div className="space-y-3">
            {nextSteps.map((s,i) => (
              <div key={i} className="flex gap-3 items-start">
                <div className="w-6 h-6 rounded-full bg-emerald-500 flex items-center justify-center text-xs font-bold text-white shrink-0 mt-0.5">{i+1}</div>
                <p className="text-sm text-zinc-300 leading-relaxed">{s}</p>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Kill Switch */}
      {deal.killSwitch && (
        <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-5">
          <p className="text-xs text-zinc-500 uppercase tracking-wider mb-3">Kill Switch</p>
          <div className={`rounded-xl p-3 border text-center mb-3 ${deal.killSwitch.overallPass ? "bg-emerald-500/10 border-emerald-500/30" : "bg-red-500/10 border-red-500/30"}`}>
            <p className={`text-sm font-bold ${deal.killSwitch.overallPass ? "text-emerald-400" : "text-red-400"}`}>{deal.killSwitch.overallPass ? "✓ Passes Kill Switch" : "✕ Deal Killed"}</p>
          </div>
          <div className="grid grid-cols-2 gap-2 text-xs">
            {[["REIT",!deal.killSwitch.reitPresent],["MHI ≥ $50k",deal.killSwitch.mhiPass],["Crime ≤ 100",deal.killSwitch.crimePass],["Flood Zone",!deal.killSwitch.floodZone]].map(([l,pass]) => (
              <div key={l} className="flex items-center justify-between bg-zinc-950 rounded-lg px-3 py-2">
                <span className="text-zinc-400">{l}</span>
                <span className={pass ? "text-emerald-400" : "text-red-400"}>{pass ? "PASS" : "FAIL"}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Buy Box */}
      {deal.buyBox && (() => { const bs = buyBoxStatus(deal.buyBox); return bs ? (
        <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-5">
          <p className="text-xs text-zinc-500 uppercase tracking-wider mb-3">Buy Box</p>
          <div className={`rounded-xl p-3 border text-center mb-3 ${bs.fits ? "bg-emerald-500/10 border-emerald-500/30" : bs.hardFails.length > 0 ? "bg-red-500/10 border-red-500/30" : "bg-amber-500/10 border-amber-500/30"}`}>
            <p className={`text-sm font-bold ${bs.fits ? "text-emerald-400" : bs.hardFails.length > 0 ? "text-red-400" : "text-amber-400"}`}>{bs.fits ? "✓ Fits Buy Box" : bs.hardFails.length > 0 ? `✕ ${bs.hardFails[0]}` : "⚡ Partial Fit"}</p>
          </div>
          <div className="flex items-center justify-between text-xs bg-zinc-950 rounded-lg px-3 py-2">
            <span className="text-zinc-400">Value-Add Score</span>
            <span className={`font-bold ${bs.leverScore >= 14 ? "text-emerald-400" : bs.leverScore >= 7 ? "text-amber-400" : "text-red-400"}`}>{bs.leverScore}/30</span>
          </div>
          {bs.demandDrivers.length > 0 && <p className="text-xs text-emerald-400 mt-2">{bs.demandDrivers.length} demand driver{bs.demandDrivers.length !== 1 ? "s" : ""} identified</p>}
        </div>
      ) : null; })()}

      {/* Offer Calc */}
      <OfferMultipleCalc scorecard={sc} />

      {/* Add more docs */}
      <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-5">
        <p className="text-xs text-zinc-500 uppercase tracking-wider mb-3">Add More Documents</p>
        <DocUploadZone onExtracted={handleExtracted} existingDocs={sc.docs||[]} dealName={deal.name} />
      </div>
    </div>
  );
}

// ── Deal Detail ───────────────────────────────────────────────────────────────
// ── Buy Box ───────────────────────────────────────────────────────────────────
const EMPTY_BUYBOX = {
  sqft: "", mhi: "", sqftPerCapita: "", lastSaleYear: "", floodZone: false,
  tenantProtection: 0, rateIncrease: 0, gmbOptimization: 0,
  websiteUtilization: 0, feeCollection: 0, expenseReduction: 0,
  notes: "",
  // Market Intel
  majorEmployers: "",
  technicalSchools: false,
  university: false,
  countySeat: false,
  militaryBase: false,
  hospitalAnchor: false,
  marketNotes: ""
};

function buyBoxStatus(bb) {
  if (!bb) return null;
  const hardFails = [];
  if (bb.sqft && parseFloat(bb.sqft) < 10000) hardFails.push("Under 10k sqft");
  if (bb.mhi && parseFloat(bb.mhi) < 50000) hardFails.push("MHI below $50k");
  if (bb.lastSaleYear && (new Date().getFullYear() - parseInt(bb.lastSaleYear)) < 10) hardFails.push("Sold within 10 years");
  if (bb.sqftPerCapita && parseFloat(bb.sqftPerCapita) >= 8) hardFails.push(`Market saturated (${bb.sqftPerCapita} sqft/capita)`);
  if (bb.floodZone) hardFails.push("In flood zone");
  const leverScore = (parseInt(bb.tenantProtection)||0) + (parseInt(bb.rateIncrease)||0) + (parseInt(bb.gmbOptimization)||0) + (parseInt(bb.websiteUtilization)||0) + (parseInt(bb.feeCollection)||0) + (parseInt(bb.expenseReduction)||0);
  const demandDrivers = ["countySeat","university","technicalSchools","militaryBase","hospitalAnchor"].filter(k => bb[k]);
  return { hardFails, leverScore, demandDrivers, fits: hardFails.length === 0 && leverScore >= 10 };
}

function LeverInput({ label, desc, value, onChange }) {
  const v = parseInt(value) || 0;
  const color = v >= 4 ? "text-emerald-400" : v >= 3 ? "text-amber-400" : v >= 1 ? "text-orange-400" : "text-zinc-600";
  return (
    <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
      <div className="flex items-start justify-between mb-2">
        <div className="flex-1 min-w-0 mr-4">
          <p className="text-sm font-medium text-zinc-200">{label}</p>
          <p className="text-xs text-zinc-500 mt-0.5">{desc}</p>
        </div>
        <span className={`text-lg font-bold shrink-0 ${color}`} style={{ fontFamily: "monospace" }}>{v}/5</span>
      </div>
      <div className="flex gap-1.5 mt-3">
        {[1,2,3,4,5].map(n => (
          <button key={n} onClick={() => onChange(v === n ? 0 : n)}
            className={`flex-1 py-2 rounded-lg text-xs font-semibold transition-colors
              ${v >= n ? "bg-amber-500 text-zinc-950" : "bg-zinc-800 text-zinc-500 hover:bg-zinc-700"}`}>
            {n}
          </button>
        ))}
      </div>
    </div>
  );
}

function BuyBoxPanel({ deal, onSave }) {
  const [bb, setBb] = useState(deal.buyBox ? { ...EMPTY_BUYBOX, ...deal.buyBox } : EMPTY_BUYBOX);
  const set = (k) => (e) => setBb(prev => ({ ...prev, [k]: e.target?.value ?? e }));
  const setLever = (k) => (v) => setBb(prev => ({ ...prev, [k]: v }));
  const status = buyBoxStatus(bb);
  const leverScore = status?.leverScore || 0;
  const leverColor = leverScore >= 22 ? "text-emerald-400" : leverScore >= 14 ? "text-amber-400" : leverScore >= 7 ? "text-orange-400" : "text-red-400";
  const leverLabel = leverScore >= 22 ? "Strong Opportunity" : leverScore >= 14 ? "Moderate Opportunity" : leverScore >= 7 ? "Limited Opportunity" : "Weak Opportunity";

  return (
    <div className="space-y-5">
      {/* Hard criteria */}
      <div>
        <p className="text-xs font-semibold text-amber-500 uppercase tracking-widest mb-3">Hard Criteria</p>
        <div className="space-y-3">
          {/* Sqft */}
          <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
            <div className="flex items-center justify-between mb-2">
              <div>
                <p className="text-sm font-medium text-zinc-200">Rentable Square Footage</p>
                <p className="text-xs text-zinc-500">Minimum 10,000 sqft</p>
              </div>
              {bb.sqft && (
                <span className={`text-xs font-semibold px-2 py-0.5 rounded-full border ${parseFloat(bb.sqft) >= 10000 ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : "bg-red-500/20 text-red-400 border-red-500/30"}`}>
                  {parseFloat(bb.sqft) >= 10000 ? "PASS" : "FAIL"}
                </span>
              )}
            </div>
            <input className={inputCls} value={bb.sqft} onChange={set("sqft")} placeholder="15300" type="number" />
          </div>

          {/* MHI */}
          <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
            <div className="flex items-center justify-between mb-2">
              <div>
                <p className="text-sm font-medium text-zinc-200">Median Household Income</p>
                <p className="text-xs text-zinc-500">Minimum $50,000 in trade area</p>
              </div>
              {bb.mhi && (
                <span className={`text-xs font-semibold px-2 py-0.5 rounded-full border ${parseFloat(bb.mhi) >= 50000 ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : "bg-red-500/20 text-red-400 border-red-500/30"}`}>
                  {parseFloat(bb.mhi) >= 50000 ? "PASS" : "FAIL"}
                </span>
              )}
            </div>
            <input className={inputCls} value={bb.mhi} onChange={set("mhi")} placeholder="62000" type="number" />
          </div>

          {/* Sqft per capita */}
          <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
            <div className="flex items-center justify-between mb-2">
              <div>
                <p className="text-sm font-medium text-zinc-200">Market Saturation (Sqft/Capita)</p>
                <p className="text-xs text-zinc-500">Must be below 8 sqft per capita</p>
              </div>
              {bb.sqftPerCapita && (
                <span className={`text-xs font-semibold px-2 py-0.5 rounded-full border ${parseFloat(bb.sqftPerCapita) < 8 ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : "bg-red-500/20 text-red-400 border-red-500/30"}`}>
                  {parseFloat(bb.sqftPerCapita) < 8 ? "PASS" : "FAIL"}
                </span>
              )}
            </div>
            <input className={inputCls} value={bb.sqftPerCapita} onChange={set("sqftPerCapita")} placeholder="6.2" type="number" step="0.1" />
          </div>

          {/* Last sale */}
          <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
            <div className="flex items-center justify-between mb-2">
              <div>
                <p className="text-sm font-medium text-zinc-200">Last Sale Year</p>
                <p className="text-xs text-zinc-500">Must not have sold in last 10 years</p>
              </div>
              {bb.lastSaleYear && (
                <span className={`text-xs font-semibold px-2 py-0.5 rounded-full border ${(new Date().getFullYear() - parseInt(bb.lastSaleYear)) >= 10 ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : "bg-red-500/20 text-red-400 border-red-500/30"}`}>
                  {(new Date().getFullYear() - parseInt(bb.lastSaleYear)) >= 10 ? "PASS" : "FAIL"}
                </span>
              )}
            </div>
            <input className={inputCls} value={bb.lastSaleYear} onChange={set("lastSaleYear")} placeholder="2008" type="number" />
          </div>

          {/* Flood zone */}
          <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-sm font-medium text-zinc-200">Flood Zone</p>
                <p className="text-xs text-zinc-500">Property must not be in a flood zone</p>
              </div>
              <div className="flex items-center gap-2">
                <span className={`text-xs ${bb.floodZone ? "text-red-400" : "text-zinc-500"}`}>{bb.floodZone ? "IN FLOOD ZONE" : "CLEAR"}</span>
                <button onClick={() => setBb(p => ({ ...p, floodZone: !p.floodZone }))}
                  className={`relative w-11 h-6 rounded-full transition-colors shrink-0 ${bb.floodZone ? "bg-red-500" : "bg-zinc-700"}`}>
                  <span className="absolute top-0.5 w-5 h-5 bg-white rounded-full shadow transition-all" style={{ left: bb.floodZone ? "1.375rem" : "0.125rem" }} />
                </button>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Hard criteria verdict */}
      {status && (
        <div className={`rounded-xl p-4 border ${status.hardFails.length === 0 ? "bg-emerald-500/10 border-emerald-500/30" : "bg-red-500/10 border-red-500/30"}`}>
          {status.hardFails.length === 0
            ? <p className="text-sm font-bold text-emerald-400 text-center">✓ Passes all hard criteria</p>
            : <div>
                <p className="text-sm font-bold text-red-400 mb-1">✕ Hard criteria failed:</p>
                {status.hardFails.map((f, i) => <p key={i} className="text-xs text-red-300">• {f}</p>)}
              </div>
          }
        </div>
      )}

      {/* Value-add levers */}
      <div>
        <div className="flex items-center justify-between mb-3">
          <p className="text-xs font-semibold text-amber-500 uppercase tracking-widest">Value-Add Opportunity</p>
          <div className="text-right">
            <span className={`text-sm font-bold ${leverColor}`} style={{ fontFamily: "monospace" }}>{leverScore}/30</span>
            <p className={`text-xs ${leverColor}`}>{leverLabel}</p>
          </div>
        </div>
        <div className="space-y-3">
          <LeverInput label="Tenant Protection" desc="Room to add or optimize tenant insurance program" value={bb.tenantProtection} onChange={setLever("tenantProtection")} />
          <LeverInput label="Rate Increases" desc="Current rates below market — pricing upside available" value={bb.rateIncrease} onChange={setLever("rateIncrease")} />
          <LeverInput label="GMB Optimization" desc="Google Business Profile missing, incomplete, or unoptimized" value={bb.gmbOptimization} onChange={setLever("gmbOptimization")} />
          <LeverInput label="Website & Online Leasing" desc="No website, poor SEO, or no online rental capability" value={bb.websiteUtilization} onChange={setLever("websiteUtilization")} />
          <LeverInput label="Fee Collection" desc="Late fees, admin fees, or other fees not being collected" value={bb.feeCollection} onChange={setLever("feeCollection")} />
          <LeverInput label="Expense Reduction" desc="Bloated expenses that can be cut — management, utilities, contracts" value={bb.expenseReduction} onChange={setLever("expenseReduction")} />
        </div>
      </div>

      {/* Market Intel */}
      <div>
        <p className="text-xs font-semibold text-amber-500 uppercase tracking-widest mb-3">Market Intel</p>
        <div className="space-y-3">
          {/* Major employers */}
          <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4">
            <label className={labelCls}>Major Employers</label>
            <input className={inputCls + " mt-1"} value={bb.majorEmployers} onChange={set("majorEmployers")} placeholder="e.g. Walmart Distribution, County Hospital, School District" />
          </div>

          {/* Demand driver toggles */}
          <div className="bg-zinc-950 border border-zinc-800 rounded-xl p-4 space-y-3">
            <p className="text-xs text-zinc-500 uppercase tracking-wider">Demand Drivers</p>
            {[
              ["countySeat", "County Seat", "Government employees, court activity, stable civic demand"],
              ["university", "University / College", "Student storage, faculty housing, recurring annual demand"],
              ["technicalSchools", "Technical / Trade School", "Workforce training programs, transient student population"],
              ["militaryBase", "Military Base Nearby", "PCS moves, deployment storage, high-turnover demand"],
              ["hospitalAnchor", "Hospital / Medical Center", "Healthcare workers, stable long-term employment anchor"],
            ].map(([key, label, desc]) => (
              <div key={key} className="flex items-start justify-between gap-3">
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-medium text-zinc-200">{label}</p>
                  <p className="text-xs text-zinc-500 mt-0.5">{desc}</p>
                </div>
                <div className="flex items-center gap-2 shrink-0 mt-0.5">
                  <span className={`text-xs ${bb[key] ? "text-emerald-400" : "text-zinc-600"}`}>{bb[key] ? "YES" : "NO"}</span>
                  <button onClick={() => setBb(p => ({ ...p, [key]: !p[key] }))}
                    className={`relative w-11 h-6 rounded-full transition-colors ${bb[key] ? "bg-amber-500" : "bg-zinc-700"}`}>
                    <span className="absolute top-0.5 w-5 h-5 bg-white rounded-full shadow transition-all" style={{ left: bb[key] ? "1.375rem" : "0.125rem" }} />
                  </button>
                </div>
              </div>
            ))}
          </div>

          {/* Demand score summary */}
          {(() => {
            const drivers = ["countySeat","university","technicalSchools","militaryBase","hospitalAnchor"].filter(k => bb[k]);
            return drivers.length > 0 ? (
              <div className="bg-emerald-500/10 border border-emerald-500/20 rounded-lg px-4 py-3">
                <p className="text-xs text-emerald-400 font-medium">{drivers.length} demand driver{drivers.length !== 1 ? "s" : ""} identified — strong market fundamentals</p>
              </div>
            ) : (
              <div className="bg-zinc-900 border border-zinc-800 rounded-lg px-4 py-3">
                <p className="text-xs text-zinc-500">No demand drivers identified yet — check local employment and anchor institutions</p>
              </div>
            );
          })()}

          {/* Market notes */}
          <div>
            <label className={labelCls}>Market Notes</label>
            <textarea className={inputCls + " resize-none mt-1"} rows={3} value={bb.marketNotes} onChange={set("marketNotes")} placeholder="e.g. Gilmer is the Upshur County seat, Gilmer ISD is largest employer, East Texas Medical Center nearby" />
          </div>
        </div>
      </div>

      {/* Notes */}
      <div>
        <label className={labelCls}>Additional Notes</label>
        <textarea className={inputCls + " resize-none"} rows={3} value={bb.notes} onChange={set("notes")} placeholder="Other value-add observations…" />
      </div>

      <button onClick={() => onSave(bb)}
        className="w-full bg-amber-500 hover:bg-amber-400 text-zinc-950 font-bold rounded-lg py-3 text-sm transition-colors">
        Save Buy Box
      </button>
    </div>
  );
}

function DealDetail({ deal, onBack, onUpdateDeal }) {
  const [tab, setTab] = useState("report");

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      <div className="border-b border-zinc-800 px-6 py-4">
        <button onClick={onBack} className="text-xs text-zinc-500 hover:text-zinc-300 mb-3 flex items-center gap-1 transition-colors">← Back to Pipeline</button>
        <div className="flex gap-1 overflow-x-auto">
          {[["report","Report"],["edit","Edit Fields"],["buybox","Buy Box"],["killswitch","Kill Switch"]].map(([id,label]) => (
            <button key={id} onClick={() => setTab(id)}
              className={`px-4 py-2 text-sm font-medium rounded-lg transition-colors shrink-0 ${tab === id ? "bg-zinc-800 text-zinc-100" : "text-zinc-500 hover:text-zinc-300"}`}>
              {label}
            </button>
          ))}
        </div>
      </div>

      <div className="p-6 max-w-2xl mx-auto">
        {tab === "report" && <ReportView deal={deal} onUpdateDeal={onUpdateDeal} />}
        {tab === "edit" && (
          <EditPanel deal={deal} onSave={(sc) => {
            onUpdateDeal({ ...deal, scorecard: sc, status: "complete" });
            setTab("report");
          }} />
        )}
        {tab === "buybox" && (
          <BuyBoxPanel deal={deal} onSave={(bb) => {
            onUpdateDeal({ ...deal, buyBox: bb });
            setTab("report");
          }} />
        )}
        {tab === "killswitch" && (
          <KillSwitchPanel deal={deal} onSave={(ks) => {
            onUpdateDeal({ ...deal, killSwitch: ks, status: ks.overallPass ? "analyzing" : "killed" });
            setTab("report");
          }} />
        )}
      </div>
    </div>
  );
}

// ── Dashboard ─────────────────────────────────────────────────────────────────
function Dashboard({ deals, onNew, onOpen, onDelete }) {
  const stats = { total: deals.length, killed: deals.filter(d => d.status === "killed").length, complete: deals.filter(d => d.status === "complete").length, go: deals.filter(d => d.scorecard?.score?.rec === "GO").length };
  const [apiStatus, setApiStatus] = useState(null); // null | "testing" | "ok" | "fail"
  const [apiMsg, setApiMsg] = useState("");

  const [showSettings, setShowSettings] = useState(false);
  const [apiUrl, setApiUrl] = useState(() => window.__STORAGEIQ_API__ || "");

  const saveApiUrl = () => {
    const url = apiUrl.replace(/\/$/, ""); // strip trailing slash
    window.__STORAGEIQ_API__ = url;
    try { localStorage.setItem("storageiq_api_url", url); } catch {}
    // In production, VITE_API_URL takes precedence
    setShowSettings(false);
    setApiStatus(null);
  };

  // Load saved URL on mount
  useEffect(() => {
    try {
      const saved = localStorage.getItem("storageiq_api_url");
      if (saved) { window.__STORAGEIQ_API__ = saved; setApiUrl(saved); }
    } catch {}
  }, []);

  const testApi = async () => {
    const endpoint = import.meta.env.VITE_API_URL || window.__STORAGEIQ_API__;
    if (!endpoint) { setApiStatus("fail"); setApiMsg("No URL set — add VITE_API_URL to environment variables"); return; }
    setApiStatus("testing");
    setApiMsg("");
    try {
      const resp = await fetch(`${endpoint}/api/extract`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ max_tokens: 20, messages: [{ role: "user", content: "Reply with just: OK" }] })
      });
      const data = await resp.json();
      if (data.error) {
        setApiStatus("fail");
        setApiMsg(`API error: ${data.error?.type || ""} — ${data.error?.message || JSON.stringify(data.error)}`);
      } else if (data.content?.[0]?.text) {
        setApiStatus("ok");
        setApiMsg(`Connected ✓  ${endpoint}`);
      } else {
        setApiStatus("fail");
        setApiMsg(`Unexpected response: ${JSON.stringify(data).slice(0, 120)}`);
      }
    } catch (e) {
      setApiStatus("fail");
      setApiMsg(`Could not reach ${endpoint} — ${e.message}`);
    }
  };

  return (
    <div className="min-h-screen bg-zinc-950 text-zinc-100">
      <div className="border-b border-zinc-800 px-6 py-5 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-amber-500 rounded-lg flex items-center justify-center text-zinc-950 font-bold text-sm">S</div>
          <div><p className="font-bold text-zinc-100 leading-none">StorageIQ</p><p className="text-xs text-zinc-500">Deal Analyzer</p></div>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={() => setShowSettings(v => !v)}
            className="text-xs text-zinc-500 hover:text-zinc-300 border border-zinc-700 hover:border-zinc-500 rounded-lg px-3 py-1.5 transition-colors">
            ⚙ Settings
          </button>
          <button onClick={testApi} disabled={apiStatus === "testing"}
            className="text-xs text-zinc-500 hover:text-zinc-300 border border-zinc-700 hover:border-zinc-500 rounded-lg px-3 py-1.5 transition-colors disabled:opacity-50">
            {apiStatus === "testing" ? "Testing…" : "Test AI"}
          </button>
          <button onClick={onNew} className="bg-amber-500 hover:bg-amber-400 text-zinc-950 font-semibold text-sm px-4 py-2 rounded-lg transition-colors">+ Analyze Deal</button>
        </div>
      </div>

      {/* Settings panel */}
      {showSettings && (
        <div className="border-b border-zinc-800 px-6 py-4 bg-zinc-900">
          <p className="text-xs text-zinc-500 uppercase tracking-wider mb-2">AI Extraction Endpoint</p>
          <p className="text-xs text-zinc-600 mb-3">Paste your Vercel URL — e.g. https://storageiq-backend.vercel.app</p>
          <div className="flex gap-2">
            <input
              className="flex-1 bg-zinc-950 border border-zinc-700 rounded-lg px-3 py-2 text-sm text-zinc-100 focus:outline-none focus:border-amber-500"
              value={apiUrl}
              onChange={e => setApiUrl(e.target.value)}
              placeholder="https://your-project.vercel.app"
            />
            <button onClick={saveApiUrl} className="bg-amber-500 hover:bg-amber-400 text-zinc-950 font-semibold text-sm px-4 py-2 rounded-lg transition-colors">
              Save
            </button>
          </div>
        </div>
      )}

      {apiStatus && apiStatus !== "testing" && (
        <div className={`px-6 py-2 text-xs border-b ${apiStatus === "ok" ? "bg-emerald-500/10 border-emerald-500/20 text-emerald-400" : "bg-red-500/10 border-red-500/20 text-red-400"}`}>
          {apiMsg}
        </div>
      )}

      <div className="px-6 py-5 grid grid-cols-4 gap-3">
        {[["Total", stats.total, "text-zinc-300"],["Killed", stats.killed, "text-red-400"],["Complete", stats.complete, "text-emerald-400"],["GO Deals", stats.go, "text-amber-400"]].map(([l, v, c]) => (
          <div key={l} className="bg-zinc-900 border border-zinc-800 rounded-xl p-4">
            <p className="text-xs text-zinc-500 mb-1">{l}</p>
            <p className={`text-2xl font-bold ${c}`} style={{ fontFamily: "monospace" }}>{v}</p>
          </div>
        ))}
      </div>

      <div className="px-6 pb-8">
        <p className="text-xs text-zinc-500 uppercase tracking-wider mb-3">Deal Pipeline</p>
        {deals.length === 0 ? (
          <div className="text-center py-16 border border-dashed border-zinc-800 rounded-2xl">
            <p className="text-4xl mb-3">🏢</p>
            <p className="font-medium text-zinc-400">No deals yet</p>
            <p className="text-sm text-zinc-600 mt-1 mb-4">Add your first self-storage deal to get started.</p>
            <button onClick={onNew} className="bg-amber-500 hover:bg-amber-400 text-zinc-950 font-semibold text-sm px-4 py-2 rounded-lg transition-colors">+ Analyze Deal</button>
          </div>
        ) : (
          <div className="space-y-2">
            {deals.map(deal => {
              const sc = deal.scorecard;
              const pct = sc ? sc.score.total : null;
              const docCount = sc?.docs?.length || 0;
              return (
                <div key={deal.id} onClick={() => onOpen(deal.id)} className="bg-zinc-900 border border-zinc-800 hover:border-zinc-600 rounded-xl p-4 cursor-pointer transition-all group">
                  <div className="flex items-center justify-between gap-4">
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2 mb-0.5 flex-wrap">
                        <p className="font-semibold text-zinc-200 text-sm truncate">{deal.name}</p>
                        <span className={`text-xs px-1.5 py-0.5 rounded-md font-medium shrink-0 ${statusColor(deal.status)}`}>{deal.status}</span>
                        {(() => { const bs = buyBoxStatus(deal.buyBox); return bs ? <span className={`text-xs px-1.5 py-0.5 rounded-md font-medium shrink-0 border ${bs.fits ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : bs.hardFails.length > 0 ? "bg-red-500/20 text-red-400 border-red-500/30" : "bg-amber-500/20 text-amber-400 border-amber-500/30"}`}>{bs.fits ? "✓ BUY BOX" : bs.hardFails.length > 0 ? "✗ NO FIT" : "⚡ PARTIAL"}</span> : null; })()}
                        {docCount > 0 && <span className="text-xs text-zinc-600 shrink-0">📎 {docCount} doc{docCount !== 1 ? "s" : ""}</span>}
                      </div>
                      <p className="text-xs text-zinc-500 truncate">{deal.address}{deal.city ? `, ${deal.city}` : ""} {deal.zip}</p>
                    </div>
                    <div className="flex items-center gap-4 shrink-0">
                      {sc ? (
                        <div className="text-right">
                          <div className="flex items-center gap-2 mb-1">
                            <span className={`text-xs px-2 py-0.5 rounded-full border ${recColor(sc.score.rec)}`}>{sc.score.rec}</span>
                            <span className="text-sm font-bold" style={{ fontFamily: "monospace", color: scoreColor(pct) }}>{sc.score.total}</span>
                          </div>
                          <div className="w-32 h-1.5 rounded-full bg-zinc-800">
                            <div className="h-full rounded-full" style={{ width: `${pct}%`, backgroundColor: scoreColor(pct) }} />
                          </div>
                        </div>
                      ) : (
                        <span className="text-xs text-zinc-600">Not scored</span>
                      )}
                      <button onClick={(e) => { e.stopPropagation(); onDelete(deal.id); }} className="opacity-0 group-hover:opacity-100 text-zinc-600 hover:text-red-400 transition-all text-lg leading-none">×</button>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

// ── Storage helpers ───────────────────────────────────────────────────────────
const STORAGE_KEY = "storageiq-deals";

async function loadDeals() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

async function saveDeals(deals) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(deals));
  } catch (e) {
    console.error("Save failed:", e);
  }
}

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const [deals, setDeals] = useState([]);
  const [view, setView] = useState("dashboard");
  const [activeDealId, setActiveDealId] = useState(null);
  const [showNewDeal, setShowNewDeal] = useState(false);
  const [loading, setLoading] = useState(true);
  const activeDeal = deals.find(d => d.id === activeDealId);

  // Load deals from storage on mount
  useEffect(() => {
    loadDeals().then(saved => {
      setDeals(saved);
      setLoading(false);
    });
  }, []);

  // Listen for Claude's storageiq JSON responses sent back through the chat
  useEffect(() => {
    const handleMessage = (event) => {
      try {
        const text = event.data?.content || event.data?.text || "";
        if (typeof text !== "string") return;
        const match = text.match(/```storageiq\s*([\s\S]*?)```/);
        if (!match) return;
        const data = JSON.parse(match[1].trim());
        // Route to the right handler
        if (window.__storageIQNewDeal) {
          window.__storageIQNewDeal(data);
        } else if (window.__storageIQFill) {
          window.__storageIQFill(data);
        }
      } catch {}
    };
    window.addEventListener("message", handleMessage);
    return () => window.removeEventListener("message", handleMessage);
  }, []);

  // Save deals to storage whenever they change
  useEffect(() => {
    if (!loading) saveDeals(deals);
  }, [deals, loading]);

  const createDeal = (form) => {
    const deal = {
      ...form,
      id: Date.now(),
      status: "pending",
      docs: [],
      createdAt: new Date().toISOString()
    };
    setDeals(prev => [deal, ...prev]);
    setActiveDealId(deal.id);
    setView("detail");
    setShowNewDeal(false);
  };

  const createBlankDeal = () => {
    const deal = { name: "New Deal", id: Date.now(), status: "pending", docs: [], createdAt: new Date().toISOString() };
    setDeals(prev => [deal, ...prev]);
    setActiveDealId(deal.id);
    setView("detail");
  };

  const updateDeal = (updated) => setDeals(prev => prev.map(d => d.id === updated.id ? updated : d));
  const deleteDeal = (id) => setDeals(prev => prev.filter(d => d.id !== id));

  if (loading) {
    return (
      <div className="min-h-screen bg-zinc-950 flex items-center justify-center">
        <div className="text-center">
          <div className="w-8 h-8 border-2 border-amber-500 border-t-transparent rounded-full animate-spin mx-auto mb-3" />
          <p className="text-sm text-zinc-500">Loading your deals…</p>
        </div>
      </div>
    );
  }

  if (view === "detail" && activeDeal) {
    return <DealDetail deal={activeDeal} onBack={() => setView("dashboard")} onUpdateDeal={updateDeal} />;
  }
  return <Dashboard deals={deals} onNew={createBlankDeal} onOpen={(id) => { setActiveDealId(id); setView("detail"); }} onDelete={deleteDeal} />;
}
