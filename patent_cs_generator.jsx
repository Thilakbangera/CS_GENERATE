import { useState, useRef, useEffect } from "react";

const PATENT_TYPES = [
  { value: "method", label: "Method only", prefix: "A METHOD FOR" },
  { value: "system_method", label: "System & Method", prefix: "A SYSTEM AND METHOD FOR" },
  { value: "method_composition", label: "Method & Composition", prefix: "A METHOD AND COMPOSITION FOR" },
];

const ALL_SECTIONS = [
  { id: "title", label: "Title" },
  { id: "abstract", label: "Abstract" },
  { id: "field", label: "Technical Field" },
  { id: "background", label: "Background" },
  { id: "objectives", label: "Objectives" },
  { id: "summary", label: "Summary" },
  { id: "detailed_description", label: "Detailed Description" },
  { id: "claims_support", label: "Claims Support / Rationale" },
];

const SECTION_LABELS = {
  title: "TITLE", abstract: "ABSTRACT", field: "TECHNICAL FIELD",
  background: "BACKGROUND", objectives: "OBJECTIVES", summary: "SUMMARY",
  detailed_description: "DETAILED DESCRIPTION", claims_support: "CLAIMS SUPPORT / RATIONALE",
};

function buildPrompt(sectionId, patentType, ctx, hasImages) {
  const prefix = PATENT_TYPES.find(t => t.value === patentType)?.prefix || "A METHOD FOR";
  const isMethodOnly = patentType === "method";
  const isComposition = patentType === "method_composition";

  const base = `You are an expert Indian patent drafter. Draft the requested section using the invention context below.
Rules: third person perspective | no camel casing | each paragraph ≤100 words | use **bold** for all reference numbers like **(100)** | output only the drafted content, no extra commentary.

INVENTION CONTEXT:
${ctx}

TASK: `;

  const prompts = {
    title: `${base}Generate the patent title in exactly 13 words. Start with "${prefix}". Output only the title.`,

    abstract: `${base}Generate an abstract in 100-140 words, third person, general description without listing each component individually. Bold reference numbers with **bold**.`,

    field: `${base}Generate Technical Field section. Sentence 1 starts 'The present invention relates to' (broad domain). Sentence 2 starts 'More specifically, the invention relates to' (specific details). Bold reference numbers.`,

    background: `${base}Generate Background — exactly 2 paragraphs each ≤100 words:
Para 1: generic problem in the broader domain.
Para 2: specific problem for this invention. Must end with a sentence starting 'Hence there is a need for' (broad perspective). Bold reference numbers.`,

    objectives: `${base}Generate 4-5 objectives:
First: 'A principal objective of the invention is to develop a...'
Rest: 'Another objective of the invention is to...'
Use 'the' not 'a' for previously mentioned things. Each 1-2 lines, general. Bold reference numbers.`,

    summary: `${base}Generate Summary — exactly 3 paragraphs each ≤100 words:
Para 1 starts: 'The present invention provides a...' — main invention details.
Para 2 starts: 'In some embodiments, the...' — overcomes prior art, variations.
Para 3 starts: 'Embodiments of the invention may further provide a method for...' — method overview.
Bold reference numbers.`,

    detailed_description: `${base}Generate Detailed Description in 4 parts:

PART 1 — DEFINITIONS:
Pick 10 broad keywords from the claims (not drawings). For each keyword write a definition paragraph starting 'As used in this application, the term "[X]" refers to...' or 'The term "[X]" refers to...'. Each ≤100 words. Bold reference numbers.

PART 2 — COMPONENTS:
${isMethodOnly ? "(Method-only patent: skip Part 2)" : `For each main component, write 5-7 sentences: (a) what it is generally, (b) its role in this invention, (c) how it connects to other components. ≤100 words per component. Bold reference numbers like **(100)**.`}

PART 3 — METHOD STEPS:
Write 5 steps:
Format: 'At step [201/203/205/207/209], the method includes [ING-word]...'
Each step continues from the previous. Use component names. Each ≤100 words.
After each step, write 'According to an example embodiment...' paragraph (≤100 words).

PART 4 — FIGURES:
${hasImages ? "For each figure, write: 'FIG. X illustrates...' then 'According to an example embodiment...' (≤100 words each)." : "(No images uploaded — skip Part 4)"}

${isMethodOnly ? "NOTE: Method-only — Part 1 and Part 3 compulsory, Part 2 skipped." : ""}
${isComposition ? "NOTE: Method & Composition — Part 2 must describe each composition/material component." : ""}`,

    claims_support: `${base}Generate Claims Support / Rationale:

Para A: State the key components and their specific quantifiable limiting parameters (ranges/values from claims) that together produce the technical effects. ≤100 words.

Para B onward (one per major technical effect): Start 'As demonstrated in [figure/result], the aforementioned combination...' — state quantifiable result, % improvement over baseline, mechanism. ≤100 words each.

Final Para: Out-of-range consequences — what happens when parameters exceed upper limit or fall below lower limit, effect on each technical outcome. ≤100 words.
Bold all reference numbers.`,
  };
  return prompts[sectionId] || `${base}Generate the ${sectionId} section.`;
}

async function callClaude(prompt) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  if (!res.ok) { const e = await res.json().catch(() => ({})); throw new Error(e.error?.message || `HTTP ${res.status}`); }
  const d = await res.json();
  return d.content?.map(b => b.text || "").join("") || "";
}

async function extractFileText(file) {
  if (!file) return "";
  const b64 = await new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("Read failed"));
    r.readAsDataURL(file);
  });
  const ext = file.name.split(".").pop().toLowerCase();
  const mediaType = ext === "pptx"
    ? "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514", max_tokens: 1000,
      messages: [{ role: "user", content: [
        { type: "document", source: { type: "base64", media_type: mediaType, data: b64 } },
        { type: "text", text: "Extract all text content from this document. Return only the raw text." }
      ]}]
    }),
  });
  if (!res.ok) return "";
  const d = await res.json();
  return d.content?.map(b => b.text || "").join("") || "";
}

function loadDocxScript() {
  return new Promise((res, rej) => {
    if (window.docx) { res(); return; }
    const s = document.createElement("script");
    s.src = "https://unpkg.com/docx@9.6.1/build/index.js";
    s.onload = res;
    s.onerror = () => rej(new Error("Failed to load docx library"));
    document.head.appendChild(s);
  });
}

async function downloadDocx(results, sections) {
  await loadDocxScript();
  const { Document, Packer, Paragraph, TextRun, HeadingLevel } = window.docx;

  function parseRuns(text) {
    const runs = []; const re = /\*\*([^*]+)\*\*/g; let last = 0, m;
    while ((m = re.exec(text)) !== null) {
      if (m.index > last) runs.push(new TextRun({ text: text.slice(last, m.index) }));
      runs.push(new TextRun({ text: m[1], bold: true }));
      last = re.lastIndex;
    }
    if (last < text.length) runs.push(new TextRun({ text: text.slice(last) }));
    return runs;
  }

  function toParagraphs(text) {
    return text.split("\n").filter(l => l.trim()).map(line =>
      new Paragraph({ children: parseRuns(line.trim()), spacing: { after: 160 } })
    );
  }

  const children = [];
  sections.forEach(id => {
    if (!results[id]) return;
    children.push(
      new Paragraph({ text: SECTION_LABELS[id] || id.toUpperCase(), heading: HeadingLevel.HEADING_1, spacing: { before: 400, after: 200 } }),
      ...toParagraphs(results[id])
    );
  });

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 24 } } },
      paragraphStyles: [{
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 },
      }]
    },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children }]
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "Patent_Complete_Specification.docx";
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

function FileBox({ label, sublabel, accept, file, onChange }) {
  const ref = useRef();
  return (
    <div onClick={() => ref.current?.click()} style={{ border: `1px dashed ${file ? "#1D9E75" : "var(--color-border-secondary)"}`, borderRadius: 10, padding: "1rem 0.75rem", textAlign: "center", cursor: "pointer", background: file ? "#E1F5EE" : "var(--color-background-secondary)", minHeight: 88, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 4 }}>
      <input ref={ref} type="file" accept={accept} style={{ display: "none" }} onChange={e => onChange(e.target.files[0])} />
      <i className={`ti ${file ? "ti-circle-check" : "ti-upload"}`} style={{ fontSize: 22, color: file ? "#1D9E75" : "var(--color-text-tertiary)" }} aria-hidden="true" />
      <div style={{ fontSize: 12, fontWeight: 500, color: file ? "#085041" : "var(--color-text-secondary)" }}>{label}</div>
      <div style={{ fontSize: 11, color: file ? "#0F6E56" : "var(--color-text-tertiary)", wordBreak: "break-all", padding: "0 4px" }}>{file ? file.name : sublabel}</div>
    </div>
  );
}

export default function App() {
  const [claims, setClaims] = useState(null);
  const [drawings, setDrawings] = useState(null);
  const [disclosure, setDisclosure] = useState(null);
  const [patentType, setPatentType] = useState("method_composition");
  const [selected, setSelected] = useState(ALL_SECTIONS.map(s => s.id));
  const [generating, setGenerating] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0, label: "" });
  const [results, setResults] = useState({});
  const [done, setDone] = useState(false);
  const [error, setError] = useState("");

  const toggleSection = id => setSelected(p => p.includes(id) ? p.filter(x => x !== id) : [...p, id]);
  const canGenerate = claims && disclosure && selected.length > 0 && !generating;
  const pct = progress.total > 0 ? Math.round((progress.current / progress.total) * 100) : 0;

  const handleGenerate = async () => {
    setError(""); setDone(false); setResults({}); setGenerating(true);
    try {
      setProgress({ current: 0, total: 1, label: "Reading uploaded files..." });
      const [claimsText, drawingsText, disclosureText] = await Promise.all([
        extractFileText(claims), extractFileText(drawings), extractFileText(disclosure),
      ]);
      const ctx = `CLAIMS:\n${claimsText}\n\nDRAWINGS:\n${drawingsText}\n\nINVENTION DISCLOSURE:\n${disclosureText}`;
      const secs = ALL_SECTIONS.filter(s => selected.includes(s.id));
      const res = {};
      for (let i = 0; i < secs.length; i++) {
        setProgress({ current: i + 1, total: secs.length, label: `Drafting: ${secs[i].label}...` });
        const out = await callClaude(buildPrompt(secs[i].id, patentType, ctx, !!drawings));
        res[secs[i].id] = out;
        setResults({ ...res });
      }
      setDone(true);
    } catch (e) { setError(e.message || "Something went wrong."); }
    finally { setGenerating(false); }
  };

  const handleDownload = async () => {
    setDownloading(true);
    try { await downloadDocx(results, selected); }
    catch (e) { setError("Download failed: " + e.message); }
    finally { setDownloading(false); }
  };

  const Card = ({ children, style = {} }) => (
    <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, padding: "1.25rem", marginBottom: "1rem", ...style }}>
      {children}
    </div>
  );

  const CardTitle = ({ icon, children }) => (
    <div style={{ fontSize: 14, fontWeight: 500, marginBottom: "0.75rem", display: "flex", alignItems: "center", gap: 8 }}>
      <i className={`ti ti-${icon}`} aria-hidden="true" style={{ fontSize: 16 }} />{children}
    </div>
  );

  return (
    <div style={{ padding: "1.5rem 0", maxWidth: 680, fontFamily: "var(--font-sans)" }}>
      <h1 style={{ fontSize: 20, fontWeight: 500, marginBottom: 4, color: "var(--color-text-primary)" }}>Patent CS Generator</h1>
      <p style={{ fontSize: 14, color: "var(--color-text-secondary)", marginBottom: "1.5rem" }}>Upload your claims, drawings, and invention disclosure — generate a complete patent specification and download as .docx.</p>

      <Card>
        <CardTitle icon="files">Upload Files</CardTitle>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10 }}>
          <FileBox label="Claims (.docx)" sublabel="Required" accept=".docx" file={claims} onChange={setClaims} />
          <FileBox label="Drawings (.pptx)" sublabel="Optional" accept=".pptx,.docx" file={drawings} onChange={setDrawings} />
          <FileBox label="Disclosure (.docx)" sublabel="Required" accept=".docx" file={disclosure} onChange={setDisclosure} />
        </div>
      </Card>

      <Card>
        <CardTitle icon="certificate">Patent Type</CardTitle>
        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
          {PATENT_TYPES.map(t => (
            <label key={t.value} style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 13, cursor: "pointer", padding: "8px 10px", borderRadius: 8, border: `0.5px solid ${patentType === t.value ? "#5DCAA5" : "var(--color-border-tertiary)"}`, background: patentType === t.value ? "#E1F5EE" : "var(--color-background-secondary)", color: patentType === t.value ? "#085041" : "var(--color-text-primary)" }}>
              <input type="radio" name="pt" value={t.value} checked={patentType === t.value} onChange={() => setPatentType(t.value)} style={{ accentColor: "#1D9E75" }} />
              <div>
                <div style={{ fontWeight: 500 }}>{t.label}</div>
                <div style={{ fontSize: 11, color: patentType === t.value ? "#0F6E56" : "var(--color-text-tertiary)" }}>{t.prefix}</div>
              </div>
            </label>
          ))}
        </div>
      </Card>

      <Card>
        <div style={{ fontSize: 14, fontWeight: 500, marginBottom: "0.75rem", display: "flex", alignItems: "center", gap: 8 }}>
          <i className="ti ti-list-check" aria-hidden="true" style={{ fontSize: 16 }} />Sections to Generate
          <button onClick={() => setSelected(selected.length === ALL_SECTIONS.length ? [] : ALL_SECTIONS.map(s => s.id))} style={{ marginLeft: "auto", fontSize: 12, color: "#1D9E75", background: "none", border: "none", cursor: "pointer", padding: 0 }}>
            {selected.length === ALL_SECTIONS.length ? "Deselect all" : "Select all"}
          </button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          {ALL_SECTIONS.map(sec => {
            const chk = selected.includes(sec.id);
            return (
              <label key={sec.id} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, cursor: "pointer", padding: "7px 10px", borderRadius: 8, border: `0.5px solid ${chk ? "#5DCAA5" : "var(--color-border-tertiary)"}`, background: chk ? "#E1F5EE" : "var(--color-background-secondary)", color: chk ? "#085041" : "var(--color-text-primary)", userSelect: "none" }}>
                <input type="checkbox" checked={chk} onChange={() => toggleSection(sec.id)} style={{ accentColor: "#1D9E75", width: 15, height: 15 }} />
                {sec.label}
              </label>
            );
          })}
        </div>
      </Card>

      <button disabled={!canGenerate} onClick={handleGenerate} style={{ width: "100%", padding: "0.8rem", fontSize: 15, fontWeight: 500, borderRadius: 8, border: "none", background: canGenerate ? "#1D9E75" : "var(--color-border-tertiary)", color: canGenerate ? "#fff" : "var(--color-text-tertiary)", cursor: canGenerate ? "pointer" : "not-allowed", marginBottom: "0.75rem" }}>
        {generating ? `Generating... ${pct}%` : "Generate Complete Specification"}
      </button>

      {generating && (
        <div style={{ background: "var(--color-background-secondary)", borderRadius: 8, padding: "1rem", marginBottom: "0.75rem" }}>
          <div style={{ fontSize: 13, color: "var(--color-text-secondary)", marginBottom: 8 }}>{progress.label}</div>
          <div style={{ height: 6, background: "var(--color-border-tertiary)", borderRadius: 99, overflow: "hidden" }}>
            <div style={{ height: "100%", width: `${pct}%`, background: "#1D9E75", borderRadius: 99, transition: "width 0.4s" }} />
          </div>
          <div style={{ fontSize: 12, color: "var(--color-text-tertiary)", marginTop: 6 }}>{progress.current} of {progress.total} sections</div>
        </div>
      )}

      {error && (
        <div style={{ background: "#FCEBEB", border: "0.5px solid #F09595", borderRadius: 8, padding: "0.75rem 1rem", fontSize: 13, color: "#791F1F", marginBottom: "0.75rem" }}>
          <i className="ti ti-alert-circle" style={{ marginRight: 6 }} aria-hidden="true" />{error}
        </div>
      )}

      {Object.keys(results).length > 0 && (
        <Card>
          <div style={{ fontSize: 14, fontWeight: 500, marginBottom: "0.75rem", display: "flex", alignItems: "center", gap: 8 }}>
            <i className="ti ti-file-text" aria-hidden="true" style={{ fontSize: 16 }} />Preview
            <span style={{ marginLeft: "auto", fontSize: 12, color: "var(--color-text-tertiary)" }}>{Object.keys(results).length}/{selected.length} sections done</span>
          </div>
          {ALL_SECTIONS.filter(s => results[s.id]).map(sec => (
            <details key={sec.id} style={{ marginBottom: 8, borderBottom: "0.5px solid var(--color-border-tertiary)", paddingBottom: 8 }}>
              <summary style={{ fontSize: 13, fontWeight: 500, cursor: "pointer", padding: "4px 0", color: "var(--color-text-primary)", listStyle: "none", display: "flex", alignItems: "center", gap: 6 }}>
                <i className="ti ti-check" style={{ color: "#1D9E75", fontSize: 14 }} aria-hidden="true" />{sec.label}
                <i className="ti ti-chevron-down" style={{ marginLeft: "auto", fontSize: 13, color: "var(--color-text-tertiary)" }} aria-hidden="true" />
              </summary>
              <pre style={{ fontSize: 12, color: "var(--color-text-secondary)", whiteSpace: "pre-wrap", wordBreak: "break-word", marginTop: 8, lineHeight: 1.6, fontFamily: "var(--font-sans)", background: "var(--color-background-secondary)", padding: "0.75rem", borderRadius: 6 }}>{results[sec.id]}</pre>
            </details>
          ))}
        </Card>
      )}

      {done && (
        <button onClick={handleDownload} disabled={downloading} style={{ width: "100%", padding: "0.8rem", fontSize: 15, fontWeight: 500, borderRadius: 8, border: "0.5px solid #0F6E56", background: downloading ? "#5DCAA5" : "#1D9E75", color: "#fff", cursor: downloading ? "not-allowed" : "pointer" }}>
          <i className="ti ti-download" style={{ marginRight: 8 }} aria-hidden="true" />
          {downloading ? "Preparing download..." : "Download as Word (.docx)"}
        </button>
      )}
    </div>
  );
}
