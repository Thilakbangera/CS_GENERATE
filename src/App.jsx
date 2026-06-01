import { useState, useRef } from 'react';
import * as docx from 'docx';
import './style.css';
import Login from './Login';


const OLLAMA_URL = import.meta.env.VITE_OLLAMA_URL || 'http://127.0.0.1:11434/v1/completions';
const OLLAMA_MODEL = import.meta.env.VITE_OLLAMA_MODEL || 'llama2';

const PATENT_TYPES = [
  { value: 'method', label: 'Method only', prefix: 'A METHOD FOR' },
  { value: 'system_method', label: 'System & Method', prefix: 'A SYSTEM AND METHOD FOR' },
  { value: 'method_composition', label: 'Method & Composition', prefix: 'A METHOD AND COMPOSITION FOR' },
];

const ALL_SECTIONS = [
  { id: 'title', label: 'Title' },
  { id: 'field', label: 'Technical Field' },
  { id: 'background', label: 'Background' },
  { id: 'objectives', label: 'Objectives' },
  { id: 'summary', label: 'Summary' },
  { id: 'brief_description', label: 'Brief Description of Drawings' },
  { id: 'detailed_description', label: 'Detailed Description' },
  { id: 'abstract', label: 'Abstract' },
];

const SECTION_LABELS = {
  title: 'TITLE', abstract: 'ABSTRACT', field: 'TECHNICAL FIELD',
  background: 'BACKGROUND', objectives: 'OBJECTIVES', summary: 'SUMMARY',
  brief_description: 'BRIEF DESCRIPTION OF DRAWINGS',
  detailed_description: 'DETAILED DESCRIPTION',
};

function buildPrompt(sectionId, patentType, ctx, hasImages) {
  const prefix = PATENT_TYPES.find(t => t.value === patentType)?.prefix || 'A METHOD FOR';
  const isMethodOnly = patentType === 'method';
  const isComposition = patentType === 'method_composition';

  const articleRule = `ARTICLE USAGE RULES (mandatory):
- Track nouns ONLY within this section. Do NOT carry tracking from any other section.
- First time a noun/element/component appears in this section: use "a" or "an".
- Second and every subsequent mention of the same noun in this section: use "the".
- Even if the first reference was in a heading or list, later references must still use "the".
- Restart tracking completely at the beginning of each new section.`;

  const base = `You are an expert Indian patent drafter producing a complete specification for a Form 2 filing. Draft ONLY the requested section. Output only the drafted content with no headings, labels, or commentary.
Rules: third person perspective | no camel casing | use **bold** for all reference numbers like **(100)** | write in full detailed sentences â€” not bullet points | NEVER use the generic shorthand "The system" or "the system" â€” always write the full complete system name with reference number (e.g., "the adaptive prefix-bucket hash table system **(100)**").
${articleRule}

INVENTION CONTEXT:
${ctx}

TASK: `;

  const prompts = {
    title: `${base}Generate the patent title in exactly 13 words. Start with "${prefix}". Output ONLY the title text, nothing else.`,

    abstract: `${base}Generate an abstract that is a direct paraphrase of Claim 1.
HARD LIMIT â€” 150 WORDS MAXIMUM. This is non-negotiable. Steps you MUST follow:
1. Write your draft paraphrase of Claim 1.
2. Count every single word in it.
3. If the count is above 150, delete words, shorten phrases, and remove adjectives until you reach 150 words or fewer.
4. Output ONLY the final trimmed version â€” do NOT explain or add commentary.
Write in the third person. Bold all component reference numbers (e.g. **(100)**, **(101)**). Apply article rules fresh.`,

    field: `${base}Generate the Technical Field of the Invention section. Write exactly two sentences:
First sentence: Start with 'The present invention relates to the field of...' â€” state the broad technical domain.
Second sentence: Start with 'More specifically, the invention relates to...' â€” state the specific aspect of this invention.
Do NOT include any reference numbers or bold component names in this section. Apply article rules for this section.`,

    background: `${base}Generate the Background of the Invention section as exactly 2 paragraphs, each 80â€“120 words:
Paragraph 1: Describe the broad problem in the general technology domain. Do NOT mention this specific invention. Write as general prior art context.
Paragraph 2: Narrow in on the specific technical problem that this invention addresses. Reference the limitations of existing approaches. The final sentence MUST start with: 'Hence there is a need for...' and must end with a broad characterization of what is needed.
Bold all reference numbers. Apply article rules for this section.`,

    objectives: `${base}Generate the Object of the Invention section with 4â€“5 objectives. CRITICAL FORMAT: each objective MUST be its own separate paragraph (separated by a blank line). Do NOT combine objectives into a single paragraph.
First objective paragraph MUST start: 'A principal objective of the invention is to develop a...'
Each subsequent objective MUST be a new paragraph starting with: 'Another objective of the invention is to...'
Each objective is one complete sentence of 1â€“2 lines. Objectives must be general and functional, not tied to specific numerical values.
Final paragraph (its own separate paragraph): 'These and other objects and characteristics of the present invention will become apparent from the further disclosure to be made in the detailed description given below.'
CRITICAL â€” DO NOT include ANY reference numbers (e.g. (100), (101), (200)) or bolded component names with numbers anywhere in this section. This section must contain only plain descriptive text with NO parenthetical numbers whatsoever. Apply article rules for this section.`,

    summary: `${base}Generate the Summary of the Invention section as a comprehensive paraphrase of all claims (Claims 1 through 8). Ensure every component, method step, and numerical limitation defined in the claims is paraphrased in a clear, narrative style.
CRITICAL — DO NOT include ANY reference numbers (e.g. (100), (101), (200)) or bolded parenthetical numbers anywhere in this section. Write component names as plain descriptive text only, with NO parenthetical numbers whatsoever. Do not omit any claim elements. Apply article rules fresh for this section.`,

    detailed_description: `${base}Generate the Detailed Description of the Invention. This is the most critical and longest section. Apply article tracking FRESH starting at Part 1 and carry it CONTINUOUSLY through all 4 parts â€” all 4 parts form ONE section.

CRITICAL NAMING RULE â€” MANDATORY, NO EXCEPTIONS:
FORBIDDEN: 'The system', 'the system', 'this system', 'said system' used alone without the full name.
REQUIRED: ALWAYS write the complete descriptive name of the invention every single time. Examples:
  âœ“ CORRECT: 'The adaptive prefix-bucket hash table system **(100)**'
  âœ“ CORRECT: 'The polymer nanocomposite **(100)**'
  âœ“ CORRECT: 'The ceramic-coated carbon nanotube reinforced thermoplastic polymer nanocomposite **(100)**'
  âœ— WRONG: 'The system performs...', 'The system includes...', 'the system is...'
If you write 'The system' anywhere without the full preceding name, your output is incorrect.

DO NOT write any opening paragraph â€” the fixed opening is already added separately by the document generator.

You MUST use structural headings like "PART 1 â€” DEFINITIONS", "PART 2 â€” COMPONENTS", "PART 3 â€” METHOD STEPS", and "PART 4 â€” FIGURES" to separate the parts of the description in your output. This is crucial for structural formatting.

PART 1 â€” DEFINITIONS (10 definitions):
For each of the 10 most important technical terms from the claims, write one definition paragraph of 60â€“90 words.
Format: 'As used in this application, the term "[TERM]" refers to...' or 'The term "[TERM]" refers to...'
Each definition must: (a) give only the general meaning of the term in plain technical language, and (b) explain its broad role in this type of invention. Do NOT include specific numerical values, parameter ranges, or component-specific implementation details â€” those belong in the figure descriptions.
CRITICAL u{2014} DO NOT include ANY reference numbers (e.g. (100), (101), (200)) or bolded parenthetical numbers in PART 1. All definitions must be written in plain descriptive language with NO parenthetical numbers whatsoever.
First mention of any noun â†’ use "a"/"an". Repeat mentions â†’ use "the".

PART 2 â€” COMPONENTS:
You MUST output Part 2 in this EXACT ORDER â€” do not deviate. This part is MANDATORY and must ALWAYS be written regardless of patent type:

[STEP 1] Write one standalone paragraph â€” a verbatim copy of the Brief Description of Drawings sentence for the figure that shows the system hardware components. Copy it word-for-word with all reference numbers bolded. This paragraph comes FIRST before any component description.
Example: 'FIG. 1B illustrates an internal hardware configuration of a device within the system **(100)**, depicting a processor **(109)**, a memory **(111)**, and a communication interface **(113)**, according to one embodiment of the invention.'

[STEP 2] For EACH main component listed in the claims, write exactly ONE separate detailed paragraph (150â€“200 words). Each paragraph MUST:
â€” Start directly with the component name and its bolded reference number, e.g.: 'The processor **(109)** is...' or 'The memory **(111)** stores...'
â€” NEVER repeat or re-copy the figure description sentence inside a component paragraph
â€” Cover: (a) what this component is in its broad technical field, (b) its specific parameters/values from the claims, (c) its role in the invention, (d) how it connects to other components, (e) why its parameter values matter for the technical effect
â€” Be one continuous paragraph with no internal paragraph breaks
Bold all reference numbers. Continue article tracking from Part 1.

PART 3 â€” METHOD STEPS:
Write 5 method steps using step numbers 201, 203, 205, 207, 209.

You MUST output Part 3 in this EXACT ORDER â€” do not deviate:

[STEP 1] Write one standalone paragraph â€” a verbatim copy of the Brief Description of Drawings sentence for the flowchart figure illustrating the method steps. Copy it word-for-word with all reference numbers bolded. This paragraph comes FIRST before any step description.
Example: 'FIG. 2 illustrates a flowchart depicting a method for prefix-based indexing and membership query routing, according to some example embodiments.'

[STEP 2] For EACH of the 5 method steps, write exactly TWO consecutive paragraphs:
PARAGRAPH A (80â€“120 words): Start exactly with 'At step **(201)**, the method includes [ING verb]...' â€” describe what happens, components involved, quantifiable parameters. Do NOT copy the flowchart figure sentence here.
PARAGRAPH B (80â€“120 words): Start exactly with 'According to an example embodiment, [ING verb]...' â€” describe HOW this step is performed, process conditions, tools, parameter values.
Steps must flow logically. Continue article tracking from previous parts.

CLAIMS COVERAGE CHECK â€” MANDATORY:
After completing all 4 parts, review every claim element, component name, and numerical parameter value from claims 1 through 8 in the INVENTION CONTEXT. For any claim element, component, or numerical value that was NOT explicitly mentioned in Parts 1â€“3 above, add a sentence to the most relevant existing paragraph to include it. Do NOT output any bracketed note or summary about claims coverage â€” just silently ensure completeness.

PART 4 â€” FIGURES:
IMPORTANT: Part 4 is MANDATORY and must ALWAYS be written. Do NOT skip Part 4.
Write figure descriptions for ALL figures from the Brief Description of Drawings using the exact 3-paragraph structure below. Even if a figure's components were mentioned in Parts 2 or 3, you must still write the full 3-paragraph entry for it in Part 4.

GROUPED FIGURES (e.g. FIG. 4A, FIG. 4B, FIG. 4C): Treat EACH sub-figure as a completely SEPARATE entry with its own 3 paragraphs. Never merge two sub-figures into one paragraph.

For EACH figure (or sub-figure), write EXACTLY THREE paragraphs in this order:

PARAGRAPH 1 (verbatim Brief Description copy):
Copy the EXACT sentence from the Brief Description of Drawings for that figure â€” word for word, with reference numbers bolded. Do not paraphrase or alter it.
Example: 'FIG. 1B illustrates an internal hardware configuration of a device within the system **(100)**, depicting a processor **(109)**, a memory **(111)**, and a communication interface **(113)**, according to one embodiment of the invention.'

PARAGRAPH 2 (extended component walkthrough):
Repeat the SAME Brief Description sentence from Paragraph 1 verbatim (with reference numbers bolded), then IMMEDIATELY continue in the SAME paragraph (no new paragraph) with a detailed walkthrough of each component/element visible in the figure â€” what each does, how it operates, what data it handles, and how it connects to the other components. This is ONE continuous paragraph.
Example: 'FIG. 1B illustrates an internal hardware configuration of a device within the system **(100)**, depicting a processor **(109)**, a memory **(111)**, and a communication interface **(113)**, according to one embodiment of the invention. The processor **(109)** executes [detailed role]. The memory **(111)** stores [detailed role]. The communication interface **(113)** facilitates [detailed role].'

PARAGRAPH 3 (quantifiable 4-step analysis â€” "As illustrated in FIG. X..."):
This MUST be ONE single continuous paragraph â€” do NOT split into sub-paragraphs or bullet points. Follow this strict 4-step structure:
  STEP 1: Start with 'As illustrated in FIG. X, the quantifiable technical effect achieved is [specific measurable outcome with value and unit], which is the direct and necessary consequence of the interaction among [list all components with bolded reference numbers].'
  STEP 2: For each component: name + bolded reference number + configured parameter value/range + specific function + quantifiable contribution to the technical effect. Write sequentially as one block.
  STEP 3: For EACH critical parameter: 'If [parameter] exceeds [upper limit], then [specific degradation] occurs, reducing performance toward [quantifiable worse value]. If [parameter] falls below [lower limit], then [specific failure mode] occurs, degrading [metric] to [quantifiable worse value].'
  STEP 4: 'The baseline system, which lacks [list of components] operating within their specified parameter ranges, produces [worse quantifiable value], confirming that the improvement is causatively attributable to the claimed combination.'
Continue article tracking from previous parts.

${isMethodOnly ? 'NOTE: Method-only patent â€” all 4 Parts are compulsory. Part 2 must describe each constituent element or functional module referenced in the claims as components.' : ''}
${isComposition ? 'NOTE: Method & Composition â€” Part 2 must describe each composition/material component in full detail.' : ''}

END with this mandatory fixed closing paragraph (copy verbatim):
"The many features and advantages of the invention are apparent from the detailed specification, and thus, it is intended by the appended claims to cover all such features and advantages of the invention which fall within the true spirit and scope of the invention. Further, since numerous modifications and variations will readily occur to those skilled in the art, it is not desired to limit the invention to the exact construction and operation illustrated and described, and accordingly, all suitable modifications and equivalents may be resorted to, falling within the scope of the invention."`,

    brief_description: `${base}Generate the Brief Description of Drawings section.
Apply article usage rules FRESH for this section only â€” restart noun tracking from scratch.
Article rules: first mention of each noun/component in THIS section â†’ use "a" or "an"; every subsequent mention in this section â†’ use "the".

CRITICAL SUB-FIGURE RULE â€” MANDATORY, NO EXCEPTIONS:
If the drawings contain sub-figures such as FIG. 1A, FIG. 1B, FIG. 2A, FIG. 2B, FIG. 4A, FIG. 4B, etc., then EACH sub-figure MUST have its OWN completely separate sentence entry.
- FORBIDDEN: Merging sub-figures into one entry (e.g., "FIG. 1 illustrates FIG. 1A and FIG. 1B..." is WRONG).
- FORBIDDEN: Grouping sub-figures under a parent label (e.g., starting with "**FIG. 1**" when the actual figures are 1A and 1B).
- REQUIRED: Write "**FIG. 1A** illustrates..." as one entry, then "**FIG. 1B** illustrates..." as a completely separate entry on its own line.
- Every sub-figure label that appears in the drawings list gets its own dedicated bold label and sentence. Never skip a sub-figure or combine two sub-figures into one line.

For EACH figure (or sub-figure) in the drawings (use the exact figures listed in the INVENTION CONTEXT):
Write ONE sentence per figure entry. The sentence must:
- Start with the figure identifier in bold: e.g. **FIG. 1A** illustrates... or **FIG. 2** illustrates...
- Describe what that specific figure/sub-figure visually shows in one clear concise sentence.
- CRITICAL u{2014} DO NOT include ANY reference numbers (e.g. (100), (101), (200)) or bolded parenthetical numbers in the figure descriptions. Write component names as plain descriptive text only, with NO parenthetical numbers whatsoever.
- Apply article rules: first time a component is introduced in this section â†’ "a"; subsequent mentions â†’ "the".
- End with: "according to some example embodiments." or "according to one embodiment of the invention."
- Must NOT exceed 3 lines per entry.

Format each figure on its OWN LINE, one sentence each:
**FIG. 1A** illustrates [what FIG. 1A shows, describing components by name without reference numbers], according to some example embodiments.
**FIG. 1B** illustrates [what FIG. 1B shows, describing components by name without reference numbers], according to some example embodiments.
**FIG. 2** illustrates [what FIG. 2 shows, describing components by name without reference numbers], according to some example embodiments.
**FIG. 4A** illustrates [what FIG. 4A shows, describing components by name without reference numbers], according to some example embodiments.
**FIG. 4B** illustrates [what FIG. 4B shows, describing components by name without reference numbers], according to some example embodiments.
[etc. â€” one entry per figure or sub-figure, never merged]

Also include TWO mandatory static sentences at the start (before listing figures), copy verbatim:
"The foregoing and other features of embodiments will become more apparent from the following detailed description of embodiments when read in conjunction with the accompanying drawings. In the drawings, like reference numerals refer to like elements."
Then the next sentence, also verbatim:
"In the following description, for the purposes of explanation, numerous specific details are set forth to provide a thorough understanding of the embodiments of the invention. It is apparent, however, to one skilled in the art that the embodiments of the invention may be practiced without these specific details or with an equivalent arrangement. In other instances, well-known structures and devices are shown in block diagram form to avoid unnecessarily obscuring the embodiments of the invention."
Then list all figures after these two sentences.`,
  };
  return prompts[sectionId] || `${base}Generate the ${sectionId} section.`;
}

async function callOllama(prompt, sectionId) {
  const url = OLLAMA_URL.replace('/v1/completions', '/api/generate');
  const numPredict = sectionId === 'detailed_description' ? 1000 : 350;
  
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: OLLAMA_MODEL,
      prompt: prompt,
      stream: false,
      options: {
        temperature: 0.2,
        num_predict: numPredict,
        num_ctx: 3072 // optimize memory footprint & computation time
      }
    }),
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Ollama error ${res.status}: ${errText}`);
  }

  const data = await res.json();
  return data.response || '';
}

async function callGemini(prompt, apiKey, onStatus) {
  // Only models actually available on the free tier with this key type
  const configs = [
    { version: 'v1', model: 'gemini-2.5-flash' },
    { version: 'v1beta', model: 'gemini-2.5-flash' },
    { version: 'v1', model: 'gemini-2.5-flash-preview-05-20' },
    { version: 'v1beta', model: 'gemini-2.5-flash-preview-05-20' },
  ];

  const errors = [];

  for (const config of configs) {
    let retries = 3;
    while (retries >= 0) {
      try {
        const url = `https://generativelanguage.googleapis.com/${config.version}/models/${config.model}:generateContent?key=${apiKey}`;
        const res = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { temperature: 0.2 }
          })
        });

        if (res.ok) {
          const data = await res.json();
          if (data.promptFeedback?.blockReason) {
            throw new Error(`Prompt blocked by Gemini safety/policy. Reason: ${data.promptFeedback.blockReason}`);
          }
          const candidate = data.candidates?.[0];
          if (candidate) {
            const text = candidate.content?.parts?.[0]?.text;
            if (text) return text;
            const finishReason = candidate.finishReason;
            if (finishReason && finishReason !== 'STOP') {
              throw new Error(`Generation blocked by Gemini. Finish reason: ${finishReason}`);
            }
          }
        }

        const errBody = await res.json().catch(() => ({}));
        const msg = errBody.error?.message || `HTTP ${res.status}`;
        const status = res.status;

        if ((status === 429 || status === 503 || status === 500) && retries > 0) {
          // Parse "Please retry in X.XXs" from the error message
          const match = msg.match(/retry in (\d+(?:\.\d+)?)s/);
          const waitMs = match ? Math.ceil(parseFloat(match[1])) * 1000 + 1000 : 15000;
          const waitSec = Math.ceil(waitMs / 1000);
          console.warn(`Gemini rate-limited. Waiting ${waitSec}s before retrying...`);
          if (onStatus) onStatus(`Rate limited by Gemini. Auto-retrying in ${waitSec}s...`);
          await new Promise(r => setTimeout(r, waitMs));
          retries--;
          continue;
        }

        // 404 means this model isn't available â€” don't retry, just move on
        errors.push(`${config.model} (${config.version}): [${status}] ${msg}`);
        break;
      } catch (e) {
        if (e.message.includes('blocked by Gemini') || e.message.includes('Finish reason')) throw e;
        if (retries > 0) {
          await new Promise(r => setTimeout(r, 5000));
          retries--;
          continue;
        }
        errors.push(`${config.model} (${config.version}): ${e.message}`);
        break;
      }
    }
  }

  throw new Error(`Gemini generation failed. Detailed log:\n` + errors.map(e => `â€¢ ${e}`).join('\n'));
}

async function callGroq(prompt, apiKey) {
  const url = 'https://api.groq.com/openai/v1/chat/completions';
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: 'llama-3.3-70b-versatile',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.2
    })
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Groq error HTTP ${res.status}`);
  }

  const data = await res.json();
  return data.choices?.[0]?.message?.content || '';
}

async function callClaude(prompt, apiKey) {
  // Routes through Vite dev-server proxy (/api/claude â†’ https://api.anthropic.com)
  // This avoids browser CORS restrictions on the Anthropic API.
  const res = await fetch('/api/claude/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: 'claude-3-5-haiku-20241022',
      max_tokens: 4096,
      messages: [{ role: 'user', content: prompt }],
    })
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error?.message || `Claude error HTTP ${res.status}`);
  }

  const data = await res.json();
  return data.content?.[0]?.text || '';
}

function loadScript(url, globalName) {
  return new Promise((res, rej) => {
    if (window[globalName]) { res(); return; }
    const s = document.createElement('script');
    s.src = url;
    s.onload = res;
    s.onerror = () => rej(new Error(`Failed to load ${globalName} library`));
    document.head.appendChild(s);
  });
}

async function extractDocxText(arrayBuffer) {
  await loadScript('https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.8.0/mammoth.browser.min.js', 'mammoth');
  const result = await window.mammoth.extractRawText({ arrayBuffer });
  return result.value || '';
}

async function extractPptxText(arrayBuffer) {
  await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js', 'JSZip');
  const zip = await window.JSZip.loadAsync(arrayBuffer);
  let text = '';
  const files = Object.keys(zip.files).filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'));
  files.sort((a, b) => {
    const numA = parseInt(a.replace(/[^0-9]/g, ''), 10);
    const numB = parseInt(b.replace(/[^0-9]/g, ''), 10);
    return numA - numB;
  });
  for (const file of files) {
    const content = await zip.files[file].async('text');
    const matches = content.match(/<a:t>([^<]*)<\/a:t>/g) || [];
    const slideText = matches.map(m => m.replace(/<\/?a:t>/g, '')).join(' ');
    if (slideText.trim()) {
      text += `\nSlide:\n${slideText}\n`;
    }
  }
  return text;
}

async function extractFileText(file) {
  if (!file) return '';
  const ext = file.name.split('.').pop().toLowerCase();
  
  const arrayBuffer = await new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(new Error('Read failed'));
    r.readAsArrayBuffer(file);
  });

  if (ext === 'docx') {
    return await extractDocxText(arrayBuffer);
  } else if (ext === 'pptx') {
    return await extractPptxText(arrayBuffer);
  } else {
    return await new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(r.result);
      r.onerror = () => rej(new Error('Read failed'));
      r.readAsText(file);
    });
  }
}

async function downloadDocx(results, sections) {
  const { Document, Packer, Paragraph, TextRun, AlignmentType } = docx;

  // Parse markdown bold and FIG/TABLE references into TextRun array
  function parseRuns(text) {
    const runs = [];
    const re = /\*\*([^*]+)\*\*|(FIGs?\.\s*\d+[A-Za-z]?(?:\s*[-â€“]\s*\d*[A-Za-z]?)*|TABLE\.\s*\d+[A-Za-z]?)/g;
    let last = 0;
    let m;
    while ((m = re.exec(text)) !== null) {
      if (m.index > last) runs.push(new TextRun({ text: text.slice(last, m.index), font: 'Arial', size: 24 }));
      runs.push(new TextRun({ text: m[1] || m[2], bold: true, font: 'Arial', size: 24 }));
      last = m.index + m[0].length;
    }
    if (last < text.length) runs.push(new TextRun({ text: text.slice(last), font: 'Arial', size: 24 }));
    return runs;
  }

  // Manual paragraph counter â€” [0001], [0002]...
  let paraNum = 0;
  function nextNum() {
    paraNum++;
    return '[' + String(paraNum).padStart(4, '0') + ']';
  }

  // Numbered paragraph: bold [0001] + space + text, no indent â€” both ruler markers at left margin (MI17 style)
  function toParagraphs(text) {
    return text.split('\n').filter(l => l.trim()).map(line => {
      return new Paragraph({
        children: [
          new TextRun({ text: nextNum() + ' ', bold: true, font: 'Arial', size: 24 }),
          ...parseRuns(line.trim()),
        ],
        spacing: { after: 200, line: 480, lineRule: 'auto' },
        alignment: AlignmentType.JUSTIFIED,
      });
    });
  }

  // Single static paragraph with bold [0001] number â€” no indent, MI17 style
  function staticParaNum(text) {
    return new Paragraph({
      children: [
        new TextRun({ text: nextNum() + ' ', bold: true, font: 'Arial', size: 24 }),
        ...parseRuns(text),
      ],
      spacing: { after: 200, line: 480, lineRule: 'auto' },
      alignment: AlignmentType.JUSTIFIED,
    });
  }

  // Header/label paragraph WITHOUT numbering
  function headerPara(text, bold = true, centered = false, sizePt = 12) {
    return new Paragraph({
      children: [new TextRun({ text, bold, font: 'Arial', size: sizePt * 2 })],
      spacing: { before: 320, after: 160 },
      alignment: centered ? AlignmentType.CENTER : AlignmentType.LEFT,
    });
  }

  function fieldRow(label, value) {
    return new Paragraph({
      children: [
        new TextRun({ text: `${label}: `, bold: true, font: 'Arial', size: 24 }),
        new TextRun({ text: value, font: 'Arial', size: 24 }),
      ],
      spacing: { after: 80 },
    });
  }

  const titleText = results.title || 'PATENT TITLE';
  const children = [];

  // ===== FORM 2 FIRST PAGE â€” BORDERED TABLE (matches reference image) =====
  const { Table, TableRow, TableCell, WidthType, BorderStyle } = docx;

  const cellBorder = {
    top:    { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    left:   { style: BorderStyle.SINGLE, size: 4, color: '000000' },
    right:  { style: BorderStyle.SINGLE, size: 4, color: '000000' },
  };
  const cellMargins = { top: 140, bottom: 140, left: 160, right: 160 };

  function cp(text, bold = false, centered = false, size = 24, spaceAfter = 80) {
    return new Paragraph({
      children: [new TextRun({ text, bold, font: 'Arial', size })],
      alignment: centered ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { after: spaceAfter },
    });
  }

  function makeRow(cellChildren) {
    return new TableRow({
      children: [
        new TableCell({
          children: cellChildren,
          borders: cellBorder,
          margins: cellMargins,
          width: { size: 100, type: WidthType.PERCENTAGE },
        }),
      ],
    });
  }

  const form2Table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      // Row 1: FORM 2 header block
      makeRow([
        cp('FORM 2', true, true, 24, 100),
        cp('THE PATENTS ACT, 1970', true, true, 24, 100),
        cp('(39 of 1970)', false, true, 24, 100),
        cp('AND', true, true, 24, 100),
        cp('The Patents Rules, 2003', true, true, 24, 100),
        cp('COMPLETE SPECIFICATION', true, true, 24, 100),
        cp('(See Section 10 and Rule 13)', false, true, 24, 0),
      ]),
      // Row 2: Title
      makeRow([
        cp('1. TITLE OF THE INVENTION:', true, false, 24, 120),
        new Paragraph({
          children: [new TextRun({ text: titleText, bold: true, font: 'Arial', size: 24 })],
          spacing: { after: 0 },
        }),
      ]),
      // Row 3: Applicant
      makeRow([
        cp('2. APPLICANT:', true, false, 24, 120),
        cp('(a) NAME:', false, false, 24, 120),
        cp('(b) NATIONALITY:', false, false, 24, 120),
        cp('(c) ADDRESS:', false, false, 24, 120),
        new Paragraph({ spacing: { after: 160 } }),
      ]),
      // Row 4: Preamble
      makeRow([
        cp('3. PREAMBLE TO THE DESCRIPTION:', true, false, 24, 120),
        new Paragraph({
          children: [new TextRun({ text: 'The following specification particularly describes the invention and the manner in which it is to be performed.', font: 'Arial', size: 24 })],
          spacing: { after: 160 },
          alignment: AlignmentType.JUSTIFIED,
        }),
        new Paragraph({ spacing: { after: 100 } }),
      ]),
      // Row 5: empty (as in reference image)
      makeRow([new Paragraph({ spacing: { after: 200 } })]),
    ],
  });

  children.push(form2Table);
  children.push(new Paragraph({ spacing: { after: 200 } }));

  // ===== DESCRIPTION HEADER =====
  children.push(
    new Paragraph({ children: [new TextRun({ text: 'DESCRIPTION OF THE INVENTION:', bold: true, font: 'Arial', size: 26 })], spacing: { before: 240, after: 240 } }),
  );


  // ===== DOCUMENT ORDER: field â†’ background â†’ objectives â†’ summary â†’ brief_description â†’ detailed_description â†’ abstract =====
  const DOC_ORDER = ['field', 'background', 'objectives', 'summary', 'brief_description', 'detailed_description', 'abstract'];

  const SECTION_DISPLAY_HEADINGS = {
    field: 'Technical Field of the Invention',
    background: 'Background',
    objectives: 'Object of the Invention',
    summary: 'Summary of the Invention',
    brief_description: 'Brief Description of Drawings',
    detailed_description: 'Detailed Description of the Invention',
    abstract: 'Abstract',
  };

  const SUMMARY_PREFIX = 'To enable the present disclosure and related ends, at least one aspect comprises the feature herein after completely described and particularly and/or specifically pointed out in the specification in the section of claims. The following drawings and descriptions outlined in detail enable certain example features of at least one aspect. Described features are indicative, however, of but a few of the many ways in which the following principles of various aspects may be employed, and the description is intended to include all such aspects and their equivalents.';
  const SUMMARY_SUFFIX = 'The above summary is descriptive and exemplary only and is not intended to be in any way restricting. In addition to the descriptive aspects, embodiments, and features described in the above summary, further features and embodiments will become apparent by reference to the accompanying drawings and the following detailed description.';
  const DD_OPEN = 'Reference will now be made in detail to the description of the present subject matter, one or more examples of which are shown in the figures. Each example is provided to explain the subject matter and not a limitation. Various changes and modifications obvious to one skilled in the art to which the invention pertains are deemed to be within the spirit, scope and contemplation of the invention.';

  for (const id of DOC_ORDER) {
    if (!results[id]) continue;

    // Section heading â€” centered for Abstract
    children.push(headerPara(SECTION_DISPLAY_HEADINGS[id], true, id === 'abstract', 12));

    if (id === 'abstract') {
      // Title â€” centered, bold, with 1440 left+right indent (MI17 style), no para number
      children.push(new Paragraph({
        children: [new TextRun({ text: titleText, bold: true, font: 'Arial', size: 24 })],
        alignment: AlignmentType.CENTER,
        indent: { left: 1440, right: 1440 },
        spacing: { after: 120, line: 480, lineRule: 'auto' },
      }));
      // Abstract body â€” justified (not centered), 1440 left+right indent (MI17 style), no para number
      results[id].split('\n').filter(l => l.trim()).forEach(line => {
        children.push(new Paragraph({
          children: parseRuns(line.trim()),
          alignment: AlignmentType.JUSTIFIED,
          indent: { left: 1440, right: 1440 },
          spacing: { after: 0, line: 480, lineRule: 'auto' },
        }));
      });
      // (FIG. 1) â€” bold, justified, 1440 left+right indent (MI17 style), no para number
      children.push(new Paragraph({
        children: [new TextRun({ text: '(FIG. 1)', bold: true, font: 'Arial', size: 24 })],
        alignment: AlignmentType.JUSTIFIED,
        indent: { left: 1440, right: 1440 },
        spacing: { before: 0, after: 200, line: 480, lineRule: 'auto' },
      }));
    } else if (id === 'summary') {
      children.push(staticParaNum(SUMMARY_PREFIX));
      children.push(new Paragraph({ spacing: { after: 80 } }));
      children.push(...toParagraphs(results[id]));
      children.push(new Paragraph({ spacing: { after: 80 } }));
      children.push(staticParaNum(SUMMARY_SUFFIX));
    } else if (id === 'detailed_description') {
      children.push(staticParaNum(DD_OPEN));
      children.push(new Paragraph({ spacing: { after: 80 } }));
      children.push(...toParagraphs(results[id]));
    } else {
      children.push(...toParagraphs(results[id]));
    }
  }

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 24 } } },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children
    }]
  });

  const blob = await Packer.toBlob(doc);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Patent_Complete_Specification.docx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}


/* â”€â”€ File Upload Box â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
function FileBoxMain({ label, sublabel, accept, file, onChange, required, icon }) {
  const ref = useRef();
  return (
    <div
      className={`file-box-main${file ? ' has-file' : ''}`}
      onClick={() => ref.current?.click()}
    >
      <input ref={ref} type="file" accept={accept} style={{ display: 'none' }} onChange={e => onChange(e.target.files[0])} />
      <span className={`material-icons-round fbm-icon`}>{file ? 'check_circle' : icon}</span>
      <div className="fbm-label">{label}</div>
      <div className="fbm-sub">{file ? file.name : sublabel}</div>
      <span className={`fbm-req ${required ? 'required' : 'optional'}`}>{required ? 'Required' : 'Optional'}</span>
    </div>
  );
}

/* â”€â”€ Section icon map â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
const SEC_ICONS = {
  title: 'title', field: 'device_hub', background: 'history_edu',
  objectives: 'track_changes', summary: 'summarize',
  brief_description: 'image', detailed_description: 'article', abstract: 'subject',
};

/* ── Extract full invention system name from title / claims / disclosure ── */
function extractSystemName(titleStr, claimsText, disclosureText) {
  let title = titleStr || '';
  if (!title && claimsText) {
    const lines = claimsText.split('\n');
    for (const line of lines) {
      const u = line.toUpperCase().trim();
      if (u.startsWith('A METHOD') || u.startsWith('A SYSTEM') || u.startsWith('A COMPOSITION')) {
        title = line.trim();
        break;
      }
    }
  }
  if (!title && disclosureText) {
    const lines = disclosureText.split('\n');
    for (const line of lines) {
      const u = line.toUpperCase().trim();
      if (u.startsWith('A METHOD') || u.startsWith('A SYSTEM') || u.startsWith('A COMPOSITION')) {
        title = line.trim();
        break;
      }
    }
  }
  if (!title) return null;
  const forIdx = title.toUpperCase().indexOf(' FOR ');
  let name = '';
  if (forIdx === -1) {
    name = title.trim();
    name = name.replace(/^(A\s+)?(METHOD\s+AND\s+COMPOSITION|SYSTEM\s+AND\s+METHOD|METHOD|SYSTEM|COMPOSITION)\s+(FOR|OF)\s+/i, '');
  } else {
    name = title.slice(forIdx + 5).trim();
  }
  name = name.replace(/^(DEVELOPING|IMPLEMENTING|CREATING|DESIGNING|PRODUCING|MANUFACTURING|FABRICATING|BUILDING|GENERATING|PROVIDING|MAKING|FORMING|PROCESSING|USING|UTILIZING|OPTIMIZING|IMPROVING|ENHANCING|ENABLING|ESTABLISHING|DETECTING|MANAGING|CONTROLLING|MONITORING|ANALYZING|PERFORMING|EXECUTING|CONFIGURING|PREPARING|APPLYING)\s+/i, '');
  if (!name) return null;
  return name.charAt(0).toUpperCase() + name.slice(1).toLowerCase();
}

function PatentApp() {

  // ── App state ─────────────────────────────────────────────────────
  const [claims, setClaims]       = useState(null);
  const [drawings, setDrawings]   = useState(null);
  const [disclosure, setDisclosure] = useState(null);
  const [patentType, setPatentType] = useState('method_composition');
  const [selected, setSelected]   = useState(ALL_SECTIONS.map(s => s.id));
  const [generating, setGenerating] = useState(false);
  const [downloading, setDownloading] = useState(false);
  const [progress, setProgress]   = useState({ current: 0, total: 0, label: '' });
  const [results, setResults]     = useState({});
  const [done, setDone]           = useState(false);
  const [error, setError]         = useState('');

  const [provider, setProvider] = useState(() => localStorage.getItem('patent_api_provider') || 'ollama');
  const [apiKey, setApiKey]     = useState(() => localStorage.getItem('patent_api_key') || '');

  const handleProviderChange = e => {
    setProvider(e.target.value);
    localStorage.setItem('patent_api_provider', e.target.value);
  };
  const handleApiKeyChange = e => {
    setApiKey(e.target.value);
    localStorage.setItem('patent_api_key', e.target.value);
  };

  const toggleSection = id => setSelected(p => p.includes(id) ? p.filter(x => x !== id) : [...p, id]);
  const canGenerate = claims && disclosure && selected.length > 0 && !generating;
  const pct = progress.total > 0 ? Math.round((progress.current / progress.total) * 100) : 0;
  const doneCount = Object.keys(results).length;

  const handleGenerate = async () => {
    setError(''); setDone(false); setResults({}); setGenerating(true);
    try {
      setProgress({ current: 0, total: 1, label: 'Reading uploaded files...' });
      const [claimsText, drawingsText, disclosureText] = await Promise.all([
        extractFileText(claims), extractFileText(drawings), extractFileText(disclosure),
      ]);
      const ctx = `CLAIMS:\n${claimsText}\n\nDRAWINGS:\n${drawingsText}\n\nINVENTION DISCLOSURE:\n${disclosureText}`;
      const secs = ALL_SECTIONS.filter(s => selected.includes(s.id));
      const res = {};
      for (let i = 0; i < secs.length; i++) {
        if (i > 0 && provider === 'gemini') {
          for (let s = 7; s > 0; s--) {
            setProgress({ current: i + 1, total: secs.length, label: `Pacing: waiting ${s}s before next section...` });
            await new Promise(r => setTimeout(r, 1000));
          }
        }
        setProgress({ current: i + 1, total: secs.length, label: `Drafting: ${secs[i].label}...` });
        let out = '';
        const prompt = buildPrompt(secs[i].id, patentType, ctx, !!drawings);
        if (provider === 'gemini') {
          if (!apiKey) throw new Error('Please enter your Gemini API key in settings.');
          out = await callGemini(prompt, apiKey, msg => setProgress(p => ({ ...p, label: msg })));
        } else if (provider === 'groq') {
          if (!apiKey) throw new Error('Please enter your Groq API key in settings.');
          out = await callGroq(prompt, apiKey);
        } else if (provider === 'claude') {
          if (!apiKey) throw new Error('Please enter your Claude API key in settings.');
          out = await callClaude(prompt, apiKey);
        } else {
          out = await callOllama(prompt, secs[i].id);
        }
        if (secs[i].id === 'abstract') {
          const words = out.trim().split(/\s+/);
          if (words.length > 150) {
            const trimmed = words.slice(0, 150).join(' ');
            const lastPeriod = Math.max(trimmed.lastIndexOf('.'), trimmed.lastIndexOf('!'), trimmed.lastIndexOf('?'));
            out = lastPeriod > 0 ? trimmed.slice(0, lastPeriod + 1) : trimmed;
          }
        }
        const sysName = extractSystemName(res.title || results.title || '', claimsText, disclosureText);
        if (sysName) {
          const systemRegex = /\b([Tt]he|[Tt]his)\s+system(?:\s*(\*\*\(\s*(\d+)\s*\)\*\*|\*\*\s*(\d+)\s*\*\*|\(\s*(\d+)\s*\)|\b(\d+)\b))?/g;
          out = out.replace(systemRegex, (match, prefix, suffix, g3, g4, g5, g6) => {
            const digits = g3 || g4 || g5 || g6 || '100';
            return `${prefix} ${sysName} **(${digits})**`;
          });
        }
        if (secs[i].id === 'detailed_description') {
          out = out.split('\n').filter(line => !/^\s*(?:\*+|#+)?\s*(PART\s+\d+|PART\s+[I|V|X]+)\b/i.test(line)).join('\n');
        }
        res[secs[i].id] = out;
        setResults({ ...res });
      }
      setDone(true);
    } catch (e) {
      setError(e.message || 'Something went wrong.');
    } finally {
      setGenerating(false);
    }
  };

  const handleDownload = async () => {
    setDownloading(true);
    try { await downloadDocx(results, selected); }
    catch (e) { setError('Download failed: ' + e.message); }
    finally { setDownloading(false); }
  };

  /* â”€â”€ step dots â”€â”€ */
  const filesReady  = !!(claims && disclosure);
  const sectionsOk  = selected.length > 0;
  const apiOk       = provider === 'ollama' || !!apiKey;

  const PROVIDER_LINKS = {
    gemini: 'https://aistudio.google.com/',
    groq:   'https://console.groq.com/keys',
    claude: 'https://console.anthropic.com/settings/keys',
  };

  return (
    <div id="app-shell">

      {/* â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â• SIDEBAR â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â• */}
      <aside id="sidebar">

        {/* Brand */}
        <div className="sb-brand">
          <div className="sb-logo">
            <div className="sb-logo-icon">
              <span className="material-icons-round" style={{ fontSize: 20 }}>description</span>
            </div>
            <div>
              <div className="sb-logo-text">Patent CS</div>
              <div className="sb-logo-sub">Complete Specification Generator</div>
            </div>
          </div>
          <div className="sb-badge">Form 2 Â· IPO India</div>
        </div>

        {/* API Provider */}
        <div className="sb-section">
          <div className="sb-section-label">API Provider</div>
          <div className="sb-card">
            <div className="sb-card-title">
              <span className="material-icons-round mi">api</span>
              Provider &amp; Key
            </div>
            <label className="g-label">Select Provider</label>
            <select className="g-select" value={provider} onChange={handleProviderChange}>
              <option value="ollama">Local Ollama (CPU)</option>
              <option value="gemini">Gemini API â€” Free Tier âš¡</option>
              <option value="groq">Groq API â€” Llama 3.3 70B</option>
              <option value="claude">Claude Haiku â€” High Quality</option>
            </select>

            {provider !== 'ollama' && (
              <div style={{ marginTop: 10 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 5 }}>
                  <label className="g-label" style={{ margin: 0 }}>API Key</label>
                  <a className="g-link" href={PROVIDER_LINKS[provider]} target="_blank" rel="noreferrer">Get free key â†’</a>
                </div>
                <input
                  className="g-input"
                  type="password"
                  placeholder={`${provider === 'gemini' ? 'Gemini' : provider === 'groq' ? 'Groq' : 'Anthropic'} API key`}
                  value={apiKey}
                  onChange={handleApiKeyChange}
                />
              </div>
            )}

            {/* Status dots */}
            <div style={{ display: 'flex', gap: 6, marginTop: 12, alignItems: 'center' }}>
              <div className={`step-dot ${apiOk ? 'done' : 'active'}`} title="API" />
              <div className={`step-dot ${filesReady ? 'done' : sectionsOk ? 'active' : ''}`} title="Files" />
              <div className={`step-dot ${done ? 'done' : generating ? 'active' : ''}`} title="Generation" />
              <span style={{ fontSize: 10, color: 'var(--grey-500)', marginLeft: 2 }}>
                {done ? 'All done!' : generating ? 'Generatingâ€¦' : apiOk && filesReady ? 'Ready to generate' : 'Setup needed'}
              </span>
            </div>
          </div>
        </div>

        {/* Patent Type */}
        <div className="sb-section">
          <div className="sb-section-label">Patent Type</div>
          <div className="sb-card">
            <div className="sb-card-title">
              <span className="material-icons-round mi">category</span>
              Claim Structure
            </div>
            {PATENT_TYPES.map(t => (
              <label key={t.value} className={`radio-pill${patentType === t.value ? ' active' : ''}`}>
                <input type="radio" name="pt" value={t.value} checked={patentType === t.value} onChange={() => setPatentType(t.value)} />
                <div>
                  <div className="pill-title">{t.label}</div>
                  <div className="pill-sub">{t.prefix}</div>
                </div>
              </label>
            ))}
          </div>
        </div>

        {/* Sections */}
        <div className="sb-section">
          <div className="sb-section-label">Sections</div>
          <div className="sb-card">
            <div className="sb-card-title">
              <span className="material-icons-round mi">checklist</span>
              To Generate
              <button className="sel-all-btn" onClick={() => setSelected(selected.length === ALL_SECTIONS.length ? [] : ALL_SECTIONS.map(s => s.id))}>
                {selected.length === ALL_SECTIONS.length ? 'Deselect all' : 'Select all'}
              </button>
            </div>
            {ALL_SECTIONS.map(sec => {
              const chk = selected.includes(sec.id);
              return (
                <label key={sec.id} className={`chk-pill${chk ? ' active' : ''}`}>
                  <input type="checkbox" checked={chk} onChange={() => toggleSection(sec.id)} />
                  <span className="material-icons-round" style={{ fontSize: 14, color: chk ? 'var(--primary-400)' : 'var(--grey-400)' }}>
                    {SEC_ICONS[sec.id] || 'article'}
                  </span>
                  {sec.label}
                </label>
              );
            })}
          </div>
        </div>

        {/* Version tag */}
        <div style={{ padding: '0 1.25rem', marginTop: 'auto' }}>
          <div style={{ fontSize: 10, color: 'var(--grey-400)', textAlign: 'center', paddingTop: 8 }}>
            Patent CS Generator Â· v2.0 Â· IPO India Form 2
          </div>
        </div>

      </aside>

      {/* â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â• MAIN CONTENT â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â• */}
      <div id="main-content">

        {/* Top bar */}
        <div id="topbar">
          <div className="topbar-title">
            Patent <span>Complete Specification</span> Drafter
          </div>
          {doneCount > 0 && (
            <div className="stat-chip">
              <span className="material-icons-round">check_circle</span>
              {doneCount} / {selected.length} sections drafted
            </div>
          )}
          <div className="topbar-step">
            {done ? 'âœ… Ready to download' : generating ? `â³ Drafting ${pct}%` : 'Upload â†’ Generate â†’ Download'}
          </div>
        </div>

        {/* Body */}
        <div id="content-body">

          {/* Hero banner */}
          <div className="hero-banner">
            <div className="hero-title">AI-Powered <span className="accent">Patent Drafting</span></div>
            <div className="hero-sub">
              Upload your claims, drawings and disclosure. The AI will auto-draft a complete Form 2 specification â€” ready for IPO India filing.
            </div>
            <div className="hero-chips">
              <div className="hero-chip"><span className="material-icons-round mi">bolt</span>Instant drafting</div>
              <div className="hero-chip"><span className="material-icons-round mi">gavel</span>IPO India Form 2</div>
              <div className="hero-chip"><span className="material-icons-round mi">format_quote</span>Article rules enforced</div>
              <div className="hero-chip"><span className="material-icons-round mi">download</span>.docx export</div>
            </div>
          </div>

          {/* Upload card */}
          <div className="m-card">
            <div className="m-card-header">
              <div className="m-card-icon blue"><span className="material-icons-round mi">upload_file</span></div>
              <div>
                <div className="m-card-title">Upload Source Files</div>
                <div className="m-card-sub">Claims and Disclosure are required Â· Drawings optional</div>
              </div>
            </div>
            <div className="file-grid-3">
              <FileBoxMain
                label="Claims"
                sublabel="Upload .docx file"
                accept=".docx"
                file={claims}
                onChange={setClaims}
                required={true}
                icon="gavel"
              />
              <FileBoxMain
                label="Drawings"
                sublabel="Upload .pptx or .docx"
                accept=".pptx,.docx"
                file={drawings}
                onChange={setDrawings}
                required={false}
                icon="image"
              />
              <FileBoxMain
                label="Disclosure"
                sublabel="Upload .docx file"
                accept=".docx"
                file={disclosure}
                onChange={setDisclosure}
                required={true}
                icon="article"
              />
            </div>
          </div>

          {/* Generate button */}
          <button
            id="btn-generate-main"
            className={`btn-generate ${canGenerate ? 'active' : 'disabled'}`}
            disabled={!canGenerate}
            onClick={handleGenerate}
          >
            {generating ? (
              <><span className="spinner" />&nbsp;Generatingâ€¦ {pct}%</>
            ) : (
              <><span className="material-icons-round" style={{ fontSize: 20 }}>auto_awesome</span>Generate Complete Specification</>
            )}
          </button>

          {/* Progress */}
          {generating && (
            <div className="progress-card" style={{ marginTop: '1rem' }}>
              <div className="progress-label">
                <span className="spinner" />
                {progress.label}
              </div>
              <div className="progress-track">
                <div className="progress-bar" style={{ width: `${pct}%` }} />
              </div>
              <div className="progress-counts">{progress.current} of {progress.total} sections complete</div>
            </div>
          )}

          {/* Error */}
          {error && (
            <div className="error-card">
              <span className="material-icons-round">error_outline</span>
              <span>{error}</span>
            </div>
          )}

          {/* Success */}
          {done && (
            <div className="success-banner">
              <span className="material-icons-round">check_circle</span>
              <div>
                <div className="success-text">All {Object.keys(results).length} sections drafted successfully!</div>
                <div className="success-sub">Review the preview below, then download your .docx specification.</div>
              </div>
            </div>
          )}

          {/* Download */}
          {done && (
            <button className="btn-download" onClick={handleDownload} disabled={downloading}>
              <span className="material-icons-round" style={{ fontSize: 20 }}>download</span>
              {downloading ? 'Preparing documentâ€¦' : 'Download as Word (.docx)'}
            </button>
          )}

          {/* Preview */}
          {Object.keys(results).length > 0 && (
            <div className="preview-card" style={{ marginTop: '1.5rem' }}>
              <div className="preview-header">
                <span className="material-icons-round" style={{ color: 'var(--primary-400)', fontSize: 20 }}>preview</span>
                <div className="preview-header-title">Section Preview</div>
                <div className="stat-chip">
                  <span className="material-icons-round">layers</span>
                  {Object.keys(results).length} sections
                </div>
              </div>
              {ALL_SECTIONS.filter(s => results[s.id]).map(sec => (
                <details key={sec.id} className="preview-item">
                  <summary className="preview-summary">
                    <span className="material-icons-round mi-check">task_alt</span>
                    <span className="material-icons-round" style={{ fontSize: 15, color: 'var(--grey-500)' }}>
                      {SEC_ICONS[sec.id] || 'article'}
                    </span>
                    {sec.label}
                    <span style={{ marginLeft: 'auto', fontSize: 11, color: 'var(--grey-500)', fontWeight: 400 }}>
                      {results[sec.id].split(/\s+/).length} words
                    </span>
                  </summary>
                  <div className="preview-content">
                    <pre className="preview-pre">{results[sec.id]}</pre>
                  </div>
                </details>
              ))}
            </div>
          )}

          {/* Bottom spacer */}
          <div style={{ height: '2rem' }} />

        </div>{/* /content-body */}
      </div>{/* /main-content */}
    </div>/* /app-shell */
  );
}

/* ── Auth gate wrapper — keeps all hooks rule-compliant ─────────── */
export default function App() {
  const [authed, setAuthed] = useState(
    () => localStorage.getItem('patent_auth') === '1'
  );

  const handleLogin = () => {
    localStorage.setItem('patent_auth', '1');
    setAuthed(true);
  };

  if (!authed) return <Login onLogin={handleLogin} />;
  return <PatentApp />;
}
