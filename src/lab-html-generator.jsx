import { useState, useCallback, useRef, useEffect } from "react";

function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) return resolve();
    const s = document.createElement("script");
    s.src = src;
    s.onload = resolve;
    s.onerror = reject;
    document.head.appendChild(s);
  });
}

// ─── Parse the DOCX XML to extract lab test fields ────────────────────────────
function parseDocxXml(xmlString, numberingXmlString = "") {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, "application/xml");
  const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const WPS_NS  = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
  const WPG_NS  = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
  const DML_NS  = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const WP_NS   = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

  const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  Array.from(doc.getElementsByTagNameNS(MC_NS, "Fallback")).forEach((fb) => {
    fb.parentNode && fb.parentNode.removeChild(fb);
  });

  
  function fillFromDrawings(data) {
    const inlines = doc.getElementsByTagNameNS(WP_NS, "inline");
    for (const inline of Array.from(inlines)) {
      // Determine total group width so we can compute column ratios
      const grpSpPrs = inline.getElementsByTagNameNS(WPG_NS, "grpSpPr");
      let groupCx = 0;
      if (grpSpPrs.length) {
        const extNode = grpSpPrs[0].getElementsByTagNameNS(DML_NS, "ext")[0];
        if (extNode) groupCx = parseInt(extNode.getAttribute("cx") || "0", 10);
      }
      if (!groupCx) continue;

      // Collect all textbox shapes with their x offset and text content
      const boxes = [];
      Array.from(inline.getElementsByTagNameNS(WPS_NS, "wsp")).forEach((wsp) => {
        const spPr = wsp.getElementsByTagNameNS(WPS_NS, "spPr")[0];
        if (!spPr) return;
        const xfrmList = spPr.getElementsByTagNameNS(DML_NS, "xfrm");
        if (!xfrmList.length) return;
        const offNode = xfrmList[0].getElementsByTagNameNS(DML_NS, "off")[0];
        if (!offNode) return;
        const xOff = parseInt(offNode.getAttribute("x") || "0", 10);

        const txbx = wsp.getElementsByTagNameNS(WPS_NS, "txbx")[0];
        if (!txbx) return;
        let lines = [];
        Array.from(txbx.getElementsByTagNameNS(WORD_NS, "p")).forEach((p) => {
          let pText = "";
          Array.from(p.getElementsByTagNameNS(WORD_NS, "r")).forEach((run) => {
            Array.from(run.childNodes).forEach((child) => {
              if (child.nodeType === 1 && child.localName === "t") pText += child.textContent || "";
            });
          });
          pText = pText.trim();
          if (pText) lines.push(pText);
        });
        if (lines.length) boxes.push({ xOff, lines });
      });
      boxes.sort((a, b) => a.xOff - b.xOff);

      // Assign to columns by x ratio:
      //   0–15%  → Test Name / Method  (leftmost textbox)
      //   15–50% → Result              (patient value — usually blank in template)
      //   50–70% → Units
      //   70–100%→ Bio. Ref. Interval
      boxes.forEach(({ xOff, lines }) => {
        const ratio = xOff / groupCx;
        if (ratio < 0.15) {
          if (!data.testName && lines[0]) data.testName = lines[0];
          if (!data.method   && lines[1]) data.method   = lines[1];
        } else if (ratio < 0.50) {
          if (!data.result && lines[0]) data.result = lines[0];
        } else if (ratio < 0.70) {
          if (!data.units && lines[0]) data.units = lines.join(" ");
        } else {
          if (!data.refInterval && lines[0]) data.refInterval = lines.join(" ");
        }
      });
    }
  }

  const paragraphs = Array.from(doc.getElementsByTagNameNS(WORD_NS, "p"));

  const data = {
    testName: "",
    method: "",
    result: "",
    units: "",
    refInterval: "",
    sections: [],
  };

  // ── Step 1: pull Test Name, Units, Ref Interval from floating textboxes ───
  // These fields are stored in a <w:drawing> group in this template — they are
  // invisible to the normal paragraph iterator below.
  fillFromDrawings(data);
  // If we got anything from the drawings, mark the result row as captured so
  // the paragraph loop doesn't overwrite with empty tab-split rows.
  let capturedResultRowFromDrawing = !!(data.units || data.refInterval || data.testName);

  // ── helpers ──────────────────────────────────────────────────────────────────
  const readValAttr = (node) => {
    if (!node || !node.attributes) return "";
    for (const attr of Array.from(node.attributes)) {
      if (attr.localName === "val") return attr.value || "";
    }
    return "";
  };

  const escapeHtml = (v = "") =>
    v.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
     .replace(/"/g, "&quot;").replace(/'/g, "&#39;");

  const normalizeHtmlBreaks = (html = "") =>
    html
      .replace(/(?:<br>\s*){3,}/g, "<br><br>")
      .replace(/^(?:\s*<br>\s*)+/i, "")
      .replace(/(?:\s*<br>\s*)+$/i, "")
      .trim();

  // ── numbering map (for list type detection) ──────────────────────────────────
  const numberingMap = (() => {
    if (!numberingXmlString) return { getNumFmt: () => "" };
    const nd = parser.parseFromString(numberingXmlString, "application/xml");
    const abstractMap = new Map();
    Array.from(nd.getElementsByTagNameNS(WORD_NS, "abstractNum")).forEach((abs) => {
      const absId = abs.getAttribute("w:abstractNumId") || abs.getAttribute("abstractNumId") || "";
      const levelFmt = new Map();
      Array.from(abs.getElementsByTagNameNS(WORD_NS, "lvl")).forEach((lvl) => {
        const ilvl = lvl.getAttribute("w:ilvl") || lvl.getAttribute("ilvl") || "0";
        const nf = lvl.getElementsByTagNameNS(WORD_NS, "numFmt")[0];
        levelFmt.set(ilvl, readValAttr(nf).toLowerCase());
      });
      if (absId) abstractMap.set(absId, levelFmt);
    });
    const numToAbstract = new Map();
    Array.from(nd.getElementsByTagNameNS(WORD_NS, "num")).forEach((numNode) => {
      const numId = numNode.getAttribute("w:numId") || numNode.getAttribute("numId") || "";
      const absNode = numNode.getElementsByTagNameNS(WORD_NS, "abstractNumId")[0];
      const absId = readValAttr(absNode);
      if (numId && absId) numToAbstract.set(numId, absId);
    });
    return {
      getNumFmt: (numId, ilvl = "0") => {
        const absId = numToAbstract.get(numId);
        if (!absId) return "";
        const lm = abstractMap.get(absId);
        if (!lm) return "";
        return lm.get(ilvl) || lm.get("0") || "";
      },
    };
  })();

  const getListType = (p, pPr) => {
    if (!pPr) return "";
    const numPr = pPr.getElementsByTagNameNS(WORD_NS, "numPr")[0];
    if (!numPr) return "";
    const numIdNode = numPr.getElementsByTagNameNS(WORD_NS, "numId")[0];
    const ilvlNode  = numPr.getElementsByTagNameNS(WORD_NS, "ilvl")[0];
    const numId = readValAttr(numIdNode);
    const ilvl  = readValAttr(ilvlNode) || "0";
    if (!numId) return "";
    const fmt = numberingMap.getNumFmt(numId, ilvl);
    if (fmt === "bullet") return "bullet";
    if (fmt) return "number";
    return p.getElementsByTagNameNS(WORD_NS, "numPr").length ? "number" : "";
  };

  // Extract plain text from a paragraph
  const getParagraphText = (p) => {
    const runs = p.getElementsByTagNameNS(WORD_NS, "r");
    let text = "";
    Array.from(runs).forEach((run) => {
      Array.from(run.childNodes).forEach((child) => {
        if (child.nodeType !== 1) return;
        if (child.localName === "t") text += child.textContent || "";
        if (child.localName === "tab") text += "\t";
        if (child.localName === "br" || child.localName === "cr") text += "\n";
      });
    });
    return text.trim();
  };

  // Extract rich HTML (preserving bold) from a paragraph
  const getParagraphHtml = (p) => {
    const runs = p.getElementsByTagNameNS(WORD_NS, "r");
    let html = "";
    Array.from(runs).forEach((run) => {
      const rPr  = run.getElementsByTagNameNS(WORD_NS, "rPr")[0];
      const bNode = rPr ? rPr.getElementsByTagNameNS(WORD_NS, "b")[0] : null;
      const bVal  = readValAttr(bNode).toLowerCase();
      const isBold = !!bNode && bVal !== "0" && bVal !== "false";

      Array.from(run.childNodes).forEach((child) => {
        if (child.nodeType !== 1) return;
        if (child.localName === "t") {
          const val = escapeHtml(child.textContent || "");
          html += isBold ? `<strong>${val}</strong>` : val;
        }
        if (child.localName === "tab") html += " ";
        if (child.localName === "br" || child.localName === "cr") html += "<br>";
      });
    });
    return normalizeHtmlBreaks(html);
  };

  // ── section helpers ──────────────────────────────────────────────────────────
  let activeSectionIdx = -1;

  const startSection = (title) => {
    data.sections.push({ title: title.trim(), items: [] });
    activeSectionIdx = data.sections.length - 1;
  };

  const addItem = (text, listType, html) => {
    if (activeSectionIdx < 0) return;
    data.sections[activeSectionIdx].items.push({ text, listType, html: html || escapeHtml(text) });
  };

  // ── state machine ─────────────────────────────────────────────────────────────
  let capturedResultRow = capturedResultRowFromDrawing;
  let inSection = false; // are we currently collecting items for a section?

  // Headings that start new bold sections within the body
  const SECTION_HEADING_RE = /^(note|intended\s*use|clinical\s*use|comments|decreased\s*levels?|increased\s*levels?|reference\s*range|interpretation|methodology|principle|specimen|limitations?|interference|background|clinical\s*significance)s*:?\s*$/i;

  paragraphs.forEach((p) => {
    // Skip paragraphs that live inside textboxes; these are already handled by
    // fillFromDrawings() and otherwise appear as duplicates.
    const isInsideTextbox = (() => {
      let node = p.parentNode;
      while (node && node.nodeType === 1) {
        const local = (node.localName || "").toLowerCase();
        if (local === "txbxcontent" || local === "textbox") return true;
        node = node.parentNode;
      }
      return false;
    })();
    if (isInsideTextbox) return;

    // Skip paragraphs that contain a drawing/inline — their text comes from
    // getParagraphText() crawling INTO the nested textboxes, which duplicates
    // whatever fillFromDrawings() already extracted above.
    const WP_DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    if (
      p.getElementsByTagNameNS(WP_DRAW_NS, "inline").length > 0 ||
      p.getElementsByTagNameNS(WP_DRAW_NS, "anchor").length > 0
    ) return;

    let text = getParagraphText(p);
    let richHtml = getParagraphHtml(p);
    if (!text) return;

    const pPr    = p.getElementsByTagNameNS(WORD_NS, "pPr")[0];
    const pStyle = pPr ? pPr.getElementsByTagNameNS(WORD_NS, "pStyle")[0] : null;
    const styleName = readValAttr(pStyle);
    const listType  = getListType(p, pPr);

    // ── Tab-separated result row ─────────────────────────────────────────────
    if (text.includes("\t")) {
      const parts = text.split("\t").map((s) => s.trim());

      // Skip pure header rows
      const isHeaderRow = parts.some((x) =>
        /^(results?|test\s*name|bio\.?\s*ref\.?\s*interval|units?)$/i.test(x)
      );
      if (isHeaderRow) return;

      if (!capturedResultRow) {
        // Try to find columns by position
        // Common layouts:
        //   4-col: [TestName, Result, Units, Ref]
        //   3-col: [Result, Units, Ref]   (test name already in first paragraph)
        //   2-col: [Units, Ref]
        const nonEmpty = parts.filter((v) => v !== "");

        if (parts.length >= 4) {
          if (!data.testName && parts[0]) data.testName = parts[0];
          if (!data.result)      data.result      = parts[1] || data.result;
          if (!data.units)       data.units       = parts[2] || data.units;
          if (!data.refInterval) data.refInterval = parts[3] || data.refInterval;
        } else if (parts.length === 3) {
          if (!data.result)      data.result      = parts[0] || data.result;
          if (!data.units)       data.units       = parts[1] || data.units;
          if (!data.refInterval) data.refInterval = parts[2] || data.refInterval;
        } else {
          // Fallback: fill from right — only if not already set from drawings
          if (!data.refInterval && nonEmpty.length >= 1) data.refInterval = nonEmpty[nonEmpty.length - 1];
          if (!data.units       && nonEmpty.length >= 2) data.units       = nonEmpty[nonEmpty.length - 2];
          if (!data.result      && nonEmpty.length >= 3) data.result      = nonEmpty[nonEmpty.length - 3];
        }
        capturedResultRow = true;
        return;
      }

      // After table capture, ignore tabbed rows unless they belong to a section.
      // This prevents duplicated "simple text" output under the table.
      if (!inSection) return;

      // Tab text inside a section → treat as plain text
      text = parts.join(" ").replace(/\s+/g, " ").trim();
      richHtml = escapeHtml(text);
      if (!text) return;
    }

    // ── Test name (ALL-CAPS block before the result row) ────────────────────
    if (!data.testName && /^[A-Z0-9][A-Z0-9,\s\-/()]+$/.test(text) && text.length > 3) {
      data.testName = text;
      return;
    }

    // ── Method in parentheses ────────────────────────────────────────────────
    if (!data.method && /^\([^)]+\)$/.test(text)) {
      data.method = text;
      return;
    }

    // ── End of report sentinel ────────────────────────────────────────────────
    if (/end\s*of\s*report/i.test(text)) {
      inSection = false;
      activeSectionIdx = -1;
      return;
    }

    // ── Ignore the IMPORTANT INSTRUCTIONS heading ─────────────────────────────
    if (/^important\s*instructions?$/i.test(text)) {
      inSection = false;
      activeSectionIdx = -1;
      return;
    }

    // ── Detect a paragraph whose ENTIRE text is bold → treat as section heading
    //    (This catches "Decreased Levels", "Increased Levels", inline headings, etc.)
    const isEntirelyBold = (() => {
      const runs = Array.from(p.getElementsByTagNameNS(WORD_NS, "r"));
      if (!runs.length) return false;
      return runs.every((run) => {
        const rPr   = run.getElementsByTagNameNS(WORD_NS, "rPr")[0];
        const bNode = rPr ? rPr.getElementsByTagNameNS(WORD_NS, "b")[0] : null;
        const bVal  = readValAttr(bNode).toLowerCase();
        // A run with no <b> tag is NOT bold, unless it's whitespace/empty
        const runText = getParagraphText(run.parentNode !== p
          ? p // safety
          : (() => { const tmp = doc.createElementNS(WORD_NS, "p"); tmp.appendChild(run.cloneNode(true)); return tmp; })()
        );
        if (!runText.trim()) return true; // blank runs don't count
        return !!bNode && bVal !== "0" && bVal !== "false";
      });
    })();

    // Word-style heading styles
    const isHeadingStyle = /^(heading|title|subtitle)/i.test(styleName);

    // Named section keyword
    const isSectionKeyword = SECTION_HEADING_RE.test(text);

    if (capturedResultRow && (isSectionKeyword || isHeadingStyle || isEntirelyBold)) {
      // Start a new section
      const sectionTitle = text.replace(/:?\s*$/, "").trim();
      startSection(sectionTitle);
      inSection = true;
      return;
    }

    // ── Collect items into active section ─────────────────────────────────────
    if (inSection) {
      addItem(text, listType, richHtml);
      return;
    }

    // ── Fallback for ref interval expressed as plain text before result row ───
    if (!capturedResultRow && !data.refInterval &&
        /^(negative|positive|reactive|non[-\s]?reactive|detected|not detected|present|absent)$/i.test(text)) {
      data.refInterval = text;
      capturedResultRow = true;
    }
  });

  return data;
}

// ─── Generate HTML ────────────────────────────────────────────────────────────
function generateHTML(data) {
  const esc = (v = "") =>
    v.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
     .replace(/"/g, "&quot;").replace(/'/g, "&#39;");

  // Render a list of items: plain paragraphs separated by <br><br>,
  // bullet/number items as <ul>/<ol>
  const renderItems = (items) => {
    if (!items?.length) return "";
    let html = "";
    let listType = "";
    let listItems = [];

    const flushList = () => {
      if (!listType || !listItems.length) return;
      const tag = listType === "bullet" ? "ul" : "ol";
      html += `<${tag} style="margin:0 0 0 22px;padding:0;line-height:1.8;">${
        listItems.map((li) => `<li>${li}</li>`).join("")
      }</${tag}>`;
      listType = "";
      listItems = [];
    };

    items.forEach((item, idx) => {
      const itemHtml = item.html || esc(item.text || "");
      if (!itemHtml) return;
      if (item.listType === "bullet" || item.listType === "number") {
        if (!listType) listType = item.listType;
        if (listType !== item.listType) flushList();
        if (!listType) listType = item.listType;
        listItems.push(itemHtml);
        return;
      }
      flushList();
      // Plain paragraphs: keep spacing tight for downstream paste targets.
      html += (html ? "<br>" : "") + itemHtml;
    });

    flushList();
    return html;
  };

  // Build section blocks (skip "important instructions")
  const sectionHTML = (data.sections || [])
    .filter((s) => s?.title && !/^important\s*instructions?$/i.test(s.title))
    .map((s) => {
      const itemsHtml = renderItems(s.items || []);
      return `<!-- ================= ${s.title.toUpperCase()} ================= -->
<div style="font-weight:700;margin:12px 0 6px 0;">
${esc(s.title)}
</div>
<div style="font-size:13px;line-height:1.7;margin:0 0 10px 0;">
${itemsHtml}
</div>`;
    })
    .join("");

  return `<!-- ================= RESULT TABLE ================= -->
<table style="width:100%;border-collapse:collapse;table-layout:fixed;margin-bottom:10px;">
<colgroup>
<col style="width:40%;">
<col style="width:20%;">
<col style="width:20%;">
<col style="width:20%;">
</colgroup>
<thead>
<tr>
<th style="border:1px solid #000;padding:6px;text-align:left;">Test Name</th>
<th style="border:1px solid #000;padding:6px;text-align:left;">Results</th>
<th style="border:1px solid #000;padding:6px;text-align:left;">Units</th>
<th style="border:1px solid #000;padding:6px;text-align:left;">Bio. Ref. Interval</th>
</tr>
</thead>
<tbody>
<tr>
<td style="border:1px solid #000;padding:6px;vertical-align:top;">
<div style="font-weight:700;text-decoration:underline;">${esc(data.testName)}</div>
${data.method ? `<div style="margin-top:6px;font-weight:600;">${esc(data.method)}</div>` : ""}
</td>
<td style="border:1px solid #000;padding:6px;color:#000;vertical-align:top;">${esc(data.result)}</td>
<td style="border:1px solid #000;padding:6px;vertical-align:top;">${esc(data.units)}</td>
<td style="border:1px solid #000;padding:6px;color:#000;vertical-align:top;">${esc(data.refInterval)}</td>
</tr>
</tbody>
</table>

${sectionHTML}

<!-- ================= END OF REPORT ================= -->
<div style="text-align:center;font-weight:700;margin-top:20px;">
-------------------------------End of report --------------------------------
</div>
<!-- ================= IMPORTANT INSTRUCTIONS ================= -->
<div style="margin-top:30px;font-size:13px;line-height:1.7;">
<p style="text-align:center;font-weight:700;text-decoration:underline;margin:0 0 10px 0;">
IMPORTANT INSTRUCTIONS
</p>
Test results released pertain to the specimen submitted.<br>
All test results are dependent on the quality of the sample received by the Laboratory.<br>
Laboratory investigations are only a tool to facilitate in arriving at a diagnosis and should be clinically correlated by the Referring Physician.<br>
Test results are not valid for medico-legal purposes.
</div>`;
}

// ─── Main Component ───────────────────────────────────────────────────────────
export default function LabHTMLGenerator() {
  const [results, setResults] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [activeIdx, setActiveIdx] = useState(null);
  const [tab, setTab] = useState("html");
  const [copiedIdx, setCopiedIdx] = useState(null);
  const [copiedFileNameIdx, setCopiedFileNameIdx] = useState(null);
  const [editedHTML, setEditedHTML] = useState("");
  const [toastMessage, setToastMessage] = useState("");
  const fileInputRef = useRef();
  const addMoreFileInputRef = useRef();
  const toastTimerRef = useRef(null);

  const showToast = useCallback((message) => {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    setToastMessage(message);
    toastTimerRef.current = setTimeout(() => setToastMessage(""), 1800);
  }, []);

  useEffect(() => {
    return () => { if (toastTimerRef.current) clearTimeout(toastTimerRef.current); };
  }, []);

  const getBaseFileName = useCallback((fileName = "") => fileName.replace(/\.docx$/i, ""), []);

  const processFiles = useCallback(async (fileList, append = false) => {
    setProcessing(true);
    await loadScript("https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js");
    const JSZip = window.JSZip;

    const processed = [];
    for (const file of fileList) {
      try {
        const arrayBuffer = await file.arrayBuffer();
        const zip = await JSZip.loadAsync(arrayBuffer);
        const docXml = await zip.file("word/document.xml").async("string");
        const numberingFile = zip.file("word/numbering.xml");
        const numberingXml = numberingFile ? await numberingFile.async("string") : "";
        const parsed = parseDocxXml(docXml, numberingXml);
        const html = generateHTML(parsed);
        processed.push({ fileName: file.name, parsed, html, status: "ok", editedHTML: html });
      } catch (err) {
        processed.push({ fileName: file.name, parsed: {}, html: "", status: "error", error: err.message, editedHTML: "" });
      }
    }

    if (append) {
      setResults((prev) => {
        const next = [...prev, ...processed];
        const newActiveIdx = prev.length;
        setActiveIdx(newActiveIdx);
        setTab("html");
        if (processed[0]) setEditedHTML(processed[0].html);
        return next;
      });
    } else {
      setResults(processed);
      setActiveIdx(0);
      setTab("html");
      if (processed[0]) setEditedHTML(processed[0].html);
    }
    setProcessing(false);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    const dropped = Array.from(e.dataTransfer.files).filter((f) => f.name.endsWith(".docx"));
    if (dropped.length) { processFiles(dropped); }
  }, [processFiles]);

  const handleFileInput = (e) => {
    const chosen = Array.from(e.target.files).filter((f) => f.name.endsWith(".docx"));
    if (chosen.length) { processFiles(chosen); }
  };

  const handleAddMoreFiles = (e) => {
    const chosen = Array.from(e.target.files).filter((f) => f.name.endsWith(".docx"));
    if (chosen.length) {
      processFiles(chosen, true);
    }
    e.target.value = "";
  };

  const selectResult = (idx) => {
    setActiveIdx(idx);
    setTab("html");
    setEditedHTML(results[idx].editedHTML);
  };

  const copyToClipboard = async (idx) => {
    await navigator.clipboard.writeText(results[idx].editedHTML || results[idx].html);
    setCopiedIdx(idx);
    setTimeout(() => setCopiedIdx(null), 2000);
  };

  const copyAll = async () => {
    const allHtml = results.filter((r) => r.status === "ok")
      .map((r) => `<!-- FILE: ${r.fileName} -->\n${r.editedHTML || r.html}`).join("\n\n");
    await navigator.clipboard.writeText(allHtml);
  };

  const copyFileName = async (idx) => {
    await navigator.clipboard.writeText(getBaseFileName(results[idx].fileName));
    setCopiedFileNameIdx(idx);
    setTimeout(() => setCopiedFileNameIdx(null), 1500);
    showToast("Filename copied");
  };

  const downloadAll = () => {
    const allHtml = results.filter((r) => r.status === "ok")
      .map((r) => `<!-- FILE: ${r.fileName} -->\n${r.editedHTML || r.html}`).join("\n\n");
    const blob = new Blob([allHtml], { type: "text/html" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "all_lab_templates.html"; a.click();
    URL.revokeObjectURL(url);
  };

  const saveEdit = () => {
    setResults((prev) => prev.map((r, i) => i === activeIdx ? { ...r, editedHTML } : r));
    setTab("html");
  };

  const removeResult = (idxToRemove) => {
    setResults((prev) => {
      const next = prev.filter((_, idx) => idx !== idxToRemove);
      setCopiedFileNameIdx((pi) => {
        if (pi === null) return pi;
        if (pi === idxToRemove) return null;
        return pi > idxToRemove ? pi - 1 : pi;
      });
      if (!next.length) { setActiveIdx(null); setEditedHTML(""); return next; }
      if (activeIdx === idxToRemove) {
        const ni = Math.min(idxToRemove, next.length - 1);
        setActiveIdx(ni);
        setEditedHTML(next[ni].editedHTML || next[ni].html || "");
      } else if (activeIdx > idxToRemove) {
        setActiveIdx(activeIdx - 1);
      }
      return next;
    });
  };

  const active = activeIdx !== null ? results[activeIdx] : null;
  const displayHTML = active
    ? tab === "edit" ? editedHTML : (active.editedHTML || active.html)
    : "";

  return (
    <div style={{ fontFamily: "'IBM Plex Mono','Courier New',monospace", background: "#0a0a0f", minHeight: "100vh", color: "#e8e8f0", display: "flex", flexDirection: "column" }}>
      {toastMessage && (
        <div style={{ position: "fixed", top: 12, left: "50%", transform: "translateX(-50%)", background: "#0f2436", border: "1px solid #00d2ff", color: "#b8e8ff", padding: "8px 14px", borderRadius: 8, fontSize: 11, letterSpacing: "0.05em", zIndex: 2000, animation: "toastSlide 0.2s ease-out" }}>
          {toastMessage}
        </div>
      )}

      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#0d1117 0%,#161b27 100%)", borderBottom: "1px solid #1e3a5f", padding: "18px 28px", display: "flex", alignItems: "center", gap: 16 }}>
        <div style={{ width: 36, height: 36, borderRadius: 8, background: "linear-gradient(135deg,#00d2ff,#0066cc)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 900, color: "#fff", boxShadow: "0 0 20px rgba(0,210,255,0.3)" }}>⚗</div>
        <div>
          <div style={{ fontWeight: 700, fontSize: 16, letterSpacing: "0.05em", color: "#00d2ff" }}>LAB HTML GENERATOR</div>
          <div style={{ fontSize: 11, color: "#5a7a9a", letterSpacing: "0.08em" }}>FIRSTMEDIX · BATCH DOCX → HTML CONVERTER</div>
        </div>
        {results.length > 0 && (
          <div style={{ marginLeft: "auto", display: "flex", gap: 10 }}>
            <button onClick={copyAll} style={btnStyle("#1a2a3a", "#00d2ff")}>📋 Copy All HTML</button>
            <button onClick={downloadAll} style={btnStyle("#162a1a", "#00ff88")}>⬇ Download All</button>
          </div>
        )}
      </div>

      <div style={{ display: "flex", flex: 1, overflow: "hidden", height: "calc(100vh - 73px)" }}>
        {/* Sidebar */}
        <div style={{ width: 240, background: "#0d1117", borderRight: "1px solid #1e2d3d", overflowY: "auto", flexShrink: 0 }}>
          {results.length === 0 ? (
            <div
              onDrop={handleDrop} onDragOver={(e) => e.preventDefault()}
              onClick={() => fileInputRef.current.click()}
              style={{ margin: 16, border: "2px dashed #1e3a5f", borderRadius: 10, padding: "32px 16px", textAlign: "center", cursor: "pointer", background: "rgba(0,210,255,0.03)" }}
              onMouseEnter={(e) => e.currentTarget.style.borderColor = "#00d2ff"}
              onMouseLeave={(e) => e.currentTarget.style.borderColor = "#1e3a5f"}
            >
              <div style={{ fontSize: 32, marginBottom: 12 }}>📂</div>
              <div style={{ fontSize: 12, color: "#4a6a8a", lineHeight: 1.6 }}>Drop .docx files here<br />or click to browse</div>
              <input ref={fileInputRef} type="file" accept=".docx" multiple style={{ display: "none" }} onChange={handleFileInput} />
            </div>
          ) : (
            <>
              <div style={{ padding: "10px 16px", fontSize: 11, color: "#4a6a8a", letterSpacing: "0.08em", borderBottom: "1px solid #1e2d3d", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span>FILES ({results.length})</span>
                <button onClick={() => { setResults([]); setActiveIdx(null); }} style={{ background: "none", border: "none", color: "#4a6a8a", cursor: "pointer", fontSize: 16 }}>×</button>
              </div>
              {results.map((r, i) => (
                <div key={i} onClick={() => selectResult(i)}
                  style={{ padding: "10px 14px", borderBottom: "1px solid #111827", cursor: "pointer", background: activeIdx === i ? "rgba(0,210,255,0.08)" : "transparent", borderLeft: activeIdx === i ? "3px solid #00d2ff" : "3px solid transparent", transition: "all 0.15s", display: "flex", alignItems: "flex-start", gap: 8 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 11, color: r.status === "error" ? "#ff4444" : "#c0d8f0", wordBreak: "break-word", lineHeight: 1.4 }}>
                      {r.status === "error" ? "⚠ " : "✓ "}{getBaseFileName(r.fileName)}
                    </div>
                    {r.parsed?.testName && (
                      <div style={{ fontSize: 10, color: "#3a6a8a", marginTop: 3 }}>{r.parsed.testName.substring(0, 30)}</div>
                    )}
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <button onClick={(e) => { e.stopPropagation(); copyFileName(i); }} title={copiedFileNameIdx === i ? "Copied" : "Copy file name"}
                      style={{ background: "none", border: "none", color: copiedFileNameIdx === i ? "#00d2ff" : "#4a6a8a", cursor: "pointer", fontSize: 12, lineHeight: 1, padding: "0 2px" }}>
                      {copiedFileNameIdx === i ? "✓" : "⧉"}
                    </button>
                    <button onClick={(e) => { e.stopPropagation(); removeResult(i); }} title="Remove file"
                      style={{ background: "none", border: "none", color: "#4a6a8a", cursor: "pointer", fontSize: 14, lineHeight: 1, padding: "0 2px" }}>×</button>
                  </div>
                </div>
              ))}
              <div onClick={() => addMoreFileInputRef.current.click()}
                style={{ padding: "12px 14px", cursor: "pointer", color: "#3a6a8a", fontSize: 11, textAlign: "center", borderTop: "1px solid #1e2d3d" }}>
                + Add more files
              </div>
              <input ref={addMoreFileInputRef} type="file" accept=".docx" multiple style={{ display: "none" }} onChange={handleAddMoreFiles} />
            </>
          )}
          {processing && (
            <div style={{ padding: 20, textAlign: "center" }}>
              <div style={{ fontSize: 12, color: "#00d2ff", animation: "pulse 1s infinite" }}>⟳ Processing...</div>
            </div>
          )}
        </div>

        {/* Main panel */}
        <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden" }}>
          {active ? (
            <>
              <div style={{ display: "flex", alignItems: "center", gap: 2, padding: "8px 16px", background: "#0d1117", borderBottom: "1px solid #1e2d3d" }}>
                {["html", "preview", "edit"].map((t) => (
                  <button key={t} onClick={() => { if (t === "edit") setEditedHTML(active.editedHTML || active.html); setTab(t); }}
                    style={{ padding: "5px 14px", fontSize: 11, letterSpacing: "0.06em", fontFamily: "inherit", cursor: "pointer", borderRadius: 5, border: tab === t ? "1px solid #00d2ff" : "1px solid #1e3a5f", background: tab === t ? "rgba(0,210,255,0.12)" : "transparent", color: tab === t ? "#00d2ff" : "#4a6a8a", textTransform: "uppercase" }}>
                    {t === "html" ? "⟨/⟩ HTML" : t === "preview" ? "👁 Preview" : "✎ Edit"}
                  </button>
                ))}
                <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
                  <span style={{ fontSize: 11, color: "#3a5a7a" }}>{active.fileName}</span>
                  {tab === "edit" ? (
                    <button onClick={saveEdit} style={btnStyle("#162a1a", "#00ff88", 11)}>💾 Save</button>
                  ) : (
                    <button onClick={() => copyToClipboard(activeIdx)} style={btnStyle("#1a2a3a", "#00d2ff", 11)}>
                      {copiedIdx === activeIdx ? "✓ Copied!" : "📋 Copy HTML"}
                    </button>
                  )}
                </div>
              </div>

              <div style={{ flex: 1, overflow: "auto" }}>
                {tab === "html" && (
                  <pre style={{ margin: 0, padding: "20px 24px", fontSize: 12, lineHeight: 1.7, color: "#a8c8e8", whiteSpace: "pre-wrap", wordBreak: "break-word", background: "#080c10" }}>
                    <code>{displayHTML}</code>
                  </pre>
                )}
                {tab === "preview" && (
                  <div style={{ padding: 24, background: "#fff" }}>
                    <iframe srcDoc={`<html><body style="font-family:Arial,sans-serif;font-size:13px;padding:20px;">${displayHTML}</body></html>`}
                      style={{ width: "100%", height: "calc(100vh - 150px)", border: "none" }} title="Preview" />
                  </div>
                )}
                {tab === "edit" && (
                  <textarea value={editedHTML} onChange={(e) => setEditedHTML(e.target.value)}
                    style={{ width: "100%", height: "100%", background: "#080c10", color: "#a8c8e8", border: "none", padding: "20px 24px", fontSize: 12, lineHeight: 1.7, fontFamily: "'IBM Plex Mono',monospace", resize: "none", outline: "none", boxSizing: "border-box" }} />
                )}
              </div>

              {active.parsed?.testName && (
                <div style={{ padding: "10px 20px", background: "#0a1520", borderTop: "1px solid #1e2d3d", display: "flex", gap: 24, flexWrap: "wrap" }}>
                  {[["Test", active.parsed.testName], ["Method", active.parsed.method], ["Result", active.parsed.result], ["Units", active.parsed.units], ["Ref. Interval", active.parsed.refInterval]]
                    .map(([label, val]) => val ? (
                      <div key={label}>
                        <span style={{ fontSize: 10, color: "#3a6a8a", letterSpacing: "0.06em" }}>{label} </span>
                        <span style={{ fontSize: 11, color: "#7ab8e8" }}>{val}</span>
                      </div>
                    ) : null)}
                </div>
              )}
            </>
          ) : (
            <div style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", color: "#2a4a6a", gap: 16 }}>
              <div onDrop={handleDrop} onDragOver={(e) => e.preventDefault()} onClick={() => fileInputRef.current?.click()}
                style={{ border: "2px dashed #1e3a5f", borderRadius: 16, padding: "60px 80px", textAlign: "center", cursor: "pointer", transition: "all 0.2s", background: "rgba(0,210,255,0.02)" }}
                onMouseEnter={(e) => { e.currentTarget.style.borderColor = "#00d2ff"; e.currentTarget.style.background = "rgba(0,210,255,0.05)"; }}
                onMouseLeave={(e) => { e.currentTarget.style.borderColor = "#1e3a5f"; e.currentTarget.style.background = "rgba(0,210,255,0.02)"; }}>
                <div style={{ fontSize: 56, marginBottom: 20 }}>📄</div>
                <div style={{ fontSize: 16, color: "#3a6a9a", marginBottom: 8, fontWeight: 700 }}>Drop your .docx files here</div>
                <div style={{ fontSize: 13, color: "#2a4a6a", lineHeight: 1.6 }}>Upload all 700 files at once<br />HTML will be generated instantly for each one</div>
              </div>
              <input ref={fileInputRef} type="file" accept=".docx" multiple style={{ display: "none" }} onChange={handleFileInput} />
              <div style={{ fontSize: 11, color: "#1e3a5f", letterSpacing: "0.06em" }}>SUPPORTS BATCH UPLOAD · NO SERVER NEEDED · 100% LOCAL</div>
            </div>
          )}
        </div>
      </div>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&display=swap');
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.4} }
        @keyframes toastSlide { from{opacity:0;transform:translateX(-50%) translateY(-8px)} to{opacity:1;transform:translateX(-50%) translateY(0)} }
        ::-webkit-scrollbar{width:6px} ::-webkit-scrollbar-track{background:#080c10} ::-webkit-scrollbar-thumb{background:#1e3a5f;border-radius:3px}
      `}</style>
    </div>
  );
}

function btnStyle(bg, border, fontSize = 12) {
  return { padding: "6px 14px", fontSize, fontFamily: "'IBM Plex Mono',monospace", cursor: "pointer", borderRadius: 6, border: `1px solid ${border}`, background: bg, color: border, letterSpacing: "0.04em", transition: "all 0.15s" };
}
