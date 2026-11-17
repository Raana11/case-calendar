/* Case Calendar – update engine.
   Reads the AP High Court cause list (PDF, via pdf.js) and the endowmentsinfo
   Court_Cases_Report (Excel, via SheetJS), matches them and builds:
     - segregation lists for verification
     - the data.json the team app reads (same format the app already uses)
   Works in the browser (window.CCEngine) and in Node (module.exports) for tests. */
(function (root) {
  "use strict";

  const TTD_RE = /\bT\.?\s?T\.?\s?D\b|TIRUMALA\s+TIRUPATI/i;
  const ENDOW_RE = /ENDOWMENT/i;
  const TEMPLE_RE = /DEVASTHANAM|DEVASTANAM|\bTEMPLE\b|\bMUTT\b|\bMATH(AM)?\b|\bMATAM\b|CHOULTRY|SATRAM|DHARMIKA|ARCHAKA|PEETHAM/i;
  const CASE_RE = /^([A-Z][A-Z().]*)\/(\d+)\/(\d{4})$/;
  const MONTHS = ["january","february","march","april","may","june","july","august","september","october","november","december"];
  const TYPE_MAP = { "WRIT (PETITIONS)": "WP", "WRIT": "WP", "WRIT (APPEAL)": "WA", "WRIT (PIL)": "WP(PIL)", "CC": "CC", "CRP": "CRP", "CMA": "CMA" };

  const pad = n => String(n).padStart(2, "0");
  const keyOf = n => String(n).toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  const normNo = (type, no, year) => `${type.toUpperCase()} ${+no}/${year}`;
  const clean = s => String(s == null ? "" : s).replace(/\s+/g, " ").trim();

  /* ---------- cause list PDF ---------- */

  // pdf.js document -> array of pages, each an array of lines [{y, items:[{x, s}]}]
  async function pdfToLines(pdf, onProgress) {
    const pages = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const tc = await page.getTextContent();
      const rows = new Map();
      for (const it of tc.items) {
        const s = (it.str || "").trim();
        if (!s) continue;
        if (Math.abs(it.transform[1]) > 0.01 || Math.abs(it.transform[2]) > 0.01) continue; // rotated watermark text
        const x = it.transform[4], y = Math.round(it.transform[5]);
        let key = null;
        for (const k of rows.keys()) if (Math.abs(k - y) <= 2) { key = k; break; }
        if (key === null) { key = y; rows.set(key, []); }
        rows.get(key).push({ x, s });
      }
      const lines = [...rows.entries()].sort((a, b) => b[0] - a[0])
        .map(([y, items]) => ({ y, items: items.sort((a, b) => a.x - b.x) }));
      pages.push(lines);
      if (onProgress && p % 10 === 0) onProgress(p, pdf.numPages);
      page.cleanup && page.cleanup();
    }
    return pages;
  }

  // lines -> {date, items:[{court, sno, caseNo, lc[], district, party, stage, text, endow, ttd}]}
  function parseCauseList(pages) {
    let date = "", court = "", stage = "", caseX = null, distX = null, cur = null;
    const out = [];
    const finish = () => { if (cur) { cur.text = clean(cur.text); out.push(cur); cur = null; } };
    for (const lines of pages) {
      for (const ln of lines) {
        ln.items = ln.items.filter(i => !/^Page \d+ of \d+$/i.test(i.s));
        if (!ln.items.length) continue;
        const text = clean(ln.items.map(i => i.s).join(" "));
        let m;
        if ((m = text.match(/COURT NO\.?\s*(\d+)/i)) && ln.items.length <= 3) { finish(); court = m[1]; stage = ""; continue; }
        if ((m = text.match(/To be Heard on \w+ The (\d+)\w* day of (\w+) (\d{4})/i))) {
          const mo = MONTHS.indexOf(m[2].toLowerCase());
          if (mo >= 0 && !date) date = `${m[3]}-${pad(mo + 1)}-${pad(+m[1])}`;
          continue;
        }
        if (/^(SNO|S\.\s?No\.?)\b/i.test(text)) {   // column header (both cause list formats)
          const c = ln.items.find(i => /^CASE( NUMBER)?$/i.test(i.s)); if (c) caseX = c.x;
          const d = ln.items.find(i => /^DISTRICT/i.test(i.s)); if (d) distX = d.x;
          continue;
        }
        if (/^(District\/Case|Remarks)$/i.test(text)) { if (/^District/i.test(text)) distX = ln.items[0].x; continue; }
        if (/^Page \d+ of \d+$/i.test(text)) continue;
        // new item: "<sno> <CASE/NO/YEAR> ..."
        const a = ln.items[0], b = ln.items[1];
        let first = a && a.s, second = b && b.s;
        // pdf.js sometimes glues "12  WP/1/2026" into one string
        if (first && !second && /^\d+(\.\d+)?\s+\S+/.test(first)) { const p = first.split(/\s+/); first = p[0]; second = p[1]; }
        else if (first && /^\d+(\.\d+)?\s+[A-Z]/.test(first)) { const p = first.split(/\s+/); first = p[0]; second = p[1]; }
        if (first && /^\d+(\.\d+)?$/.test(first) && second && CASE_RE.test(second.split(/\s+/)[0])) {
          finish();
          const cm = second.split(/\s+/)[0].match(CASE_RE);
          const distItem = distX != null ? ln.items.filter(i => i.x >= distX - 45).map(i => i.s).join(" ") : "";
          cur = { court, sno: first, caseNo: normNo(cm[1], cm[2], cm[3]), lc: [], district: clean(distItem), stage, text: text };
          if (cm[1] === "IA") { cur.ia = cur.caseNo; cur.needMain = true; }
          continue;
        }
        if (cur) {
          // connected case in brackets e.g. (WP/9663/2025)
          if (cur.needMain) {   // "IA/4/2026" items: the main case is on a following line, e.g. WP/10547/2006
            const mi = ln.items.find(i => CASE_RE.test(i.s));
            if (mi) { const mm = mi.s.match(CASE_RE); cur.caseNo = normNo(mm[1], mm[2], mm[3]); cur.needMain = false; }
          }
          for (const it of ln.items) {
            const lm = it.s.match(/^\(([A-Z][A-Z().]*)\/(\d+)\/(\d{4})\)$/);
            if (lm) cur.lc.push(normNo(lm[1], lm[2], lm[3]));
          }
          // heading line after the block? (starts left of the case column, all caps, no digits-first)
          if (caseX != null && a.x < caseX - 8 && !/^\d/.test(a.s) && /^[A-Z ()\-&.,/]+$/.test(text) && text.length > 6) {
            finish(); stage = text; continue;
          }
          cur.text += " " + text;
          continue;
        }
        if (caseX != null && a && a.x < caseX - 8 && /^[A-Z][A-Z ()\-&.,/]+$/.test(text) && !/along with/i.test(text)) stage = text;
      }
    }
    finish();
    for (const it of out) { it.endow = ENDOW_RE.test(it.text); it.ttd = TTD_RE.test(it.text); }
    return { date, items: out };
  }

  /* ---------- endowmentsinfo Excel ---------- */

  function dmyToIso(v) {
    if (v == null || v === "") return "";
    if (typeof v === "number" && v > 20000 && v < 80000) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return `${m[1]}-${pad(+m[2])}-${pad(+m[3])}`;
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/); if (m) return `${m[3]}-${pad(+m[2])}-${pad(+m[1])}`;
    return "";
  }

  // rows: array of objects keyed by header (SheetJS sheet_to_json)
  function parseEndowmentsInfo(rows) {
    const map = new Map(); let skipped = 0;
    for (const r of rows) {
      const type = TYPE_MAP[clean(r["Case Type"]).toUpperCase()];
      const no = clean(r["Case Number"]), year = clean(r["Year"]);
      if (!type || !/^\d+$/.test(no) || !/^\d{4}$/.test(year)) { skipped++; continue; }
      const caseNo = normNo(type, no, year);
      const cf = clean(r["Counter Filed or Not"]).toLowerCase();
      const counter = cf === "filed" ? "filed" : cf === "not filed" ? "notfiled" : cf.startsWith("need not") ? "proforma" : "blank";
      const sec = clean(r["Sections"]);
      const rec = {
        caseNo, counter, sec: sec === "-" ? "" : sec, temple: clean(r["Temple"]), district: clean(r["District"]),
        listing: dmyToIso(r["Date of Listing"]), status: clean(r["Court Case Status"]),
        petitioner: clean(r["Petitioner Name"]), whoFiles: clean(r["Counter to be Filed"]),
        subject: clean(r["Subject"]) || clean(r["Category"]), remarks: clean(r["Remarks"])
      };
      rec.ttd = TTD_RE.test([rec.temple, rec.petitioner, rec.remarks].join(" "));
      const prev = map.get(caseNo);
      if (!prev || (prev.status === "Disposed" && rec.status === "Pending")) map.set(caseNo, rec);
    }
    return { map, skipped };
  }

  // cases-info.json for the team app's "Dashboard" tab (all endowmentsinfo cases, TTD left out)
  function buildCasesInfo(rows, nowIso) {
    const out = []; let ttd = 0;
    for (const r of rows) {
      const n = clean(r["Case Number"]).replace(/^0+(?=\d)/, ""), y = clean(r["Year"]);
      if (!/^\d+$/.test(n) || !/^\d{4}$/.test(y)) continue;
      const tem = clean(r["Temple"]);
      if (TTD_RE.test(tem)) { ttd++; continue; }
      out.push([n, y, dmyToIso(r["Date of Listing"]), clean(r["Sections"]), clean(r["District"]), clean(r["Mandal"]), tem,
        clean(r["Petitioner Name"]), clean(r["Counter Filed or Not"]).toLowerCase(), clean(r["Court Case Status"])]);
    }
    return { meta: { at: nowIso, source: "endowmentsinfo Excel (Update Dashboard)", count: out.length, ttdExcluded: ttd },
      cols: ["n", "y", "d", "s", "dist", "man", "tem", "pet", "ctr", "st"], rows: out };
  }


  /* ---------- Legal Cell results sheet (photo → OCR words) ---------- */
  const RES_TYPES = ["WP(PIL)", "CRLRC", "CRLA", "CRP", "CMA", "WP", "WA", "CC", "SA", "AS"];
  const OCR_DIGIT = { O: "0", Q: "0", D: "0", I: "1", L: "1", "|": "1", "!": "1", S: "5", B: "8", Z: "2", G: "6" };
  function fixCaseVariants(raw) {
    let t = String(raw || "").toUpperCase().replace(/\s+/g, "");
    const pil = /\(?PIL\)?/.test(t);
    const lm = t.match(/[A-Z]+/); if (!lm) return [];
    let letters = lm[0].replace(/^V?W+P/, "WP").replace(/^VP/, "WP").replace(/^V?W+A/, "WA").replace(/^VA/, "WA");
    const type = pil ? "WP(PIL)" : (["CRLRC", "CRLA", "CRP", "CMA", "WP", "WA", "CC", "SA", "AS"].find(x => letters.startsWith(x)) || "");
    if (!type) return [];
    let rest = (pil ? t.replace(/.*PIL\)?/, "") : letters.slice(type.replace("(PIL)", "").length) + t.slice(lm.index + lm[0].length));
    const amb = rest.replace(/[A-Z|!]/g, ch => OCR_DIGIT[ch] || "");
    const alts = new Set([amb]); if (/B/.test(rest)) alts.add(rest.replace(/[A-Z|!]/g, ch => ch === "B" ? "6" : (OCR_DIGIT[ch] || "")));
    const out = [];
    for (const r of alts) {
      const digits = r.replace(/[^0-9]/g, ""); if (digits.length < 5) continue;
      let parts = r.split(/[^0-9]+/).filter(Boolean), no, year;
      if (parts.length >= 2 && /^(19|20)\d\d$/.test(parts[parts.length - 1])) { year = parts.pop(); no = parts.join(""); }
      else { year = digits.slice(-4); no = digits.slice(0, -4); }
      if (/^(19|20)\d\d$/.test(year) && no) out.push(`${type} ${+no}/${year}`);
    }
    return out;
  }
  function fixCase(raw) { return fixCaseVariants(raw)[0] || null; }
  // section cell says TTD (read as "TTD", "TID", "TD", "1D"…): no "(SEC)" suffix and only T/I/1 letters before a D
  const TTD_SEC = { test: s => { const c = String(s).toUpperCase().replace(/[^A-Z0-9|]/g, ""); return /T[TI1L|]\W?[D0O]/.test(c) || (!/S[E]?C/.test(c) && /^[TI1L|]{1,3}[D0O]$/.test(c)); } };
  function classifyRemark(r) {
    if (/CLOS|DISPOS|DISMISS|INFRUCT|WITHDRA/i.test(r)) return "C5";
    if (/RE?ACH|REACT|N[O0]T\s*RE|WEEK\s*O/i.test(r)) return "C6";
    if (/INTERIM|EXTEND/i.test(r)) return "C3";
    if (/RESERV|JUDG/i.test(r)) return "C4";
    if (/\bC\.?A\b|COUNTER|F[O0]R\s*CA|[lI1]orca/i.test(r)) return "C2";
    return "C1";
  }
  function remarkDates(r, listIso) {
    const out = []; const re = /(\d{1,2})\s*[.,\-\/:]\s*(\d{1,2})\s*[.,\-\/:]\s*(\d{4})/g; let m;
    const y0 = +listIso.slice(0, 4);
    while ((m = re.exec(r))) {
      const d = +m[1], mo = +m[2], y = +m[3];
      if (mo < 1 || mo > 12 || d < 1 || d > 31) { out.push({ bad: m[0] }); continue; }
      const iso = `${y}-${pad(mo)}-${pad(d)}`;
      if (y >= y0 && y <= y0 + 1 && iso > listIso) { out.push({ iso }); continue; }
      const fix = [y0, y0 + 1].map(yy => `${yy}-${pad(mo)}-${pad(d)}`).find(x => x > listIso);
      out.push(fix ? { iso: fix, fixed: m[0] } : { bad: m[0] });
    }
    return out;
  }


  function editDist(a, b) {
    const m = a.length, n = b.length, d = Array.from({ length: m + 1 }, (_, i) => [i, ...Array(n).fill(0)]);
    for (let j = 1; j <= n; j++) d[0][j] = j;
    for (let i = 1; i <= m; i++) for (let j = 1; j <= n; j++) d[i][j] = Math.min(d[i-1][j] + 1, d[i][j-1] + 1, d[i-1][j-1] + (a[i-1] === b[j-1] ? 0 : 1));
    return d[m][n];
  }
  function nearestKnown(caseNo, known) {
    const m = caseNo.match(/^(\S+) (\d+)\/(\d{4})$/); if (!m) return [];
    let best = [], bd = 3;
    for (const k of known) {
      const q = k.match(/^(\S+) (\d+)\/(\d{4})$/); if (!q) continue;
      const typeOk = q[1] === m[1] || (q[1] === "WP(PIL)" && m[1] === "WP") || (q[1] === "WP" && m[1] === "WP(PIL)");
      if (!typeOk) continue;
      const dist = editDist(m[2], q[2]) + (q[3] === m[3] ? 0 : 1);
      if (dist < bd) { bd = dist; best = [k]; } else if (dist === bd) best.push(k);
    }
    return best;
  }
  // words: [{t, x0, y0, x1, y1, c}] from one photo. opts.known: Set of case numbers we already track.
  function parseResultsWords(words, opts = {}) {
    const known = opts.known || new Set();
    const H = words.map(w => w.y1 - w.y0).sort((a, b) => a - b)[Math.floor(words.length / 2)] || 20;
    const lines = [];
    for (const w of [...words].sort((a, b) => (a.y0 + a.y1) - (b.y0 + b.y1))) {
      const cy = (w.y0 + w.y1) / 2;
      let ln = lines.find(l => Math.abs(l.cy - cy) < H * 0.6);
      if (!ln) { ln = { cy, words: [] }; lines.push(ln); }
      ln.words.push(w);
    }
    lines.forEach(l => l.words.sort((a, b) => a.x0 - b.x0)); lines.sort((a, b) => a.cy - b.cy);
    let listDate = "", cols = null;
    for (const l of lines) {
      const txt = l.words.map(w => w.t).join(" ");
      const dm = txt.match(/CAUSE\s*LIST\s*(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})/i);
      if (dm && !listDate) listDate = `${dm[3]}-${pad(+dm[2])}-${pad(+dm[1])}`;
      if (/REMARK/i.test(txt) && /DISTRICT/i.test(txt) && !cols) {
        const xOf = re => { const w = l.words.find(w => re.test(w.t)); return w ? w.x0 : null; };
        const cw = l.words.filter(w => /^CASE/i.test(w.t));
        const item = xOf(/^ITEM/i), cas = cw[0] && cw[0].x0, stage = cw[1] && cw[1].x0, dist = xOf(/^DISTRICT/i), sec = xOf(/^SECTION/i), rem = xOf(/^REMARK/i);
        if ([item, cas, stage, dist, sec, rem].every(v => v != null))
          cols = { b1: (item + cas) / 2, b2: (cas + stage) / 2, b3: (stage + dist) / 2, b4: (dist + sec) / 2, b5: (sec + rem) / 2 };
      }
    }
    if (!listDate) listDate = opts.listDate || "";
    let colsGuessed = false;
    if (!cols && opts.imgW) {   // header not readable: use the sheet's usual column proportions
      const W = opts.imgW; colsGuessed = true;
      cols = { b1: W * 0.096, b2: W * 0.226, b3: W * 0.394, b4: W * 0.557, b5: W * 0.723 };
    }
    if (!cols) return { listDate, rows: [], error: "Table header (ITEM / CASE NO / … / REMARKS) not found in photo" };
    const colOf = x => x < cols.b1 ? "item" : x < cols.b2 ? "case" : x < cols.b3 ? "stage" : x < cols.b4 ? "dist" : x < cols.b5 ? "sec" : "rem";
    // segment into court blocks and anchor rows
    let court = "", rows = [], pendingRem = [];
    const anchors = [];
    for (const l of lines) {
      const txt = l.words.map(w => w.t).join(" ");
      const cm = txt.match(/C[O0]URT\s*N[O0]\.?\s*[:;.]?\s*(\d{1,2})\b/i);
      if (cm) { court = cm[1]; anchors.push({ sep: true, cy: l.cy }); continue; }
      if (/REMARK/i.test(txt) && /DISTRICT/i.test(txt)) { anchors.push({ sep: true, cy: l.cy }); continue; }
      const cell = { item: [], case: [], stage: [], dist: [], sec: [], rem: [] };
      for (const w of l.words) cell[colOf(w.x0)].push(w);
      // try all words in the case column, then fewer (stray fragments from the next line can get attached)
      let caseTxt = "", vs = [];
      for (let k = cell.case.length; k >= 1 && !vs.length; k--) { caseTxt = cell.case.slice(0, k).map(w => w.t).join(""); vs = caseTxt.length >= 5 ? fixCaseVariants(caseTxt) : []; }
      const cn = vs.find(v => known.has(v)) || vs[0] || null;
      if (cn) anchors.push({ cy: l.cy, court, cell, caseRaw: caseTxt, caseNo: cn, conf: Math.min(...cell.case.map(w => w.c ?? 100)), extra: [] });
      else anchors.push({ cy: l.cy, loose: cell });
    }
    // attach loose lines (remarks above/below the main line) to nearest anchor row within the same court block
    const real = anchors.map((a, i) => ({ a, i })).filter(o => o.a.caseNo);
    for (let i = 0; i < anchors.length; i++) {
      const a = anchors[i]; if (!a.loose) continue;
      let best = null, bd = Infinity;
      for (const o of real) {
        const between = anchors.slice(Math.min(i, o.i), Math.max(i, o.i)).some(x => x.sep);
        if (between) continue;
        const d = Math.abs(o.a.cy - a.cy); if (d < bd) { bd = d; best = o.a; }
      }
      if (best && bd < H * 4) best.extra.push({ cy: a.cy, cell: a.loose });
    }
    // second reading of the remarks column (grayscale, sharper digits) – opts.remWords already in this photo's coordinates
    const remLines = [];
    for (const w of (opts.remWords || []).slice().sort((p, q) => (p.y0 + p.y1) - (q.y0 + q.y1))) {
      const cy = (w.y0 + w.y1) / 2; let ln = remLines.find(l => Math.abs(l.cy - cy) < H * 0.6);
      if (!ln) { ln = { cy, words: [] }; remLines.push(ln); } ln.words.push(w);
    }
    remLines.forEach(l => l.words.sort((p, q) => p.x0 - q.x0));
    const ys = anchors.map(x => x.cy);
    for (const { a, i } of real) {
      const all = [{ cy: a.cy, cell: a.cell }, ...a.extra].sort((p, q) => p.cy - q.cy);
      const join = k => all.map(p => p.cell[k].map(w => w.t).join(" ")).filter(Boolean).join(" ").trim();
      // row band: halfway to the previous / next real row or separator
      let lo = -Infinity, hi = Infinity;
      for (let j = i - 1; j >= 0; j--) if (anchors[j].sep || anchors[j].caseNo) { lo = (ys[j] + a.cy) / 2; break; }
      for (let j = i + 1; j < anchors.length; j++) if (anchors[j].sep || anchors[j].caseNo) { hi = (ys[j] + a.cy) / 2; break; }
      const remark2 = remLines.filter(l => l.cy > lo && l.cy <= hi).map(l => l.words.map(w => w.t).join(" ")).join(" ").trim();
      const remark1 = join("rem"), secRaw = join("sec").toUpperCase();
      const remark = remark2 || remark1;
      const flags = [];
      let caseNo = a.caseNo;
      if (!known.has(caseNo)) {
        const num = caseNo.split(" ")[1].split("/")[0];
        let alt = [...known].filter(k => k.split(" ")[1] && k.split(" ")[1].split("/")[0] === num);
        if (!alt.length) alt = nearestKnown(caseNo, known);
        if (alt.length === 1) { flags.push(`Photo read ${caseNo} → taken as ${alt[0]}`); caseNo = alt[0]; }
        else if (alt.length > 1) flags.push(`Not in our list – could be ${alt.slice(0, 3).join(" / ")}`);
        else flags.push("Case not in our list – check number");
      }
      if (a.conf < 70 && !known.has(caseNo)) flags.push("Case number unclear in photo");
      let C = classifyRemark(remark2 || remark1);
      let ds = remarkDates(remark2, listDate || "2000-01-01");
      let weak = false;
      if (!ds.some(d => d.iso && !d.fixed)) { const d1 = remarkDates(remark1, listDate || "2000-01-01"); if (d1.some(d => d.iso)) { ds = d1; weak = true; } }
      let next = "";
      const good = ds.filter(d => d.iso);
      const wk = (remark2 + " " + remark1).match(/WEEK\W*\w{0,3}\W*(\d{1,2}\s*[.,]\s*\d{1,2}\s*[.,]\s*\d{4})/i);
      const wkd = wk ? remarkDates(wk[1], listDate || "2000-01-01")[0] : null;
      if (C === "C6") next = wkd && wkd.iso ? wkd.iso : (good[0] ? good[0].iso : ""); else if (C !== "C5") next = good.length ? good[good.length - 1].iso : "";
      if (ds.some(d => d.fixed)) flags.push("Year corrected from " + ds.filter(d => d.fixed).map(d => d.fixed).join(", "));
      if (ds.some(d => d.bad)) flags.push("Date unclear: " + ds.filter(d => d.bad).map(d => d.bad).join(", "));
      if (C === "C5" && good.length) {   // closed cases have no next date – text from the row above probably leaked in
        C = "C1"; next = good[good.length - 1].iso; flags.push("Says CLOSED but has a next date – check");
      }
      if (C !== "C5" && !next) flags.push("No next date found");
      else if (weak && C !== "C5") flags.push("Date not clear in photo – check");
      const sc = secRaw.replace(/[^A-Z(]/g, "");
      const sec = (sc.match(/^(DP|[A-Z])[(I1L]?S[EC]/) || sc.match(/^(DP|[A-Z])/) || [])[1] || "";
      const kSec = opts.knownSec && opts.knownSec[caseNo];
      if (sec && kSec && kSec.toUpperCase().replace(/[^A-Z]/g, "") !== sec && !TTD_SEC.test(secRaw)) flags.push(`Section: photo ${sec}, our list ${kSec}`);
      rows.push({
        court: a.court, item: join("item").replace(/[^0-9.\-, ]/g, "").trim(), caseNo, caseRaw: a.caseRaw,
        stage: join("stage"), district: join("dist").replace(/[^A-Za-z .]/g, "").trim(), sec: TTD_SEC.test(secRaw) ? "TTD" : sec,
        ttd: TTD_SEC.test(secRaw), remark, C, next, flags
      });
    }
    return { listDate, rows, colsGuessed };
  }

  // several readings of the same photo → keep, per case, the reading that matches a known case with fewest doubts
  function mergeResultReadings(readings, known) {
    const best = new Map(), order = [];
    const score = r => (known.has(r.caseNo) ? 10 : 0) - r.flags.length + (r.next ? 1 : 0);
    for (const rows of readings) for (const r of rows) {
      const cur = best.get(r.caseNo);
      if (!cur) { best.set(r.caseNo, r); order.push(r.caseNo); continue; }
      let pick = score(r) > score(cur) ? r : cur, other = pick === r ? cur : r;
      if (!pick.next && other.next && pick.C === other.C) pick = { ...pick, next: other.next, flags: [...other.flags.filter(f => /Date|date/.test(f)), ...pick.flags.filter(f => !/next date/i.test(f))] };
      best.set(r.caseNo, pick);
    }
    // a misread number that was not matched to a known case is dropped when the known case was found in another reading at the same court/item
    const out = order.map(k => best.get(k));
    return out.filter(r => known.has(r.caseNo) || !out.some(o => o !== r && known.has(o.caseNo) && o.court === r.court && o.item && o.item === r.item));
  }


  // apply checked result rows to the team data.json (same shape the app reads)
  function applyResults(old, listDate, rows, nowIso) {
    const cases = new Map((old.cases || []).map(c => [c.k, { ...c, ds: [...(c.ds || (c.d ? [c.d] : []))] }]));
    const status = { ...(old.status || {}) };
    let added = 0, updated = 0;
    const dm = listDate ? listDate.slice(8, 10) + "-" + listDate.slice(5, 7) : "";
    for (const r of rows) {
      const k = keyOf(r.caseNo); let c = cases.get(k);
      if (!c) { c = { k, n: r.caseNo, ds: [] }; cases.set(k, c); added++; } else updated++;
      const ds = new Set(c.ds); if (listDate) ds.add(listDate); if (r.next) ds.add(r.next);
      c.ds = [...ds].sort(); c.d = c.ds[c.ds.length - 1];
      if (r.sec) c.s = r.sec.toUpperCase();
      if (r.district && !c.dist) c.dist = r.district;
      if (r.court) c.ct = `Court ${r.court}` + (r.item ? ` · Item ${r.item}` : "");
      const st = { ...(status[k] || {}), C: r.C, note: (dm ? dm + ": " : "") + (r.remark || "").trim() };
      if (r.next) st.nd = r.next; else delete st.nd;
      status[k] = st;
    }
    const list = [...cases.values()].sort((a, b) => a.n.localeCompare(b.n));
    return { data: { ...old, meta: { ...(old.meta || {}), at: nowIso, count: list.length, results: listDate }, cases: list, status }, added, updated };
  }

  /* ---------- match + segregate ---------- */

  const COUNTER_LABEL = { filed: "Counter filed", notfiled: "Counter NOT filed", proforma: "Formal party", blank: "Status blank in endowmentsinfo", missing: "Not in endowmentsinfo" };

  function buildReport(cl, ei, opts = {}) {
    const listed = [], ttdOut = [], possible = [];
    for (const it of cl.items) {
      const rec = ei.map.get(it.caseNo) || it.lc.map(x => ei.map.get(x)).find(Boolean);
      const known = opts.known && opts.known.has(it.caseNo);   // already in the team app (Legal Cell list)
      const isOurs = it.endow || !!ei.map.get(it.caseNo) || known || (opts.include && opts.include.has(it.caseNo));
      if (!isOurs) {
        if (!it.ttd && TEMPLE_RE.test(it.text) && !possible.some(p => p.caseNo === it.caseNo))
          possible.push({ caseNo: it.caseNo, court: it.court, item: it.sno, district: it.district, text: it.text.slice(0, 220) });
        continue;
      }
      if (opts.exclude && opts.exclude.has(it.caseNo)) continue;
      if (it.ttd || (rec && rec.ttd)) { ttdOut.push({ ...it, rec }); continue; }
      const dup = listed.find(x => x.caseNo === it.caseNo && x.court === it.court);
      if (dup) { if (!String(dup.item).split(", ").includes(it.sno)) dup.item += ", " + it.sno; continue; }
      const direct = ei.map.get(it.caseNo);
      const secFix = opts.secFix && opts.secFix[it.caseNo];
      const counter = direct ? direct.counter : "missing";
      listed.push({
        caseNo: it.caseNo, court: it.court, item: it.sno, district: it.district || (direct && direct.district) || "",
        stage: it.stage, lc: it.lc, sec: secFix || (direct ? direct.sec : (rec ? rec.sec : "")), temple: (direct || rec || {}).temple || "",
        counter, counterLabel: COUNTER_LABEL[counter], whoFiles: direct ? direct.whoFiles : "", inSite: !!direct,
        viaConnected: !direct && !!rec, knownOnly: !it.endow && !direct && !!known, contempt: /^CC /.test(it.caseNo), endowText: it.endow
      });
    }
    const by = f => listed.filter(f);
    const seg = {
      listed,
      filed: by(c => c.counter === "filed"),
      toFile: by(c => c.counter === "notfiled" || c.counter === "blank"),
      notNeeded: by(c => c.counter === "proforma"),
      notInSite: by(c => c.counter === "missing"),
      noSection: by(c => !c.sec),
      contempt: by(c => c.contempt),
      possible,
      ttd: ttdOut.map(t => ({ caseNo: t.caseNo, court: t.court, item: t.sno, district: t.district }))
    };
    return { date: cl.date, seg };
  }

  /* ---------- merge into the team app's data.json (same format) ---------- */

  function mergeData(old, report, ei, nowIso) {
    const cases = new Map((old && old.cases || []).map(c => [c.k, { ...c, ds: [...(c.ds || (c.d ? [c.d] : []))] }]));
    const status = { ...(old && old.status || {}) };
    const day = report.date;
    let added = 0, updated = 0;
    for (const c of report.seg.listed) {
      const k = keyOf(c.caseNo);
      let e = cases.get(k);
      if (!e) { e = { k, n: c.caseNo, ds: [] }; cases.set(k, e); added++; } else updated++;
      if (day && !e.ds.includes(day)) e.ds.push(day);
      e.ds.sort(); e.d = e.ds[e.ds.length - 1];
      if (c.sec) e.s = c.sec.toUpperCase();
      if (c.temple) e.sub = c.temple;
      if (c.district) e.dist = c.district;
      e.ct = `Court ${c.court} · Item ${c.item}`;
      e.site = c.counterLabel;
    }
    // refresh counter status / section of every known case from the fresh Excel
    for (const e of cases.values()) {
      const rec = ei.map.get(e.n);
      if (!rec) continue;
      e.site = COUNTER_LABEL[rec.counter];
      if (rec.sec) e.s = rec.sec.toUpperCase();
      if (rec.temple && !e.sub) e.sub = rec.temple;
    }
    const list = [...cases.values()].sort((a, b) => a.n.localeCompare(b.n));
    return {
      data: { meta: { at: nowIso, count: list.length, src: "Dashboard", causeListDate: day }, cases: list, status },
      added, updated
    };
  }

  /* ---------- outgoing tables ---------- */

  function tables(report, data) {
    const L = report.seg.listed;
    const note = c => (data.status[keyOf(c.caseNo)] || {}).note || "";
    const code = c => (data.status[keyOf(c.caseNo)] || {}).C || "";
    const base = c => ({ "Case": c.caseNo, "Court": c.court, "Item": c.item, "Section": c.sec || "—", "District": c.district, "Counter": c.counterLabel });
    const bySec = {};
    for (const c of L) (bySec[(c.sec || "No section").toUpperCase()] ||= []).push(c);
    const urgent = L.filter(c => c.contempt || /APPEARANCE|IMPORTANT|FINALLY/i.test(note(c)) || (c.counter === "notfiled" && /C2/.test(code(c))));
    return {
      alo1: L.map(c => ({ ...base(c), "Last result": code(c), "Last remark": note(c) })),
      alo2: [...report.seg.toFile].map(c => ({ ...base(c), "Counter by": c.whoFiles || "", "Last remark": note(c) })),
      sections: Object.fromEntries(Object.entries(bySec).sort().map(([s, arr]) => [s, arr.map(c => ({ ...base(c), "Temple": c.temple, "Last remark": note(c) }))])),
      deo: [
        ...report.seg.notInSite.map(c => ({ "Action": "Add to endowmentsinfo", ...base(c) })),
        ...report.seg.listed.filter(c => c.counter === "blank").map(c => ({ "Action": "Update counter status", ...base(c) })),
        ...report.seg.noSection.filter(c => c.inSite).map(c => ({ "Action": "Assign section", ...base(c) }))
      ],
      head: {
        summary: [
          ["Cases listed", L.length], ["Counter filed", report.seg.filed.length], ["Counter to be filed", report.seg.toFile.length],
          ["Counter not necessary", report.seg.notNeeded.length], ["Not in endowmentsinfo", report.seg.notInSite.length],
          ["Contempt (CC)", report.seg.contempt.length], ["TTD left out", report.seg.ttd.length]
        ],
        urgent: urgent.map(c => ({ ...base(c), "Remark": note(c) }))
      }
    };
  }

  function toWhatsApp(title, rows, cols) {
    if (!rows.length) return `*${title}*\nNone`;
    cols = cols || Object.keys(rows[0]);
    return `*${title}* (${rows.length})\n` + rows.map((r, i) => `${i + 1}. ` + cols.map(c => r[c]).filter(v => v !== "" && v != null && v !== "—").join(" | ")).join("\n");
  }

  const api = { mergeResultReadings, applyResults, parseResultsWords, classifyRemark, fixCase, fixCaseVariants, buildCasesInfo, pdfToLines, parseCauseList, parseEndowmentsInfo, buildReport, mergeData, tables, toWhatsApp, keyOf, COUNTER_LABEL };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.CCEngine = api;
})(typeof window !== "undefined" ? window : globalThis);
