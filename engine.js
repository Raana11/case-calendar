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

  const api = { buildCasesInfo, pdfToLines, parseCauseList, parseEndowmentsInfo, buildReport, mergeData, tables, toWhatsApp, keyOf, COUNTER_LABEL };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.CCEngine = api;
})(typeof window !== "undefined" ? window : globalThis);
