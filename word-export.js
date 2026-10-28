/* Case Notice System — "Download Word" (.docx) for the Notice letter and the Case List / Cause List.
   Read-only add-on: it only READS what the screen already shows (notice paper, case table, report summary)
   and builds a Word file in the browser. It changes no data and no existing button, print or PDF behaviour.
   Library: docx (self-hosted copy from /bill-register-system/lib/, with CDN fallbacks). */
(function(){
  "use strict";
  var DOCX_SRC = ["/bill-register-system/lib/docx.iife.js",
                  "https://cdn.jsdelivr.net/npm/docx@9.7.1/dist/index.iife.js",
                  "https://unpkg.com/docx@9.7.1/dist/index.iife.js"];
  var AR_RE = /[؀-ۿ]/;
  var FONT = "Arial";
  var PAPER = { A4:{w:11906,h:16838}, Legal:{w:12240,h:20160}, Foolscap:{w:12240,h:18720} };

  function $(id){ return document.getElementById(id); }
  function toast(msg){ var t = $("toast"); if(!t){ alert(msg); return; } t.textContent = msg; t.classList.add("show"); clearTimeout(toast._h); toast._h = setTimeout(function(){ t.classList.remove("show"); }, 3200); }
  function isUr(){ return document.documentElement.getAttribute("data-ui-lang") === "ur"; }
  function addScript(src){ return new Promise(function(res, rej){ var s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = function(){ rej(new Error("load " + src)); }; document.head.appendChild(s); }); }
  async function loadDocx(){
    if(window.docx && window.docx.Document) return;
    var err = null;
    for(var i = 0; i < DOCX_SRC.length; i++){
      try{ await addScript(DOCX_SRC[i]); if(window.docx && window.docx.Document) return; }catch(e){ err = e; }
    }
    throw err || new Error("Word library could not be loaded");
  }
  function download(blob, name){
    var a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function(){ URL.revokeObjectURL(a.href); }, 4000);
  }
  function txt(el){ return el ? String(el.innerText || el.textContent || "").replace(/ /g, " ").trim() : ""; }
  function lines(s){ return String(s == null ? "" : s).split(/\r?\n/).map(function(l){ return l.replace(/\s+$/, ""); }); }
  function today(){ var d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
  function safe(s){ return String(s || "").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "file"; }

  /* ---------- small docx helpers ---------- */
  function H(D){
    var AL = D.AlignmentType;
    function run(t, o){ o = o || {}; return new D.TextRun({ text: String(t == null ? "" : t), font: FONT, size: o.size || 22, bold: !!o.bold, italics: !!o.italics, color: o.color, rightToLeft: !!o.rtl || AR_RE.test(String(t)) }); }
    function para(runs, o){ o = o || {}; return new D.Paragraph({ children: runs, alignment: o.align, spacing: { before: o.before || 0, after: o.after == null ? 0 : o.after, line: o.line }, bidirectional: !!o.rtl, border: o.border, keepNext: !!o.keepNext }); }
    function textParas(s, o){
      o = o || {};
      return lines(s).map(function(l){ var rtl = o.rtl != null ? o.rtl : AR_RE.test(l); return para([run(l, { size: o.size, bold: o.bold, color: o.color, rtl: rtl })], { align: o.align, rtl: rtl, after: o.after, line: o.line }); });
    }
    function cell(children, w, o){
      o = o || {};
      return new D.TableCell({ children: children, width: { size: w, type: D.WidthType.DXA }, verticalAlign: o.valign || D.VerticalAlign.TOP,
        shading: o.fill ? { type: D.ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
        margins: { top: o.mv == null ? 70 : o.mv, bottom: o.mv == null ? 70 : o.mv, left: 90, right: 90 }, columnSpan: o.span });
    }
    return { AL: AL, run: run, para: para, textParas: textParas, cell: cell };
  }
  function ruleParas(D){
    return [
      new D.Paragraph({ spacing: { before: 40, after: 0, line: 20 }, border: { bottom: { style: D.BorderStyle.SINGLE, size: 18, color: "B8860B", space: 1 } }, children: [new D.TextRun({ text: "", size: 2 })] }),
      new D.Paragraph({ spacing: { before: 0, after: 120, line: 20 }, border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: "1B3F63", space: 1 } }, children: [new D.TextRun({ text: "", size: 2 })] })
    ];
  }
  function footerOf(D, rtl){
    return new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [
      new D.TextRun({ text: "Camp Office Prosecutor General Punjab  |  Supreme Court of Pakistan, Islamabad  |  Generated " + today() + "  |  Page ", font: FONT, size: 14, color: "666666" }),
      new D.TextRun({ children: [D.PageNumber.CURRENT], font: FONT, size: 14, bold: true, color: "1B3F63" }),
      new D.TextRun({ text: " of ", font: FONT, size: 14, color: "666666" }),
      new D.TextRun({ children: [D.PageNumber.TOTAL_PAGES], font: FONT, size: 14, bold: true, color: "1B3F63" })
    ] })] });
  }

  /* ---------- 1) the Notice letter ---------- */
  function readNotice(){
    var paper = $("noticePaper"); if(!paper || !paper.firstElementChild) return null;
    var rtl = paper.classList.contains("rtl") || paper.getAttribute("dir") === "rtl";
    var meta = [];
    Array.prototype.forEach.call(paper.querySelectorAll(".notice-meta > div"), function(d){
      var lab = d.querySelector(".label"); var l = txt(lab); var full = txt(d);
      var v = lab ? full.slice(l.length).replace(/^\s*:?\s*/, "") : full;
      meta.push([l.replace(/\s*:\s*$/, ""), v]);
    });
    var extra = [];
    Array.prototype.forEach.call(paper.querySelectorAll(".pb-print-extra-fields .hf-row"), function(r){ extra.push(txt(r).replace(/\s*\n\s*/g, "      ")); });
    return {
      rtl: rtl,
      title: txt(paper.querySelector(".notice-header h2")),
      sub: txt(paper.querySelector(".notice-header .sub")),
      extra: extra, meta: meta,
      body: String((paper.querySelector(".notice-body") || {}).innerText || "").replace(/ /g, " ").trim(),
      sender: txt(paper.querySelector(".notice-sender")),
      stamp: txt(paper.querySelector(".stamp"))
    };
  }
  async function buildNoticeDocx(N, paperKey){
    var D = window.docx, h = H(D), AL = h.AL, pg = PAPER[paperKey] || PAPER.A4;
    var M = 1134, TOTAL_W = pg.w - 2 * M;
    var rtl = N.rtl, kids = [];
    kids.push(h.para([h.run(N.title, { size: 32, bold: true, color: "1B3F63", rtl: rtl })], { align: AL.CENTER, after: 40, rtl: rtl }));
    kids.push(h.para([h.run(N.sub, { size: 19, color: "555555", rtl: rtl })], { align: AL.CENTER, after: 0, rtl: rtl }));
    kids = kids.concat(ruleParas(D));
    N.extra.forEach(function(l){ kids.push(h.para([h.run(l, { size: 18, color: "444444", rtl: rtl })], { after: 30, rtl: rtl })); });
    if(N.extra.length) kids.push(h.para([h.run("")], { after: 60 }));
    if(N.meta.length){
      var bd = { style: D.BorderStyle.SINGLE, size: 4, color: "BFCAD9" };
      var w1 = Math.round(TOTAL_W * 0.34), w2 = TOTAL_W - w1;
      kids.push(new D.Table({
        width: { size: TOTAL_W, type: D.WidthType.DXA }, columnWidths: [w1, w2], layout: D.TableLayoutType.FIXED, visuallyRightToLeft: rtl,
        borders: { top: bd, bottom: bd, left: bd, right: bd, insideHorizontal: bd, insideVertical: bd },
        rows: N.meta.map(function(m){
          return new D.TableRow({ cantSplit: true, children: [
            h.cell([h.para([h.run(m[0], { size: 19, bold: true, color: "1B3F63", rtl: rtl })], { rtl: rtl })], w1, { fill: "F2F5FA" }),
            h.cell(h.textParas(m[1] || "—", { size: 20, rtl: AR_RE.test(m[1]) ? true : (rtl && !/[A-Za-z]/.test(m[1])) }), w2)
          ] });
        })
      }));
      kids.push(h.para([h.run("")], { after: 120 }));
    }
    h.textParas(N.body, { size: 23, rtl: rtl, align: rtl ? AL.RIGHT : AL.JUSTIFIED, after: 120, line: 320 }).forEach(function(p){ kids.push(p); });
    kids.push(h.para([h.run("")], { before: 240 }));
    var NONE = { style: D.BorderStyle.NONE, size: 0, color: "FFFFFF" };
    var half = Math.floor(TOTAL_W / 2);
    var senderKids = h.textParas(N.sender, { size: 21, rtl: rtl, align: rtl ? AL.RIGHT : AL.LEFT });
    var stampKids = N.stamp ? h.textParas(N.stamp, { size: 21, bold: true, rtl: rtl, align: AL.CENTER }) : [h.para([h.run("")])];
    kids.push(new D.Table({
      width: { size: TOTAL_W, type: D.WidthType.DXA }, columnWidths: [half, TOTAL_W - half], layout: D.TableLayoutType.FIXED, visuallyRightToLeft: rtl,
      borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
      rows: [new D.TableRow({ cantSplit: true, children: [h.cell(senderKids, half), h.cell(stampKids, TOTAL_W - half, { valign: D.VerticalAlign.CENTER })] })]
    }));
    var doc = new D.Document({
      creator: "Camp Office, Prosecutor General Punjab", title: N.title || "Notice",
      sections: [{ properties: { page: { size: { width: pg.w, height: pg.h }, margin: { top: 1000, bottom: 1000, left: M, right: M, footer: 400 } } }, footers: { default: footerOf(D, rtl) }, children: kids }]
    });
    return await D.Packer.toBlob(doc);
  }

  /* ---------- 2) the Case List / Cause List + summary ---------- */
  function hiddenCols(){
    var st = $("caseListColStyle"), out = {};
    var m = (st && st.textContent || "").match(/\[data-col='([A-Za-z]+)'\]/g) || [];
    m.forEach(function(x){ out[x.replace(/.*'([A-Za-z]+)'.*/, "$1")] = true; });
    return out;
  }
  function readList(){
    var body = $("caseTableBody"); if(!body) return null;
    var table = body.closest("table"), hid = hiddenCols(), cols = [], idx = [];
    Array.prototype.forEach.call(table.querySelectorAll("thead th"), function(th, i){
      var k = th.getAttribute("data-col"); if(!k || hid[k]) return;
      cols.push({ key: k, label: txt(th) }); idx.push(i);
    });
    var rows = [];
    Array.prototype.forEach.call(body.querySelectorAll("tr"), function(tr){
      if(tr.classList.contains("empty-row")) return;
      var tds = tr.querySelectorAll("td"); if(!tds.length || tds.length < idx.length) return;
      rows.push(idx.map(function(i){ return txt(tds[i]); }));
    });
    var sum = [];
    Array.prototype.forEach.call(document.querySelectorAll("#reportBenchBody tr"), function(tr){ if(tr.classList.contains("empty-row")) return; var t = tr.querySelectorAll("td"); if(t.length >= 2) sum.push([txt(t[0]), txt(t[1])]); });
    var tot = document.querySelectorAll("#reportBenchTotals td");
    return { cols: cols, rows: rows, sum: sum, total: tot.length >= 2 ? txt(tot[1]) : String(rows.length),
             from: ($("f_from") || {}).value || "", to: ($("f_to") || {}).value || "",
             sumHead: txt(document.querySelector("#reportCard thead th")) || "Bench", sumCount: txt(document.querySelectorAll("#reportCard thead th")[1]) || "Cases" };
  }
  async function buildListDocx(L, paperKey){
    var D = window.docx, h = H(D), AL = h.AL, pg = PAPER[paperKey] || PAPER.A4;
    var landscape = L.cols.length > 5;
    var pw = landscape ? pg.h : pg.w, ph = landscape ? pg.w : pg.h, M = 567, TOTAL_W = pw - 2 * M;
    var ur = isUr(), kids = [];
    kids.push(h.para([h.run("Camp Office Prosecutor General Punjab", { size: 28, bold: true, color: "1B3F63" })], { align: AL.CENTER, after: 20 }));
    kids.push(h.para([h.run("Supreme Court of Pakistan, Islamabad", { size: 19, color: "555555" })], { align: AL.CENTER }));
    kids = kids.concat(ruleParas(D));
    kids.push(h.para([h.run(ur ? "کیس فہرست / کاز لسٹ اور خلاصہ" : "CASE LIST / CAUSE LIST & SUMMARY", { size: 26, bold: true, color: "1B3F63", rtl: ur })], { align: AL.CENTER, before: 40, after: 40, rtl: ur }));
    var sub = "Fixed date: " + (L.from || "All-time") + " to " + (L.to || "today") + "   |   Cases: " + L.total + "   |   Date: " + today();
    kids.push(h.para([h.run(sub, { size: 18, color: "555555" })], { align: AL.CENTER, after: 120 }));
    var w = L.cols.map(function(c){ return c.key === "title" || c.key === "vs" ? 2.2 : (c.key === "serial" || c.key === "caseType" || c.key === "firNo" || c.key === "bench" ? 0.9 : 1.2); });
    var ws = w.reduce(function(a, b){ return a + b; }, 0);
    var cw = w.map(function(x){ return Math.floor(TOTAL_W * x / ws); });
    var bd = { style: D.BorderStyle.SINGLE, size: 4, color: "BFCAD9" };
    var head = new D.TableRow({ tableHeader: true, cantSplit: true, children: L.cols.map(function(c, i){
      return h.cell([h.para([h.run(c.label, { size: 16, bold: true, color: "FFFFFF" })], { align: AL.CENTER })], cw[i], { fill: "2A5A84", valign: D.VerticalAlign.CENTER, mv: 60 });
    }) });
    var body = L.rows.map(function(r, k){
      return new D.TableRow({ cantSplit: true, children: r.map(function(v, i){ return h.cell(h.textParas(v, { size: 16, color: "1C3363" }), cw[i], { fill: (k % 2) ? "F4F7FB" : undefined, mv: 50 }); }) });
    });
    kids.push(new D.Table({ rows: [head].concat(body), width: { size: TOTAL_W, type: D.WidthType.DXA }, columnWidths: cw, layout: D.TableLayoutType.FIXED,
      borders: { top: bd, bottom: bd, left: bd, right: bd, insideHorizontal: bd, insideVertical: bd } }));
    if(L.sum.length){
      kids.push(h.para([h.run(ur ? "رپورٹ — بینچ وار خلاصہ" : "REPORT — SUMMARY BY BENCH", { size: 22, bold: true, color: "1B3F63", rtl: ur })], { before: 280, after: 80, rtl: ur, keepNext: true }));
      var sw1 = Math.min(TOTAL_W, 6000), sw2 = 1800;
      var srow = function(a, b, o){ o = o || {}; return new D.TableRow({ cantSplit: true, children: [
        h.cell([h.para([h.run(a, { size: 18, bold: !!o.bold, color: o.color })])], sw1, { fill: o.fill }),
        h.cell([h.para([h.run(b, { size: 18, bold: !!o.bold, color: o.color })], { align: AL.RIGHT })], sw2, { fill: o.fill })] }); };
      var srows = [srow(L.sumHead, L.sumCount, { bold: true, color: "FFFFFF", fill: "2A5A84" })];
      L.sum.forEach(function(s){ srows.push(srow(s[0], s[1])); });
      srows.push(srow("Total", L.total, { bold: true, fill: "E8EEF6" }));
      kids.push(new D.Table({ rows: srows, width: { size: sw1 + sw2, type: D.WidthType.DXA }, columnWidths: [sw1, sw2], layout: D.TableLayoutType.FIXED,
        borders: { top: bd, bottom: bd, left: bd, right: bd, insideHorizontal: bd, insideVertical: bd } }));
    }
    var doc = new D.Document({
      creator: "Camp Office, Prosecutor General Punjab", title: "Case List / Cause List",
      sections: [{ properties: { page: { size: { width: pg.w, height: pg.h, orientation: landscape ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT }, margin: { top: 700, bottom: 760, left: M, right: M, footer: 300 } } }, footers: { default: footerOf(D, false) }, children: kids }]
    });
    return await D.Packer.toBlob(doc);
  }

  /* ---------- buttons ---------- */
  var busy = false;
  async function doNotice(){
    if(busy) return; busy = true;
    try{
      var overlay = $("noticeOverlay");
      var sel = $("n_paperSize"); var key = sel ? sel.value : "A4";
      /* make sure the paper shows the current choices (language, fields) — same as the Print/PDF buttons do */
      var N = readNotice(); if(!N){ toast(isUr() ? "پہلے نوٹس کھولیں۔" : "Please open a notice first."); return; }
      await loadDocx();
      var blob = await buildNoticeDocx(N, key);
      download(blob, "CampOffice-Notice-" + safe(overlay && overlay.dataset.caseId) + ".docx");
    }catch(e){ console.error(e); toast((isUr() ? "Word فائل نہیں بن سکی: " : "Could not build the Word file: ") + (e && e.message || e)); }
    finally{ busy = false; }
  }
  async function doList(){
    if(busy) return; busy = true;
    try{
      var L = readList();
      if(!L || !L.rows.length){ toast(isUr() ? "موجودہ فلٹر میں کوئی کیس نہیں۔" : "No cases match the current filters."); return; }
      await loadDocx();
      var sel = $("globalPaperSize"); var blob = await buildListDocx(L, sel ? sel.value : "A4");
      download(blob, "CampOffice-CauseList-" + today() + ".docx");
    }catch(e){ console.error(e); toast((isUr() ? "Word فائل نہیں بن سکی: " : "Could not build the Word file: ") + (e && e.message || e)); }
    finally{ busy = false; }
  }
  function label(){ return isUr() ? "📝 Word ڈاؤن لوڈ کریں" : "📝 Download Word"; }
  function addBtn(afterId, id, cls, handler){
    var ref = $(afterId); if(!ref || $(id)) return;
    var b = document.createElement("button"); b.type = "button"; b.id = id; b.className = cls; b.setAttribute("data-wx", "1");
    b.title = "Downloads an editable Word (.docx) file"; b.textContent = label(); b.addEventListener("click", handler);
    ref.parentNode.insertBefore(b, ref.nextSibling);
  }
  function init(){
    addBtn("n_pdfBtn", "n_wordBtn", "btn ghost", doNotice);
    addBtn("r_pdfBtn", "r_wordBtn", "btn ghost", doList);
    addBtn("globalPdfBtn", "globalWordBtn", "", doList);
    new MutationObserver(function(){ Array.prototype.forEach.call(document.querySelectorAll("[data-wx]"), function(b){ b.textContent = label(); }); })
      .observe(document.documentElement, { attributes: true, attributeFilter: ["data-ui-lang"] });
  }
  window.__cnWord = { readNotice: readNotice, readList: readList, buildNoticeDocx: buildNoticeDocx, buildListDocx: buildListDocx, loadDocx: loadDocx };
  if(document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
