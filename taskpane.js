/*  SAZAN Distribution Map - Excel task pane
    Reads and writes the equipment registry named  Equip_Registry  (Dashboard!A100:J349)
    Columns (0-based):  0 Tag  1 Type  2 Voltage  3 Fed From  5 Demand kVA
                        6 Sheet  8 Feed Type  9 Ckt / space                */

var REG_SHEET = "Dashboard";
var REG_RANGE = "A101:J350";   /* fallback only - see regRange() below */
var REG_NAME  = "Equip_Registry";   /* authoritative - resolved at runtime */

/*  Resolve the registry from its defined name so this file cannot drift out of
    sync when the registry is resized in the master. REG_RANGE is used only if
    the name is missing.                                                     */
function regRange(ctx) {
  var n = ctx.workbook.names.getItemOrNullObject(REG_NAME);
  n.load("isNullObject");
  return ctx.sync().then(function () {
    return n.isNullObject
      ? ctx.workbook.worksheets.getItem(REG_SHEET).getRange(REG_RANGE)
      : n.getRange();
  });
}
var C_TAG = 0, C_TYPE = 1, C_VOLT = 2, C_FED = 3, C_MAINS = 4, C_KVA = 5, C_SHEET = 6, C_FT = 8, C_CKT = 9;
var MAXDEPTH = 12;

var items = [];          // { tag, type, volt, fed, kva, sheet, ft, ckt, row }
var xfmrKva = {};        // v2.10: T-tag -> { std, load } from the Transformers schedule
var collapsed = {};      // tag -> true when its children are hidden
var selTag = "";

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) return;
  applyTheme();
  document.getElementById("btnRefresh").onclick = refresh;
  document.getElementById("btnExpand").onclick = function () { collapsed = {}; render(); };
  document.getElementById("btnCollapse").onclick = function () {
    items.forEach(function (it) { if (hasChildren(it.tag)) collapsed[it.tag] = true; });
    render();
  };
  document.getElementById("btnApply").onclick = applyLink;
  document.getElementById("btnUnlink").onclick = unlink;
  document.getElementById("btnGoto").onclick = gotoSheet;
  document.getElementById("btnAdd").onclick = addEquipment;
  document.getElementById("btnRename").onclick = renameSelected;
  document.getElementById("btnRepair").onclick = repairSelected;
  refresh();
});

/*  ---------- theme: follow the Excel UI theme ---------- */
function applyTheme() {
  try {
    var t = Office.context.officeTheme;
    if (!t || !t.bodyBackgroundColor) return;
    var hex = String(t.bodyBackgroundColor).replace("#", "");
    var r = parseInt(hex.substr(0, 2), 16), g = parseInt(hex.substr(2, 2), 16), b = parseInt(hex.substr(4, 2), 16);
    var lum = (0.299 * r + 0.587 * g + 0.114 * b);
    if (lum < 128) document.body.classList.add("dark");
  } catch (e) { /* light stays default */ }
}

function setStatus(msg) { document.getElementById("status").textContent = msg || ""; }

/*  ---------- read the registry ---------- */
function refresh() {
  return Excel.run(function (ctx) {
    return (typeof SAZ !== "undefined" && SAZ.healRegistryOwnKva ? SAZ.healRegistryOwnKva(ctx).catch(function () { return null; }) : Promise.resolve(null)).then(function () { return regRange(ctx); }).then(function (rg) {   /* v2.10: refresh rating / demand links first */
      rg.load("values");
      /* Utility service transformers live on the Dashboard (Util_Tags row, up to
         four sources), not in the registry. Each is shown as a tree root so main
         boards can be linked to it; a Fed From equal to any source tag resolves
         to 'Source' in the Feeder Schedule.  */
      var uN = ["Util_Tags", "Util_KVA", "Cfg_Volt"].map(function (nm) {
        var n = ctx.workbook.names.getItemOrNullObject(nm); n.load("isNullObject"); return n;
      });
      return ctx.sync().then(function () {
        var uR = uN.map(function (n) {
          if (n.isNullObject) return null;
          var r = n.getRange(); r.load("values"); return r;
        });
        return ctx.sync().then(function () {
          var uT = uR[0] ? uR[0].values[0] : [], uK = uR[1] ? uR[1].values[0] : [];
          var uVolt = uR[2] ? String(uR[2].values[0][0] || "") : "";
          var utils = uT.map(function (t, k) {
            return { tag: String(t || "").trim(), kva: Number(uK[k]) || 0, volt: uVolt };
          }).filter(function (u) { return u.tag; });
          items = [];
          rg.values.forEach(function (row, i) {
            var tag = String(row[C_TAG] === null ? "" : row[C_TAG]).trim();
            if (!tag) return;
            items.push({
              tag: tag,
              type: String(row[C_TYPE] || "").trim(),
              volt: String(row[C_VOLT] || "").trim(),
              fed: String(row[C_FED] || "").trim(),
              kva: Number(row[C_KVA]) || 0,
              sheet: String(row[C_SHEET] || "").trim(),
              ft: String(row[C_FT] || "").trim(),
              ckt: String(row[C_CKT] || "").trim(),
              mains: String(row[C_MAINS] || "").trim(),
              row: i                                   // offset inside the registry
            });
          });
          utils.slice().reverse().forEach(function (util) {
            if (!byTag(util.tag)) items.unshift({ tag: util.tag, type: "Utility Transformer", volt: util.volt,
              fed: "", kva: util.kva, sheet: "Dashboard", ft: "", ckt: "", row: -1, head: true });
          });
          var altR = ctx.workbook.worksheets.getItem(REG_SHEET).getRange("L101:W350"); altR.load("values");   /* L Alt Source | M own kVA | N UPS out V | O Main/Bus override (A) | P rating (A) | Q demand (A) | R:W feeder inputs (v2.11) */
          var xfS = ctx.workbook.worksheets.getItemOrNullObject("Transformers"); xfS.load("isNullObject");
          return ctx.sync().then(function () {
            altR.values.forEach(function (r, k) { items.forEach(function (x) { if (x.row !== k) return; var a = String(r[0] || "").trim(); if (a) x.alt = a; x.outv = String(r[2] || "").trim(); x.ovr = Number(r[3]) || 0; x.ampR = Number(r[4]) || 0; x.ampD = Number(r[5]) || 0; x.len = r[6] === "" ? "" : r[6]; x.lenPri = r[7] === "" ? "" : r[7]; x.lenEm = r[8] === "" ? "" : r[8]; x.mat = String(r[9] || "").trim().toUpperCase(); x.race = String(r[10] || "").trim(); x.exist = /^y/i.test(String(r[11] || "")) ? "Y" : ""; }); });
            xfmrKva = {};
            if (xfS.isNullObject) return null;
            var xfR = xfS.getRange("A6:G105"); xfR.load("values");
            return ctx.sync().then(function () { xfR.values.forEach(function (r) { var t = String(r[0] || "").trim(); if (t) xfmrKva[t] = { std: r[6], load: Number(r[5]) || 0 }; }); });
          }).then(function () {
            addTransformerNodes();
            addEmergencyNodes();
            render();
            fillFedList();
            if (selTag) select(selTag); else clearProps();
            setStatus("");
          });
        });
      });
    });
  }).catch(function (e) { setStatus("Read failed: " + e.message); });
}

/*  v2.3: a transformer sits between a parent and a child whose line voltages
    differ. The Transformers schedule derives a T-<child> row from exactly that
    registry edge, so the tree shows the same node: parent -> T-<child> -> child.
    The node is display-only (row -1, virtual); the child's real Fed From in the
    registry stays the parent, which is what the schedule's detection reads.  */
function leadVolts(s) { var m = /([0-9]+(?:\.[0-9]+)?)/.exec(String(s || "")); return m ? Number(m[1]) : 0; }
/*  v2.10: what the map shows beside each tag.
    Boards: Main/Bus rating and demand amps (% of rating); an override is marked "ovr".
    Transformers: standard kVA and load kVA.  UPS / transfer switch: their rating.
    Utility sources: kVA.  Anything else: demand kVA as before.  */
function nodeRating(it) {
  var f1 = function (n) { return (Math.round(n * 10) / 10).toLocaleString(); };
  if (it.head) return it.kva ? "   " + f1(it.kva) + " kVA" : "";
  if (/^transformer/i.test(String(it.type || "")) || /^T-/.test(it.tag)) {
    var x = xfmrKva[it.tag]; if (!x) return it.kva ? "   " + f1(it.kva) + " kVA" : "";
    var std = Number(x.std);
    return "   " + (std ? f1(std) + " kVA" : String(x.std || "") + " kVA") + (x.load ? " \u00B7 " + f1(x.load) + " kVA load" : "");
  }
  if (it.ampR) {
    var pct = it.ampD ? Math.round(it.ampD / it.ampR * 100) : 0;
    return "   " + f1(it.ampR) + " A" + (it.ovr ? " ovr" : "") + (it.ampD ? " \u00B7 " + f1(it.ampD) + " A dem (" + pct + "%)" : "");
  }
  if (it.mains && /^(UPS|ATS|MTS|Transfer)/i.test(String(it.type || ""))) { var m = String(it.mains); return "   " + (/^[0-9.]+$/.test(m) ? m + " A" : m); }
  return it.kva ? "   " + f1(it.kva) + " kVA" : "";
}
function addTransformerNodes() {
  items.slice().forEach(function (it) {
    if (it.virtual || !it.fed) return;
    var p = byTag(it.fed);
    if (!p) return;
    var pv = leadVolts(p.outv || p.volt), cv = leadVolts(it.volt);   /* v2.10: a UPS with an internal transformer feeds at its output voltage */
    if (!pv || !cv || Math.abs(pv - cv) <= 1) return;
    var t = "T-" + it.tag;
    if (!byTag(t)) {
      items.push({ tag: t, type: "Transformer", volt: pv + "V \u2192 " + cv + "V", fed: it.fed,
        kva: 0, sheet: "Transformers", ft: "auto", ckt: it.ckt, row: -1, virtual: true });
    }
    it.treeFed = t;
  });
}

/*  v2.8: a transfer switch has two sources. Its registry Fed From is the normal
    source (utility / upstream board); column L names the generator. The map
    shows the generator as a source with a display-only emergency link down to
    the switch, so the switch reads as fed from both directions.            */
function addEmergencyNodes() {
  items.slice().forEach(function (it) {
    if (it.virtual || !it.alt) return;
    var t = "\u21C4 " + it.tag + " (emergency)";
    if (!byTag(t)) {
      items.push({ tag: t, type: "Emergency feed \u2192 " + it.tag, volt: it.volt, fed: it.alt,
        kva: it.kva, sheet: "Generator", ft: "auto", ckt: "", row: -1, virtual: true, emerg: it.tag });
    }
  });
}

/*  ---------- tree helpers ---------- */
function byTag(tag) {
  for (var i = 0; i < items.length; i++) if (items[i].tag === tag) return items[i];
  return null;
}
function childrenOf(tag) {
  return items.filter(function (it) { return (it.treeFed || it.fed) === tag; });
}
function hasChildren(tag) { return childrenOf(tag).length > 0; }
function isRoot(it) { return !it.fed || !byTag(it.fed); }

function descendants(tag, depth) {
  depth = depth || 0;
  if (depth > MAXDEPTH) return [];
  var out = [];
  childrenOf(tag).forEach(function (c) {
    out.push(c);
    out = out.concat(descendants(c.tag, depth + 1));
  });
  return out;
}
function downstreamKva(tag) {
  return descendants(tag).reduce(function (s, c) { return s + c.kva; }, 0);
}
function statusOf(it) {
  if (it.emerg) return "\u21C4 Emergency source for " + it.emerg + " (Generator sheet)";
  if (it.alt) return "\u2714 Normal: " + (it.fed || "(none)") + "  \u00B7  Emergency: " + it.alt;
  if (/^UPS$/i.test(String(it.type || ""))) return "\u2714 UPS  \u00B7  Input: " + (it.fed || "(none)") + "  \u00B7  sized on UPS \u25B8 UPS SYSTEMS";
  if (it.virtual) return "\u2714 Auto (Transformers schedule)";
  if (!it.fed) return "\u25C6 Source";
  if (!byTag(it.fed)) return "\u26A0 Parent not in registry";
  var chain = [], p = it, guard = 0;
  while (p && p.fed && guard++ < MAXDEPTH) {
    if (p.fed === it.tag) return "\u26A0 Circular feed";
    chain.push(p.fed); p = byTag(p.fed);
  }
  if (it.fed && it.fed !== "(unassigned)" && !it.ft) return "\u26A0 Feed type not set";
  return "\u2714 OK";
}

/*  ---------- render ---------- */
function render() {
  var host = document.getElementById("tree");
  host.innerHTML = "";
  var count = 0;

  function emit(it, depth) {
    count++;
    var kids = childrenOf(it.tag);
    var div = document.createElement("div");
    div.className = "node" + (it.tag === selTag ? " sel" : "");
    div.setAttribute("data-tag", it.tag);

    var twist = document.createElement("span");
    twist.className = "twist";
    twist.textContent = kids.length ? (collapsed[it.tag] ? "+" : "-") : " ";
    twist.onclick = function (ev) {
      ev.stopPropagation();
      if (!kids.length) return;
      if (collapsed[it.tag]) delete collapsed[it.tag]; else collapsed[it.tag] = true;
      render();
    };
    div.appendChild(twist);

    var st = statusOf(it);
    var lead = "  ".repeat(depth) + (depth === 0 ? "\u25CF  " : "\u2514\u2500 ");
    var label = document.createElement("span");
    label.textContent = lead + it.tag + nodeRating(it);   /* v2.10: rating + demand (transformers keep kVA) */
    div.appendChild(label);

    if (st.indexOf("\u26A0") === 0) {
      var f = document.createElement("span");
      f.className = "flag";
      f.textContent = "   " + st;
      div.appendChild(f);
    }

    div.onclick = function () { select(it.tag); };
    host.appendChild(div);

    if (!collapsed[it.tag] && depth < MAXDEPTH) {
      kids.forEach(function (k) { emit(k, depth + 1); });
    }
  }

  items.filter(isRoot).forEach(function (r) { emit(r, 0); });
  document.getElementById("count").textContent = count + " item" + (count === 1 ? "" : "s");
}

/*  ---------- selection ---------- */
function clearProps() {
  document.getElementById("vTag").textContent = "(none selected)";
  ["vTV", "vKVA", "vDown", "vStat"].forEach(function (id) { document.getElementById(id).textContent = ""; });
  document.getElementById("selFT").value = "Sub-Feed";
  document.getElementById("txtCkt").value = "";
  document.getElementById("selMains").value = "";
  document.getElementById("txtSpaces").value = "";
  document.getElementById("txtRename").value = "";
  ["txtLen", "txtLenPri", "txtLenEm", "selMat", "selRace", "selExist"].forEach(function (id) { var e = document.getElementById(id); if (e) e.value = ""; });   /* v2.11 */
}

function select(tag) {
  selTag = tag;
  var it = byTag(tag);
  if (!it) { clearProps(); return; }
  document.getElementById("vTag").textContent = it.tag;
  document.getElementById("vTV").textContent = [it.type, it.volt].filter(Boolean).join("  /  ");
  document.getElementById("vKVA").textContent = it.kva ? it.kva.toFixed(1) : "\u2013";
  document.getElementById("vDown").textContent = downstreamKva(it.tag).toFixed(1);
  document.getElementById("vStat").textContent = statusOf(it);
  var ft = /feed-through/i.test(it.ft) ? "Feed-Through" : (/tap/i.test(it.ft) ? "Tap / Splice" : "Sub-Feed");
  document.getElementById("selFT").value = ft;
  document.getElementById("txtCkt").value = (it.ckt === "FT") ? "" : (it.ckt || "");
  document.getElementById("selMains").value = it.mains || "";
  /* v2.11: feeder inputs - the registry (Dashboard R:W) is their only home */
  document.getElementById("txtLen").value = (it.len === undefined || it.len === null) ? "" : it.len;
  document.getElementById("txtLenPri").value = (it.lenPri === undefined || it.lenPri === null) ? "" : it.lenPri;
  document.getElementById("txtLenEm").value = (it.lenEm === undefined || it.lenEm === null) ? "" : it.lenEm;
  document.getElementById("selMat").value = it.mat || "";
  document.getElementById("selRace").value = it.race || "";
  document.getElementById("selExist").value = it.exist || "";
  document.getElementById("txtSpaces").value = "";
  document.getElementById("txtRename").value = it.tag;
  fillFedList();
  document.getElementById("selFed").value = (it.fed === "(unassigned)") ? "" : (it.fed || "");
  render();
  /* v2.5: the space count is read live from the schedule's GEO_MAP */
  if (!it.head && !it.virtual && it.sheet && typeof SAZ !== "undefined") {
    Excel.run(function (ctx) {
      return SAZ.geo(ctx, it.sheet).then(function (g) { return SAZ.currentSpaces(g); });
    }).then(function (n) { if (selTag === tag) { document.getElementById("txtSpaces").value = n || ""; } })
      .catch(function () { /* not a schedule sheet */ });
  }
}

/*  parents offered exclude self and anything downstream of self - no circular feeds */
function fillFedList() {
  var sel = document.getElementById("selFed");
  var keep = sel.value;
  var bad = {};
  if (selTag) {
    bad[selTag] = true;
    descendants(selTag).forEach(function (d) { bad[d.tag] = true; });
  }
  sel.innerHTML = "";
  var blank = document.createElement("option");
  blank.value = ""; blank.textContent = "\u2014 utility / source";
  sel.appendChild(blank);
  items.forEach(function (it) {
    if (bad[it.tag] || it.virtual) return;
    var o = document.createElement("option");
    o.value = it.tag; o.textContent = it.tag;
    sel.appendChild(o);
  });
  sel.value = keep;
}

/*  ---------- writes ---------- */
/*  The Dashboard is a PROTECTED sheet. Every registry write runs inside an
    unprotect/reprotect pair (house options, restored on failure too) and
    resolves the registry from Equip_Registry rather than a fixed address.  */
var PROTECT_OPTS = {
  allowAutoFilter: false, allowDeleteColumns: false, allowDeleteRows: false,
  allowEditObjects: false, allowEditScenarios: false, allowFormatCells: false,
  allowFormatColumns: false, allowFormatRows: false, allowInsertColumns: false,
  allowInsertHyperlinks: false, allowInsertRows: false, allowPivotTables: false,
  allowSort: false, selectionMode: "Normal"
};
function withUnprotected(ctx, sheet, fn) {
  sheet.protection.load("protected");
  return ctx.sync().then(function () {
    var was = sheet.protection.protected;
    var re = function () {
      if (!was) return Promise.resolve();
      sheet.protection.protect(PROTECT_OPTS);
      return ctx.sync();
    };
    if (was) sheet.protection.unprotect();
    return ctx.sync().then(fn).then(function (out) {
      return re().then(function () { return out; });
    }, function (err) {
      return re().then(function () { throw err; }, function () { throw err; });
    });
  });
}
function writeCells(rowOffset, pairs) {
  return Excel.run(function (ctx) {
    return regRange(ctx).then(function (base) {
      base.load("worksheet/name");
      return ctx.sync().then(function () {
        var ws = ctx.workbook.worksheets.getItem(base.worksheet.name);
        return withUnprotected(ctx, ws, function () {
          pairs.forEach(function (p) {
            base.getCell(rowOffset, p.col).values = [[p.val]];
          });
          return ctx.sync();
        });
      });
    });
  });
}

/*  v2.5: Apply runs the full link through the shared engine (commands.js):
    circuit on the parent, poles from the child's voltage, live demand link,
    feed-through rows, mains, space count and Fed From - one source of truth. */
function busy(on) {
  ["btnApply", "btnUnlink", "btnRepair", "btnRename", "btnAdd"].forEach(function (id) {
    var b = document.getElementById(id); if (b) { b.disabled = !!on; }
  });
}
function engineReady() {
  if (typeof SAZ === "undefined" || !SAZ.mapApply) {
    setStatus("The shared engine (commands.js) did not load - publish commands.js next to taskpane.html and reload.");
    return false;
  }
  return true;
}
function applyLink() {
  var it = byTag(selTag);
  if (!it) { setStatus("Select a piece of equipment first."); return; }
  if (it.head) { setStatus(it.tag + " is the utility head end - edit it on the Dashboard."); return; }
  if (it.virtual) { setStatus(it.tag + " is derived automatically - link its secondary board instead."); return; }
  if (!engineReady()) { return; }
  var req = {
    tag: it.tag,
    fed: document.getElementById("selFed").value,
    ft: document.getElementById("selFT").value,
    ckt: document.getElementById("txtCkt").value,
    mains: document.getElementById("selMains").value,
    spaces: document.getElementById("txtSpaces").value,
    len: document.getElementById("txtLen").value,          /* v2.11 feeder inputs */
    lenPri: document.getElementById("txtLenPri").value,
    lenEm: document.getElementById("txtLenEm").value,
    mat: document.getElementById("selMat").value,
    race: document.getElementById("selRace").value,
    exist: document.getElementById("selExist").value
  };
  if (req.fed === it.tag) { setStatus("Equipment cannot feed itself."); return; }
  busy(true); setStatus("Applying...");
  SAZ.exec(function (ctx) { return SAZ.mapApply(ctx, req); })
    .then(function (msg) { busy(false); return refresh().then(function () { setStatus(msg || "Applied."); }); })
    .catch(function (e) { busy(false); setStatus("Apply failed: " + (e && e.message ? e.message : e)); });
}

function unlink() {
  var it = byTag(selTag);
  if (!it) { setStatus("Select a piece of equipment first."); return; }
  if (it.head) { setStatus(it.tag + " is the utility head end - it has no parent."); return; }
  if (it.virtual) { setStatus(it.tag + " is derived automatically - unlink its secondary board instead."); return; }
  if (!engineReady()) { return; }
  busy(true);
  SAZ.exec(function (ctx) { return SAZ.mapUnlink(ctx, it.tag); })
    .then(function (msg) { busy(false); return refresh().then(function () { setStatus(msg); }); })
    .catch(function (e) { busy(false); setStatus("Unlink failed: " + (e && e.message ? e.message : e)); });
}

function repairSelected() {
  var it = byTag(selTag);
  if (!it || it.head || it.virtual || !it.sheet) { setStatus("Select a schedule (panelboard, switchboard or MCC) first."); return; }
  if (!engineReady()) { return; }
  busy(true); setStatus("Checking " + it.tag + "...");
  SAZ.exec(function (ctx) { return SAZ.mapRepair(ctx, it.sheet); })
    .then(function (msg) { busy(false); setStatus(msg); })
    .catch(function (e) { busy(false); setStatus("Repair failed: " + (e && e.message ? e.message : e)); });
}

function gotoSheet() {
  var it = byTag(selTag);
  if (!it) { setStatus("Select a piece of equipment first."); return; }
  if (!it.sheet) { setStatus(it.tag + " has no schedule sheet in the registry."); return; }
  Excel.run(function (ctx) {
    var ws = ctx.workbook.worksheets.getItemOrNullObject(it.sheet);
    ws.load("name");
    return ctx.sync().then(function () {
      if (ws.isNullObject) { setStatus("Sheet not found: " + it.sheet); return; }
      ws.activate();
      return ctx.sync();
    });
  }).catch(function (e) { setStatus("Could not activate sheet: " + e.message); });
}

/*  registry row style (same as commands.js SAZ.formatRegistryRow) - applied to every row we write */
var REG_ALIGN = ["Left","Left","Center","Left","Center","Center","Left","Left","Center","Center"];
function formatRegistryRow(row) {
  var f = row.format;
  f.font.name = "Segoe UI"; f.font.size = 10; f.font.color = "#FF0000";
  f.font.bold = false; f.font.italic = false;
  f.fill.clear(); f.verticalAlignment = "Center";
  /* borders + grey/blue input cells (Feed Type, Feeder Ckt) come from the
     registry's conditional formats - they appear only once a Tag is entered */
  REG_ALIGN.forEach(function (a, j) { row.getCell(0, j).format.horizontalAlignment = a; });
  row.getCell(0, 5).numberFormat = [["#,##0.0;(#,##0.0);-"]];
  f.protection.locked = false;
}

/*  adds a placeholder row fed from the current selection - rename it on the Dashboard */
/*  v2.5: Add equipment deploys a REAL schedule (same create dialog as the ribbon:
    name, voltage, spaces) and, when a board is selected, links it straight in
    as a Sub-Feed at the first open circuit. No placeholder registry rows. */
function addEquipment() {
  if (!engineReady()) { return; }
  var parent = byTag(selTag);
  var fedTag = (parent && !parent.virtual) ? parent.tag : "";
  var before = {};
  items.forEach(function (x) { before[x.tag.toLowerCase()] = true; });
  busy(true);
  SAZ.exec(function (ctx) {
    return SAZ.form("Add equipment", fedTag ? "New schedule fed from " + fedTag + "." : "New schedule at the top of the map.", [
      { id: "kind", label: "Type", type: "select", options: ["Panelboard", "Switchboard", "MCC"], value: "Panelboard", required: true }
    ], "", 35).then(function (a) {
      if (!a || !a.ok || !a.values) { return null; }
      var k = a.values.kind;
      var tpl = k === "Switchboard" ? "TPL_Switchboard" : (k === "MCC" ? "TPL_MCC" : "TPL_Panelboard");
      return SAZ.uiCreate(ctx, tpl, k).then(function () {
        return SAZ.registryRecords(ctx);
      }).then(function (reg) {
        var added = reg.filter(function (r) { return !before[r.tag.toLowerCase()]; })[0];
        if (!added) { return null; }
        if (!fedTag) { return "Added " + added.tag + "."; }
        return SAZ.mapApply(ctx, { tag: added.tag, fed: fedTag, ft: "Sub-Feed" }).then(function (m) {
          selTag = added.tag; return "Added " + added.tag + ". " + m;
        });
      });
    });
  }).then(function (msg) { busy(false); return refresh().then(function () { if (msg) { setStatus(msg); } }); })
    .catch(function (e) { busy(false); setStatus("Add failed: " + (e && e.message ? e.message : e)); });
}

/*  Rename the selected board everywhere (tab, schedule name, registry, every
    text reference). Uses SAZ_RENAME below - identical copy of the commands.js core. */
function renameSelected() {
  var it = byTag(selTag);
  if (!it) { setStatus("Select a piece of equipment first."); return; }
  if (it.head) { setStatus(it.tag + " is a utility source - rename it on the Dashboard."); return; }
  if (it.virtual) { setStatus(it.tag + " is derived automatically - rename its secondary board instead."); return; }
  var nv = document.getElementById("txtRename").value;
  if (!String(nv || "").trim()) { setStatus("Type the new tag in Rename to."); return; }
  Excel.run(function (ctx) { return SAZ_RENAME.run(ctx, it.tag, nv); })
    .then(function (res) {
      setStatus("Renamed " + res.from + " to " + res.to + " - " + res.cells + " reference(s) updated.");
      selTag = res.to;
      return refresh();
    })
    .catch(function (e) { setStatus("Rename not done: " + (e.message || e)); });
}

/* ==========================================================
   RENAME EQUIPMENT  (v2.4)  -  IDENTICAL COPY in commands.js and taskpane.js.
   Edit both together.
   Renames a board everywhere its tag is stored as TEXT. Formula links (registry
   Voltage / Demand kVA, parent circuit kVA, feed-through links) follow the tab
   automatically when it is renamed; the text copies do not, so they are
   rewritten here:
     - the sheet tab and the schedule's own name cell (GEO NAME_CELL, locked)
     - the equipment registry: Tag, Fed From and Sheet columns
     - every other schedule: circuit descriptions, feed-through descriptions and
       the Fed From cell, including the "T-TAG" and "P via T-TAG" forms
   Registry row ORDER is unchanged, so study rows keep their typed inputs.
   Resolves to { from, to, sheetRenamed, cells }.
   ========================================================== */
var SAZ_RENAME = (function () {
  var OPTS = {
    allowAutoFilter: false, allowDeleteColumns: false, allowDeleteRows: false,
    allowEditObjects: false, allowEditScenarios: false, allowFormatCells: false,
    allowFormatColumns: false, allowFormatRows: false, allowInsertColumns: false,
    allowInsertHyperlinks: false, allowInsertRows: false, allowPivotTables: false,
    allowSort: false, selectionMode: "Normal"
  };
  var REG_NAME = "Equip_Registry", REG_SHEET = "Dashboard", REG_RANGE = "A101:J350";
  var C_TAG = 0, C_FED = 3, C_SHEET = 6;

  function esc(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
  function tokenRe(t) { return new RegExp("(^|[\\s(,]|T-)" + esc(t) + "(?=$|[\\s),])", "gi"); }
  function swap(text, oldT, newT) {
    var s = String(text == null ? "" : text);
    if (!s || !tokenRe(oldT).test(s)) { return null; }
    return s.replace(tokenRe(oldT), function (m, pre) { return pre + newT; });
  }
  function clean(raw) {
    return String(raw || "").replace(/[:\\\/\?\*\[\]]/g, "").trim().substring(0, 31);
  }
  function check(oldT, newT, names) {
    if (!newT) { return "The new tag contains no usable characters."; }
    if (/^'|'$/.test(newT)) { return "A tag cannot start or end with an apostrophe."; }
    if (newT.toLowerCase() === "history") { return "\"History\" is reserved by Excel."; }
    if (newT === oldT) { return "The new tag is the same as the current one."; }
    var clash = names.filter(function (n) {
      return n.toLowerCase() === newT.toLowerCase() && n.toLowerCase() !== oldT.toLowerCase();
    })[0];
    if (clash) { return "\"" + clash + "\" is already in use. Pick another tag."; }
    return "";
  }
  function letters(a) { var m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(String(a || "")); return m ? [m[1].toUpperCase(), Number(m[2])] : null; }
  function usable(v) { var s = String(v == null ? "" : v).trim(); return (s && s !== "-") ? s : ""; }

  function withUnprot(ctx, sheet, fn) {
    sheet.protection.load("protected");
    return ctx.sync().then(function () {
      var was = sheet.protection.protected;
      var re = function () { if (!was) { return Promise.resolve(); } sheet.protection.protect(OPTS); return ctx.sync(); };
      if (was) { sheet.protection.unprotect(); }
      return ctx.sync().then(fn).then(function (o) { return re().then(function () { return o; }); },
        function (e) { return re().then(function () { throw e; }, function () { throw e; }); });
    });
  }

  function registry(ctx) {
    var n = ctx.workbook.names.getItemOrNullObject(REG_NAME);
    n.load("isNullObject");
    return ctx.sync().then(function () {
      var sh = ctx.workbook.worksheets.getItemOrNullObject(REG_SHEET);
      sh.load("isNullObject");
      return ctx.sync().then(function () {
        if (!n.isNullObject) { return n.getRange(); }
        return sh.isNullObject ? null : sh.getRange(REG_RANGE);
      });
    });
  }

  function run(ctx, oldRaw, newRaw) {
    var oldT = String(oldRaw || "").trim(), newT = clean(newRaw);
    var sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    return ctx.sync().then(function () {
      var names = sheets.items.map(function (s) { return s.name; });
      return registry(ctx).then(function (reg) {
        if (reg) { reg.load("values,rowIndex,columnIndex,worksheet/name"); }
        return ctx.sync().then(function () {
          var regTags = reg ? reg.values.map(function (r) { return String(r[C_TAG] || "").trim(); }).filter(Boolean) : [];
          var target = sheets.items.filter(function (s) { return s.name.toLowerCase() === oldT.toLowerCase(); })[0];
          var inReg = regTags.some(function (t) { return t.toLowerCase() === oldT.toLowerCase(); });
          if (!target && !inReg) { throw new Error("\"" + oldT + "\" is not a schedule sheet or a registry tag."); }
          if (target) { oldT = target.name; }
          var err = check(oldT, newT, names.concat(regTags));
          if (err) { throw new Error(err); }

          /* sheet-scoped GEO_MAP on every sheet = the set of schedules */
          var geoNames = sheets.items.map(function (s) {
            var nm = s.names.getItemOrNullObject("GEO_MAP"); nm.load("isNullObject"); return nm;
          });
          return ctx.sync().then(function () {
            var geoRanges = geoNames.map(function (nm) {
              if (nm.isNullObject) { return null; }
              var r = nm.getRange(); r.load("values"); return r;
            });
            return ctx.sync().then(function () {
              var boards = [];
              geoRanges.forEach(function (r, i) {
                if (!r) { return; }
                var g = {};
                r.values.forEach(function (row) { var k = String(row[0] || "").trim(); if (k && k !== "GEO_MAP") { g[k] = row[1]; } });
                boards.push({ sheet: sheets.items[i], name: sheets.items[i].name, g: g });
              });

              /* 1. rename the tab first - formula links follow it */
              var renamed = false;
              if (target) { target.name = newT; renamed = true; }
              return ctx.sync().then(function () {
                /* 2. collect text cells on every schedule */
                var scans = [];
                boards.forEach(function (b) {
                  var g = b.g, spans = [];
                  var dc = usable(g.CKT_DESC_COL), f1 = Number(g.CKT_FIRST) || 0, l1 = Number(g.CKT_LAST) || 0;
                  if (dc && f1 && l1 >= f1) { spans.push(dc + f1 + ":" + dc + l1); }
                  var fd = usable(g.FEEDTHRU_DESC_COL), t1 = Number(g.FEEDTHRU_FIRST) || 0, t2 = Number(g.FEEDTHRU_LAST) || 0;
                  if (fd && t1 && t2 >= t1) { spans.push(fd + t1 + ":" + fd + (t2 + 1)); }
                  var nc = letters(usable(g.NAME_CELL));
                  if (nc) { spans.push(nc[0] + (nc[1] + 1)); }          /* Fed From = cell below the name */
                  var rs = spans.map(function (a) { var x = b.sheet.getRange(a); x.load("formulas,address"); return x; });
                  scans.push({ b: b, rs: rs, nc: nc });
                });
                return ctx.sync().then(function () {
                  var cells = 0, jobs = [];
                  scans.forEach(function (sc) {
                    var edits = [];
                    sc.rs.forEach(function (x) {
                      x.formulas.forEach(function (row, i) {
                        var v = row[0];
                        if (typeof v !== "string" || v.charAt(0) === "=") { return; }
                        var nv = swap(v, oldT, newT);
                        if (nv !== null) { edits.push({ r: x, i: i, v: nv }); }
                      });
                    });
                    var isTarget = !!target && sc.b.name.toLowerCase() === oldT.toLowerCase();
                    if (isTarget && sc.nc) { edits.push({ addr: sc.nc[0] + sc.nc[1], v: newT }); }
                    if (!edits.length) { return; }
                    cells += edits.length;
                    jobs.push(function () {
                      return withUnprot(ctx, sc.b.sheet, function () {
                        edits.forEach(function (e) {
                          var cell = e.addr ? sc.b.sheet.getRange(e.addr) : e.r.getCell(e.i, 0);
                          cell.values = [[e.v]];
                        });
                        return ctx.sync();
                      });
                    });
                  });
                  /* 3. registry text columns */
                  if (reg) {
                    var redits = [];
                    reg.values.forEach(function (row, i) {
                      [C_TAG, C_FED, C_SHEET].forEach(function (c) {
                        var v = String(row[c] || "").trim();
                        if (v && v.toLowerCase() === oldT.toLowerCase()) { redits.push([i, c]); }
                      });
                    });
                    if (redits.length) {
                      cells += redits.length;
                      var rsheet = ctx.workbook.worksheets.getItem(reg.worksheet.name);
                      jobs.push(function () {
                        return withUnprot(ctx, rsheet, function () {
                          redits.forEach(function (e) { reg.getCell(e[0], e[1]).values = [[newT]]; });
                          return ctx.sync();
                        });
                      });
                    }
                  }
                  return jobs.reduce(function (p, j) { return p.then(j); }, Promise.resolve()).then(function () {
                    return { from: oldT, to: newT, sheetRenamed: renamed, cells: cells };
                  });
                });
              });
            });
          });
        });
      });
    });
  }
  return { run: run, clean: clean };
})();
