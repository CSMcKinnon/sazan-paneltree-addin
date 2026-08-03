/*  SAZAN Panel Tree - Excel task pane
    Reads and writes the equipment registry on  Dashboard!A75:J114
    Columns (0-based):  0 Tag  1 Type  2 Voltage  3 Fed From  5 Demand kVA
                        6 Sheet  8 Feed Type  9 Ckt / space                */

var REG_SHEET = "Dashboard";
var REG_RANGE = "A75:J114";
var C_TAG = 0, C_TYPE = 1, C_VOLT = 2, C_FED = 3, C_KVA = 5, C_SHEET = 6, C_FT = 8, C_CKT = 9;
var MAXDEPTH = 12;

var items = [];          // { tag, type, volt, fed, kva, sheet, ft, ckt, row }
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
    var rg = ctx.workbook.worksheets.getItem(REG_SHEET).getRange(REG_RANGE);
    rg.load("values");
    return ctx.sync().then(function () {
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
          row: i                                   // offset inside REG_RANGE
        });
      });
      render();
      fillFedList();
      if (selTag) select(selTag); else clearProps();
      setStatus("");
    });
  }).catch(function (e) { setStatus("Read failed: " + e.message); });
}

/*  ---------- tree helpers ---------- */
function byTag(tag) {
  for (var i = 0; i < items.length; i++) if (items[i].tag === tag) return items[i];
  return null;
}
function childrenOf(tag) {
  return items.filter(function (it) { return it.fed === tag; });
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
  if (!it.fed) return "\u25C6 Source";
  if (!byTag(it.fed)) return "\u26A0 Parent not in registry";
  var chain = [], p = it, guard = 0;
  while (p && p.fed && guard++ < MAXDEPTH) {
    if (p.fed === it.tag) return "\u26A0 Circular feed";
    chain.push(p.fed); p = byTag(p.fed);
  }
  if (!it.ft) return "\u26A0 Feed type not set";
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
    label.textContent = lead + it.tag + (it.kva ? "   " + it.kva.toFixed(1) + " kVA" : "");
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
  document.getElementById("selFT").value = "";
  document.getElementById("txtCkt").value = "";
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
  document.getElementById("selFT").value = it.ft || "";
  document.getElementById("txtCkt").value = it.ckt || "";
  fillFedList();
  document.getElementById("selFed").value = it.fed || "";
  render();
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
    if (bad[it.tag]) return;
    var o = document.createElement("option");
    o.value = it.tag; o.textContent = it.tag;
    sel.appendChild(o);
  });
  sel.value = keep;
}

/*  ---------- writes ---------- */
function writeCells(rowOffset, pairs) {
  return Excel.run(function (ctx) {
    var ws = ctx.workbook.worksheets.getItem(REG_SHEET);
    var base = ws.getRange(REG_RANGE);
    pairs.forEach(function (p) {
      base.getCell(rowOffset, p.col).values = [[p.val]];
    });
    return ctx.sync();
  });
}

function applyLink() {
  var it = byTag(selTag);
  if (!it) { setStatus("Select a piece of equipment first."); return; }
  var fed = document.getElementById("selFed").value;
  var ft = document.getElementById("selFT").value;
  var ckt = document.getElementById("txtCkt").value;
  if (fed === it.tag) { setStatus("Equipment cannot feed itself."); return; }
  writeCells(it.row, [{ col: C_FED, val: fed }, { col: C_FT, val: ft }, { col: C_CKT, val: ckt }])
    .then(function () { setStatus("Linked " + it.tag + (fed ? " to " + fed : " as a source.")); return refresh(); })
    .catch(function (e) { setStatus("Write failed: " + e.message); });
}

function unlink() {
  var it = byTag(selTag);
  if (!it) { setStatus("Select a piece of equipment first."); return; }
  writeCells(it.row, [{ col: C_FED, val: "" }, { col: C_FT, val: "" }, { col: C_CKT, val: "" }])
    .then(function () { setStatus(it.tag + " is now an unlinked source."); return refresh(); })
    .catch(function (e) { setStatus("Write failed: " + e.message); });
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

/*  adds a placeholder row fed from the current selection - rename it on the Dashboard */
function addEquipment() {
  return Excel.run(function (ctx) {
    var base = ctx.workbook.worksheets.getItem(REG_SHEET).getRange(REG_RANGE);
    base.load("values");
    return ctx.sync().then(function () {
      var free = -1;
      for (var i = 0; i < base.values.length; i++) {
        if (String(base.values[i][C_TAG] || "").trim() === "") { free = i; break; }
      }
      if (free < 0) { setStatus("Registry is full - extend Dashboard!A75:J114."); return; }
      var n = 1, tag;
      do { tag = "NEW-" + n++; } while (byTag(tag));
      base.getCell(free, C_TAG).values = [[tag]];
      if (selTag) base.getCell(free, C_FED).values = [[selTag]];
      return ctx.sync().then(function () {
        setStatus("Added " + tag + (selTag ? " fed from " + selTag : "") + " - rename it on the Dashboard.");
        selTag = tag;
        return refresh();
      });
    });
  }).catch(function (e) { setStatus("Add failed: " + e.message); });
}
