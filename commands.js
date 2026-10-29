/*  SAZAN PowerCalc  -  commands.js
    Ribbon command handlers for the web add-in.
    Loaded by commands.html, which the manifest declares as <FunctionFile>.
    Every FunctionName in manifest.xml must have a matching Office.actions.associate here.
*/

var CFG = {
  host:        "https://csmckinnon.github.io/sazan-paneltree-addin",
  /* templateFile retired 2026-08-31: one hosted artifact, two roles. The seed
     carries TPL_* already, so ensureTemplate now pulls them from projectFile.
     Two files holding the same 2,787 formulas would inevitably drift. */
  projectFile: "sazan-project-template.xlsx",   /* Artifact A - the full 19-sheet seed */
  regSheet:    "Dashboard",   regName: "Equip_Registry",
  regRange:    "A101:J350"   /* fallback only - resolved from Equip_Registry at runtime */
};

var SAZ = {};
SAZ.DIALOG_VER = "2.12.0";   /* sent as ?_v= on every dialog URL so Office / GitHub Pages cannot serve a cached dialog.html - bump when dialog.html changes */
/* ---------- dialog helpers: replace VBA MsgBox / InputBox ---------- */

/* v2.5: Office refuses to open a dialog while the previous one is still closing
   (error 12007) - that is why the second box of any two-step command never
   appeared. open() retries for ~2 s, and each answer resolves only after the
   close has had time to finish, so callers can chain dialogs safely. */
SAZ.dialog = function (params) {
  return new Promise(function (resolve) {
    var qs = Object.keys(params).map(function (k) {
      return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]);
    }).join("&");
    var url = CFG.host + "/dialog.html?" + qs + "&_v=" + encodeURIComponent(SAZ.DIALOG_VER);
    var done = false;
    var finish = function (out) { if (done) { return; } done = true; setTimeout(function () { resolve(out); }, 350); };
    var open = function (tries) {
      Office.context.ui.displayDialogAsync(url,
        { height: params._h || 42, width: params._w || 30, displayInIframe: true },
        function (res) {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            if (res.error && res.error.code === 12007 && tries < 8) {
              setTimeout(function () { open(tries + 1); }, 250);
              return;
            }
            finish({ ok: false }); return;
          }
          var dlg = res.value;
          dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            var out = { ok: false };
            try { out = JSON.parse(arg.message); } catch (e) { }
            dlg.close();
            finish(out);
          });
          dlg.addEventHandler(Office.EventType.DialogEventReceived, function () {
            finish({ ok: false });
          });
        });
    };
    open(0);
  });
};

SAZ.notify = function (title, text) {
  return SAZ.dialog({ mode: "notify", title: title, text: text });
};

SAZ.prompt = function (title, label, dflt, required) {
  return SAZ.dialog({ mode: "prompt", title: title, label: label,
                      "default": dflt || "", required: required ? "1" : "0", text: "" });
};

/* v2.3 multi-field form: fields = [{ id, label, type: text|number|select, options,
   value, required, min, max, hint, list }]. Resolves { ok, values: { id: value } }.
   err is shown in amber - callers re-open the form with the reason after a failed check. */
SAZ.form = function (title, text, fields, err, heightPct) {
  return SAZ.dialog({ mode: "form", title: title, text: text || "",
                      form: JSON.stringify({ fields: fields, err: err || "" }),
                      _h: heightPct || 60 }).then(function (ans) {
    /* Loop guard: a pre-2.3 dialog.html (old copy on GitHub, or a cached one) has no
       form mode - it shows the title with OK / Cancel only and returns { ok:true } with
       no values. Every caller then re-opened the form with an error the old page could
       not display, so OK just brought the same box back forever. Stop and say why. */
    if (ans && ans.ok && !ans.values) {
      return SAZ.notify("Dialog page out of date",
        "The dialog page the add-in loaded is an older version without form fields, so this command cannot continue.\n\n" +
        "Publish dialog.html (v" + SAZ.DIALOG_VER + ") together with commands.js, clear the Office add-in cache, then reopen the workbook."
      ).then(function () { return { ok: false, stale: true }; });
    }
    return ans;
  });
};

/* ---------- run wrapper: guarantees event.completed() ---------- */

/* Seed lock: the calculation tabs below ship protected so a deployed project
   cannot be edited outside its input cells (grey fill / blue text). Every
   ribbon command runs through SAZ.run, which unprotects these tabs for the
   duration of the command and re-protects ALL of them afterwards, success or
   failure. Sheets that do not exist yet are skipped. Panel, study and Dashboard
   sheets keep their own withUnprotected handling. */
SAZ.SEED_LOCKED = ["Instructions", "Load Calc", "Logging Demand", "Generator",
  "UPS", "Report", "NEC_Data", "Energy_Code_Data"];
SAZ.seedSheets = function (ctx) {
  var out = SAZ.SEED_LOCKED.map(function (n) {
    var w = ctx.workbook.worksheets.getItemOrNullObject(n);
    w.load("isNullObject");
    return w;
  });
  return ctx.sync().then(function () {
    var live = out.filter(function (w) { return !w.isNullObject; });
    live.forEach(function (w) { w.protection.load("protected"); });
    return ctx.sync().then(function () { return live; });
  });
};
SAZ.seedUnlock = function (ctx) {
  return SAZ.seedSheets(ctx).then(function (live) {
    live.forEach(function (w) { if (w.protection.protected) { w.protection.unprotect(); } });
    return ctx.sync();
  });
};
SAZ.seedRelock = function (ctx) {
  return SAZ.seedSheets(ctx).then(function (live) {
    live.forEach(function (w) { if (!w.protection.protected) { w.protection.protect(SAZ.PROTECT_OPTS); } });
    return ctx.sync();
  });
};

SAZ.run = function (name, body) {
  return function (event) {
    Excel.run(function (ctx) {
      return SAZ.seedUnlock(ctx).then(function () { return body(ctx); }).then(
        function (out) { return SAZ.seedRelock(ctx).then(function () { return out; }); },
        function (err) { return SAZ.seedRelock(ctx).then(function () { throw err; }, function () { throw err; }); });
    })
      .then(function () { event.completed(); })
      .catch(function (err) {
        var msg = (err && err.message) ? err.message : String(err);
        SAZ.notify(name + " failed", msg).then(function () { event.completed(); },
                                                function () { event.completed(); });
      });
  };
};

/* ---------- sheet name helpers: this is what makes multiples work ---------- */

SAZ.sanitise = function (raw) {
  var s = String(raw || "").replace(/[:\\\/\?\*\[\]]/g, "").trim();
  return s.substring(0, 31);
};

SAZ.uniqueSheetName = function (ctx, base) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  return ctx.sync().then(function () {
    var taken = {};
    sheets.items.forEach(function (s) { taken[s.name.toLowerCase()] = true; });
    if (!taken[base.toLowerCase()]) { return base; }
    for (var i = 2; i < 200; i++) {
      var cand = (base + " (" + i + ")").substring(0, 31);
      if (!taken[cand.toLowerCase()]) { return cand; }
    }
    throw new Error("Could not find a free sheet name based on " + base);
  });
};

/* ---------- template cache: fetch once, then copy locally forever ---------- */

SAZ.toBase64 = function (buf) {
  var bytes = new Uint8Array(buf), chunk = 0x8000, out = [];
  for (var i = 0; i < bytes.length; i += chunk) {
    out.push(String.fromCharCode.apply(null, bytes.subarray(i, i + chunk)));
  }
  return btoa(out.join(""));
};

SAZ.ensureTemplate = function (ctx, tplName) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  return ctx.sync().then(function () {
    var hit = sheets.items.filter(function (s) {
      return s.name.toLowerCase() === tplName.toLowerCase();
    })[0];
    if (hit) { return hit.name; }
    return fetch(CFG.host + "/" + CFG.projectFile).then(function (res) {
      if (!res.ok) {
        throw new Error("Project seed not reachable (" + res.status + "). " +
          "Add " + tplName + " to this workbook, or check your connection and try again.");
      }
      return res.arrayBuffer();
    }).then(function (buf) {
      var b64 = SAZ.toBase64(buf);
      ctx.application.load("calculationMode");
      return ctx.sync().then(function () {
        var saved = ctx.application.calculationMode;
        ctx.application.calculationMode = Excel.CalculationMode.manual;
        return ctx.sync()
          .then(function () {
            return ctx.workbook.insertWorksheetsFromBase64(b64, { sheetNamesToInsert: [tplName] });
          })
          .then(function () { return ctx.sync(); })
          .then(function () {
            var t = ctx.workbook.worksheets.getItem(tplName);
            t.visibility = Excel.SheetVisibility.veryHidden;  /* match arrangeScaffold */
            ctx.application.calculationMode = saved;
            return ctx.sync();
          })
          .then(function () { return tplName; });
      });
    });
  });
};

/* ---------- registry write: the missing link that fills the Distribution Map ---------- */

SAZ.registryRange = function (ctx) {
  var n = ctx.workbook.names.getItemOrNullObject(CFG.regName);
  n.load("isNullObject");
  return ctx.sync().then(function () {
    if (!n.isNullObject) { return n.getRange(); }
    return ctx.workbook.worksheets.getItem(CFG.regSheet).getRange(CFG.regRange);
  });
};

/* Registry row style - committed to every row the ribbon writes so a paste
   or a manual clear can never leave an unformatted row behind.
   Dashboard legend: red = auto-filled, no fill; grid via conditional format. */
SAZ.REG_ALIGN = ["Left","Left","Center","Left","Center","Center","Left","Left","Center","Center"];
SAZ.formatRegistryRow = function (row) {
  var f = row.format;
  f.font.name = "Segoe UI"; f.font.size = 10; f.font.color = "#FF0000";
  f.font.bold = false; f.font.italic = false;
  f.fill.clear(); f.verticalAlignment = "Center";
  /* borders + grey/blue input cells (Feed Type, Feeder Ckt) come from the
     registry's conditional formats - they appear only once a Tag is entered */
  SAZ.REG_ALIGN.forEach(function (a, j) {
    row.getCell(0, j).format.horizontalAlignment = a;
  });
  row.getCell(0, 5).numberFormat = [["#,##0.0;(#,##0.0);-"]];
  f.protection.locked = false;
};

SAZ.registerEquipment = function (ctx, rec) {
  /* Dashboard is a PROTECTED sheet and the registry cells were locked, so this
     write used to fail with AccessDenied AFTER the schedule had been cloned -
     the sheet appeared but never reached the registry or the Distribution Map.
     Fixed 2026-10-22: registry A:J unlocked on the Dashboard AND the write runs
     inside withUnprotected, so either fix alone is sufficient. */
  return SAZ.registryRange(ctx).then(function (rg) {
    rg.load("values,rowIndex,columnIndex,rowCount,worksheet/name");
    return ctx.sync().then(function () {
      var free = -1;
      for (var i = 0; i < rg.values.length; i++) {
        if (!rg.values[i][0]) { free = i; break; }
      }
      if (free < 0) {
        throw new Error("Equipment registry is full (" + rg.rowCount +
          " rows on " + rg.worksheet.name + "). Extend the Equip_Registry" +
          " defined name and Cfg_MaxNodes, then reload the add-in.");
      }
      var sheet = ctx.workbook.worksheets.getItem(rg.worksheet.name);
      return SAZ.withUnprotected(ctx, sheet, function () {
        var row = sheet.getRangeByIndexes(rg.rowIndex + free, rg.columnIndex, 1, 10);
        row.formulas = [[ rec.tag, rec.type, rec.volts || "", rec.fed || "", rec.mains || "",
                        rec.kva || "", rec.sheet || rec.tag, rec.notes || "",
                        rec.feedType || "", rec.ckt || "" ]];
        SAZ.formatRegistryRow(row);
        return ctx.sync();
      });
    });
  });
};
/* ---------- shared workers ---------- */

SAZ.goOrCreate = function (ctx, sheetName) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name,items/visibility");
  return ctx.sync().then(function () {
    var hit = sheets.items.filter(function (s) {
      return s.name.toLowerCase() === sheetName.toLowerCase();
    })[0];
    if (!hit) {
      return SAZ.notify("Sheet not found",
        "There is no \"" + sheetName + "\" sheet in this workbook.\n\n" +
        "Use the matching New command to create it first.");
    }
    if (hit.visibility !== Excel.SheetVisibility.visible) {
      hit.visibility = Excel.SheetVisibility.visible;
    }
    hit.activate();
    return ctx.sync();
  });
};

/* Copy a sheet whatever its visibility. Excel refuses to copy a veryHidden
   sheet (InvalidOperation - "not permitted for the current object") and a copy
   of a hidden sheet arrives hidden, which cannot be activated (InvalidReference -
   "reference is invalid"). New Project leaves TPL_* veryHidden and every module
   hidden, so every New command hit one or the other. Fix: lift the source to
   visible, copy, show the copy, then put the source back - on failure too.
   Resolves to the new sheet's name. */
SAZ.cloneSheet = function (ctx, srcName, newName) {
  var src = ctx.workbook.worksheets.getItem(srcName);
  src.load("visibility");
  return ctx.sync().then(function () {
    var was = src.visibility;
    var restore = function () {
      if (was !== Excel.SheetVisibility.visible) { src.visibility = was; }
      return ctx.sync();
    };
    if (was !== Excel.SheetVisibility.visible) {
      src.visibility = Excel.SheetVisibility.visible;
    }
    return ctx.sync().then(function () {
      var copy = src.copy(Excel.WorksheetPositionType.end);
      return ctx.sync().then(function () {
        copy.name = newName;
        copy.visibility = Excel.SheetVisibility.visible;
        return ctx.sync();
      });
    }).then(function () {
      return restore().then(function () { return newName; });
    }, function (err) {
      return restore().then(function () { throw err; }, function () { throw err; });
    });
  });
};

/* ---------- v2.3 deploy: one dialog commits name, voltage and size ----------
   New Panelboard / Switchboard / MCC asks for all three at once. On OK the
   schedule is cloned, the name is stamped into NAME_CELL and the voltage into
   VOLT_CELL (a plain input - it stays editable from the cell's own dropdown),
   the board is registered, and the circuit grid is trimmed or grown to the
   requested size through the same resizePanel used by Resize Panel.
   Voltage is pre-selected from the Dashboard project voltage (Cfg_Volt) when
   one is set; the list is Volt_Sys, the same list the schedule cell uses. */
SAZ.DEPLOY_DEFAULT_SPACES = { Panelboard: 42, Switchboard: 20, MCC: 20 };
SAZ.DEPLOY_MAX_SPACES = 200;

SAZ.voltChoices = function (ctx) {
  var nm = ctx.workbook.names.getItemOrNullObject("Volt_Sys");
  var pv = ctx.workbook.names.getItemOrNullObject("Cfg_Volt");
  nm.load("isNullObject"); pv.load("isNullObject");
  return ctx.sync().then(function () {
    var a = nm.isNullObject ? null : nm.getRange(), b = pv.isNullObject ? null : pv.getRange();
    if (a) { a.load("values"); }
    if (b) { b.load("values"); }
    return ctx.sync().then(function () {
      var list = a ? a.values.map(function (r) { return String(r[0] || "").trim(); }).filter(Boolean) : [];
      var proj = b ? String(b.values[0][0] || "").trim() : "";
      return { list: list, project: (list.indexOf(proj) >= 0) ? proj : "" };
    });
  });
};

/* Spaces a board currently has, read from its GEO_MAP. */
SAZ.currentSpaces = function (g) {
  if (String(g.RESIZE_MODE || "").trim().toUpperCase() === "TWOBLOCK") {
    return 2 * (Number(g.BLOCK_SIZE) || 0);
  }
  return (Number(g.CKT_LAST) || 0) - (Number(g.CKT_FIRST) || 0) + 1;
};

/* Write the deploy choices into the new schedule's header. */
SAZ.stampHeader = function (ctx, sheetName, volt) {
  return SAZ.geo(ctx, sheetName).then(function (g) {
    var sheet = ctx.workbook.worksheets.getItem(sheetName);
    var nc = SAZ.geoUsable(g.NAME_CELL), vc = SAZ.geoUsable(g.VOLT_CELL);
    return SAZ.withUnprotected(ctx, sheet, function () {
      if (nc) { sheet.getRange(nc).values = [[sheetName]]; }
      if (vc && volt) { sheet.getRange(vc).values = [[volt]]; }
      if (vc) { var vR = sheet.getRange(vc); vR.format.protection.locked = true; vR.format.fill.clear(); vR.format.font.color = SAZ.CALC_FONT; }   /* v2.12: voltage owned by the Distribution Map */
      return ctx.sync();
    });
  });
};

SAZ.uiCreate = function (ctx, tplName, typeLabel, err, last) {
  return SAZ.requireSheets(ctx, SAZ.CLONE_REQUIRES).then(function (ok) { if (!ok) { return null; }
  return SAZ.voltChoices(ctx).then(function (vc) {
    var dSp = SAZ.DEPLOY_DEFAULT_SPACES[typeLabel] || 42;
    last = last || {};
    var fields = [
      { id: "tag", label: typeLabel + " name / tag", type: "text", value: last.tag || "", required: true,
        hint: "Becomes the sheet tab, the name on the schedule and the Distribution Map tag." },
      { id: "volt", label: "System voltage", type: "select", options: vc.list,
        value: last.volt || vc.project, required: true,
        hint: vc.project ? "Pre-selected from the Dashboard project voltage. Change it here, or later in the Distribution Map (the schedule cell is locked)."
                         : "No project voltage on the Dashboard yet - pick one. It can be changed later in the Distribution Map." },
      { id: "spaces", label: (typeLabel === "Panelboard" ? "Spaces (circuits)" : (typeLabel === "MCC" ? "Units / buckets" : "Circuits / sections")),
        type: "number", value: String(last.spaces || dSp), min: 1, max: SAZ.DEPLOY_MAX_SPACES, required: true,
        hint: typeLabel === "Panelboard" ? "Odd counts round up to even (odd circuits left, even right)."
                                         : "Can be changed later with Resize Panel." }
    ];
    return SAZ.form("New " + typeLabel, "", fields, err || "", 72).then(function (ans) {
      if (!ans.ok) { return null; }
      var v = ans.values || {};
      var tag = SAZ.sanitise(v.tag);
      if (!tag) { return SAZ.uiCreate(ctx, tplName, typeLabel, "That tag contains no usable characters.", v); }
      if (vc.list.length && vc.list.indexOf(v.volt) < 0) {
        return SAZ.uiCreate(ctx, tplName, typeLabel, "Pick a system voltage from the list.", v);
      }
      var spaces = parseInt(v.spaces, 10) || dSp;
      if (typeLabel === "Panelboard" && spaces % 2) { spaces += 1; }
      return SAZ.ensureTemplate(ctx, tplName).then(function (tpl) {
        return SAZ.uniqueSheetName(ctx, tag).then(function (finalName) {
          return SAZ.cloneSheet(ctx, tpl, finalName);
        });
      }).then(function (finalName) {
        return SAZ.stampHeader(ctx, finalName, v.volt).then(function () {
          return SAZ.registerFromGeo(ctx, finalName, typeLabel);
        }).then(function () {
          return SAZ.geo(ctx, finalName);
        }).then(function (g) {
          if (SAZ.currentSpaces(g) === spaces) { return null; }
          return SAZ.resizePanel(ctx, finalName, spaces);
        }).then(function () {
          ctx.workbook.worksheets.getItem(finalName).activate();
          ctx.workbook.worksheets.getItem(finalName).getRange("A1").select();
          return ctx.sync();
        }).then(function () {
          if (finalName !== tag) {
            return SAZ.notify("Tag already in use",
              "\"" + tag + "\" already exists, so the new " + typeLabel +
              " was created as \"" + finalName + "\".");
          }
        });
      });
    });
  }); });
};

SAZ.createModule = function (ctx, sheetName) {
  /*  First use in a project: New Project put the module away hidden, so New
      simply reveals it - the scaffold notice promises "its tab will appear".
      Once it is showing, New makes a further copy (Generator (2), ...).
      Cross-sheet formulas inside a duplicated module still need review per
      module - see PORT_Audit. */
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name,items/visibility");
  return ctx.sync().then(function () {
    var hit = sheets.items.filter(function (s) {
      return s.name.toLowerCase() === sheetName.toLowerCase();
    })[0];
    if (!hit) {
      return SAZ.notify("Module not in this workbook",
        "\"" + sheetName + "\" does not exist here yet.\n\n" +
        "Run New Project from a blank workbook to bring in the full scaffold.");
    }
    if (hit.visibility !== Excel.SheetVisibility.visible) {
      hit.visibility = Excel.SheetVisibility.visible;
      return ctx.sync().then(function () {
        hit.activate();
        return ctx.sync();
      });
    }
    return SAZ.uniqueSheetName(ctx, hit.name).then(function (finalName) {
      return SAZ.cloneSheet(ctx, hit.name, finalName);
    }).then(function (finalName) {
      ctx.workbook.worksheets.getItem(finalName).activate();
      return ctx.sync();
    });
  });
};

/* ---------- v1.2 project modules: Generator / UPS ----------
   Neither module is in a project by default. New Generator / New UPS deploys
   it: the Dashboard status cell (Tog_Gen / Tog_UPS) reads ON, the Gen or UPS
   column appears in every Load Calc section, the matching connected-loads
   block appears at the foot of the Load Calc, and the module tab is shown.
   Pressed again it offers go-to / remove / add-another-copy. Removal is
   refused while any load is still tagged to the module. */
SAZ.MODULES = {
  Generator: { tog: "Tog_Gen", col: "R:R", banner: "GENERATOR-CONNECTED LOADS", total: "LC_GenTotal", byCat: "LC_GenByCategory", tag: "Gen" },
  UPS:       { tog: "Tog_UPS", col: "S:S", banner: "UPS-CONNECTED LOADS",       total: "LC_UPSTotal", byCat: "LC_UPSByCategory", tag: "UPS" }
};

/* Row bands of the two connected-loads blocks on the Load Calc, found by
   banner text: banner row through the row before the next banner. */
SAZ.moduleBands = function (ctx, sheet) {
  var used = sheet.getUsedRange(true);
  used.load("rowIndex,rowCount");
  var colA = sheet.getRange("A1:A400");
  colA.load("values");
  return ctx.sync().then(function () {
    var last = used.rowIndex + used.rowCount;
    var heads = [];
    colA.values.forEach(function (r, i) { if (String(r[0] || "").trim() !== "") { heads.push({ row: i + 1, text: String(r[0]) }); } });
    var out = {};
    Object.keys(SAZ.MODULES).forEach(function (k) {
      var n = heads.findIndex(function (h) { return h.text.indexOf(SAZ.MODULES[k].banner) === 0; });
      if (n < 0) { return; }
      out[k] = { first: heads[n].row, last: n + 1 < heads.length ? heads[n + 1].row - 1 : last };
    });
    return out;
  });
};

SAZ.moduleState = function (ctx) {
  var out = {}, rs = {};
  Object.keys(SAZ.MODULES).forEach(function (k) {
    var nm = ctx.workbook.names.getItemOrNullObject(SAZ.MODULES[k].tog);
    nm.load("isNullObject");
    rs[k] = nm;
  });
  return ctx.sync().then(function () {
    var cells = {};
    Object.keys(rs).forEach(function (k) { if (!rs[k].isNullObject) { cells[k] = rs[k].getRange(); cells[k].load("values"); } });
    return ctx.sync().then(function () {
      Object.keys(SAZ.MODULES).forEach(function (k) {
        out[k] = cells[k] ? String(cells[k].values[0][0]).toUpperCase() === "ON" : true;   /* no status cell = legacy workbook, treat as deployed */
      });
      return out;
    });
  });
};

SAZ.setModuleDeployed = function (ctx, key, on) {
  var m = SAZ.MODULES[key];
  var lc = ctx.workbook.worksheets.getItem(SAZ.EQUIP_REGISTRY_SHEET);
  var dash = ctx.workbook.worksheets.getItemOrNullObject("Dashboard");
  var mod = ctx.workbook.worksheets.getItemOrNullObject(key);
  dash.load("isNullObject"); mod.load("isNullObject,visibility");
  return ctx.sync().then(function () {
    /* status first - the roll-up helper reads it to decide what it may unhide */
    if (dash.isNullObject) { return null; }
    return SAZ.withUnprotected(ctx, dash, function () {
      ctx.workbook.names.getItem(m.tog).getRange().values = [[on ? "ON" : "OFF"]];
      return ctx.sync();
    });
  }).then(function () {
    return SAZ.moduleBands(ctx, lc);
  }).then(function (bands) {
    lc.getRange(m.col).columnHidden = !on;
    var b = bands[key];
    if (b) { lc.getRange(b.first + ":" + b.last).rowHidden = !on; }
    return ctx.sync().then(function () {
      if (!on || !b) { return null; }
      /* re-hide the roll-up lines of sections that are switched off */
      return SAZ.sectionRegistry(ctx).then(function (list) {
        var chain = Promise.resolve();
        list.forEach(function (e) {
          if (String(e.state).toUpperCase() === "NO") {
            chain = chain.then(function () { return SAZ.setRollupRows(ctx, lc, e.label, false); });
          }
        });
        return chain;
      });
    });
  }).then(function () {
    if (mod.isNullObject) { return null; }
    if (on) {
      mod.visibility = Excel.SheetVisibility.visible;
      mod.activate();
    } else {
      lc.activate();
      mod.visibility = Excel.SheetVisibility.hidden;
    }
    return ctx.sync();
  });
};

SAZ.uiModule = function (ctx, key) {
  var m = SAZ.MODULES[key];
  return SAZ.moduleState(ctx).then(function (st) {
    if (!st[key]) { return SAZ.setModuleDeployed(ctx, key, true); }
    return SAZ.prompt(key + " is in this project",
      "1 = go to " + key + "\n2 = remove " + key + " from the project\n3 = add another " + key + " sheet (copy)", "1", true)
      .then(function (ans) {
        var pick = parseInt(String(ans || "").trim(), 10);
        if (pick === 1) { return SAZ.goOrCreate(ctx, key); }
        if (pick === 3) { return SAZ.createModule(ctx, key); }
        if (pick !== 2) { return null; }
        var tot = ctx.workbook.names.getItem(m.total).getRange();
        var cat = ctx.workbook.names.getItem(m.byCat).getRange();
        tot.load("values"); cat.load("values,rowIndex,columnIndex");
        return ctx.sync().then(function () {
          if (Number(tot.values[0][0]) > 0) {
            var lc = ctx.workbook.worksheets.getItem(SAZ.EQUIP_REGISTRY_SHEET);
            var lab = lc.getRangeByIndexes(cat.rowIndex, cat.columnIndex - 1, cat.values.length, 1);
            lab.load("values");
            return ctx.sync().then(function () {
              var lines = [];
              cat.values.forEach(function (r, i) { if (Number(r[0]) > 0) { lines.push("  " + lab.values[i][0] + ": " + Number(r[0]).toFixed(1) + " kVA"); } });
              return SAZ.notify(key + " not removed",
                Number(tot.values[0][0]).toFixed(1) + " kVA is still tagged " + m.tag + " in the Load Calc:\n\n" + lines.join("\n") +
                "\n\nClear those " + m.tag + " tags first, then run this again.");
            });
          }
          return SAZ.setModuleDeployed(ctx, key, false);
        });
      });
  });
};


SAZ.todo = function (label, worker) {
  return SAZ.notify(label + " - not ported yet",
    "This command is still backed by VBA (" + worker + ").\n\n" +
    "See the PORT_Audit tab for its scope and status.");
};
/* ---------- refresh feeders ----------
   The Feeder Schedule is DERIVED: rows spill out of the equipment registry and
   every printed column is a formula. So a refresh must NOT rebuild rows the way
   the old VBA (Rbn_RefreshFeeders) did - writing rows would overwrite the engine.
   Instead we force a full recalculation and report what the sheet now says, so
   the user still gets the confirmation they are used to. */

SAZ.FEEDER_FIRST_ROW = 5;
SAZ.FEEDER_LAST_ROW  = 254;

SAZ.refreshFeeders = function (ctx) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name,items/visibility");
  return ctx.sync().then(function () {
    var hit = sheets.items.filter(function (s) {
      return s.name.toLowerCase() === "feeder schedule";
    })[0];
    if (!hit) {
      return SAZ.notify("Refresh Feeders",
        "There is no \"Feeder Schedule\" sheet in this workbook.\n\n" +
        "Use New Feeder Schedule to create it first.");
    }
    if (hit.visibility !== Excel.SheetVisibility.visible) {
      hit.visibility = Excel.SheetVisibility.visible;
    }
    hit.activate();
    ctx.workbook.application.calculate(Excel.CalculationType.fullRebuild);

    var first = SAZ.FEEDER_FIRST_ROW, last = SAZ.FEEDER_LAST_ROW;
    var tags   = hit.getRange("B" + first + ":B" + last);
    var exist  = hit.getRange("N" + first + ":N" + last);
    var lens   = hit.getRange("H" + first + ":H" + last);
    var status = hit.getRange("V" + first + ":V" + last);
    tags.load("values"); exist.load("values");
    lens.load("values"); status.load("values");

    return ctx.sync().then(function () {
      var n = 0, nExist = 0, nNoLen = 0, warn = [];
      for (var i = 0; i < tags.values.length; i++) {
        var tag = String(tags.values[i][0] || "");
        if (!tag) { continue; }
        n++;
        if (String(exist.values[i][0] || "").toUpperCase() === "Y") { nExist++; }
        if (String(lens.values[i][0] || "") === "") { nNoLen++; }
        var st = String(status.values[i][0] || "");
        if (st.charAt(0) === "\u26A0") {
          warn.push(tag + " - " + st.replace(/^\u26A0\s*/, ""));
        }
      }
      if (n === 0) {
        return SAZ.notify("Refresh Feeders",
          "Recalculated. No feeders yet.\n\n" +
          "Feeders derive from the equipment registry on the Dashboard: give a " +
          "device a Fed From and its feeder appears here automatically.");
      }
      var msg = "Recalculated. " + n + " feeder(s).";
      if (nExist) { msg += "\n" + nExist + " marked EXIST (sizing bypassed)."; }
      if (nNoLen) { msg += "\n" + nNoLen + " still need a Length before voltage drop can be checked."; }
      if (warn.length) {
        msg += "\n\nFlagged:\n" + warn.slice(0, 12).join("\n");
        if (warn.length > 12) { msg += "\n... and " + (warn.length - 12) + " more."; }
      } else {
        msg += "\nNo warnings.";
      }
      return SAZ.notify("Refresh Feeders", msg);
    });
  });
};


/* ---------- handlers ---------- */

function NewProject (event) { SAZ.run("New Project", function (ctx) {
  return SAZ.ensureProject(ctx, true);
})(event); }
function GoInstructions (event) { SAZ.run("Instructions", function (ctx) {
  return SAZ.goOrCreate(ctx, "Instructions");
})(event); }
function GoDashboard (event) { SAZ.run("Dashboard", function (ctx) {
  return SAZ.goOrCreate(ctx, "Dashboard");
})(event); }
function NewLoadCalc (event) { SAZ.run("New Load Calc", function (ctx) {
  return SAZ.createModule(ctx, "Load Calc");
})(event); }
function GoLoadCalc (event) { SAZ.run("Go To Load Calc", function (ctx) {
  return SAZ.goOrCreate(ctx, "Load Calc");
})(event); }
function AddSectionRow (event) { SAZ.run("Add Row", function (ctx) {
  return SAZ.lcAddRow(ctx);
})(event); }
function TrimSectionRows (event) { SAZ.run("Trim Empty Rows", function (ctx) {
  return SAZ.lcTrimRows(ctx);
})(event); }
function SendToPanel (event) { SAZ.run("Send to Panel", function (ctx) {
  return SAZ.uiSendToPanel(ctx);
})(event); }
function RefreshPlacement (event) { SAZ.run("Refresh Placement", function (ctx) {
  return SAZ.uiRefreshPlacement(ctx);
})(event); }
function ToggleSection (event) { SAZ.run("Toggle Section", function (ctx) {
  return SAZ.uiToggleSection(ctx);
})(event); }
function NewUtility (event) { SAZ.run("New Utility Demand (12 mo)", function (ctx) {
  return SAZ.createModule(ctx, "Utility Demand");
})(event); }
function GoUtility (event) { SAZ.run("Go To Utility Demand", function (ctx) {
  return SAZ.goOrCreate(ctx, "Utility Demand");
})(event); }
function NewLogging (event) { SAZ.run("New Logging Demand (30 day)", function (ctx) {
  return SAZ.createModule(ctx, "Logging Demand");
})(event); }
function GoLogging (event) { SAZ.run("Go To Logging Demand", function (ctx) {
  return SAZ.goOrCreate(ctx, "Logging Demand");
})(event); }
function AddMeter (event) { SAZ.run("Add Meter", function (ctx) {
  return SAZ.udAddMeter(ctx);
})(event); }
function NewPanelboard (event) { SAZ.run("New Panelboard", function (ctx) {
  return SAZ.uiCreate(ctx, "TPL_Panelboard", "Panelboard");
})(event); }
function NewSwitchboard (event) { SAZ.run("New Switchboard", function (ctx) {
  return SAZ.uiCreate(ctx, "TPL_Switchboard", "Switchboard");
})(event); }
function NewMCC (event) { SAZ.run("New MCC", function (ctx) {
  return SAZ.uiCreate(ctx, "TPL_MCC", "MCC");
})(event); }
function Link (event) { SAZ.run("Link Equipment", function (ctx) {
  return SAZ.uiLink(ctx);
})(event); }
function SubFeed (event) { SAZ.run("Sub-Feed", function (ctx) {
  return SAZ.uiSubFeed(ctx);
})(event); }
function FeedThru (event) { SAZ.run("Feed-Through", function (ctx) {
  return SAZ.uiFeedThru(ctx);
})(event); }
function Resize (event) { SAZ.run("Resize Panel", function (ctx) {
  return SAZ.uiResize(ctx);
})(event); }
function AddEquipRows (event) { SAZ.run("Add Circuit Rows", function (ctx) {
  return SAZ.uiEquipAddRows(ctx);
})(event); }
function TrimEquipRows (event) { SAZ.run("Trim Circuit Rows", function (ctx) {
  return SAZ.uiEquipTrimRows(ctx);
})(event); }
function RebuildGeo (event) { SAZ.run("Rebuild GEO_MAP", function (ctx) {
  return SAZ.uiRepair(ctx);   /* v2.5: Repair Schedule = restore damaged cells + rebuild GEO_MAP */
})(event); }
function RenameBoard (event) { SAZ.run("Rename Equipment", function (ctx) {
  return SAZ.uiRename(ctx);
})(event); }
function NewTransferSwitch (event) { SAZ.run("New Transfer Switch", function (ctx) {
  return SAZ.uiTransferSwitch(ctx);
})(event); }
function NewUpsSystem (event) { SAZ.run("New UPS System", function (ctx) {
  return SAZ.uiUpsSystem(ctx);
})(event); }
function NewGenerator (event) { SAZ.run("New Generator", function (ctx) {
  return SAZ.uiModule(ctx, "Generator");
})(event); }
function GoGenerator (event) { SAZ.run("Go To Generator", function (ctx) {
  return SAZ.goOrCreate(ctx, "Generator");
})(event); }
function NewUPS (event) { SAZ.run("New UPS", function (ctx) {
  return SAZ.uiModule(ctx, "UPS");
})(event); }
function GoUPS (event) { SAZ.run("Go To UPS", function (ctx) {
  return SAZ.goOrCreate(ctx, "UPS");
})(event); }
function NewTransformers (event) { SAZ.run("New Transformers", function (ctx) {
  return SAZ.createModule(ctx, "Transformers");
})(event); }
function GoTransformers (event) { SAZ.run("Go To Transformers", function (ctx) {
  return SAZ.goOrCreate(ctx, "Transformers");
})(event); }
function NewShortCircuit (event) { SAZ.run("New Short Circuit", function (ctx) {
  return SAZ.createModule(ctx, "Short Circuit");
})(event); }
function GoShortCircuit (event) { SAZ.run("Go To Short Circuit", function (ctx) {
  return SAZ.goOrCreate(ctx, "Short Circuit");
})(event); }
function NewVoltageDrop (event) { SAZ.run("New Voltage Drop", function (ctx) {
  return SAZ.createModule(ctx, "Voltage Drop");
})(event); }
function GoVoltageDrop (event) { SAZ.run("Go To Voltage Drop", function (ctx) {
  return SAZ.goOrCreate(ctx, "Voltage Drop");
})(event); }
function NewArcFlash (event) { SAZ.run("New Arc Flash", function (ctx) {
  return SAZ.createModule(ctx, "Arc Flash");
})(event); }
function GoArcFlash (event) { SAZ.run("Go To Arc Flash", function (ctx) {
  return SAZ.goOrCreate(ctx, "Arc Flash");
})(event); }
function NewFeederSchedule (event) { SAZ.run("New Feeder Schedule", function (ctx) {
  return SAZ.createModule(ctx, "Feeder Schedule");
})(event); }
function GoFeederSchedule (event) { SAZ.run("Go To Feeder Schedule", function (ctx) {
  return SAZ.goOrCreate(ctx, "Feeder Schedule");
})(event); }
function GoStudiesView (event) { SAZ.run("Go To Combined Studies", function (ctx) {
  return SAZ.goOrCreate(ctx, "Studies");
})(event); }
function RefreshFeeders (event) { SAZ.run("Refresh Feeders", function (ctx) {
  return SAZ.refreshFeeders(ctx);
})(event); }
function RefreshStudies (event) { SAZ.run("Refresh Studies", function (ctx) {
  return SAZ.refreshStudies(ctx);
})(event); }
function RefreshVoltageDrop (event) { SAZ.run("Refresh Voltage Drop", function (ctx) {
  return SAZ.refreshStudySheet(ctx, "vd");
})(event); }
function RefreshShortCircuit (event) { SAZ.run("Refresh Short Circuit", function (ctx) {
  return SAZ.refreshStudySheet(ctx, "sc");
})(event); }
function RefreshArcFlash (event) { SAZ.run("Refresh Arc Flash", function (ctx) {
  return SAZ.refreshStudies(ctx);
})(event); }
function NewReport (event) { SAZ.run("New Report", function (ctx) {
  return SAZ.createModule(ctx, "Report");
})(event); }
function GoReport (event) { SAZ.run("Go To Report", function (ctx) {
  return SAZ.goOrCreate(ctx, "Report");
})(event); }
function GoRecon (event) { SAZ.run("Reconciliation", function (ctx) {
  return SAZ.uiGoRecon(ctx);
})(event); }
function ToggleSizing (event) { SAZ.run("Toggle Sizing Columns", function (ctx) {
  return SAZ.uiToggleSizingColumns(ctx);
})(event); }
function ToggleXfmrOverrides (event) { SAZ.run("Show / Hide Transformer Overrides", function (ctx) {
  return SAZ.uiToggleXfmrOverrides(ctx);
})(event); }

function ToggleMatrix (event) { SAZ.run("Toggle Load Matrix", function (ctx) {
  return SAZ.uiToggleLoadMatrix(ctx);
})(event); }


/* ==========================================================
   v2.5 DISTRIBUTION MAP ENGINE  (called from the task pane - taskpane.html loads
   this file too, so the map and the ribbon run the SAME code; nothing is copied)
   One Apply does everything for the selected board:
     - Spaces:     resizes the schedule (repair runs first)
     * Mains:      registry Mains column -> the schedule's Mains Type formula
     - Fed from + Feed type + Ckt:
         Sub-Feed      breaker on the parent at circuit #ckt (first fit if blank):
                       description, poles from the child's voltage/phase, LIVE
                       demand link into the parent's PNL column; transformer
                       auto-inserted when voltages differ. Changing ckt = move.
         Feed-Through  / Tap: no breaker - the child joins the parent's
                       FEED-THROUGH LOADS (demand, connected, phase A/B/C,
                       receptacle) and its source fault follows the parent bus.
         utility / blank  top of the tree: any old parent link is removed.
       Switching feed type removes the old link first, so a load is never
       counted twice.
   ========================================================== */
SAZ.exec = function (body) {
  return Excel.run(function (ctx) {
    return SAZ.seedUnlock(ctx).then(function () { return body(ctx); }).then(
      function (out) { return SAZ.seedRelock(ctx).then(function () { return out; }); },
      function (err) { return SAZ.seedRelock(ctx).then(function () { throw err; }, function () { throw err; }); });
  });
};
SAZ.isUtilityTag = function (ctx, tag) {
  if (!tag) { return Promise.resolve(true); }
  var n = ctx.workbook.names.getItemOrNullObject("Util_Tags");
  n.load("isNullObject");
  return ctx.sync().then(function () {
    if (n.isNullObject) { return false; }
    var rg = n.getRange(); rg.load("values");
    return ctx.sync().then(function () {
      return rg.values[0].some(function (v) { return String(v || "").trim().toLowerCase() === String(tag).toLowerCase(); });
    });
  });
};
SAZ.registrySet = function (ctx, tag, cols) {          /* cols: { 3: fed, 4: mains, 8: ft, 9: ckt } */
  return SAZ.registryRange(ctx).then(function (rg) {
    rg.load("values,rowIndex,columnIndex,worksheet/name");
    return ctx.sync().then(function () {
      for (var i = 0; i < rg.values.length; i++) {
        if (String(rg.values[i][0] || "").trim().toLowerCase() === String(tag).toLowerCase()) {
          var sh = ctx.workbook.worksheets.getItem(rg.worksheet.name), at = i;
          return SAZ.withUnprotected(ctx, sh, function () {
            Object.keys(cols).forEach(function (k) {
              sh.getRangeByIndexes(rg.rowIndex + at, rg.columnIndex + Number(k), 1, 1).values = [[cols[k]]];
            });
            return ctx.sync().then(function () { return true; });
          });
        }
      }
      return false;
    });
  });
};
/* remove every upstream link of a board: breaker on any parent + feed-through rows */
SAZ.mapDetach = function (ctx, childTag) {
  return SAZ.registryRecords(ctx).then(function (reg) {
    var me = reg.filter(function (r) { return r.tag.toLowerCase() === childTag.toLowerCase(); })[0];
    var old = me ? me.fed : "";
    var sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    return ctx.sync().then(function () {
      var has = function (t) { return t && sheets.items.some(function (s) { return s.name.toLowerCase() === String(t).toLowerCase(); }); };
      if (!has(old)) { return { from: old, cleared: 0 }; }
      return SAZ.clearParentCircuit(ctx, old, [childTag, "T-" + childTag]).catch(function () { return 0; })
        .then(function (a) {
          return SAZ.feedThruRemove(ctx, old, childTag).catch(function () { return 0; })
            .then(function (b) { return { from: old, cleared: (a || 0) + (b || 0) }; });
        });
    });
  });
};
/* v2.12: System Voltage has ONE home - the Distribution Map (or the deploy dialog for a new
   board). The schedule's VOLT_CELL is locked and shown as a calculated (red) value; this
   writer is the only path that changes it. Returns true when the voltage actually changed. */
SAZ.setBoardVoltage = function (ctx, tag, volt) {
  return SAZ.geo(ctx, tag).then(function (g) {
    var vc = SAZ.geoUsable(g.VOLT_CELL);
    if (!vc) { return false; }
    var sh = ctx.workbook.worksheets.getItem(tag), cell = sh.getRange(vc);
    cell.load("values");
    return ctx.sync().then(function () {
      var cur = String(cell.values[0][0] || "").trim(), want = String(volt || "").trim();
      var change = !!want && want !== cur;
      return SAZ.withUnprotected(ctx, sh, function () {
        var c = sh.getRange(vc);
        c.format.protection.locked = true; c.format.fill.clear(); c.format.font.color = SAZ.CALC_FONT;
        if (change) { c.values = [[want]]; }
        return ctx.sync();
      }).then(function () { return change; });
    });
  });
};

SAZ.mapApply = function (ctx, req) {
  /* req: { tag, fed, ft, ckt, mains, spaces, len, lenPri, lenEm, mat, race, exist, volt } */
  var C = String(req.tag || "").trim(), P = String(req.fed || "").trim();
  var ft = String(req.ft || "").trim(), msgs = [];
  if (!C) { return Promise.reject(new Error("Select a board first.")); }
  if (P && P.toLowerCase() === C.toLowerCase()) { return Promise.reject(new Error(C + " cannot feed itself.")); }
  var step = Promise.resolve();
  /* 1. spaces */
  if (req.spaces) {
    step = step.then(function () {
      return SAZ.geo(ctx, C).then(function (g) {
        var cur = SAZ.currentSpaces(g), want = parseInt(req.spaces, 10);
        if (!want || want === cur) { return SAZ.repairSchedule(ctx, C); }
        return SAZ.resizePanel(ctx, C, want).then(function (o) {
          msgs.push("Resized to " + want + " spaces (was " + cur + ").");
          return o;
        });
      });
    });
  } else {
    step = step.then(function () { return SAZ.repairSchedule(ctx, C).catch(function () { return null; }); });
  }
  /* 1b. v2.12 system voltage - set here (Distribution Map) and pushed to the schedule; the
      sub-feed link below then re-decides whether a transformer is required. */
  step = step.then(function () {
    if (req.volt === undefined) { return null; }
    return SAZ.voltChoices(ctx).then(function (vc) {
      var want = String(req.volt || "").trim();
      if (want && vc.list.length && vc.list.indexOf(want) < 0) { throw new Error("\"" + want + "\" is not on the system voltage list (Volt_Sys)."); }
      return SAZ.setBoardVoltage(ctx, C, want).then(function (ch) {
        if (ch) { ctx.workbook.application.calculate(Excel.CalculationType.full); msgs.push("System voltage set to " + want + "."); }
        return ctx.sync();
      });
    });
  });
  /* 2. mains */
  step = step.then(function () {
    if (req.mains === undefined) { return null; }
    return SAZ.registrySet(ctx, C, { 4: req.mains || "" });
  });
  /* 2b. v2.11 feeder inputs - single home in the registry (Dashboard R:W); the Feeder Schedule
     (Length, Material, Conduit Type, EXIST?) and every schedule read them by tag, so a typed
     value can never drift onto another feeder when the derived schedule re-sorts. */
  step = step.then(function () {
    var nb = function (v) { var s = String(v == null ? "" : v).replace(/[^0-9.]/g, ""); var n = parseFloat(s); return isNaN(n) ? "" : n; };
    var f = {};
    if (req.len !== undefined) { f[17] = nb(req.len); }
    if (req.lenPri !== undefined) { f[18] = nb(req.lenPri); }
    if (req.lenEm !== undefined) { f[19] = nb(req.lenEm); }
    if (req.mat !== undefined) { f[20] = /^(CU|AL)$/i.test(String(req.mat)) ? String(req.mat).toUpperCase() : ""; }
    if (req.race !== undefined) { f[21] = String(req.race || "").trim(); }
    if (req.exist !== undefined) { f[22] = /^y/i.test(String(req.exist || "")) ? "Y" : ""; }
    if (!Object.keys(f).length) { return null; }
    return SAZ.registrySet(ctx, C, f);
  });
  /* 3. upstream link */
  return step.then(function () { return SAZ.isUtilityTag(ctx, P); }).then(function (isUtil) {
    if (isUtil) {
      return SAZ.mapDetach(ctx, C).then(function (d) {
        if (d.cleared) { msgs.push("Removed from " + d.from + "."); }
        return SAZ.registrySet(ctx, C, { 3: P, 8: P ? "Utility / Service" : "", 9: "" });
      }).then(function () { msgs.push(P ? "Fed from utility " + P + "." : "Unlinked (top of the map)."); });
    }
    if (ft === "Feed-Through" || ft === "Tap / Splice") {
      return SAZ.mapDetach(ctx, C).then(function () {
        return SAZ.geo(ctx, C).then(function (cg) {
          return SAZ.geo(ctx, P).then(function (pg) {
            var ref = SAZ.sheetRef(C, cg.DEMAND_CELL);
            return SAZ.feedThruAppend(ctx, P, C, ref, C).then(function (res) {
              var src = SAZ.geoUsable(cg.SRCFAULT_CELL), bus = SAZ.geoUsable(pg.BUSFAULT_CELL);
              var csh = ctx.workbook.worksheets.getItem(C);
              return SAZ.withUnprotected(ctx, csh, function () {
                if (src && bus) { csh.getRange(src).formulas = [[ SAZ.sheetRef(P, bus) ]]; }
                return ctx.sync();
              }).then(function () {
                return SAZ.registrySet(ctx, C, { 3: P, 8: ft, 9: "FT" });
              }).then(function () {
                msgs.push(C + " added to " + P + " FEED-THROUGH LOADS (row " + res.row + ", no breaker)." +
                  (res.warn && res.warn.length ? " Only demand linked for: " + res.warn.join(", ") + "." : ""));
              });
            });
          });
        });
      });
    }
    /* Sub-Feed (default) */
    var reg0 = null;
    return SAZ.registryRecords(ctx).then(function (reg) {
      reg0 = reg.filter(function (r) { return r.tag.toLowerCase() === C.toLowerCase(); })[0];
      /* leaving a feed-through parent? take it out of that block first */
      if (reg0 && reg0.fed && /feed-through|tap/i.test(reg0.ft)) {
        return SAZ.feedThruRemove(ctx, reg0.fed, C).catch(function () { return 0; });
      }
      return 0;
    }).then(function () {
      var num = parseInt(String(req.ckt || "").replace(/[^0-9]/g, ""), 10);
      if (!num) { return null; }
      return SAZ.cktGrid(ctx, P).then(function (grid) {
        var rec = grid.byNum[num];
        if (!rec) { throw new Error("Circuit " + num + " does not exist on " + P + "."); }
        return { row: rec.row };
      });
    }).then(function (slot) {
      return SAZ.linkEquipment(ctx, P, C, slot, { quiet: true });
    }).then(function (r) {
      if (r && r.text) { msgs.push(r.text); }
      return SAZ.registrySet(ctx, C, { 8: "Sub-Feed" });
    });
  }).then(function () {
    ctx.workbook.application.calculate(Excel.CalculationType.recalculate);
    return ctx.sync().then(function () { return msgs.join("\n"); });
  });
};
SAZ.mapUnlink = function (ctx, tag) {
  return SAZ.mapDetach(ctx, tag).then(function (d) {
    return SAZ.registrySet(ctx, tag, { 3: "", 8: "", 9: "" }).then(function () {
      return d.cleared ? tag + " removed from " + d.from + "." : tag + " is now unlinked.";
    });
  });
};
SAZ.mapRepair = function (ctx, tag) {
  return SAZ.repairSchedule(ctx, tag).then(function (o) {
    return o.repaired ? o.repaired + " damaged cell(s) restored on " + tag + "." : tag + ": nothing damaged.";
  });
};

/* ---------- associations ---------- */

Office.onReady(function () {
  /* v2.5: taskpane.html also loads this file for the Distribution Map; there is no
     ribbon runtime in the pane, so associate only where it is supported. */
  if (!Office.actions || typeof Office.actions.associate !== "function") { return; }
  try {
  Office.actions.associate("NewProject", NewProject);
  Office.actions.associate("GoInstructions", GoInstructions);
  Office.actions.associate("GoDashboard", GoDashboard);
  Office.actions.associate("NewLoadCalc", NewLoadCalc);
  Office.actions.associate("GoLoadCalc", GoLoadCalc);
Office.actions.associate("AddSectionRow", AddSectionRow);
Office.actions.associate("TrimSectionRows", TrimSectionRows);
Office.actions.associate("ToggleSection", ToggleSection);
Office.actions.associate("SendToPanel", SendToPanel);
  Office.actions.associate("MoveCircuit", MoveCircuit);
Office.actions.associate("NewTransferSwitch", NewTransferSwitch);
Office.actions.associate("NewUpsSystem", NewUpsSystem);
Office.actions.associate("RefreshPlacement", RefreshPlacement);
  Office.actions.associate("NewUtility", NewUtility);
  Office.actions.associate("GoUtility", GoUtility);
  Office.actions.associate("NewLogging", NewLogging);
  Office.actions.associate("GoLogging", GoLogging);
  Office.actions.associate("AddMeter", AddMeter);
  Office.actions.associate("NewPanelboard", NewPanelboard);
  Office.actions.associate("NewSwitchboard", NewSwitchboard);
  Office.actions.associate("NewMCC", NewMCC);
  Office.actions.associate("Link", Link);
  Office.actions.associate("SubFeed", SubFeed);
  Office.actions.associate("FeedThru", FeedThru);
  Office.actions.associate("Resize", Resize);
  Office.actions.associate("AddEquipRows", AddEquipRows);
  Office.actions.associate("TrimEquipRows", TrimEquipRows);
  Office.actions.associate("RebuildGeo", RebuildGeo);
  Office.actions.associate("NewGenerator", NewGenerator);
  Office.actions.associate("GoGenerator", GoGenerator);
  Office.actions.associate("NewUPS", NewUPS);
  Office.actions.associate("GoUPS", GoUPS);
  Office.actions.associate("NewTransformers", NewTransformers);
  Office.actions.associate("GoTransformers", GoTransformers);
  Office.actions.associate("NewShortCircuit", NewShortCircuit);
  Office.actions.associate("GoShortCircuit", GoShortCircuit);
  Office.actions.associate("NewVoltageDrop", NewVoltageDrop);
  Office.actions.associate("GoVoltageDrop", GoVoltageDrop);
  Office.actions.associate("NewArcFlash", NewArcFlash);
  Office.actions.associate("GoArcFlash", GoArcFlash);
  Office.actions.associate("NewFeederSchedule", NewFeederSchedule);
  Office.actions.associate("GoFeederSchedule", GoFeederSchedule);
  Office.actions.associate("GoStudiesView", GoStudiesView);
  Office.actions.associate("RefreshFeeders", RefreshFeeders);
  Office.actions.associate("RefreshStudies", RefreshStudies);
  Office.actions.associate("RefreshVoltageDrop", RefreshVoltageDrop);
  Office.actions.associate("RefreshShortCircuit", RefreshShortCircuit);
  Office.actions.associate("RefreshArcFlash", RefreshArcFlash);
  Office.actions.associate("NewReport", NewReport);
  Office.actions.associate("GoReport", GoReport);
  Office.actions.associate("GoRecon", GoRecon);
  Office.actions.associate("ToggleSizing", ToggleSizing);
  Office.actions.associate("ToggleXfmrOverrides", ToggleXfmrOverrides);
  Office.actions.associate("ToggleMatrix", ToggleMatrix);
Office.actions.associate("RenameBoard", RenameBoard);
  ["01","02","03","04","05","06","07","08","09","10","11","12","13","14","15"].forEach(function (n) {   /* v2.10 Help ▸ Instructions chapters */
    Office.actions.associate("GoCh" + n, SAZ.run("Instructions - chapter " + n, function (ctx) { return SAZ.goChapter(ctx, n); }));
  });
  } catch (e) { /* not a ribbon runtime */ }
});

/* ---------- Load Calc section rows  (added: replaces the VBA Add/Trim Area Rows) ---------- */

/* Every Load Calc section is a real Excel Table, so adding a table row inherits
   column formatting and calculated-column formulas automatically. That is why
   these two commands can serve all twelve sections instead of one per section. */

SAZ.lcSections = [
  { table: "LC_Areas",   label: "Baseline Square Footage" },
  { table: "LC_WSF",     label: "Watts per Square Foot" },
  { table: "LC_Site",    label: "Site / Exterior" },
  { table: "LC_House",   label: "House / Miscellaneous Equipment" },
  { table: "LC_Dept",    label: "Departmental Equipment" },
  { table: "LC_Refrig",  label: "Refrigeration" },
  { table: "LC_Mech",    label: "Mechanical" },
  { table: "LC_Plumb",   label: "Plumbing" },
  { table: "LC_Kitchen", label: "Kitchen" },
  { table: "LC_Elev",    label: "Elevators" },
  { table: "LC_EV",      label: "EV Charging" },
  { table: "LC_Data",    label: "Data / IT" },
  { table: "LC_Fire",    label: "Fire / Life Safety" },
  { table: "LC_Process", label: "Process / Industrial" }
];

/* Add Row / Trim Empty Rows also serve departments added at run time. */
SAZ.lcSectionAtCursor = function (ctx) {
  var sel = ctx.workbook.getSelectedRange();
  sel.load("address");
  return SAZ.sectionRegistry(ctx).catch(function () { return []; }).then(function (reg) {
    var specs = SAZ.lcSections.slice();
    reg.forEach(function (e) {
      if (!specs.some(function (s) { return s.table === e.table; })) { specs.push({ table: e.table, label: SAZ.sectionName(e.label) }); }
    });
    var probes = specs.map(function (s) {
      var t = ctx.workbook.tables.getItemOrNullObject(s.table);
      var r = t.getRange().getIntersectionOrNullObject(sel);
      r.load("isNullObject");
      t.load("isNullObject");
      return { spec: s, table: t, hit: r };
    });
    return ctx.sync().then(function () {
      for (var i = 0; i < probes.length; i++) {
        if (!probes[i].table.isNullObject && !probes[i].hit.isNullObject) { return probes[i]; }
      }
      return null;
    });
  });
};

SAZ.lcNoSection = function () {
  var names = SAZ.lcSections.map(function (s) { return "\u2022 " + s.label; }).join("\n");
  return SAZ.notify("Pick a section first",
    "Click any cell inside the section you want to change, then run this again.\n\n" +
    "Load Calc sections:\n" + names);
};


/* Remove trailing empty rows from the section under the cursor. Only rows at the
   BOTTOM are removed and at least one data row is always kept, so a gap in the
   middle of the section is never silently closed up and the table can never be
   reduced to a header with no fill pattern to copy from. */
SAZ.lcTrimRows = function (ctx) {
  return SAZ.lcSectionAtCursor(ctx).then(function (found) {
    if (!found) return SAZ.lcNoSection();
    var body = found.table.getDataBodyRange();
    body.load("values,rowCount");
    return ctx.sync().then(function () {
      var vals = body.values, keep = vals.length;
      while (keep > 1) {
        var row = vals[keep - 1], blank = true;
        for (var c = 0; c < row.length; c++) {
          var v = row[c];
          if (v !== "" && v !== null && v !== 0) { blank = false; break; }
        }
        if (!blank) break;
        keep--;
      }
      var remove = vals.length - keep;
      if (remove === 0) {
        return SAZ.notify("Nothing to trim",
          found.spec.label + " has no empty rows at the bottom.");
      }
      for (var i = 0; i < remove; i++) { found.table.rows.getItemAt(keep).delete(); }
      return ctx.sync().then(function () {
        return SAZ.notify("Rows trimmed",
          "Removed " + remove + " empty row(s) from " + found.spec.label +
          ". " + keep + " row(s) remain.");
      });
    });
  });
};

/* ---------- narrow-section fallback ---------- */

/* Baseline Square Footage (LC_Areas) and Watts per Square Foot (LC_WSF) occupy
   columns B:G, while the eight equipment sections below them span A:U. A table
   row insert shifts only its own columns, so inserting into a B:G table would
   partially move the wider tables underneath - Excel refuses this outright, and
   it is also why Tab does not extend those two sections.

   A FULL sheet-row insert shifts every table whole, so it is allowed. This
   fallback inserts an entire worksheet row inside the table body, which Excel
   absorbs into the table and fills with the calculated-column formulas. */

SAZ.lcAddRowFullWidth = function (ctx, found) {
  var body = found.table.getDataBodyRange();
  body.load("address,rowCount");
  return ctx.sync().then(function () {
    var m = body.address.match(/\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)/);
    var lastRow = m ? parseInt(m[4], 10) : null;
    if (!lastRow) {
      return SAZ.notify("Could not add a row",
        "The " + found.spec.label + " table did not report a usable range.");
    }
    /* insert a whole row at the last data row: it lands inside the table body,
       so the table grows rather than the row landing outside it */
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.getRange(lastRow + ":" + lastRow).insert(Excel.InsertShiftDirection.down);
    return ctx.sync().then(function () {
      /* the new sheet row sits at lastRow; the old last row is now lastRow + 1.
         Paint the whole new row from it so every column keeps its look. */
      sheet.getRange(lastRow + ":" + lastRow)
        .copyFrom(sheet.getRange((lastRow + 1) + ":" + (lastRow + 1)), Excel.RangeCopyType.formats);
      return ctx.sync();
    }).then(function () {
      var after = found.table.getDataBodyRange();
      after.load("rowCount");
      return ctx.sync().then(function () {
        return SAZ.notify("Row added",
          found.spec.label + " now has " + after.rowCount + " rows." +
          "\n\nA full worksheet row was inserted because this section is narrower than the tables below it.");

      });
    });
  });
};

/* Restore the section's look on rows a table insert just created. A table row
   insert takes the table STYLE's formatting, not the hand-applied fill, font
   colour, borders and number formats of the rows above - which is what blew
   the formatting out. Paint the new rows from the FIRST data row (the section's
   pattern row) with a formats-only copy, so values and formulas are untouched.
   Spans the table's own columns only, so neighbouring tables are never touched. */
SAZ.matchTableRowFormats = function (ctx, table, added) {
  var body = table.getDataBodyRange();
  body.load("rowCount");
  return ctx.sync().then(function () {
    if (body.rowCount <= 1 || added < 1) { return null; }
    var pattern = body.getRow(0);
    var from = Math.max(1, body.rowCount - added);
    var target = body.getRow(from).getResizedRange(body.rowCount - 1 - from, 0);
    target.copyFrom(pattern, Excel.RangeCopyType.formats);
    return ctx.sync();
  });
};

/* Add one blank row to the section under the cursor. Try the normal table insert,
   and if Excel refuses because the section is narrower than the tables below it,
   fall back to a full worksheet row. */
SAZ.lcAddRow = function (ctx) {
  return SAZ.lcSectionAtCursor(ctx).then(function (found) {
    if (!found) return SAZ.lcNoSection();
    found.table.rows.add();
    return ctx.sync().then(function () {
      return SAZ.matchTableRowFormats(ctx, found.table, 1);
    }).then(function () {
      var after = found.table.getDataBodyRange();
      after.load("rowCount");
      return ctx.sync().then(function () {
        var last = found.table.getDataBodyRange().getLastRow();
        last.select();
        return ctx.sync().then(function () {
          return SAZ.notify("Row added",
            found.spec.label + " now has " + after.rowCount + " rows.");
        });
      });
    }, function () {
      /* InsertDeleteConflict - narrow section, use the full-row route */
      return SAZ.lcAddRowFullWidth(ctx, found);
    });
  });
};

/* ---------- Utility Demand: add a meter ---------- */

/* A meter block is 19 rows: the METER n banner, the utility/rate/serves inputs,
   the column headers, twelve billing cycles, the METER PEAK summary and a blank
   spacer. The RESULTS section below already aggregates with whole-column
   COUNTIF/SUMIF on "METER PEAK", so a new block is picked up with no change to
   any results formula. Template geometry is taken from the FIRST block so the
   command keeps working if the sheet is edited. */

SAZ.UD_BLOCK_ROWS = 20;   /* v1.2: +1 spacer row under the Source / Utility lines */
/* Utility Demand is a PROTECTED sheet (house standard: all thirteen allowances
   false, selectionMode Normal). Row inserts, copyFrom and clears of locked cells
   all fail under protection, so the core runs inside an unprotect/reprotect pair
   - the same pattern as SAZ.resizePanel - and the sheet is reprotected on the
   FAILURE path as well, so an error can never leave it unprotected. */
SAZ.PROTECT_OPTS = {
  allowAutoFilter: false, allowDeleteColumns: false, allowDeleteRows: false,
  allowEditObjects: false, allowEditScenarios: false, allowFormatCells: false,
  allowFormatColumns: false, allowFormatRows: false, allowInsertColumns: false,
  allowInsertHyperlinks: false, allowInsertRows: false, allowPivotTables: false,
  allowSort: false, selectionMode: "Normal"
};
SAZ.UD_PROTECT_OPTS = SAZ.PROTECT_OPTS;   /* kept for existing callers */

/* Run fn with the sheet unprotected, then restore house-standard protection on
   BOTH the success and failure paths, so an error can never leave a sheet open.
   fn must return a promise. */
SAZ.withUnprotected = function (ctx, sheet, fn) {
  sheet.protection.load("protected");
  return ctx.sync().then(function () {
    var wasProt = sheet.protection.protected;
    var reprotect = function () {
      if (!wasProt) { return Promise.resolve(); }
      sheet.protection.protect(SAZ.PROTECT_OPTS);
      return ctx.sync();
    };
    if (wasProt) { sheet.protection.unprotect(); }
    return ctx.sync().then(function () { return fn(); }).then(function (out) {
      return reprotect().then(function () { return out; });
    }, function (err) {
      return reprotect().then(function () { throw err; }, function () { throw err; });
    });
  });
};

SAZ.udAddMeter = function (ctx) {
  var sheet = ctx.workbook.worksheets.getItem("Utility Demand");
  sheet.protection.load("protected");
  return ctx.sync().then(function () {
    var wasProt = sheet.protection.protected;
    var reprotect = function () {
      if (!wasProt) { return Promise.resolve(); }
      sheet.protection.protect(SAZ.UD_PROTECT_OPTS);
      return ctx.sync();
    };
    /* rows cannot be inserted or deleted on a protected sheet */
    if (wasProt) { sheet.protection.unprotect(); }
    return ctx.sync().then(function () {
      return SAZ.udAddMeterCore(ctx);
    }).then(function (out) {
      return reprotect().then(function () { return out; });
    }, function (err) {
      return reprotect().then(function () { throw err; }, function () { throw err; });
    });
  });
};


SAZ.udAddMeterCore = function (ctx) {
  var sheet = ctx.workbook.worksheets.getItem("Utility Demand");
  var colB = sheet.getRange("B1:B400");
  var colA = sheet.getRange("A1:A400");
  colB.load("values");
  colA.load("values");
  return ctx.sync().then(function () {
    var firstBanner = -1, lastPeak = -1, count = 0;
    for (var i = 0; i < colA.values.length; i++) {
      var a = String(colA.values[i][0] || "");
      if (a.indexOf("METER ") === 0 && firstBanner < 0) firstBanner = i + 1;
    }
    for (var j = 0; j < colB.values.length; j++) {
      if (String(colB.values[j][0] || "") === "METER PEAK") { lastPeak = j + 1; count++; }
    }
    if (firstBanner < 0 || lastPeak < 0) {
      return SAZ.notify("Could not find a meter block",
        "Expected a \"METER 1\" heading in column A and a \"METER PEAK\" row in column B on the Utility Demand sheet.");
    }
    var insertAt = lastPeak + 2;
    var newBanner = insertAt;
    var offset = newBanner - firstBanner;
    var rows = sheet.getRange(insertAt + ":" + (insertAt + SAZ.UD_BLOCK_ROWS - 1));
    rows.insert(Excel.InsertShiftDirection.down);
    return ctx.sync().then(function () {
      var src = sheet.getRange(firstBanner + ":" + (firstBanner + SAZ.UD_BLOCK_ROWS - 1));
      var dst = sheet.getRange(newBanner + ":" + (newBanner + SAZ.UD_BLOCK_ROWS - 1));
      dst.copyFrom(src, Excel.RangeCopyType.all);
      return ctx.sync().then(function () {
        return SAZ.udCleanMeterBlock(ctx, sheet, firstBanner, newBanner, offset, count + 1);
      });
    });
  });
};

/* Renumber the new block, clear the copied bill data, and repoint the one formula
   that uses an absolute reference. The METER PEAK label cell reads the block own
   meter number via $C$15 in block 1; copied verbatim it would still point at
   block 1, so it is rewritten against the new block row. */
SAZ.udCleanMeterBlock = function (ctx, sheet, firstBanner, newBanner, offset, meterNo) {
  sheet.getRange("A" + newBanner).values = [["METER " + meterNo]];
  /* inputs on the identity row: meter #, source (utility name below it is a Dashboard link), rate schedule, serves */
  var idRow = newBanner + 2;
  sheet.getRange("C" + idRow).clear(Excel.ClearApplyTo.contents);
  sheet.getRange("F" + idRow).clear(Excel.ClearApplyTo.contents);
  sheet.getRange("I" + idRow).clear(Excel.ClearApplyTo.contents);
  sheet.getRange("L" + idRow).clear(Excel.ClearApplyTo.contents);
  sheet.getRange("L" + (idRow + 1)).clear(Excel.ClearApplyTo.contents);   /* v2.11: Metered Board - blank = service board */
  /* twelve billing cycles: clear the entered columns only, keeping Cycle numbers,
     PF Used defaults and every calculated column intact */
  var firstCycle = newBanner + 6;
  var lastCycle = firstCycle + 11;
  sheet.getRange("C" + firstCycle + ":I" + lastCycle).clear(Excel.ClearApplyTo.contents);
  /* repoint the peak-row label at this block own meter number */
  var peakRow = newBanner + 18;
  sheet.getRange("C" + peakRow).formulas =
    [["=IF($C$" + idRow + "=\"\",\"(enter meter #)\",\"Meter \"&$C$" + idRow + ")"]];
  return ctx.sync().then(function () {
    sheet.getRange("C" + idRow).select();
    return ctx.sync().then(function () {
      return SAZ.notify("Meter " + meterNo + " added",
        "A blank meter block was added below the last one. Enter the meter number, pick its Source (the utility name follows the Dashboard), the rate schedule and twelve billing cycles.\n\nThe NEC 220.87 results below update on their own \u2014 they count and sum every METER PEAK row, so nothing needs rewiring.");
    });
  });
};

/* ==========================================================
   NEW PROJECT  -  Artifact A, inserted WHOLE
   Port of VBA EnsureProject (VBA_Code rows 166-219).

   The calculation sheets are ONE connected component: Dashboard reads 8 sheets
   and is read by 8 sheets. A sheet inserted without its dependencies present
   gets a PERMANENT #REF! that does NOT heal when the others arrive later, and
   the dependency graph contains cycles regardless. So the core travels whole or
   it does not work at all. See WEB_Templates A26-A32.

   Calculation is forced to manual across the insert: the seed is formula-heavy
   and recalculating a half-built dependency graph can bring Excel down.
   ========================================================== */

SAZ.PROJECT_ANCHOR = "Dashboard";

SAZ.ensureProject = function (ctx, announce) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  return ctx.sync().then(function () {
    var before = sheets.items.map(function (s) { return s.name; });
    var hasAnchor = before.some(function (n) {
      return n.toLowerCase() === SAZ.PROJECT_ANCHOR.toLowerCase();
    });
    if (hasAnchor) {
      return SAZ.notify("Project already present",
        "This workbook already has a " + SAZ.PROJECT_ANCHOR + " tab, so the scaffold " +
        "is already in place.\n\nRun New Project from a blank workbook instead.");
    }
    return fetch(CFG.host + "/" + CFG.projectFile).then(function (res) {
      if (!res.ok) {
        throw new Error("Project seed not reachable (" + res.status + "). The SÄZÄN PowerCalc template library could not be reached - check your connection and try again.");
      }
      return res.arrayBuffer();
    }).then(function (buf) {
      var b64 = SAZ.toBase64(buf);
      ctx.application.load("calculationMode");
      return ctx.sync().then(function () {
        var saved = ctx.application.calculationMode;
        ctx.application.calculationMode = Excel.CalculationMode.manual;
        return ctx.sync().then(function () {
          /* NO sheetNamesToInsert - every sheet in the seed comes across together */
          return ctx.workbook.insertWorksheetsFromBase64(b64, {
            positionType: Excel.WorksheetPositionType.end
          });
        }).then(function () {
          return ctx.sync();
        }).then(function () {
          return SAZ.arrangeScaffold(ctx, before);
        }).then(function () {
          ctx.application.calculationMode = saved;
          return ctx.sync();
        }).then(function () {
          ctx.application.calculate(Excel.CalculationType.fullRebuild);
          return ctx.sync();
        }).then(function () {
          if (!announce) { return null; }
          return SAZ.notify("Project scaffold ready",
            "Only the Dashboard is showing. Choose a module from the ribbon and " +
            "its tab will appear.");
        });
      });
    });
  });
};

/* Dashboard to the front, everything else that arrived put away.
   Sheet names are DIFFED against the pre-insert list rather than hardcoded, so
   adding or renaming a sheet in the seed needs no change here.
   Two tiers of concealment, and the difference matters:
     module sheets -> hidden      (ribbon reveals them via SAZ.goOrCreate)
     TPL_* sheets  -> veryHidden  (clone sources; users must not wander in) */
SAZ.arrangeScaffold = function (ctx, before) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  return ctx.sync().then(function () {
    var anchor = null;
    var arrived = [];
    sheets.items.forEach(function (s) {
      if (before.indexOf(s.name) !== -1) { return; }
      if (s.name.toLowerCase() === SAZ.PROJECT_ANCHOR.toLowerCase()) { anchor = s; }
      else { arrived.push(s); }
    });
    if (!anchor) {
      throw new Error("The seed inserted but no " + SAZ.PROJECT_ANCHOR + " tab arrived. " +
        "Nothing was hidden - check the sheet list inside " + CFG.projectFile + ".");
    }
    anchor.position = 0;
    arrived.forEach(function (s) {
      var tpl = s.name.toUpperCase().indexOf("TPL_") === 0;
      s.visibility = tpl ? Excel.SheetVisibility.veryHidden
                         : Excel.SheetVisibility.hidden;
    });
    anchor.activate();
    return ctx.sync().then(function () { return SAZ.removeBlankStarters(ctx, before); });   /* v1.2: drop the blank workbook's own empty Sheet1 */
  });
};

/* v1.2: after New Project lands in a blank workbook, delete the sheets that
   were there BEFORE the seed (Sheet1, or its localised name) - but only if
   they are truly empty: no cells, no formatting, no charts, no shapes.
   Anything the user already put in the workbook is left alone. */
SAZ.removeBlankStarters = function (ctx, before) {
  var probes = (before || []).map(function (n) {
    var s = ctx.workbook.worksheets.getItemOrNullObject(n);
    var u = s.getUsedRangeOrNullObject(false);
    var ch = s.charts; ch.load("count");
    s.load("isNullObject");
    return { s: s, u: u, ch: ch };
  });
  probes.forEach(function (p) { p.u.load("isNullObject"); });
  return ctx.sync().then(function () {
    var shapeProbes = probes.map(function (p) {
      if (p.s.isNullObject) { return null; }
      try { var sh = p.s.shapes; sh.load("count"); return sh; } catch (e) { return null; }
    });
    return ctx.sync().catch(function () { return null; }).then(function () {
      probes.forEach(function (p, i) {
        if (p.s.isNullObject || !p.u.isNullObject || p.ch.count > 0) { return; }
        var sh = shapeProbes[i];
        if (sh && sh.count > 0) { return; }
        p.s.delete();
      });
      return ctx.sync();
    });
  });
};


/* ==========================================================
   GEO_MAP READER  -  never hardcode an equipment cell address

   Each equipment sheet carries a sheet-scoped GEO_MAP name (AS:AX) whose
   column AS holds the key and column AT the target cell. Resizing a panel
   moves those targets, and RebuildGeoMap rewrites them - so any handler that
   hardcodes "H10" is correct only until someone adds a circuit row.
   The VBA got this wrong: RegisterEquipment pins the demand cell to K7 for
   non-panelboards while the switchboard GEO_MAP says DEMAND_CELL = H10.
   ========================================================== */

SAZ.geo = function (ctx, sheetName) {
  var sheet = ctx.workbook.worksheets.getItem(sheetName);
  var nm = sheet.names.getItemOrNullObject("GEO_MAP");
  nm.load("isNullObject");
  return ctx.sync().then(function () {
    if (nm.isNullObject) {
      throw new Error("Sheet \"" + sheetName + "\" has no GEO_MAP. It was not created" +
        " from a TPL_ template, or the name was deleted. Run Rebuild GEO_MAP.");
    }
    var rg = nm.getRange();
    rg.load("values");
    return ctx.sync().then(function () {
      var map = {};
      rg.values.forEach(function (row) {
        var k = String(row[0] || "").trim();
        if (k && k !== "GEO_MAP") { map[k] = row[1]; }
      });
      return map;
    });
  });
};

/* A1 reference to another sheet, quoted correctly. Sheet names may contain
   spaces or apostrophes; an apostrophe is escaped by doubling it. */
SAZ.sheetRef = function (sheetName, cell) {
  var safe = String(sheetName).replace(/'/g, "''");
  var a1 = String(cell).replace(/([A-Za-z]+)([0-9]+)/, "$$$1$$$2");
  return "='" + safe + "'!" + a1;
};

/* Registry row for a newly created board.
   Voltage and demand are written as LIVE LINKS, not copied values: demand
   changes every time a circuit is added, so a static number would be wrong
   the moment the user typed anything. 'fed' stays (unassigned) until Link
   Equipment sets a parent. */
SAZ.registerFromGeo = function (ctx, sheetName, typeLabel) {
  return SAZ.geo(ctx, sheetName).then(function (g) {
    var voltCell = g.VOLT_CELL;
    var kvaCell  = g.DEMAND_CELL;
    return SAZ.registerEquipment(ctx, {
      tag:   sheetName,
      type:  typeLabel,
      volts: voltCell ? SAZ.sheetRef(sheetName, voltCell) : "",
      fed:   "(unassigned)",
      kva:   kvaCell  ? SAZ.sheetRef(sheetName, kvaCell)  : "",
      sheet: sheetName
    });
  });
};


/* ==========================================================
   FEED-THROUGH  -  carry a downstream board's load forward to this
   panel as a pass-through. Adds to demand, consumes NO breaker.

   NOTE ON THE VBA: FeedThruAppend's fallback is WRONG for the current
   templates and was deliberately not ported as written. It computes
   firstRow = headerRow + 1, which is the Description/kVA sub-header, and
   defaults the kVA column to "H" for switchboards when the real column is L.
   As written it would overwrite a header and post demand into the engine
   output column, where the block subtotal would never see it.
   This port derives the block from its actual structure instead:
     header row      = the cell containing "FEED-THROUGH"  (desc column)
     sub-header row  = header + 1, carries the "kVA" label (kVA column)
     first data row  = header + 2
     last data row   = the row above the subtotal row
   GEO_MAP keys are honoured first if a template ever defines them.
   ========================================================== */

SAZ.FEEDTHRU_MAX_SCAN = 40;

SAZ.feedThruBlock = function (ctx, sheetName) {
  return SAZ.geo(ctx, sheetName).then(function (g) {
    if (g.FEEDTHRU_DESC_COL && g.FEEDTHRU_KVA_COL && g.FEEDTHRU_FIRST) {
      return {
        descCol:  String(g.FEEDTHRU_DESC_COL),
        kvaCol:   String(g.FEEDTHRU_KVA_COL),
        firstRow: Number(g.FEEDTHRU_FIRST),
        lastRow:  Number(g.FEEDTHRU_LAST) || Number(g.FEEDTHRU_FIRST) + 3,
        source:   "GEO_MAP"
      };
    }
    var sheet = ctx.workbook.worksheets.getItem(sheetName);
    var scan = sheet.getRangeByIndexes(0, 0, SAZ.FEEDTHRU_MAX_SCAN, 20);
    scan.load("values");
    return ctx.sync().then(function () {
      var v = scan.values, hr = -1, hc = -1;
      for (var r = 0; r < v.length && hr < 0; r++) {
        for (var c = 0; c < v[r].length; c++) {
          if (String(v[r][c] || "").toUpperCase().indexOf("FEED-THROUGH") === 0) {
            hr = r; hc = c; break;
          }
        }
      }
      if (hr < 0) {
        throw new Error("Sheet \"" + sheetName + "\" has no FEED-THROUGH block." +
          " It predates the current template - re-copy from TPL_.");
      }
      /* kVA column: the "kVA" label on the sub-header row */
      var kc = -1, subRow = v[hr + 1] || [];
      for (var c2 = hc; c2 < subRow.length; c2++) {
        var lab = String(subRow[c2] || "").trim().toLowerCase();
        if (lab === "demand kva" || lab === "kva") { kc = c2; break; }
      }
      if (kc < 0) {
        throw new Error("Found the FEED-THROUGH header on \"" + sheetName +
          "\" but no kVA column beneath it. Do not guess - re-copy the template.");
      }
      /* last data row = the row above whichever row carries the subtotal */
      var stop = -1;
      for (var r2 = hr + 2; r2 < v.length; r2++) {
        if (/subtotal/i.test(String(v[r2][hc] || ""))) { stop = r2; break; }
      }
      if (stop < 0) { stop = hr + 6; }
      var col = function (n) { return String.fromCharCode(65 + n); };
      return {
        descCol: col(hc), kvaCol: col(kc),
        firstRow: hr + 3, lastRow: Math.min(stop, hr + 5),   /* 3 managed slots max - the manual row above the subtotal is never auto-filled */
        source: "derived"
      };
    });
  });
};


/* Write (or refresh) one feed-through row. Re-running with the same tag
   updates that row in place rather than consuming another slot (the
   same-tag search runs over ALL slots before an empty slot is taken, so
   a gap left by a removal never creates a duplicate). The managed rows
   are LOCKED on the template - writes go through withUnprotected. The
   manual row below the managed slots is never touched by the command. */
/* #245: besides the demand link, link the child's connected kVA, phase A/B/C
   kVA and receptacle totals so the parent's connected total, phase balance
   and combined receptacle demand factor include the downstream section.
   Locations come from GEO keys FT_CONN_COL / FT_DETAIL_COL (parent) and
   CONNECTED_CELL, PHASE_A/B/C_CELL, REC_TOT_CELL, REC_DEM_CELL (child).
   Missing keys (boards made before #245) link 0 and are reported in warn. */
SAZ.FT_DETAIL_KEYS = ["PHASE_A_CELL", "PHASE_B_CELL", "PHASE_C_CELL", "REC_TOT_CELL", "REC_DEM_CELL"];
SAZ.feedThruDetail = function (sheet, slot, pg, cg, childSheet) {
  var warn = [];
  var link = function (key) {
    var c = SAZ.geoUsable(cg[key]);
    if (!c) { warn.push(childSheet + " " + key); return 0; }
    return SAZ.sheetRef(childSheet, c);
  };
  var cc = SAZ.geoUsable(pg.FT_CONN_COL), dc = SAZ.geoUsable(pg.FT_DETAIL_COL);
  if (cc) { sheet.getRange(cc + slot).formulas = [[ link("CONNECTED_CELL") ]]; }
  if (dc) {
    var c0 = SAZ.gridColNum(dc);
    sheet.getRange(dc + slot + ":" + SAZ.gridColLet(c0 + 4) + slot).formulas =
      [ SAZ.FT_DETAIL_KEYS.map(link) ];
  }
  if (!cc || !dc) { warn.push("parent layout (board predates #245)"); }
  return warn;
};

SAZ.feedThruAppend = function (ctx, sheetName, childTag, demandRef, childSheet) {
  return SAZ.feedThruBlock(ctx, sheetName).then(function (b) {
    var gg = childSheet
      ? SAZ.geo(ctx, sheetName).then(function (pg) {      /* sequential: one sync in flight */
        return SAZ.geo(ctx, childSheet).then(function (cg) { return [pg, cg]; });
      })
      : Promise.resolve(null);
    return gg.then(function (geos) {
    var sheet = ctx.workbook.worksheets.getItem(sheetName);
    var span = sheet.getRange(b.descCol + b.firstRow + ":" + b.descCol + b.lastRow);
    span.load("values");
    return ctx.sync().then(function () {
      var slot = -1, i, key = childTag.toLowerCase();
      for (i = 0; i < span.values.length && slot < 0; i++) {
        if (String(span.values[i][0] || "").trim().toLowerCase() === key) { slot = b.firstRow + i; }
      }
      for (i = 0; i < span.values.length && slot < 0; i++) {
        if (String(span.values[i][0] || "").trim() === "") { slot = b.firstRow + i; }
      }
      if (slot < 0) {
        throw new Error("All " + (b.lastRow - b.firstRow + 1) + " feed-through slots on \"" +
          sheetName + "\" are in use. Remove an entry (Feed-Through, type -TAG) first.");
      }
      var warn = [];
      return SAZ.withUnprotected(ctx, sheet, function () {
        sheet.getRange(b.descCol + slot).values  = [[childTag]];
        sheet.getRange(b.kvaCol  + slot).formulas = [[demandRef]];
        if (geos) { warn = SAZ.feedThruDetail(sheet, slot, geos[0], geos[1], childSheet); }
        return ctx.sync();
      }).then(function () { return { row: slot, block: b, warn: warn }; });
    });
    });
  });
};

/* Clear every managed feed-through slot on sheetName that carries childTag.
   Resolves to the number of rows cleared (0 = not found). */
SAZ.feedThruRemove = function (ctx, sheetName, childTag) {
  return SAZ.feedThruBlock(ctx, sheetName).then(function (b) {
    var sheet = ctx.workbook.worksheets.getItem(sheetName);
    var span = sheet.getRange(b.descCol + b.firstRow + ":" + b.descCol + b.lastRow);
    span.load("values");
    return ctx.sync().then(function () {
      var rows = [], key = childTag.toLowerCase();
      for (var i = 0; i < span.values.length; i++) {
        if (String(span.values[i][0] || "").trim().toLowerCase() === key) { rows.push(b.firstRow + i); }
      }
      if (!rows.length) { return 0; }
      return SAZ.geo(ctx, sheetName).catch(function () { return {}; }).then(function (pg) {
        var dc = SAZ.geoUsable(pg.FT_DETAIL_COL);
        var dEnd = dc ? SAZ.gridColLet(SAZ.gridColNum(dc) + 4) : "";
        return SAZ.withUnprotected(ctx, sheet, function () {
          rows.forEach(function (r) {
            sheet.getRange(b.descCol + r + ":" + b.kvaCol + r).clear(Excel.ClearApplyTo.contents);
            if (dc) { sheet.getRange(dc + r + ":" + dEnd + r).clear(Excel.ClearApplyTo.contents); }
          });
          return ctx.sync();
        }).then(function () { return rows.length; });
      });
    });
  });
};

/* Read the child's current Fed From (registry column D). "" if unregistered. */
SAZ.registryFedFrom = function (ctx, childTag) {
  return SAZ.registryRange(ctx).then(function (rg) {
    rg.load("values");
    return ctx.sync().then(function () {
      for (var i = 0; i < rg.values.length; i++) {
        if (String(rg.values[i][0] || "").toLowerCase() === childTag.toLowerCase()) {
          return String(rg.values[i][3] || "").trim();
        }
      }
      return "";
    });
  });
};

/* Stamp the parent tag into the child's registry row (column D, Fed From). */
SAZ.updateRegistryFedFrom = function (ctx, childTag, parentTag) {
  return SAZ.registryRange(ctx).then(function (rg) {
    rg.load("values,rowIndex,columnIndex,worksheet/name");
    return ctx.sync().then(function () {
      for (var i = 0; i < rg.values.length; i++) {
        if (String(rg.values[i][0] || "").toLowerCase() === childTag.toLowerCase()) {
          var sheet = ctx.workbook.worksheets.getItem(rg.worksheet.name);
          var at = i;
          return SAZ.withUnprotected(ctx, sheet, function () {
            sheet.getRangeByIndexes(rg.rowIndex + at, rg.columnIndex + 3, 1, 1)
              .values = [[parentTag]];
            return ctx.sync().then(function () { return true; });
          });
        }
      }
      return false;   /* child not registered - caller decides whether that matters */
    });
  });
};

/* Run FROM the primary board. Prompts for the downstream board whose load
   feeds through to it. Demand is a LIVE LINK to the child's GEO DEMAND_CELL,
   so the pass-through tracks the child as its circuits change.
   Type -TAG to REMOVE a feed-through. When a board is pointed at a new
   parent, its row on the OLD parent (registry Fed From) is cleared first so
   its load is never counted twice. */
SAZ.uiFeedThru = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    var primary = active.name;
    return SAZ.prompt("Feed-Through",
      "Downstream board whose load feeds THROUGH to " + primary +
      "  (type -TAG to remove one)", "", true)
      .then(function (ans) {
        if (!ans.ok) { return null; }
        var raw = String(ans.value || "").trim();
        var removing = raw.charAt(0) === "-";
        var childTag = SAZ.sanitise(removing ? raw.slice(1) : raw);
        if (!childTag) { return SAZ.notify("Invalid tag", "No usable characters in that tag."); }
        if (removing) {
          return SAZ.feedThruRemove(ctx, primary, childTag).then(function (n) {
            if (!n) {
              return SAZ.notify("Not found", "\"" + childTag + "\" is not a feed-through on " + primary + ".");
            }
            return SAZ.registryFedFrom(ctx, childTag).then(function (ff) {
              var clr = ff.toLowerCase() === primary.toLowerCase()
                ? SAZ.updateRegistryFedFrom(ctx, childTag, "") : Promise.resolve(false);
              return clr.then(function () {
                return SAZ.notify("Feed-through removed", childTag + " removed from " + primary + ".");
              });
            });
          });
        }
        if (childTag.toLowerCase() === primary.toLowerCase()) {
          return SAZ.notify("Cannot feed itself",
            "\"" + primary + "\" cannot be its own feed-through source.");
        }
        var sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        return ctx.sync().then(function () {
          var find = function (nm) {
            return sheets.items.filter(function (s) { return s.name.toLowerCase() === String(nm).toLowerCase(); })[0];
          };
          var child = find(childTag);
          if (!child) {
            return SAZ.notify("No such board",
              "There is no sheet named \"" + childTag + "\". Create it first.");
          }
          return SAZ.geo(ctx, child.name).then(function (cg) {
            if (!cg.DEMAND_CELL) {
              throw new Error("\"" + child.name + "\" has no DEMAND_CELL in its GEO_MAP.");
            }
            var ref = SAZ.sheetRef(child.name, cg.DEMAND_CELL);
            return SAZ.registryFedFrom(ctx, child.name).then(function (oldParent) {
              var old = find(oldParent);
              var purge = (old && old.name.toLowerCase() !== primary.toLowerCase())
                ? SAZ.feedThruRemove(ctx, old.name, child.name).catch(function () { return 0; })
                : Promise.resolve(0);
              return purge.then(function (moved) {
                return SAZ.feedThruAppend(ctx, primary, child.name, ref, child.name).then(function (res) {
                  return SAZ.updateRegistryFedFrom(ctx, child.name, primary).then(function (reg) {
                    return SAZ.notify("Feed-through added",
                      child.name + " added to " + primary + " at row " + res.row +
                      " (no breaker used)." +
                      (moved ? "\n\nRemoved from its previous parent " + old.name + " so it is not counted twice." : "") +
                      (reg ? "" : "\n\nNote: " + child.name +
                       " is not in the equipment registry, so Fed From was not updated.") +
                      (res.warn && res.warn.length ? "\n\nOnly demand was linked for: " + res.warn.join(", ") +
                       ". Recreate that board from the current template to carry connected, phase and receptacle loads." : ""));
                  });
                });
              });
            });
          });
        });
      });
  });
};


/* ==========================================================
   DISTRIBUTION TREE  -  Link Equipment / Sub-Feed
   Port of VBA LinkEquipment (1022-1123), AddParentCircuit (1132-1182),
   AddParentCircuitAt (1599-1616), TxAppendRow (1441-1462).

   Link and Sub-Feed differ ONLY in slot choice: Link takes the first open
   circuit, Sub-Feed takes the row you selected. They share one implementation
   here, which is what the VBA comment at 1153 said it wanted but did not do.

   A '-' in a GEO_MAP column key means 'this template has no such column' and
   MUST be skipped, not written to. MSB for instance has CKT_QTY_COL = '-'.
   ========================================================== */

SAZ.TX_SHEET = "Transformers";
/* Default transformer impedance, %Z. A business assumption, not a constant -
   the VBA buried 5.75 inline at line 1095. It now lives in a workbook input
   cell. The workbook already had one: Xfmr_DefZ on NEC_Data, which the sheet
   formulas use too, so the code reads that rather than a second copy. */
/* %Z is no longer held in code. The Transformers sheet defaults it from the
   Xfmr_DefZ name and column T overrides per unit, so a constant here would be
   a third home for one business assumption - the drift hazard already retired
   twice on this project (Cfg_TxZPct, and the second hosted template file). */
/* Voltage equality tolerance in volts - below this, no transformer is needed. */
SAZ.VOLT_MATCH_TOLERANCE = 1;

SAZ.geoUsable = function (val) {
  var s = String(val == null ? "" : val).trim();
  return (s !== "" && s !== "-") ? s : "";
};

/* Value of the cell a GEO key POINTS AT (not the key's own value). */
SAZ.geoCellValue = function (ctx, sheetName, key) {
  return SAZ.geo(ctx, sheetName).then(function (g) {
    var addr = SAZ.geoUsable(g[key]);
    if (!addr) { return null; }
    var cell = ctx.workbook.worksheets.getItem(sheetName).getRange(addr);
    cell.load("values");
    return ctx.sync().then(function () {
      var v = cell.values[0][0];
      return (typeof v === "number") ? v : (parseFloat(v) || 0);
    });
  });
};

/* First circuit row with an empty description. */
SAZ.addParentCircuit = function (ctx, parentTag, descTag, kvaRef) {
  return SAZ.geo(ctx, parentTag).then(function (g) {
    var first = Number(g.CKT_FIRST) || 0;
    var lastR = Number(g.CKT_LAST) || 0;
    var dc = SAZ.geoUsable(g.CKT_DESC_COL);
    if (!first || lastR < first || !dc) {
      throw new Error("GEO_MAP on \"" + parentTag + "\" is missing circuit-grid keys" +
        " (CKT_FIRST, CKT_LAST, CKT_DESC_COL). Run Rebuild GEO_MAP.");
    }
    var span = ctx.workbook.worksheets.getItem(parentTag)
      .getRange(dc + first + ":" + dc + lastR);
    span.load("values");
    return ctx.sync().then(function () {
      for (var i = 0; i < span.values.length; i++) {
        if (String(span.values[i][0] || "").trim() === "") {
          return SAZ.addParentCircuitAt(ctx, parentTag, descTag, kvaRef, first + i);
        }
      }
      throw new Error("No open circuit on \"" + parentTag + "\" between rows " +
        first + " and " + lastR + ". Resize the panel first.");
    });
  });
};


/* Row on the Transformers schedule whose Tag matches, or 0 if absent.

   Replaces txAppendRow. That helper looked for the first BLANK row in column
   A, but every row of column A is a formula returning "" while the detection
   spill is shorter than the grid - so it always returned the first data row,
   and the caller then overwrote that row's derived formulas. Rows on this
   sheet are never appended; they appear on their own and are found by tag. */
SAZ.txRowForTag = function (ctx, tag) {
  var sheet = ctx.workbook.worksheets.getItem(SAZ.TX_SHEET);
  var col = sheet.getRange("A1:A400");
  col.load("values");
  return ctx.sync().then(function () {
    var v = col.values, want = String(tag).trim().toLowerCase();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0] || "").trim().toLowerCase() === want) { return i + 1; }
    }
    return 0;
  });
};

/* Link a child board to a parent. targetRow optional: omit for the first open
   circuit (Link Equipment), pass a row for a chosen slot (Sub-Feed).

   If parent and child line voltages differ by more than the tolerance, a
   transformer row is created on the Transformers schedule and the child's
   source fault is fed from the transformer's secondary rather than the parent
   bus. Every cross-sheet value is a LIVE LINK, never a copied number. */
/* ==========================================================
   v2.3 LINK EQUIPMENT  -  circuit-position aware
   The child is placed in a real circuit position on the parent: the
   description, pole count and a LIVE demand link land on the chosen row, and
   multi-pole breakers tie the rows below through the template's own AN
   continuation logic (writing the pole count on the first row is enough).
   The slot dialog defaults to the first open position that fits the pole
   count and refuses an occupied one - occupied means a description, any
   connected load, or a row already tied into a multi-pole breaker above.
   Every write to a schedule runs inside withUnprotected: the child's Source
   Fault cell is a LOCKED formula, and writing it on a protected sheet was why
   linking failed part-way before v2.3.
   ========================================================== */

/* Circuit grid of a board: one record per row with its circuit number,
   description, poles, connected load and the row that owns it if it is tied
   into a multi-pole breaker. TWOBLOCK panelboards are two independent blocks
   (odd circuits, even circuits) - a breaker may never span across them. */
/* v2.12: the board tag a circuit description identifies. Linked sub-feeds now carry a
   readable label - "PANEL L1" or "PANEL L1 (VIA 75 KVA XFMR T-L1)" - so every identity
   match goes through here instead of comparing the raw text to the tag. */
SAZ.BOARD_WORD = { panelboard: "PANEL", switchboard: "SWBD", mcc: "MCC" };
SAZ.descKey = function (d) {
  var s = String(d == null ? "" : d).trim();
  var m = /\(\s*VIA\b[^)]*?\b(T-[^\s)]+)\s*\)\s*$/i.exec(s);
  if (m) { return m[1]; }
  m = /^(PANEL|SWBD|MCC)\s+(\S+)$/i.exec(s);
  return m ? m[2] : s;
};
SAZ.descIs = function (d, tags) { var k = SAZ.descKey(d).toLowerCase(), raw = String(d || "").trim().toLowerCase(); return tags.some(function (t) { t = String(t).toLowerCase(); return t === k || t === raw; }); };
/* Live label formula for a child board on its parent's circuit row. Reads the child's own
   NAME_CELL (follows a rename) and the transformer's Std kVA (follows a re-size). */
SAZ.linkDescFormula = function (childTag, childType, nameCell, viaTx) {
  var w = SAZ.BOARD_WORD[String(childType || "").trim().toLowerCase()] || "PANEL";
  var nm = SAZ.sheetRef(childTag, nameCell).substring(1);
  var word = new RegExp("^" + w + "\\b", "i").test(childTag) ? "" : "\"" + w + " \"&";
  if (!viaTx) { return "=" + word + nm; }
  var tx = "'" + SAZ.TX_SHEET.replace(/'/g, "''") + "'!";
  var kva = "IFERROR(TEXT(XLOOKUP(\"T-\"&" + nm + "," + tx + "$A$1:$A$400," + tx + "$G$1:$G$400),\"0\")&\" KVA \",\"\")";
  return "=" + word + nm + "&\" (VIA \"&" + kva + "&\"XFMR T-\"&" + nm + "&\")\"";
};

SAZ.cktGrid = function (ctx, sheetName) {
  return SAZ.geo(ctx, sheetName).then(function (g) {
    var first = Number(g.CKT_FIRST) || 0, last = Number(g.CKT_LAST) || 0;
    var dc = SAZ.geoUsable(g.CKT_DESC_COL), nc = SAZ.geoUsable(g.CKT_NUM_COL) || "B";
    var pc = SAZ.geoUsable(g.CKT_POLES_COL), kc = SAZ.geoUsable(g.CKT_KVA_COL);
    if (!first || last < first || !dc) {
      throw new Error("GEO_MAP on \"" + sheetName + "\" is missing circuit-grid keys" +
        " (CKT_FIRST, CKT_LAST, CKT_DESC_COL). Run Rebuild GEO_MAP.");
    }
    var blocks = [{ f: first, l: last }];
    /* Only panelboards tie the rows below a multi-pole breaker (column AN
       continuation logic). Switchboard and MCC rows are one device each,
       whatever the pole count, so there a breaker always fills one row. */
    var spans = String(g.RESIZE_MODE || "").trim().toUpperCase() === "TWOBLOCK";
    if (spans) {
      var b1F = Number(g.CKT_BLOCK1_FIRST) || first, b1L = Number(g.CKT_BLOCK1_LAST) || last;
      var b2F = Number(g.CKT_BLOCK2_FIRST) || 0,     b2L = Number(g.CKT_BLOCK2_LAST) || last;
      blocks = b2F ? [{ f: b1F, l: b1L }, { f: b2F, l: b2L }] : [{ f: b1F, l: b1L }];
    }
    var sh = ctx.workbook.worksheets.getItem(sheetName);
    var hdr = sh.getRange("A" + (first - 1) + ":CZ" + (first - 1));
    var rN = sh.getRange(nc + first + ":" + nc + last);
    var rD = sh.getRange(dc + first + ":" + dc + last);
    var rP = pc ? sh.getRange(pc + first + ":" + pc + last) : null;
    hdr.load("values"); rN.load("values"); rD.load("values"); if (rP) { rP.load("values"); }
    return ctx.sync().then(function () {
      var hc = -1;
      hdr.values[0].forEach(function (v, j) {
        if (hc < 0 && String(v || "").trim().toLowerCase() === "conn. kva") { hc = j; }
      });
      var rH = null;
      if (hc >= 0) {
        rH = sh.getRangeByIndexes(first - 1, hc, last - first + 1, 1);
        rH.load("values");
      }
      return ctx.sync().then(function () {
        var rows = [], byNum = {};
        for (var i = 0; i <= last - first; i++) {
          var p = rP ? parseInt(rP.values[i][0], 10) : 1;
          rows.push({
            row: first + i,
            num: Number(rN.values[i][0]) || 0,
            desc: String(rD.values[i][0] == null ? "" : rD.values[i][0]).trim(),
            poles: (p >= 1 && p <= 3) ? p : 1,
            conn: rH ? String(rH.values[i][0] == null ? "" : rH.values[i][0]).trim() : "",
            owner: 0
          });
        }
        /* multi-pole ownership, walked block by block so nothing crosses a block edge */
        if (spans) blocks.forEach(function (b) {
          var left = 0, own = 0;
          for (var r = b.f; r <= b.l; r++) {
            var rec = rows[r - first];
            if (!rec) { continue; }
            if (left > 0) { rec.owner = own; left--; continue; }
            if (rec.poles > 1 && (rec.desc || rec.conn)) { own = r; left = rec.poles - 1; }
          }
        });
        rows.forEach(function (rec) { if (rec.num) { byNum[rec.num] = rec; } });
        return { name: sheetName, g: g, first: first, last: last, blocks: blocks,
                 rows: rows, byNum: byNum, spans: spans, cols: { dc: dc, pc: pc, kc: kc } };
      });
    });
  });
};

/* Is circuit 'num' (and the rows a 'poles' breaker would tie below it) free?
   'mine' lists descriptions that count as free - the child's own current
   circuit, so re-linking to the same slot is not refused. */
SAZ.slotCheck = function (grid, num, poles, mine) {
  var start = grid.byNum[num];
  var nums = Object.keys(grid.byNum).map(Number).sort(function (a, b) { return a - b; });
  if (!start) {
    return { ok: false, msg: "Circuit " + num + " does not exist on " + grid.name +
      " (circuits " + nums[0] + "-" + nums[nums.length - 1] + ")." };
  }
  var blk = grid.blocks.filter(function (b) { return start.row >= b.f && start.row <= b.l; })[0];
  var span = grid.spans ? poles : 1;
  if (!blk || start.row + span - 1 > blk.l) {
    return { ok: false, msg: "A " + poles + "-pole breaker at circuit " + num +
      " would run past the bottom of the schedule. Pick a lower circuit or resize " + grid.name + "." };
  }
  var mineL = (mine || []).map(function (s) { return String(s).toLowerCase(); });
  var isMine = function (rec) { return rec.desc && SAZ.descIs(rec.desc, mineL);   /* v2.12 */ };
  var used = [], rowsOut = [], numsOut = [];
  for (var k = 0; k < span; k++) {
    var rec = grid.rows[start.row + k - grid.first];
    rowsOut.push(rec.row); numsOut.push(rec.num);
    if (isMine(rec)) { continue; }
    if (rec.owner) {
      var o = grid.rows[rec.owner - grid.first];
      if (!isMine(o)) { used.push(rec.num + " (tied to ckt " + o.num + (o.desc ? " \"" + o.desc + "\"" : "") + ")"); }
    } else if (rec.desc || rec.conn) {
      used.push(rec.num + (rec.desc ? " \"" + rec.desc + "\"" : " (has load)"));
    }
  }
  if (used.length) {
    return { ok: false, msg: "Occupied on " + grid.name + ": circuit " + used.join(", ") + ". Pick another position." };
  }
  return { ok: true, row: start.row, poles: poles, nums: numsOut, label: numsOut.join(",") };
};

/* Lowest circuit number where a 'poles' breaker fits, or 0. */
SAZ.firstFit = function (grid, poles, mine) {
  var nums = Object.keys(grid.byNum).map(Number).sort(function (a, b) { return a - b; });
  for (var i = 0; i < nums.length; i++) {
    if (SAZ.slotCheck(grid, nums[i], poles, mine).ok) { return nums[i]; }
  }
  return 0;
};

/* Default breaker poles for a child from its phase count and line voltage. */
SAZ.defaultPoles = function (ph, vll) {
  if (Number(ph) === 1) { return (Number(vll) > 0 && Number(vll) <= 120) ? 1 : 2; }
  return 3;
};

/* Write one circuit on the parent at a chosen row. Column mappings live here only. */
SAZ.addParentCircuitAt = function (ctx, parentTag, descTag, kvaRef, targetRow, poles) {
  return SAZ.geo(ctx, parentTag).then(function (g) {
    var sheet = ctx.workbook.worksheets.getItem(parentTag);
    return SAZ.withUnprotected(ctx, sheet, function () {
      var put = function (key, val, isFormula) {
        var col = SAZ.geoUsable(g[key]);
        if (!col) { return; }
        var rg = sheet.getRange(col + targetRow);
        if (isFormula) { rg.formulas = [[val]]; } else { rg.values = [[val]]; }
      };
      put("CKT_DESC_COL",  descTag, String(descTag).charAt(0) === "=");   /* v2.12: live label formula allowed */
      put("CKT_KVA_COL",   kvaRef,  true);
      put("CKT_QTY_COL",   1,       false);
      put("CKT_POLES_COL", poles || 3, false);
      put("CKT_CLASS_COL", "Panel", false);
      return ctx.sync().then(function () { return targetRow; });
    });
  });
};

/* Remove a child's existing circuit(s) from a parent before it is re-linked,
   so a move never leaves a stale feeder behind. Blanks by writing "" (the
   description cells are merged; Range.clear() throws on them). */
SAZ.clearParentCircuit = function (ctx, parentTag, descs) {
  var want = descs.map(function (s) { return String(s).toLowerCase(); });
  return SAZ.cktGrid(ctx, parentTag).then(function (grid) {
    var hits = grid.rows.filter(function (r) { return r.desc && SAZ.descIs(r.desc, want);   /* v2.12 */ });
    if (!hits.length) { return 0; }
    var sheet = ctx.workbook.worksheets.getItem(parentTag);
    return SAZ.withUnprotected(ctx, sheet, function () {
      hits.forEach(function (h) {
        sheet.getRange(grid.cols.dc + h.row).values = [[""]];
        if (grid.cols.kc) { sheet.getRange(grid.cols.kc + h.row).values = [[""]]; }
        if (grid.cols.pc) { sheet.getRange(grid.cols.pc + h.row).values = [[""]]; }
      });
      return ctx.sync().then(function () { return hits.length; });
    });
  });
};

/* Registry rows as simple records. */
SAZ.registryRecords = function (ctx) {
  return SAZ.registryRange(ctx).then(function (rg) {
    rg.load("values");
    return ctx.sync().then(function () {
      var out = [];
      rg.values.forEach(function (r, i) {
        var t = String(r[0] == null ? "" : r[0]).trim();
        if (t) { out.push({ tag: t, type: String(r[1] || ""), fed: String(r[3] || "").trim(),
                            ft: String(r[8] || "").trim(), ckt: String(r[9] || "").trim(), at: i }); }
      });
      return out;
    });
  });
};

/* Stamp Fed From (D), Feed Type (I, only if blank) and Feeder Ckt / Space (J). */
SAZ.updateRegistryLink = function (ctx, childTag, parentTag, cktLabel) {
  return SAZ.registryRange(ctx).then(function (rg) {
    rg.load("values,rowIndex,columnIndex,worksheet/name");
    return ctx.sync().then(function () {
      for (var i = 0; i < rg.values.length; i++) {
        if (String(rg.values[i][0] || "").toLowerCase() === childTag.toLowerCase()) {
          var sheet = ctx.workbook.worksheets.getItem(rg.worksheet.name);
          var at = i, ftBlank = String(rg.values[i][8] || "").trim() === "";
          return SAZ.withUnprotected(ctx, sheet, function () {
            sheet.getRangeByIndexes(rg.rowIndex + at, rg.columnIndex + 3, 1, 1).values = [[parentTag]];
            if (ftBlank) { sheet.getRangeByIndexes(rg.rowIndex + at, rg.columnIndex + 8, 1, 1).values = [["Sub-Feed"]]; }
            sheet.getRangeByIndexes(rg.rowIndex + at, rg.columnIndex + 9, 1, 1).values = [[cktLabel || ""]];
            return ctx.sync().then(function () { return true; });
          });
        }
      }
      return false;
    });
  });
};

/* The child's Fed From cell: GEO key FEDFROM_CELL if a template defines one,
   otherwise the cell directly below NAME_CELL (D5 -> D6 on all three templates). */
SAZ.fedFromCell = function (g) {
  var k = SAZ.geoUsable(g.FEDFROM_CELL);
  if (k) { return k; }
  var m = /^([A-Z]+)([0-9]+)$/i.exec(SAZ.geoUsable(g.NAME_CELL));
  return m ? m[1] + (Number(m[2]) + 1) : "";
};

/* Link a child board to a parent at a slot { row, poles, label }.
   slot may also be a bare row number (legacy callers) or omitted (first fit).
   If parent and child line voltages differ by more than the tolerance, the
   Transformers schedule derives a T-<child> row from the registry on its own;
   this command reveals that schedule if it is put away, finds the row by tag
   and points the child's source fault at its secondary. Every cross-sheet
   value is a LIVE LINK, never a copied number. */
SAZ.linkEquipment = function (ctx, parentTag, childTag, slot, opts) {
  var say = function (title, text) { return (opts && opts.quiet) ? Promise.resolve({ title: title, text: text }) : SAZ.notify(title, text); };
  if (parentTag.toLowerCase() === childTag.toLowerCase()) {
    return say("Cannot feed itself", "\"" + parentTag + "\" cannot feed itself.");
  }
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name,items/visibility");
  return ctx.sync().then(function () {
    var byName = function (t) {
      return sheets.items.filter(function (s) { return s.name.toLowerCase() === t.toLowerCase(); })[0];
    };
    var p = byName(parentTag), c = byName(childTag);
    if (!p) { return say("Parent not found", "No sheet named \"" + parentTag + "\"."); }
    if (!c) { return say("Child not found",  "No sheet named \"" + childTag + "\"."); }
    var P = p.name, C = c.name;
    var gs = [];
    return SAZ.geo(ctx, P).then(function (x) { gs[0] = x; return SAZ.geo(ctx, C); })
    .then(function (x) { gs[1] = x; return SAZ.registryRecords(ctx); })
    .then(function (x) {
      gs[2] = x;
      var gp = gs[0], gc = gs[1], reg = gs[2];
      var childDemand = SAZ.geoUsable(gc.DEMAND_CELL);
      var childSrc    = SAZ.geoUsable(gc.SRCFAULT_CELL);
      var parentBus   = SAZ.geoUsable(gp.BUSFAULT_CELL);
      var fedCell     = SAZ.fedFromCell(gc);
      if (!childDemand || !childSrc || !parentBus) {
        throw new Error("GEO_MAP incomplete: need DEMAND_CELL and SRCFAULT_CELL on \"" +
          C + "\" and BUSFAULT_CELL on \"" + P + "\".");
      }
      var demandRef = SAZ.sheetRef(C, childDemand);
      var mine = [C, "T-" + C];
      var regC = reg.filter(function (r) { return r.tag.toLowerCase() === C.toLowerCase(); })[0];
      var oldP = regC ? regC.fed : "";
      if (oldP === "(unassigned)") { oldP = ""; }
      var vals = [];
      return SAZ.geoCellValue(ctx, P, "ENG_VLL_CELL").then(function (x) { vals[0] = x; return SAZ.geoCellValue(ctx, C, "ENG_VLL_CELL"); })
      .then(function (x) { vals[1] = x; return SAZ.geoCellValue(ctx, C, "ENG_PH_CELL"); })
      .then(function (x) {
        vals[2] = x;
        var pV = vals[0] || 0, cV = vals[1] || 0, cPh = vals[2] || 0;
        if (!pV || !cV) {
          throw new Error("Set the System Voltage on " + (!pV ? "\"" + P + "\"" : "\"" + C + "\"") +
            " first - the link needs both voltages to decide whether a transformer is required.");
        }
        var needsTx = Math.abs(pV - cV) > SAZ.VOLT_MATCH_TOLERANCE;
        var desc = SAZ.linkDescFormula(C, regC ? regC.type : "", SAZ.geoUsable(gc.NAME_CELL) || "D5", needsTx);   /* v2.12: "PANEL L1 (VIA 75 KVA XFMR T-L1)" */
        var childSheet = ctx.workbook.worksheets.getItem(C);

        /* resolve the slot: explicit, legacy row, or first fit */
        var resolveSlot = SAZ.cktGrid(ctx, P).then(function (grid) {
          var poles = (slot && slot.poles) || SAZ.defaultPoles(cPh, cV);
          if (slot && slot.row) {
            var rec = grid.rows[slot.row - grid.first];
            var chk = rec ? SAZ.slotCheck(grid, rec.num, poles, mine) : { ok: false, msg: "Row " + slot.row + " is not a circuit row." };
            if (!chk.ok) { throw new Error(chk.msg); }
            return chk;
          }
          if (typeof slot === "number") {
            var rec2 = grid.rows[slot - grid.first];
            var chk2 = rec2 ? SAZ.slotCheck(grid, rec2.num, poles, mine) : { ok: false, msg: "Row " + slot + " is not a circuit row." };
            if (!chk2.ok) { throw new Error(chk2.msg); }
            return chk2;
          }
          var n = SAZ.firstFit(grid, poles, mine);
          if (!n) { throw new Error("No open " + poles + "-pole position on \"" + P + "\". Resize the panel first."); }
          return SAZ.slotCheck(grid, n, poles, mine);
        });

        return resolveSlot.then(function (s) {
          /* free the child's previous circuit wherever it was (old parent, or the same parent) */
          var clearOld = oldP ? (byName(oldP) ? SAZ.clearParentCircuit(ctx, byName(oldP).name, mine) : Promise.resolve(0))
                              : SAZ.clearParentCircuit(ctx, P, mine);
          return clearOld.then(function (cleared) {
            var txReady = Promise.resolve(0);
            if (needsTx) {
              var tx = byName(SAZ.TX_SHEET);
              if (!tx) {
                throw new Error("\"" + P + "\" is " + pV + "V and \"" + C + "\" is " + cV +
                  "V, so a transformer is required, but this workbook has no " + SAZ.TX_SHEET +
                  " schedule to put it on. Run New Project to bring in the full scaffold.");
              }
              if (tx.visibility !== Excel.SheetVisibility.visible) {
                tx.visibility = Excel.SheetVisibility.visible;   /* deploy the schedule; do not move the user */
              }
              txReady = SAZ.updateRegistryLink(ctx, C, P, s.label).then(function () {
                ctx.workbook.application.calculate(Excel.CalculationType.full);
                return ctx.sync();
              }).then(function () { return SAZ.txRowForTag(ctx, "T-" + C); }).then(function (tr) {
                if (!tr) {
                  throw new Error("\"" + C + "\" is " + cV + "V under \"" + P + "\" at " + pV +
                    "V, but no transformer row appeared on " + SAZ.TX_SHEET +
                    ". The usual cause is a registry voltage string the config list does " +
                    "not recognise - see the check at " + SAZ.TX_SHEET + "!BS2.");
                }
                return tr;
              });
            }
            return txReady.then(function (tr) {
              var srcF = needsTx
                ? "='" + SAZ.TX_SHEET.replace(/'/g, "''") + "'!$T$" + tr    /* T = Sec Fault Sym (A) */
                : SAZ.sheetRef(P, parentBus);
              return SAZ.withUnprotected(ctx, childSheet, function () {
                childSheet.getRange(childSrc).formulas = [[ srcF ]];
          /* v2.5: Fed From is a formula on the schedule (reads the registry) - nothing to write */
                return ctx.sync();
              }).then(function () {
                return SAZ.addParentCircuitAt(ctx, P, desc, demandRef, s.row, s.poles);
              }).then(function () {
                return needsTx ? true : SAZ.updateRegistryLink(ctx, C, P, s.label);
              }).then(function () {
                var moved = cleared ? "\n\nPrevious circuit on " + (oldP || P) + " cleared." : "";
                if (!needsTx) {
                  return say("Linked",
                    C + " linked to " + P + " at circuit " + s.label + " (" + s.poles + "-pole), same voltage." + moved);
                }
                return say("Linked via transformer",
                  C + " (" + cV + "V) linked to " + P + " (" + pV + "V) at circuit " + s.label +
                  " (" + s.poles + "-pole).\n\nTransformer T-" + C + " was added to the " + SAZ.TX_SHEET +
                  " schedule (row " + tr + ") and the Distribution Map. Set %Z on that row if the Xfmr_DefZ " +
                  "default is not right for this unit." + moved);
              });
            });
          });
        });
      });
    });
  });
};

/* Slot dialog: circuit number + poles, pre-filled with the first open
   position. Re-opens with the reason when the chosen slot is occupied. */
SAZ.uiPickSlot = function (ctx, P, C, wantNum, wantPoles, err) {
  var r = [];
  return SAZ.cktGrid(ctx, P).then(function (g) { r[0] = g; return SAZ.geoCellValue(ctx, C, "ENG_PH_CELL"); })
  .then(function (v) { r[1] = v; return SAZ.geoCellValue(ctx, C, "ENG_VLL_CELL"); })
  .then(function (v) { r[2] = v;
    var grid = r[0], mine = [C, "T-" + C];
    var poles = Number(wantPoles) || SAZ.defaultPoles(r[1], r[2]);
    var nums = Object.keys(grid.byNum).map(Number).sort(function (a, b) { return a - b; });
    var fit = SAZ.firstFit(grid, poles, mine);
    var num = Number(wantNum) || fit || nums[0];
    return SAZ.form("Link " + C + " \u2192 " + P,
      C + " will be fed from " + P + " (" + nums.length + " circuits).",
      [ { id: "ckt", label: "Circuit position on " + P, type: "number", value: String(num),
          min: nums[0], max: nums[nums.length - 1], required: true,
          hint: fit ? "First open slot for a " + poles + "-pole breaker: " + fit + "." : "No open " + poles + "-pole slot - resize " + P + " first." },
        { id: "poles", label: "Breaker poles", type: "select", options: ["1", "2", "3"], value: String(poles), required: true } ],
      err || "", 62).then(function (ans) {
        if (!ans.ok) { return null; }
        var n = parseInt(ans.values.ckt, 10), pl = parseInt(ans.values.poles, 10) || poles;
        var chk = SAZ.slotCheck(grid, n, pl, mine);
        if (!chk.ok) { return SAZ.uiPickSlot(ctx, P, C, n, pl, chk.msg); }
        return SAZ.linkEquipment(ctx, P, C, chk);
      });
  });
};

/* Link Equipment: pick parent and child (type-ahead from the registry), then
   the circuit position. The active sheet pre-fills the parent when it is a board. */
SAZ.uiLink = function (ctx, presetParent, presetNum, err) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () { return SAZ.registryRecords(ctx); }).then(function (reg) {
    var tags = reg.filter(function (x) { return SAZ.isBoardType(x.type); }).map(function (x) { return x.tag; }).slice(0, 150);
    var isBoard = function (t) { return tags.some(function (x) { return x.toLowerCase() === String(t).toLowerCase(); }); };
    var pDefault = presetParent || (isBoard(active.name) ? active.name : "");
    return SAZ.form("Link Equipment", "Pick the upstream board and the board it feeds.",
      [ { id: "parent", label: "PARENT (upstream board)", type: "text", value: pDefault, list: tags, required: true },
        { id: "child",  label: "CHILD (fed from the parent)", type: "text", value: "", list: tags, required: true } ],
      err || "", 58).then(function (ans) {
        if (!ans.ok) { return null; }
        var P = SAZ.sanitise(ans.values.parent), C = SAZ.sanitise(ans.values.child);
        if (!P || !C) { return SAZ.uiLink(ctx, P, presetNum, "Both tags are required."); }
        if (P.toLowerCase() === C.toLowerCase()) { return SAZ.uiLink(ctx, P, presetNum, "A board cannot feed itself."); }
        if (!isBoard(P)) { return SAZ.uiLink(ctx, "", presetNum, "\"" + P + "\" is not in the equipment registry."); }
        if (!isBoard(C)) { return SAZ.uiLink(ctx, P, presetNum, "\"" + C + "\" is not in the equipment registry."); }
        var canon = function (t) { return tags.filter(function (x) { return x.toLowerCase() === t.toLowerCase(); })[0]; };
        return SAZ.uiPickSlot(ctx, canon(P), canon(C), presetNum, 0, "");
      });
  });
};

/* Sub-Feed: run FROM the parent with a circuit row selected - same dialogs,
   pre-filled with this board and the selected circuit number. */
SAZ.uiSubFeed = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  var sel = ctx.workbook.getSelectedRange();
  sel.load("rowIndex");
  return ctx.sync().then(function () {
    var parentTag = active.name;
    var row = sel.rowIndex + 1;
    return SAZ.cktGrid(ctx, parentTag).then(function (grid) {
      var rec = grid.rows[row - grid.first];
      return SAZ.uiLink(ctx, parentTag, rec ? rec.num : 0, "");
    }, function () {
      return SAZ.notify("Not an equipment sheet",
        "\"" + parentTag + "\" has no circuit grid in its GEO_MAP." +
        " Run Sub-Feed from a panelboard, switchboard or MCC.");
    });
  });
};


/* ==========================================================
   GEO_MAP SELF-HEAL  -  port of RebuildGeoMap (2326-2371),
   GeoDeriveRows (2374-2408) and SafeWrite (2421-2433).

   Each GEO_MAP row is six columns wide:
     +0 key   +1 value   +2 ANCHOR_TEXT   +3 ANCHOR_COL   +4 TARGET_COL   +5 ROW_OFFSET
   The value is REDERIVED by finding ANCHOR_TEXT inside ANCHOR_COL and
   composing TARGET_COL & (foundRow + ROW_OFFSET). That is what makes the map
   survive inserted rows - the anchor is a label the user can see, not an
   address. Rows with no anchor text are left alone.
   ========================================================== */

SAZ.GEO_WIDTH = 6;

/* Write that survives a protected sheet, then restores protection. */
SAZ.safeWrite = function (ctx, sheetName, addr, value) {
  var sheet = ctx.workbook.worksheets.getItem(sheetName);
  return SAZ.withUnprotected(ctx, sheet, function () {
    sheet.getRange(addr).values = [[value]];
    return ctx.sync();
  });
};

/* Rebuild every anchored key on one sheet. Returns a per-key change report. */
SAZ.rebuildGeoMap = function (ctx, sheetName) {
  var sheet = ctx.workbook.worksheets.getItem(sheetName);
  var nm = sheet.names.getItemOrNullObject("GEO_MAP");
  nm.load("isNullObject");
  return ctx.sync().then(function () {
    if (nm.isNullObject) {
      throw new Error("\"" + sheetName + "\" has no GEO_MAP block. NOTHING WAS" +
        " REBUILT. The block is located by a sheet-scoped name - AS1 on the" +
        " panelboard and switchboard templates, BJ1 on the MCC.");
    }
    var blk = nm.getRange();
    blk.load("values,rowIndex,columnIndex,rowCount");
    /* the anchor search needs whole columns, so read a generous slab */
    var slab = sheet.getRangeByIndexes(0, 0, 200, 60);
    slab.load("values");
    return ctx.sync().then(function () {
      var v = blk.values, changes = [], writes = {};
      for (var i = 0; i < v.length; i++) {
        var key = String(v[i][0] || "").trim();
        if (!key || key.toUpperCase() === "GEO_MAP") { continue; }
        var aTxt = String(v[i][2] || "").trim();
        var aCol = String(v[i][3] || "").trim();
        var tCol = String(v[i][4] || "").trim();
        var off  = Number(v[i][5]) || 0;
        if (!aTxt || !aCol) { continue; }   /* unanchored - leave as authored */
        var ci = aCol.toUpperCase().charCodeAt(0) - 65;
        if (aCol.length > 1) { ci = (aCol.toUpperCase().charCodeAt(0) - 64) * 26 +
                                   (aCol.toUpperCase().charCodeAt(1) - 65); }
        var hitRow = -1;
        for (var r = 0; r < slab.values.length; r++) {
          if (String(slab.values[r][ci] || "").trim().toLowerCase() === aTxt.toLowerCase()) {
            hitRow = r + 1; break;
          }
        }
        if (hitRow < 0) { continue; }       /* anchor text gone - do not guess */
        var newVal = tCol ? (tCol + (hitRow + off)) : String(hitRow + off);
        var oldVal = String(v[i][1] == null ? "" : v[i][1]).trim();
        if (newVal !== oldVal) {
          changes.push({ key: key, from: oldVal, to: newVal });
          writes[i] = newVal;
        }
      }
      var rows = Object.keys(writes);
      if (!rows.length) { return { sheet: sheetName, changed: 0, changes: [] }; }
      return SAZ.withUnprotected(ctx, sheet, function () {
        rows.forEach(function (i) {
          sheet.getRangeByIndexes(blk.rowIndex + Number(i), blk.columnIndex + 1, 1, 1)
            .values = [[ writes[i] ]];
        });
        return ctx.sync();
      }).then(function () {
        return { sheet: sheetName, changed: changes.length, changes: changes };
      });
    });
  });
};


/* Keys that are COMPUTED from other keys rather than anchored to a label.
   Driven by BLOCK_SIZE / FT_SLOTS / VD_COL / RESIZE_MODE.

   NOTE: the current TPL_Panelboard, TPL_Switchboard and TPL_MCC GEO_MAPs carry
   19 keys and none of these drivers, so this pass is a no-op on today's
   templates. Ported anyway because ResizePanel is what sets BLOCK_SIZE, and
   Phase 3's resize work is where these keys start to matter. */
SAZ.geoDeriveRows = function (ctx, sheetName) {
  return SAZ.geo(ctx, sheetName).then(function (g) {
    var c1    = Number(g.CKT_FIRST) || 0;
    var blkSz = Number(g.BLOCK_SIZE) || 0;
    var ft1   = Number(g.FEEDTHRU_FIRST) || 0;
    var slots = Number(g.FT_SLOTS) || 0;
    var vdCol = SAZ.geoUsable(g.VD_COL);
    var twoBlock = String(g.RESIZE_MODE || "").toUpperCase() === "TWOBLOCK";
    var set = {};
    if (c1 > 0 && blkSz > 0) {
      var b1L = c1 + blkSz - 1;
      set.CKT_BLOCK1_FIRST = String(c1);
      set.CKT_BLOCK1_LAST  = String(b1L);
      if (twoBlock) {
        var b2F = b1L + 1, b2L = b2F + blkSz - 1;
        set.CKT_BLOCK2_FIRST = String(b2F);
        set.CKT_BLOCK2_LAST  = String(b2L);
        set.CKT_LAST = String(b2L);
      } else {
        set.CKT_LAST = String(b1L);
      }
      if (vdCol) { set.VD_RANGE = vdCol + c1 + ":" + vdCol + set.CKT_LAST; }
    }
    if (ft1 > 0 && slots > 0) { set.FEEDTHRU_LAST = String(ft1 + slots - 1); }
    var keys = Object.keys(set);
    if (!keys.length) { return { derived: 0 }; }
    return SAZ.geoSetMany(ctx, sheetName, set).then(function () {
      return { derived: keys.length, values: set };
    });
  });
};

/* Write several GEO values by key, adding rows for keys not already present. */
SAZ.geoSetMany = function (ctx, sheetName, pairs) {
  var sheet = ctx.workbook.worksheets.getItem(sheetName);
  var nm = sheet.names.getItemOrNullObject("GEO_MAP");
  nm.load("isNullObject");
  return ctx.sync().then(function () {
    if (nm.isNullObject) { throw new Error("No GEO_MAP on \"" + sheetName + "\"."); }
    var blk = nm.getRange();
    blk.load("values,rowIndex,columnIndex,rowCount");
    sheet.protection.load("protected");
    return ctx.sync().then(function () {
      var v = blk.values, idx = {}, firstFree = -1;
      for (var i = 0; i < v.length; i++) {
        var k = String(v[i][0] || "").trim();
        if (k && k.toUpperCase() !== "GEO_MAP") { idx[k] = i; }
        else if (!k && firstFree < 0) { firstFree = i; }
      }
      var missing = Object.keys(pairs).filter(function (k) { return idx[k] === undefined; });
      if (missing.length && firstFree < 0) {
        throw new Error("GEO_MAP on \"" + sheetName + "\" has no spare rows for " +
          missing.join(", ") + ". Extend the GEO_MAP named range.");
      }
      return SAZ.withUnprotected(ctx, sheet, function () {
        var slot = firstFree;
        Object.keys(pairs).forEach(function (k) {
          var row = (idx[k] !== undefined) ? idx[k] : slot++;
          sheet.getRangeByIndexes(blk.rowIndex + row, blk.columnIndex, 1, 2)
            .values = [[ k, pairs[k] ]];
        });
        return ctx.sync();
      });
    });
  });
};

/* Rebuild GEO_MAP on the active sheet, then report what moved. */
SAZ.uiRebuildGeoMap = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    var name = active.name;
    return SAZ.rebuildGeoMap(ctx, name).then(function (res) {
      return SAZ.geoDeriveRows(ctx, name).then(function (d) {
        if (!res.changed && !d.derived) {
          return SAZ.notify("GEO_MAP already correct",
            "Nothing moved on " + name + ". Every anchored key already points at" +
            " the right cell.");
        }
        var detail = res.changes.map(function (c) {
          return "  " + c.key + ":  " + (c.from || "(blank)") + "  ->  " + c.to;
        }).join("\n");
        return SAZ.notify("GEO_MAP rebuilt",
          name + ": " + res.changed + " key" + (res.changed === 1 ? "" : "s") +
          " repointed" + (d.derived ? ", " + d.derived + " derived" : "") + "." +
          (detail ? "\n\n" + detail : ""));
      });
    });
  });
};


/* ==========================================================
   RESIZE  -  port of ResizePanel (922-1016) and SetBlockRows (2453-2489)

   Grow: insert rows after the block, clone the last existing row for format
   and formulas, then clear ONLY the constants so calculated columns survive.
   The VBA did that with SpecialCells(xlCellTypeConstants), which has no
   Office.js equivalent - here a cell is a constant if its formula does not
   start with '='. Getting this wrong wipes every engine formula in the grid.

   Shrink: delete the surplus rows from the bottom.
   ========================================================== */

/* v2.5: the grid's right edge for row shifts = two columns left of the engine
   (AO on panelboards / switchboards, BF on the MCC). Everything from the engine
   column rightwards - engine formulas, GEO_MAP, feed-through detail - shares
   rows with the circuit grid and must NEVER move when the grid grows/shrinks. */
SAZ.gridLastCol = function (g) {
  var m = /^([A-Z]+)/i.exec(String(SAZ.geoUsable(g && g.ENG_VLL_CELL) || ""));
  return m ? SAZ.gridColLet(SAZ.gridColNum(m[1]) - 2) : "AO";
};

/* v2.5 REWRITE. Shrink / grow by shifting CELLS in A:lastCol only.
   Before v2.5 this deleted / inserted ENTIRE sheet rows. The engine block
   (AQ26:AQ38), GEO_MAP (AS1:AX40) and feed-through detail sit on the same rows
   as the circuit grid, so trimming a switchboard to 10 circuits deleted AQ38
   (receptacle demand) and the GEO_MAP tail - every #REF! on Main/Bus Rating,
   Fault at Bus, Feeder, Demand and Horiz. Bus traced back to that. Row heights
   below the grid travel with the shifted cells so the Notes banner keeps its
   height. Growing clones the last circuit row (grid columns only) and blanks
   its typed constants by writing "" (merged description cells reject clear()). */
SAZ.setBlockRows = function (ctx, sheetName, firstRow, lastRowIn, wantRows, polesCol, lastCol) {
  var sheet = ctx.workbook.worksheets.getItem(sheetName);
  lastCol = lastCol || "AO";
  if (wantRows < 1) { wantRows = 1; }
  var have = lastRowIn - firstRow + 1;
  if (have < 1) { have = 1; }
  var endRow = firstRow + wantRows - 1;
  if (wantRows === have) { return Promise.resolve({ added: 0, removed: 0, endRow: endRow }); }
  var shift = wantRows - have;
  var used = sheet.getUsedRange();
  used.load("rowIndex,rowCount");
  return ctx.sync().then(function () {
    var tailFirst = lastRowIn + 1;
    var tailLast = Math.max(used.rowIndex + used.rowCount, tailFirst) + 2;
    var probes = [];
    for (var r = tailFirst; r <= tailLast; r++) {
      var pr = sheet.getRange(r + ":" + r); pr.format.load("rowHeight"); probes.push(pr);
    }
    var gridRow = sheet.getRange(lastRowIn + ":" + lastRowIn);
    gridRow.format.load("rowHeight");
    return ctx.sync().then(function () {
      var heights = probes.map(function (p) { return p.format.rowHeight; });
      if (shift < 0) {
        sheet.getRange("A" + (endRow + 1) + ":" + lastCol + lastRowIn).delete(Excel.DeleteShiftDirection.up);
      } else {
        sheet.getRange("A" + tailFirst + ":" + lastCol + endRow).insert(Excel.InsertShiftDirection.down);
      }
      return ctx.sync().then(function () {
        heights.forEach(function (h, i) {
          var nr = tailFirst + i + shift;
          if (nr >= 1) { sheet.getRange(nr + ":" + nr).format.rowHeight = h; }
        });
        if (shift > 0) { sheet.getRange(tailFirst + ":" + endRow).format.rowHeight = gridRow.format.rowHeight; }
        return ctx.sync();
      }).then(function () {
        if (shift < 0) { return { added: 0, removed: -shift, endRow: endRow }; }
        var band = sheet.getRange("A" + tailFirst + ":" + lastCol + endRow);
        band.copyFrom(sheet.getRange("A" + lastRowIn + ":" + lastCol + lastRowIn), Excel.RangeCopyType.all);
        return ctx.sync().then(function () {
          band.load("formulas");
          return ctx.sync();
        }).then(function () {
          var f = band.formulas;
          for (var i = 0; i < f.length; i++) {
            for (var j = 0; j < f[i].length; j++) {
              var cell = f[i][j];
              var isFormula = (typeof cell === "string" && cell.charAt(0) === "=");
              if (cell !== "" && cell !== null && !isFormula) { band.getCell(i, j).values = [[""]]; }
            }
          }
          if (polesCol && polesCol !== "-") {
            var ones = [];
            for (var k = 0; k < shift; k++) { ones.push([1]); }
            sheet.getRange(polesCol + tailFirst + ":" + polesCol + endRow).values = ones;
          }
          return ctx.sync();
        }).then(function () { return { added: shift, removed: 0, endRow: endRow }; });
      });
    });
  });
};

/* Re-draw the circuit-grid row lines after any resize (all TPL_ layouts).
   Panelboards: the odd/even seam is a Medium bottom border on the LAST ODD row;
   shrinking deletes that row and growing clones it onto every new row.
   Switchboards / MCCs: shrinking drops the closing line under the dashboard.
   So the lines are rebuilt from the final geometry: every grid row gets its
   normal inner line, the seam row (TWOBLOCK only) gets Medium, the last row
   closes the grid in Thin. Spans come from GEO_MAP so any layout works: the
   circuit entry dashboard is the 9 class columns ending at CKT_KVA_COL and the
   schedule runs from B to two columns left of it (one gutter column). The inner
   weight is read from the sheet's own first grid row (panelboard dashboard =
   Hairline, everything else Thin). Must run inside withUnprotected. */
SAZ.gridColNum = function (L) {
  var n = 0;
  String(L).toUpperCase().replace(/[^A-Z]/g, "").split("").forEach(function (ch) { n = n * 26 + ch.charCodeAt(0) - 64; });
  return n;
};
SAZ.gridColLet = function (n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};
SAZ.gridSpans = function (g) {
  var kva = SAZ.gridColNum(SAZ.geoUsable(g.CKT_KVA_COL) || "W"), d1 = kva - 8;
  return [ { first: "B", last: SAZ.gridColLet(d1 - 2) },            /* circuit schedule */
           { first: SAZ.gridColLet(d1), last: SAZ.gridColLet(kva) } ]; /* circuit entry dashboard */
};
SAZ.restampGrid = function (ctx, sheetName, g, firstRow, seamRow, lastRow) {
  var sheet = ctx.workbook.worksheets.getItem(sheetName), spans = SAZ.gridSpans(g);
  var probes = spans.map(function (sp) {
    var b = sheet.getRange(sp.first + firstRow).format.borders.getItem("EdgeBottom");
    b.load("style,weight");
    return b;
  });
  var setB = function (rg, edge, weight) {
    var b = rg.format.borders.getItem(edge);
    b.style = "Continuous"; b.weight = weight; b.color = "#000000";
  };
  return ctx.sync().then(function () {
    spans.forEach(function (sp, i) {
      var p = probes[i], inner = "Thin";
      if (p.style === "Continuous" && p.weight !== "Medium" && firstRow !== seamRow && firstRow !== lastRow) { inner = p.weight; }
      var span = function (r1, r2) { return sheet.getRange(sp.first + r1 + ":" + sp.last + r2); };
      if (lastRow > firstRow) { setB(span(firstRow, lastRow), "InsideHorizontal", inner); }
      setB(span(lastRow, lastRow), "EdgeBottom", "Thin");
      if (seamRow && seamRow >= firstRow && seamRow < lastRow) { setB(span(seamRow, seamRow), "EdgeBottom", "Medium"); }
    });
    return ctx.sync();
  });
};

/* Grid totals (load summary K6:K14 / H14:H16, engine VD / phase cells) are
   absolute ranges that start on the first circuit row, e.g. SUM($O$28:$O$47).
   Growing inserts rows BELOW the last circuit row, which Excel does not fold
   into those ranges - the new circuits silently dropped out of every total and
   the demand that feeds the feeder sizing. After every resize, any absolute
   range anywhere on the sheet that starts on the first circuit row is repointed
   to end on the last one. Runs every resize, so panels grown earlier heal too. */
SAZ.healGridTotals = function (ctx, sheetName, firstRow, lastRow) {
  var sheet = ctx.workbook.worksheets.getItem(sheetName), used = sheet.getUsedRange();
  used.load("formulas,rowIndex");
  return ctx.sync().then(function () {
    var re = new RegExp("\\$([A-Z]{1,3})\\$" + firstRow + ":\\$([A-Z]{1,3})\\$(\\d+)(?![0-9])", "g"), fixed = 0;
    used.formulas.forEach(function (row, i) {
      /* no row skip: the engine cells (AQ / BH) sit beside the grid rows */
      row.forEach(function (f, j) {
        if (typeof f !== "string" || f.charAt(0) !== "=") { return; }
        var nf = f.replace(re, function (m, c1, c2) { return "$" + c1 + "$" + firstRow + ":$" + c2 + "$" + lastRow; });
        if (nf !== f) { used.getCell(i, j).formulas = [[nf]]; fixed++; }
      });
    });
    return ctx.sync().then(function () { return fixed; });
  });
};


/* Multi-pole tie flag (column AN, 1 = this row is the 2nd/3rd pole of the breaker
   above). Row n checks n-1 for >=2 poles and n-2 for 3 poles. The SECOND row of
   each block must check only the row above: the old template's ckt-4 row looked
   two rows up into the last ODD row, so a shrink (which deletes that row) left
   #REF! cascading down the even block, and a 3-pole at the last odd ckt wrongly
   tied ckt 4. Rewritten on every TWOBLOCK resize so older panels self-heal. */
SAZ.PNL_TIE_COL = "AN";
SAZ.healPanelTies = function (ctx, sheetName, b1F, b2F, perBlock, polesCol) {
  if (!polesCol || polesCol === "-" || perBlock < 2) { return Promise.resolve(); }
  var sheet = ctx.workbook.worksheets.getItem(sheetName), tc = SAZ.PNL_TIE_COL;
  [b1F, b2F].forEach(function (top) {
    sheet.getRange(tc + top).values = [[0]];
    sheet.getRange(tc + (top + 1)).formulas = [[
      "=IF(AND($" + polesCol + top + ">=2,$" + tc + top + "=0),1,0)" ]];
  });
  return ctx.sync();
};

/* Resize a panel to a target space count, then self-heal GEO_MAP.
   TWOBLOCK (panelboards) keeps both blocks the SAME size - odd circuits above,
   even below - or the numbering goes asymmetric. The bottom block is resized
   FIRST so the upper rows do not shift underneath the operation. */
SAZ.resizePanel = function (ctx, tag, spaces) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  return ctx.sync().then(function () {
    var hit = sheets.items.filter(function (s) {
      return s.name.toLowerCase() === tag.toLowerCase();
    })[0];
    if (!hit) { return SAZ.notify("Not found", "No sheet named \"" + tag + "\"."); }
    var name = hit.name;
    var sheet = ctx.workbook.worksheets.getItem(name);
    return SAZ.repairSchedule(ctx, name).then(function () { return SAZ.geo(ctx, name); }).then(function (g) {   /* v2.5: heal first - a damaged GEO_MAP would mis-size the grid */
      var ckt1    = Number(g.CKT_FIRST) || 28;
      var gridMax = Number(g.CKT_LAST) || ckt1;
      if (gridMax < ckt1) { gridMax = ckt1; }
      var rowsPer = Number(g.CKT_PER_ROW) || 1;
      var mode    = String(g.RESIZE_MODE || "").trim().toUpperCase();
      var cktCol  = SAZ.geoUsable(g.CKT_NUM_COL) || "B";
      var pc      = SAZ.geoUsable(g.CKT_POLES_COL);
      var lc      = SAZ.gridLastCol(g);   /* v2.5: shift only the schedule columns - never the engine / GEO_MAP */
        /* rows cannot be inserted or deleted on a protected sheet. withUnprotected
           restores house-standard protection on success AND failure; GEO_MAP is
           healed inside, before reprotecting */
        return SAZ.withUnprotected(ctx, sheet, function () {
          return ctx.sync().then(function () {
          if (mode === "TWOBLOCK") {
            var b1F = Number(g.CKT_BLOCK1_FIRST) || ckt1;
            var b1L = Number(g.CKT_BLOCK1_LAST)  || b1F;
            var b2F = Number(g.CKT_BLOCK2_FIRST) || 0;
            var b2L = Number(g.CKT_BLOCK2_LAST)  || 0;
            if (b1L < b1F) { b1L = b1F; }
            if (b2F <= b1L) { b2F = b1L + 1; }
            if (b2L < b2F) { b2L = b2F; }
            var newBlk = Math.max(1, Math.ceil(spaces / 2));
            /* bottom block FIRST - order matters */
            return SAZ.setBlockRows(ctx, name, b2F, b2L, newBlk, pc, lc).then(function () {
              return SAZ.setBlockRows(ctx, name, b1F, b1L, newBlk, pc, lc);
            }).then(function () {
              var n1L = b1F + newBlk - 1;
              var n2F = n1L + 1;
              var n2L = n2F + newBlk - 1;
              sheet.getRange(cktCol + b1F).formulas = [[ "=2*(ROW()-" + b1F + ")+1" ]];
              sheet.getRange(cktCol + n2F).formulas = [[ "=2*(ROW()-" + n2F + ")+2" ]];
              return ctx.sync().then(function () {
                if (newBlk > 1) {
                  sheet.getRange(cktCol + (b1F + 1) + ":" + cktCol + n1L)
                    .copyFrom(sheet.getRange(cktCol + b1F), Excel.RangeCopyType.formulas);
                  sheet.getRange(cktCol + (n2F + 1) + ":" + cktCol + n2L)
                    .copyFrom(sheet.getRange(cktCol + n2F), Excel.RangeCopyType.formulas);
                }
                return ctx.sync();
              }).then(function () {
                return SAZ.geoSetMany(ctx, name, {
                  BLOCK_SIZE: String(newBlk),
                  CKT_BLOCK1_FIRST: String(b1F), CKT_BLOCK1_LAST: String(n1L),
                  CKT_BLOCK2_FIRST: String(n2F), CKT_BLOCK2_LAST: String(n2L),
                  CKT_LAST: String(n2L),
                  VD_RANGE: (SAZ.geoUsable(g.VD_COL) || "K") + b1F + ":" + (SAZ.geoUsable(g.VD_COL) || "K") + n2L
                });
              }).then(function () { return SAZ.healPanelTies(ctx, name, b1F, n2F, newBlk, pc); }).then(function () { return SAZ.restampGrid(ctx, name, g, b1F, n1L, n2L); }).then(function () { return SAZ.healGridTotals(ctx, name, b1F, n2L); }).then(function () { return { mode: "TWOBLOCK", perBlock: newBlk, cktLast: n2L }; });
            });
          }
          /* single contiguous block - CKT = 1,2,3... */
          var lastRow = Math.max(ckt1, (ckt1 - 1) + Math.ceil(spaces / rowsPer));
          return SAZ.setBlockRows(ctx, name, ckt1, gridMax, lastRow - ckt1 + 1, pc, lc)
            .then(function () {
              sheet.getRange(cktCol + ckt1).formulas = [[ "=ROW()-" + (ckt1 - 1) ]];
              return ctx.sync();
            }).then(function () {
              if (lastRow > ckt1) {
                sheet.getRange(cktCol + (ckt1 + 1) + ":" + cktCol + lastRow)
                  .copyFrom(sheet.getRange(cktCol + ckt1), Excel.RangeCopyType.formulas);
                return ctx.sync();
              }
            }).then(function () {
              return SAZ.geoSetMany(ctx, name, {
                BLOCK_SIZE: String(lastRow - ckt1 + 1),
                CKT_LAST:   String(lastRow),
                CKT_BLOCK1_FIRST: String(ckt1), CKT_BLOCK1_LAST: String(lastRow),
                VD_RANGE: (SAZ.geoUsable(g.VD_COL) || "K") + ckt1 + ":" + (SAZ.geoUsable(g.VD_COL) || "K") + lastRow
              });
            }).then(function () { return SAZ.restampGrid(ctx, name, g, ckt1, 0, lastRow); }).then(function () { return SAZ.healGridTotals(ctx, name, ckt1, lastRow); }).then(function () { return { mode: "SINGLE", cktLast: lastRow }; });
          }).then(function (res) {
            /* row moves invalidate every anchored address - heal before reprotecting */
            return SAZ.rebuildGeoMap(ctx, name).then(function (rb) {
              return { res: res, rb: rb };
            });
          });
        });
    });
  });
};

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

SAZ.renameBoard = function (ctx, oldTag, newTag) { return SAZ_RENAME.run(ctx, oldTag, newTag); };

/* Menu: Rename Equipment. Defaults to the active schedule. */
SAZ.uiRename = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  var geo = active.names.getItemOrNullObject("GEO_MAP");
  geo.load("isNullObject");
  return ctx.sync().then(function () {
    var dflt = geo.isNullObject ? "" : active.name;
    return SAZ.prompt("Rename Equipment", "Current tag of the board to rename", dflt, true)
      .then(function (a) {
        if (!a.ok) { return null; }
        var oldTag = String(a.value || "").trim();
        return SAZ.prompt("Rename Equipment",
          "New tag for " + oldTag + "  (renames the tab, the schedule name and every reference)",
          oldTag, true).then(function (b) {
            if (!b.ok) { return null; }
            return SAZ.renameBoard(ctx, oldTag, b.value).then(function (res) {
              return SAZ.notify("Equipment renamed",
                res.from + " is now " + res.to + "." +
                (res.sheetRenamed ? " The tab was renamed." : " (Registry entry only - no schedule sheet.)") +
                "\n\n" + res.cells + " text reference(s) updated; formula links follow automatically.");
            }, function (e) {
              return SAZ.notify("Rename not done", e.message || String(e));
            });
          });
      });
  });
};


/* ==========================================================
   v2.5 REPAIR SCHEDULE
   Heals a schedule damaged by the pre-v2.5 resize (whole-row deletes). Works
   from the board's own TPL_ template, cell by cell, so typed inputs, circuit
   loads and live feed-through links are never overwritten:
     * schedule area (A : grid edge): any formula showing #REF! is re-copied from
       the template's matching cell (above the grid = same cell; circuit rows =
       the template's row pattern; below the grid = offset by the size change)
     - engine / GEO_MAP area: #REF! cells and cells the template has but the
       board lost are restored from the template
   then GEO_MAP's named range, grid keys, grid totals and panel ties are
   re-derived from the board's real circuit rows. Safe to run on a healthy
   board - it reports 0 repairs. Runs automatically before every resize.
   ========================================================== */
SAZ.templateFor = function (ctx, sheetName) {
  var head = ctx.workbook.worksheets.getItem(sheetName).getRange("A1:H3");
  head.load("values");
  return ctx.sync().then(function () {
    var t = head.values.map(function (r) { return r.join(" "); }).join(" ").toUpperCase();
    if (/MOTOR CONTROL|\bMCC\b/.test(t)) { return "TPL_MCC"; }
    if (/PANELBOARD/.test(t)) { return "TPL_Panelboard"; }
    if (/SWITCHBOARD|SWITCHGEAR/.test(t)) { return "TPL_Switchboard"; }
    return "";
  });
};
SAZ.SYNC_CELLS  = ["D6", "D11", "D13"];                 /* Fed From, Mains Type, Std Device AIC - calculated since v2.5 */
SAZ.SYNC_GRID   = { "TPL_Panelboard": ["H", "AD", "AH"] }; /* per-pole kVA display + total-based FLA / VA */
SAZ.SYNC_ENGINE = { "TPL_Panelboard": ["AQ34"] };            /* largest motor from VA#, not the per-pole display */
SAZ.repairSchedule = function (ctx, name) {
  var sheet = ctx.workbook.worksheets.getItem(name);
  return SAZ.templateFor(ctx, name).then(function (tpl) {
    if (!tpl) { throw new Error("\"" + name + "\" is not a panelboard, switchboard or MCC schedule."); }
    return SAZ.ensureTemplate(ctx, tpl).then(function () { return SAZ.geo(ctx, tpl); }).then(function (tg) {
      var t = ctx.workbook.worksheets.getItem(tpl);
      var tNm = t.names.getItem("GEO_MAP").getRange(); tNm.load("address");
      var used = sheet.getUsedRange(); used.load("formulas,rowIndex,columnIndex");
      var tu = t.getUsedRange(); tu.load("formulas,rowIndex,columnIndex");
      return ctx.sync().then(function () {
        var F = function (rg, r, c) {           /* 1-based row / column */
          var row = rg.formulas[r - 1 - rg.rowIndex];
          if (!row) { return ""; }
          var v = row[c - 1 - rg.columnIndex];
          return (v === null || v === undefined) ? "" : String(v);
        };
        var engCol = SAZ.gridColNum(String(tg.ENG_VLL_CELL).replace(/[0-9$]/g, "")) - 1;
        var gridLast = SAZ.gridColNum(SAZ.gridLastCol(tg));
        var tFirst = Number(tg.CKT_FIRST), tLast = Number(tg.CKT_LAST);
        var two = String(tg.RESIZE_MODE || "").toUpperCase() === "TWOBLOCK";
        var tB2F = Number(tg.CKT_BLOCK2_FIRST) || 0;
        var nc = SAZ.gridColNum(SAZ.geoUsable(tg.CKT_NUM_COL) || "B");
        /* the board's real circuit rows, read from its own CKT numbering formulas */
        var b1L = 0, b2F = 0, cur = tFirst - 1;
        for (var r = tFirst; r < tFirst + 400; r++) {
          var f = F(used, r, nc).replace(/\s/g, "");
          if (!/^=(2\*\()?ROW\(\)/i.test(f)) { break; }
          cur = r;
          if (two) { if (/\+2$/.test(f)) { if (!b2F) { b2F = r; } } else { b1L = r; } }
        }
        if (cur < tFirst) { throw new Error("Could not find the circuit rows on \"" + name + "\" (column " + SAZ.gridColLet(nc) + " numbering is missing)."); }
        if (two && !b2F) { b2F = b1L + 1; }
        var tplRowFor = function (r) {
          if (r < tFirst) { return r; }
          if (r > cur) { return r + (tLast - cur); }
          var bs = (two && r >= b2F) ? b2F : tFirst, ts = (two && r >= b2F) ? tB2F : tFirst;
          return ts + Math.min(r - bs, 2);
        };
        var fixes = [], maxR = Math.max(used.rowIndex + used.formulas.length, tu.rowIndex + tu.formulas.length);
        var maxC = Math.max(used.columnIndex + (used.formulas[0] || []).length, tu.columnIndex + (tu.formulas[0] || []).length);
        for (var rr = 1; rr <= maxR; rr++) {
          for (var cc = 1; cc <= maxC; cc++) {
            var mine = F(used, rr, cc);
            if (cc <= gridLast) {
              if (mine.indexOf("#REF!") >= 0) { fixes.push({ r: rr, c: cc, tr: tplRowFor(rr), all: false }); }
            } else if (cc >= engCol) {
              var theirs = F(tu, rr, cc);
              if (mine.indexOf("#REF!") >= 0 || (mine === "" && theirs !== "")) { fixes.push({ r: rr, c: cc, tr: rr, all: true }); }
            }
          }
        }
        return SAZ.withUnprotected(ctx, sheet, function () {
          fixes.forEach(function (x) {
            var a = SAZ.gridColLet(x.c);
            sheet.getRange(a + x.r).copyFrom(t.getRange(a + x.tr), x.all ? Excel.RangeCopyType.all : Excel.RangeCopyType.formulas);
          });
          /* GEO_MAP name back to the template's full block */
          /* v2.5: cells the Distribution Map now drives (Fed From, Mains Type, Std Device AIC)
             and the panelboard per-pole kVA columns are re-stamped from the template on
             every repair, so boards deployed before v2.5 pick the automation up. */
          SAZ.SYNC_CELLS.forEach(function (a) {
            var c = sheet.getRange(a); c.dataValidation.clear();
            c.copyFrom(t.getRange(a), Excel.RangeCopyType.all);
          });
          (SAZ.SYNC_GRID[tpl] || []).forEach(function (cl) {
            sheet.getRange(cl + tFirst + ":" + cl + cur).copyFrom(t.getRange(cl + tFirst), Excel.RangeCopyType.formulas);
          });
          (SAZ.SYNC_ENGINE[tpl] || []).forEach(function (a) {
            sheet.getRange(a).copyFrom(t.getRange(a), Excel.RangeCopyType.formulas);
          });
          var nm = sheet.names.getItemOrNullObject("GEO_MAP");
          nm.load("isNullObject");
          return ctx.sync().then(function () {
            var addr = tNm.address.split("!").pop();
            if (!nm.isNullObject) { nm.delete(); }
            sheet.names.add("GEO_MAP", sheet.getRange(addr));
            return ctx.sync();
          }).then(function () {
            var pairs = { CKT_FIRST: String(tFirst), CKT_LAST: String(cur), RESIZE_MODE: String(tg.RESIZE_MODE || "SINGLE") };
            var vd = SAZ.geoUsable(tg.VD_COL) || "K";
            pairs.VD_RANGE = vd + tFirst + ":" + vd + cur;
            if (two) {
              pairs.BLOCK_SIZE = String(b1L - tFirst + 1);
              pairs.CKT_BLOCK1_FIRST = String(tFirst); pairs.CKT_BLOCK1_LAST = String(b1L);
              pairs.CKT_BLOCK2_FIRST = String(b2F);    pairs.CKT_BLOCK2_LAST = String(cur);
            } else {
              pairs.BLOCK_SIZE = String(cur - tFirst + 1);
              pairs.CKT_BLOCK1_FIRST = String(tFirst); pairs.CKT_BLOCK1_LAST = String(cur);
              pairs.CKT_BLOCK2_FIRST = "-"; pairs.CKT_BLOCK2_LAST = "-";
            }
            return SAZ.geoSetMany(ctx, name, pairs);
          }).then(function () {
            return SAZ.healGridTotals(ctx, name, tFirst, cur);
          }).then(function (healed) {
            var pc = SAZ.geoUsable(tg.CKT_POLES_COL);
            var p = two ? SAZ.healPanelTies(ctx, name, tFirst, b2F, b1L - tFirst + 1, pc) : Promise.resolve();
            return p.then(function () { return SAZ.rebuildGeoMap(ctx, name); })
              .then(function () { return { sheet: name, template: tpl, repaired: fixes.length, totals: healed, cktLast: cur }; });
          });
        });
      });
    });
  });
};
SAZ.uiRepair = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    return SAZ.repairSchedule(ctx, active.name).then(function (o) {
      return SAZ.notify("Schedule checked", o.repaired ?
        (o.sheet + ": " + o.repaired + " damaged cell" + (o.repaired === 1 ? "" : "s") + " restored from " + o.template + "; grid keys and totals re-derived.") :
        (o.sheet + ": nothing damaged. Grid keys and totals re-checked."));
    });
  });
};
SAZ.uiResize = function (ctx, presetTag, err) {
  /* v2.5: ONE form (panel + space count). The old version chained two prompts;
     Office cannot open a second dialog while the first is still closing, so the
     space-count box never appeared and the command silently stopped. The
     Distribution Map sets spaces directly and calls SAZ.resizePanel itself. */
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    var tag0 = presetTag || active.name;
    return SAZ.geo(ctx, tag0).then(function (g) { return SAZ.currentSpaces(g); }, function () { return ""; })
      .then(function (cur) {
        return SAZ.form("Resize Panel", "Grows or shrinks the circuit grid. Panelboards keep an even odd/even split.", [
          { id: "tag", label: "Panel / board", type: "text", value: tag0, required: true },
          { id: "spaces", label: "New space count" + (cur ? " (now " + cur + ")" : ""), type: "number", value: String(cur || 42), required: true, min: 1, max: 200 }
        ], err || "", 45).then(function (a) {
          if (!a || !a.ok || !a.values) { return null; }
          var tag = SAZ.sanitise(a.values.tag);
          var spaces = parseInt(String(a.values.spaces).replace(/[^0-9]/g, ""), 10);
          if (!tag) { return SAZ.uiResize(ctx, "", "No usable characters in the panel tag."); }
          if (!spaces || spaces < 1) { return SAZ.uiResize(ctx, tag, "Enter a whole number of spaces, 1 or more."); }
          return SAZ.resizePanel(ctx, tag, spaces).then(function (out) {
            if (!out || !out.res) { return null; }
            return SAZ.notify("Panel resized", tag + " resized to " + spaces + " spaces. Circuit grid now ends at row " + out.res.cktLast + ".");
          });
        });
      });
  });
};


/* ==========================================================
   ADD / TRIM CIRCUIT ROWS  -  port of PickEquipTable (2016-2065)

   The ribbon labels say "Circuit Rows" but these buttons do NOT touch an
   equipment circuit grid - equipment sheets carry no ListObjects at all.
   They operate on the named tables listed in the registry on Load Calc:
   column W holds a numbered label, column X the table name. Row 11 of that
   range is deliberately blank and is skipped. The number is PARSED FROM THE
   LABEL, so the mapping stays correct if the registry is ever renumbered or
   reordered. Generator and UPS load tables are appended after the highest
   registry number.
   ========================================================== */

SAZ.EQUIP_REGISTRY_SHEET = "Load Calc";
SAZ.EQUIP_REGISTRY_RANGE = "AA7:AB40";
SAZ.EQUIP_EXTRA_TABLES = ["GEN_Loads", "UPS_Loads"];

SAZ.equipTableRegistry = function (ctx) {
  var sheet = ctx.workbook.worksheets.getItemOrNullObject(SAZ.EQUIP_REGISTRY_SHEET);
  sheet.load("isNullObject");
  return ctx.sync().then(function () {
    if (sheet.isNullObject) {
      throw new Error("No \"" + SAZ.EQUIP_REGISTRY_SHEET + "\" sheet, so the table" +
        " registry cannot be read. Create the Load Calc module first.");
    }
    var rg = sheet.getRange(SAZ.EQUIP_REGISTRY_RANGE);
    rg.load("values");
    return ctx.sync().then(function () {
      var out = [], maxNum = 0;
      rg.values.forEach(function (row) {
        var label = String(row[0] || "").trim();
        var name  = String(row[1] || "").trim();
        if (!name) { return; }                 /* blank registry row - skip */
        var m = label.match(/^\s*(\d+)/);
        var num = m ? parseInt(m[1], 10) : 0;  /* parsed, never positional */
        if (num > maxNum) { maxNum = num; }
        out.push({ num: num, label: label, table: name });
      });
      SAZ.EQUIP_EXTRA_TABLES.forEach(function (t, i) {
        out.push({ num: maxNum + i + 1, label: (maxNum + i + 1) + " " +
          t.replace(/_/g, " "), table: t });
      });
      return out;
    });
  });
};

SAZ.pickEquipTable = function (ctx, verb) {
  return SAZ.equipTableRegistry(ctx).then(function (list) {
    var menu = list.map(function (e) { return e.label; }).join("\n");
    return SAZ.prompt("Which table do you want to " + verb + "?", menu, "1", true)
      .then(function (a) {
        if (!a.ok) { return null; }
        var pick = parseInt(String(a.value).replace(/[^0-9]/g, ""), 10);
        var hit = list.filter(function (e) { return e.num === pick; })[0];
        if (!hit) {
          return SAZ.notify("No such table",
            "Nothing is numbered " + pick + ". Choose a number from the list.")
            .then(function () { return null; });
        }
        return hit;
      });
  });
};

/* Resolve a table by name anywhere in the workbook, with a clear failure. */
SAZ.tableByName = function (ctx, tableName) {
  var t = ctx.workbook.tables.getItemOrNullObject(tableName);
  t.load("isNullObject,name,worksheet/name");
  return ctx.sync().then(function () {
    if (t.isNullObject) {
      throw new Error("Table \"" + tableName + "\" is registered on " +
        SAZ.EQUIP_REGISTRY_SHEET + " but does not exist in this workbook." +
        " The module that owns it has not been created yet.");
    }
    return t;
  });
};


/* Add N rows to the chosen table. Table row insertion carries calculated-column
   formulas but NOT hand-applied formatting, so the new rows are repainted from
   the pattern row by matchTableRowFormats. */
SAZ.uiEquipAddRows = function (ctx) {
  return SAZ.pickEquipTable(ctx, "add rows to").then(function (pickd) {
    if (!pickd) { return null; }
    return SAZ.prompt("Add Rows", "How many rows to add to " + pickd.table + "?", "1", true)
      .then(function (a) {
        if (!a.ok) { return null; }
        var n = parseInt(String(a.value).replace(/[^0-9]/g, ""), 10);
        if (!n || n < 1) {
          return SAZ.notify("Invalid count", "Enter a whole number of rows, 1 or more.");
        }
        return SAZ.tableByName(ctx, pickd.table).then(function (t) {
          for (var i = 0; i < n; i++) { t.rows.add(); }
          return ctx.sync().then(function () {
            return SAZ.matchTableRowFormats(ctx, t, n);
          }).then(function () {
            var body = t.getDataBodyRange();
            body.load("rowCount,address");
            return ctx.sync().then(function () {
              t.getDataBodyRange().getLastRow().select();
              return ctx.sync().then(function () {
                return SAZ.notify("Rows added",
                  n + " row" + (n === 1 ? "" : "s") + " added to " + pickd.table +
                  ". It now has " + body.rowCount + " rows.");
              });
            });
          });
        });
      });
  });
};

/* Trim N rows from the bottom of the chosen table.
   Refuses to empty a table completely - a ListObject with no data body is a
   nuisance to recover, and every formula pointed at the column breaks. */
SAZ.uiEquipTrimRows = function (ctx) {
  return SAZ.pickEquipTable(ctx, "trim rows from").then(function (pickd) {
    if (!pickd) { return null; }
    return SAZ.tableByName(ctx, pickd.table).then(function (t) {
      var body = t.getDataBodyRange();
      body.load("rowCount");
      return ctx.sync().then(function () {
        var have = body.rowCount;
        if (have <= 1) {
          return SAZ.notify("Nothing to trim",
            pickd.table + " has only " + have + " row" + (have === 1 ? "" : "s") +
            ". At least one must remain.");
        }
        return SAZ.prompt("Trim Rows",
          "How many rows to remove from the bottom of " + pickd.table +
          "?  (" + have + " present, at least 1 must remain)", "1", true)
          .then(function (a) {
            if (!a.ok) { return null; }
            var n = parseInt(String(a.value).replace(/[^0-9]/g, ""), 10);
            if (!n || n < 1) {
              return SAZ.notify("Invalid count", "Enter a whole number of rows, 1 or more.");
            }
            if (n > have - 1) {
              return SAZ.notify("Too many",
                "Removing " + n + " would leave " + pickd.table + " empty." +
                " The most you can trim is " + (have - 1) + ".");
            }
            t.rows.load("items/index");
            return ctx.sync().then(function () {
              var rows = t.rows.items.slice();
              rows.sort(function (x, y) { return y.index - x.index; });
              for (var i = 0; i < n; i++) { rows[i].delete(); }
              return ctx.sync();
            }).then(function () {
              return SAZ.notify("Rows trimmed",
                n + " row" + (n === 1 ? "" : "s") + " removed from " + pickd.table +
                ". It now has " + (have - n) + " rows.");
            });
          });
      });
    });
  });
};


/* ==========================================================
   LOAD CALC SECTION TOGGLE  -  port of Rbn_ToggleSection (2558-2613)
   and SetSectionState (2618-2645).

   Same registry as Add/Trim Table Rows, one column wider: W = numbered
   label, X = table name, Y = the Yes/No on-off flag the engine reads.

   DEVIATION: the VBA guards against nesting outline levels with
   'If ws.Rows(firstRow).OutlineLevel < 2'. Office.js exposes no readable
   outline level, so instead of testing we NORMALISE - ungroup then group,
   which lands on exactly one level whatever the starting state.
   ========================================================== */

SAZ.SECTION_REGISTRY_RANGE = "AA7:AC40";   /* v2.6: rows 19-40 take departments added at run time; v2.8 registry moved W:AD -> AA:AH (collapsed group Z:AH) */
SAZ.SECTION_OFF_COLOUR = "#7E868C";
SAZ.SECTION_ON_COLOUR  = "#FFFFFF";

SAZ.sectionRegistry = function (ctx) {
  var sheet = ctx.workbook.worksheets.getItemOrNullObject(SAZ.EQUIP_REGISTRY_SHEET);
  sheet.load("isNullObject");
  return ctx.sync().then(function () {
    if (sheet.isNullObject) {
      throw new Error("No \"" + SAZ.EQUIP_REGISTRY_SHEET + "\" sheet in this workbook.");
    }
    var rg = sheet.getRange(SAZ.SECTION_REGISTRY_RANGE);
    rg.load("values,rowIndex");
    return ctx.sync().then(function () {
      var out = [];
      rg.values.forEach(function (row, i) {
        var label = String(row[0] || "").trim();
        var name  = String(row[1] || "").trim();
        if (!name) { return; }
        var m = label.match(/^\s*(\d+)/);
        out.push({
          num:   m ? parseInt(m[1], 10) : 0,
          label: label,
          table: name,
          state: String(row[2] || "").trim(),
          flagRow: rg.rowIndex + i + 1        /* 1-based worksheet row */
        });
      });
      return out;
    });
  });
};

/* REWORKED 2026-10-22 - optional sections.
   Baseline sections are always on. A section whose registry label (column W)
   ends "(optional)" starts OFF and is added from the ribbon when the project
   needs it. OFF = Include flag "No" (every total already honours it) AND the
   whole block - banner through its demand-factor subtotal row - collapsed in
   a row group. The spacer row under each block is left OUT of the group so
   neighbouring optional sections stay separate groups (Excel merges touching
   groups into one, which is what broke the Transformers columns earlier).
   Rows are never deleted, so no table, name or fixed reference can break. */
SAZ.SECTION_OPTIONAL_TAG = /\(optional\)\s*$/i;

SAZ.sectionBand = function (ctx, sheet, table) {
  var hdr = table.getHeaderRowRange();
  var whole = table.getRange();
  hdr.load("rowIndex");
  whole.load("rowIndex,rowCount");
  return ctx.sync().then(function () {
    var headerRow = hdr.rowIndex + 1;                      /* 1-based */
    var tableEnd  = whole.rowIndex + whole.rowCount;       /* 1-based last row */
    var probe = sheet.getRange("B" + (tableEnd + 1) + ":B" + (tableEnd + 6));
    probe.load("values");
    return ctx.sync().then(function () {
      var sub = tableEnd;
      for (var i = 0; i < probe.values.length; i++) {
        if (String(probe.values[i][0] || "").toUpperCase().indexOf("NEC DEMAND FACTOR") === 0) {
          sub = tableEnd + 1 + i; break;
        }
      }
      return { first: Math.max(1, headerRow - 2), last: sub, banner: headerRow - 2 };
    });
  });
};

/* Collapse or expand a block of rows as ONE outline group. Normalise first
   (ungroup, then group) so repeated toggles never stack outline levels. */
SAZ.setBandState = function (ctx, sheet, first, last, turnOn) {
  var band = sheet.getRange(first + ":" + last);
  band.ungroup(Excel.GroupOption.byRows);
  return ctx.sync().catch(function () { return null; }).then(function () {
    band.group(Excel.GroupOption.byRows);
    return ctx.sync().catch(function () { return null; });
  }).then(function () {
    band.rowHidden = !turnOn;
    return ctx.sync();
  });
};

SAZ.setSectionState = function (ctx, entry, turnOn) {
  var sheet = ctx.workbook.worksheets.getItem(SAZ.EQUIP_REGISTRY_SHEET);
  var t = sheet.tables.getItemOrNullObject(entry.table);
  t.load("isNullObject");
  return ctx.sync().then(function () {
    if (t.isNullObject) { return false; }   /* section not built yet - skip quietly */
    return SAZ.sectionBand(ctx, sheet, t).then(function (b) {
      sheet.getRange("AC" + entry.flagRow).values = [[turnOn ? "Yes" : "No"]];
      return SAZ.setBandState(ctx, sheet, b.first, b.last + 1, turnOn  /* +1: trailing spacer row */).then(function () {
        if (b.banner >= 1) {
          sheet.getRange("A" + b.banner).format.font.color =
            turnOn ? SAZ.SECTION_ON_COLOUR : SAZ.SECTION_OFF_COLOUR;
        }
        return SAZ.setRollupRows(ctx, sheet, entry.label, turnOn).then(function () { return SAZ.setSectionLists(ctx, entry.label, turnOn); }).then(function () { return true; });
      });
    });
  });
};

/* v1.2: hide/show an excluded section's one-line roll-ups (service summary,
   generator and UPS breakdowns). Rows are found from the TOTAL SERVICE LOAD
   SUMMARY banner down, by the section number leading the label in B or D. */
SAZ.setRollupRows = function (ctx, sheet, label, turnOn) {
  var name = String(label || "").replace(/^\s*\d+\s+/, "").replace(/\(optional\)/i, "").trim().toLowerCase();
  if (!name) { return ctx.sync(); }
  var used = sheet.getUsedRange(true);
  used.load("rowCount,rowIndex");
  var colA = sheet.getRange("A1:A400");
  colA.load("values");
  return ctx.sync().then(function () { return SAZ.moduleState(ctx); }).then(function (ms) { return SAZ.moduleBands(ctx, sheet).then(function (bd) { return [ms, bd]; }); }).then(function (mb) {
    var start = 0;
    colA.values.forEach(function (r, i) { if (!start && String(r[0]).indexOf("TOTAL SERVICE LOAD SUMMARY") === 0) { start = i + 1; } });
    if (!start) { start = 145; }
    var last = used.rowIndex + used.rowCount;
    var bd = sheet.getRange("B" + start + ":D" + last);
    bd.load("values");
    return ctx.sync().then(function () {
      var same = function (v) { return String(v || "").trim().toLowerCase() === name; };
      bd.values.forEach(function (r, i) {
        if (same(r[0]) || same(r[2])) {
          var n = start + i;
          var skip = Object.keys(mb[1]).some(function (k) { return !mb[0][k] && n >= mb[1][k].first && n <= mb[1][k].last; });   /* never surface a line inside an undeployed Generator/UPS block */
          if (!skip) { sheet.getRange(n + ":" + n).rowHidden = !turnOn; }
        }
      });
      return ctx.sync();
    });
  });
};
SAZ.setSectionLists = function (ctx, label, turnOn) {
  /* v1.2: Generator and UPS section lists follow the Load Calc include state. */
  var name = String(label || "").replace(/^\s*\d+\s+/, "").replace(/\(optional\)/i, "").trim().toLowerCase();
  if (!name) { return ctx.sync(); }
  var sheets = ["Generator", "UPS"].map(function (n) { var w = ctx.workbook.worksheets.getItemOrNullObject(n); w.load("isNullObject"); return w; });
  return ctx.sync().then(function () {
    var cols = sheets.map(function (w) { if (w.isNullObject) { return null; } var r = w.getRange("B1:B40"); r.load("values"); return r; });
    return ctx.sync().then(function () {
      cols.forEach(function (r, k) {
        if (!r) { return; }
        for (var i = 0; i < r.values.length; i++) {
          var t = String(r.values[i][0] || "").trim().toLowerCase();
          if (t.indexOf("total included") === 0) { break; }
          if (t === name) { sheets[k].getRange((i + 1) + ":" + (i + 1)).rowHidden = !turnOn; }
        }
      });
      return ctx.sync();
    });
  });
};


/* Logging Demand: the interval-data approach is three blocks - the interval
   SUMMARY, the DEMAND PROFILE (load duration curve + chart) and the RAW DATA
   paste area. Found by banner text, never by fixed row. */
SAZ.LD_SHEET = "Logging Demand";
SAZ.LD_INTERVAL_MODE = "Raw Interval Data";
SAZ.LD_MANUAL_MODE = "Manual Peak Entry";
SAZ.LD_BANNERS = ["RAW INTERVAL DATA \u2014 SUMMARY", "DEMAND PROFILE", "RAW INTERVAL DATA \u2014 PASTE"];

SAZ.ldBands = function (ctx, sheet) {
  var col = sheet.getRange("A1:A200");
  var used = sheet.getUsedRange();
  col.load("values");
  used.load("rowIndex,rowCount");
  return ctx.sync().then(function () {
    var banners = [];
    col.values.forEach(function (r, i) {
      var v = String(r[0] || "");
      /* a banner is a row whose first 12 characters are all capitals - the
         logger-export banner carries lowercase "kVAR" later in its text */
      var head = v.slice(0, 12);
      if (v.length > 6 && /[A-Z]/.test(head) && head === head.toUpperCase()) { banners.push({ row: i + 1, text: v }); }
    });
    var lastUsed = used.rowIndex + used.rowCount;
    var out = [];
    SAZ.LD_BANNERS.forEach(function (key) {
      var k = banners.findIndex(function (b) { return b.text.indexOf(key) === 0; });
      if (k < 0) { return; }
      var first = banners[k].row;
      var last = k + 1 < banners.length ? banners[k + 1].row - 1 : lastUsed   /* -1: band includes its trailing spacer row */;
      if (last < first) { last = first; }
      out.push({ key: key, first: first, last: last });
    });
    return out;
  });
};

SAZ.uiToggleInterval = function (ctx) {
  var sheet = ctx.workbook.worksheets.getItem(SAZ.LD_SHEET);
  var mode = sheet.getRange("C6");
  mode.load("values");
  return ctx.sync().then(function () {
    var isOn = String(mode.values[0][0] || "") === SAZ.LD_INTERVAL_MODE;
    var turnOn = !isOn;
    return SAZ.ldBands(ctx, sheet).then(function (bands) {
      if (bands.length === 0) {
        return SAZ.notify("Layout not recognised", "The interval-data blocks on " + SAZ.LD_SHEET + " could not be found.");
      }
      var chain = Promise.resolve();
      bands.forEach(function (b) {
        chain = chain.then(function () { return SAZ.setBandState(ctx, sheet, b.first, b.last, turnOn); });
      });
      return chain.then(function () {
        mode.values = [[turnOn ? SAZ.LD_INTERVAL_MODE : SAZ.LD_MANUAL_MODE]];
        return ctx.sync();
      }).then(function () {
        return SAZ.notify(turnOn ? "Interval data added" : "Interval data removed",
          turnOn ? "Data Source is now Raw Interval Data. The interval summary, load duration curve and raw-data paste area are showing - paste the logger export under the Timestamp header."
                 : "Data Source is back to Manual Peak Entry. The interval blocks are collapsed (any pasted data is kept, not deleted).");
      });
    });
  });
};


/* ==========================================================
   v2.6  LOAD CALC SECTIONS BY NAME + MULTIPLE DEPARTMENTS
   The Add / Remove Section picker lists sections by NAME (no numbers)
   with their state, plus "+ Add a new Department section". Each added
   department is a real LC_DeptN table with its own demand factor and its
   own line in every roll-up (panel space counter, service summary,
   generator and UPS breakdowns, Generator / UPS section lists). Formulas
   that enumerate every section (Dashboard reconciliation, largest motor)
   are extended to include the new table.
   ========================================================== */
SAZ.ADD_DEPT_CHOICE = "+ Add a new Department section";
SAZ.DEPT_TABLE_RX = /^LC_Dept\d*$/;
SAZ.DEPT_BASE_TABLE = "LC_Dept";
SAZ.sectionName = function (label) {
  return String(label || "").replace(/^\s*\d+\s+/, "").replace(SAZ.SECTION_OPTIONAL_TAG, "").trim();
};
SAZ.reEsc = function (s) { return String(s).replace(/[.*+?^\${}()|[\]\\]/g, "\\$&"); };

SAZ.uiToggleSection = function (ctx, err) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    if (active.name.toLowerCase() === SAZ.LD_SHEET.toLowerCase()) { return SAZ.uiToggleInterval(ctx); }
    return SAZ.sectionRegistry(ctx).then(function (list) {
      var opt = list.filter(function (e) { return SAZ.SECTION_OPTIONAL_TAG.test(e.label); });
      var base = list.filter(function (e) { return !SAZ.SECTION_OPTIONAL_TAG.test(e.label); })
                     .map(function (e) { return SAZ.sectionName(e.label); });
      var opts = opt.map(function (e) {
        var on = String(e.state).toUpperCase() === "YES";
        return SAZ.sectionName(e.label) + (on ? "   \u2014 included (pick to remove)" : "   \u2014 not added (pick to add)");
      });
      opts.push(SAZ.ADD_DEPT_CHOICE);
      var fields = [
        { id: "sec", label: "Section", type: "select", options: opts, value: opts[0], required: true,
          hint: "Always included: " + base.join(", ") + "." },
        { id: "name", label: "New department name (only for + Add)", type: "text", value: "",
          hint: "e.g. Pharmacy, Bakery, Lab 2. Each department gets its own table, demand factor and subtotal." }
      ];
      return SAZ.form("Add / remove a Load Calc section", "", fields, err || "", 62).then(function (ans) {
        if (!ans || !ans.ok) { return null; }
        var v = ans.values || {};
        if (v.sec === SAZ.ADD_DEPT_CHOICE) {
          var nm = String(v.name || "").trim();
          if (!nm) { return SAZ.uiToggleSection(ctx, "Type a name for the new department."); }
          return SAZ.addDepartment(ctx, nm).then(function (res) {
            if (res.err) { return SAZ.uiToggleSection(ctx, res.err); }
            return SAZ.notify("Department added",
              res.label + " is now its own Load Calc section (table " + res.table + ", rows " + res.first + "-" + res.last +
              ") with its own demand factor, and has a line in the service summary, panel space counter, generator and UPS breakdowns." +
              (res.warn ? "\n\n" + res.warn : ""));
          });
        }
        var hit = opt[opts.indexOf(v.sec)];
        if (!hit) { return SAZ.uiToggleSection(ctx, "Pick a section from the list."); }
        var turnOn = String(hit.state).toUpperCase() !== "YES";
        var name = SAZ.sectionName(hit.label);
        return SAZ.setSectionState(ctx, hit, turnOn).then(function (did) {
          if (!did) { return SAZ.notify("Section not built", "Table \"" + hit.table + "\" does not exist in this workbook."); }
          return SAZ.notify(turnOn ? "Section added" : "Section removed",
            name + (turnOn ? " is now included in the load calc." :
              " is collapsed and excluded from every total. Its rows are kept, so adding it back restores any entries."));
        });
      });
    });
  });
};

/* Clone every aggregate term that reads the base table so the new table is
   counted too. SUMIF/COUNTIF style terms become (a+b), MAXIFS/MINIFS become
   MAX(a,b)/MIN(a,b), bare refs inside VSTACK/HSTACK gain a sibling ref. */
SAZ.cloneTableTerms = function (f, base, nt) {
  var refRx = new RegExp(SAZ.reEsc(base) + "\\[", "g");
  var ph = [];
  var callRx = /\b(SUMIFS?|COUNTIFS?|AVERAGEIFS?|MAXIFS|MINIFS)\(([^()]*)\)/g;
  var out = f.replace(callRx, function (m, fn, args) {
    if (!new RegExp(SAZ.reEsc(base) + "\\[").test(args)) { return m; }
    if (/LC_[A-Za-z0-9]+\[/.test(args.replace(refRx, ""))) { return m; }
    var clone = fn + "(" + args.replace(refRx, nt + "[") + ")";
    var rep = fn === "MAXIFS" ? "MAX(" + m + "," + clone + ")" :
              fn === "MINIFS" ? "MIN(" + m + "," + clone + ")" : "(" + m + "+" + clone + ")";
    ph.push(rep);
    return "\u0001" + (ph.length - 1) + "\u0001";
  });
  if (/\b[VH]STACK\(/.test(out)) {
    out = out.replace(new RegExp(SAZ.reEsc(base) + "\\[([^\\[\\]]+)\\](?=\\s*[,)])", "g"), function (m, col) {
      return m + "," + nt + "[" + col + "]";
    });
  }
  return out.replace(/\u0001(\d+)\u0001/g, function (m, i) { return ph[+i]; });
};

SAZ.extendEnumerations = function (ctx, base, nt) {
  var wss = ctx.workbook.worksheets;
  wss.load("items/name");
  return ctx.sync().then(function () {
    var dev = /^(WEB_|VBA_|XML_|CODE_|PORT_|DESIGN_|Claude Log|Instructions)/;
    var probes = wss.items.filter(function (w) { return !dev.test(w.name); }).map(function (w) {
      var u = w.getUsedRangeOrNullObject(true);
      u.load("formulas,rowIndex,columnIndex");
      return { ws: w, u: u };
    });
    return ctx.sync().then(function () {
      var baseRx = new RegExp(SAZ.reEsc(base) + "\\[");
      var otherRx = new RegExp("LC_(?!" + SAZ.reEsc(base.slice(3)) + "\\[)[A-Za-z0-9]+\\[");
      var chain = Promise.resolve(), count = 0;
      probes.forEach(function (p) {
        if (p.u.isNullObject) { return; }
        var edits = [];
        p.u.formulas.forEach(function (row, i) {
          row.forEach(function (f, j) {
            if (typeof f !== "string" || f.charAt(0) !== "=" || !baseRx.test(f) || !otherRx.test(f)) { return; }
            var g = SAZ.cloneTableTerms(f, base, nt);
            if (g !== f) { edits.push({ r: p.u.rowIndex + i, c: p.u.columnIndex + j, f: g }); }
          });
        });
        if (!edits.length) { return; }
        count += edits.length;
        chain = chain.then(function () {
          return SAZ.withUnprotected(ctx, p.ws, function () {
            edits.forEach(function (e) { p.ws.getCell(e.r, e.c).formulas = [[e.f]]; });
            return ctx.sync();
          });
        });
      });
      return chain.then(function () { return count; });
    });
  });
};

/* Insert one roll-up line for the new table under the last department line of
   every list that carries the base "Department / Owner" line. Formulas are
   copied from the base line and re-pointed (table, include flag, DF cell). */
SAZ.deptRollups = function (ctx, sheet, info) {
  var colA = sheet.getRange("A1:A600"); colA.load("values");
  var used = sheet.getUsedRange(true); used.load("rowIndex,rowCount");
  return ctx.sync().then(function () {
    var start = 0;
    colA.values.forEach(function (r, i) { if (!start && String(r[0]).indexOf("PANEL SPACE COUNTER") === 0) { start = i + 1; } });
    if (!start) { start = info.blockLast + 1; }
    var last = used.rowIndex + used.rowCount;
    var bd = sheet.getRange("B" + start + ":D" + last); bd.load("values");
    return ctx.sync().then(function () {
      var lc = function (v) { return String(v || "").trim().toLowerCase(); };
      var baseLbl = lc(info.baseLabel), deptSet = info.deptLabels.map(lc);
      var hits = [];
      bd.values.forEach(function (r, i) {
        var k = lc(r[0]) === baseLbl ? 0 : (lc(r[2]) === baseLbl ? 2 : -1);
        if (k < 0) { return; }
        var end = i;
        while (end + 1 < bd.values.length && deptSet.indexOf(lc(bd.values[end + 1][k])) >= 0) { end++; }
        hits.push({ base: start + i, after: start + end, col: k === 0 ? "B" : "D" });
      });
      hits.sort(function (a, b) { return b.after - a.after; });       /* bottom-up: inserts never shift pending rows */
      var chain = Promise.resolve();
      hits.forEach(function (h) {
        chain = chain.then(function () {
          var ir = h.after + 1;
          var br = sheet.getRange(h.base + ":" + h.base); br.load("rowHidden");
          return ctx.sync().then(function () {
          sheet.getRange(ir + ":" + ir).insert(Excel.InsertShiftDirection.down);
          sheet.getRange("A" + ir + ":U" + ir).copyFrom(sheet.getRange("A" + h.base + ":U" + h.base), Excel.RangeCopyType.all);
          sheet.getRange(ir + ":" + ir).rowHidden = br.rowHidden;   /* inherit: a line inside a hidden band (panel space counter, undeployed Generator/UPS) stays hidden */
          var rg = sheet.getRange("A" + ir + ":U" + ir); rg.load("formulas");
          return ctx.sync().then(function () {
            var tRx = new RegExp(SAZ.reEsc(info.baseTable) + "\\[", "g");
            var yRx = new RegExp("\\$AC\\$" + info.baseFlag + "(?!\\d)", "g");
            var pRx = new RegExp("\\$P\\$" + info.baseDF + "(?!\\d)", "g");
            rg.formulas = [rg.formulas[0].map(function (f) {
              if (typeof f !== "string" || f.charAt(0) !== "=") { return f; }
              return f.replace(tRx, info.table + "[").replace(yRx, "$AC$" + info.flagRow).replace(pRx, "$P$" + info.dfRow);
            })];
            sheet.getRange(h.col + ir).values = [[info.label]];
            return ctx.sync();
          });
          });
        });
      });
      return chain.then(function () { return hits.length; });
    });
  });
};

/* Generator / UPS "LOADS FROM LOAD CALC" lists must stay row-aligned with the
   Load Calc generator / UPS breakdown, so the same line is inserted there and
   every Connected kVA cell is made position-based (INDEX by row). */
SAZ.deptModuleList = function (ctx, sheetName, rangeName, info) {
  var ws = ctx.workbook.worksheets.getItemOrNullObject(sheetName); ws.load("isNullObject");
  return ctx.sync().then(function () {
    if (ws.isNullObject) { return false; }
    var b = ws.getRange("B1:B60"); b.load("values");
    return ctx.sync().then(function () {
      var lc = function (v) { return String(v || "").trim().toLowerCase(); };
      var hdr = 0, tot = 0, baseR = 0;
      b.values.forEach(function (r, i) {
        var t = lc(r[0]);
        if (!hdr && t === "section") { hdr = i + 1; }
        if (hdr && !tot && t.indexOf("total included") === 0) { tot = i + 1; }
        if (hdr && !tot && t === lc(info.baseLabel)) { baseR = i + 1; }
      });
      if (!hdr || !tot || !baseR) { return false; }
      var end = baseR;
      while (end + 1 < tot && info.deptLabels.map(lc).indexOf(lc(b.values[end][0])) >= 0) { end++; }
      var ir = end + 1;
      return SAZ.withUnprotected(ctx, ws, function () {
        ws.getRange(ir + ":" + ir).insert(Excel.InsertShiftDirection.down);
        ws.getRange("A" + ir + ":F" + ir).copyFrom(ws.getRange("A" + baseR + ":F" + baseR), Excel.RangeCopyType.all);
        ws.getRange(ir + ":" + ir).rowHidden = false;
        ws.getRange("B" + ir).values = [[info.label]];
        ws.getRange("D" + ir).values = [["Y"]];
        /* labelled rows only (never the blank row or the TOTAL row); rows at or
           below the insert point have moved down one */
        for (var r = hdr + 1; r < tot; r++) {
          if (!String(b.values[r - 1][0] || "").trim()) { continue; }
          var rr = r >= ir ? r + 1 : r;
          ws.getRange("C" + rr).formulas = [["=INDEX(" + rangeName + ",ROW()-ROW($B$" + hdr + "))"]];
        }
        ws.getRange("C" + ir).formulas = [["=INDEX(" + rangeName + ",ROW()-ROW($B$" + hdr + "))"]];
        return ctx.sync().then(function () { return true; });
      });
    });
  });
};

SAZ.addDepartment = function (ctx, rawName) {
  var name = String(rawName || "").replace(/[\[\]"'#]/g, "").replace(/\s+/g, " ").trim().slice(0, 40);
  if (!name) { return Promise.resolve({ err: "That name contains no usable characters." }); }
  var label = "Department \u2013 " + name;
  var sheet = ctx.workbook.worksheets.getItem(SAZ.EQUIP_REGISTRY_SHEET);
  return SAZ.sectionRegistry(ctx).then(function (list) {
    var depts = list.filter(function (e) { return SAZ.DEPT_TABLE_RX.test(e.table); });
    var baseE = depts.filter(function (e) { return e.table === SAZ.DEPT_BASE_TABLE; })[0];
    if (!baseE) { return { err: "The base Department / Owner section (" + SAZ.DEPT_BASE_TABLE + ") is not in the section registry." }; }
    var clash = list.some(function (e) {
      var s = SAZ.sectionName(e.label).toLowerCase();
      return s === label.toLowerCase() || s === name.toLowerCase();
    });
    if (clash) { return { err: "A section called \"" + name + "\" already exists." }; }
    var n = 1;
    depts.forEach(function (e) { var m = e.table.match(/^LC_Dept(\d+)$/); if (m) { n = Math.max(n, +m[1]); } });
    var tname = "LC_Dept" + (n + 1);
    var baseT = sheet.tables.getItem(baseE.table), lastT = sheet.tables.getItem(depts[depts.length - 1].table);
    var regRg = sheet.getRange(SAZ.SECTION_REGISTRY_RANGE); regRg.load("values,rowIndex");
    var bHdr = baseT.getHeaderRowRange(); bHdr.load("rowIndex");
    var bBody = baseT.getDataBodyRange(); bBody.load("formulas");
    baseT.load("showFilterButton,showBandedRows");
    return ctx.sync().then(function () {
      var flagRow = 0;
      for (var i = 0; i < regRg.values.length; i++) {
        if (!String(regRg.values[i][1] || "").trim()) { flagRow = regRg.rowIndex + i + 1; break; }
      }
      if (!flagRow) { return { err: "The section registry (" + SAZ.SECTION_REGISTRY_RANGE + ") is full." }; }
      return SAZ.sectionBand(ctx, sheet, baseT).then(function (bb) {
        return SAZ.sectionBand(ctx, sheet, lastT).then(function (lb) {
          var hdrRow = bHdr.rowIndex + 1;
          var src = [bb.banner, bb.banner + 1, hdrRow, hdrRow + 1, bb.last - 1, bb.last, bb.last + 1];
          var at = lb.last + 2;
          sheet.getRange(at + ":" + (at + 6)).insert(Excel.InsertShiftDirection.down);
          return ctx.sync().then(function () {
            var hs = src.map(function (r) { var x = sheet.getRange("A" + r); x.load("height"); return x; });
            src.forEach(function (r, k) {
              sheet.getRange("A" + (at + k) + ":" + SAZ.LC_LAST_COL + (at + k)).copyFrom(sheet.getRange("A" + r + ":" + SAZ.LC_LAST_COL + r), Excel.RangeCopyType.all);
            });
            sheet.getRange(at + ":" + (at + 6)).rowHidden = false;
            return ctx.sync().then(function () {
              hs.forEach(function (x, k) { sheet.getRange((at + k) + ":" + (at + k)).format.rowHeight = x.height; });
              sheet.getRange("A" + (at + 3) + ":" + SAZ.LC_LAST_COL + (at + 3)).clear(Excel.ClearApplyTo.contents);
              var t = sheet.tables.add("A" + (at + 2) + ":" + SAZ.LC_LAST_COL + (at + 3), true);
              t.name = tname;
              t.showBandedRows = baseT.showBandedRows;
              t.showFilterButton = baseT.showFilterButton;
              return ctx.sync().then(function () {
                /* table creation applies a default style - put the section's own look back */
                t.style = null;            /* the built-in sections carry NO table style - a style would bold the headers */
                t.showBandedRows = false;
                return ctx.sync().catch(function () { return null; });
              }).then(function () {
                sheet.getRange("A" + (at + 2) + ":" + SAZ.LC_LAST_COL + (at + 3)).copyFrom(sheet.getRange("A" + hdrRow + ":" + SAZ.LC_LAST_COL + (hdrRow + 1)), Excel.RangeCopyType.formats);
                var row0 = bBody.formulas[0].map(function (f) { return (typeof f === "string" && f.charAt(0) === "=") ? f : ""; });
                t.getDataBodyRange().formulas = [row0];
                t.getHeaderRowRange().format.font.bold = false;   /* table style bolds headers; the sections do not */
                sheet.getRange("A" + at).values = [["DEPARTMENT \u2014 " + name.toUpperCase()]];
                 (function () { var bn = sheet.getRange("A" + at + ":" + SAZ.LC_LAST_COL + (at + 1)); bn.unmerge(); bn.merge(false); bn.format.horizontalAlignment = "Left"; bn.format.verticalAlignment = "Center"; })();   /* banner = 2 rows merged A:X, like every section */
                 sheet.getRange("A" + (at + 2) + ":" + SAZ.LC_LAST_COL + (at + 2)).format.font.bold = false;   /* column headers are never bold */
                sheet.getRange("G" + (at + 5)).values = [[1]];
                sheet.getRange("O" + (at + 5)).formulas = [["=SUM(" + tname + "[Total kVA])"]];
                sheet.getRange("P" + (at + 5)).formulas = [["=ROUND(IFERROR($O" + (at + 5) + "*$G" + (at + 5) + ",0),1)"]];
                var pr = flagRow - 1 >= regRg.rowIndex + 1 ? flagRow - 1 : flagRow;
                sheet.getRange("AA" + flagRow + ":AD" + flagRow).copyFrom(sheet.getRange("AA" + pr + ":AD" + pr), Excel.RangeCopyType.formats);
                sheet.getRange("AA" + flagRow + ":AC" + flagRow).values = [[label + " (optional)", tname, "Yes"]];
                sheet.getRange("AD" + flagRow).formulas = [["=IF($AC" + flagRow + "=\"No\",\"n/a\",SUM(" + tname + "[Total kVA]))"]]; SAZ.lcTrackRegRow(sheet, flagRow, tname);
                return ctx.sync();
              }).then(function () {
                var info = { table: tname, label: label, baseTable: baseE.table, baseLabel: SAZ.sectionName(baseE.label),
                  deptLabels: depts.map(function (e) { return SAZ.sectionName(e.label); }),
                  baseFlag: baseE.flagRow, flagRow: flagRow, baseDF: bb.last, dfRow: at + 5, blockLast: at + 6 };
                return SAZ.deptRollups(ctx, sheet, info).then(function (nRoll) {
                  return SAZ.deptModuleList(ctx, "Generator", "LC_GenByCategory", info).then(function (g) {
                    return SAZ.deptModuleList(ctx, "UPS", "LC_UPSByCategory", info).then(function (u) {
                      return SAZ.extendEnumerations(ctx, baseE.table, tname).then(function (nEnum) {
                        var warn = (nRoll < 4 ? "Only " + nRoll + " of 4 Load Calc roll-up lists were found and extended. " : "") +
                                   (!g ? "Generator section list not updated. " : "") + (!u ? "UPS section list not updated." : "");
                        return { table: tname, label: label, first: at, last: at + 6, rollups: nRoll, enums: nEnum, gen: g, ups: u, warn: warn.trim() };
                      });
                    });
                  });
                });
              });
            });
          });
        });
      });
    });
  });
};



/* ==========================================================
   VIEW TOGGLES  -  sizing columns, load matrix, reconciliation jump.

   Sheet kind is identified by the banner text in A1, exactly as the VBA does:
   "...PANELBOARD..." vs "...SCHEDULE..." (switchboard / MCC). These are the
   only handlers that legitimately hide columns - the ribbon button IS the
   affordance for getting them back, so there is no hidden-data trap.
   ========================================================== */

/* Sheet kind from the schedule banner.
   BUG INHERITED FROM THE VBA AND FIXED HERE: the VBA read Range("A1"), but A1
   is EMPTY on all three templates. The banner lives at B2 - "PANELBOARD
   SCHEDULE", "SWITCHBOARD SCHEDULE", "MOTOR CONTROL CENTER SCHEDULE". That
   test could never match, which silently disabled every handler depending on
   it. Scan a small block rather than pinning one cell, and test PANELBOARD
   before SCHEDULE because "PANELBOARD SCHEDULE" contains both words. */
SAZ.sheetKind = function (ctx, sheetName) {
  var band = ctx.workbook.worksheets.getItem(sheetName).getRange("A1:C4");
  band.load("values");
  return ctx.sync().then(function () {
    var t = "";
    band.values.forEach(function (row) {
      row.forEach(function (cell) { t += " " + String(cell || "").toUpperCase(); });
    });
    if (t.indexOf("PANELBOARD") >= 0) { return "panelboard"; }
    if (t.indexOf("SCHEDULE") >= 0)   { return "schedule"; }
    return "other";
  });
};

/* v2.12: Transformers override columns (Z:AF) live in a collapsed column group - editable,
   out of sight on day one. A protected sheet blocks the outline +/- buttons, so this
   command opens / closes the group with the sheet briefly unprotected. */
SAZ.XFMR_OVR_COLS = "Z:AF";
SAZ.uiToggleXfmrOverrides = function (ctx) {
  var ws = ctx.workbook.worksheets.getItemOrNullObject(SAZ.TX_SHEET); ws.load("isNullObject,visibility");
  return ctx.sync().then(function () {
    if (ws.isNullObject) { return SAZ.notify("No Transformers sheet", "Create it with New Transformers first."); }
    if (ws.visibility !== Excel.SheetVisibility.visible) { ws.visibility = Excel.SheetVisibility.visible; }
    ws.activate();
    var band = ws.getRange(SAZ.XFMR_OVR_COLS); var probe = ws.getRange("Z:Z"); probe.load("columnHidden");
    return ctx.sync().then(function () {
      var show = probe.columnHidden;
      return SAZ.withUnprotected(ctx, ws, function () {
        try { if (show) { band.showGroupDetails(Excel.GroupOption.byColumns); } else { band.hideGroupDetails(Excel.GroupOption.byColumns); } }
        catch (e) { band.columnHidden = !show; }
        return ctx.sync().catch(function () { band.columnHidden = !show; return ctx.sync(); });
      }).then(function () {
        if (show) { ws.getRange("Z6").select(); }
        return ctx.sync();
      });
    });
  });
};

SAZ.uiToggleSizingColumns = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    return SAZ.sheetKind(ctx, active.name).then(function (kind) {
      if (kind === "other") {
        return SAZ.notify("Wrong sheet",
          "Run this from a panelboard, switchboard, or MCC schedule.");
      }
      /* Probe column reflects current state; col Q on a panelboard is left
         alone because it carries the fault labels. */
      /* v2.7: Length Ft / VD % / Wire-Gnd are J:L on all three templates (old bands pointed at the dashboard and engine columns) */
      var probeCol = "J";
      var bands = ["J:L"];
      var probe = active.getRange(probeCol + ":" + probeCol);
      probe.load("columnHidden");
      return ctx.sync().then(function () {
        var hide = !probe.columnHidden;
        return SAZ.withUnprotected(ctx, active, function () { bands.forEach(function (b) { active.getRange(b).columnHidden = hide; }); return ctx.sync(); }).then(function () {
          return SAZ.notify("Sizing columns " + (hide ? "hidden" : "shown"),
            "Length / VD% / wire columns " + (hide ? "hidden" : "restored") +
            " on " + active.name + ".");
        });
      });
    });
  });
};

SAZ.uiToggleLoadMatrix = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    return SAZ.sheetKind(ctx, active.name).then(function (kind) {
      if (kind === "other") {
        return SAZ.notify("Wrong sheet", "Run this from a panelboard, switchboard or MCC schedule.");
      }
      /* v2.7: the load matrix is the Circuit Entry Dashboard - the 9 class columns ending at CKT_KVA_COL (O:W, MCC AF:AN) */
      return SAZ.geo(ctx, active.name).then(function (g) {
        var endN = SAZ.colNum(SAZ.geoUsable(g.CKT_KVA_COL) || "W");
        var band = SAZ.colLet(endN - 8) + ":" + SAZ.colLet(endN);
        var probe = active.getRange(band);
        probe.load("columnHidden");
        return ctx.sync().then(function () {
          var hide = !probe.columnHidden;
          return SAZ.withUnprotected(ctx, active, function () { active.getRange(band).columnHidden = hide; return ctx.sync(); }).then(function () {
            return SAZ.notify("Circuit Entry Dashboard " + (hide ? "hidden" : "shown"),
              "Columns " + band + " " + (hide ? "hidden" : "restored") + " on " + active.name + ".");
          });
        });
      });
    });
  });
};

/* Jump to the Panel Load Reconciliation block on Dashboard. Port of
   Rbn_GoRecon (1979-2004): find the banner in column A, land 5 rows below. */
SAZ.RECON_BANNER = "PANEL LOAD RECONCILIATION";

SAZ.uiGoRecon = function (ctx) {
  var sheet = ctx.workbook.worksheets.getItemOrNullObject(CFG.regSheet);
  sheet.load("isNullObject");
  return ctx.sync().then(function () {
    if (sheet.isNullObject) {
      return SAZ.notify("No Dashboard", "This workbook has no " + CFG.regSheet + " yet.");
    }
    var col = sheet.getRange("A1:A400");
    col.load("values");
    return ctx.sync().then(function () {
      var row = -1;
      for (var i = 0; i < col.values.length; i++) {
        if (String(col.values[i][0] || "").trim().toUpperCase() === SAZ.RECON_BANNER) {
          row = i + 1; break;
        }
      }
      sheet.activate();
      if (row < 0) {
        return ctx.sync().then(function () {
          return SAZ.notify("Section not found",
            "The Panel Load Reconciliation section was not found on " + CFG.regSheet + ".");
        });
      }
      sheet.getRange("A" + (row + 5)).select();
      return ctx.sync();
    });
  });
};


/* ==========================================================
   REFRESH STUDIES  -  port of RefreshStudies (1250-1371) plus
   FindHeaderRow, NoteRowBelow, SizeStudyTable, ProtectStudySheet.

   THE CRITICAL PROPERTY, straight from the VBA's own comment:
   Short Circuit and Voltage Drop are SELF-DERIVING. Every row and every
   value on those two sheets comes from formulas fed by the equipment
   registry, ordered by tier. This routine must NEVER write to them -
   doing so overwrites the spill formulas and breaks the upstream fault
   chain. Arc Flash is still row-driven, so it is the only output.

   The bus list is read from Short Circuit rather than the registry so that
   transformer secondary buses are included and Arc Flash covers exactly the
   same buses as the fault study. Tier 0 is the utility service - the source,
   not our equipment - and is skipped.
   ========================================================== */

SAZ.SC_SHEET = "Short Circuit";
SAZ.AF_SHEET = "Arc Flash";
SAZ.AF_INPUT_COLS = ["A", "D", "F", "H", "I", "Q"];
SAZ.AF_DEFAULT_GAP = 0.05;
SAZ.SC_SCAN_ROWS = 251;

SAZ.findHeaderRow = function (ctx, sheetName, colLetter, headerText, corroborate) {
  /* A study sheet can carry a COLUMN-GROUP BANNER above its real header row
     that repeats one of the header labels. Short Circuit row 3 holds "Node"
     as the group label for the Node columns; row 4 is the true header. A
     single-cell match picks the banner, shifts the whole bus scan up one row,
     and feeds the header row itself in as a phantom bus tagged "Node".
     `corroborate` is a list of { col, text } that must match on the SAME row,
     which no group-banner row satisfies. */
  var sheet = ctx.workbook.worksheets.getItem(sheetName);
  var rg = sheet.getRange(colLetter + "1:" + colLetter + "400");
  rg.load("values");
  var extra = (corroborate || []).map(function (c) {
    var r = sheet.getRange(c.col + "1:" + c.col + "400");
    r.load("values");
    return { rg: r, want: String(c.text).trim().toLowerCase() };
  });
  return ctx.sync().then(function () {
    var want = String(headerText).trim().toLowerCase();
    for (var i = 0; i < rg.values.length; i++) {
      if (String(rg.values[i][0] || "").trim().toLowerCase() !== want) { continue; }
      var ok = true;
      for (var j = 0; j < extra.length; j++) {
        if (String(extra[j].rg.values[i][0] || "").trim().toLowerCase() !== extra[j].want) {
          ok = false; break;
        }
      }
      if (ok) { return i + 1; }
    }
    return 0;
  });
};

/* The templates mark the end of a study table with a note in column A
   beginning "Rows are added". That note is the table's lower fence. */
SAZ.noteRowBelow = function (ctx, sheetName, fromRow) {
  var rg = ctx.workbook.worksheets.getItem(sheetName)
    .getRange("A" + fromRow + ":A" + (fromRow + 1000));
  rg.load("values");
  return ctx.sync().then(function () {
    for (var i = 0; i < rg.values.length; i++) {
      if (String(rg.values[i][0] || "").indexOf("Rows are added") === 0) {
        /* v2.8.1: the note sits under a blank, unformatted spacer row - the
           table ends at the spacer, so the spacer is never counted as a row */
        if (i >= 2 && String(rg.values[i - 1][0] || "").trim() === "") { return fromRow + i - 1; }
        return fromRow + i;
      }
    }
    return 0;
  });
};

/* Grow or shrink a study table to exactly `needed` rows. Growing clones the
   first row so column formats and formula columns come with it, then clears
   ONLY the named input columns - an explicit list, not a guess at what is a
   constant, so formula columns are never at risk. */
SAZ.sizeStudyTable = function (ctx, sheetName, firstRow, needed, clearCols) {
  if (needed < 1) { needed = 1; }
  var sheet = ctx.workbook.worksheets.getItem(sheetName);
  return SAZ.noteRowBelow(ctx, sheetName, firstRow).then(function (noteRow) {
    var have = noteRow === 0 ? 1 : noteRow - firstRow;
    if (have < 1) { have = 1; }
    if (needed === have) { return { have: have, needed: needed, delta: 0 }; }
    if (needed > have) {
      var addN = needed - have;
      sheet.getRange((firstRow + 1) + ":" + (firstRow + addN))
        .insert(Excel.InsertShiftDirection.down);
      return ctx.sync().then(function () {
        sheet.getRange((firstRow + 1) + ":" + (firstRow + addN))
          .copyFrom(sheet.getRange(firstRow + ":" + firstRow), Excel.RangeCopyType.all);
        return ctx.sync();
      }).then(function () {
        (clearCols || []).forEach(function (col) {
          sheet.getRange(col + (firstRow + 1) + ":" + col + (firstRow + addN))
            .clear(Excel.ClearApplyTo.contents);
        });
        return ctx.sync().then(function () {
          return { have: have, needed: needed, delta: addN };
        });
      });
    }
    sheet.getRange((firstRow + needed) + ":" + (firstRow + have - 1))
      .delete(Excel.DeleteShiftDirection.up);
    return ctx.sync().then(function () {
      return { have: have, needed: needed, delta: needed - have };
    });
  });
};

SAZ.protectStudySheet = function (ctx, sheetName) {
  var sheet = ctx.workbook.worksheets.getItem(sheetName);
  sheet.protection.protect(SAZ.PROTECT_OPTS);
  return ctx.sync().catch(function () { return null; });   /* already protected */
};


/* Arc Flash equipment class defaults, applied only to buses that are new to
   the table. An existing row's user input always wins. */
SAZ.afDefaultClass = function (busType) {
  switch (String(busType || "").trim()) {
    case "Switchboard":
    case "Switchgear":   return "LV Switchgear";
    case "MCC":          return "MCC";
    case "Transformer":  return "Cable/Junction";
    default:             return "Panelboard";
  }
};

/* Default working distance (in) by equipment class - IEEE 1584-2002 Table 3
   typicals: LV switchgear 24, MCC / panelboard / cable 18, >1 kV 36. */
SAZ.afWorkDistFormula = function (r) {
  return "=IF($D" + r + "=\"\",\"\",IF(N($B" + r + ")>1000,36,IF($D" + r + "=\"LV Switchgear\",24,18)))";
};

SAZ.refreshStudies = function (ctx) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name,items/visibility");
  return ctx.sync().then(function () {
    var byName = {};
    sheets.items.forEach(function (s) { byName[s.name.toLowerCase()] = s; });
    var sc = byName[SAZ.SC_SHEET.toLowerCase()];
    var af = byName[SAZ.AF_SHEET.toLowerCase()];
    if (!sc) {
      return SAZ.notify("No Short Circuit sheet",
        "The bus list is derived on \"" + SAZ.SC_SHEET + "\", which this workbook" +
        " does not have. Create it with New Short Circuit first.");
    }
    if (!af) {
      /* DEVIATION: the VBA calls CreateModule to conjure Arc Flash mid-refresh.
         Creating a whole module as a side effect of a refresh is a surprise the
         user did not ask for, so this refuses and says what to do instead. */
      return SAZ.notify("No Arc Flash sheet",
        "Arc Flash is the output of this refresh but does not exist yet." +
        " Create it with New Arc Flash, then run Refresh Studies again.");
    }
    if (af.visibility !== Excel.SheetVisibility.visible) {
      af.visibility = Excel.SheetVisibility.visible;
    }
    /* let the self-deriving sheets settle before their values are read */
    ctx.application.calculate(Excel.CalculationType.full);
    return ctx.sync().then(function () {
      return SAZ.findHeaderRow(ctx, SAZ.SC_SHEET, "A", "Node",
            [{ col: "B", text: "Type" }, { col: "C", text: "Tier" }]).then(function (scHdr) {
        return SAZ.findHeaderRow(ctx, SAZ.AF_SHEET, "C", "Ibf (kA)",
              [{ col: "A", text: "Node" }, { col: "B", text: "Volts" }]).then(function (afHdr) {
          if (scHdr === 0 || afHdr === 0) {
            return SAZ.notify("Study layout not recognised",
              "The Short Circuit bus list or the Arc Flash table header could not" +
              " be found. Those sheets may have been altered.");
          }
          var afFirst = afHdr + 1;
          /* READ ONLY from Short Circuit - never write to it */
          var busRg = ctx.workbook.worksheets.getItem(SAZ.SC_SHEET)
            .getRange("A" + (scHdr + 1) + ":C" + (scHdr + SAZ.SC_SCAN_ROWS));
          busRg.load("values");
          return ctx.sync().then(function () {
            var buses = [];
            busRg.values.forEach(function (row) {
              var tag = String(row[0] || "").trim();
              if (!tag) { return; }
              var tier = row[2];
              if (typeof tier === "number" && tier === 0) { return; }  /* utility source */
              buses.push({ tag: tag, type: String(row[1] || "").trim() });
            });
            var afSheet = ctx.workbook.worksheets.getItem(SAZ.AF_SHEET);
            afSheet.protection.unprotect();
            return ctx.sync().catch(function () { return null; }).then(function () {
              return SAZ.noteRowBelow(ctx, SAZ.AF_SHEET, afFirst).then(function (nr) {
                var afLast = nr === 0 ? afFirst + 1 : nr;
                /* preserve existing user inputs, keyed by node tag */
                var keep = afSheet.getRange("A" + afFirst + ":Q" + (afLast - 1));
                keep.load("values,formulas");
                return ctx.sync().then(function () {
                  var saved = {};
                  keep.values.forEach(function (row, k) {
                    var t = String(row[0] || "").trim();
                    if (t && !saved[t]) {
                      /* G (working distance) is a DEFAULT FORMULA by equipment class
                         unless the user typed a number over it - keep only overrides */
                      var gF = String(keep.formulas[k][6] || "");
                      saved[t] = { D: row[3], F: row[5], H: row[7], I: row[8], Q: row[16],
                                   G: gF.charAt(0) === "=" ? null : row[6] };
                    }
                  });
                  return SAZ.sizeStudyTable(ctx, SAZ.AF_SHEET, afFirst,
                    buses.length, SAZ.AF_INPUT_COLS).then(function () {
                    if (buses.length === 0) {
                      SAZ.AF_INPUT_COLS.forEach(function (col) {
                        afSheet.getRange(col + afFirst).clear(Excel.ClearApplyTo.contents);
                      });
                      return ctx.sync();
                    }
                    buses.forEach(function (b, i) {
                      var r = afFirst + i;
                      afSheet.getRange("A" + r).values = [[b.tag]];
                      var s = saved[b.tag];
                      if (s && s.G !== null && s.G !== "" && s.G !== undefined) {
                        afSheet.getRange("G" + r).values = [[s.G]];
                      } else {
                        afSheet.getRange("G" + r).formulas = [[SAZ.afWorkDistFormula(r)]];
                      }
                      if (s) {
                        afSheet.getRange("D" + r).values = [[s.D]];
                        afSheet.getRange("F" + r).values = [[s.F]];
                        afSheet.getRange("H" + r).values = [[s.H]];
                        afSheet.getRange("I" + r).values = [[s.I]];
                        afSheet.getRange("Q" + r).values = [[s.Q]];
                      } else {
                        afSheet.getRange("D" + r).values = [[SAZ.afDefaultClass(b.type)]];
                        afSheet.getRange("H" + r).values = [[SAZ.AF_DEFAULT_GAP]];
                      }
                    });
                    return ctx.sync();
                  }).then(function () {
                    return SAZ.protectStudySheet(ctx, SAZ.AF_SHEET);
                  }).then(function () {
                    afSheet.activate();
                    return ctx.sync();
                  }).then(function () {
                    var reused = buses.filter(function (b) { return !!saved[b.tag]; }).length;
                    return SAZ.notify("Studies refreshed",
                      "Arc Flash rebuilt for " + buses.length + " bus" +
                      (buses.length === 1 ? "" : "es") + ".\n" +
                      reused + " kept existing inputs, " + (buses.length - reused) +
                      " seeded with defaults.\n\n" +
                      "Short Circuit and Voltage Drop derive themselves from the" +
                      " equipment registry and were not modified.");
                  });
                });
              });
            }).catch(function (err) {
              /* never leave Arc Flash unprotected if the rebuild fails part-way */
              return SAZ.protectStudySheet(ctx, SAZ.AF_SHEET).then(function () { throw err; });
            });
          });
        });
      });
    });
  });
};


/* ==========================================================
   LOAD-BEARING SHEET NAMES

   Every TPL_ sheet carries an AIC check that reaches into 'Short Circuit'
   BY NAME (XLOOKUP over 'Short Circuit'!A6:Q255). Clone a panel into a
   workbook that has no such tab and the lookup cannot resolve - and because
   the formula wraps it in IFERROR, it does not shout. It quietly drops the
   fault-study cross-check and the AIC verdict is then based on the local
   calc alone. A silently weakened safety check is worse than a loud failure,
   so creation is blocked up front instead.
   ========================================================== */

SAZ.CLONE_REQUIRES = ["Short Circuit"];

SAZ.requireSheets = function (ctx, names) {
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  return ctx.sync().then(function () {
    var have = sheets.items.map(function (s) { return s.name.toLowerCase(); });
    var missing = names.filter(function (n) { return have.indexOf(n.toLowerCase()) < 0; });
    if (missing.length === 0) { return true; }
    return SAZ.notify("Missing required sheet",
      "This workbook has no \"" + missing.join("\", \"") + "\" tab.\n\n" +
      "Equipment sheets look that name up for their AIC check, so a board" +
      " created now would show an incomplete verdict. Create it first, and do" +
      " not rename it afterwards.").then(function () { return false; });
  });
};


/* Transformer default %Z, read from the workbook rather than baked into code.
   Xfmr_DefZ already exists on NEC_Data and is what the sheet formulas use,
   is a blue-on-grey input cell. Falls back to the constant only if the name
   has been deleted, so an older workbook still works. */

/* ==========================================================
   LOAD CALC -> PANEL  (v2.6, phase 2)
   Send to Panel places the selected Load Calc rows on a schedule as LIVE
   links (description, kVA and HP follow the Load Calc row). Placement is
   tracked on every equipment section in three command-managed columns:
     Ckts   - where the load landed, e.g. "LP1-5,7; LP2-2"
     Placed - units placed (dedicated loads: one per circuit; Ded. Ckt = N: 1)
     Status - Placed / Partial / Unplaced  (live formula, filterable)
   Refresh Placement rescans every schedule for circuits linked to Load Calc,
   so clearing or moving a circuit on a panel is picked up.
   ========================================================== */
SAZ.LC_SHEET = "Load Calc";
SAZ.LC_LAST_COL = "X";
SAZ.LC_TRACK = ["Ckts", "Placed", "Status"];
SAZ.LC_STATUS_F = '=IF([@Description]="","",IF(N([@Placed])=0,"Unplaced",IF([@Placed]>=IF([@[Ded. Ckt]]="N",1,MAX(1,N([@Qty]))),"Placed","Partial")))';
SAZ.CLASS_CODES = ["LTG", "REC", "MTR", "HVC", "HEAT", "KIT", "CONT", "GEN", "PNL"];

SAZ.colLetter = function (n) { var s = ""; n = n + 1; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

/* Equipment tables on Load Calc = every table with a Panel Tag column. */
SAZ.lcEquipTables = function (ctx) {
  var sheet = ctx.workbook.worksheets.getItem(SAZ.LC_SHEET);
  var ts = sheet.tables; ts.load("items/name");
  return ctx.sync().then(function () {
    var hs = ts.items.map(function (t) { var h = t.getHeaderRowRange(); h.load("values,rowIndex,columnIndex"); var b = t.getDataBodyRange(); b.load("rowIndex,rowCount"); return { t: t, h: h, b: b }; });
    return ctx.sync().then(function () {
      return hs.filter(function (x) { return x.h.values[0].indexOf("Panel Tag") >= 0; }).map(function (x) {
        var cols = {}; x.h.values[0].forEach(function (v, j) { cols[String(v)] = SAZ.colLetter(x.h.columnIndex + j); });
        return { name: x.t.name, table: x.t, cols: cols, hdrRow: x.h.rowIndex + 1, first: x.b.rowIndex + 1, last: x.b.rowIndex + x.b.rowCount, headers: x.h.values[0] };
      });
    });
  });
};

/* Add Ckts / Placed / Status to any equipment section that lacks them (older
   workbooks heal on first use). Columns are command-managed: locked, red. */
SAZ.lcTrackRegRow = function (sheet, r, t) {
  var g = function (st, what) { return what === "n" ? 'COUNTIFS(' + t + '[Status],"' + st + '")' : 'SUMIFS(' + t + '[Total kVA],' + t + '[Status],"' + st + '")'; };
  sheet.getRange("AG" + r + ":AH" + r).formulas = [[
    '=IF($AC' + r + '="No",0,' + g("Unplaced", "n") + '+' + g("Partial", "n") + ')',
    '=IF($AC' + r + '="No",0,' + g("Unplaced", "k") + '+' + g("Partial", "k") + ')']];
};
/* Registry columns AG:AH = loads / kVA not fully placed per section (Dashboard check 14). */
SAZ.lcTrackRegistry = function (ctx, list) {
  var sheet = ctx.workbook.worksheets.getItem(SAZ.LC_SHEET);
  var rg = sheet.getRange(SAZ.SECTION_REGISTRY_RANGE); rg.load("values,rowIndex");
  var cur = sheet.getRange("AG7:AH40"); cur.load("formulas");
  return ctx.sync().then(function () {
    var names = list.filter(function (e) { return e.cols.Status; }).map(function (e) { return e.name; });
    var hdr = sheet.getRange("AG6:AH6"); hdr.values = [["Not placed (loads)", "Not placed (kVA)"]];
    hdr.copyFrom(sheet.getRange("AD6"), Excel.RangeCopyType.formats);
    rg.values.forEach(function (v, i) {
      var r = rg.rowIndex + i + 1, t = String(v[1] || "").trim();
      if (names.indexOf(t) < 0) { return; }
      if (String(cur.formulas[r - 7][0]).indexOf(t + "[Status]") >= 0) { return; }
      SAZ.lcTrackRegRow(sheet, r, t);
      sheet.getRange("AG" + r + ":AH" + r).copyFrom(sheet.getRange("AD" + r), Excel.RangeCopyType.formats);
    });
    sheet.getRange("AH7:AH40").numberFormat = Array.apply(null, Array(34)).map(function () { return ["#,##0.00"]; });
    return ctx.sync();
  });
};

SAZ.lcEnsureTrack = function (ctx) {
  var sheet = ctx.workbook.worksheets.getItem(SAZ.LC_SHEET);
  return SAZ.lcEquipTables(ctx).then(function (list) {
    var todo = list.filter(function (e) { return !e.cols.Status; });
    if (!todo.length) { return SAZ.withUnprotected(ctx, sheet, function () { return SAZ.lcTrackRegistry(ctx, list); }).then(function () { return list; }); }
    return SAZ.withUnprotected(ctx, sheet, function () {
      todo.forEach(function (e) {
        SAZ.LC_TRACK.forEach(function (nm) { if (!e.cols[nm]) { e.table.columns.add(null, null, nm); } });
      });
      return ctx.sync().then(function () { return SAZ.lcEquipTables(ctx); }).then(function (l2) {
        l2.forEach(function (e) {
          if (!todo.some(function (x) { return x.name === e.name; })) { return; }
          var n = e.cols.Notes, c = e.cols.Ckts, s = e.cols.Status, k = e.cols["kVA Each"];
          sheet.getRange(c + e.hdrRow + ":" + s + e.hdrRow).copyFrom(sheet.getRange(n + e.hdrRow), Excel.RangeCopyType.formats);
          var body = sheet.getRange(c + e.first + ":" + s + e.last);
          body.copyFrom(sheet.getRange(k + e.first + ":" + k + e.last), Excel.RangeCopyType.formats);
          body.format.protection.locked = true;
          body.format.horizontalAlignment = "Center";
          e.table.columns.getItem("Status").getDataBodyRange().formulas =
            Array.apply(null, Array(e.last - e.first + 1)).map(function () { return [SAZ.LC_STATUS_F]; });
          e.table.columns.getItem("Placed").getDataBodyRange().numberFormat =
            Array.apply(null, Array(e.last - e.first + 1)).map(function () { return ["0;-0;"]; });
        });
        sheet.getRange("V:V").format.columnWidth = 110;
        sheet.getRange("W:W").format.columnWidth = 60;
        sheet.getRange("X:X").format.columnWidth = 66;
        return ctx.sync().then(function () { return SAZ.lcStatusFormats(ctx, sheet, l2); }).then(function () { return SAZ.lcTrackRegistry(ctx, l2); }).then(function () { return l2; });
      });
    });
  });
};

/* Status colours: Unplaced amber, Partial orange, Placed green (text only). */
SAZ.lcStatusFormats = function (ctx, sheet, list) {
  list.forEach(function (e) {
    if (!e.cols.Status) { return; }
    var rg = sheet.getRange(e.cols.Status + e.first + ":" + e.cols.Status + e.last);
    rg.conditionalFormats.clearAll();
    [["Unplaced", "#C55A11"], ["Partial", "#B8860B"], ["Placed", "#2E7D32"]].forEach(function (p) {
      var cf = rg.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
      cf.cellValue.format.font.color = p[1];
      cf.cellValue.rule = { formula1: '="' + p[0] + '"', operator: "EqualTo" };
    });
  });
  return ctx.sync().catch(function () { return null; });
};

/* Records for the Load Calc rows under the selection. */
SAZ.lcSelectedLoads = function (ctx) {
  var sheet = ctx.workbook.worksheets.getItem(SAZ.LC_SHEET);
  var sel = ctx.workbook.getSelectedRange(); sel.load("rowIndex,rowCount,worksheet/name");
  return ctx.sync().then(function () { return SAZ.lcEnsureTrack(ctx); }).then(function (list) {
    var r0 = sel.rowIndex + 1, r1 = sel.rowIndex + sel.rowCount;
    var hits = list.filter(function (e) { return e.last >= r0 && e.first <= r1; });
    var reads = hits.map(function (e) { var a = Math.max(r0, e.first), b = Math.min(r1, e.last); var rg = sheet.getRange("A" + a + ":" + SAZ.LC_LAST_COL + b); rg.load("values"); return { e: e, a: a, rg: rg }; });
    return ctx.sync().then(function () {
      var out = [];
      reads.forEach(function (x) {
        var idx = function (h) { return x.e.headers.indexOf(h); };
        x.rg.values.forEach(function (v, i) {
          var g = function (h) { var j = idx(h); return j < 0 ? "" : v[j]; };
          if (String(g("Description")).trim() === "") { return; }
          out.push({ table: x.e.name, cols: x.e.cols, row: x.a + i, desc: SAZ.lcLabel(g("No."), g("Description")), no: String(g("No.") || "").trim(), qty: Number(g("Qty")) || 1,
            ded: String(g("Ded. Ckt")).trim().toUpperCase(), unit: String(g("Unit")).trim(), volts: Number(g("Volts")) || 0,
            ph: Number(g("Ph")) || 0, recept: String(g("Recept Type")).trim(), placed: Number(g("Placed")) || 0,
            kvaEach: Number(g("kVA Each")) || 0, panelTag: String(g("Panel Tag")).trim() });
        });
      });
      return out;
    });
  });
};

/* v2.12: panel label for a Load Calc row - "No. - Description" (e.g. "RTU-1 - ROOFTOP UNIT"), or the
   description alone when No. is blank. descFormula is the live-link twin written to the schedule. */
SAZ.lcLabel = function (no, desc) { var n = String(no || "").trim(), d = String(desc || "").trim(); return n ? n + " - " + d : d; };
SAZ.lcDescFormula = function (rec, tagN) {
  var R = "'" + SAZ.LC_SHEET + "'!$", d = R + rec.cols.Description + "$" + rec.row, sfx = tagN ? "&\" #" + tagN + "\"" : "";
  if (!rec.cols["No."]) { return "=" + d + sfx; }
  var n = R + rec.cols["No."] + "$" + rec.row;
  return "=IF(TRIM(" + n + ")=\"\"," + d + ",TRIM(" + n + ")&\" - \"&" + d + ")" + sfx;
};
SAZ.lcUnits = function (rec) { return rec.ded === "N" ? 1 : Math.max(1, rec.qty); };

/* Classification guess from the section and the load. */
SAZ.lcClassFor = function (rec) {
  var d = rec.desc.toLowerCase(), t = rec.table;
  if (/heat|heater|ehc|uh-|duct heater/.test(d)) { return "HEAT"; }
  if (rec.recept) { return "REC"; }
  if (/light|ltg|lighting|sign/.test(d)) { return "LTG"; }
  if (t === "LC_Kitchen") { return "KIT"; }
  if (t === "LC_Mech") { return "HVC"; }
  if (rec.unit.toUpperCase() === "HP" || t === "LC_Refrig" || t === "LC_Elev" || t === "LC_Plumb") { return "MTR"; }
  if (t === "LC_EV" || t === "LC_Data") { return "CONT"; }
  return "GEN";
};

/* Breaker poles from the load's phase / voltage. */
SAZ.lcPoles = function (rec) {
  if (rec.ph === 3) { return 3; }
  return (rec.volts > 0 && rec.volts <= 277) ? 1 : 2;
};

/* Can a load of 'v' volts be served by a board of line voltage 'vll'? */
SAZ.lcVoltOk = function (v, vll, ph) {
  if (!v || !vll) { return true; }
  var near = function (a, b) { return Math.abs(a - b) / b <= 0.06; };
  if (near(v, vll)) { return true; }
  if (Number(ph) !== 3 && near(v, vll / 2)) { return true; }             /* 240/120 1-phase */
  if (near(v, vll / Math.sqrt(3))) { return true; }                      /* L-N of a wye board */
  return false;
};

/* Class-entry column on a board for a classification code. */
SAZ.boardClassCols = function (ctx, tag, first) {
  var sh = ctx.workbook.worksheets.getItem(tag);
  var h = sh.getRange("A" + (first - 1) + ":CZ" + (first - 1)); h.load("values");
  return ctx.sync().then(function () {
    var m = {};
    h.values[0].forEach(function (v, j) { var k = String(v || "").trim().toUpperCase(); if (SAZ.CLASS_CODES.indexOf(k) >= 0 && !m[k]) { m[k] = SAZ.colLetter(j); } });
    var hp = h.values[0].findIndex(function (v) { return String(v || "").trim().toUpperCase() === "HP"; });
    m._HP = hp >= 0 ? SAZ.colLetter(hp) : "";
    return m;
  });
};
SAZ.CKT_LEN_COL = "J";   /* Length Ft on all three schedule templates (manual entry, drives branch VD %) */

/* Place loads. recs: from lcSelectedLoads; opt: { panel, start, poles(0=auto), cls(""=auto), units(0=remaining) }.
   Returns { placed:[{desc,label}], skipped:[msg] }. */
SAZ.sendToPanel = function (ctx, recs, opt) {
  var P = opt.panel;
  var sheet = ctx.workbook.worksheets.getItem(P);
  var parts = [];
  return SAZ.cktGrid(ctx, P).then(function (grid) { parts[0] = grid; return SAZ.boardClassCols(ctx, P, grid.first); })
  .then(function (cc) { parts[1] = cc; return SAZ.geoCellValue(ctx, P, "ENG_VLL_CELL"); })
  .then(function (vll) {
    var grid = parts[0], cc = parts[1], done = [], skipped = [];
    var nums = Object.keys(grid.byNum).map(Number).sort(function (a, b) { return a - b; });
    var cursor = Number(opt.start) || 0;
    var nextFit = function (poles) {
      var startAt = cursor ? nums.indexOf(cursor) : 0; if (startAt < 0) { startAt = 0; }
      for (var i = startAt; i < nums.length; i++) { var c = SAZ.slotCheck(grid, nums[i], poles, []); if (c.ok) { return c; } }
      return null;
    };
    var occupy = function (chk, desc) {
      chk.nums.forEach(function (n, k) { var r = grid.byNum[n]; if (k === 0) { r.desc = desc; } else { r.owner = chk.row; } });
    };
    return SAZ.withUnprotected(ctx, sheet, function () {
      recs.forEach(function (rec, ri) {
        if (!SAZ.lcVoltOk(rec.volts, vll, rec.ph)) { skipped.push(rec.desc + ": " + rec.volts + " V load cannot go on " + P + " (" + vll + " V)."); return; }
        var poles = Number(opt.poles) || SAZ.lcPoles(rec);
        var cls = (opt.cls || SAZ.lcClassFor(rec)).toUpperCase();
        var col = cc[cls] || cc.GEN || SAZ.geoUsable(grid.g.CKT_KVA_COL);
        var units = SAZ.lcUnits(rec), left = Math.max(0, units - rec.placed);
        var n = (recs.length === 1 && Number(opt.units)) ? Math.min(Number(opt.units), left) : left;
        if (n <= 0) { skipped.push(rec.desc + ": already fully placed (" + rec.placed + " of " + units + ")."); return; }
        var kvaCell = rec.ded === "N" ? rec.cols["Total kVA"] : rec.cols["kVA Each"];
        for (var u = 0; u < n; u++) {
          var chk;
          if (ri === 0 && u === 0 && opt.start) { chk = SAZ.slotCheck(grid, Number(opt.start), poles, []); if (!chk.ok) { throw new Error(chk.msg); } }
          else { chk = nextFit(poles); }
          if (!chk) { skipped.push(rec.desc + ": no open " + poles + "-pole position left on " + P + " (" + (n - u) + " unit(s) not placed)."); break; }
          var tagN = units > 1 && rec.ded !== "N" ? (rec.placed + u + 1) : 0;
          var dF = SAZ.lcDescFormula(rec, tagN);   /* v2.12: No. - Description */
          sheet.getRange(grid.cols.dc + chk.row).formulas = [[dF]];
          if (grid.cols.pc) { sheet.getRange(grid.cols.pc + chk.row).values = [[poles]]; }
          SAZ.CLASS_CODES.forEach(function (k) { if (cc[k]) { sheet.getRange(cc[k] + chk.row).values = [[""]]; } });
          sheet.getRange(col + chk.row).formulas = [["='" + SAZ.LC_SHEET + "'!$" + kvaCell + "$" + rec.row]];
          if (opt.len > 0 && SAZ.CKT_LEN_COL) { sheet.getRange(SAZ.CKT_LEN_COL + chk.row).values = [[opt.len]]; }
          if (cc._HP && rec.unit.toUpperCase() === "HP") { sheet.getRange(cc._HP + chk.row).formulas = [["='" + SAZ.LC_SHEET + "'!$" + rec.cols["Load Value"] + "$" + rec.row]]; }
          occupy(chk, rec.desc + (tagN ? " #" + tagN : ""));
          cursor = chk.nums[chk.nums.length - 1];
          done.push({ desc: rec.desc + (tagN ? " #" + tagN : ""), label: P + "-" + chk.label });
        }
      });
      return ctx.sync();
    }).then(function () { return SAZ.lcRefreshPlacement(ctx); }).then(function () { return { placed: done, skipped: skipped }; });
  });
};

/* Rescan every schedule for circuits linked to Load Calc rows and restamp
   Ckts / Placed (and Panel Tag when it is blank). */
SAZ.lcRefreshPlacement = function (ctx) {
  var lcs = ctx.workbook.worksheets.getItem(SAZ.LC_SHEET);
  var rx = new RegExp("'" + SAZ.reEsc(SAZ.LC_SHEET) + "'!\\$?[A-Z]{1,2}\\$?(\\d+)");
  var hits = {};   /* lc row -> [{tag,num}] */
  return SAZ.registryRecords(ctx).then(function (reg) {
    var ws = ctx.workbook.worksheets; ws.load("items/name");
    return ctx.sync().then(function () {
      var names = ws.items.map(function (s) { return s.name; });
      var tags = reg.map(function (r) { return r.tag; }).filter(function (t) { return names.indexOf(t) >= 0; });
      var chain = Promise.resolve();
      tags.forEach(function (t) {
        chain = chain.then(function () {
          return SAZ.cktGrid(ctx, t).then(function (grid) {
            var sh = ctx.workbook.worksheets.getItem(t);
            var rD = sh.getRange(grid.cols.dc + grid.first + ":" + grid.cols.dc + grid.last); rD.load("formulas");
            return ctx.sync().then(function () {
              rD.formulas.forEach(function (f, i) {
                var m = String(f[0] || "").match(rx);
                if (m) { var lr = Number(m[1]); (hits[lr] = hits[lr] || []).push({ tag: t, num: grid.rows[i].num }); }
              });
            });
          }, function () { return null; });
        });
      });
      return chain;
    });
  }).then(function () { return SAZ.lcEquipTables(ctx); }).then(function (list) {
    return SAZ.withUnprotected(ctx, lcs, function () {
      var reads = list.filter(function (e) { return e.cols.Ckts && e.last >= e.first; }).map(function (e) {
        var pt = lcs.getRange(e.cols["Panel Tag"] + e.first + ":" + e.cols["Panel Tag"] + e.last); pt.load("values");
        return { e: e, pt: pt };
      });
      return ctx.sync().then(function () {
        reads.forEach(function (x) {
          var e = x.e, ck = [], pl = [], tg = [], changed = false;
          for (var r = e.first; r <= e.last; r++) {
            var h = hits[r] || [];
            var by = {}; h.forEach(function (z) { (by[z.tag] = by[z.tag] || []).push(z.num); });
            ck.push([Object.keys(by).map(function (k) { return k + "-" + by[k].sort(function (a, b) { return a - b; }).join(","); }).join("; ")]);
            pl.push([h.length]);
            var cur = String(x.pt.values[r - e.first][0] || "").trim();
            if (!cur && h.length) { cur = h[0].tag; changed = true; }
            tg.push([cur]);
          }
          lcs.getRange(e.cols.Ckts + e.first + ":" + e.cols.Ckts + e.last).values = ck;
          lcs.getRange(e.cols.Placed + e.first + ":" + e.cols.Placed + e.last).values = pl;
          if (changed) { lcs.getRange(e.cols["Panel Tag"] + e.first + ":" + e.cols["Panel Tag"] + e.last).values = tg; }
        });
        return ctx.sync();
      });
    });
  });
};

/* Ribbon: Send to Panel. Select one or more load rows on Load Calc first. */
SAZ.uiSendToPanel = function (ctx, keep, err) {
  var active = ctx.workbook.worksheets.getActiveWorksheet(); active.load("name");
  return ctx.sync().then(function () {
    if (active.name !== SAZ.LC_SHEET) { return SAZ.notify("Select loads on Load Calc", "Go to Load Calc, select the load row(s) to place (any cell in each row), then run Send to Panel."); }
    return SAZ.lcSelectedLoads(ctx).then(function (recs) {
      if (!recs.length) { return SAZ.notify("No loads selected", "Select one or more filled-in rows inside a Load Calc equipment section, then run Send to Panel."); }
      return SAZ.registryRecords(ctx).then(function (reg) {
        var tags = reg.filter(function (r) { return SAZ.isBoardType(r.type); }).map(function (r) { return r.tag; });
        if (!tags.length) { return SAZ.notify("No schedules", "Create a panelboard, switchboard or MCC first."); }
        keep = keep || {};
        var p0 = keep.panel || recs[0].panelTag; if (tags.indexOf(p0) < 0) { p0 = tags[0]; }
        var one = recs.length === 1, rec = recs[0];
        var left = Math.max(0, SAZ.lcUnits(rec) - rec.placed);
        var text = one ? "\"" + rec.desc + "\" - " + SAZ.lcUnits(rec) + " unit(s), " + rec.placed + " placed." :
          recs.length + " loads selected. Each dedicated unit gets its own circuit, filled from the start circuit upward.";
        var fields = [
          { id: "panel", label: "Panel / board", type: "select", options: tags, value: p0, required: true },
          { id: "start", label: "Start at circuit (blank = first open)", type: "number", value: keep.start || "" },
          { id: "poles", label: "Breaker poles", type: "select", options: ["Auto", "1", "2", "3"], value: keep.poles || "Auto", hint: "Auto: 3-pole for 3-phase loads, 1-pole up to 277 V, else 2-pole." },
          { id: "cls", label: "Classification", type: "select", options: ["Auto"].concat(SAZ.CLASS_CODES), value: keep.cls || (one ? SAZ.lcClassFor(rec) : "Auto") },
          { id: "len", label: "Branch circuit length (ft, optional)", type: "number", value: keep.len || "", hint: "Written to Length Ft on each circuit placed, for voltage drop. It stays a manual entry on the schedule - edit it there any time." } ];
        if (one && rec.ded !== "N" && SAZ.lcUnits(rec) > 1) { fields.push({ id: "units", label: "Units to place now", type: "number", value: keep.units || String(left || 1), min: 1, max: Math.max(1, left) }); }
        return SAZ.form("Send to Panel", text, fields, err || "", 70).then(function (ans) {
          if (!ans.ok) { return null; }
          var v = ans.values;
          var o = { panel: v.panel, start: parseInt(v.start, 10) || 0, poles: v.poles === "Auto" ? 0 : parseInt(v.poles, 10), cls: v.cls === "Auto" ? "" : v.cls, units: parseInt(v.units, 10) || 0, len: Number(v.len) > 0 ? Number(v.len) : 0 };
          return SAZ.sendToPanel(ctx, recs, o).then(function (res) {
            var msg = (res.placed.length ? res.placed.length + " circuit(s) written:\n" + res.placed.map(function (p) { return "\u2022 " + p.desc + "  \u2192  " + p.label; }).join("\n") : "Nothing was placed.") +
              (res.skipped.length ? "\n\nNot placed:\n" + res.skipped.map(function (s) { return "\u2022 " + s; }).join("\n") : "");
            return SAZ.notify("Send to Panel", msg);
          }, function (e) { return SAZ.uiSendToPanel(ctx, v, e.message || String(e)); });
        });
      });
    });
  });
};

SAZ.uiRefreshPlacement = function (ctx) {
  return SAZ.lcEnsureTrack(ctx).then(function () { return SAZ.lcRefreshPlacement(ctx); }).then(function () {
    return SAZ.notify("Placement refreshed", "Ckts / Placed / Status on Load Calc now match the circuits linked on every schedule.");
  });
};


/* ==========================================================
   v2.7  PHASE 3 - panel schedule tools
   - SAZ.lockLinkedRows: circuit rows carrying a live link (sub-feed demand,
     Load Calc load, any cross-sheet formula in an input cell) are LOCKED and
     shown as calculated (no fill, red text) on both the circuit schedule and
     the Circuit Entry Dashboard; every other circuit row is an ordinary input
     row (grey fill, blue text, unlocked). Multi-pole continuation rows follow
     their breaker. Runs after every command that writes circuits.
   - SAZ.moveCircuit / uiMoveCircuit: move any circuit (sub-feed, linked load
     or manual entry) to a chosen circuit number, with optional swap.
   ========================================================== */
SAZ.CKT_NOTES_COL = { std: "M", mcc: "AD" };   /* v2.12: Notes column per layout (Panelboard / Switchboard M, MCC AD) */
SAZ.CKT_INPUT_RUNS = {
  std: [["C", "E"], ["I", "J"], ["M", "M"], ["O", "W"]],
  mcc: [["C", "E"], ["I", "J"], ["M", "R"], ["V", "AD"], ["AF", "AN"]]
};
SAZ.INPUT_FILL = "#D9D9D9"; SAZ.INPUT_FONT = "#0000FF"; SAZ.CALC_FONT = "#FF0000";
SAZ.colNum = function (L) { var n = 0; String(L).toUpperCase().split("").forEach(function (ch) { n = n * 26 + (ch.charCodeAt(0) - 64); }); return n; };
SAZ.colLet = function (n) { var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
SAZ.boardRuns = function (g) { return String(SAZ.geoUsable(g.CKT_KVA_COL) || "").toUpperCase() === "AN" ? SAZ.CKT_INPUT_RUNS.mcc : SAZ.CKT_INPUT_RUNS.std; };
SAZ.isLinkFormula = function (f) { return typeof f === "string" && f.charAt(0) === "=" && f.indexOf("!") > 0; };

/* Read every input cell of the grid once: returns { grid, runs, lastCol, f: [[...]] (B..lastCol per row) } */
SAZ.gridInputs = function (ctx, tag) {
  return SAZ.cktGrid(ctx, tag).then(function (grid) {
    var runs = SAZ.boardRuns(grid.g);
    var lastCol = runs[runs.length - 1][1];
    var sh = ctx.workbook.worksheets.getItem(tag);
    var rg = sh.getRange("B" + grid.first + ":" + lastCol + grid.last);
    rg.load("formulas");
    return ctx.sync().then(function () { return { grid: grid, runs: runs, lastCol: lastCol, f: rg.formulas }; });
  });
};
SAZ.rowHasInput = function (gi, i) {
  var row = gi.f[i];
  return gi.runs.some(function (run) {
    for (var c = SAZ.colNum(run[0]); c <= SAZ.colNum(run[1]); c++) { var v = row[c - 2]; if (v !== "" && v !== null && v !== undefined) { return true; } }
    return false;
  });
};
SAZ.rowIsLinked = function (gi, i) {
  var row = gi.f[i];
  return gi.runs.some(function (run) {
    for (var c = SAZ.colNum(run[0]); c <= SAZ.colNum(run[1]); c++) { if (SAZ.isLinkFormula(row[c - 2])) { return true; } }
    return false;
  });
};

SAZ.lockLinkedRows = function (ctx, tag) {
  return SAZ.gridInputs(ctx, tag).then(function (gi) {
    var grid = gi.grid, linkedRow = {};
    grid.rows.forEach(function (r, i) { if (SAZ.rowIsLinked(gi, i)) { linkedRow[r.row] = true; } });
    grid.rows.forEach(function (r) { if (r.owner && linkedRow[r.owner]) { linkedRow[r.row] = true; } });
    var sheet = ctx.workbook.worksheets.getItem(tag);
    return SAZ.withUnprotected(ctx, sheet, function () {
      var n = 0;
      grid.rows.forEach(function (r) {
        var lk = !!linkedRow[r.row]; if (lk) { n++; }
        gi.runs.forEach(function (run) {
          var cell = sheet.getRange(run[0] + r.row + ":" + run[1] + r.row);
          cell.format.protection.locked = lk;
          if (lk) { cell.format.fill.clear(); cell.format.font.color = SAZ.CALC_FONT; }
          else { cell.format.fill.color = SAZ.INPUT_FILL; cell.format.font.color = SAZ.INPUT_FONT; }
        });
        if (lk && SAZ.CKT_LEN_COL) { var lc = sheet.getRange(SAZ.CKT_LEN_COL + r.row); lc.format.protection.locked = false; lc.format.fill.color = SAZ.INPUT_FILL; lc.format.font.color = SAZ.INPUT_FONT; }   /* v2.10: Length Ft stays a manual entry on linked rows */
        if (lk) { var ncL = SAZ.CKT_NOTES_COL[gi.runs === SAZ.CKT_INPUT_RUNS.mcc ? "mcc" : "std"]; var nC = sheet.getRange(ncL + r.row); nC.format.protection.locked = false; nC.format.fill.color = SAZ.INPUT_FILL; nC.format.font.color = SAZ.INPUT_FONT; }   /* v2.12: Notes stay an editable (grey) entry on linked rows */
      });
      return ctx.sync().then(function () { return n; });
    });
  });
};

/* Relock after every command that writes or rebuilds circuits. */
SAZ.afterHook = function (name, tagOf) {
  var f = SAZ[name];
  if (typeof f !== "function" || f._sazHooked) { return; }
  var w = function () {
    var a = arguments, ctx = a[0], t = tagOf(a);
    return f.apply(SAZ, a).then(function (r) {
      if (!t) { return r; }
      return SAZ.lockLinkedRows(ctx, t).catch(function () { return null; }).then(function () { return r; });
    });
  };
  w._sazHooked = true;
  SAZ[name] = w;
};
SAZ.afterHook("linkEquipment",      function (a) { return a[1]; });
SAZ.afterHook("clearParentCircuit", function (a) { return a[1]; });
SAZ.afterHook("sendToPanel",        function (a) { return a[2] && a[2].panel; });
SAZ.afterHook("repairSchedule",     function (a) { return a[1]; });
SAZ.afterHook("resizePanel",        function (a) { return a[1]; });

/* Rows a breaker at grid row index i with p poles occupies (panelboard: p rows
   down the same block; switchboard / MCC: one row). null if it crosses an edge. */
SAZ.spanRows = function (grid, i, p) {
  var r0 = grid.rows[i].row;
  if (!grid.spans || p <= 1) { return [i]; }
  var blk = grid.blocks.filter(function (b) { return r0 >= b.f && r0 <= b.l; })[0];
  if (!blk || r0 + p - 1 > blk.l) { return null; }
  var out = []; for (var k = 0; k < p; k++) { out.push(i + k); } return out;
};
SAZ.ownerIndex = function (grid, i) {
  var r = grid.rows[i]; if (!r.owner) { return i; }
  return r.owner - grid.first;
};

SAZ.moveCircuit = function (ctx, tag, fromNum, toNum, opt) {
  opt = opt || {};
  return SAZ.gridInputs(ctx, tag).then(function (gi) {
    var grid = gi.grid;
    var idxOf = function (num) { var rec = grid.byNum[num]; return rec ? rec.row - grid.first : -1; };
    var si = idxOf(fromNum), di = idxOf(toNum);
    if (si < 0) { throw new Error("Circuit " + fromNum + " does not exist on " + tag + "."); }
    if (di < 0) { throw new Error("Circuit " + toNum + " does not exist on " + tag + "."); }
    si = SAZ.ownerIndex(grid, si);
    if (!SAZ.rowHasInput(gi, si) && !grid.rows[si].desc) { throw new Error("Circuit " + fromNum + " on " + tag + " is empty - nothing to move."); }
    var p = grid.rows[si].poles || 1;
    var src = SAZ.spanRows(grid, si, p) || [si];
    if (grid.rows[di].owner && src.indexOf(SAZ.ownerIndex(grid, di)) < 0) {
      throw new Error("Circuit " + toNum + " is tied into the multi-pole breaker at circuit " + grid.rows[SAZ.ownerIndex(grid, di)].num + ".");
    }
    var dst = SAZ.spanRows(grid, di, p);
    if (!dst) { throw new Error("A " + p + "-pole breaker at circuit " + toNum + " would cross the end of its block."); }
    if (dst[0] === src[0]) { return { text: "Circuit " + fromNum + " is already there.", moved: false }; }
    /* occupants of the destination that are not the source itself */
    var occ = {};
    dst.forEach(function (i) {
      if (src.indexOf(i) >= 0) { return; }
      var o = SAZ.ownerIndex(grid, i);
      if (src.indexOf(o) >= 0) { return; }
      if (SAZ.rowHasInput(gi, o) || grid.rows[o].desc) { occ[o] = true; }
    });
    var occList = Object.keys(occ).map(Number);
    var swap = null;
    if (occList.length) {
      var names = occList.map(function (o) { return "ckt " + grid.rows[o].num + (grid.rows[o].desc ? " (" + grid.rows[o].desc + ")" : ""); }).join(", ");
      if (!opt.swap) { throw new Error("Circuit " + toNum + " is occupied by " + names + ". Choose an open circuit or pick Swap."); }
      if (occList.length > 1) { throw new Error("Swap needs exactly one circuit at the destination; found " + names + "."); }
      var oi = occList[0], q = grid.rows[oi].poles || 1;
      var oRows = SAZ.spanRows(grid, oi, q) || [oi];
      var back = SAZ.spanRows(grid, si, q);
      var overlap = oRows.some(function (i) { return src.indexOf(i) >= 0; }) || (back || []).some(function (i) { return dst.indexOf(i) >= 0 && oRows.indexOf(i) < 0; });
      if (!back || overlap || oRows[0] !== dst[0] || q !== p) {
        throw new Error("Swap only works between two breakers with the same pole count whose circuits do not overlap (" + names + " is " + q + "-pole, the source is " + p + "-pole).");
      }
      swap = { from: oRows, to: back };
    }
    var sheet = ctx.workbook.worksheets.getItem(tag);
    var snap = function (rows) { return rows.map(function (i) { return gi.f[i]; }); };
    var srcF = snap(src), swF = swap ? snap(swap.from) : null;
    var writeRow = function (i, rowF) {
      var r = grid.rows[i].row;
      gi.runs.forEach(function (run) {
        var a = SAZ.colNum(run[0]), b = SAZ.colNum(run[1]);
        var vals = []; for (var c = a; c <= b; c++) { var v = rowF ? rowF[c - 2] : ""; vals.push(v === null || v === undefined ? "" : v); }
        sheet.getRange(run[0] + r + ":" + run[1] + r).formulas = [vals];
      });
    };
    return SAZ.withUnprotected(ctx, sheet, function () {
      src.forEach(function (i) { writeRow(i, null); });
      if (swap) { swap.from.forEach(function (i) { writeRow(i, null); }); }
      dst.forEach(function (i, k) { writeRow(i, srcF[k]); });
      if (swap) { swap.to.forEach(function (i, k) { writeRow(i, swF[k]); }); }
      return ctx.sync();
    }).then(function () {
      var lbl = function (rows) { return rows.map(function (i) { return grid.rows[i].num; }).join(","); };
      var descOf = function (rows) { return String(grid.rows[rows[0]].desc || ""); };
      var moves = [{ desc: descOf(src), label: lbl(dst), from: lbl(src) }];
      if (swap) { moves.push({ desc: descOf(swap.from), label: lbl(swap.to), from: lbl(swap.from) }); }
      var touchLC = srcF.concat(swF || []).some(function (row) { return row.some(function (v) { return typeof v === "string" && v.indexOf("Load Calc") >= 0; }); });
      return SAZ.registryRecords(ctx).then(function (reg) {
        var chain = Promise.resolve();
        moves.forEach(function (m) {
          var hit = reg.filter(function (r) { return SAZ.descIs(m.desc, [r.tag]) && r.fed.toLowerCase() === tag.toLowerCase(); })[0];
          if (hit) { chain = chain.then(function () { return SAZ.updateRegistryLink(ctx, hit.tag, tag, m.label); }); }
        });
        return chain;
      }).then(function () {
        return touchLC && typeof SAZ.lcRefreshPlacement === "function" ? SAZ.lcRefreshPlacement(ctx).catch(function () { return null; }) : null;
      }).then(function () { return SAZ.lockLinkedRows(ctx, tag); }).then(function () {
        ctx.workbook.application.calculate(Excel.CalculationType.recalculate);
        return ctx.sync().then(function () {
          return { moved: true, text: moves.map(function (m) { return "\u2022 " + (m.desc || "(circuit)") + ":  ckt " + m.from + "  \u2192  ckt " + m.label; }).join("\n") };
        });
      });
    });
  });
};

SAZ.uiMoveCircuit = function (ctx, keep, err) {
  var active = ctx.workbook.worksheets.getActiveWorksheet(); active.load("name");
  var sel = ctx.workbook.getSelectedRange(); sel.load("rowIndex");
  return ctx.sync().then(function () {
    var tag = active.name;
    return SAZ.cktGrid(ctx, tag).catch(function () { return null; }).then(function (grid) {
      if (!grid) { return SAZ.notify("Open a schedule", "Go to a panelboard, switchboard or MCC schedule, click the circuit you want to move (schedule or Circuit Entry Dashboard), then run Move Circuit."); }
      keep = keep || {};
      var at = grid.rows.filter(function (r) { return r.row === sel.rowIndex + 1; })[0];
      if (at && at.owner) { at = grid.rows[at.owner - grid.first]; }
      var from = keep.from || (at ? String(at.num) : "");
      var fields = [
        { id: "from", label: "Move circuit", type: "number", value: from, required: true },
        { id: "to",   label: "To circuit",   type: "number", value: keep.to || "", required: true },
        { id: "mode", label: "If the destination is occupied", type: "select", options: ["Refuse", "Swap the two circuits"], value: keep.mode || "Refuse" } ];
      var text = (at && (at.desc || at.conn) ? "Selected: ckt " + at.num + " - " + (at.desc || "(no description)") + (at.poles > 1 ? " (" + at.poles + "-pole)" : "") + ". " : "") +
        "Sub-feeds, Load Calc loads and manual circuits move with every linked cell; the registry and Load Calc tracking follow.";
      return SAZ.form("Move Circuit - " + tag, text, fields, err || "", 55).then(function (ans) {
        if (!ans.ok) { return null; }
        var v = ans.values;
        return SAZ.moveCircuit(ctx, tag, parseInt(v.from, 10), parseInt(v.to, 10), { swap: /^Swap/.test(v.mode) }).then(function (res) {
          return SAZ.notify("Move Circuit", res.text);
        }, function (e) { return SAZ.uiMoveCircuit(ctx, v, e.message || String(e)); });
      });
    });
  });
};

function MoveCircuit (event) { SAZ.run("Move Circuit", function (ctx) {
  return SAZ.uiMoveCircuit(ctx);
})(event); }


/* ==========================================================
   v2.8  DEVICE AIC LINK  -  Short Circuit Device AIC is no longer a hand entry.
   Registry column K carries a live link to each schedule's Std Device AIC (D13,
   which honours the schedule's E13 override). Short Circuit R looks it up by
   node tag. A direct reference (not INDIRECT) follows sheet renames.
   ========================================================== */
SAZ.REG_AIC_COL = "K";
SAZ.healRegistryAic = function (ctx) {
  var dash = ctx.workbook.worksheets.getItem("Dashboard");
  var tags = dash.getRange("A101:A350"); tags.load("values");
  var sheets = ctx.workbook.worksheets; sheets.load("items/name");
  return ctx.sync().then(function () {
    var byName = {};
    sheets.items.forEach(function (s) { byName[s.name.toLowerCase()] = s.name; });
    var probes = tags.values.map(function (r) {
      var t = String(r[0] || "").trim();
      var nm = t ? byName[t.toLowerCase()] : null;
      if (!nm || /^TPL_/i.test(nm)) { return null; }
      var c = ctx.workbook.worksheets.getItem(nm).getRange("C13"); c.load("values");
      return { name: nm, c: c };
    });
    return ctx.sync().then(function () {
      var f = probes.map(function (p) {
        if (!p || String(p.c.values[0][0] || "").indexOf("Std Device AIC") !== 0) { return [""]; }
        return ["='" + p.name.replace(/'/g, "''") + "'!$D$13"];
      });
      return SAZ.withUnprotected(ctx, dash, function () {
        dash.getRange("K100").values = [["Std AIC (A)"]];
        var k = dash.getRange("K101:K350");
        k.formulas = f;
        k.numberFormat = f.map(function () { return ["#,##0"]; });
        k.format.font.name = "Segoe UI"; k.format.font.size = 10; k.format.font.color = "#FF0000";
        k.format.horizontalAlignment = "Center";
        return ctx.sync();
      });
    });
  });
};
(function () {
  var reg = SAZ.registerEquipment, rs = SAZ.refreshStudies;
  SAZ.registerEquipment = function (ctx, rec) {
    return reg.apply(SAZ, arguments).then(function (v) { return SAZ.healRegistryAic(ctx).then(function () { return v; }); });
  };
  if (rs) {
    SAZ.refreshStudies = function (ctx) {
      return SAZ.healRegistryAic(ctx).then(function () { return rs.apply(SAZ, [ctx].concat([].slice.call(arguments, 1))); });
    };
  }
})();

/* ==========================================================
   v2.8  TRANSFER SWITCHES (ATS / MTS) + GENERATOR SOURCE
   - New Transfer Switch: one dialog (tag, ATS/MTS, poles, SE rated, scope,
     load-side board, generator unit, optional rating / WCR overrides).
   - The switch is a registry row (Type ATS/MTS, Fed From = normal source,
     column L Alt Source = generator unit) and a row in Generator ▸ TRANSFER
     SWITCHES, where rating and WCR size themselves.  The generator unit is a
     registry row too (Type Generator, no Fed From = a source in the map).
   - The load-side board is re-fed from the switch; the Feeder Schedule derives
     the normal feeder (source -> ATS) and an emergency feeder (GEN -> ATS,
     code F-<ATS>-E); Short Circuit uses the worse of the utility path and the
     generator's own fault contribution at the switch.
   ========================================================== */
SAZ.XFER_SHEET = "Generator";
SAZ.XFER_FIRST = 62; SAZ.XFER_LAST = 69;      /* Generator!B62:U69 */
SAZ.XFER_BRANCH_COL = "W";   /* v2.12: NEC branch served by each switch (Generator!W62:W69) */
SAZ.NEC_BRANCHES = ["700 Emergency", "701 Legally Required", "702 Optional Standby"];
SAZ.GEN_UNIT_FIRST = 53; SAZ.GEN_UNIT_LAST = 58;
SAZ.REG_ALT_COL = 11;                          /* registry column L (0-based from A) */
SAZ.isXferType = function (t) { return /^(ATS|MTS)$/i.test(String(t || "").trim()); };
SAZ.isSourceType = function (t) { return /^(ATS|MTS|Generator|UPS)$/i.test(String(t || "").trim()); };
SAZ.isBoardType = function (t) { return !SAZ.isSourceType(t); };

/* full registry rows A:L with sheet address info */
SAZ.registryFull = function (ctx) {
  var dash = ctx.workbook.worksheets.getItem(CFG.regSheet);
  var rg = dash.getRange("A101:L350"); rg.load("values,formulas");
  return ctx.sync().then(function () {
    var out = [];
    rg.values.forEach(function (r, i) {
      var t = String(r[0] == null ? "" : r[0]).trim();
      if (t) { out.push({ tag: t, type: String(r[1] || "").trim(), volts: String(r[2] || ""), fed: String(r[3] || "").trim(),
        ft: String(r[8] || "").trim(), ckt: String(r[9] || "").trim(), alt: String(r[11] || "").trim(), row: 101 + i, f: rg.formulas[i] }); }
    });
    return { dash: dash, rows: out };
  });
};
SAZ.registryWriteRow = function (ctx, row, vals12) {
  var dash = ctx.workbook.worksheets.getItem(CFG.regSheet);
  return SAZ.withUnprotected(ctx, dash, function () {
    var r = dash.getRange("A" + row + ":J" + row);
    r.formulas = [vals12.slice(0, 10)];
    SAZ.formatRegistryRow(r);
    var l = dash.getRange("L" + row); l.values = [[vals12[11] || ""]];
    return ctx.sync();
  });
};
SAZ.registryFreeRow = function (ctx) {
  return SAZ.registryFull(ctx).then(function (R) {
    var used = {}; R.rows.forEach(function (x) { used[x.row] = 1; });
    for (var r = 101; r <= 350; r++) { if (!used[r]) { return r; } }
    throw new Error("Equipment registry is full.");
  });
};

/* generator units on the Generator sheet */
SAZ.genUnits = function (ctx) {
  var g = ctx.workbook.worksheets.getItem(SAZ.XFER_SHEET);
  var rg = g.getRange("B" + SAZ.GEN_UNIT_FIRST + ":B" + SAZ.GEN_UNIT_LAST); rg.load("values");
  return ctx.sync().then(function () {
    return rg.values.map(function (r, i) { return { tag: String(r[0] || "").trim(), row: SAZ.GEN_UNIT_FIRST + i }; })
      .filter(function (u) { return u.tag; });
  });
};
SAZ.ensureGenUnit = function (ctx) {
  return SAZ.genUnits(ctx).then(function (u) {
    if (u.length) { return u; }
    var g = ctx.workbook.worksheets.getItem(SAZ.XFER_SHEET);
    return SAZ.withUnprotected(ctx, g, function () {
      g.getRange("B" + SAZ.GEN_UNIT_FIRST).values = [["GEN-1"]];   /* v2.12: Description (C) is a formula - named from the NEC branches its switches serve */
      return ctx.sync();
    }).then(function () { return SAZ.genUnits(ctx); });
  });
};
SAZ.xferRows = function (ctx) {
  var g = ctx.workbook.worksheets.getItem(SAZ.XFER_SHEET);
  var rg = g.getRange("B" + SAZ.XFER_FIRST + ":I" + SAZ.XFER_LAST); rg.load("values");
  return ctx.sync().then(function () {
    return rg.values.map(function (r, i) { return { tag: String(r[0] || "").trim(), row: SAZ.XFER_FIRST + i, v: r }; });
  });
};
SAZ.scFaultRef = function (node) {
  return "=IFERROR(XLOOKUP(\"" + String(node).replace(/"/g, "\"\"") + "\",'Short Circuit'!$A$6:$A$255,'Short Circuit'!$M$6:$M$255,\"\"),\"\")";
};
/* point a board's Source Fault cell at the Short Circuit bus fault of a non-sheet source (ATS) */
SAZ.linkSrcFaultToNode = function (ctx, boardTag, node) {
  return SAZ.geo(ctx, boardTag).then(function (g) {
    var src = SAZ.geoUsable(g.SRCFAULT_CELL);
    if (!src) { return null; }
    var sh = ctx.workbook.worksheets.getItem(boardTag);
    return SAZ.withUnprotected(ctx, sh, function () { sh.getRange(src).formulas = [[SAZ.scFaultRef(node)]]; return ctx.sync(); });
  }).catch(function () { return null; });
};
/* rename the parent's breaker description from the load board to the switch */
SAZ.relabelParentCircuit = function (ctx, parent, fromTag, toTag) {
  var ws = ctx.workbook.worksheets.getItemOrNullObject(parent); ws.load("isNullObject");
  return ctx.sync().then(function () {
    if (ws.isNullObject) { return 0; }
    return SAZ.cktGrid(ctx, parent).then(function (grid) {
      var hits = grid.rows.filter(function (r) { return SAZ.descIs(r.desc, [fromTag]);   /* v2.12 */ });
      if (!hits.length) { return 0; }
      return SAZ.withUnprotected(ctx, ws, function () {
        hits.forEach(function (h) { ws.getRange(grid.cols.dc + h.row).values = [[toTag]]; });
        return ctx.sync();
      }).then(function () { return typeof SAZ.lockLinkedRows === "function" ? SAZ.lockLinkedRows(ctx, parent).catch(function () { return null; }) : null; })
        .then(function () { return hits.length; });
    });
  }).catch(function () { return 0; });
};

/* registry Std AIC for transfer switches = WCR used on the Generator sheet */
SAZ.healXferAic = function (ctx) {
  return SAZ.registryFull(ctx).then(function (R) {
    var x = R.rows.filter(function (r) { return SAZ.isXferType(r.type); });
    if (!x.length) { return null; }
    return SAZ.withUnprotected(ctx, R.dash, function () {
      x.forEach(function (r) {
        var k = R.dash.getRange("K" + r.row);
        k.formulas = [["=IFERROR(XLOOKUP($A" + r.row + "," + SAZ.XFER_SHEET + "!$B$" + SAZ.XFER_FIRST + ":$B$" + SAZ.XFER_LAST + "," +
          SAZ.XFER_SHEET + "!$S$" + SAZ.XFER_FIRST + ":$S$" + SAZ.XFER_LAST + ",\"\"),\"\")"]];
        k.numberFormat = [["#,##0"]];
      });
      return ctx.sync();
    });
  });
};
(function () {
  var base = SAZ.healRegistryAic;
  if (typeof base === "function" && !base._xfer) {
    SAZ.healRegistryAic = function (ctx) { return base.apply(SAZ, arguments).then(function (o) { return SAZ.healXferAic(ctx).then(function () { return o; }); }); };
    SAZ.healRegistryAic._xfer = true;
  }
})();

/* Distribution Map: switches and generators have no schedule sheet */
(function () {
  var base = SAZ.mapApply;
  if (typeof base !== "function" || base._xfer) { return; }
  SAZ.mapApply = function (ctx, req) {
    var C = String(req.tag || "").trim(), P = String(req.fed || "").trim();
    return SAZ.registryFull(ctx).then(function (R) {
      var find = function (t) { return R.rows.filter(function (r) { return r.tag.toLowerCase() === String(t).toLowerCase(); })[0]; };
      var cr = find(C), pr = P ? find(P) : null;
      if (cr && SAZ.isSourceType(cr.type)) {
        if (/^Generator$/i.test(cr.type)) { return Promise.reject(new Error(C + " is a generator - it is a source and is not fed from anything. Its emergency feed is set on the transfer switch.")); }
        return SAZ.isUtilityTag(ctx, P).then(function (isU) {
          return SAZ.registrySet(ctx, C, { 3: P, 8: P ? (isU ? "Utility / Service" : "Sub-Feed") : "" }).then(function () {
            var g = ctx.workbook.worksheets.getItem(SAZ.XFER_SHEET);
            return SAZ.xferRows(ctx).then(function (xr) {
              var h = xr.filter(function (x) { return x.tag.toLowerCase() === C.toLowerCase(); })[0];
              if (!h) { return null; }
              return SAZ.withUnprotected(ctx, g, function () { g.getRange("G" + h.row).values = [[P]]; return ctx.sync(); });
            });
          }).then(function () {
            return C + " normal source set to " + (P || "(none)") + "." + (P && !isU ? " Assign its breaker on " + P + " by hand (Move Circuit / circuit entry)." : "");
          });
        });
      }
      if (pr && /^Generator$/i.test(pr.type)) {
        return Promise.reject(new Error(P + " is a generator. Boards connect to a generator through a transfer switch (New Transfer Switch)."));
      }
      if (pr && SAZ.isXferType(pr.type)) {
        return SAZ.mapDetach(ctx, C).then(function () {
          return SAZ.registrySet(ctx, C, { 3: P, 8: "Transfer Switch", 9: "" });
        }).then(function () { return SAZ.linkSrcFaultToNode(ctx, C, P); })
          .then(function () { return SAZ.healXferDemand(ctx); })
          .then(function () { return C + " is now fed from transfer switch " + P + "."; });
      }
      return base.call(SAZ, ctx, req);
    });
  };
  SAZ.mapApply._xfer = true;
})();

/* registry demand of a switch = live sum of the demand of the boards it feeds;
   its Generator row reads that (no circular range). */
SAZ.healXferDemand = function (ctx) {
  return SAZ.registryFull(ctx).then(function (R) {
    var sw = R.rows.filter(function (r) { return SAZ.isXferType(r.type); });
    if (!sw.length) { return null; }
    var ws = ctx.workbook.worksheets; ws.load("items/name");
    return ctx.sync().then(function () {
      var names = ws.items.map(function (s) { return s.name.toLowerCase(); });
      var refs = {};
      var pass = [];
      sw.forEach(function (s) {
        var kids = R.rows.filter(function (r) { return r.fed.toLowerCase() === s.tag.toLowerCase() && names.indexOf(r.tag.toLowerCase()) >= 0; });
        pass.push({ s: s, kids: kids });
      });
      return Promise.all(pass.map(function (p) {
        return Promise.all(p.kids.map(function (k) { return SAZ.geo(ctx, k.tag).then(function (g) { return SAZ.sheetRef(k.tag, g.DEMAND_CELL); }).catch(function () { return null; }); }))
          .then(function (rs) { refs[p.s.tag] = rs.filter(Boolean); });
      })).then(function () {
        return SAZ.withUnprotected(ctx, R.dash, function () {
          pass.forEach(function (p) {
            var rr = refs[p.s.tag];
            R.dash.getRange("F" + p.s.row).formulas = [[rr.length ? "=" + rr.map(function (x) { return String(x).replace(/^=/, ""); }).join("+") : 0]];
          });
          return ctx.sync();
        });
      });
    });
  });
};

SAZ.xferRemove = function (ctx, tag) {
  return SAZ.registryFull(ctx).then(function (R) {
    var me = R.rows.filter(function (r) { return r.tag.toLowerCase() === tag.toLowerCase(); })[0];
    if (!me) { throw new Error(tag + " is not in the registry."); }
    var kids = R.rows.filter(function (r) { return r.fed.toLowerCase() === tag.toLowerCase(); });
    var chain = Promise.resolve();
    kids.forEach(function (k, i) {
      chain = chain.then(function () {
        return SAZ.registrySet(ctx, k.tag, { 3: me.fed, 8: i === 0 ? me.ft : "", 9: i === 0 ? me.ckt : "" });
      }).then(function () {
        if (i === 0 && me.fed) { var pv = R.rows.filter(function (r) { return r.tag.toLowerCase() === me.fed.toLowerCase(); })[0], num = function (x) { var m = /([0-9]+)/.exec(String(x || "")); return m ? Number(m[1]) : 0; }; var lab = (pv && num(pv.volts) && num(k.volts) && num(pv.volts) !== num(k.volts)) ? "T-" + k.tag : k.tag; return SAZ.relabelParentCircuit(ctx, me.fed, tag, lab).then(function () { return SAZ.relabelParentCircuit(ctx, me.fed, "T-" + tag, lab); }); }
        return null;
      }).then(function () {
        return SAZ.isUtilityTag(ctx, me.fed).then(function (isU) {
          if (!me.fed) { return null; }
          if (!isU) { return SAZ.geo(ctx, me.fed).then(function (pg) {
              var bus = SAZ.geoUsable(pg.BUSFAULT_CELL);
              return SAZ.geo(ctx, k.tag).then(function (cg) {
                var src = SAZ.geoUsable(cg.SRCFAULT_CELL); if (!src || !bus) { return null; }
                var sh = ctx.workbook.worksheets.getItem(k.tag);
                return SAZ.withUnprotected(ctx, sh, function () { sh.getRange(src).formulas = [[SAZ.sheetRef(me.fed, bus)]]; return ctx.sync(); });
              });
            }).catch(function () { return null; }); }
          return SAZ.geo(ctx, k.tag).then(function (cg) {   /* back on the utility: template source-fault formula */
            var src = SAZ.geoUsable(cg.SRCFAULT_CELL); if (!src) { return null; }
            var sh = ctx.workbook.worksheets.getItem(k.tag);
            return SAZ.withUnprotected(ctx, sh, function () { sh.getRange(src).formulas = [["=IF(Cfg_SCA=\"\",\"\",Cfg_SCA)"]]; return ctx.sync(); });
          }).catch(function () { return null; });
        });
      });
    });
    return chain.then(function () {
      return SAZ.withUnprotected(ctx, R.dash, function () {
        R.dash.getRange("A" + me.row + ":L" + me.row).clear(Excel.ClearApplyTo.contents);
        return ctx.sync();
      });
    }).then(function () { return SAZ.xferRows(ctx); }).then(function (xr) {
      var h = xr.filter(function (x) { return x.tag.toLowerCase() === tag.toLowerCase(); })[0];
      if (!h) { return null; }
      var g = ctx.workbook.worksheets.getItem(SAZ.XFER_SHEET);
      return SAZ.withUnprotected(ctx, g, function () {
        ["B" + h.row + ":I" + h.row, "O" + h.row, "R" + h.row, SAZ.XFER_BRANCH_COL + h.row].forEach(function (a) { g.getRange(a).clear(Excel.ClearApplyTo.contents); });
        return ctx.sync();
      });
    }).then(function () {
      /* drop the generator's registry row when no switch uses it any more */
      return SAZ.registryFull(ctx).then(function (R2) {
        var gen = me.alt, still = R2.rows.some(function (r) { return r.alt.toLowerCase() === gen.toLowerCase(); });
        var gr = R2.rows.filter(function (r) { return r.tag.toLowerCase() === gen.toLowerCase() && /^Generator$/i.test(r.type); })[0];
        if (still || !gr) { return null; }
        return SAZ.withUnprotected(ctx, R2.dash, function () { R2.dash.getRange("A" + gr.row + ":L" + gr.row).clear(Excel.ClearApplyTo.contents); return ctx.sync(); });
      });
    }).then(function () { return { moved: kids.map(function (k) { return k.tag; }), to: me.fed }; });
  });
};

SAZ.xferCreate = function (ctx, a) {
  /* a: { tag, type, poles, se, scope, branch, load, gen, rating, wcr } */
  var R, load, genRow;
  return SAZ.registryFull(ctx).then(function (x) {
    R = x;
    load = R.rows.filter(function (r) { return r.tag.toLowerCase() === a.load.toLowerCase(); })[0];
    if (!load) { throw new Error("Load-side board " + a.load + " is not in the registry."); }
    if (R.rows.some(function (r) { return r.tag.toLowerCase() === a.tag.toLowerCase(); })) { throw new Error(a.tag + " is already used by another piece of equipment."); }
    return SAZ.genUnits(ctx);
  }).then(function (units) {
    var u = units.filter(function (x) { return x.tag.toLowerCase() === a.gen.toLowerCase(); })[0];
    if (!u) { throw new Error("Generator unit " + a.gen + " is not in Generator ▸ GENERATOR UNITS."); }
    genRow = u.row;
    return SAZ.xferRows(ctx);
  }).then(function (xr) {
    var slot = xr.filter(function (x) { return !x.tag; })[0];
    if (!slot) { throw new Error("All " + (SAZ.XFER_LAST - SAZ.XFER_FIRST + 1) + " transfer-switch rows on the Generator sheet are in use."); }
    var g = ctx.workbook.worksheets.getItem(SAZ.XFER_SHEET);
    return SAZ.withUnprotected(ctx, g, function () {
      g.getRange("B" + slot.row + ":I" + slot.row).values = [[a.tag, a.type, Number(a.poles) || 3, a.se, a.scope, load.fed, a.gen, load.tag]];
      g.getRange("O" + slot.row).values = [[Number(a.rating) > 0 ? Number(a.rating) : ""]];
      g.getRange("R" + slot.row).values = [[Number(a.wcr) > 0 ? Number(a.wcr) : ""]];
      g.getRange(SAZ.XFER_BRANCH_COL + slot.row).values = [[a.branch || ""]];   /* v2.12 */
      return ctx.sync();
    }).then(function () { return slot.row; });
  }).then(function (xrow) {
    /* generator registry row (source) */
    var hasGen = R.rows.some(function (r) { return r.tag.toLowerCase() === a.gen.toLowerCase(); });
    var step = Promise.resolve();
    if (!hasGen) {
      step = SAZ.registryFreeRow(ctx).then(function (row) {
        return SAZ.registryWriteRow(ctx, row, [a.gen, "Generator", load.volts, "", "=" + SAZ.XFER_SHEET + "!$R$" + genRow + "&\" kW\"",
          "=" + SAZ.XFER_SHEET + "!$J$" + genRow, SAZ.XFER_SHEET, "Emergency / standby source", "", "", "", ""]);
      });
    }
    return step.then(function () { return SAZ.registryFreeRow(ctx); }).then(function (row) {
      var note = a.poles + "P " + a.type + (/^y/i.test(a.se) ? ", SE rated" : "") + " - " + a.scope + (a.branch ? " - NEC " + a.branch : "");
      return SAZ.registryWriteRow(ctx, row, [a.tag, a.type, load.volts, load.fed, "=" + SAZ.XFER_SHEET + "!$P$" + xrow,
        0, SAZ.XFER_SHEET, note, load.ft || "", load.ckt || "", "", a.gen]);
    });
  }).then(function () {
    return SAZ.registrySet(ctx, load.tag, { 3: a.tag, 8: "Transfer Switch", 9: "" });
  }).then(function () {
    return SAZ.isUtilityTag(ctx, load.fed).then(function (isU) {
      if (!isU && load.fed) { return SAZ.relabelParentCircuit(ctx, load.fed, load.tag, a.tag); }
      return 0;
    });
  }).then(function () { return SAZ.linkSrcFaultToNode(ctx, load.tag, a.tag); })
    .then(function () { return SAZ.healXferDemand(ctx); })
    .then(function () { return SAZ.healXferAic(ctx); })
    .then(function () { return SAZ.xferRows(ctx); })
    .then(function (xr) {
      var g = ctx.workbook.worksheets.getItem(SAZ.XFER_SHEET), br = g.getRange(SAZ.XFER_BRANCH_COL + SAZ.XFER_FIRST + ":" + SAZ.XFER_BRANCH_COL + SAZ.XFER_LAST), gs = g.getRange("H" + SAZ.XFER_FIRST + ":H" + SAZ.XFER_LAST);
      br.load("values"); gs.load("values");
      return ctx.sync().then(function () {
        var seen = {};
        gs.values.forEach(function (r, k) { if (String(r[0]).toLowerCase() === String(a.gen).toLowerCase() && br.values[k][0]) { seen[String(br.values[k][0]).substring(0, 3)] = true; } });
        return { tag: a.tag, normal: load.fed, load: load.tag, gen: a.gen, branches: Object.keys(seen).length };
      });
    });
};

SAZ.uiTransferSwitch = function (ctx, keep, err) {
  return SAZ.moduleState(ctx).then(function (st) {
    if (!st.Generator) { return SAZ.setModuleDeployed(ctx, "Generator", true); }
    return null;
  }).then(function () { return SAZ.ensureGenUnit(ctx); }).then(function (units) {
    return SAZ.registryFull(ctx).then(function (R) {
      var boards = R.rows.filter(function (r) { return SAZ.isBoardType(r.type); });
      var sw = R.rows.filter(function (r) { return SAZ.isXferType(r.type); });
      if (!boards.length) { return SAZ.notify("No boards yet", "Create the board the transfer switch will feed (for a whole-service switch, the main switchboard) first."); }
      var n = 1; while (R.rows.some(function (r) { return r.tag.toLowerCase() === ("ats-" + n); })) { n++; }
      keep = keep || {};
      var top = boards.filter(function (b) { return !b.fed || !R.rows.some(function (r) { return r.tag.toLowerCase() === b.fed.toLowerCase(); }); })[0] || boards[0];
      var actions = ["Create a new transfer switch"].concat(sw.map(function (s) { return "Remove " + s.tag + " (re-feed its loads from " + (s.fed || "nothing") + ")"; }));
      return SAZ.form("New Transfer Switch",
        "Sized automatically from the load-side demand and the generator unit's running load. Whole service = the switch sits between the utility source and the main board.",
        [ { id: "act", label: "Action", type: "select", options: actions, value: keep.act || actions[0] },
          { id: "tag", label: "Tag", type: "text", value: keep.tag || ("ATS-" + n) },
          { id: "type", label: "Type", type: "select", options: ["ATS", "MTS"], value: keep.type || "ATS" },
          { id: "poles", label: "Poles", type: "select", options: ["3", "4", "2"], value: keep.poles || "3", hint: "4-pole switches the neutral (separately derived generator)." },
          { id: "se", label: "Service-entrance rated", type: "select", options: ["Yes", "No"], value: keep.se || "Yes" },
          { id: "scope", label: "Scope", type: "select", options: ["Whole service", "Branch"], value: keep.scope || "Whole service" },
          { id: "branch", label: "NEC branch served", type: "select", options: SAZ.NEC_BRANCHES, value: keep.branch || SAZ.NEC_BRANCHES[2],
            hint: "One switch per branch: emergency (700), legally required (701) and optional standby (702) loads each get their own switch and downstream boards (NEC 700.10(B), 701.10). Several switches can share one generator - run this once per branch." },
          { id: "load", label: "Load side (board fed by the switch)", type: "select", options: boards.map(function (b) { return b.tag; }), value: keep.load || top.tag },
          { id: "gen", label: "Emergency source (generator unit)", type: "select", options: units.map(function (u) { return u.tag; }), value: keep.gen || units[0].tag },
          { id: "rating", label: "Rating override (A, optional)", type: "number", value: keep.rating || "" },
          { id: "wcr", label: "WCR override (A, optional)", type: "number", value: keep.wcr || "" } ],
        err || "", 80).then(function (ans) {
          if (!ans.ok) { return null; }
          var v = ans.values;
          if (v.act !== actions[0]) {
            var t = v.act.replace(/^Remove\s+/, "").replace(/\s+\(.*$/, "");
            return SAZ.xferRemove(ctx, t).then(function (o) {
              return SAZ.notify("Transfer switch removed", t + " removed. " + (o.moved.length ? o.moved.join(", ") + " fed from " + (o.to || "nothing") + " again." : ""));
            });
          }
          var tag = SAZ.cleanTag ? SAZ.cleanTag(v.tag) : String(v.tag || "").trim();
          if (/^700/.test(v.branch || "") && v.scope === "Whole service") {
            return SAZ.uiTransferSwitch(ctx, v, "An emergency (NEC 700) transfer switch may supply emergency loads only (700.5(D)). Set Scope to Branch and feed a dedicated emergency board.");
          }
          if (!tag) { return SAZ.uiTransferSwitch(ctx, v, "Enter a tag."); }
          var ld = R.rows.filter(function (r) { return r.tag === v.load; })[0];
          if (ld && R.rows.some(function (r) { return r.tag.toLowerCase() === ld.fed.toLowerCase() && SAZ.isXferType(r.type); })) {
            return SAZ.uiTransferSwitch(ctx, v, v.load + " is already fed from transfer switch " + ld.fed + ".");
          }
          if (R.rows.some(function (r) { return r.tag.toLowerCase() === tag.toLowerCase(); })) { return SAZ.uiTransferSwitch(ctx, v, tag + " is already in use."); }
          v.tag = tag;
          return SAZ.xferCreate(ctx, v).then(function (o) {
            var g = ctx.workbook.worksheets.getItem(SAZ.XFER_SHEET); g.visibility = Excel.SheetVisibility.visible; g.activate();
            return ctx.sync().then(function () {
              return SAZ.notify("Transfer switch created",
                o.tag + ": normal " + (o.normal || "(none)") + "  ->  " + o.tag + "  ->  " + o.load + ", emergency " + o.gen + ".\n\n" +
                (o.branches > 1 ? o.gen + " now serves " + o.branches + " NEC branches - its description updates automatically. Size the set for the combined load or provide automatic load shedding so 700 loads take priority over 701, and 701 over 702 (NEC 700.4(B), 701.4(B)).\n\n" : "") + "Rating and WCR size themselves on Generator ▸ TRANSFER SWITCHES. Feeder Schedule now carries F-" + o.tag + " (normal) and F-" + o.tag + "-E (emergency); run Refresh Studies to update Short Circuit.");
            });
          });
        });
    });
  });
};


/* ==========================================================
   v2.8  UPS SYSTEMS  -  loose mirror of the transfer-switch logic.
   A UPS unit (UPS ▸ UPS UNITS) is placed in the distribution between an
   input board and an output (critical) board. Sized on UPS ▸ UPS SYSTEMS:
   rating = next std kVA >= max(output demand x margin, unit sizing); input
   kVA = rating / efficiency x (1 + recharge); input / output OCPD at 125 %.
   Short Circuit: with a bypass the upstream fault passes through; with
   'None' the output is limited to the inverter (x FLA). Registry type "UPS".
   ========================================================== */
SAZ.UPS_SHEET = "UPS";
SAZ.UPS_UNIT_FIRST = 48; SAZ.UPS_UNIT_LAST = 53;
SAZ.UPSSYS_FIRST = 57; SAZ.UPSSYS_LAST = 64;
SAZ.isUpsType = function (t) { return /^UPS$/i.test(String(t || "").trim()); };

SAZ.upsUnits = function (ctx) {
  var u = ctx.workbook.worksheets.getItem(SAZ.UPS_SHEET);
  var rg = u.getRange("B" + SAZ.UPS_UNIT_FIRST + ":B" + SAZ.UPS_UNIT_LAST); rg.load("values");
  return ctx.sync().then(function () {
    return rg.values.map(function (r, i) { return { tag: String(r[0] || "").trim(), row: SAZ.UPS_UNIT_FIRST + i }; })
      .filter(function (x) { return x.tag; });
  });
};
SAZ.ensureUpsUnit = function (ctx) {
  return SAZ.upsUnits(ctx).then(function (u) {
    if (u.length) { return u; }
    var sh = ctx.workbook.worksheets.getItem(SAZ.UPS_SHEET);
    return SAZ.withUnprotected(ctx, sh, function () {
      sh.getRange("B" + SAZ.UPS_UNIT_FIRST + ":C" + SAZ.UPS_UNIT_FIRST).values = [["UPS-1", "UPS System"]];
      return ctx.sync();
    }).then(function () { return SAZ.upsUnits(ctx); });
  });
};
SAZ.upsRows = function (ctx) {
  var sh = ctx.workbook.worksheets.getItem(SAZ.UPS_SHEET);
  var rg = sh.getRange("B" + SAZ.UPSSYS_FIRST + ":E" + SAZ.UPSSYS_LAST); rg.load("values");
  return ctx.sync().then(function () {
    return rg.values.map(function (r, i) { return { tag: String(r[0] || "").trim(), row: SAZ.UPSSYS_FIRST + i, v: r }; });
  });
};

/* a board's Source Fault back to its (new) parent: board bus fault, switch/UPS node, or utility */
SAZ.restoreSrcFault = function (ctx, kidTag, parentTag) {
  if (!parentTag) { return Promise.resolve(null); }
  return SAZ.registryFull(ctx).then(function (R) {
    var p = R.rows.filter(function (r) { return r.tag.toLowerCase() === String(parentTag).toLowerCase(); })[0];
    if (p && SAZ.isSourceType(p.type)) { return SAZ.linkSrcFaultToNode(ctx, kidTag, p.tag); }
    return SAZ.isUtilityTag(ctx, parentTag).then(function (isU) {
      return SAZ.geo(ctx, kidTag).then(function (cg) {
        var src = SAZ.geoUsable(cg.SRCFAULT_CELL); if (!src) { return null; }
        var setF = function (f) { var sh = ctx.workbook.worksheets.getItem(kidTag);
          return SAZ.withUnprotected(ctx, sh, function () { sh.getRange(src).formulas = [[f]]; return ctx.sync(); }); };
        if (isU) { return setF("=IF(Cfg_SCA=\"\",\"\",Cfg_SCA)"); }
        return SAZ.geo(ctx, parentTag).then(function (pg) {
          var bus = SAZ.geoUsable(pg.BUSFAULT_CELL); if (!bus) { return null; }
          return setF(SAZ.sheetRef(parentTag, bus));
        });
      });
    });
  }).catch(function () { return null; });
};

/* UPS output load = live sum of the demand of the boards it feeds (on the UPS sheet,
   the registry row reads the UPS input kVA - no circular range) */
SAZ.healUpsDemand = function (ctx) {
  return SAZ.registryFull(ctx).then(function (R) {
    var ups = R.rows.filter(function (r) { return SAZ.isUpsType(r.type); });
    if (!ups.length) { return null; }
    var ws = ctx.workbook.worksheets; ws.load("items/name");
    return ctx.sync().then(function () { return SAZ.upsRows(ctx); }).then(function (rows) {
      var names = ws.items.map(function (s) { return s.name.toLowerCase(); });
      var jobs = ups.map(function (u) {
        var h = rows.filter(function (x) { return x.tag.toLowerCase() === u.tag.toLowerCase(); })[0];
        var kids = R.rows.filter(function (r) { return r.fed.toLowerCase() === u.tag.toLowerCase() && names.indexOf(r.tag.toLowerCase()) >= 0; });
        return Promise.all(kids.map(function (k) { return SAZ.geo(ctx, k.tag).then(function (g) { return SAZ.sheetRef(k.tag, g.DEMAND_CELL); }).catch(function () { return null; }); }))
          .then(function (rs) { return { h: h, refs: rs.filter(Boolean) }; });
      });
      return Promise.all(jobs).then(function (js) {
        var sh = ctx.workbook.worksheets.getItem(SAZ.UPS_SHEET);
        return SAZ.withUnprotected(ctx, sh, function () {
          js.forEach(function (j) {
            if (!j.h) { return; }
            var body = j.refs.length ? j.refs.map(function (x) { return String(x).replace(/^=/, ""); }).join("+") : "0";
            sh.getRange("G" + j.h.row).formulas = [["=IF($B" + j.h.row + "=\"\",\"\"," + body + ")"]];
          });
          return ctx.sync();
        });
      });
    });
  });
};
/* registry Std AIC for a UPS = SCCR used on the UPS sheet */
SAZ.healUpsAic = function (ctx) {
  return SAZ.registryFull(ctx).then(function (R) {
    var x = R.rows.filter(function (r) { return SAZ.isUpsType(r.type); });
    if (!x.length) { return null; }
    return SAZ.withUnprotected(ctx, R.dash, function () {
      x.forEach(function (r) {
        var k = R.dash.getRange("K" + r.row);
        k.formulas = [["=IFERROR(XLOOKUP($A" + r.row + "," + SAZ.UPS_SHEET + "!$B$" + SAZ.UPSSYS_FIRST + ":$B$" + SAZ.UPSSYS_LAST + "," +
          SAZ.UPS_SHEET + "!$W$" + SAZ.UPSSYS_FIRST + ":$W$" + SAZ.UPSSYS_LAST + ",\"\"),\"\")"]];
        k.numberFormat = [["#,##0"]];
      });
      return ctx.sync();
    });
  });
};
(function () {
  var base = SAZ.healRegistryAic;
  if (typeof base === "function" && !base._ups) {
    SAZ.healRegistryAic = function (ctx) { return base.apply(SAZ, arguments).then(function (o) { return SAZ.healUpsAic(ctx).then(function () { return o; }); }); };
    SAZ.healRegistryAic._ups = true; SAZ.healRegistryAic._xfer = true;
  }
})();

SAZ.upsCreate = function (ctx, a) {
  /* a: { unit, load, bypass, input ("" = output board's current feed), upsIn ("" = input board V),
          internal (bool: UPS converts upsIn -> output board V), autoX (bool: allow T-<unit> ahead of the UPS) }
     Registry: C = UPS input voltage, N = UPS output voltage (only when internal) so the Transformers
     schedule and the Distribution Map treat the UPS output as the output board's voltage. */
  var R, load, slotRow, inTag, inV = "", vin, vout, isU = false, feed = { ft: "", ckt: "" }, xfmr = "";
  var lv = function (s) { var m = /([0-9]+(?:\.[0-9]+)?)/.exec(String(s || "")); return m ? Number(m[1]) : 0; };
  var phOf = function (s) { return /1\s*\u00D8|1\s*PH|1-?PHASE/i.test(String(s || "")) ? 1 : 3; };
  return SAZ.registryFull(ctx).then(function (x) {
    R = x;
    var find = function (t) { return R.rows.filter(function (r) { return r.tag.toLowerCase() === String(t).toLowerCase(); })[0]; };
    load = find(a.load);
    if (!load) { throw new Error("Output board " + a.load + " is not in the registry."); }
    if (!SAZ.isBoardType(load.type)) { throw new Error(a.load + " is a " + load.type + ", not a board."); }
    if (find(a.unit)) { throw new Error(a.unit + " is already placed in the distribution."); }
    if (!load.volts) { throw new Error(load.tag + " has no system voltage set - set it on its schedule first."); }
    inTag = String(a.input || load.fed || "").trim();
    if (!inTag) { throw new Error(load.tag + " has no feed yet - pick the board that feeds the UPS."); }
    if (inTag.toLowerCase() === load.tag.toLowerCase()) { throw new Error("The UPS cannot be fed from its own output board."); }
    var ip = find(inTag);
    if (ip && SAZ.isUpsType(ip.type)) { throw new Error("A UPS input must come from a board or the utility, not from another UPS."); }
    if (ip && !SAZ.isBoardType(ip.type) && !SAZ.isXferType(ip.type)) { throw new Error("A UPS input must come from a board, transfer switch or the utility, not from " + inTag + "."); }
    return SAZ.isUtilityTag(ctx, inTag).then(function (u) {
      isU = u;
      if (ip) { inV = ip.volts; return null; }
      if (!isU) { throw new Error(inTag + " is not in the registry."); }
      var n = ctx.workbook.names.getItemOrNullObject("Cfg_Volt"); n.load("isNullObject");
      return ctx.sync().then(function () {
        if (n.isNullObject) { return null; }
        var r = n.getRange(); r.load("values");
        return ctx.sync().then(function () { inV = String(r.values[0][0] || ""); });
      });
    });
  }).then(function () {
    vout = load.volts;
    vin = a.internal ? (a.upsIn || inV || vout) : vout;
    if (!a.internal && a.upsIn && lv(a.upsIn) !== lv(vout)) {
      throw new Error("Without an internal transformer the UPS input runs at the output board voltage (" + vout + "). Set Internal transformer = Yes for a " + a.upsIn + " input.");
    }
    if (inV && lv(inV) && lv(vin) && Math.abs(lv(inV) - lv(vin)) > 1) {
      if (!a.autoX) { throw new Error(inTag + " is " + inV + " but the UPS input is " + vin + ". Allow a transformer ahead of the UPS, use an internal transformer, or pick a " + vin + " input board."); }
      xfmr = "T-" + a.unit;
    }
    if (String(inTag).toLowerCase() === String(load.fed || "").toLowerCase()) {
      feed = { ft: load.ft || "", ckt: load.ckt || "", same: true };
      return null;
    }
    /* new input board: find the UPS breaker slot first, then free the output board's old breaker */
    var pick = isU ? Promise.resolve(null) : SAZ.cktGrid(ctx, inTag).then(function (grid) {
      var bv = inV || vin, poles = SAZ.defaultPoles(phOf(bv), lv(bv));
      var fit = SAZ.firstFit(grid, poles, [a.unit]);
      if (!fit) { throw new Error("No open " + poles + "-pole slot on " + inTag + " - resize it first."); }
      var chk = SAZ.slotCheck(grid, fit, poles, [a.unit]);
      feed = { ft: "Sub-Feed", ckt: chk.label, row: chk.row, poles: poles };
    });
    return pick.then(function () {
      if (isU) { feed = { ft: "Utility / Service", ckt: "" }; }
      return SAZ.mapDetach(ctx, load.tag);
    });
  }).then(function () { return SAZ.upsRows(ctx); }).then(function (rows) {
    var slot = rows.filter(function (x) { return !x.tag; })[0];
    if (!slot) { throw new Error("All " + (SAZ.UPSSYS_LAST - SAZ.UPSSYS_FIRST + 1) + " UPS-system rows on the UPS sheet are in use."); }
    slotRow = slot.row;
    var sh = ctx.workbook.worksheets.getItem(SAZ.UPS_SHEET);
    return SAZ.withUnprotected(ctx, sh, function () {
      sh.getRange("B" + slotRow + ":E" + slotRow).values = [[a.unit, inTag, load.tag, a.bypass || "Static + Maint."]];
      return ctx.sync();
    });
  }).then(function () {
    if (feed.row) { return SAZ.addParentCircuitAt(ctx, inTag, a.unit, "=N(" + SAZ.UPS_SHEET + "!$P$" + slotRow + ")", feed.row, feed.poles); }
    return null;
  }).then(function () { return SAZ.registryFreeRow(ctx); }).then(function (row) {
    var note = "UPS - " + (a.bypass || "Static + Maint.") + " bypass" + (a.internal ? ", internal transformer " + vin + " -> " + vout : "");
    return SAZ.registryWriteRow(ctx, row, [a.unit, "UPS", vin, inTag, "=" + SAZ.UPS_SHEET + "!$L$" + slotRow + "&\" kVA\"",
      "=N(" + SAZ.UPS_SHEET + "!$P$" + slotRow + ")", SAZ.UPS_SHEET, note, feed.ft, feed.ckt, "", ""]).then(function () {
      return SAZ.withUnprotected(ctx, R.dash, function () {
        R.dash.getRange("N100").values = [["Out Voltage"]];
        R.dash.getRange("N" + row).values = [[a.internal ? vout : ""]];
        return ctx.sync();
      });
    });
  }).then(function () {
    return SAZ.registrySet(ctx, load.tag, { 3: a.unit, 8: "UPS Output", 9: "" });
  }).then(function () {
    if (feed.same && !isU && inTag) { var nl = xfmr || a.unit;   /* the parent breaker may be labelled with the board or its T- transformer */ return SAZ.relabelParentCircuit(ctx, inTag, load.tag, nl).then(function () { return SAZ.relabelParentCircuit(ctx, inTag, "T-" + load.tag, nl); }); }
    if (feed.row && typeof SAZ.lockLinkedRows === "function") { return SAZ.lockLinkedRows(ctx, inTag).catch(function () { return null; }); }
    return 0;
  }).then(function () { return SAZ.linkSrcFaultToNode(ctx, load.tag, a.unit); })
    .then(function () { return SAZ.healUpsDemand(ctx); })
    .then(function () { return SAZ.healUpsAic(ctx); })
    .then(function () { return { tag: a.unit, input: inTag, load: load.tag, row: slotRow, vin: vin, vout: vout, internal: !!a.internal, xfmr: xfmr, ckt: feed.row ? feed.ckt : "" }; });
};

SAZ.upsRemove = function (ctx, tag) {
  return SAZ.registryFull(ctx).then(function (R) {
    var me = R.rows.filter(function (r) { return r.tag.toLowerCase() === tag.toLowerCase(); })[0];
    if (!me) { throw new Error(tag + " is not in the distribution."); }
    var kids = R.rows.filter(function (r) { return r.fed.toLowerCase() === tag.toLowerCase(); });
    var chain = Promise.resolve();
    kids.forEach(function (k, i) {
      chain = chain.then(function () {
        return SAZ.registrySet(ctx, k.tag, { 3: me.fed, 8: i === 0 ? me.ft : "", 9: i === 0 ? me.ckt : "" });
      }).then(function () {
        if (i === 0 && me.fed) { var pv = R.rows.filter(function (r) { return r.tag.toLowerCase() === me.fed.toLowerCase(); })[0], num = function (x) { var m = /([0-9]+)/.exec(String(x || "")); return m ? Number(m[1]) : 0; }; var lab = (pv && num(pv.volts) && num(k.volts) && num(pv.volts) !== num(k.volts)) ? "T-" + k.tag : k.tag; return SAZ.relabelParentCircuit(ctx, me.fed, tag, lab).then(function () { return SAZ.relabelParentCircuit(ctx, me.fed, "T-" + tag, lab); }); }
        return null;
      });
    });
    return chain.then(function () {
      return SAZ.withUnprotected(ctx, R.dash, function () {
        R.dash.getRange("A" + me.row + ":N" + me.row).clear(Excel.ClearApplyTo.contents);
        return ctx.sync();
      });
    }).then(function () {
      var c2 = Promise.resolve();
      kids.forEach(function (k) { c2 = c2.then(function () { return SAZ.restoreSrcFault(ctx, k.tag, me.fed); }); });
      return c2;
    }).then(function () { return SAZ.upsRows(ctx); }).then(function (rows) {
      var h = rows.filter(function (x) { return x.tag.toLowerCase() === tag.toLowerCase(); })[0];
      if (!h) { return null; }
      var sh = ctx.workbook.worksheets.getItem(SAZ.UPS_SHEET);
      return SAZ.withUnprotected(ctx, sh, function () {
        ["B" + h.row + ":E" + h.row, "K" + h.row, "V" + h.row].forEach(function (x) { sh.getRange(x).clear(Excel.ClearApplyTo.contents); });
        sh.getRange("G" + h.row).formulas = [["=IF($B" + h.row + "=\"\",\"\",0)"]];
        return ctx.sync();
      });
    }).then(function () { return { moved: kids.map(function (k) { return k.tag; }), to: me.fed }; });
  });
};

/* Distribution Map: a board placed under a UPS, or a UPS input moved */
(function () {
  var base = SAZ.mapApply;
  if (typeof base !== "function" || base._ups) { return; }
  SAZ.mapApply = function (ctx, req) {
    var C = String(req.tag || "").trim(), P = String(req.fed || "").trim();
    return SAZ.registryFull(ctx).then(function (R) {
      var find = function (t) { return R.rows.filter(function (r) { return r.tag.toLowerCase() === String(t).toLowerCase(); })[0]; };
      var cr = find(C), pr = P ? find(P) : null;
      if (cr && SAZ.isUpsType(cr.type)) {
        if (pr && !SAZ.isBoardType(pr.type)) { return Promise.reject(new Error("A UPS input must come from a board or the utility, not from " + P + ".")); }
        return SAZ.isUtilityTag(ctx, P).then(function (isU) {
          return SAZ.registrySet(ctx, C, { 3: P, 8: P ? (isU ? "Utility / Service" : "Sub-Feed") : "" });
        }).then(function () { return SAZ.upsRows(ctx); }).then(function (rows) {
          var h = rows.filter(function (x) { return x.tag.toLowerCase() === C.toLowerCase(); })[0];
          if (!h) { return null; }
          var sh = ctx.workbook.worksheets.getItem(SAZ.UPS_SHEET);
          return SAZ.withUnprotected(ctx, sh, function () { sh.getRange("C" + h.row).values = [[P]]; return ctx.sync(); });
        }).then(function () { return C + " input set to " + (P || "(none)") + ". Assign its breaker on " + (P || "the parent") + " by hand (Move Circuit / circuit entry)."; });
      }
      if (pr && SAZ.isUpsType(pr.type)) {
        return SAZ.mapDetach(ctx, C).then(function () {
          return SAZ.registrySet(ctx, C, { 3: P, 8: "UPS Output", 9: "" });
        }).then(function () { return SAZ.linkSrcFaultToNode(ctx, C, P); })
          .then(function () { return SAZ.healUpsDemand(ctx); })
          .then(function () { return C + " is now fed from UPS " + P + "."; });
      }
      return base.call(SAZ, ctx, req);
    });
  };
  SAZ.mapApply._ups = true; SAZ.mapApply._xfer = true;
})();

SAZ.uiUpsSystem = function (ctx, keep, err) {
  /* v2.10: explicit input board + voltage definition (no override boxes - the
     rating / SCCR overrides live on UPS \u25B8 UPS SYSTEMS columns K and V). */
  var SAME = "(output board's current feed)", MATCH = "(same as input board)";
  return SAZ.moduleState(ctx).then(function (st) {
    if (!st.UPS) { return SAZ.setModuleDeployed(ctx, "UPS", true); }
    return null;
  }).then(function () { return SAZ.ensureUpsUnit(ctx); }).then(function (units) {
    var z = [];   /* sequential: parallel syncs on one context can read before load */ return SAZ.registryFull(ctx).then(function (a) { z[0] = a; return SAZ.voltChoices(ctx); }).then(function (b) { z[1] = b; return SAZ.utilTags(ctx); }).then(function (c) { z[2] = c;
      var R = z[0], VC = z[1], UT = z[2];
      var boards = R.rows.filter(function (r) { return SAZ.isBoardType(r.type); });
      var placed = R.rows.filter(function (r) { return SAZ.isUpsType(r.type); });
      var free = units.filter(function (u) { return !R.rows.some(function (r) { return r.tag.toLowerCase() === u.tag.toLowerCase(); }); });
      if (!boards.length) { return SAZ.notify("No boards yet", "Create the critical (output) board the UPS will feed first."); }
      var actions = (free.length ? ["Place a UPS in the distribution"] : []).concat(placed.map(function (s) { return "Remove " + s.tag + " (re-feed its loads from " + (s.fed || "nothing") + ")"; }));
      if (!actions.length) { return SAZ.notify("No free UPS unit", "Every unit on UPS \u25B8 UPS UNITS is already placed. Add another unit tag there first."); }
      keep = keep || {};
      var outs = boards.filter(function (b) { var p = R.rows.filter(function (r) { return r.tag.toLowerCase() === b.fed.toLowerCase(); })[0]; return !(p && SAZ.isUpsType(p.type)); });
      var inputs = [SAME].concat(UT).concat(boards.map(function (b) { return b.tag + "  (" + (b.volts || "voltage not set") + ")"; }));
      return SAZ.form("New UPS System",
        "Places a UPS between an input board and the critical board it serves. Rating, input / output OCPD and SCCR size themselves on UPS \u25B8 UPS SYSTEMS (overrides are on that sheet).",
        [ { id: "act", label: "Action", type: "select", options: actions, value: keep.act || actions[0] },
          { id: "unit", label: "UPS unit", type: "select", options: free.length ? free.map(function (u) { return u.tag; }) : ["-"], value: keep.unit || (free[0] ? free[0].tag : "-") },
          { id: "load", label: "Output (critical) board", type: "select", options: (outs.length ? outs : boards).map(function (b) { return b.tag; }), value: keep.load || (outs[0] || boards[0]).tag },
          { id: "input", label: "Input from (board feeding the UPS)", type: "select", options: inputs, value: keep.input || SAME,
            hint: "Default keeps the output board's present breaker. Pick another board to feed the UPS from a new breaker there (first open slot)." },
          { id: "upsIn", label: "UPS input voltage", type: "select", options: [MATCH].concat(VC.list), value: keep.upsIn || MATCH },
          { id: "internal", label: "Internal transformer (UPS input V -> output board V)", type: "select", options: ["No", "Yes"], value: keep.internal || "No",
            hint: "No: the UPS input runs at the output board voltage. Yes: the UPS converts its input voltage to the output board voltage itself." },
          { id: "autoX", label: "If the input board voltage differs from the UPS input", type: "select", options: ["Create a transformer ahead of the UPS", "Stop and tell me"], value: keep.autoX || "Create a transformer ahead of the UPS",
            hint: "The transformer T-<UPS> is added to the Transformers schedule and the Distribution Map automatically." },
          { id: "bypass", label: "Bypass", type: "select", options: ["Static + Maint.", "Static only", "None"], value: keep.bypass || "Static + Maint.",
            hint: "With a bypass the upstream fault passes through; None limits the output to the inverter." } ],
        err || "", 85).then(function (ans) {
          if (!ans.ok) { return null; }
          var v = ans.values;
          if (/^Remove\s+/.test(v.act)) {
            var t = v.act.replace(/^Remove\s+/, "").replace(/\s+\(.*$/, "");
            return SAZ.upsRemove(ctx, t).then(function (o) {
              return SAZ.notify("UPS removed", t + " taken out of the distribution. " + (o.moved.length ? o.moved.join(", ") + " fed from " + (o.to || "nothing") + " again." : ""));
            });
          }
          if (!v.unit || v.unit === "-") { return SAZ.uiUpsSystem(ctx, v, "Pick a UPS unit."); }
          var a = { unit: v.unit, load: v.load, bypass: v.bypass,
            input: v.input === SAME ? "" : String(v.input).replace(/\s+\(.*$/, "").trim(),
            upsIn: v.upsIn === MATCH ? "" : v.upsIn, internal: v.internal === "Yes", autoX: /^Create/.test(v.autoX) };
          return SAZ.upsCreate(ctx, a).then(function (o) {
            var sh = ctx.workbook.worksheets.getItem(SAZ.UPS_SHEET); sh.visibility = Excel.SheetVisibility.visible; sh.activate();
            return ctx.sync().then(function () {
              return SAZ.notify("UPS placed", (o.input || "(none)") + (o.xfmr ? "  ->  " + o.xfmr : "") + "  ->  " + o.tag + "  ->  " + o.load + ".\n" +
                "UPS input " + (o.vin || "-") + ", output " + (o.vout || "-") + (o.internal ? " (internal transformer)" : "") + (o.ckt ? ".  Breaker on " + o.input + ": ckt " + o.ckt : "") + ".\n\n" +
                "Rating, input / output OCPD and SCCR size themselves on UPS \u25B8 UPS SYSTEMS. Run Refresh Feeders and Refresh Short Circuit to update the studies.");
            });
          }).catch(function (e) { return SAZ.uiUpsSystem(ctx, v, e.message || String(e)); });
        });
    });
  });
};
SAZ.utilTags = function (ctx) {
  var n = ctx.workbook.names.getItemOrNullObject("Util_Tags"); n.load("isNullObject");
  return ctx.sync().then(function () {
    if (n.isNullObject) { return []; }
    var r = n.getRange(); r.load("values");
    return ctx.sync().then(function () { return r.values.reduce(function (a, x) { return a.concat(x); }, []).map(function (t) { return String(t || "").trim(); }).filter(Boolean); });
  });
};

/* ==========================================================
   v2.8.1  PANEL LOAD RECONCILIATION - automatic "Entered on Panel".
   Registry column M carries each board's OWN connected kVA (its circuit
   categories K6:K13, i.e. total less sub-panels / feed-through), linked
   live like the Std AIC in column K. The Dashboard reconciliation looks it
   up by tag, so Load Calc "assigned" vs "entered on panel" compares the same
   thing and no hand entry is needed.
   ========================================================== */
SAZ.REG_OWN_KVA_COL = "M";
SAZ.healRegistryOwnKva = function (ctx) {
  var dash = ctx.workbook.worksheets.getItem("Dashboard");
  var tags = dash.getRange("A101:A350"); tags.load("values");
  var sheets = ctx.workbook.worksheets; sheets.load("items/name");
  return ctx.sync().then(function () {
    var byName = {};
    sheets.items.forEach(function (s) { byName[s.name.toLowerCase()] = s.name; });
    var probes = tags.values.map(function (r) {
      var t = String(r[0] || "").trim();
      var nm = t ? byName[t.toLowerCase()] : null;
      if (!nm || /^TPL_/i.test(nm)) { return null; }
      var c = ctx.workbook.worksheets.getItem(nm).getRange("C13"); c.load("values");
      return { name: nm, c: c };
    });
    return ctx.sync().then(function () {
      var q = function (p) { return "'" + p.name.replace(/'/g, "''") + "'!"; };
      var isB = function (p) { return p && String(p.c.values[0][0] || "").indexOf("Std Device AIC") === 0; };
      /* M own conn. kVA | O Main/Bus override (A, blank = none) | P effective Main/Bus rating (A) | Q demand FLA (A)
         (N = UPS output voltage, written by upsCreate - left alone here) */
      var f = probes.map(function (p) { return isB(p) ? ["=SUM(" + q(p) + "$K$6:$K$13)"] : [""]; });
      var opq = probes.map(function (p) {
        if (!isB(p)) { return ["", "", ""]; }
        return ["=IF(N(" + q(p) + "$E$12)>0," + q(p) + "$E$12,\"\")", "=IFERROR(N(" + q(p) + "$D$12),\"\")", "=IFERROR(N(" + q(p) + "$H$10),\"\")"];
      });
      return SAZ.withUnprotected(ctx, dash, function () {
        dash.getRange("M100").values = [["Panel Conn. kVA"]];
        dash.getRange("O100:Q100").values = [["Rating Ovr (A)", "Rating (A)", "Demand (A)"]];
        var m = dash.getRange("M101:M350"), o = dash.getRange("O101:Q350");
        m.formulas = f; o.formulas = opq;
        m.numberFormat = f.map(function () { return ["#,##0.00"]; });
        o.numberFormat = opq.map(function () { return ["#,##0", "#,##0", "#,##0.00"]; });
        [m, o].forEach(function (x) { x.format.font.name = "Segoe UI"; x.format.font.size = 10; x.format.font.color = "#FF0000"; x.format.horizontalAlignment = "Center"; });
        return ctx.sync();
      });
    });
  });
};
(function () {
  var h = SAZ.healRegistryAic;
  SAZ.healRegistryAic = function (ctx) {
    return h.apply(SAZ, arguments).then(function (v) { return SAZ.healRegistryOwnKva(ctx).then(function () { return v; }); });
  };
})();

/* ==========================================================
   v2.10  MAIN/BUS OVERRIDE  -  E12 on every schedule (E11 = 'Override' caption).
   D12 = IF(N(E12)>0, E12, calc); calc min kept in engine row 41 (AQ / BH on MCC).
   Boards deployed from an older template are upgraded in place the first time
   the registry heals (Refresh Feeders / Studies / Distribution Map refresh).
   ========================================================== */
SAZ.healMainsOverride = function (ctx, name) {
  var sh = ctx.workbook.worksheets.getItem(name);
  var d = sh.getRange("D12"), lab = sh.getRange("C12"), mcc = sh.getRange("BH30"), aq = sh.getRange("AQ30");
  d.load("formulas"); lab.load("values"); mcc.load("formulas"); aq.load("formulas");
  return ctx.sync().then(function () {
    var f = String(d.formulas[0][0] || "");
    if (String(lab.values[0][0] || "").indexOf("Main/Bus") !== 0 || f.indexOf("$E$12") >= 0 || f.charAt(0) !== "=") { return false; }
    var X = /\$BH\$30/.test(f) ? "BH" : "AQ", Lb = X === "BH" ? "BG" : "AP";
    return SAZ.withUnprotected(ctx, sh, function () {
      sh.getRange(X + "41").formulas = [[f]];
      sh.getRange(Lb + "41").values = [["Main/Bus \u2014 calc. min (A)"]];
      var c = sh.getRange("E11:F11"); c.unmerge(); c.merge(false); c.values = [["Override", ""]];
      c.format.font.name = "Segoe UI"; c.format.font.size = 8; c.format.font.italic = true; c.format.font.color = "#7F7F7F";
      c.format.horizontalAlignment = "Center"; c.format.verticalAlignment = "Bottom";
      var e = sh.getRange("E12:F12"); e.unmerge(); e.values = [["", ""]]; e.merge(false);
      e.format.fill.color = "#D9D9D9"; e.format.font.name = "Segoe UI"; e.format.font.size = 10; e.format.font.italic = false; e.format.font.color = "#0000FF";
      e.format.horizontalAlignment = "Center"; e.format.verticalAlignment = "Center"; e.numberFormat = [["#,##0", "#,##0"]];
      ["EdgeLeft", "EdgeRight", "EdgeTop", "EdgeBottom"].forEach(function (k) { var b = e.format.borders.getItem(k); b.style = "Continuous"; b.weight = "Thin"; b.color = "#000000"; }); var gl = sh.getRange("G11").format.borders.getItem("EdgeLeft"); gl.style = "Continuous"; gl.weight = "Thin"; gl.color = "#000000";   /* merged caption must not strip the Feeder Sets box edge */
      e.format.protection.locked = false;
      var dv = sh.getRange("E12").dataValidation; dv.clear(); dv.rule = { list: { inCellDropDown: true, source: "=OCPD_List" } }; dv.errorAlert = { showAlert: false };
      var cf = e.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      cf.custom.rule.formula = "=AND(N($E$12)>0,N($E$12)<N($" + X + "$41))"; cf.custom.format.font.color = "#FF0000"; cf.custom.format.fill.color = "#FFC7CE";
      d.formulas = [["=IF(N($E$12)>0,$E$12,$" + X + "$41)"]];
      return ctx.sync().then(function () { return true; });
    });
  }).catch(function () { return false; });
};
(function () {
  var h = SAZ.healRegistryOwnKva;
  if (typeof h !== "function" || h._mains) { return; }
  SAZ.healRegistryOwnKva = function (ctx) {
    var dash = ctx.workbook.worksheets.getItem("Dashboard"), tags = dash.getRange("A101:B350"); tags.load("values");
    var ws = ctx.workbook.worksheets; ws.load("items/name");
    return ctx.sync().then(function () {
      var names = {}; ws.items.forEach(function (x) { names[x.name.toLowerCase()] = x.name; });
      var list = tags.values.filter(function (r) { return SAZ.isBoardType(String(r[1] || "")) && names[String(r[0] || "").trim().toLowerCase()]; })
        .map(function (r) { return names[String(r[0]).trim().toLowerCase()]; });
      var chain = Promise.resolve();
      list.forEach(function (nm) { chain = chain.then(function () { return SAZ.healMainsOverride(ctx, nm); }); });
      return chain;
    }).then(function () { return h.call(SAZ, ctx); });
  };
  SAZ.healRegistryOwnKva._mains = true;
})();

/* ==========================================================
   v2.10  PER-STUDY REFRESH  (ribbon: Voltage Drop / Short Circuit menus;
   Arc Flash uses the existing study rebuild). Same pattern as Refresh Feeders:
   heal the registry links, full recalc, open the sheet, summarise the flags.
   ========================================================== */
SAZ.STUDY_REFRESH = {
  vd: { sheet: "Voltage Drop",  first: 6, last: 260, node: "A", status: "K", title: "Refresh Voltage Drop", noun: "bus",
        none: "No buses yet - Voltage Drop derives one row per bus from the equipment registry and the Feeder Schedule.",
        hint: "Feeder lengths, material and conduit are set in the Distribution Map (stored on the registry, read by the Feeder Schedule) and branch lengths in each schedule's Length Ft column." },
  sc: { sheet: "Short Circuit", first: 6, last: 255, node: "A", status: "T", title: "Refresh Short Circuit", noun: "bus",
        none: "No buses yet - Short Circuit derives one row per bus from the equipment registry.",
        hint: "Device AIC comes from each schedule's Std Device AIC (override in E13). Run Refresh Arc Flash after this to carry the new fault currents into the arc-flash table." }
};
SAZ.refreshStudySheet = function (ctx, key) {
  var d = SAZ.STUDY_REFRESH[key];
  var sheets = ctx.workbook.worksheets; sheets.load("items/name,items/visibility");
  return ctx.sync().then(function () {
    var hit = sheets.items.filter(function (s) { return s.name.toLowerCase() === d.sheet.toLowerCase(); })[0];
    if (!hit) { return SAZ.notify(d.title, "There is no \"" + d.sheet + "\" sheet in this workbook. Create it from the ribbon first."); }
    var heal = typeof SAZ.healRegistryAic === "function" ? SAZ.healRegistryAic(ctx).catch(function () { return null; }) : Promise.resolve(null);
    return heal.then(function () {
      if (hit.visibility !== Excel.SheetVisibility.visible) { hit.visibility = Excel.SheetVisibility.visible; }
      hit.activate();
      ctx.workbook.application.calculate(Excel.CalculationType.fullRebuild);
      var nodes = hit.getRange(d.node + d.first + ":" + d.node + d.last), st = hit.getRange(d.status + d.first + ":" + d.status + d.last);
      nodes.load("values"); st.load("values");
      return ctx.sync().then(function () {
        var n = 0, warn = [], fail = 0;
        nodes.values.forEach(function (r, i) {
          var t = String(r[0] || "").trim(), s = String(st.values[i][0] || "").trim(); if (!t || !s) { return; }
          n++;
          /* a bus row always carries a status; other text in the node column (notes) is skipped */
          if (/^FAIL/i.test(s)) { fail++; warn.push(t + " - " + s); }
          else if (s.charAt(0) === "\u26A0") { warn.push(t + " - " + s.replace(/^\u26A0\s*/, "")); }
        });
        if (!n) { return SAZ.notify(d.title, "Recalculated. " + d.none); }
        var msg = "Recalculated. " + n + " " + d.noun + (n === 1 ? "" : "es") + "." + (fail ? "  " + fail + " FAIL." : "");
        msg += warn.length ? "\n\nFlagged:\n" + warn.slice(0, 12).join("\n") + (warn.length > 12 ? "\n... and " + (warn.length - 12) + " more." : "") : "\nNo warnings.";
        return SAZ.notify(d.title, msg + "\n\n" + d.hint);
      });
    });
  });
};

/* v2.10  Help ▸ Instructions: open the manual at a chapter band ("CHAPTER nn · ...", column A). */
SAZ.goChapter = function (ctx, n) {
  return Promise.resolve(SAZ.goOrCreate(ctx, "Instructions")).then(function () {
    var sh = ctx.workbook.worksheets.getItem("Instructions");
    var col = sh.getRange("A1:A800"); col.load("values");
    return ctx.sync().then(function () {
      var key = "CHAPTER " + n, row = 0;
      for (var i = 0; i < col.values.length; i++) { if (String(col.values[i][0] || "").indexOf(key) === 0) { row = i + 1; break; } }
      sh.activate();
      if (!row) { return SAZ.notify("Instructions", "Chapter " + n + " was not found - this project may carry an older Instructions tab. Copy the current one from the master / seed."); }
      sh.getRange("A" + Math.min(row + 30, 800)).select();
      sh.getRange("A" + row).select();
      return ctx.sync();
    });
  });
};
