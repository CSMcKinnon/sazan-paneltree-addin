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
  regRange:    "A100:J349"   /* fallback only - resolved from Equip_Registry at runtime */
};

var SAZ = {};

/* ---------- dialog helpers: replace VBA MsgBox / InputBox ---------- */

SAZ.dialog = function (params) {
  return new Promise(function (resolve) {
    var qs = Object.keys(params).map(function (k) {
      return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]);
    }).join("&");
    var url = CFG.host + "/dialog.html?" + qs;
    Office.context.ui.displayDialogAsync(url,
      { height: 42, width: 30, displayInIframe: true },
      function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) { resolve({ ok: false }); return; }
        var dlg = res.value;
        dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          var out = { ok: false };
          try { out = JSON.parse(arg.message); } catch (e) { }
          dlg.close();
          resolve(out);
        });
        dlg.addEventHandler(Office.EventType.DialogEventReceived, function () {
          resolve({ ok: false });
        });
      });
  });
};

SAZ.notify = function (title, text) {
  return SAZ.dialog({ mode: "notify", title: title, text: text });
};

SAZ.prompt = function (title, label, dflt, required) {
  return SAZ.dialog({ mode: "prompt", title: title, label: label,
                      "default": dflt || "", required: required ? "1" : "0", text: "" });
};

/* ---------- run wrapper: guarantees event.completed() ---------- */

SAZ.run = function (name, body) {
  return function (event) {
    Excel.run(function (ctx) { return body(ctx); })
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
          "Add " + tplName + " to this workbook, or check that " + CFG.projectFile +
          " is published to " + CFG.host + ".");
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

/* ---------- registry write: the missing link that fills the Panel Tree ---------- */

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
     the sheet appeared but never reached the registry or the Panel Tree.
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

SAZ.uiCreate = function (ctx, tplName, typeLabel) {
  return SAZ.requireSheets(ctx, SAZ.CLONE_REQUIRES).then(function (ok) { if (!ok) { return null; }
  return SAZ.prompt("New " + typeLabel, "Equipment tag", "", true).then(function (ans) {
    if (!ans.ok) { return null; }
    var tag = SAZ.sanitise(ans.value);
    if (!tag) { return SAZ.notify("Invalid tag", "That tag contains no usable characters."); }
    return SAZ.ensureTemplate(ctx, tplName).then(function (tpl) {
      return SAZ.uniqueSheetName(ctx, tag).then(function (finalName) {
        return SAZ.cloneSheet(ctx, tpl, finalName);
      }).then(function (finalName) {
        return SAZ.registerFromGeo(ctx, finalName, typeLabel).then(function () {
          ctx.workbook.worksheets.getItem(finalName).activate();
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
  return SAZ.uiRebuildGeoMap(ctx);
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
function ToggleMatrix (event) { SAZ.run("Toggle Load Matrix", function (ctx) {
  return SAZ.uiToggleLoadMatrix(ctx);
})(event); }

/* ---------- associations ---------- */

Office.onReady(function () {
  Office.actions.associate("NewProject", NewProject);
  Office.actions.associate("GoInstructions", GoInstructions);
  Office.actions.associate("GoDashboard", GoDashboard);
  Office.actions.associate("NewLoadCalc", NewLoadCalc);
  Office.actions.associate("GoLoadCalc", GoLoadCalc);
Office.actions.associate("AddSectionRow", AddSectionRow);
Office.actions.associate("TrimSectionRows", TrimSectionRows);
Office.actions.associate("ToggleSection", ToggleSection);
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
  Office.actions.associate("NewReport", NewReport);
  Office.actions.associate("GoReport", GoReport);
  Office.actions.associate("GoRecon", GoRecon);
  Office.actions.associate("ToggleSizing", ToggleSizing);
  Office.actions.associate("ToggleMatrix", ToggleMatrix);
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

/* Which section is the cursor sitting in? Returns null if none. */
SAZ.lcSectionAtCursor = function (ctx) {
  var sel = ctx.workbook.getSelectedRange();
  sel.load("address");
  var probes = SAZ.lcSections.map(function (s) {
    var t = ctx.workbook.tables.getItemOrNullObject(s.table);
    var r = t.getRange().getIntersectionOrNullObject(sel);
    r.load("isNullObject");
    t.load("isNullObject");
    return { spec: s, table: t, hit: r };
  });
  return ctx.sync().then(function () {
    for (var i = 0; i < probes.length; i++) {
      if (!probes[i].table.isNullObject && !probes[i].hit.isNullObject) return probes[i];
    }
    return null;
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
        throw new Error("Project seed not reachable (" + res.status + "). Check that " +
          CFG.projectFile + " is published to " + CFG.host + ".");
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
        if (String(subRow[c2] || "").trim().toLowerCase() === "kva") { kc = c2; break; }
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
        firstRow: hr + 3, lastRow: stop,
        source: "derived"
      };
    });
  });
};


/* Write (or refresh) one feed-through row. Re-running with the same tag
   updates that row in place rather than consuming another slot. */
SAZ.feedThruAppend = function (ctx, sheetName, childTag, demandRef) {
  return SAZ.feedThruBlock(ctx, sheetName).then(function (b) {
    var sheet = ctx.workbook.worksheets.getItem(sheetName);
    var span = sheet.getRange(b.descCol + b.firstRow + ":" + b.descCol + b.lastRow);
    span.load("values");
    return ctx.sync().then(function () {
      var slot = -1;
      for (var i = 0; i < span.values.length; i++) {
        var cur = String(span.values[i][0] || "").trim();
        if (cur === "" || cur.toLowerCase() === childTag.toLowerCase()) {
          slot = b.firstRow + i; break;
        }
      }
      if (slot < 0) {
        throw new Error("All " + (b.lastRow - b.firstRow + 1) + " feed-through slots on \"" +
          sheetName + "\" are in use. Resize the block or remove an entry.");
      }
      sheet.getRange(b.descCol + slot).values  = [[childTag]];
      sheet.getRange(b.kvaCol  + slot).formulas = [[demandRef]];
      return ctx.sync().then(function () { return { row: slot, block: b }; });
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
   so the pass-through tracks the child as its circuits change. */
SAZ.uiFeedThru = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    var primary = active.name;
    return SAZ.prompt("Feed-Through",
      "Downstream board whose load feeds THROUGH to " + primary, "", true)
      .then(function (ans) {
        if (!ans.ok) { return null; }
        var childTag = SAZ.sanitise(ans.value);
        if (!childTag) { return SAZ.notify("Invalid tag", "No usable characters in that tag."); }
        if (childTag.toLowerCase() === primary.toLowerCase()) {
          return SAZ.notify("Cannot feed itself",
            "\"" + primary + "\" cannot be its own feed-through source.");
        }
        var sheets = ctx.workbook.worksheets;
        sheets.load("items/name");
        return ctx.sync().then(function () {
          var child = sheets.items.filter(function (s) {
            return s.name.toLowerCase() === childTag.toLowerCase();
          })[0];
          if (!child) {
            return SAZ.notify("No such board",
              "There is no sheet named \"" + childTag + "\". Create it first.");
          }
          return SAZ.geo(ctx, child.name).then(function (cg) {
            if (!cg.DEMAND_CELL) {
              throw new Error("\"" + child.name + "\" has no DEMAND_CELL in its GEO_MAP.");
            }
            var ref = SAZ.sheetRef(child.name, cg.DEMAND_CELL);
            return SAZ.feedThruAppend(ctx, primary, child.name, ref).then(function (res) {
              return SAZ.updateRegistryFedFrom(ctx, child.name, primary).then(function (reg) {
                return SAZ.notify("Feed-through added",
                  child.name + " added to " + primary + " at row " + res.row +
                  " (no breaker used)." +
                  (reg ? "" : "\n\nNote: " + child.name +
                   " is not in the equipment registry, so Fed From was not updated."));
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

/* Write one circuit row on the parent. Column mappings live here only. */
SAZ.addParentCircuitAt = function (ctx, parentTag, descTag, kvaRef, targetRow) {
  return SAZ.geo(ctx, parentTag).then(function (g) {
    var sheet = ctx.workbook.worksheets.getItem(parentTag);
    var put = function (key, val, isFormula) {
      var col = SAZ.geoUsable(g[key]);
      if (!col) { return; }
      var rg = sheet.getRange(col + targetRow);
      if (isFormula) { rg.formulas = [[val]]; } else { rg.values = [[val]]; }
    };
    put("CKT_DESC_COL",  descTag, false);
    put("CKT_KVA_COL",   kvaRef,  true);
    put("CKT_QTY_COL",   1,       false);
    put("CKT_POLES_COL", 3,       false);
    put("CKT_CLASS_COL", "Panel", false);
    return ctx.sync().then(function () { return targetRow; });
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
SAZ.linkEquipment = function (ctx, parentTag, childTag, targetRow) {
  if (parentTag.toLowerCase() === childTag.toLowerCase()) {
    return SAZ.notify("Cannot feed itself",
      "\"" + parentTag + "\" cannot feed itself.");
  }
  var sheets = ctx.workbook.worksheets;
  sheets.load("items/name");
  return ctx.sync().then(function () {
    var byName = function (t) {
      return sheets.items.filter(function (s) {
        return s.name.toLowerCase() === t.toLowerCase();
      })[0];
    };
    var p = byName(parentTag), c = byName(childTag);
    if (!p) { return SAZ.notify("Parent not found", "No sheet named \"" + parentTag + "\"."); }
    if (!c) { return SAZ.notify("Child not found",  "No sheet named \"" + childTag + "\"."); }
    var P = p.name, C = c.name;
    return Promise.all([ SAZ.geo(ctx, P), SAZ.geo(ctx, C) ]).then(function (gs) {
      var gp = gs[0], gc = gs[1];
      var childDemand = SAZ.geoUsable(gc.DEMAND_CELL);
      var childSrc    = SAZ.geoUsable(gc.SRCFAULT_CELL);
      var parentBus   = SAZ.geoUsable(gp.BUSFAULT_CELL);
      if (!childDemand || !childSrc || !parentBus) {
        throw new Error("GEO_MAP incomplete: need DEMAND_CELL and SRCFAULT_CELL on \"" +
          C + "\" and BUSFAULT_CELL on \"" + P + "\".");
      }
      var demandRef = SAZ.sheetRef(C, childDemand);
      return Promise.all([
        SAZ.geoCellValue(ctx, P, "ENG_VLL_CELL"),
        SAZ.geoCellValue(ctx, C, "ENG_VLL_CELL"),
        SAZ.geoCellValue(ctx, C, "ENG_PH_CELL")
      ]).then(function (vals) {
        var pV = vals[0] || 0, cV = vals[1] || 0, cPh = vals[2] || 0;
        var needsTx = Math.abs(pV - cV) > SAZ.VOLT_MATCH_TOLERANCE;
        var childSheet = ctx.workbook.worksheets.getItem(C);
        var place = function (desc) {
          return (targetRow)
            ? SAZ.addParentCircuitAt(ctx, P, desc, demandRef, targetRow)
            : SAZ.addParentCircuit(ctx, P, desc, demandRef);
        };
        if (!needsTx) {
          childSheet.getRange(childSrc).formulas = [[ SAZ.sheetRef(P, parentBus) ]];
          return ctx.sync()
            .then(function () { return place(C); })
            .then(function (row) {
              return SAZ.updateRegistryFedFrom(ctx, C, P).then(function () {
                return SAZ.notify("Linked",
                  C + " linked to " + P + " at the same voltage (row " + row + ").");
              });
            });
        }
        /* Voltages differ, so a transformer is required.

           DO NOT create one here. The Transformers schedule is SELF-DERIVING:
           BL4 spills one row per registry edge where the voltage changes, and
           columns A-F of every data row read that spill through
           INDEX($BL$4#,ROW()-5,n). Only R, S, T and U are inputs. The VBA
           appended a row by hand and this port copied it, which would have
           overwritten six derived formulas on the very row detection wants to
           use. This is the same rule already honoured for Short Circuit and
           Voltage Drop - registry-fed sheets are READ ONLY to handlers.

           So: update the registry, let the sheet derive the row, find it by
           tag, and point the child source fault at its secondary. */
        if (!sheets.items.some(function (s) { return s.name === SAZ.TX_SHEET; })) {
          throw new Error("\"" + P + "\" is " + pV + "V and \"" + C + "\" is " + cV +
            "V, so a transformer is required, but there is no " + SAZ.TX_SHEET +
            " sheet. Create it first with New Transformers.");
        }
        return SAZ.updateRegistryFedFrom(ctx, C, P).then(function () {
          ctx.workbook.application.calculate(Excel.CalculationType.full);
          return ctx.sync();
        }).then(function () {
          return SAZ.txRowForTag(ctx, "T-" + C);
        }).then(function (tr) {
          if (!tr) {
            throw new Error("\"" + C + "\" is " + cV + "V under \"" + P + "\" at " +
              pV + "V, but no transformer row appeared on " + SAZ.TX_SHEET +
              ". The usual cause is a registry voltage string the config list does " +
              "not recognise - see the check at " + SAZ.TX_SHEET + "!BS2.");
          }
          childSheet.getRange(childSrc).formulas =
            [[ "='" + SAZ.TX_SHEET.replace(/'/g, "''") + "'!$T$" + tr ]];   /* T = Sec Fault Sym (A) on the printed schedule */
          return ctx.sync()
            .then(function () { return place("T-" + C); })
            .then(function (row) {
              return SAZ.notify("Linked via transformer",
                C + " (" + cV + "V) linked to " + P + " (" + pV + "V).\n\n" +
                SAZ.TX_SHEET + " row " + tr + " was derived automatically from the " +
                "registry - this command did not write it. Parent circuit at row " +
                row + ".\n\nIf the default impedance from Xfmr_DefZ is not right for " +
                "this unit, set %Z in column T on that row.");
            });
        });
      });
    });
  });
};


/* Link Equipment: prompts for both tags, takes the first open circuit. */
SAZ.uiLink = function (ctx) {
  return SAZ.prompt("Link Equipment", "PARENT tag (the upstream board)", "", true)
    .then(function (a) {
      if (!a.ok) { return null; }
      var parentTag = SAZ.sanitise(a.value);
      if (!parentTag) { return SAZ.notify("Invalid tag", "No usable characters."); }
      return SAZ.prompt("Link Equipment", "CHILD tag (fed from " + parentTag + ")", "", true)
        .then(function (b) {
          if (!b.ok) { return null; }
          var childTag = SAZ.sanitise(b.value);
          if (!childTag) { return SAZ.notify("Invalid tag", "No usable characters."); }
          return SAZ.linkEquipment(ctx, parentTag, childTag);
        });
    });
};

/* Sub-Feed: run FROM the parent with a circuit row selected. The chosen row
   is used instead of the first open one. Confirms before overwriting a row
   that already has a description. */
SAZ.uiSubFeed = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  var sel = ctx.workbook.getSelectedRange();
  sel.load("rowIndex");
  return ctx.sync().then(function () {
    var parentTag = active.name;
    var row = sel.rowIndex + 1;
    return SAZ.geo(ctx, parentTag).then(function (g) {
      var first = Number(g.CKT_FIRST) || 0;
      var lastR = Number(g.CKT_LAST) || 0;
      var dc = SAZ.geoUsable(g.CKT_DESC_COL);
      if (!first || !dc) {
        return SAZ.notify("Not an equipment sheet",
          "\"" + parentTag + "\" has no circuit grid in its GEO_MAP." +
          " Run Sub-Feed from a panelboard, switchboard or MCC.");
      }
      if (row < first || row > lastR) {
        return SAZ.notify("Select a circuit row",
          "Click a cell in the schedule grid (rows " + first + "-" + lastR +
          ") to pick the slot, then run Sub-Feed. You selected row " + row + ".");
      }
      var descCell = ctx.workbook.worksheets.getItem(parentTag).getRange(dc + row);
      descCell.load("values");
      return ctx.sync().then(function () {
        var occupied = String(descCell.values[0][0] || "").trim() !== "";
        var ask = occupied
          ? SAZ.prompt("Row " + row + " is not empty",
              "Row " + row + " currently reads \"" + descCell.values[0][0] +
              "\". Type the child tag to overwrite it, or cancel", "", true)
          : SAZ.prompt("Sub-Feed", "Child board to sub-feed from " + parentTag +
              " at row " + row, "", true);
        return ask.then(function (ans) {
          if (!ans.ok) { return null; }
          var childTag = SAZ.sanitise(ans.value);
          if (!childTag) { return SAZ.notify("Invalid tag", "No usable characters."); }
          return SAZ.linkEquipment(ctx, parentTag, childTag, row);
        });
      });
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

SAZ.setBlockRows = function (ctx, sheetName, firstRow, lastRowIn, wantRows, polesCol) {
  var sheet = ctx.workbook.worksheets.getItem(sheetName);
  if (wantRows < 1) { wantRows = 1; }
  var have = lastRowIn - firstRow + 1;
  if (have < 1) { have = 1; }
  var endRow = firstRow + wantRows - 1;

  if (wantRows === have) { return Promise.resolve({ added: 0, removed: 0, endRow: endRow }); }

  if (wantRows < have) {
    sheet.getRange((endRow + 1) + ":" + lastRowIn).delete(Excel.DeleteShiftDirection.up);
    return ctx.sync().then(function () {
      return { added: 0, removed: have - wantRows, endRow: endRow };
    });
  }

  var newFirst = lastRowIn + 1;
  sheet.getRange(newFirst + ":" + endRow).insert(Excel.InsertShiftDirection.down);
  return ctx.sync().then(function () {
    /* clone the last real row - carries column formats AND calculated formulas */
    sheet.getRange(newFirst + ":" + endRow)
      .copyFrom(sheet.getRange(lastRowIn + ":" + lastRowIn), Excel.RangeCopyType.all);
    return ctx.sync();
  }).then(function () {
    var used = sheet.getUsedRange();
    used.load("columnIndex,columnCount");
    return ctx.sync().then(function () {
      var width = used.columnIndex + used.columnCount;
      var band = sheet.getRangeByIndexes(newFirst - 1, 0, endRow - newFirst + 1, width);
      band.load("formulas");
      return ctx.sync().then(function () {
        /* Blank typed constants by WRITING "" - never Range.clear().
           Every circuit description in column C is a single-cell MERGED area,
           and Office.js throws GeneralException on .clear() against a merged
           cell. The original per-cell .clear() loop therefore failed on any
           panel that had a description in the clone source - i.e. every real
           panel - and it failed AFTER the rows were inserted and cloned,
           leaving the sheet half-resized. An empty-string write succeeds on
           merged and unmerged cells alike. Verified on a scratch panel:
           C40.clear() throws, C40.values=[[""]] succeeds. */
        var f = band.formulas, cleared = 0;
        for (var r = 0; r < f.length; r++) {
          for (var c = 0; c < f[r].length; c++) {
            var cell = f[r][c];
            var isFormula = (typeof cell === "string" && cell.charAt(0) === "=");
            if (cell !== "" && cell !== null && !isFormula) {
              band.getCell(r, c).values = [[""]];
              cleared++;
            }
          }
        }
        return ctx.sync().then(function () { return cleared; });
      });
    });
  }).then(function () {
    if (polesCol && polesCol !== "-") {
      var n = endRow - newFirst + 1, ones = [];
      for (var i = 0; i < n; i++) { ones.push([1]); }
      sheet.getRange(polesCol + newFirst + ":" + polesCol + endRow).values = ones;
      return ctx.sync();
    }
  }).then(function () {
    return { added: wantRows - have, removed: 0, endRow: endRow };
  });
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
    return SAZ.geo(ctx, name).then(function (g) {
      var ckt1    = Number(g.CKT_FIRST) || 28;
      var gridMax = Number(g.CKT_LAST) || ckt1;
      if (gridMax < ckt1) { gridMax = ckt1; }
      var rowsPer = Number(g.CKT_PER_ROW) || 1;
      var mode    = String(g.RESIZE_MODE || "").trim().toUpperCase();
      var cktCol  = SAZ.geoUsable(g.CKT_NUM_COL) || "B";
      var pc      = SAZ.geoUsable(g.CKT_POLES_COL);
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
            return SAZ.setBlockRows(ctx, name, b2F, b2L, newBlk, pc).then(function () {
              return SAZ.setBlockRows(ctx, name, b1F, b1L, newBlk, pc);
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
                  CKT_LAST: String(n2L)
                });
              }).then(function () { return { mode: "TWOBLOCK", perBlock: newBlk, cktLast: n2L }; });
            });
          }
          /* single contiguous block - CKT = 1,2,3... */
          var lastRow = Math.max(ckt1, (ckt1 - 1) + Math.ceil(spaces / rowsPer));
          return SAZ.setBlockRows(ctx, name, ckt1, gridMax, lastRow - ckt1 + 1, pc)
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
                CKT_LAST:   String(lastRow)
              });
            }).then(function () { return { mode: "SINGLE", cktLast: lastRow }; });
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

SAZ.uiResize = function (ctx) {
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    return SAZ.prompt("Resize Panel", "Panel tag to resize", active.name, true)
      .then(function (a) {
        if (!a.ok) { return null; }
        var tag = SAZ.sanitise(a.value);
        if (!tag) { return SAZ.notify("Invalid tag", "No usable characters."); }
        return SAZ.prompt("Resize Panel",
          "New space count (grows or shrinks; panelboards round up to an even split)",
          "42", true).then(function (b) {
            if (!b.ok) { return null; }
            var spaces = parseInt(String(b.value).replace(/[^0-9]/g, ""), 10);
            if (!spaces || spaces < 1) {
              return SAZ.notify("Invalid count", "Enter a whole number of spaces, 1 or more.");
            }
            return SAZ.resizePanel(ctx, tag, spaces).then(function (out) {
              if (!out || !out.res) { return null; }
              return SAZ.notify("Panel resized",
                tag + " resized to " + spaces + " spaces (" + out.res.mode +
                (out.res.perBlock ? ", " + out.res.perBlock + " rows per block" : "") +
                "). Circuit grid now ends at row " + out.res.cktLast + "." +
                "\n\nGEO_MAP: " + out.rb.changed + " key" +
                (out.rb.changed === 1 ? "" : "s") + " repointed.");
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
SAZ.EQUIP_REGISTRY_RANGE = "W7:X19";
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

SAZ.SECTION_REGISTRY_RANGE = "W7:Y19";
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
      sheet.getRange("Y" + entry.flagRow).values = [[turnOn ? "Yes" : "No"]];
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

SAZ.uiToggleSection = function (ctx) {
  /* Same ribbon button, context-aware: on Logging Demand it adds/removes the
     interval-data approach; everywhere else it manages Load Calc sections. */
  var active = ctx.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  return ctx.sync().then(function () {
    if (active.name.toLowerCase() === SAZ.LD_SHEET.toLowerCase()) { return SAZ.uiToggleInterval(ctx); }
    return SAZ.sectionRegistry(ctx).then(function (list) {
      var opt = list.filter(function (e) { return SAZ.SECTION_OPTIONAL_TAG.test(e.label); });
      if (opt.length === 0) { return SAZ.notify("No optional sections", "No Load Calc section is marked (optional)."); }
      var menu = opt.map(function (e) {
        return e.label.replace(SAZ.SECTION_OPTIONAL_TAG, "").trim() + "   [" +
          (String(e.state).toUpperCase() === "YES" ? "IN" : "not added") + "]";
      }).join("\n") + "\n\nBaseline sections are always included.";
      return SAZ.prompt("Add or remove a Load Calc section - enter its number", menu, "", true)
        .then(function (a) {
          if (!a.ok) { return null; }
          var pick = parseInt(String(a.value).replace(/[^0-9]/g, ""), 10);
          var hit = opt.filter(function (e) { return e.num === pick; })[0];
          if (!hit) { return SAZ.notify("Not an optional section", "Enter one of the numbers listed."); }
          var turnOn = String(hit.state).toUpperCase() !== "YES";
          var name = hit.label.replace(SAZ.SECTION_OPTIONAL_TAG, "").trim();
          return SAZ.setSectionState(ctx, hit, turnOn).then(function (did) {
            if (!did) {
              return SAZ.notify("Section not built", "Table \"" + hit.table + "\" does not exist in this workbook.");
            }
            return SAZ.notify(turnOn ? "Section added" : "Section removed",
              name + (turnOn ? " is now included in the load calc." :
                " is collapsed and excluded from every total. Its rows are kept, so adding it back restores any entries."));
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
      var probeCol = (kind === "panelboard") ? "P" : "I";
      var bands = (kind === "panelboard") ? ["O:P", "X:Z"] : ["G:G", "I:J"];
      var probe = active.getRange(probeCol + ":" + probeCol);
      probe.load("columnHidden");
      return ctx.sync().then(function () {
        var hide = !probe.columnHidden;
        bands.forEach(function (b) { active.getRange(b).columnHidden = hide; });
        return ctx.sync().then(function () {
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
      if (kind !== "panelboard") {
        return SAZ.notify("Wrong sheet", "Run this from a panelboard schedule.");
      }
      var probe = active.getRange("C:C");
      probe.load("columnHidden");
      return ctx.sync().then(function () {
        var hide = !probe.columnHidden;
        ["B:J", "AE:AM"].forEach(function (b) { active.getRange(b).columnHidden = hide; });
        return ctx.sync().then(function () {
          return SAZ.notify("Load matrix " + (hide ? "hidden" : "shown"),
            "Classification matrix " + (hide ? "hidden" : "restored") +
            " on " + active.name + ".");
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
