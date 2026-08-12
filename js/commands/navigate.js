/* ==========================================================
   SAZAN PowerCalc  -  js/commands/navigate.js
   Wave 1. Ports GoOrCreate (13 buttons) and GoRecon (1).
   VBA source: VBA_Code row 386 (GoOrCreate), row 1828 (GoRecon).
   ========================================================== */

var SAZ = SAZ || {};

SAZ.navigate = (function () {
  "use strict";

  /* goOrCreate
     Activate the sheet if it is there, otherwise offer to build it.
     Mirrors the VBA exactly: nothing is ever created silently. */
  function goOrCreate(sheetName) {
    return SAZ.excel.run(function (context) {
      return SAZ.excel.exists(context, sheetName).then(function (found) {
        if (found) { return SAZ.excel.activate(context, sheetName); }
        return SAZ.dialog.confirm(
          sheetName,
          "There is no " + sheetName + " sheet in this workbook yet. Create it now?",
          "Create"
        ).then(function (yes) {
          if (!yes) { return false; }
          /* Wave 3 swaps this stub for SAZ.templates.insert(sheetName). */
          return SAZ.dialog.alert(
            sheetName,
            "Sheet creation arrives in Wave 3. For now use the New button on the ribbon."
          ).then(function () { return false; });
        });
      });
    });
  }

  /* goRecon
     Jump to the reconciliation block on the Dashboard. The VBA warned with a
     MsgBox when there was no Dashboard, so the same guard is kept here.
     Prefers the Recon_Anchor defined name and falls back to A75, which is
     where the block sits in the current template. */
  function goRecon() {
    return SAZ.excel.run(function (context) {
      return SAZ.excel.exists(context, "Dashboard").then(function (found) {
        if (!found) {
          return SAZ.dialog.alert(
            "Reconciliation",
            "This workbook has no Dashboard sheet, so there is nothing to reconcile. Run New Project first."
          ).then(function () { return false; });
        }
        var sh = context.workbook.worksheets.getItem("Dashboard");
        sh.activate();
        var anchor = context.workbook.names.getItemOrNullObject("Recon_Anchor");
        anchor.load("isNullObject");
        return context.sync().then(function () {
          var target = anchor.isNullObject ? sh.getRange("A75") : anchor.getRange();
          target.select();
          return context.sync().then(function () { return true; });
        });
      });
    });
  }

  return { goOrCreate: goOrCreate, goRecon: goRecon };
})();

/* ----------------------------------------------------------
   Ribbon entry points.
   Each name below appears in the manifest as a FunctionName.
   Office hands in an event object and EVERY path must call
   event.completed() or that ribbon button locks until restart.
   ---------------------------------------------------------- */

function sazGo(sheetName, event) {
  SAZ.navigate.goOrCreate(sheetName)
    .catch(function () { /* already logged by core/excel.js */ })
    .then(function () { event.completed(); });
}

function goInstructions(e)   { sazGo("Instructions", e); }
function goDashboard(e)      { sazGo("Dashboard", e); }
function goLoadCalc(e)       { sazGo("Load Calc", e); }
function goUtility(e)        { sazGo("Utility Demand", e); }
function goLogging(e)        { sazGo("Logging Demand", e); }
function goGenerator(e)      { sazGo("Generator", e); }
function goUPS(e)            { sazGo("UPS", e); }
function goTransformers(e)   { sazGo("Transformers", e); }
function goShortCircuit(e)   { sazGo("Short Circuit", e); }
function goVoltageDrop(e)    { sazGo("Voltage Drop", e); }
function goArcFlash(e)       { sazGo("Arc Flash", e); }
function goFeederSchedule(e) { sazGo("Feeder Schedule", e); }
function goReport(e)         { sazGo("Report", e); }

function goRecon(e) {
  SAZ.navigate.goRecon()
    .catch(function () {})
    .then(function () { e.completed(); });
}

/* Office.actions.associate is required for the function file to bind these
   names. Without it the buttons appear but do nothing. */
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("goInstructions",   goInstructions);
  Office.actions.associate("goDashboard",      goDashboard);
  Office.actions.associate("goLoadCalc",       goLoadCalc);
  Office.actions.associate("goUtility",        goUtility);
  Office.actions.associate("goLogging",        goLogging);
  Office.actions.associate("goGenerator",      goGenerator);
  Office.actions.associate("goUPS",            goUPS);
  Office.actions.associate("goTransformers",   goTransformers);
  Office.actions.associate("goShortCircuit",   goShortCircuit);
  Office.actions.associate("goVoltageDrop",    goVoltageDrop);
  Office.actions.associate("goArcFlash",       goArcFlash);
  Office.actions.associate("goFeederSchedule", goFeederSchedule);
  Office.actions.associate("goReport",         goReport);
  Office.actions.associate("goRecon",          goRecon);
}
