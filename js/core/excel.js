/* ==========================================================
   SAZAN PowerCalc  -  js/core/excel.js
   Thin wrappers over Excel.run. Every command goes through
   SAZ.excel.run(). No command file calls Excel.run directly,
   so suspension and error handling live in exactly one place.
   ========================================================== */

var SAZ = SAZ || {};

SAZ.excel = (function () {
  "use strict";

  /* Sheets the user must never create, delete or treat as a module. */
  var INTERNAL = ["NEC_Data", "Energy_Code_Data", "_TplMeta", "Claude Log"];

  function isInternal(name) {
    return INTERNAL.indexOf(name) !== -1 || /^TPL_/i.test(name);
  }

  /* run(fn)
     Single entry point for every command.
     Suspends recalculation and screen updating for the whole batch, which
     is the Office.js equivalent of Application.ScreenUpdating = False, and
     funnels failures into one handler instead of 40 try/catch blocks. */
  function run(fn) {
    return Excel.run(function (context) {
      context.application.suspendApiCalculationUntilNextSync();
      context.application.suspendScreenUpdatingUntilNextSync();
      return fn(context);
    }).catch(function (e) {
      SAZ.excel.lastError = e;
      console.error("[SAZAN]", e && e.message, e && e.debugInfo);
      throw e;
    });
  }

  /* exists — getItemOrNullObject avoids throwing on a missing sheet. */
  function exists(context, name) {
    var sh = context.workbook.worksheets.getItemOrNullObject(name);
    sh.load("isNullObject");
    return context.sync().then(function () { return !sh.isNullObject; });
  }

  function sheetNames(context) {
    var shs = context.workbook.worksheets;
    shs.load("items/name");
    return context.sync().then(function () {
      return shs.items.map(function (s) { return s.name; });
    });
  }

  function activate(context, name) {
    context.workbook.worksheets.getItem(name).activate();
    return context.sync().then(function () { return true; });
  }

  /* hide — insertWorksheetsFromBase64 cannot insert a sheet already hidden,
     so Wave 3 lands the data tabs visible and calls this in the same sync
     to keep the flicker to a single frame. */
  function hide(context, names) {
    names.forEach(function (n) {
      var sh = context.workbook.worksheets.getItemOrNullObject(n);
      sh.visibility = Excel.SheetVisibility.hidden;
    });
    return context.sync();
  }

  /* namedRange — returns a Range or null. Replaces wb.Names("x").RefersToRange. */
  function namedRange(context, name) {
    var nm = context.workbook.names.getItemOrNullObject(name);
    nm.load("isNullObject");
    return context.sync().then(function () {
      return nm.isNullObject ? null : nm.getRange();
    });
  }

  /* readBlock — one bulk read of a whole region.
     The VBA habit of touching one cell at a time is what makes ported code
     crawl. Read the block, work on the array, write the block back. */
  function readBlock(context, sheet, address) {
    var rng = context.workbook.worksheets.getItem(sheet).getRange(address);
    rng.load("values,text,rowCount,columnCount");
    return context.sync().then(function () { return rng; });
  }

  /* selectedRow — replaces ActiveCell.Row. */
  function selectedRow(context) {
    var sel = context.workbook.getSelectedRange();
    sel.load("rowIndex,columnIndex,worksheet/name");
    return context.sync().then(function () {
      return { row: sel.rowIndex + 1, col: sel.columnIndex + 1, sheet: sel.worksheet.name };
    });
  }

  return {
    run: run,
    exists: exists,
    sheetNames: sheetNames,
    activate: activate,
    hide: hide,
    namedRange: namedRange,
    readBlock: readBlock,
    selectedRow: selectedRow,
    isInternal: isInternal,
    INTERNAL: INTERNAL,
    lastError: null
  };
})();
