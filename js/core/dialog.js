/* ==========================================================
   SAZAN PowerCalc  -  js/core/dialog.js
   Replaces every VBA MsgBox and InputBox.

   Ribbon commands run in a headless runtime with no document
   object, so window.confirm and window.prompt are unavailable.
   All prompting goes through Office dialogs, which also keeps
   the same code working in Excel on the web and on Mac.
   ========================================================== */

var SAZ = SAZ || {};

SAZ.dialog = (function () {
  "use strict";

  /* Must be same-origin with the add-in or the AppDomains list has to grow. */
  var BASE = "https://csmckinnon.github.io/sazan-paneltree-addin/dialog.html";

  /* open — renders dialog.html against a spec and resolves its reply.
     The spec rides in the URL fragment so dialog.html needs no round trip. */
  function open(spec, width, height) {
    return new Promise(function (resolve) {
      var url = BASE + "#" + encodeURIComponent(JSON.stringify(spec));
      Office.context.ui.displayDialogAsync(
        url,
        { width: width || 30, height: height || 22, displayInIframe: true },
        function (res) {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("[SAZAN] dialog failed", res.error && res.error.message);
            resolve(null);
            return;
          }
          var dlg = res.value;
          dlg.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            var payload = null;
            try { payload = JSON.parse(arg.message); } catch (e) { payload = null; }
            dlg.close();
            resolve(payload);
          });
          /* Fires when the user closes with the X. Treat as cancel, never hang. */
          dlg.addEventHandler(Office.EventType.DialogEventReceived, function () {
            resolve(null);
          });
        }
      );
    });
  }

  /* confirm — replaces MsgBox vbYesNo. Resolves true or false. */
  function confirm(title, message, okLabel) {
    return open({ kind: "confirm", title: title, message: message, ok: okLabel || "Yes" })
      .then(function (r) { return !!(r && r.ok); });
  }

  /* alert — replaces MsgBox vbOKOnly. */
  function alert(title, message) {
    return open({ kind: "alert", title: title, message: message }, 30, 18)
      .then(function () { return true; });
  }

  /* prompt — replaces InputBox. Resolves the string, or null if cancelled. */
  function prompt(title, message, defaultValue) {
    return open({ kind: "prompt", title: title, message: message, value: defaultValue || "" })
      .then(function (r) { return (r && r.ok) ? r.value : null; });
  }

  /* pick — replaces the numbered InputBox menus such as PickEquipTable and
     PickVoltage. options is an array of { value, label }. */
  function pick(title, message, options) {
    return open({ kind: "pick", title: title, message: message, options: options }, 30, 26)
      .then(function (r) { return (r && r.ok) ? r.value : null; });
  }

  return { open: open, confirm: confirm, alert: alert, prompt: prompt, pick: pick };
})();
