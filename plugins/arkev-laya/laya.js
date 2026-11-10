// ARKEV's ONLYOFFICE plugin: the live half of "Ask Laya to edit".
//
// Protocol 1 with Core (arkev-core internal/office, docs/office-editors.md):
//   options   editorConfig.plugins.options[<this guid>] = {"core": <Core URL>, "token": <plugin token>}
//             in the config Core signed. The token is for this person in this document only.
//   GET  <core>/v1/office/plugin/next          long poll (Core holds it up to ~25 s)
//        -> {"task": null} | {"task": {"id", "kind": "outline"}} | {"task": {"id", "kind": "edits", "edits": [...]}}
//   POST <core>/v1/office/plugin/tasks/{id}    {"outline": {...}} | {"result": {"applied", "skipped"}} | {"error": "..."}
//   401/403 ends the loop (token expired or the document was closed); other errors back off.
// Edits run in the editor through callCommand, so everyone in the document sees them at once.
(function (window) {
  "use strict";

  var plugin = window.Asc.plugin;
  var core = "", token = "", running = false, delay = 1000;
  var MAX_DELAY = 30000;

  function editorType() {
    var t = plugin.info && (plugin.info.editorType || plugin.info.editorSubType);
    return t === "cell" || t === "slide" ? t : "word";
  }

  function call(method, path, body) {
    var init = {method: method, headers: {"Authorization": "Bearer " + token}, cache: "no-store", credentials: "omit", mode: "cors"};
    if (body !== undefined) {
      init.headers["Content-Type"] = "application/json";
      init.body = JSON.stringify(body);
    }
    return fetch(core + path, init).then(function (r) {
      if (r.status === 401 || r.status === 403) { var e = new Error("stopped"); e.stop = true; throw e; }
      if (!r.ok) throw new Error("Core answered " + r.status);
      return r.status === 204 ? null : r.json();
    });
  }

  function run(fn, scope) {
    return new Promise(function (resolve, reject) {
      window.Asc.scope = scope || {};
      try {
        plugin.callCommand(fn, false, true, function (out) {
          try { resolve(typeof out === "string" ? JSON.parse(out) : out); } catch (e) { reject(e); }
        });
      } catch (e) { reject(e); }
    });
  }

  function handle(task) {
    var commands = window.ARKEV_COMMANDS[editorType()];
    var work = task.kind === "outline"
      ? run(commands.outline).then(function (outline) { return {outline: outline}; })
      : task.kind === "edits"
        ? run(commands.apply, {edits: task.edits || []}).then(function (result) { return {result: result}; })
        : Promise.resolve({error: "unknown task " + task.kind});
    return work.catch(function (e) { return {error: String(e && e.message || e).slice(0, 500)}; })
      .then(function (body) { return call("POST", "/v1/office/plugin/tasks/" + encodeURIComponent(task.id), body); });
  }

  function loop() {
    call("GET", "/v1/office/plugin/next").then(function (r) {
      delay = 1000;
      return r && r.task ? handle(r.task) : null;
    }).then(function () {
      setTimeout(loop, 50);
    }, function (e) {
      if (e && e.stop) { running = false; return; }
      setTimeout(loop, delay);
      delay = Math.min(delay * 2, MAX_DELAY);
    });
  }

  plugin.init = function () {
    var o = (plugin.info && plugin.info.options) || {};
    var allowed = window.ARKEV_LAYA_CORE || [];
    if (typeof o.core !== "string" || allowed.indexOf(o.core) < 0 || typeof o.token !== "string" || !o.token) return;
    core = o.core;
    token = o.token;
    if (!running) { running = true; loop(); }
  };
  plugin.button = function () {};
})(window);
