/* Home page: Start here path, task tiles (one open at a time), search by title, Advanced and Reference. */
(function () {
  "use strict";
  var LS = window.EbobLessons, D = window.EbobDisplay;
  var PKEY = "ebob-guide-progress";
  function $(id) { return document.getElementById(id); }
  function progress() { try { return JSON.parse(localStorage.getItem(PKEY) || "{}") || {}; } catch (e) { return {}; } }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function href(l) { return l.page === "installer" ? "installer.html?lesson=" + l.id : l.kind === "doc" ? "doc.html?lesson=" + l.id : l.scenario ? "emulator.html?tutorial=" + l.scenario + "&fresh=1" : "emulator.html?tutorial=free-mode&lesson=" + l.id; }
  var I = {
    download: '<path d="M12 4v11M7 10.5l5 5 5-5"/><path d="M5 19.5h14"/>',
    silo: '<path d="M7 9a5 3 0 0 1 10 0v11H7z"/><path d="M7 13h10M7 17h10"/>',
    bell: '<path d="M6 16V11a6 6 0 0 1 12 0v5l1.5 2h-15zM10 20.5a2 2 0 0 0 4 0"/>',
    gauge: '<path d="M4 17a8 8 0 1 1 16 0"/><path d="M12 17l4-5"/><circle cx="12" cy="17" r="1.3"/>',
    chart: '<path d="M4 20V4M4 20h16"/><rect x="7" y="11" width="3" height="6" rx=".5"/><rect x="12" y="7" width="3" height="10" rx=".5"/><rect x="17" y="13" width="3" height="4" rx=".5"/>',
    users: '<circle cx="9" cy="8" r="3.2"/><path d="M3 19c0-3.3 2.7-5.5 6-5.5s6 2.2 6 5.5"/><path d="M16 5a3 3 0 0 1 0 6M18 13.7c1.9.7 3 2.7 3 5.3"/>',
    wrench: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.2L4 16.8V20h3.2l5.3-5.3a4 4 0 0 0 5.2-5.4l-2.6 2.6-2.4-.6-.6-2.4z"/>',
    check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>', chevR: '<path d="M9 6l6 6-6 6"/>', chevD: '<path d="M6 9l6 6 6-6"/>', play: '<path d="M8 5.5v13l10-6.5z"/>'
  };
  function ico(n, cls) { return '<svg class="eg-ico ' + (cls || "") + '" viewBox="0 0 24 24" aria-hidden="true">' + I[n] + "</svg>"; }

  /* the five essentials, in order */
  var START = [
    { id: "install", label: "Install Binventory", line: "On the main PC (Server/Host)." },
    { id: "tour", label: "First launch", line: "A quick look around the screen." },
    { id: "sensor-network", label: "Add a sensor network", line: "Tell Binventory how your sensors connect." },
    { id: "add-vessel", label: "Add a vessel", line: "Name, shape, contents and sensor." },
    { id: "level-alarms", label: "Set up an alert", line: "Get a warning when a level is high or low." }
  ];
  /* task tiles: every lesson except the Advanced ones */
  var TASKS = [
    { id: "install", title: "Install and upgrade", line: "Put Binventory on a PC, or upgrade it.", icon: "download", ids: ["plan", "install", "network-install", "install-import", "upgrade"] },
    { id: "site", title: "Set up my site", line: "Sensors, vessels and site settings.", icon: "silo", ids: ["tour", "system-setup", "site-setup", "sensor-network", "add-vessel", "more-vessels", "measure"] },
    { id: "alerts", title: "Alerts and email", line: "Get an email when a level is high or low.", icon: "bell", ids: ["level-alarms", "email-setup", "contacts", "email-alerts"] },
    { id: "readings", title: "Readings and schedules", line: "Measure vessels on a schedule.", icon: "gauge", ids: ["vessel-details", "groups", "schedules", "temporary-group"] },
    { id: "reports", title: "Reports", line: "Print or email inventory reports.", icon: "chart", ids: ["inventory-report", "history-report", "site-status-report"] },
    { id: "users", title: "Users and backups", line: "Logins and backing up your setup.", icon: "users", ids: ["users", "export"] },
    { id: "trouble", title: "Fix a problem", line: "Step-by-step fixes for common problems.", icon: "wrench", wide: true, ids: ["ts-comm-error", "ts-readings-wrong", "ts-no-schedule", "ts-no-email", "ts-pending-unknown", "ts-db-read-only", "ts-installer"] }
  ];
  var ADV = ["install-site", "install-external-sql"];
  var STATUS = [
    ["Unknown", "The sensor has not measured since the PC was restarted or the Binventory Engine was restarted."],
    ["Pending", "A measurement was requested and its result is not known yet."],
    ["Descending", "A SmartBob is measuring and dropping its bob."],
    ["Retracting", "A SmartBob is measuring and raising its bob."],
    ["Ready", "The last measurement worked. The sensor is ready."],
    ["Bob In Override", "The SmartBob ignored the request because its override input is active or open circuit."],
    ["Communications Error", "Binventory did not get the right reply from the sensor. Usually RS-485 wiring."],
    ["Bob Stuck Top", "The bob did not drop. The vessel may be very full, or the SmartBob has a mechanical fault."],
    ["Bob Stuck Bottom", "The bob did not fully retract. It may be buried in product, or the SmartBob has a mechanical fault."],
    ["Database Read Only", "Data shows but measurements cannot start. The Engine service may be stopped, or Site Setup has the wrong host IP."],
    ["Error", "One or more errors. Hover over the message or open Vessel Details for more."]
  ];
  var byId = function (id) { return LS.byId(id); };

  function row(l, p) {
    var done = !!p[l.id];
    return '<li><a class="hm-row' + (done ? " is-done" : "") + '" href="' + href(l) + '" data-lesson="' + l.id + '">' +
      '<span class="hm-mark" aria-hidden="true">' + (done ? ico("check") : "") + "</span>" +
      '<span class="hm-row-text"><span class="hm-row-title">' + esc(l.title) + (done ? '<span class="eg-sr-only"> (done)</span>' : "") + "</span>" +
      '<span class="hm-row-line">' + esc(l.desc) + "</span></span>" +
      '<span class="hm-row-meta">' + l.stepCount + " steps</span>" +
      '<span class="hm-row-go" aria-hidden="true">Start</span></a></li>';
  }

  /* ---------- Start here ---------- */
  function renderStart(p) {
    var done = 0, next = null, html = "";
    START.forEach(function (s, i) {
      var l = byId(s.id), isDone = !!p[s.id];
      if (isDone) done++; else if (!next) next = { s: s, l: l, n: i + 1 };
      var isNext = next && next.s === s;
      html += '<li><a class="hm-step' + (isDone ? " is-done" : "") + (isNext ? " is-next" : "") + '" href="' + href(l) + '" data-lesson="' + s.id + '">' +
        '<span class="hm-step-n" aria-hidden="true">' + (isDone ? ico("check") : i + 1) + "</span>" +
        '<span class="hm-step-text"><span class="hm-step-title"><span class="eg-sr-only">Step ' + (i + 1) + ": </span>" + esc(s.label) + (isDone ? '<span class="eg-sr-only"> (done)</span>' : "") + "</span>" +
        '<span class="hm-step-line">' + esc(s.line) + "</span></span>" +
        (isDone ? '<span class="hm-tag hm-tag--done" aria-hidden="true">' + ico("check") + "Done</span>" : isNext ? '<span class="hm-tag" aria-hidden="true">Next</span>' : "") +
        ico("chevR", "hm-step-chev") + "</a></li>";
    });
    $("startSteps").innerHTML = html;
    $("startProgress").textContent = done + " of " + START.length + " done";
    var c = $("hmContinue");
    if (next) {
      c.hidden = false; c.href = href(next.l);
      c.innerHTML = ico("play") + (done === 0 ? "Start step 1" : "Continue: " + esc(next.s.label));
    } else {
      c.hidden = true;
      $("startProgress").textContent = "All five done. Nice work.";
    }
  }

  /* ---------- task tiles ---------- */
  var openTask = null;
  function renderTiles(p) {
    var html = "";
    TASKS.forEach(function (t, i) {
      var ls = t.ids.map(byId).filter(Boolean), isOpen = openTask === t.id;
      var doneN = ls.filter(function (l) { return p[l.id]; }).length;
      html += '<h3 class="hm-tile-h' + (t.wide ? " hm-tile-h--wide" : "") + '" data-i="' + i + '" data-task-id="' + t.id + '">' +
        '<button type="button" class="hm-tile' + (isOpen ? " is-open" : "") + '" id="task-' + t.id + '" data-task="' + t.id + '" aria-expanded="' + isOpen + '" aria-controls="panel-' + t.id + '">' +
        '<span class="hm-tile-icon">' + ico(t.icon) + "</span>" +
        '<span class="hm-tile-text"><span class="hm-tile-title">' + esc(t.title) + '</span><span class="hm-tile-line">' + esc(t.line) + "</span>" +
        '<span class="hm-tile-count">' + ls.length + " lessons" + (doneN ? ", " + doneN + " done" : "") + "</span></span>" +
        ico("chevD", "hm-chev") + "</button></h3>" +
        '<div class="hm-panel" id="panel-' + t.id + '" role="region" aria-labelledby="task-' + t.id + '"' + (isOpen ? "" : " hidden") + ">" +
        '<ul class="hm-rows">' + ls.map(function (l) { return row(l, p); }).join("") + "</ul>" +
        (t.id === "trouble" ? '<details class="hm-status"><summary>What each vessel status means</summary><dl>' +
          STATUS.map(function (s) { return "<dt>" + s[0] + "</dt><dd>" + s[1] + "</dd>"; }).join("") + "</dl></details>" : "") +
        "</div>";
    });
    $("tiles").innerHTML = html;
    placePanel();
  }
  /* the open panel shows full width right under the row of tiles that holds its tile (DOM order stays tile, panel) */
  function placePanel() {
    var heads = [].slice.call($("tiles").querySelectorAll(".hm-tile-h"));
    heads.forEach(function (h, i) { h.style.order = String(i * 2); if (h.nextElementSibling) h.nextElementSibling.style.order = String(i * 2 + 1); });
    if (!openTask) return;
    var at = TASKS.findIndex(function (t) { return t.id === openTask; });
    var panel = $("panel-" + openTask); if (!panel || at < 0) return;
    panel.hidden = true;
    var top = heads[at].offsetTop, end = at;
    for (var k = at + 1; k < heads.length; k++) { if (Math.abs(heads[k].offsetTop - top) < 4) end = k; else break; }
    panel.style.order = String(end * 2 + 1);
    panel.hidden = false;
  }
  function setTask(id, focus) {
    openTask = openTask === id ? null : id;
    renderTiles(progress());
    if (focus) { var b = $("task-" + id); if (b) b.focus(); }
  }
  $("tiles").addEventListener("click", function (e) {
    var b = e.target.closest(".hm-tile"); if (!b) return;
    setTask(b.getAttribute("data-task"), true);
  });
  function openFix() {
    if (search.value) { search.value = ""; renderSearch(); }
    openTask = null; setTask("trouble", false);
    var b = $("task-trouble"); b.scrollIntoView({ block: "start", behavior: D.reduceMotion() ? "auto" : "smooth" }); b.focus({ preventScroll: true });
  }
  $("topFix").addEventListener("click", function (e) { e.preventDefault(); openFix(); try { history.replaceState(null, "", "#task-trouble"); } catch (x) {} });

  /* ---------- Advanced ---------- */
  function renderAdv(p) { $("advBody").innerHTML = '<ul class="hm-rows">' + ADV.map(byId).filter(Boolean).map(function (l) { return row(l, p); }).join("") + "</ul>"; }
  $("advBtn").addEventListener("click", function () {
    var b = $("advBtn"), open = b.getAttribute("aria-expanded") !== "true";
    b.setAttribute("aria-expanded", String(open)); $("advBody").hidden = !open; b.classList.toggle("is-open", open);
  });

  /* ---------- search by title ---------- */
  var search = $("lessonSearch");
  function renderSearch() {
    var q = search.value.trim().toLowerCase(), p = progress();
    if (!q) { $("results").hidden = true; $("results").innerHTML = ""; $("tiles").hidden = false; $("searchCount").textContent = ""; return; }
    var hits = LS.lessons.filter(function (l) { return l.title.toLowerCase().indexOf(q) >= 0; });
    $("tiles").hidden = true; $("results").hidden = false;
    $("results").innerHTML = hits.length ? '<ul class="hm-rows">' + hits.map(function (l) { return row(l, p); }).join("") + "</ul>"
      : '<p class="hm-empty">No lesson titles match "' + esc(search.value.trim()) + '". Try vessel, alarm, email, report or install.</p>';
    $("searchCount").textContent = hits.length + (hits.length === 1 ? " lesson matches" : " lessons match");
  }
  search.addEventListener("input", renderSearch);
  document.addEventListener("keydown", function (e) {
    if (e.key === "/" && document.activeElement !== search && !/input|textarea|select/i.test(document.activeElement.tagName)) { e.preventDefault(); search.focus(); }
    if (e.key === "Escape" && document.activeElement === search && search.value) { search.value = ""; renderSearch(); }
  });

  /* ---------- display settings: one button, one panel ---------- */
  D.attachPanel($("displayBtn"), {
    title: "Display settings",
    place: function (panel, btn) {
      var r = btn.getBoundingClientRect();
      panel.style.position = "absolute";
      panel.style.top = (r.bottom + window.scrollY + 8) + "px";
      panel.style.right = Math.max(12, document.documentElement.clientWidth - r.right) + "px";
    }
  });

  $("resetProgress").addEventListener("click", function () {
    if (!confirm("Clear the done marks on every lesson?")) return;
    try { localStorage.removeItem(PKEY); } catch (e) {}
    render();
  });

  function render() {
    var p = progress();
    renderStart(p); renderTiles(p); renderAdv(p); renderSearch();
  }
  window.addEventListener("storage", function (e) { if (e.key === PKEY) render(); });
  window.addEventListener("pageshow", render);
  var rz; window.addEventListener("resize", function () { clearTimeout(rz); rz = setTimeout(placePanel, 80); });
  D.onChange(function () { setTimeout(placePanel, 0); });
  render();
  /* old links: #sec-trouble and #task-trouble open Fix a problem; #sec-setup goes to Start here */
  var h = location.hash;
  if (/^#(sec-trouble|task-trouble|trouble)$/.test(h)) setTimeout(openFix, 0);
  else if (/^#(sec-setup|setup|start)$/.test(h)) setTimeout(function () { $("start").focus(); }, 0);
  else if (/^#(sec-advanced|advanced)$/.test(h)) setTimeout(function () { $("advBtn").click(); $("advBtn").focus(); }, 0);
  window.EbobHome = { TASKS: TASKS, START: START, ADV: ADV, href: href };
})();
