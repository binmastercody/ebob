/* Quick menu on the Binventory practice screen (emulator.html with no lesson running):
 *  - Welcome card (first visit, ?welcome=1) over the live Binventory screen, like the BinDisc, CNCR and 3D guides.
 *  - Floating "What would you like to do?" box with five choices. Each choice opens its lessons in the same box.
 *  - Copy-link icon on every choice and lesson so a tech can send a direct link.
 * Deep links: emulator.html?do=<choice>  opens the box on that choice (no welcome card).
 *             emulator.html?menu=1       opens the box (no welcome card). */
(function () {
    "use strict";
    var LS = window.EbobLessons;
    var sp = new URLSearchParams(location.search);
    if (!LS || sp.get("lesson") || String(sp.get("tutorial") || "").toLowerCase() !== "free-mode") return;

    var CHOICES = [
        { id: "install", title: "Install, upgrade or back up", blurb: "Put Binventory on a PC, upgrade it, or save a backup.",
          ids: ["plan", "install", "network-install", "install-import", "upgrade", "export", "install-site", "install-external-sql"] },
        { id: "site", title: "Set up my site", blurb: "A quick tour, sensor networks, vessels and users.",
          ids: ["tour", "system-setup", "site-setup", "sensor-network", "add-vessel", "more-vessels", "measure", "users"] },
        { id: "alerts", title: "Set up alarms and email", blurb: "Get an email when a level is high or low.",
          ids: ["level-alarms", "email-setup", "contacts", "email-alerts"] },
        { id: "readings", title: "Readings and reports", blurb: "Schedules, groups, history and inventory reports.",
          ids: ["vessel-details", "groups", "schedules", "temporary-group", "inventory-report", "history-report", "site-status-report"] },
        { id: "trouble", title: "Fix a problem", blurb: "Step-by-step fixes for common problems.",
          ids: ["ts-comm-error", "ts-readings-wrong", "ts-no-schedule", "ts-no-email", "ts-pending-unknown", "ts-db-read-only", "ts-installer"] }
    ];
    var PKEY = "ebob-guide-progress", MINI_KEY = "ebob-qm-mini";
    var SVG = function (p) { return '<svg viewBox="0 0 24 24" aria-hidden="true">' + p + "</svg>"; };
    var ICO = {
        link: SVG('<path d="M10 14a4.5 4.5 0 0 0 6.4 0l3-3a4.5 4.5 0 0 0-6.4-6.4l-1.2 1.2"/><path d="M14 10a4.5 4.5 0 0 0-6.4 0l-3 3a4.5 4.5 0 0 0 6.4 6.4l1.2-1.2"/>'),
        check: SVG('<path d="M5 12.5l4.5 4.5L19 7.5"/>'),
        chev: '<svg class="qm-chev" viewBox="0 0 24 24" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>',
        back: SVG('<path d="M15 6l-6 6 6 6"/>'),
        min: SVG('<path d="M6 12h12"/>'),
        max: SVG('<path d="M6 12h12M12 6v12"/>')
    };
    function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
    function progress() { try { return JSON.parse(localStorage.getItem(PKEY) || "{}") || {}; } catch (e) { return {}; } }
    function lessonHref(l) {
        return l.page === "installer" ? "installer.html?lesson=" + l.id : l.kind === "doc" ? "doc.html?lesson=" + l.id : l.scenario ? "emulator.html?tutorial=" + l.scenario + "&fresh=1" : "emulator.html?tutorial=free-mode&lesson=" + l.id;
    }
    function abs(href) { try { return new URL(href, location.href).href; } catch (e) { return href; } }
    function choiceById(id) { for (var i = 0; i < CHOICES.length; i++) if (CHOICES[i].id === id) return CHOICES[i]; return null; }

    /* clean the address bar: the welcome flag is one-shot, do/menu stay out of copied URLs */
    var showWelcome = sp.get("welcome") === "1";
    var startChoice = choiceById(String(sp.get("do") || "").toLowerCase());
    ["welcome", "menu", "do"].forEach(function (k) { sp.delete(k); });
    try { history.replaceState(null, "", location.pathname + "?" + sp.toString() + location.hash); } catch (e) {}
    document.documentElement.classList.add("eg-quickmenu");

    /* ---------------- floating box ---------------- */
    var panel = document.createElement("section");
    panel.className = "eg-ui qm-panel";
    panel.setAttribute("aria-labelledby", "qmTitle");
    panel.innerHTML =
        '<div class="qm-hdr">' +
        '<button type="button" class="qm-hbtn qm-back" aria-label="Back to all choices" title="Back" hidden>' + ICO.back + "</button>" +
        '<img class="qm-hdr-ico" src="assets/eBob-V5-Application.ico" alt="">' +
        '<h2 class="qm-hdr-title" id="qmTitle">What would you like to do?</h2>' +
        '<button type="button" class="qm-hbtn qm-min" aria-expanded="true" aria-controls="qmBody" title="Minimize">' + ICO.min + "</button>" +
        "</div>" +
        '<div class="qm-body" id="qmBody"></div>' +
        '<div class="qm-foot"><span class="qm-status" role="status" aria-live="polite"></span><a href="lessons.html">Full lesson list</a></div>';
    document.body.appendChild(panel);
    var body = panel.querySelector(".qm-body"), titleEl = panel.querySelector(".qm-hdr-title"),
        backBtn = panel.querySelector(".qm-back"), minBtn = panel.querySelector(".qm-min"), status = panel.querySelector(".qm-status");
    var current = null;

    function linkBtn(url, label) {
        return '<button type="button" class="qm-link" data-copy="' + esc(url) + '" aria-label="Copy link: ' + esc(label) + '" title="Copy link">' + ICO.link + "</button>";
    }
    function renderHome() {
        current = null;
        titleEl.textContent = "What would you like to do?";
        backBtn.hidden = true;
        body.innerHTML = CHOICES.map(function (c, i) {
            return '<div class="qm-row"><button type="button" class="qm-item" data-choice="' + c.id + '">' +
                '<span class="qm-num" aria-hidden="true">' + (i + 1) + "</span>" +
                '<span class="qm-item-text"><span class="qm-item-title">' + esc(c.title) + '</span><span class="qm-item-blurb">' + esc(c.blurb) + "</span></span>" +
                ICO.chev + "</button>" + linkBtn(abs("emulator.html?do=" + c.id), c.title) + "</div>";
        }).join("");
    }
    function renderChoice(c) {
        current = c;
        var p = progress();
        titleEl.textContent = c.title;
        backBtn.hidden = false;
        body.innerHTML = '<p class="qm-sub">Pick a lesson. It runs on this practice copy of Binventory.</p>' +
            c.ids.map(function (id) { return LS.byId(id); }).filter(Boolean).map(function (l) {
                var done = !!p[l.id];
                return '<div class="qm-row"><a class="qm-item" href="' + esc(lessonHref(l)) + '">' +
                    (done ? '<span class="qm-num qm-done" aria-hidden="true">' + ICO.check + "</span>" : "") +
                    '<span class="qm-item-text"><span class="qm-item-title">' + esc(l.title) + (done ? '<span class="eg-sr-only"> (done)</span>' : "") + "</span>" +
                    (l.desc ? '<span class="qm-item-blurb">' + esc(l.desc) + "</span>" : "") + "</span>" +
                    ICO.chev + "</a>" + linkBtn(abs(lessonHref(l)), l.title) + "</div>";
            }).join("");
    }
    function show(c, focus) {
        if (c) renderChoice(c); else renderHome();
        body.scrollTop = 0;
        setMini(false);
        if (focus) { var f = body.querySelector(".qm-item"); if (f) f.focus({ preventScroll: true }); }
    }

    function setMini(on) {
        panel.classList.toggle("is-mini", on);
        minBtn.innerHTML = on ? ICO.max : ICO.min;
        minBtn.title = on ? "Show choices" : "Minimize";
        minBtn.setAttribute("aria-label", on ? "Show choices" : "Minimize");
        minBtn.setAttribute("aria-expanded", String(!on));
        try { sessionStorage.setItem(MINI_KEY, on ? "1" : "0"); } catch (e) {}
        dock();
    }
    minBtn.addEventListener("click", function () { setMini(!panel.classList.contains("is-mini")); });
    backBtn.addEventListener("click", function () { show(null, true); });

    var statusTimer = 0;
    function copyText(text) {
        function legacy() {
            return new Promise(function (ok, fail) {
                var ta = document.createElement("textarea");
                ta.value = text; ta.setAttribute("readonly", ""); ta.style.position = "fixed"; ta.style.opacity = "0";
                document.body.appendChild(ta); ta.select();
                var r = false; try { r = document.execCommand("copy"); } catch (e) {}
                ta.remove(); r ? ok() : fail();
            });
        }
        if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text).catch(legacy);
        return legacy();
    }
    body.addEventListener("click", function (e) {
        var cp = e.target.closest(".qm-link");
        if (cp) {
            e.preventDefault();
            copyText(cp.getAttribute("data-copy")).then(function () {
                cp.classList.add("is-copied"); cp.innerHTML = ICO.check;
                status.textContent = "Link copied";
                clearTimeout(statusTimer);
                statusTimer = setTimeout(function () { status.textContent = ""; }, 2400);
                setTimeout(function () { cp.classList.remove("is-copied"); cp.innerHTML = ICO.link; }, 1600);
            }, function () { window.prompt("Copy this link:", cp.getAttribute("data-copy")); });
            return;
        }
        var ch = e.target.closest("[data-choice]");
        if (ch) { show(choiceById(ch.getAttribute("data-choice")), true); }
    });

    /* wide screens: shift the Binventory window left so the box does not cover it */
    function dock() {
        var root = document.documentElement, w = panel.getBoundingClientRect().width + 24;
        var room = window.innerWidth - w - 32 >= 900 && !panel.classList.contains("is-mini");
        root.style.setProperty("--qm-dock-w", (room ? w : 0) + "px");
        root.classList.toggle("qm-docked", room);
        root.classList.toggle("qm-docked-welcome", room);
        if (window.__ebobUpdateAppColumnLayout) window.__ebobUpdateAppColumnLayout();
        requestAnimationFrame(besideWindow);
    }
    /* sit the box just right of the Binventory window, top-aligned with it, when there is room */
    function besideWindow() {
        var sh = document.getElementById("appShell"), r = sh && sh.getBoundingClientRect();
        var pw = panel.getBoundingClientRect().width;
        panel.style.top = ""; var top = parseFloat(getComputedStyle(panel).top) || 0;
        if (r && r.width && r.right + 20 + pw + 16 <= window.innerWidth) {
            panel.style.left = Math.round(r.right + 20) + "px"; panel.style.right = "auto";
            if (r.top > top) panel.style.top = Math.round(r.top) + "px";
        } else { panel.style.left = ""; panel.style.right = ""; panel.style.top = ""; }
    }
    setInterval(besideWindow, 400);
    window.addEventListener("resize", function () { requestAnimationFrame(dock); });
    if (window.EbobDisplay && window.EbobDisplay.onChange) window.EbobDisplay.onChange(function () { requestAnimationFrame(dock); });

    show(startChoice, false);
    var mini = false; try { mini = sessionStorage.getItem(MINI_KEY) === "1"; } catch (e) {}
    if (mini && !showWelcome && !startChoice) setMini(true); else dock();

    /* ---------------- welcome card ---------------- */
    if (!showWelcome) return;
    var wel = document.createElement("div");
    wel.className = "eg-ui qm-welcome";
    wel.setAttribute("role", "dialog"); wel.setAttribute("aria-modal", "false"); wel.setAttribute("aria-labelledby", "qmWelTitle");
    wel.innerHTML =
        '<div class="qm-card">' +
        '<button type="button" class="qm-trouble" aria-label="Troubleshooting: fix a problem"><span class="qm-trouble-ico" aria-hidden="true"></span><span class="qm-trouble-label">Troubleshooting</span></button>' +
        '<div class="qm-shot"><img src="app/preview.png" alt="The Binventory Workstation main screen with vessel panels"></div>' +
        '<div class="qm-copy">' +
        '<h1 class="qm-title" id="qmWelTitle"><span>Welcome to the</span><span class="qm-accent">Binventory Workstation Guide</span></h1>' +
        '<p class="qm-tag">Practice on a copy of Binventory Workstation 1.0. Press <b>Get Started</b>, then pick what you want to do from the box on the right.</p>' +
        '<button type="button" class="qm-start">Get Started</button>' +
        '<p class="qm-ver">Binventory Workstation 1.0 Training</p>' +
        "</div></div>";
    document.body.appendChild(wel);
    document.documentElement.classList.add("qm-welcome-open");
    function fitWelcome() { var vw = window.visualViewport ? window.visualViewport.width : window.innerWidth; document.documentElement.style.setProperty("--qm-vw", Math.round(Math.min(vw, window.innerWidth)) + "px"); }
    fitWelcome(); window.addEventListener("resize", fitWelcome);
    var startBtn = wel.querySelector(".qm-start");
    setTimeout(function () { try { startBtn.focus({ preventScroll: true }); } catch (e) {} }, 60);

    function closeWelcome(then) {
        if (!wel.parentNode) return;
        wel.classList.add("is-leaving");
        document.documentElement.classList.remove("qm-welcome-open");
        document.removeEventListener("keydown", onKey);
        setTimeout(function () { wel.remove(); }, 360);
        if (then) then();
        panel.classList.remove("qm-pulse"); void panel.offsetWidth; panel.classList.add("qm-pulse");
    }
    function onKey(e) { if (e.key === "Escape") closeWelcome(function () { show(current, true); }); }
    document.addEventListener("keydown", onKey);
    startBtn.addEventListener("click", function () { closeWelcome(function () { show(null, true); }); });
    wel.querySelector(".qm-trouble").addEventListener("click", function () { closeWelcome(function () { show(choiceById("trouble"), true); }); });
    /* picking from the box while the welcome card is up also closes it */
    panel.addEventListener("click", function (e) { if (e.target.closest("[data-choice], a.qm-item")) closeWelcome(); });
})();
