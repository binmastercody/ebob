/*
 * eBob Guide engine for emulator.html.
 *  - ?lesson=<id>      runs a guided lesson from app/lessons.js on the sample site (free mode)
 *  - ?tutorial=db-read-only | pending-unknown-status  dresses the built-in troubleshooting scenarios
 *    (their steps live in ebob-sim.js; pacing is patched there to use EbobGuidePace)
 *
 * Pacing rule: a step accepts your action only after its highlight has finished moving in ("settle"),
 * and a finished step holds a green check before the next one appears ("hold"). Clicks that arrive
 * too early are held back with a visible message instead of being lost.
 */
(function () {
    "use strict";
    var D = window.EbobDisplay, LS = window.EbobLessons, bootAt = Date.now();
    if (!D || !LS) return;
    var H = LS.helpers;
    /* the page hosting the lesson: the Binventory emulator by default, or the installer replica */
    var HOST = window.EbobGuideHost || {
        shellId: "appShell",
        minShellW: 880,
        refit: function () { if (window.__ebobUpdateAppColumnLayout) window.__ebobUpdateAppColumnLayout(); },
        ready: function () {
            var ov = document.getElementById("startupOverlay");
            if (ov && ov.classList.contains("startup-overlay--visible")) return false;
            var grid = document.getElementById("vesselGrid");
            if (!grid || !grid.children.length) return false;
            if (grid.querySelector(".vessel--initializing") && Date.now() - bootAt < 20000) return false;
            return true;
        }
    };
    var PKEY = "ebob-guide-progress";
    var sp = new URLSearchParams(location.search);
    var lessonId = sp.get("lesson");
    var scenario = String(sp.get("tutorial") || "").toLowerCase();

    function progress() { try { return JSON.parse(localStorage.getItem(PKEY) || "{}") || {}; } catch (e) { return {}; } }
    function markDone(id) { var p = progress(); p[id] = Date.now(); try { localStorage.setItem(PKEY, JSON.stringify(p)); } catch (e) {} }
    function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
    function stripTags(s) { var d = document.createElement("div"); d.innerHTML = s; return d.textContent; }
    function now() { return Date.now(); }

    var ICON_CHECK = '<svg viewBox="0 0 24 24" aria-hidden="true" class="eg-ico"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>';
    var ICON_BACK = '<svg viewBox="0 0 24 24" aria-hidden="true" class="eg-ico"><path d="M15 6l-6 6 6 6"/></svg>';
    var ICON_NEXT = '<svg viewBox="0 0 24 24" aria-hidden="true" class="eg-ico"><path d="M9 6l6 6-6 6"/></svg>';
    var ICON_GRID = '<svg viewBox="0 0 24 24" aria-hidden="true" class="eg-ico"><rect x="4" y="4" width="6.5" height="6.5" rx="1.5"/><rect x="13.5" y="4" width="6.5" height="6.5" rx="1.5"/><rect x="4" y="13.5" width="6.5" height="6.5" rx="1.5"/><rect x="13.5" y="13.5" width="6.5" height="6.5" rx="1.5"/></svg>';
    var HAND = '<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M27 6c-2.8 0-5 2.2-5 5v22.5l-3.6-3.7c-2-2-5.2-2-7.1 0-1.9 1.9-2 5-.2 7l11.6 13.4c2.9 3.4 7.2 5.3 11.6 5.3H40c7.7 0 14-6.3 14-14V30.5c0-2.6-2.1-4.7-4.7-4.7-1 0-1.9.3-2.6.8v-.6c0-2.6-2.1-4.7-4.7-4.7-1.1 0-2.1.4-2.9 1-.6-2-2.4-3.4-4.5-3.4-.9 0-1.8.3-2.6.7V11c0-2.8-2.2-5-5-5z" fill="#fff" stroke="#111827" stroke-width="3" stroke-linejoin="round"/></svg>';

    /* ---------- shared chrome: top bar ---------- */
    var lesson = lessonId ? LS.byId(lessonId) : null;
    var scenarioLesson = null;
    LS.lessons.forEach(function (l) { if (l.scenario && l.scenario === scenario) scenarioLesson = l; });
    var shownLesson = lesson || scenarioLesson;

    function buildTopBar() {
        var top = document.createElement("header");
        top.className = "eg-ui eg-topbar";
        var sec = shownLesson ? LS.sections.filter(function (s) { return s.id === shownLesson.section; })[0] : null;
        top.innerHTML =
            '<a class="eg-tb-btn eg-home" href="index.html">' + ICON_GRID + "<span>All lessons</span></a>" +
            '<div class="eg-top-title"><span class="eg-top-kicker">' + esc(sec ? sec.title : "Free practice") + "</span>" +
            '<span class="eg-top-name">' + esc(shownLesson ? shownLesson.title : "Binventory Workstation") + "</span></div>";
        document.body.appendChild(top);
        D.buildToolbar(top);
        var fab = document.getElementById("guideBackToMenu");
        if (fab) fab.classList.add("eg-hidden-fab");
    }
    buildTopBar();
    /* frmSplash and frmWait open centered on the Binventory window (as in Windows, CenterParent),
       so they line up with the window wherever the layout puts it */
    function syncWinCenter() {
        if (!document.documentElement.classList.contains("eg-emulator") || window.EbobGuideHost) return;
        var sh = document.getElementById("appShell"), root = document.documentElement;
        var r = sh && sh.getBoundingClientRect();
        if (!r || !r.width) { root.style.removeProperty("--eg-win-cx"); root.style.removeProperty("--eg-win-cy"); return; }
        root.style.setProperty("--eg-win-cx", Math.round(r.left + r.width / 2) + "px");
        root.style.setProperty("--eg-win-cy", Math.round(r.top + r.height / 2) + "px");
    }
    syncWinCenter(); setInterval(syncWinCenter, 120); window.addEventListener("resize", function () { requestAnimationFrame(syncWinCenter); });

    if (scenarioLesson) { initScenario(); return; }
    if (!lesson || !lesson.steps) { return; } /* free practice: just the top bar */

    /* =================================================================
     *  Guided lessons
     * ================================================================= */
    var steps = lesson.steps;
    var S = { idx: 0, phase: "loading", enteredAt: 0, doneAt: 0, clicked: false, ack: false, review: -1,
              missingSince: 0, lastTargetEl: null, keyboard: false, finished: false, lastHintAt: 0 };

    /* ---- DOM ---- */
    var coach = document.createElement("section");
    coach.className = "eg-ui eg-coach";
    coach.setAttribute("aria-label", "Lesson guide");
    coach.innerHTML =
        '<div class="eg-coach-head"><span class="eg-goal">' + esc(lesson.title) + '</span><span class="eg-count" aria-live="off"></span><button type="button" class="eg-min" aria-pressed="false"><span>Hide text</span></button></div>' +
        '<div class="eg-bar" role="progressbar" aria-label="Lesson progress" aria-valuemin="0"><span class="eg-bar-fill"></span></div>' +
        '<div class="eg-state" aria-hidden="true"><span class="eg-state-dot"></span><span class="eg-state-text"></span></div>' +
        '<div class="eg-live" aria-live="polite" aria-atomic="true"><h2 class="eg-title" tabindex="-1"></h2><p class="eg-msg"></p></div>' +
        '<p class="eg-hint" role="status" aria-live="assertive"></p>' +
        '<div class="eg-actions">' +
        '<button type="button" class="eg-btn eg-btn--ghost eg-back">' + ICON_BACK + "<span>Back</span></button>" +
        '<button type="button" class="eg-btn eg-btn--ghost eg-doit">Do this step for me</button>' +
        '<button type="button" class="eg-btn eg-next"><span>Next</span>' + ICON_NEXT + "</button>" +
        "</div>";
    document.body.appendChild(coach);
    var Q = function (s) { return coach.querySelector(s); };
    var elCount = Q(".eg-count"), elBar = Q(".eg-bar"), elFill = Q(".eg-bar-fill"), elTitle = Q(".eg-title"), elMsg = Q(".eg-msg"),
        elHint = Q(".eg-hint"), elState = Q(".eg-state"), elStateText = Q(".eg-state-text"),
        bBack = Q(".eg-back"), bDo = Q(".eg-doit"), bNext = Q(".eg-next"), bMin = Q(".eg-min");
    bMin.addEventListener("click", function () { userMin = !S.compact; lastSpot = null; tick(); });
    elBar.setAttribute("aria-valuemax", String(steps.length));

    var ring = document.createElement("div");
    ring.className = "eg-ring";
    ring.setAttribute("aria-hidden", "true");
    ring.innerHTML = '<span class="eg-ring-check">' + ICON_CHECK + "</span>";
    var hand = document.createElement("div");
    hand.className = "eg-hand";
    hand.setAttribute("aria-hidden", "true");
    hand.innerHTML = HAND;
    document.body.appendChild(ring);
    document.body.appendChild(hand);

    /* ---- helpers ---- */
    function target(step) {
        var t = step && step.target;
        if (!t) return null;
        try { return typeof t === "function" ? t() : document.querySelector(t); } catch (e) { return null; }
    }
    function vis(e) { return !!e && H.visible(e); }
    function norm(s) { return String(s == null ? "" : s).trim().replace(/\s+/g, " ").toLowerCase(); }
    function sameValue(a, b) {
        var na = parseFloat(a), nb = parseFloat(b);
        if (/^\s*-?[\d.]+\s*$/.test(a) && /^\s*-?[\d.]+\s*$/.test(b) && isFinite(na) && isFinite(nb)) return na === nb;
        return norm(a) === norm(b);
    }
    function selectedLabels(sel) {
        var out = [];
        for (var i = 0; i < sel.options.length; i++) if (sel.options[i].selected) out.push(norm(sel.options[i].text));
        return out;
    }
    function isDone(step, t) {
        switch (step.do) {
            case "click": return S.clicked && (!step.until || !!step.until());
            case "type": return !!t && sameValue(t.value, step.value);
            case "select": return !!t && t.options && selectedLabels(t).indexOf(norm(step.value)) >= 0;
            case "check": return !!t && t.checked === true;
            case "uncheck": return !!t && t.checked === false;
            case "wait": return !!step.until && !!step.until();
            case "read": return S.ack;
        }
        return false;
    }
    function verb(step) {
        return { click: "Click", type: "Type", select: "Choose", check: "Check the box",
                 uncheck: "Uncheck the box", read: "Read, then press Next", wait: "Watch" }[step.do] || "Action";
    }
    function setState(kind, text) {
        elState.className = "eg-state eg-state--" + kind;
        elStateText.textContent = text;
        coach.setAttribute("data-phase", kind);
        relayout();
    }
    /* re-place the card as soon as its text or size changes, not on the next tick */
    function relayout() {
        if (!coach || S.finished || S.phase === "loading" || coach.hidden) return;
        var i = S.review >= 0 ? S.review : S.idx, st = steps[i], t = target(st);
        if (frozen(t)) return;
        placeCoach(vis(t) ? t : null, st);
    }
    /* once a step is done and the item it used has gone (a Close or Finish button), the card stays put:
       the window behind it is closing, so there is nothing to move away from */
    function frozen(t) { return S.review < 0 && S.phase === "done" && !vis(t) && !!lastSpot; }
    var hintTimer = 0;
    function hint(text, strong) {
        elHint.textContent = text;
        elHint.classList.add("on");
        relayout();
        clearTimeout(hintTimer);
        hintTimer = setTimeout(function () { elHint.classList.remove("on"); elHint.textContent = ""; }, 3200);
        if (strong && !D.reduceMotion()) {
            coach.classList.remove("eg-nudge"); void coach.offsetWidth; coach.classList.add("eg-nudge");
            ring.classList.remove("eg-ring--ping"); void ring.offsetWidth; ring.classList.add("eg-ring--ping");
        }
    }

    /* ---- rendering ---- */
    function renderText(i, reviewing) {
        var st = steps[i];
        elCount.textContent = (reviewing ? "Reviewing step " : "Step ") + (i + 1) + " of " + steps.length;
        elTitle.textContent = st.title;
        elMsg.innerHTML = (typeof st.msg === "function" ? st.msg() : st.msg) + (st.do === "type" ? ' <span class="eg-tip">Tip: click the box, select the old text, then type.</span>' : "");
        elFill.style.width = ((S.idx + (S.phase === "done" ? 1 : 0)) / steps.length * 100) + "%";
        elBar.setAttribute("aria-valuenow", String(S.idx));
        elBar.setAttribute("aria-valuetext", "Step " + (S.idx + 1) + " of " + steps.length);
        relayout();
    }
    function updateButtons() {
        var st = steps[S.idx];
        bBack.disabled = (S.review >= 0 ? S.review : S.idx) === 0;
        var nextOn = false, nextLabel = "Next";
        if (S.review >= 0) { nextOn = true; nextLabel = S.review + 1 >= S.idx ? "Back to current step" : "Next"; }
        else if (S.phase === "ready" && st.do === "read") nextOn = true;
        else if (S.phase === "done" && D.get().pace === "manual") nextOn = true;
        bNext.disabled = !nextOn;
        bNext.querySelector("span").textContent = nextLabel;
        bNext.classList.toggle("eg-btn--pulse", nextOn && S.review < 0);
        bNext.setAttribute("aria-label", nextOn ? nextLabel : "Next, available after you finish this step");
        var h0 = coach.offsetHeight;
        bDo.hidden = S.review >= 0 || st.do === "read" || st.do === "wait";
        bDo.disabled = S.phase !== "ready";
        if (coach.offsetHeight !== h0) relayout();   /* the button row changed the card height */
    }

    function enter(i, opts) {
        S.idx = i; S.phase = "entering"; S.enteredAt = now(); S.clicked = false; S.ack = false; S.missingSince = 0; S.review = -1; S.scrolled = false;
        ring.className = "eg-ring"; hand.className = "eg-hand"; lastSpot = null;
        renderText(i, false);
        setState("wait", "Please wait");
        updateButtons();
        if (!opts || !opts.quiet) { try { elTitle.focus({ preventScroll: true }); } catch (e) {} }
    }

    function placeRing(t, st) {
        var r = t.getBoundingClientRect();
        var pad = 5;
        ring.style.left = (r.left - pad) + "px";
        ring.style.top = (r.top - pad) + "px";
        ring.style.width = (r.width + pad * 2) + "px";
        ring.style.height = (r.height + pad * 2) + "px";
        var showHand = st.do !== "read" && st.do !== "wait" && S.phase !== "done" && S.review < 0;
        hand.classList.toggle("on", showHand && ring.classList.contains("on"));
        /* the hand sits just outside the target, its fingertip touching the ring, so it never hides the target */
        var hw = hand.offsetWidth || 40, hh = hand.offsetHeight || 40;
        var tipX = r.left + Math.min(r.width * 0.6, r.width - 6);
        var below = r.bottom + pad + hh + 4 <= window.innerHeight;
        hand.classList.toggle("eg-hand--up", !below);
        hand.style.left = Math.max(2, Math.min(window.innerWidth - hw - 2, tipX - hw * 0.42)) + "px";
        hand.style.top = (below ? r.bottom + pad : r.top - pad - hh) + "px";
    }

    /* ---- no-overlap placement ----
     * The coach card may never cover the circled target, its ring and check badge, the hand,
     * or a reading the step talks about (step.reading). Candidate spots are tried in order and the
     * first clear one wins. If nothing is clear the card shrinks to a compact bar and tries again. */
    function box(r) { return { l: r.left, t: r.top, r: r.right, b: r.bottom }; }
    /* a top bar button that sits over the circled item steps aside until the item moves */
    function clearTopbar(rects) {
        var tb = document.querySelector(".eg-topbar"); if (!tb) return;
        for (var i = 0; i < tb.children.length; i++) {
            var ch = tb.children[i]; ch.classList.remove("eg-tb-away");
            var r = ch.getBoundingClientRect(); if (!r.width) continue;
            var me = { l: r.left - 4, t: r.top - 4, r: r.right + 4, b: r.bottom + 4 };
            for (var k = 0; k < rects.length; k++) { var a = rects[k]; if (a && me.l < a.r && me.r > a.l && me.t < a.b && me.b > a.t) { ch.classList.add("eg-tb-away"); break; } }
        }
    }
    function grow(a, p) { return { l: a.l - p, t: a.t - p, r: a.r + p, b: a.b + p }; }
    function hits(a, b) { return a.l < b.r && a.r > b.l && a.t < b.b && a.b > b.t; }
    function readingEls(st) {
        var rd = st && st.reading, out = [];
        if (!rd) return out;
        try { rd = typeof rd === "function" ? rd() : document.querySelectorAll(rd); } catch (e) { rd = null; }
        if (!rd) return out;
        if (rd.nodeType === 1) rd = [rd];
        for (var i = 0; i < rd.length; i++) if (rd[i] && vis(rd[i])) out.push(rd[i]);
        return out;
    }
    function avoidRects(t, st) {
        var list = [];
        if (t && st && st.ring === false) t = null;
        if (t) {
            var sc = parseFloat(getComputedStyle(document.documentElement).getPropertyValue("--eg-scale")) || 1;
            list.push(grow(box(t.getBoundingClientRect()), 8 + Math.ceil(16 * sc)));
            if (hand.classList.contains("on")) list.push(grow(box(hand.getBoundingClientRect()), 4));
        }
        readingEls(st).forEach(function (e) { list.push(grow(box(e.getBoundingClientRect()), 6)); });
        return list;
    }
    function topLimit() {
        var tb = document.querySelector(".eg-topbar");
        return tb ? Math.ceil(tb.getBoundingClientRect().bottom) + 6 : 12;
    }
    var lastSpot = null;
    function findSpot(cw, ch, avoid, shellRect) {
        var vw = window.innerWidth, vh = window.innerHeight, m = 12, top = topLimit();
        if (cw > vw - 2 * m || ch > vh - top - m) return null;
        function clampX(x) { return Math.max(m, Math.min(vw - cw - m, x)); }
        function clampY(y) { return Math.max(top, Math.min(vh - ch - m, y)); }
        var c = [];
        if (composed && shellRect && shellRect.width > 0) {
            /* the reserved slot: top-aligned with the window, or centered on it when the card is taller */
            var dy = ch <= shellRect.height ? shellRect.top : shellRect.top - (ch - shellRect.height) / 2;
            c.push({ x: shellRect.left - cw - 16, y: dy, dock: true });
        }
        if (lastSpot) c.push({ x: lastSpot.x, y: lastSpot.y, dock: lastSpot.dock });
        if (shellRect && shellRect.width > 0) {
            if (shellRect.left >= cw + 28) c.push({ x: shellRect.left - cw - 16, y: shellRect.top, dock: true });
            if (vw - shellRect.right >= cw + 28) c.push({ x: shellRect.right + 16, y: shellRect.top, dock: true });
        }
        c.push({ x: (vw - cw) / 2, y: vh - ch - m }, { x: (vw - cw) / 2, y: top });
        avoid.forEach(function (a) {
            var cx = (a.l + a.r) / 2 - cw / 2, cy = (a.t + a.b) / 2 - ch / 2;
            c.push({ x: cx, y: a.b + 4 }, { x: cx, y: a.t - ch - 4 }, { x: a.r + 4, y: cy }, { x: a.l - cw - 4, y: cy },
                   { x: a.r + 4, y: a.t }, { x: a.l - cw - 4, y: a.t }, { x: a.r + 4, y: a.b - ch }, { x: a.l - cw - 4, y: a.b - ch });
        });
        c.push({ x: m, y: vh - ch - m }, { x: vw - cw - m, y: vh - ch - m }, { x: m, y: top }, { x: vw - cw - m, y: top },
               { x: m, y: (vh - ch) / 2 }, { x: vw - cw - m, y: (vh - ch) / 2 });
        var inWin = shellRect && shellRect.width > 0 ? avoid.map(function (a) {
            return a.r > shellRect.left && a.l < shellRect.right ? { l: Math.max(a.l, shellRect.left - 4), t: a.t, r: Math.min(a.r, shellRect.right + 4), b: a.b } : a;
        }) : avoid;
        for (var i = 0; i < c.length; i++) {
            var x = clampX(c[i].x), y = clampY(c[i].y);
            var me = { l: x, t: y, r: x + cw, b: y + ch }, ok = true, av = c[i].dock ? inWin : avoid;
            for (var k = 0; k < av.length; k++) if (hits(me, av[k])) { ok = false; break; }
            if (ok) return { x: x, y: y, dock: !!c[i].dock };
        }
        return null;
    }
    /* Centered composition: when the docked card and the window fit side by side, the host page
       reserves --eg-dock-w to the left of the window (html.eg-composed), so card + gap + window are
       centered as one unit. Decided per viewport and text size, not per step, so the window never jumps. */
    var composeKey = "", composed = false;
    function compose() {
        var root = document.documentElement, tb = document.querySelector(".eg-topbar");
        var topPx = tb ? Math.ceil(tb.getBoundingClientRect().bottom) : 0;
        var hadDock = coach.classList.contains("eg-coach--dock"), hadMin = coach.classList.contains("eg-coach--min");
        coach.classList.add("eg-coach--dock"); coach.classList.remove("eg-coach--min");
        var cw = coach.offsetWidth;
        coach.classList.toggle("eg-coach--dock", hadDock); coach.classList.toggle("eg-coach--min", hadMin);
        var key = window.innerWidth + "x" + window.innerHeight + ":" + topPx + ":" + cw;
        if (key === composeKey) return composed;
        /* wait for the host to be ready, or its window is not sized yet */
        if (!HOST.ready || HOST.ready()) { composeKey = key; markPlaced(); }
        root.style.setProperty("--eg-top", topPx + "px");
        root.style.setProperty("--eg-dock-w", (cw + 16) + "px");
        root.classList.add("eg-composed");
        if (HOST.refit) HOST.refit();
        var shell = document.getElementById(HOST.shellId), sr = shell ? shell.getBoundingClientRect() : null;
        composed = !!(sr && sr.width > 0 && sr.left - cw - 16 >= 12 && sr.right <= window.innerWidth - 12 &&
            (!HOST.minShellW || sr.width >= HOST.minShellW));
        if (!composed) { root.classList.remove("eg-composed"); if (HOST.refit) HOST.refit(); }
        lastSpot = null;
        return composed;
    }
    /* the window is shown only once its place is decided, so it never appears in one spot and jumps to another */
    function markPlaced() { requestAnimationFrame(function () { document.documentElement.classList.add("eg-placed"); }); }
    var userMin = false, lastSR = null;
    function placeCoach(t, st) {
        compose();
        var shell = document.getElementById(HOST.shellId);
        var sr = shell ? shell.getBoundingClientRect() : null;
        /* while the window is closed for a moment (a restart), keep the card in its column */
        if (sr && sr.width > 0) lastSR = sr; else if (composed && lastSR) sr = lastSR;
        var avoid = avoidRects(t, st);
        var tries = userMin ? [true] : [false, true];
        var spot = null, compact = userMin;
        coach.classList.remove("eg-coach--narrow");
        for (var i = 0; i < tries.length && !spot; i++) {
            compact = tries[i];
            coach.classList.toggle("eg-coach--min", compact);
            /* measure as docked width first, then as the wide card */
            var dockFirst = !!(sr && sr.width > 0 && (sr.left >= 300 || window.innerWidth - sr.right >= 300));
            var widths = dockFirst ? [true, false] : [false];
            for (var w = 0; w < widths.length && !spot; w++) {
                coach.classList.toggle("eg-coach--dock", widths[w]);
                var cw = coach.offsetWidth, ch = coach.offsetHeight;
                spot = findSpot(cw, ch, avoid, widths[w] ? sr : null);
                if (spot && widths[w] && !spot.dock) spot = null;
                if (!spot && widths[w]) lastSpot = null;
            }
        }
        if (!spot) {
            /* a narrow compact card can still fit beside the target on short, large-text screens */
            coach.classList.add("eg-coach--min", "eg-coach--narrow"); coach.classList.remove("eg-coach--dock");
            spot = findSpot(coach.offsetWidth, coach.offsetHeight, avoid, null);
            if (spot) compact = true; else coach.classList.remove("eg-coach--narrow");
        }
        if (!spot) {
            /* nothing is clear: keep the compact bar at the bottom or top, whichever covers less */
            coach.classList.add("eg-coach--min"); coach.classList.remove("eg-coach--dock");
            var cw2 = coach.offsetWidth, ch2 = coach.offsetHeight, vh = window.innerHeight;
            var tr = t ? t.getBoundingClientRect() : null;
            spot = { x: Math.max(12, (window.innerWidth - cw2) / 2), y: tr && tr.top + tr.height / 2 > vh / 2 ? topLimit() : vh - ch2 - 12 };
            compact = true;
        }
        lastSpot = spot;
        coach.style.left = Math.round(spot.x) + "px";
        coach.style.top = Math.round(spot.y) + "px";
        coach.style.bottom = "auto";
        bMin.setAttribute("aria-pressed", String(userMin));
        bMin.querySelector("span").textContent = compact ? "Show text" : "Hide text";
        S.compact = compact;
        clearTopbar(t && !(st && st.ring === false) ? [box(t.getBoundingClientRect()), ring.classList.contains("on") ? box(ring.getBoundingClientRect()) : null] .concat(readingEls(st).map(function (e) { return box(e.getBoundingClientRect()); })) : []);
    }

    /* the card can change height after it was placed (state badge, hint, fonts): place it again */
    var roPending = false;
    if (window.ResizeObserver) new ResizeObserver(function () {
        if (roPending) return; roPending = true;
        requestAnimationFrame(function () { roPending = false; relayout(); });
    }).observe(coach);

    function tick() {
        if (S.finished) return;
        if (S.phase === "loading") {
            if (appReady()) { enter(0, { quiet: true }); }
            else { setState("wait", "Binventory is starting"); placeCoach(null, null); return; }
        }
        var i = S.review >= 0 ? S.review : S.idx;
        var st = steps[i];
        var t = target(st);
        var tv = vis(t);
        if (tv && !S.scrolled) {
            S.scrolled = true;
            var r0 = t.getBoundingClientRect();
            if (r0.top < 56 || r0.bottom > window.innerHeight - 8) {
                try { t.scrollIntoView({ block: "center", behavior: D.reduceMotion() ? "auto" : "smooth" }); } catch (e) {}
            }
        }
        if (tv) placeRing(t, st);
        if (!frozen(t)) placeCoach(tv ? t : null, st);

        if (S.review >= 0) {
            if (tv) { ring.classList.add("on", "eg-ring--review"); placeRing(t, st); } else { ring.classList.remove("on"); hand.classList.remove("on"); }
            return;
        }

        if (!tv) {
            ring.classList.remove("on"); hand.classList.remove("on");
            /* a clicked menu item or Close button disappears as it works: finish the step without the ring */
            if ((S.phase === "ready" && st.do === "click" && S.clicked) || S.phase === "done") { S.missingSince = 0; stepProgress(st, t); return; }
            if (!S.missingSince) S.missingSince = now();
            var prev = steps[S.idx - 1];
            if (S.phase !== "done" && prev && prev.opensMenu && now() - S.missingSince > 900) {
                enter(S.idx - 1, { quiet: true });
                hint("The menu closed. Open it again.", false);
                return;
            }
            if (S.phase === "ready") setState("wait", HOST.waitText || "Waiting for Binventory");
            return;
        }
        S.missingSince = 0;
        if (!ring.classList.contains("on") && st.ring !== false) {
            ring.classList.add("on");
            ring.classList.toggle("eg-ring--look", st.do === "read" || st.do === "wait");
            ring.style.setProperty("--eg-ring-ms", D.pace().ring + "ms");
            ring.classList.add("eg-ring--enter");
            placeRing(t, st);
            placeCoach(t, st);
        }
        placeRing(t, st);
        if (st.do === "type" && !t.__egSelectOnFocus) {
            t.__egSelectOnFocus = true;
            t.addEventListener("focus", function () { try { if (steps[S.idx] && target(steps[S.idx]) === t) t.select(); } catch (e) {} });
        }

        if (S.phase === "entering") {
            if (now() - S.enteredAt >= D.pace().settle) {
                S.phase = "ready";
                ring.classList.remove("eg-ring--enter");
                setState(st.do === "read" ? "read" : (st.do === "wait" ? "watch" : "go"), verb(st));
                updateButtons();
                if (S.keyboard && st.do !== "read" && st.do !== "wait") { try { t.focus(); } catch (e) {} }
            }
            return;
        }
        stepProgress(st, t);
    }

    function stepProgress(st, t) {
        if (S.phase === "ready") {
            if (isDone(st, t)) {
                S.phase = "done"; S.doneAt = now();
                ring.classList.add("eg-ring--done"); hand.classList.remove("on");
                setState("done", D.get().pace === "manual" && st.do !== "read" ? "Done. Press Next" : "Done");
                renderText(S.idx, false);
                updateButtons();
            }
            return;
        }
        if (S.phase === "done") {
            /* a typed value or checkbox that changes back undoes the step */
            if (t && (st.do === "type" || st.do === "select" || st.do === "check" || st.do === "uncheck") && !isDone(st, t)) {
                S.phase = "ready"; ring.classList.remove("eg-ring--done");
                setState("go", verb(st)); updateButtons(); return;
            }
            if (D.get().pace === "manual" && st.do !== "read" && !(S.idx + 1 >= steps.length && !vis(t))) return;
            var holdMs = S.idx + 1 >= steps.length && !vis(t) ? Math.min(450, D.pace().hold) : D.pace().hold;
            if (now() - S.doneAt >= holdMs) advance();
        }
    }

    function advance() {
        if (S.idx + 1 >= steps.length) return finish();
        enter(S.idx + 1);
    }

    function appReady() { return HOST.ready(); }

    /* ---- input guard: only the circled item, and only once it is ready ---- */
    function inGuide(e) { return !!(e.target && e.target.closest && e.target.closest(".eg-ui, .eg-success")); }
    function inTarget(e) {
        var st = steps[S.idx], t = target(st);
        return !!t && (t === e.target || t.contains(e.target) || (e.target.tagName === "OPTION" && t.contains(e.target)));
    }
    function guard(e) {
        if (S.finished || S.phase === "loading" || inGuide(e)) return;
        var block = true, msg = "";
        var st = steps[S.idx];
        if (S.review >= 0) msg = "Earlier step. Press Back to current step to continue.";
        else if (S.phase === "entering") msg = "Wait for the red circle, then " + st.title.toLowerCase() + ".";
        else if (S.phase === "done") msg = "Done.";
        else if (st.do === "read") msg = "Read this step, then press Next.";
        else if (st.do === "wait") msg = "Wait. This step continues on its own.";
        else if (inTarget(e)) block = false;
        else msg = "Use the item in the red circle.";
        if (!block) { if (e.type === "click") S.clicked = true; return; }
        e.preventDefault(); e.stopPropagation(); if (e.stopImmediatePropagation) e.stopImmediatePropagation();
        if (e.type === "pointerdown" || (e.type === "click" && now() - S.lastHintAt > 400)) { S.lastHintAt = now(); hint(msg, true); }
    }
    ["pointerdown", "mousedown", "mouseup", "click", "dblclick", "contextmenu", "auxclick"].forEach(function (t) {
        window.addEventListener(t, guard, true);
    });
    window.addEventListener("keydown", function (e) {
        if (e.key === "Tab") S.keyboard = true;
        if (S.finished || S.phase === "loading" || inGuide(e)) return;
        if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); return; }
        var st = steps[S.idx];
        var typing = e.key.length === 1 || e.key === "Backspace" || e.key === "Delete";
        var activate = e.key === "Enter" || e.key === " ";
        var focusedTarget = inTarget(e);
        /* typing or pressing Enter in a field that is not this step's field is held back */
        if ((typing || activate || e.key.indexOf("Arrow") === 0) && !(focusedTarget && S.phase === "ready" && S.review < 0)) {
            var tag = (e.target.tagName || "").toLowerCase();
            if (tag === "input" || tag === "select" || tag === "textarea" || tag === "button" || activate) {
                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
                hint(S.phase === "ready" ? "Use the item in the red circle." : "Wait for the red circle.", true);
            }
        }
    }, true);
    window.addEventListener("pointerdown", function () { S.keyboard = false; }, true);

    /* ---- buttons ---- */
    bNext.addEventListener("click", function () {
        if (bNext.disabled) return;
        if (S.review >= 0) {
            S.review++;
            if (S.review >= S.idx) { S.review = -1; ring.classList.remove("eg-ring--review"); ring.classList.remove("on"); renderText(S.idx, false); setState(S.phase === "done" ? "done" : (S.phase === "ready" ? "go" : "wait"), S.phase === "ready" ? verb(steps[S.idx]) : "Please wait"); }
            else { renderText(S.review, true); setState("review", "Earlier step"); }
            ring.classList.remove("on");
            updateButtons(); return;
        }
        var st = steps[S.idx];
        if (S.phase === "ready" && st.do === "read") { S.ack = true; return; }
        if (S.phase === "done") advance();
    });
    bBack.addEventListener("click", function () {
        var from = S.review >= 0 ? S.review : S.idx;
        if (from <= 0) return;
        S.review = from - 1;
        ring.classList.remove("on"); hand.classList.remove("on");
        renderText(S.review, true);
        setState("review", "Earlier step. Your changes are kept.");
        updateButtons();
    });
    bDo.addEventListener("click", function () {
        if (S.phase !== "ready" || S.review >= 0) return;
        var st = steps[S.idx], t = target(st);
        if (!t) return;
        doStep(st, t);
    });
    function fire(t, type) { t.dispatchEvent(new Event(type, { bubbles: true })); }
    function doStep(st, t) {
        if (st.do === "click") { t.click(); }
        else if (st.do === "type") { t.focus(); t.value = st.value; fire(t, "input"); fire(t, "change"); }
        else if (st.do === "select") {
            for (var i = 0; i < t.options.length; i++) {
                var hit = norm(t.options[i].text) === norm(st.value);
                if (t.multiple || t.size > 1) t.options[i].selected = hit; else if (hit) t.selectedIndex = i;
            }
            fire(t, "input"); fire(t, "change");
        }
        else if (st.do === "check" && !t.checked) t.click();
        else if (st.do === "uncheck" && t.checked) t.click();
    }

    /* ---- finish ----
     * The card fades out where it is, then one centered "Lesson complete" panel fades in.
     * Nothing re-places during the change; with reduced motion both happen at once. */
    function finish() {
        if (S.finished) return;
        S.finished = true;
        markDone(lesson.id);
        ring.className = "eg-ring"; hand.className = "eg-hand";   /* gone at once, no fading ring left at the old spot */
        clearTimeout(hintTimer); elHint.classList.remove("on");
        var list = LS.lessons, at = list.indexOf(lesson), next = null;
        for (var k = at + 1; k < list.length; k++) { if (list[k].steps || list[k].scenario || list[k].doc) { next = list[k]; break; } }
        var ov = document.createElement("div");
        ov.className = "eg-success eg-success--enter" + (HOST.shellId === "appShell" ? " eg-success--dim" : "");
        ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true"); ov.setAttribute("aria-labelledby", "egSuccessTitle");
        ov.innerHTML =
            '<div class="eg-success-card">' +
            '<svg class="eg-check" viewBox="0 0 52 52" aria-hidden="true"><circle class="eg-check-c" cx="26" cy="26" r="24"/><path class="eg-check-p" d="M15 27l7.5 7.5L38 19"/></svg>' +
            '<p class="eg-success-kicker">Lesson complete</p>' +
            '<h2 id="egSuccessTitle" tabindex="-1">' + esc(lesson.title) + "</h2>" +
            (lesson.outro ? '<p class="eg-success-text">' + esc(lesson.outro) + "</p>" : "") +
            (next ? '<p class="eg-success-next">Next: ' + esc(next.title) + "</p>" : "") +
            '<div class="eg-success-actions">' +
            (next ? '<a class="eg-btn eg-btn--big" href="' + lessonHref(next) + '" aria-label="Next lesson: ' + esc(next.title) + '">Next lesson' + ICON_NEXT + "</a>" : "") +
            '<div class="eg-success-row">' +
            '<a class="eg-btn eg-btn--light" href="index.html">' + ICON_GRID + "Back to guide</a>" +
            '<a class="eg-btn eg-btn--light" href="' + lessonHref(lesson) + '">Repeat</a>' +
            "</div></div></div>";
        var fadeMs = D.reduceMotion() ? 0 : 160;
        coach.classList.add("eg-coach--leaving");
        setTimeout(function () {
            coach.hidden = true;
            document.documentElement.classList.add("eg-lesson-done");
            document.body.appendChild(ov);
            void ov.offsetWidth;
            ov.classList.remove("eg-success--enter");
            var h2 = ov.querySelector("h2"); if (h2) { try { h2.focus({ preventScroll: true }); } catch (e) { h2.focus(); } }
        }, fadeMs);
    }

    /* expose a small, read-only view for tests and support */
    window.EbobGuide = {
        /* test hook: where the guide is and what it must not cover */
        layout: function () {
            function R(e) { if (!e) return null; var r = e.getBoundingClientRect(); return { l: r.left, t: r.top, r: r.right, b: r.bottom }; }
            var i = S.review >= 0 ? S.review : S.idx, st = steps[i], t = target(st);
            return { idx: i, phase: S.phase, coach: coach.hidden ? null : R(coach), compact: !!S.compact,
                     target: vis(t) && st.ring !== false ? R(t) : null, ring: ring.classList.contains("on") ? R(ring) : null,
                     hand: hand.classList.contains("on") ? R(hand) : null,
                     reading: readingEls(st).map(R), vw: window.innerWidth, vh: window.innerHeight };
        },
        state: function () { return { lesson: lesson.id, idx: S.idx, total: steps.length, phase: S.phase, review: S.review, finished: S.finished }; },
        /* test hook: tag the current target so automated checks can find it */
        markTarget: function () {
            var old = document.querySelectorAll("[data-eg-current]");
            for (var i = 0; i < old.length; i++) old[i].removeAttribute("data-eg-current");
            var st = steps[S.idx], t = target(st);
            if (!t) return null;
            t.setAttribute("data-eg-current", "1");
            return { do: st.do, value: st.value == null ? null : String(st.value), tag: t.tagName, title: st.title };
        }
    };

    setInterval(tick, 60);
    window.addEventListener("resize", function () { lastSpot = null; lastSR = null; tick(); });
    window.addEventListener("scroll", tick, true);
    D.onChange(function () { lastSpot = null; if (!S.finished && S.phase !== "loading") { renderText(S.review >= 0 ? S.review : S.idx, S.review >= 0); updateButtons(); } });

    /* =================================================================
     *  Built-in troubleshooting scenarios
     * ================================================================= */
    function initScenario() {
        document.documentElement.classList.add("eg-scenario");
        var tip = document.getElementById("ebobTutorialTooltip");
        if (!tip) return;
        /* a status line and hint inside the scenario card */
        var stateEl = document.createElement("div");
        stateEl.className = "eg-state eg-state--wait eg-native-state";
        stateEl.innerHTML = '<span class="eg-state-dot"></span><span class="eg-state-text">Please wait</span>';
        var hintEl = document.createElement("p");
        hintEl.className = "eg-hint"; hintEl.setAttribute("role", "status"); hintEl.setAttribute("aria-live", "assertive");
        var reviewEl = document.createElement("div");
        reviewEl.className = "eg-ui eg-native-review"; reviewEl.hidden = true;
        var backBtn = document.createElement("button");
        backBtn.type = "button"; backBtn.className = "eg-ui eg-native-back"; backBtn.innerHTML = ICON_BACK + "<span>Back</span>";
        backBtn.setAttribute("aria-label", "Show the previous step again");
        var titleEl = document.getElementById("ebobTutorialTitle");
        tip.insertBefore(stateEl, titleEl);
        var actions = tip.querySelector(".ebob-tutorial-actions");
        tip.insertBefore(hintEl, actions);
        tip.insertBefore(reviewEl, actions);
        if (actions) actions.insertBefore(backBtn, actions.firstChild);
        var bar = document.createElement("div");
        bar.className = "eg-bar"; bar.setAttribute("role", "progressbar"); bar.setAttribute("aria-label", "Scenario progress");
        bar.innerHTML = '<span class="eg-bar-fill"></span>';
        tip.insertBefore(bar, stateEl);

        var lastIdx = -1, enteredAt = 0, reviewAt = -1, markedDone = false, ht = 0;
        /* keep the scenario card off the circled item and its pointer */
        function nb(e) { if (!e || e.classList.contains("ebob-tutorial-hidden")) return null; var r = e.getBoundingClientRect(); if (!r.width) return null; return { l: r.left - 14, t: r.top - 14, r: r.right + 14, b: r.bottom + 14 }; }
        function hitN(a, b) { return a.l < b.r && a.r > b.l && a.t < b.b && a.b > b.t; }
        /* same layout rule as the guided lessons: when the card and the window fit side by side, reserve a
           column left of the window (html.eg-composed) so card + window are centered as one unit */
        var nKey = "", nComposed = false;
        function composeNative() {
            var root = document.documentElement, tb = document.querySelector(".eg-topbar");
            var topPx = tb ? Math.ceil(tb.getBoundingClientRect().bottom) : 0;
            var had = tip.className;
            tip.classList.remove("eg-native--narrow", "eg-native--min"); tip.classList.add("eg-native--dock");
            var cw = tip.offsetWidth;
            tip.className = had;
            var key = innerWidth + "x" + innerHeight + ":" + topPx + ":" + cw;
            if (key === nKey) return nComposed;
            if (HOST.ready()) nKey = key;
            root.style.setProperty("--eg-top", topPx + "px");
            root.style.setProperty("--eg-dock-w", (cw + 16) + "px");
            root.classList.add("eg-composed");
            HOST.refit();
            var sh = document.getElementById(HOST.shellId), sr = sh ? sh.getBoundingClientRect() : null;
            nComposed = !!(sr && sr.width > 0 && sr.left - cw - 16 >= 12 && sr.right <= innerWidth - 12 && sr.width >= HOST.minShellW);
            if (!nComposed) { root.classList.remove("eg-composed"); HOST.refit(); }
            return nComposed;
        }
        function placeNative() {
            if (tip.classList.contains("ebob-tutorial-hidden")) return;
            var av = [nb(document.getElementById("ebobTutorialCircle")), nb(document.getElementById("ebobTutorialPointer"))].filter(Boolean);
            var vw = window.innerWidth, vh = window.innerHeight, m = 12, tb = document.querySelector(".eg-topbar");
            var top = tb ? tb.getBoundingClientRect().bottom + 6 : 12;
            clearTopbar(av);
            tip.style.setProperty("transform", "none", "important"); tip.style.bottom = "auto";
            var docked = composeNative();
            /* the pointer stays out of the card column: it points from the right when the item is near the window's left edge */
            var shR = docked ? document.getElementById(HOST.shellId).getBoundingClientRect() : null;
            var minX = shR ? Math.round(shR.left) : 0;
            window.__ebobPointerMinX = minX;
            if (docked) {
                var sr = document.getElementById(HOST.shellId).getBoundingClientRect();
                tip.classList.remove("eg-native--narrow", "eg-native--min"); tip.classList.add("eg-native--dock");
                var dw = tip.offsetWidth, dh = tip.offsetHeight, dx = sr.left - dw - 16;
                var dy = Math.max(top, Math.min(vh - dh - m, dh <= sr.height ? sr.top : sr.top - (dh - sr.height) / 2));
                var me0 = { l: dx, t: dy, r: dx + dw, b: dy + dh }, clear = true;
                av.forEach(function (a) { var c = a.r > sr.left ? { l: Math.max(a.l, sr.left - 4), t: a.t, r: a.r, b: a.b } : a; if (hitN(me0, c)) clear = false; });
                if (clear && dh <= vh - top - m) { tip.style.left = Math.round(dx) + "px"; tip.style.top = Math.round(dy) + "px"; return; }
                tip.classList.remove("eg-native--dock");
            } else tip.classList.remove("eg-native--dock");
            /* full card, then a narrower card, then a compact card; each tries the usual spots and then a grid */
            var modes = ["", "eg-native--narrow", "eg-native--narrow eg-native--min"];
            for (var md = 0; md < modes.length; md++) {
                tip.classList.remove("eg-native--narrow", "eg-native--min");
                if (modes[md]) modes[md].split(" ").forEach(function (k) { tip.classList.add(k); });
                var cw = tip.offsetWidth, ch = tip.offsetHeight;
                var c = [[(vw - cw) / 2, vh - ch - 20], [(vw - cw) / 2, top]];
                av.forEach(function (a) { c.push([(a.l + a.r - cw) / 2, a.b + 4], [(a.l + a.r - cw) / 2, a.t - ch - 4], [a.r + 4, a.t], [a.l - cw - 4, a.t], [a.r + 4, a.b - ch], [a.l - cw - 4, a.b - ch]); });
                c.push([m, vh - ch - m], [vw - cw - m, vh - ch - m], [m, top], [vw - cw - m, top]);
                for (var gy = 0; gy <= 8; gy++) for (var gx = 0; gx <= 8; gx++) c.push([m + (vw - cw - 2 * m) * gx / 8, top + (vh - ch - m - top) * gy / 8]);
                for (var i = 0; i < c.length; i++) {
                    var x = Math.max(m, Math.min(vw - cw - m, c[i][0])), y = Math.max(top, Math.min(vh - ch - m, c[i][1]));
                    var me = { l: x, t: y, r: x + cw, b: y + ch }, ok = true;
                    for (var k = 0; k < av.length; k++) if (hitN(me, av[k])) ok = false;
                    if (ok) { tip.style.left = Math.round(x) + "px"; tip.style.top = Math.round(y) + "px"; return; }
                }
            }
            /* nothing is clear: compact card on the side away from the circle */
            var c0 = av[0], cw2 = tip.offsetWidth, ch2 = tip.offsetHeight;
            tip.style.left = Math.round(c0 && (c0.l + c0.r) / 2 < vw / 2 ? vw - cw2 - m : m) + "px";
            tip.style.top = Math.round(c0 && (c0.t + c0.b) / 2 > vh / 2 ? top : vh - ch2 - m) + "px";
        }
        window.EbobGuide = { layout: function () {
            function R(e) { if (!e || e.classList.contains("ebob-tutorial-hidden")) return null; var r = e.getBoundingClientRect(); return r.width ? { l: r.left, t: r.top, r: r.right, b: r.bottom } : null; }
            var s = st();
            return { native: true, idx: s ? s.index : -1, coach: R(tip), ring: R(document.getElementById("ebobTutorialCircle")), target: null, reading: [], vw: innerWidth, vh: innerHeight };
        }, state: function () { var s = st(); return s ? { idx: s.index, total: s.total, finished: s.completed, phase: s.pending ? "done" : "ready" } : null; } };
        window.addEventListener("resize", placeNative);
        function st() { return window.__ebobNativeTutorial ? window.__ebobNativeTutorial.state() : null; }
        function setNative(kind, text) { stateEl.className = "eg-state eg-state--" + kind + " eg-native-state"; stateEl.lastChild.textContent = text; }
        function showReview() {
            var s = st(); if (!s || reviewAt < 0) { reviewEl.hidden = true; return; }
            var p = s.steps[reviewAt];
            reviewEl.hidden = false;
            reviewEl.innerHTML = '<p class="eg-native-review-head">Earlier: step ' + (reviewAt + 1) + " of " + s.total + "</p><p><b>" + esc(p.title) + "</b> " + esc(p.message) + "</p>" +
                '<div class="eg-native-review-btns">' + (reviewAt > 0 ? '<button type="button" class="eg-ui" data-r="prev">Earlier</button>' : "") +
                '<button type="button" class="eg-ui" data-r="close">Back to current step</button></div>';
        }
        reviewEl.addEventListener("click", function (e) {
            var b = e.target.closest("button"); if (!b) return;
            if (b.getAttribute("data-r") === "prev" && reviewAt > 0) reviewAt--; else reviewAt = -1;
            showReview();
        });
        backBtn.addEventListener("click", function () {
            var s = st(); if (!s) return;
            reviewAt = reviewAt < 0 ? s.index - 1 : Math.max(0, reviewAt - 1);
            showReview();
        });
        window.addEventListener("ebobguide:native-blocked", function (e) {
            var d = e.detail || {};
            hintEl.textContent = d.done ? "Done." : (d.busy ? "Wait for the red circle." : "Use the item in the red circle.");
            hintEl.classList.add("on");
            clearTimeout(ht); ht = setTimeout(function () { hintEl.classList.remove("on"); hintEl.textContent = ""; }, 3000);
            if (!D.reduceMotion()) { tip.classList.remove("eg-nudge"); void tip.offsetWidth; tip.classList.add("eg-nudge"); }
        });
        setInterval(function () {
            var s = st(); if (!s) return;
            if (s.completed) { if (!markedDone) { markedDone = true; markDone(scenarioLesson.id); } return; }
            if (s.index !== lastIdx) { lastIdx = s.index; enteredAt = now(); reviewAt = -1; showReview(); }
            backBtn.disabled = s.index === 0;
            bar.firstChild.style.width = (s.index / s.total * 100) + "%";
            bar.setAttribute("aria-valuenow", String(s.index)); bar.setAttribute("aria-valuemax", String(s.total));
            placeNative();
            if (s.pending) setNative("done", "Done");
            else if (now() - enteredAt < D.pace().settle) setNative("wait", "Please wait");
            else setNative("go", "Action");
        }, 80);
    }

    function lessonHref(l) {
        return l.page === "installer" ? "installer.html?lesson=" + l.id : l.kind === "doc" ? "doc.html?lesson=" + l.id : l.scenario ? "emulator.html?tutorial=" + l.scenario + "&fresh=1" : "emulator.html?tutorial=free-mode&lesson=" + l.id;
    }
})();
