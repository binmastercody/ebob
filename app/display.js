/*
 * eBob Guide display settings, shared by the lesson menu and the lesson pages.
 * Saved in localStorage "ebob-guide-display": { size, pace, contrast, motion }.
 *   size     0..4  -> 90%, 100% (default), 115%, 130%, 150% of the guide text
 *   pace     "normal" | "slow" | "manual" (wait for Next after every step)
 *   contrast true for the high contrast guide
 *   motion   "system" (follow the OS setting) | "reduce" | "full"
 */
(function () {
    "use strict";
    var KEY = "ebob-guide-display";
    var SIZES = [
        { label: "Smaller", pct: 90 },
        { label: "Default", pct: 100 },
        { label: "Large", pct: 115 },
        { label: "Larger", pct: 130 },
        { label: "Largest", pct: 150 }
    ];
    var DEFAULT_SIZE = 1;
    var PACES = {
        /* settle: how long a new step's highlight animates in before your action is accepted
           hold:   how long a finished step stays on screen (green check) before the next one */
        normal: { label: "Normal", settle: 500, hold: 750, ring: 450 },
        slow: { label: "Slow", settle: 1200, hold: 1600, ring: 900 },
        manual: { label: "Step by step", settle: 500, hold: 750, ring: 450 }
    };
    var listeners = [];

    function load() {
        var d = {};
        try { d = JSON.parse(localStorage.getItem(KEY) || "{}") || {}; } catch (e) { d = {}; }
        if (typeof d.size !== "number" || d.size < 0 || d.size >= SIZES.length) d.size = DEFAULT_SIZE;
        if (!PACES[d.pace]) d.pace = "normal";
        d.contrast = !!d.contrast;
        if (["system", "reduce", "full"].indexOf(d.motion) < 0) d.motion = "system";
        return d;
    }
    var cur = load();

    function systemReduce() {
        return !!(window.matchMedia && matchMedia("(prefers-reduced-motion: reduce)").matches);
    }
    function reduceMotion() { return cur.motion === "reduce" || (cur.motion === "system" && systemReduce()); }

    function apply() {
        var c = document.documentElement.classList;
        for (var i = 0; i < SIZES.length; i++) c.remove("eg-size-" + i);
        c.add("eg-size-" + cur.size);
        document.documentElement.style.setProperty("--eg-scale", String(SIZES[cur.size].pct / 100));
        c.toggle("eg-hc", cur.contrast);
        c.toggle("eg-reduce-motion", reduceMotion());
        c.toggle("eg-pace-slow", cur.pace === "slow");
    }
    function save() {
        try { localStorage.setItem(KEY, JSON.stringify(cur)); } catch (e) { /* storage blocked */ }
        apply();
        listeners.forEach(function (fn) { try { fn(cur); } catch (e) {} });
    }
    function set(k, v) { cur[k] = v; save(); }

    /* also follow storage changes from another open tab */
    window.addEventListener("storage", function (e) { if (e.key === KEY) { cur = load(); apply(); refreshUi(); listeners.forEach(function (fn) { fn(cur); }); } });
    apply();

    var api = {
        SIZES: SIZES, PACES: PACES,
        get: function () { return cur; },
        set: set,
        onChange: function (fn) { listeners.push(fn); },
        reduceMotion: reduceMotion,
        pace: function () { return PACES[cur.pace]; },
        bigger: function () { if (cur.size < SIZES.length - 1) set("size", cur.size + 1); announce(); },
        smaller: function () { if (cur.size > 0) set("size", cur.size - 1); announce(); },
        resetSize: function () { set("size", DEFAULT_SIZE); announce(); },
        resetAll: function () { cur = { size: DEFAULT_SIZE, pace: "normal", contrast: false, motion: "system" }; save(); refreshUi(); }
    };
    window.EbobDisplay = api;
    /* used by the patched ebob-sim.js for the built-in troubleshooting scenarios */
    window.EbobGuidePace = {
        settleMs: function () { return api.pace().settle; },
        doneHoldMs: function () { return api.pace().hold; }
    };

    /* ---------- toolbar: A- A A+ and the Display panel ---------- */
    var bar, panel, live, btn;
    function h(tag, attrs, html) {
        var e = document.createElement(tag);
        if (attrs) Object.keys(attrs).forEach(function (k) { e.setAttribute(k, attrs[k]); });
        if (html != null) e.innerHTML = html;
        return e;
    }
    function announce() {
        if (live) live.textContent = "Text size " + SIZES[cur.size].label + ", " + SIZES[cur.size].pct + " percent";
    }
    function refreshUi() {
        if (!bar) return;
        bar.querySelector("[data-eg-size=down]").disabled = cur.size === 0;
        bar.querySelector("[data-eg-size=up]").disabled = cur.size === SIZES.length - 1;
        bar.querySelector(".eg-size-val").textContent = SIZES[cur.size].pct + "%";
        if (panel) {
            panel.querySelectorAll("input[name=eg-size]").forEach(function (r) { r.checked = +r.value === cur.size; });
            panel.querySelectorAll("input[name=eg-pace]").forEach(function (r) { r.checked = r.value === cur.pace; });
            panel.querySelector("#egHc").checked = cur.contrast;
            panel.querySelector("#egMotion").checked = reduceMotion();
        }
    }
    function closePanel(focusBtn) {
        if (!panel || panel.hidden) return;
        panel.hidden = true; btn.setAttribute("aria-expanded", "false");
        if (focusBtn) btn.focus();
    }
    function buildToolbar(host) {
        bar = h("div", { "class": "eg-ui eg-toolbar", role: "toolbar", "aria-label": "Display settings" });
        bar.innerHTML =
            '<div class="eg-sizer" role="group" aria-label="Text size">' +
            '<button type="button" class="eg-tb-btn" data-eg-size="down" aria-label="Make text smaller" title="Smaller text"><span aria-hidden="true">A<sup>&minus;</sup></span></button>' +
            '<button type="button" class="eg-tb-btn eg-tb-reset" data-eg-size="reset" aria-label="Reset text size" title="Reset text size"><span class="eg-size-val">100%</span></button>' +
            '<button type="button" class="eg-tb-btn" data-eg-size="up" aria-label="Make text larger" title="Larger text"><span aria-hidden="true">A<sup>+</sup></span></button>' +
            "</div>" +
            '<button type="button" class="eg-tb-btn eg-tb-display" aria-expanded="false" aria-controls="egDisplayPanel"><span class="eg-aa" aria-hidden="true">Aa</span> Display</button>';
        live = h("div", { "class": "eg-sr-only", role: "status", "aria-live": "polite" });
        bar.appendChild(live);
        btn = bar.querySelector(".eg-tb-display");
        bar.querySelector("[data-eg-size=down]").addEventListener("click", api.smaller);
        bar.querySelector("[data-eg-size=up]").addEventListener("click", api.bigger);
        bar.querySelector("[data-eg-size=reset]").addEventListener("click", api.resetSize);

        panel = h("div", { "class": "eg-ui eg-panel", id: "egDisplayPanel", role: "dialog", "aria-label": "Display settings", hidden: "" });
        var sizes = SIZES.map(function (s, i) {
            return '<label class="eg-opt"><input type="radio" name="eg-size" value="' + i + '"> ' + s.label + ' <span class="eg-opt-sub">' + s.pct + "%</span></label>";
        }).join("");
        panel.innerHTML =
            '<div class="eg-panel-head"><h2 class="eg-panel-title">Display</h2><button type="button" class="eg-panel-close">Close</button></div>' +
            '<fieldset class="eg-group"><legend>Text size</legend><div class="eg-opts">' + sizes + "</div></fieldset>" +
            '<fieldset class="eg-group"><legend>Lesson pace</legend><div class="eg-opts eg-opts--col">' +
            '<label class="eg-opt"><input type="radio" name="eg-pace" value="normal"> Normal <span class="eg-opt-sub">moves on by itself after each step</span></label>' +
            '<label class="eg-opt"><input type="radio" name="eg-pace" value="slow"> Slow <span class="eg-opt-sub">longer pauses before and after each step</span></label>' +
            '<label class="eg-opt"><input type="radio" name="eg-pace" value="manual"> Step by step <span class="eg-opt-sub">waits for you to press Next</span></label>' +
            "</div></fieldset>" +
            '<fieldset class="eg-group"><legend>Viewing</legend><div class="eg-opts eg-opts--col">' +
            '<label class="eg-opt"><input type="checkbox" id="egHc"> High contrast guide</label>' +
            '<label class="eg-opt"><input type="checkbox" id="egMotion"> Reduce motion</label>' +
            "</div></fieldset>" +
            '<button type="button" class="eg-panel-reset">Reset all display settings</button>';
        panel.addEventListener("change", function (e) {
            var t = e.target;
            if (t.name === "eg-size") { set("size", +t.value); announce(); }
            else if (t.name === "eg-pace") set("pace", t.value);
            else if (t.id === "egHc") set("contrast", t.checked);
            else if (t.id === "egMotion") set("motion", t.checked ? "reduce" : (systemReduce() ? "full" : "system"));
        });
        panel.querySelector(".eg-panel-close").addEventListener("click", function () { closePanel(true); });
        panel.querySelector(".eg-panel-reset").addEventListener("click", function () { api.resetAll(); announce(); });
        btn.addEventListener("click", function () {
            var open = panel.hidden;
            panel.hidden = !open; btn.setAttribute("aria-expanded", String(open));
            if (open) { refreshUi(); var f = panel.querySelector("input:checked") || panel.querySelector("input"); if (f) f.focus(); }
        });
        document.addEventListener("keydown", function (e) { if (e.key === "Escape" && panel && !panel.hidden) { e.stopPropagation(); closePanel(true); } }, true);
        document.addEventListener("pointerdown", function (e) {
            if (panel && !panel.hidden && !panel.contains(e.target) && !btn.contains(e.target)) closePanel(false);
        }, true);
        (host || document.body).appendChild(bar);
        document.body.appendChild(panel);
        refreshUi();
        return bar;
    }
    api.buildToolbar = buildToolbar;
    api.onChange(refreshUi);
})();
