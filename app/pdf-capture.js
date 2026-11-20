/* PDF capture. Used only by tools/build-pdfs.mjs, which opens a lesson page with ?pdf=1,
   gives the page window.__ebobPdfShot(box) (a real browser screenshot of that area) and calls
   EbobPdfCapture.driveSelf(). The lesson runs itself step by step; each step is pictured with
   the red circle on it. Does nothing on a normal visit. */
(function () {
    "use strict";
    var LS = window.EbobLessons;
    if (!LS || new URLSearchParams(location.search).get("pdf") !== "1") return;

    function qs() { return new URLSearchParams(location.search); }
    function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
    function until(fn, ms) {
        ms = ms || 45000;
        return new Promise(function (res, rej) {
            var t0 = Date.now();
            (function poll() {
                var ok = false;
                try { ok = !!fn(); } catch (e) { ok = false; }
                if (ok) return res(true);
                if (Date.now() - t0 > ms) return rej(new Error("timed out waiting for the lesson"));
                setTimeout(poll, 40);
            })();
        });
    }

    /* the area to picture: the open window or menu, the red circle and the hand, at least 640 x 400 */
    function cropBox(doc) {
        var win = doc.defaultView;
        function R(e) {
            if (!e) return null;
            var r = e.getBoundingClientRect();
            if (r.width <= 2 || r.height <= 2) return null;
            var cs = win.getComputedStyle(e);
            if (cs.visibility === "hidden" || cs.display === "none") return null;
            return { l: r.left, t: r.top, r: r.right, b: r.bottom };
        }
        function U(a, b) { return !a ? b : !b ? a : { l: Math.min(a.l, b.l), t: Math.min(a.t, b.t), r: Math.max(a.r, b.r), b: Math.max(a.b, b.b) }; }
        var shell = R(doc.getElementById("appShell"));
        var u = null;
        u = U(u, R(doc.querySelector(".eg-ring")));
        u = U(u, R(doc.getElementById("ebobTutorialCircle")));
        u = U(u, R(doc.querySelector(".eg-hand.on")));
        var mark = u; /* the red circle and hand are never cut off, even outside the Binventory window (the taskbar) */
        /* topmost open Binventory window */
        var dlg = null, z = -1;
        doc.querySelectorAll(".backdrop.show").forEach(function (b) {
            var zz = parseInt(win.getComputedStyle(b).zIndex, 10) || 0;
            var r = R(b.querySelector(".modal, .modal-dialog, [class*='modal']"));
            if (r && zz >= z) { z = zz; dlg = r; }
        });
        u = U(u, dlg);
        u = U(u, R(doc.querySelector("#menuStrip .menu-root.open .menu-drop, #menuStrip .menu-root.open ul, #menuStrip .menu-root.open .dropdown")));
        /* installer window and its message boxes; the Windows start menu and search in the troubleshooting scenarios */
        doc.querySelectorAll(".iw, #mbLayer > *, .inst-closed:not([hidden]), #winStartMenu:not([hidden]), #winSearchFlyout:not([hidden])").forEach(function (e) { u = U(u, R(e)); });
        var S = shell ? U(shell, mark) : { l: 0, t: 0, r: win.innerWidth, b: win.innerHeight }; /* no Binventory window (installer): the whole page */
        S = { l: Math.max(0, S.l), t: Math.max(0, S.t), r: Math.min(win.innerWidth, S.r), b: Math.min(win.innerHeight, S.b) };
        if (!u) u = S;
        var pad = 18;
        u = { l: u.l - pad, t: u.t - pad, r: u.r + pad, b: u.b + pad };
        var minW = 640, minH = 400, cx = (u.l + u.r) / 2, cy = (u.t + u.b) / 2;
        if (u.r - u.l < minW) { u.l = cx - minW / 2; u.r = cx + minW / 2; }
        if (u.b - u.t < minH) { u.t = cy - minH / 2; u.b = cy + minH / 2; }
        function fit(lo, hi, a, b) { if (a < lo) { b += lo - a; a = lo; } if (b > hi) { a -= b - hi; b = hi; } return [Math.max(lo, a), Math.min(hi, b)]; }
        var x = fit(S.l, S.r, u.l, u.r), y = fit(S.t, S.b, u.t, u.b);
        return { x: Math.round(x[0]), y: Math.round(y[0]), width: Math.round(x[1] - x[0]), height: Math.round(y[1] - y[0]) };
    }
    async function shoot() {
        if (typeof window.__ebobPdfShot !== "function") throw new Error("no screenshot hook");
        await sleep(60);
        return window.__ebobPdfShot(cropBox(document));
    }

    /* typing and list steps are pictured with the value already in, so the reader sees what it should look like */
    function prefill(st) {
        if (!st || (st.do !== "type" && st.do !== "select") || typeof st.target !== "string") return;
        var e = document.querySelector(st.target);
        if (!e) return;
        if (st.do === "type") e.value = st.value;
        else for (var i = 0; i < e.options.length; i++) if (e.options[i].text.trim() === st.value) e.value = e.options[i].value;
        e.dispatchEvent(new Event("input", { bubbles: true }));
        e.dispatchEvent(new Event("change", { bubbles: true }));
        try { e.blur(); } catch (x) {}
    }

    function stepInfo(l, i) {
        if (l.steps && l.steps[i]) {
            var st = l.steps[i];
            return { title: st.title, html: typeof st.msg === "function" ? st.msg() : st.msg };
        }
        var ns = window.__ebobNativeTutorial && window.__ebobNativeTutorial.state();
        var step = ns && ns.steps && ns.steps[i];
        if (step) return { title: step.title, html: step.message };
        return { title: "Step " + (i + 1), html: "" };
    }

    function fireEnter(el) {
        if (!el) return;
        try { el.focus(); } catch (e) {}
        ["keydown", "keypress", "keyup"].forEach(function (type) {
            el.dispatchEvent(new KeyboardEvent(type, { key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true, cancelable: true }));
        });
    }
    function clickEl(el, how) {
        if (!el) return;
        var right = how === "right";
        var r = el.getBoundingClientRect(); /* click in the middle of the element, so a right-click menu opens beside it */
        var opts = { bubbles: true, cancelable: true, view: el.ownerDocument.defaultView, button: right ? 2 : 0, buttons: right ? 2 : 1,
            clientX: Math.round(r.left + Math.min(r.width / 2, 60)), clientY: Math.round(r.top + r.height / 2) };
        ["pointerdown", "mousedown", right ? "contextmenu" : "mouseup", "pointerup"].forEach(function (type) {
            if (type === "mouseup" || type === "pointerup") opts.buttons = 0;
            el.dispatchEvent(new MouseEvent(type, opts));
        });
        if (!right) el.click();
    }
    function clickRing(doc) {
        var circle = doc.getElementById("ebobTutorialCircle");
        if (!circle) return;
        var r = circle.getBoundingClientRect();
        var el = doc.elementFromPoint((r.left + r.right) / 2, (r.top + r.bottom) / 2);
        if (el && !el.closest("#ebobTutorialTooltip")) clickEl(el, "left");
    }
    /* the two troubleshooting scenarios run on the emulator's own tutorial engine */
    function actNative(win) {
        var raw = win.__ebobNativeTutorial && win.__ebobNativeTutorial.state();
        if (!raw || raw.completed) return "done";
        var step = raw.steps[raw.index] || {};
        var doc = win.document;
        if (raw.pending) return "pending";
        if (step.next) {
            var nb = doc.getElementById("ebobTutorialNext");
            if (nb) nb.click();
            return "next";
        }
        if (step.input) {
            var el = doc.querySelector(step.input);
            var m = (step.message || "").match(/["“]([^"”]+)["”]/);
            var txt = m ? m[1] : ((step.message || "").match(/\btype\s+(\S+?)(?:\s|\.?$)/i) || [])[1];
            if (el) {
                try { el.focus(); clickEl(el, "left"); } catch (e2) {}
                if (txt && !/^clear/i.test(step.message || "")) {
                    el.value = txt;
                    el.dispatchEvent(new Event("input", { bubbles: true }));
                    el.dispatchEvent(new Event("change", { bubbles: true }));
                } else if (/^clear/i.test(step.message || "")) {
                    el.value = "";
                    el.dispatchEvent(new Event("input", { bubbles: true }));
                }
            }
            return "type";
        }
        if (/press enter/i.test(step.message || "")) {
            fireEnter(doc.querySelector(step.input || "#cmdInput") || doc.querySelector(step.selector || ""));
            return "enter";
        }
        if (step.selector) {
            var t = doc.querySelector(step.selector);
            var right = /right-click/i.test(step.message || "");
            var dbl = /double-click/i.test(step.message || "");
            if (t) {
                if (dbl) { clickEl(t, "left"); t.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true })); }
                else clickEl(t, right ? "right" : "left");
            } else clickRing(doc);
            if (/press enter/i.test(step.message || "")) fireEnter(t || doc.querySelector("#cmdInput"));
            return "click";
        }
        clickRing(doc);
        return "wait";
    }

    async function driveSelf() {
        var sp = qs();
        var id = sp.get("lesson"), tutorial = sp.get("tutorial");
        var l = id ? LS.byId(id) : null;
        if (!l && tutorial && tutorial !== "free-mode") LS.lessons.forEach(function (x) { if (x.scenario === tutorial) l = x; });
        if (!l) throw new Error("no lesson in this page");
        if (window.EbobDisplay && window.EbobDisplay.PACES && window.EbobDisplay.PACES.picture) {
            try { window.EbobDisplay.set("pace", "picture"); window.EbobDisplay.set("size", 1); window.EbobDisplay.set("contrast", false); } catch (e) {}
        }
        document.documentElement.classList.add("eg-picture-mode");
        var n = l.stepCount, out = [];
        if (l.scenario) {
            await until(function () { return window.__ebobNativeTutorial && window.__ebobNativeTutorial.state() && window.__ebobNativeTutorial.state().total; }, 90000);
            for (var j = 0; j < n; j++) {
                await until(function () { var s = window.__ebobNativeTutorial.state(); return s && s.index === j; }, 30000);
                await sleep(500);
                var ni = stepInfo(l, j);
                out.push({ title: ni.title, html: ni.html, file: await shoot() });
                if (j < n - 1) {
                    var tries = 0;
                    while (tries < 8) {
                        var sNow = window.__ebobNativeTutorial.state();
                        if (sNow && (sNow.index > j || sNow.completed)) break;
                        if (sNow && sNow.pending) { await sleep(200); tries++; continue; }
                        actNative(window);
                        await sleep(350);
                        sNow = window.__ebobNativeTutorial.state();
                        if (sNow && sNow.index === j && !sNow.pending && tries % 2 === 1) {
                            fireEnter(document.querySelector((sNow.steps[j] && sNow.steps[j].input) || "#cmdInput"));
                            clickRing(document);
                        }
                        tries++;
                    }
                    await until(function () { var s = window.__ebobNativeTutorial.state(); return s && (s.index > j || s.completed); }, 20000);
                }
            }
        } else {
            var G = window.EbobGuide;
            await until(function () { var s = G && G.state && G.state(); return s && s.phase && s.phase !== "loading" && s.total; });
            for (var k = 0; k < n; k++) {
                await until(function () { var s = G.state(); return s && s.idx === k && (s.phase === "ready" || s.phase === "done" || s.finished); });
                window.__ebobPictureLock = true;
                await sleep(600);
                prefill(l.steps && l.steps[k]);
                await sleep(300);
                var gi = stepInfo(l, k);
                out.push({ title: gi.title, html: gi.html, file: await shoot() });
                window.__ebobPictureLock = false;
                if (k < n - 1) {
                    if (G.act) G.act();
                    await until(function () { var s = G.state(); return s && (s.idx > k || s.finished); });
                }
            }
        }
        if (out.length !== n) throw new Error("step count mismatch: got " + out.length + " expected " + n);
        return out;
    }

    window.EbobPdfCapture = { driveSelf: driveSelf, cropBox: cropBox };
})();
